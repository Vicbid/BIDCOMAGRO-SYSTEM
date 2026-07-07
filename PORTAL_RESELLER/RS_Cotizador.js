// ============================================================
// @version 1.0
//  PORTAL RESELLER — Cotizador de Presupuestos (cliente final)
//  Similar al carrito de repuestos (RS_Pedidos) pero:
//   - Precio base = PVP de lista (sin el 40% del reseller).
//   - Descuento POR ÍTEM, editable, arranca en 25%.
//   - Genera PDF + email + guarda historial (hoja COTIZACIONES).
//  Reutiliza helpers de RS_Pedidos/Env: _buildPriceMap, _lookupResellerMeta,
//  _fmtUsd, getSheetValues, getStockSheetValues, _construirEmailHTML, PORTAL_CONFIG.
// ============================================================

var _COT_DESC_DEFECTO = 25; // % de descuento inicial por ítem

function _asegurarHojaCotizaciones() {
  var ss = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.COTIZACIONES);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.COTIZACIONES);
    hoja.appendRow(['ID','Fecha','Reseller','Email Reseller','Cliente','Email Cliente','Cant. Ítems','Items JSON','Total USD','PDF URL','Observaciones']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 11).setBackground('#6c5ce7').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(1, 100);
    hoja.setColumnWidth(3, 160);
    hoja.setColumnWidth(5, 160);
    hoja.setColumnWidth(8, 300);
    hoja.setColumnWidth(10, 200);
  }
  return hoja;
}

function _siguienteNumeroCotizacion() {
  _asegurarHojaCotizaciones();
  var datos = getSheetValues(SCHEMA.SHEETS.COTIZACIONES);
  var max   = 0;
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][0] || '');
    var m  = id.match(/COT-(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1]));
  }
  var num = String(max + 1);
  while (num.length < 4) num = '0' + num;
  return 'COT-' + num;
}

// Normaliza un ítem del cotizador: fuerza precio de lista del catálogo cuando existe,
// clampa el descuento a 0–100 y calcula precio final + subtotal.
function _cotNormalizarItems(items, priceMap) {
  var out = [], total = 0;
  for (var i = 0; i < items.length; i++) {
    var it   = items[i] || {};
    var sku  = String(it.sku || '').trim();
    var skuU = sku.toUpperCase();
    var cant = Number(it.cantidad) || 1;
    if (cant < 1) cant = 1;

    // Precio de lista (PVP): del catálogo (priceMap está al 40%, se revierte a base) o el enviado.
    var precioLista;
    if (skuU && priceMap[skuU] !== undefined && priceMap[skuU] > 0) {
      precioLista = Math.round(priceMap[skuU] / 0.60 * 100) / 100;
    } else {
      precioLista = Number(it.precioLista) || 0;
    }

    var desc = Number(it.descuento);
    if (isNaN(desc)) desc = _COT_DESC_DEFECTO;
    if (desc < 0)   desc = 0;
    if (desc > 100) desc = 100;

    var precioFinal = Math.round(precioLista * (1 - desc / 100) * 100) / 100;
    var subtotal    = Math.round(precioFinal * cant * 100) / 100;
    total += subtotal;

    out.push({
      sku:         sku,
      descripcion: String(it.descripcion || '').trim(),
      modelos:     String(it.modelos || '').trim(),
      cantidad:    cant,
      precioLista: precioLista,
      descuento:   desc,
      precioFinal: precioFinal,
      subtotal:    subtotal
    });
  }
  return { items: out, total: Math.round(total * 100) / 100 };
}

// ── Confirmar cotización — genera PDF, guarda historial, manda email ──
// params: { reseller, cliente, clienteEmail, observaciones,
//           items: [{ sku, descripcion, cantidad, precioLista, descuento, modelos }] }
function RS_confirmarCotizacion(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    params = params || {};
    var reseller     = String(params.reseller     || '').trim();
    var cliente      = String(params.cliente      || '').trim();
    var clienteEmail = String(params.clienteEmail || '').trim();
    var obs          = String(params.observaciones|| '').trim();
    var itemsIn      = params.items || [];

    if (!reseller)        return { ok: false, error: 'Falta el reseller.' };
    if (!itemsIn.length)  return { ok: false, error: 'El carrito está vacío.' };
    if (!cliente)         return { ok: false, error: 'Ingresá el nombre del cliente.' };

    _asegurarHojaCotizaciones();

    // Precios de catálogo (enforcement) + normalización
    var priceMap = _buildPriceMap();
    var norm  = _cotNormalizarItems(itemsIn, priceMap);
    var items = norm.items;
    var total = norm.total;

    // Email del reseller + metadatos comerciales
    var emailReseller = '';
    try {
      var dRes = getSheetValues(SCHEMA.SHEETS.RESELLERS);
      var rLow = reseller.toLowerCase();
      for (var ri = 1; ri < dRes.length; ri++) {
        if (String(dRes[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rLow) {
          emailReseller = String(dRes[ri][SCHEMA.RESELLERS.EMAIL] || '').trim();
          break;
        }
      }
    } catch(eR) { Logger.log('RS_confirmarCotizacion email: ' + eR); }
    var resellerMeta = _lookupResellerMeta(reseller);

    var numero = _siguienteNumeroCotizacion();

    // PDF
    var pdfUrl = _generarPdfCotizacion(numero, resellerMeta, cliente, items, total, obs);

    // Guardar en COTIZACIONES
    var hoja = getSheet(SCHEMA.SHEETS.COTIZACIONES);
    if (hoja) {
      hoja.appendRow([
        numero,
        new Date(),
        reseller,
        emailReseller,
        cliente,
        clienteEmail,
        items.length,
        JSON.stringify(items),
        total > 0 ? total : '',
        pdfUrl || '',
        obs
      ]);
      invalidateSheetValues(SCHEMA.SHEETS.COTIZACIONES);
    }

    // Email (al reseller + cliente si cargó email)
    _enviarEmailCotizacion(numero, reseller, emailReseller, cliente, clienteEmail, items, total, pdfUrl, obs);

    return { ok: true, numero: numero, pdfUrl: pdfUrl || '', total: total };

  } catch(e) {
    Logger.log('RS_confirmarCotizacion ERROR: ' + e);
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}

// ── PDF de la cotización (hoja temporal → DriveApp.getAs) ─────────
function _generarPdfCotizacion(numero, resellerMeta, cliente, items, total, obs) {
  var tempSs = null;
  try {
    var meta     = resellerMeta || { nombre: '', direccion: '', telefono: '' };
    var fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    tempSs = SpreadsheetApp.create('TEMP_COT_' + numero);
    var sheet = tempSs.getActiveSheet();
    sheet.setName('Cotizacion');

    var ri = 1;

    // 1. Cabecera
    sheet.getRange(ri, 1, 3, 4).merge().setValue('BIDCOMAGRO')
      .setFontSize(20).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#6c5ce7').setVerticalAlignment('middle').setHorizontalAlignment('left');
    sheet.getRange(ri,     5, 1, 3).merge().setValue('PRESUPUESTO')
      .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#5a4bd1').setHorizontalAlignment('right').setVerticalAlignment('bottom');
    sheet.getRange(ri + 1, 5, 1, 3).merge().setValue('N\xba ' + numero)
      .setFontSize(10).setFontColor('#e4e0fb').setBackground('#5a4bd1').setHorizontalAlignment('right');
    sheet.getRange(ri + 2, 5, 1, 3).merge().setValue('Fecha: ' + fechaStr)
      .setFontSize(9).setFontColor('#e4e0fb').setBackground('#5a4bd1').setHorizontalAlignment('right').setVerticalAlignment('top');
    sheet.setRowHeight(ri, 24); sheet.setRowHeight(ri + 1, 18); sheet.setRowHeight(ri + 2, 20);
    ri += 3;

    // 2. Emisor (reseller) + Cliente
    sheet.setRowHeight(ri, 6); sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff'); ri++;

    sheet.getRange(ri, 1, 1, 7).merge().setValue('EMISOR')
      .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#2d3436').setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 20); ri++;

    var emisorRows = [
      ['Reseller',   meta.nombre    || '—'],
      ['Dirección',  meta.direccion || '—'],
      ['Teléfono',   meta.telefono  || '—']
    ];
    var bg1 = ['#f7f8fa', '#ffffff', '#f7f8fa'];
    for (var er = 0; er < emisorRows.length; er++) {
      sheet.getRange(ri, 1, 1, 2).merge().setValue(emisorRows[er][0])
        .setFontSize(9).setFontWeight('bold').setFontColor('#5e6778').setBackground(bg1[er]);
      sheet.getRange(ri, 3, 1, 5).merge().setValue(emisorRows[er][1])
        .setFontSize(9).setFontColor('#1a1a2e').setBackground(bg1[er]);
      sheet.setRowHeight(ri, 20); ri++;
    }

    sheet.setRowHeight(ri, 6); sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff'); ri++;
    sheet.getRange(ri, 1, 1, 7).merge().setValue('CLIENTE')
      .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#2d3436').setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 20); ri++;
    sheet.getRange(ri, 1, 1, 2).merge().setValue('Nombre')
      .setFontSize(9).setFontWeight('bold').setFontColor('#5e6778').setBackground('#f7f8fa');
    sheet.getRange(ri, 3, 1, 5).merge().setValue(cliente || '—')
      .setFontSize(9).setFontWeight('bold').setFontColor('#1a1a2e').setBackground('#f7f8fa');
    sheet.setRowHeight(ri, 20); ri++;

    // 3. Tabla de ítems
    sheet.setRowHeight(ri, 8); sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff'); ri++;

    var headers = ['SKU', 'Descripción', 'Cant.', 'PVP USD', 'Desc.', 'Precio USD', 'Subtotal USD'];
    sheet.getRange(ri, 1, 1, 7).setValues([headers])
      .setBackground('#6c5ce7').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22); ri++;

    for (var i = 0; i < items.length; i++) {
      var it    = items[i];
      var rowBg = (i % 2 === 0) ? '#ffffff' : '#f5f7fa';
      var rowVals = [
        it.sku         || '—',
        it.descripcion || '—',
        it.cantidad,
        it.precioLista > 0 ? _fmtUsd(it.precioLista) : '—',
        (it.descuento || 0) + '%',
        it.precioFinal > 0 ? _fmtUsd(it.precioFinal) : '—',
        it.subtotal    > 0 ? _fmtUsd(it.subtotal)    : '—'
      ];
      sheet.getRange(ri, 1, 1, 7).setValues([rowVals]).setFontSize(9).setBackground(rowBg).setVerticalAlignment('middle');
      sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#6c5ce7');
      sheet.getRange(ri, 3).setHorizontalAlignment('center');
      sheet.getRange(ri, 5).setHorizontalAlignment('center');
      sheet.getRange(ri, 4, 1, 4).setHorizontalAlignment('right');
      sheet.setRowHeight(ri, 20); ri++;
    }

    // 4. Total
    if (total > 0) {
      sheet.getRange(ri, 1, 1, 6).merge().setValue('TOTAL (no incluye impuestos)')
        .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff').setBackground('#2d3436')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 7).setValue(_fmtUsd(total))
        .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff').setBackground('#6c5ce7')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.setRowHeight(ri, 24); ri++;
    }

    if (obs) {
      sheet.setRowHeight(ri, 8); ri++;
      sheet.getRange(ri, 1, 1, 7).merge().setValue('Observaciones: ' + obs)
        .setFontSize(9).setFontColor('#5e6778').setWrap(true);
      ri++;
    }

    sheet.setRowHeight(ri + 1, 8); ri += 2;
    sheet.getRange(ri, 1, 1, 7).merge()
      .setValue('Presupuesto válido salvo variación de precios. Generado desde el Portal Resellers de BIDCOMAGRO.')
      .setFontSize(8).setFontColor('#9ba5b4').setHorizontalAlignment('center');

    // Anchos de columna
    sheet.setColumnWidth(1, 90); sheet.setColumnWidth(2, 230); sheet.setColumnWidth(3, 45);
    sheet.setColumnWidth(4, 90); sheet.setColumnWidth(5, 50);  sheet.setColumnWidth(6, 90);
    sheet.setColumnWidth(7, 100);
    sheet.getRange(1, 1, sheet.getMaxRows(), 7).setBorder(false, false, false, false, false, false);

    SpreadsheetApp.flush();

    // Exportar a PDF y mover a Drive
    var blob = tempSs.getAs('application/pdf').setName(numero + '.pdf');
    var file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = file.getUrl();
    return url;

  } catch(e) {
    Logger.log('_generarPdfCotizacion: ' + e);
    return '';
  } finally {
    try { if (tempSs) DriveApp.getFileById(tempSs.getId()).setTrashed(true); } catch(eD) {}
  }
}

// ── Email de la cotización ────────────────────────────────────
function _enviarEmailCotizacion(numero, reseller, emailReseller, cliente, clienteEmail, items, total, pdfUrl, obs) {
  try {
    var destinatarios = [];
    if (emailReseller && emailReseller.indexOf('@') !== -1) destinatarios.push(emailReseller);
    if (clienteEmail  && clienteEmail.indexOf('@')  !== -1) destinatarios.push(clienteEmail);
    if (!destinatarios.length) return;

    var filas = '';
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      filas +=
        '<tr style="background:' + (i % 2 === 0 ? '#ffffff' : '#f7f8fa') + '">' +
          '<td style="padding:7px 10px;font-size:11px;color:#6c5ce7;font-weight:700;border-bottom:1px solid #eef2f6">' + (it.sku || '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;color:#333;border-bottom:1px solid #eef2f6">' + (it.descripcion || '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:center;border-bottom:1px solid #eef2f6">' + it.cantidad + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:center;border-bottom:1px solid #eef2f6">' + (it.descuento || 0) + '%</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;border-bottom:1px solid #eef2f6">' + (it.precioFinal > 0 ? _fmtUsd(it.precioFinal) : '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;font-weight:700;border-bottom:1px solid #eef2f6">' + (it.subtotal > 0 ? _fmtUsd(it.subtotal) : '—') + '</td>' +
        '</tr>';
    }

    var pdfBloque = pdfUrl
      ? '<div style="margin-top:16px;text-align:center"><a href="' + pdfUrl + '" target="_blank" style="display:inline-block;padding:11px 26px;background:#6c5ce7;color:#fff;border-radius:7px;text-decoration:none;font-size:13px;font-weight:700">📄 Descargar presupuesto (PDF)</a></div>'
      : '';

    var cuerpo =
      '<p style="font-size:14px;color:#444;margin:0 0 6px">Presupuesto <strong style="color:#6c5ce7">' + numero + '</strong></p>' +
      '<p style="font-size:13px;color:#555;margin:0 0 18px">Cliente: <strong>' + (cliente || '—') + '</strong></p>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #e8e8e8;font-family:Arial,sans-serif">' +
        '<thead><tr style="background:#6c5ce7">' +
          '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:left">SKU</th>' +
          '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:left">Descripción</th>' +
          '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:center">Cant.</th>' +
          '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:center">Desc.</th>' +
          '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:right">Precio</th>' +
          '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:right">Subtotal</th>' +
        '</tr></thead><tbody>' + filas + '</tbody></table>' +
      (total > 0 ? '<p style="text-align:right;font-size:15px;font-weight:700;color:#1a1f2e;margin:12px 0 2px">Total: ' + _fmtUsd(total) + '</p>' +
                   '<p style="text-align:right;font-size:10px;color:#999;margin:0">No incluye impuestos</p>' : '') +
      (obs ? '<p style="font-size:12px;color:#666;margin-top:14px"><strong>Observaciones:</strong> ' + obs + '</p>' : '') +
      pdfBloque;

    var html = _construirEmailHTML('Presupuesto ' + numero, reseller, cuerpo, 'Portal Resellers BIDCOMAGRO · ' + reseller);

    GmailApp.sendEmail(destinatarios.join(','), 'Presupuesto ' + numero + (cliente ? ' — ' + cliente : ''), '', {
      htmlBody: html,
      name:     PORTAL_CONFIG.NOMBRE_REMITENTE,
      replyTo:  emailReseller && emailReseller.indexOf('@') !== -1 ? emailReseller : PORTAL_CONFIG.EMAIL_SUPERVISOR
    });

    try {
      var hojaLog = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
      if (hojaLog) hojaLog.appendRow([new Date(), numero, destinatarios.join(','), 'Cotización (Portal)', 'Presupuesto ' + numero, 'OK', '']);
    } catch(eL) {}

  } catch(e) {
    Logger.log('_enviarEmailCotizacion: ' + e);
  }
}

// ── Historial de cotizaciones del reseller ────────────────────
function obtenerHistorialCotizaciones(reseller) {
  try {
    _asegurarHojaCotizaciones();
    var datos = getSheetValues(SCHEMA.SHEETS.COTIZACIONES);
    var rLow  = String(reseller || '').trim().toLowerCase();
    var out   = [];
    for (var i = datos.length - 1; i >= 1; i--) {
      var f = datos[i];
      if (String(f[2] || '').trim().toLowerCase() !== rLow) continue;
      var fecha = f[1];
      out.push({
        id:        String(f[0] || '').trim(),
        fecha:     fecha instanceof Date
                     ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
                     : String(fecha || ''),
        cliente:   String(f[4] || '').trim(),
        cantItems: Number(f[6]) || 0,
        totalUsd:  Number(f[8]) || 0,
        pdfUrl:    String(f[9] || '').trim(),
        obs:       String(f[10] || '').trim()
      });
      if (out.length >= 30) break;
    }
    return { ok: true, historial: out };
  } catch(e) {
    Logger.log('obtenerHistorialCotizaciones: ' + e);
    return { ok: false, historial: [] };
  }
}
