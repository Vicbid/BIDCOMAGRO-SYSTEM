// ============================================================
// @version 2.3
//  WOS — Gestión de hilos Gmail · V-1.0 (Hitos 2–5)
//
//  Hito 1 vive en PORTAL_RESELLER/RS_Pedidos.js.
//  Este módulo requiere que COL.THREAD_ID (col R) ya esté
//  poblado por el Portal al confirmar el pedido.
//
//  TRIGGER a instalar: WOS_instalarTriggerDetector()
//  (ejecutar UNA VEZ desde el editor de Apps Script)
// ============================================================

// ── Helpers internos ─────────────────────────────────────────

// ── Template de email idéntico al Portal Reseller ─────────────
function _wosPortalHead(titulo) {
  return "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>" +
    "<body style='margin:0;padding:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",Helvetica,Arial,sans-serif'>" +
    "<table width='100%' cellpadding='0' cellspacing='0' style='background:#f2f4f7;padding:28px 12px'><tr><td>" +
    "<table width='600' align='center' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%'>" +
    "<tr><td style='background:#00a3e0;border-radius:10px 10px 0 0;padding:24px 32px'>" +
      "<table width='100%' cellpadding='0' cellspacing='0'><tr>" +
        "<td><div style='background:#fff;border-radius:5px;padding:4px 10px;display:inline-block'>" +
          "<span style='font-size:14px;font-weight:700;color:#00a3e0;letter-spacing:.1em'>DJI</span></div></td>" +
        "<td align='right'><span style='color:rgba(255,255,255,.85);font-size:11px'>BIDCOMAGRO · Portal Resellers</span></td>" +
      "</tr></table>" +
      "<h1 style='color:#fff;font-size:18px;font-weight:600;margin:18px 0 0;line-height:1.35'>" + titulo + "</h1>" +
    "</td></tr>" +
    "<tr><td style='background:#fff;padding:28px 32px'>";
}
function _wosPortalFoot(footerTxt) {
  return "</td></tr>" +
    "<tr><td style='background:#f9f9f9;border-top:1px solid #eee;border-radius:0 0 10px 10px;padding:16px 32px'>" +
      "<p style='font-size:11px;color:#aaa;margin:0;line-height:1.7'>" + (footerTxt || '') +
      "<br>Generado automaticamente por WOS · BIDCOMAGRO.</p>" +
    "</td></tr></table></td></tr></table></body></html>";
}

function _formatMoneda(n) {
  n = Number(n) || 0;
  var partes  = n.toFixed(2).split('.');
  var entero  = partes[0];
  var cents   = partes[1];
  var res = '';
  var len = entero.length;
  for (var i = 0; i < len; i++) {
    if (i > 0 && (len - i) % 3 === 0) res += '.';
    res += entero[i];
  }
  return res + ',' + cents;
}

function _trackingUrl(transp, track) {
  if (!track) return '';
  var t = String(transp || '').toLowerCase();
  if (t.indexOf('correo') >= 0)   return 'https://www.correoargentino.com.ar';
  if (t.indexOf('oca') >= 0)      return 'https://www.oca.com.ar';
  if (t.indexOf('andreani') >= 0) return 'https://www.andreani.com';
  if (t.indexOf('liqen') >= 0)    return 'https://www.liqen.com.ar';
  if (t.indexOf('via cargo') >= 0 || t.indexOf('viacargo') >= 0)
                                   return 'https://www.viacargo.com.ar';
  if (t.indexOf('credifin') >= 0) return 'https://www.credifin.com.ar';
  return '';
}

function _wosHoja() {
  return SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName(HOJA_PEDIDOS);
}

function _wosGetEmailReseller(nombre) {
  try {
    var d = SpreadsheetApp.openById(MASTER_SS_ID)
              .getSheetByName('Resellers').getDataRange().getValues();
    var n = String(nombre || '').trim().toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][COL_RS.NOMBRE] || '').trim().toLowerCase() === n)
        return String(d[i][COL_RS.EMAIL] || '').trim();
    }
  } catch(e) { Logger.log('_wosGetEmailReseller: ' + e); }
  return '';
}

// Lee todas las filas de un pedido desde NOTAS — resuelve la hoja por prefijo de número
function _wosLeerPedido(numero) {
  var hoja  = _getHojaPorNumero(numero);
  if (!hoja) return { hoja: null, datos: [], reseller: '', envio: '', pago: '', obs: '', threadId: '', tracking: '', items: [] };
  var datos = hoja.getDataRange().getValues();
  var res   = {
    hoja: hoja, datos: datos,
    reseller: '', envio: '', pago: '', obs: '',
    threadId: '', tracking: '',
    items: []
  };
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (!res.reseller) {
      res.reseller = String(datos[i][COL.RESELLER]     || '');
      res.envio    = String(datos[i][COL.ENVIO]        || '');
      res.pago     = String(datos[i][COL.PAGO]         || '');
      res.obs      = String(datos[i][COL.OBS]          || '');
      res.threadId = String(datos[i][COL.THREAD_ID]    || '').trim();
      res.tracking = String(datos[i][COL.TRACKING]     || '').trim();
    }
    res.items.push({
      row:      i + 1,                                        // 1-indexed para getRange
      sku:      String(datos[i][COL.SKU]       || ''),
      desc:     String(datos[i][COL.DESC]      || ''),
      cantSol:  Number(datos[i][COL.CANT_SOL]) || 0,
      cantDesp: Number(datos[i][COL.CANT_DESP])|| 0,
      cantPend: Number(datos[i][COL.CANT_PEND])|| 0,         // valor calculado por fórmula
      precio:   Number(datos[i][COL.PRECIO])   || 0,
      estado:   String(datos[i][COL.ESTADO]    || '')
    });
  }
  return res;
}

// Actualiza ESTADO y FECHA_ESTADO para todas las filas del pedido.
// Limpia la validación de celda antes de escribir para que cualquier
// estado definido en el código funcione sin depender del dropdown del sheet.
function _wosSetEstado(hoja, datos, numero, nuevoEstado) {
  var ahora = new Date();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    var rEst = hoja.getRange(i + 1, COL.ESTADO + 1);
    rEst.clearDataValidations();
    rEst.setValue(nuevoEstado);
    hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
  }
}

// Cambia ESTADO solo para filas cuyo estado actual sea estadoFiltro.
// Las demás filas del pedido no se modifican.
function _wosSetEstadoFiltrado(hoja, datos, numero, estadoFiltro, nuevoEstado) {
  var ahora = new Date();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (String(datos[i][COL.ESTADO] || '').trim() !== estadoFiltro) continue;
    var rEst = hoja.getRange(i + 1, COL.ESTADO + 1);
    rEst.clearDataValidations();
    rEst.setValue(nuevoEstado);
    hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
  }
}

// Para la notificación de faltante: fija estado por SKU.
// SKUs en skuFaltSet → estFalt (En_Espera_Reseller)
// SKUs restantes del mismo pedido → estDisp (Preparado)
function _wosSetEstadoPorSku(hoja, datos, numero, skuFaltSet, estFalt, estDisp) {
  var ahora = new Date();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    var sku = String(datos[i][COL.SKU] || '').trim().toUpperCase();
    var nuevoEst = skuFaltSet[sku] ? estFalt : estDisp;
    var rEst = hoja.getRange(i + 1, COL.ESTADO + 1);
    rEst.clearDataValidations();
    rEst.setValue(nuevoEst);
    hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
  }
}

// Lookup de metadata del reseller desde hoja master
function _wosGetResellerMeta(nombre) {
  try {
    var d = SpreadsheetApp.openById(MASTER_SS_ID)
              .getSheetByName('Resellers').getDataRange().getValues();
    var n = String(nombre || '').trim().toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() === n) {
        return {
          nombre:    String(d[i][0] || '').trim(),
          cuit:      String(d[i][1] || '').trim(),
          direccion: String(d[i][2] || '').trim(),
          telefono:  String(d[i][6] || '').trim()
        };
      }
    }
  } catch(e) { Logger.log('_wosGetResellerMeta: ' + e); }
  return { nombre: String(nombre || ''), cuit: '', direccion: '', telefono: '' };
}

// Devuelve el próximo número de nota para un pedido ("01", "02", ...) buscando en la carpeta Drive
function _wosNextNotaNum(numero) {
  try {
    var folder = DriveApp.getFolderById(_wosConfig().pdfFolderId);
    var prefix = 'NE_' + numero + '-';
    // searchFiles filtra en Drive — mucho más rápido que iterar getFiles() completo
    var files  = folder.searchFiles('title contains "' + prefix + '"');
    var max    = 0;
    while (files.hasNext()) {
      var fname  = files.next().getName();
      if (fname.indexOf(prefix) === 0) {
        var parsed = parseInt(fname.substring(prefix.length)) || 0;
        if (parsed > max) max = parsed;
      }
    }
    var n = max + 1;
    return n < 10 ? '0' + n : String(n);
  } catch(e) {
    Logger.log('_wosNextNotaNum: ' + e);
    return '01';
  }
}

// Genera PDF de Nota de Entrega, lo guarda en Drive y devuelve { blob, url, nombreNota }
// notaNumStr: "01", "02", etc. (usar _wosNextNotaNum para obtenerlo)
// transp y costoEnvio son opcionales. bultos = [{tracking, peso}, ...]
function _wosGenerarPDF(numero, notaNumStr, reseller, items, fecha, transp, bultos, costoEnvio) {
  var tempSs = null;
  try {
    var meta     = _wosGetResellerMeta(reseller);
    var fechaStr = fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    transp       = String(transp || '').trim();
    bultos       = bultos || [];
    costoEnvio   = Number(costoEnvio) || 0;
    var pesoTotal = 0;
    for (var bi = 0; bi < bultos.length; bi++) pesoTotal += Number(bultos[bi].peso) || 0;
    pesoTotal = Math.round(pesoTotal * 1000) / 1000;

    tempSs = SpreadsheetApp.create('TEMP_NE_' + numero);
    var sheet = tempSs.getActiveSheet();
    sheet.setName('Nota de Entrega');
    var ri = 1;

    // ── 1. CABECERA ───────────────────────────────────────────
    sheet.getRange(ri, 1, 3, 4).merge().setValue('BIDCOMAGRO')
      .setFontSize(20).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#00a3e0').setVerticalAlignment('middle').setHorizontalAlignment('left');
    sheet.getRange(ri,     5, 1, 3).merge().setValue('NOTA DE ENTREGA')
      .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('bottom');
    sheet.getRange(ri + 1, 5, 1, 3).merge().setValue('Nº ' + numero + '-' + notaNumStr)
      .setFontSize(10).setFontColor('#cce8f4')
      .setBackground('#007ab3').setHorizontalAlignment('right');
    sheet.getRange(ri + 2, 5, 1, 3).merge().setValue('Fecha: ' + fechaStr)
      .setFontSize(9).setFontColor('#cce8f4')
      .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('top');
    sheet.setRowHeight(ri, 24); sheet.setRowHeight(ri+1, 18); sheet.setRowHeight(ri+2, 20);
    ri += 3;

    // ── 2. DATOS DEL RESELLER ─────────────────────────────────
    sheet.setRowHeight(ri, 6);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    sheet.getRange(ri, 1, 1, 7).merge().setValue('DATOS DEL RESELLER DESTINATARIO')
      .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#2d3436').setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22);
    ri++;

    var resellerRows = [
      ['Nombre / Razón Social', meta.nombre    || '—'],
      ['Dirección Comercial',   meta.direccion || '—'],
      ['Teléfono de Contacto',  meta.telefono  || '—'],
      ['Transportista',         transp         || '—'],
      ['Costo de Envío (ARS)',        costoEnvio > 0 ? '$ ' + _formatMoneda(costoEnvio) : '—']
    ];
    for (var bri = 0; bri < bultos.length; bri++) {
      var bt     = bultos[bri];
      var bLabel = bultos.length > 1 ? 'Bulto ' + (bri + 1) : 'Nº de Seguimiento';
      var bVal   = (bt.tracking || '—') + (bt.peso > 0 ? '  ·  ' + bt.peso + ' kg' : '');
      resellerRows.push([bLabel, bVal]);
    }
    if (pesoTotal > 0 && bultos.length > 1) resellerRows.push(['Peso Total', pesoTotal + ' kg']);

    for (var rri = 0; rri < resellerRows.length; rri++) {
      var rBg = rri % 2 === 0 ? '#f7f8fa' : '#ffffff';
      sheet.getRange(ri, 1, 1, 2).merge().setValue(resellerRows[rri][0])
        .setFontSize(9).setFontWeight('bold').setFontColor('#5e6778').setBackground(rBg);
      sheet.getRange(ri, 3, 1, 5).merge().setValue(resellerRows[rri][1])
        .setFontSize(9).setFontWeight(rri === 0 ? 'bold' : 'normal')
        .setFontColor('#1a1a2e').setBackground(rBg);
      sheet.setRowHeight(ri, 20);
      ri++;
    }

    // ── 3. TABLA DE ÍTEMS ─────────────────────────────────────
    sheet.setRowHeight(ri, 8);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    sheet.getRange(ri, 1, 1, 7).setValues([['SKU', 'Descripción', 'Cant.', 'PVP USD', 'Precio Res. USD', 'Subtotal USD', '']])
      .setBackground('#00a3e0').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22);
    ri++;

    // Determinar si renderizar agrupado por caja
    var hasCajaAssign = bultos.length > 1;
    if (hasCajaAssign) {
      var anyAssigned = false;
      for (var ci = 0; ci < items.length; ci++) {
        if (items[ci].cajaIdx !== undefined) { anyAssigned = true; break; }
      }
      hasCajaAssign = anyAssigned;
    }

    var total = 0;
    if (hasCajaAssign) {
      // Ordenar ítems por cajaIdx
      var sortedItems = items.slice().sort(function(a, b) {
        return (a.cajaIdx || 0) - (b.cajaIdx || 0);
      });
      var currentCajaIdx = -1;
      var altRow = 0;
      for (var i = 0; i < sortedItems.length; i++) {
        var it      = sortedItems[i];
        var cajaIdx = it.cajaIdx !== undefined ? it.cajaIdx : 0;
        // Fila separadora de caja
        if (cajaIdx !== currentCajaIdx) {
          currentCajaIdx = cajaIdx;
          altRow = 0;
          var blt      = bultos[cajaIdx] || {};
          var cajaLbl  = 'Caja ' + (cajaIdx + 1);
          if (blt.tracking) cajaLbl += '  ·  Tracking: ' + blt.tracking;
          if (blt.peso > 0) cajaLbl += '  ·  ' + blt.peso + ' kg';
          sheet.getRange(ri, 1, 1, 7).merge().setValue(cajaLbl)
            .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
            .setBackground('#005082').setHorizontalAlignment('left').setVerticalAlignment('middle');
          sheet.setRowHeight(ri, 20);
          ri++;
        }
        var cant  = Number(it.cantDesp) || 0;
        var prec  = Number(it.precio)   || 0;
        var pvp   = prec > 0 ? Math.round(prec / 0.60 * 100) / 100 : 0;
        var sub   = cant * prec;
        total    += sub;
        var rowBg = (altRow % 2 === 0) ? '#ffffff' : '#f5f7fa';
        sheet.getRange(ri, 1, 1, 7).setValues([[
          it.sku  || '—',
          it.desc || '—',
          cant,
          pvp  > 0 ? 'USD ' + _formatMoneda(pvp)  : '—',
          prec > 0 ? 'USD ' + _formatMoneda(prec) : '—',
          sub  > 0 ? 'USD ' + _formatMoneda(sub)  : '—',
          ''
        ]]).setFontSize(9).setBackground(rowBg).setVerticalAlignment('middle');
        sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#00a3e0');
        sheet.getRange(ri, 3).setHorizontalAlignment('center');
        sheet.getRange(ri, 4, 1, 3).setHorizontalAlignment('right');
        sheet.setRowHeight(ri, 20);
        ri++;
        altRow++;
      }
    } else {
      for (var i = 0; i < items.length; i++) {
        var it   = items[i];
        var cant = Number(it.cantDesp) || 0;
        var prec = Number(it.precio)   || 0;
        var pvp  = prec > 0 ? Math.round(prec / 0.60 * 100) / 100 : 0;
        var sub  = cant * prec;
        total   += sub;
        var rowBg = (i % 2 === 0) ? '#ffffff' : '#f5f7fa';
        sheet.getRange(ri, 1, 1, 7).setValues([[
          it.sku  || '—',
          it.desc || '—',
          cant,
          pvp  > 0 ? 'USD ' + _formatMoneda(pvp)  : '—',
          prec > 0 ? 'USD ' + _formatMoneda(prec) : '—',
          sub  > 0 ? 'USD ' + _formatMoneda(sub)  : '—',
          ''
        ]]).setFontSize(9).setBackground(rowBg).setVerticalAlignment('middle');
        sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#00a3e0');
        sheet.getRange(ri, 3).setHorizontalAlignment('center');
        sheet.getRange(ri, 4, 1, 3).setHorizontalAlignment('right');
        sheet.setRowHeight(ri, 20);
        ri++;
      }
    }

    // ── 4. FILA TOTAL ─────────────────────────────────────────
    if (total > 0) {
      sheet.getRange(ri, 1, 1, 5).merge()
        .setValue('TOTAL (precio reseller · no incluye impuestos)')
        .setFontSize(10).setFontWeight('bold').setFontColor('#5e6778')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 6).setValue('USD ' + _formatMoneda(total))
        .setFontSize(11).setFontWeight('bold').setFontColor('#00a3e0')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 7).setValue('').setBackground('#e8f5fc');
      sheet.setRowHeight(ri, 24);
      ri++;
    }

    // ── 5. PIE ────────────────────────────────────────────────
    sheet.setRowHeight(ri, 4);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#00a3e0');
    ri++;
    sheet.getRange(ri, 1, 1, 7).merge()
      .setValue('Documento generado automáticamente — WOS · BIDCOMAGRO · ' + meta.nombre)
      .setFontSize(8).setFontStyle('italic').setFontColor('#9ba5b4')
      .setHorizontalAlignment('center').setBackground('#ffffff');
    sheet.setRowHeight(ri, 16);

    // Anchos de columna
    sheet.setColumnWidth(1, 110); // SKU
    sheet.setColumnWidth(2, 210); // Descripción
    sheet.setColumnWidth(3, 48);  // Cant.
    sheet.setColumnWidth(4, 78);  // PVP USD
    sheet.setColumnWidth(5, 82);  // Precio Res.
    sheet.setColumnWidth(6, 82);  // Subtotal
    sheet.setColumnWidth(7, 1);   // spacer

    // ── HOJA ANEXO: NÚMEROS DE SERIE ─────────────────────────────
    var itemsConSerial = [];
    for (var sci = 0; sci < items.length; sci++) {
      if (items[sci].seriales) itemsConSerial.push(items[sci]);
    }
    if (itemsConSerial.length > 0) {
      // Nota al pie de la NE principal
      sheet.getRange(ri, 1, 1, 7).merge()
        .setValue('Ver Anexo — N\xfameros de Serie (p\xe1gina siguiente)')
        .setFontSize(8).setFontStyle('italic').setFontColor('#00a3e0')
        .setHorizontalAlignment('center').setBackground('#f0f8ff');
      sheet.setRowHeight(ri, 14);
      ri++;

      var sSheet = tempSs.insertSheet('N\xba de Serie');
      var sri = 1;

      // Cabecera
      sSheet.getRange(sri, 1, 3, 4).merge().setValue('BIDCOMAGRO')
        .setFontSize(16).setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#00a3e0').setVerticalAlignment('middle');
      sSheet.getRange(sri,     5, 1, 3).merge().setValue('ANEXO — N\xdaMEROS DE SERIE')
        .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('bottom');
      sSheet.getRange(sri + 1, 5, 1, 3).merge().setValue('NE ' + numero + '-' + notaNumStr)
        .setFontSize(9).setFontColor('#cce8f4')
        .setBackground('#007ab3').setHorizontalAlignment('right');
      sSheet.getRange(sri + 2, 5, 1, 3).merge().setValue(meta.nombre || '')
        .setFontSize(9).setFontColor('#cce8f4')
        .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('top');
      sSheet.setRowHeight(sri, 24); sSheet.setRowHeight(sri+1, 18); sSheet.setRowHeight(sri+2, 18);
      sri += 3;

      sSheet.setRowHeight(sri, 8);
      sSheet.getRange(sri, 1, 1, 7).merge().setBackground('#ffffff');
      sri++;

      for (var sii = 0; sii < itemsConSerial.length; sii++) {
        var sIt    = itemsConSerial[sii];
        var sNums  = String(sIt.seriales).split(',');

        // Sub-cabecera del ítem
        sSheet.getRange(sri, 1, 1, 7).merge()
          .setValue(sIt.sku + '   \xb7   ' + sIt.desc + '   (' + sIt.cantDesp + ' u.)')
          .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#2d3436');
        sSheet.setRowHeight(sri, 22);
        sri++;

        // Seriales numerados
        for (var sni = 0; sni < sNums.length; sni++) {
          var snVal = String(sNums[sni]).trim();
          if (!snVal) continue;
          var snBg = sni % 2 === 0 ? '#f7f8fa' : '#ffffff';
          sSheet.getRange(sri, 1, 1, 2).merge()
            .setValue(sni + 1)
            .setFontSize(9).setFontColor('#5e6778').setBackground(snBg)
            .setHorizontalAlignment('center').setFontWeight('bold');
          sSheet.getRange(sri, 3, 1, 5).merge()
            .setValue(snVal)
            .setFontSize(10).setFontColor('#1a1a2e').setBackground(snBg)
            .setFontFamily('Courier New').setFontWeight('bold');
          sSheet.setRowHeight(sri, 22);
          sri++;
        }

        if (sii < itemsConSerial.length - 1) {
          sSheet.setRowHeight(sri, 8);
          sSheet.getRange(sri, 1, 1, 7).merge().setBackground('#ffffff');
          sri++;
        }
      }

      // Pie
      sSheet.setRowHeight(sri, 4);
      sSheet.getRange(sri, 1, 1, 7).merge().setBackground('#00a3e0');
      sri++;
      sSheet.getRange(sri, 1, 1, 7).merge()
        .setValue('Documento generado autom\xe1ticamente — WOS \xb7 BIDCOMAGRO \xb7 ' + meta.nombre)
        .setFontSize(8).setFontStyle('italic').setFontColor('#9ba5b4')
        .setHorizontalAlignment('center').setBackground('#ffffff');
      sSheet.setRowHeight(sri, 16);

      // Anchos
      sSheet.setColumnWidth(1, 40);
      sSheet.setColumnWidth(2, 10);
      sSheet.setColumnWidth(3, 350);
      sSheet.setColumnWidth(4, 1);
      sSheet.setColumnWidth(5, 1);
      sSheet.setColumnWidth(6, 1);
      sSheet.setColumnWidth(7, 1);
    }

    SpreadsheetApp.flush();

    var nombreNota = 'NE_' + numero + '-' + notaNumStr + '.pdf';
    var pdfBlob    = DriveApp.getFileById(tempSs.getId()).getAs('application/pdf');
    pdfBlob.setName(nombreNota);
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);

    var pdfUrl = '';
    try {
      var folder  = DriveApp.getFolderById(_wosConfig().pdfFolderId);
      var pdfFile = folder.createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = pdfFile.getUrl();
    } catch(eDrive) {
      Logger.log('_wosGenerarPDF Drive save: ' + eDrive);
    }

    return { blob: pdfBlob, url: pdfUrl, nombreNota: nombreNota };
  } catch(e) {
    Logger.log('_wosGenerarPDF: ' + e);
    if (tempSs) { try { DriveApp.getFileById(tempSs.getId()).setTrashed(true); } catch(eT) {} }
    return null;
  }
}


// ─────────────────────────────────────────────────────────────
//  HITO 4b — Despachar Completo (reemplaza WOS_despacharPedido)
//  Registra cantidades, guarda tracking, responde en el hilo
//  original y notifica a facturación en un solo paso.
//
//  despachos:    [{row (1-indexed), cantDesp}]
//  transportista: string requerido ("Correo Argentino", "Credifin", "Via Cargo")
//  bultos:        [{tracking, peso}] — un objeto por bulto físico
//  costoEnvio:    costo total del envío (opcional)
// ─────────────────────────────────────────────────────────────
function WOS_despacharCompleto(numero, despachos, transportista, bultos, costoEnvio, operario) {
  try {
    var ped = _wosLeerPedido(numero);
    if (!ped.reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    var esRetiroDesp = String(ped.envio || '').toLowerCase().indexOf('retiro') >= 0;
    if (!ped.threadId && !esRetiroDesp) return { ok: false, error: 'Sin Thread_ID en col R. Verificar que el Portal guardó el hilo al crear el pedido.' };

    var email = _wosGetEmailReseller(ped.reseller);
    if (!email) return { ok: false, error: 'Email no encontrado para: ' + ped.reseller };

    operario    = String(operario || '');
    var transp  = String(transportista || '').trim();
    bultos      = bultos || [];
    var costo   = Number(costoEnvio) || 0;
    // Computar tracking string y peso total desde los bultos
    var trackParts = [];
    var pesoTotal  = 0;
    for (var bi = 0; bi < bultos.length; bi++) {
      var btTrack = String(bultos[bi].tracking || '').trim();
      if (btTrack) trackParts.push(btTrack);
      pesoTotal += Number(bultos[bi].peso) || 0;
    }
    var track = trackParts.join(' | ');
    var peso  = Math.round(pesoTotal * 1000) / 1000;
    var tz     = Session.getScriptTimeZone();
    var fecha  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    var ahora  = new Date();

    var despMap    = {};
    var serialMap  = {};
    var cajaMap    = {};
    var ubicMap    = {};
    for (var d = 0; d < despachos.length; d++) {
      despMap[despachos[d].row]   = Number(despachos[d].cantDesp) || 0;
      serialMap[despachos[d].row] = String(despachos[d].seriales || '').trim();
      cajaMap[despachos[d].row]   = despachos[d].cajaIdx !== undefined ? Number(despachos[d].cajaIdx) : 0;
      ubicMap[despachos[d].row]   = String(despachos[d].ubicacion || '').trim();
    }

    var carmenSS    = null;
    var carmenHoja  = null;
    var carmenUbicH = null;
    try {
      carmenSS    = SpreadsheetApp.openById(CARMEN_SS_ID);
      carmenHoja  = carmenSS.getSheetByName('Entregados');
      carmenUbicH = carmenSS.getSheetByName(CARMEN_UBICACIONES_TAB);
    } catch(eC) {}

    // Calcular número de nota antes del loop para poder escribirlo en col P por ítem
    var notaNumStr   = _wosNextNotaNum(numero);
    var notaEntrega  = 'NE_' + numero + '-' + notaNumStr;

    var itemsDesp = [];
    var filasDesp = [];   // filas (1-indexed) donde dispNow > 0 para escribir NE_URL post-PDF
    var totalUSD  = 0;

    for (var i = 1; i < ped.datos.length; i++) {
      if (String(ped.datos[i][COL.NUMERO] || '').trim() !== numero) continue;

      // Despacho aditivo: el modal envía cuánto se despacha AHORA;
      // el sheet guarda el acumulado total (oldCantDesp + dispNow).
      var oldCantDesp = Number(ped.datos[i][COL.CANT_DESP]) || 0;
      var dispNow     = (despMap[i + 1] !== undefined) ? (Number(despMap[i + 1]) || 0) : 0;
      var cantFinal   = oldCantDesp + dispNow;

      if (dispNow > 0) {
        // CANT_DESP aislado (col 6)
        ped.hoja.getRange(i + 1, COL.CANT_DESP + 1).setValue(cantFinal);
        // FECHA_DESPACHO (col 15), NOTA_ENTREGA (col 16), TRACKING (col 17) — bloque contiguo
        ped.hoja.getRange(i + 1, COL.FECHA_DESPACHO + 1, 1, 3).setValues([[ahora, notaEntrega, track || '']]);
        // FECHA_ESTADO (col 19), TRANSPORTISTA_DESP (col 20), COSTO_ENVIO (col 21), PESO_ENVIO (col 22) — bloque contiguo
        ped.hoja.getRange(i + 1, COL.FECHA_ESTADO + 1, 1, 4).setValues([[ahora, transp, costo > 0 ? costo : '', peso > 0 ? peso : '']]);
        if (operario) ped.hoja.getRange(i + 1, COL.OPERARIO + 1).setValue(operario);
        var rowSeriales = serialMap[i + 1] || '';
        if (rowSeriales) ped.hoja.getRange(i + 1, COL.SERIALES + 1).setValue(rowSeriales);
        filasDesp.push(i + 1);
      }

      if (dispNow > 0 && carmenHoja) {
        var _skuDesp  = String(ped.datos[i][COL.SKU]  || '').trim().toUpperCase();
        var _ubicDesp = ubicMap[i + 1] || '';
        carmenHoja.appendRow([_skuDesp, String(ped.datos[i][COL.DESC] || ''), dispNow, String(numero || ''), '', '', fecha, _ubicDesp]);
        // Restar de UBICACIONES si hay ubicación seleccionada
        if (_ubicDesp && carmenUbicH) {
          var _dU = carmenUbicH.getDataRange().getValues();
          for (var ui = 1; ui < _dU.length; ui++) {
            if (String(_dU[ui][0] || '').trim().toUpperCase() === _skuDesp &&
                String(_dU[ui][1] || '').trim().toUpperCase() === _ubicDesp.toUpperCase()) {
              var _nueva = Math.max(0, (parseFloat(_dU[ui][2]) || 0) - dispNow);
              carmenUbicH.getRange(ui + 1, 3).setValue(_nueva);
              break;
            }
          }
        }
      }

      if (dispNow > 0) {
        var prec = Number(ped.datos[i][COL.PRECIO]) || 0;
        itemsDesp.push({
          sku:      String(ped.datos[i][COL.SKU]  || ''),
          desc:     String(ped.datos[i][COL.DESC] || ''),
          cantDesp: dispNow, precio: prec,
          seriales: serialMap[i + 1] || '',
          cajaIdx:  cajaMap[i + 1] !== undefined ? cajaMap[i + 1] : 0
        });
        totalUSD += dispNow * prec;
      }
    }

    if (!itemsDesp.length) return { ok: false, error: 'Ningún ítem con cantidad > 0.' };

    var pdfResult  = _wosGenerarPDF(numero, notaNumStr, ped.reseller, itemsDesp, fecha, transp, bultos, costo);
    var pdfBlob    = pdfResult ? pdfResult.blob : null;
    var pdfUrl     = pdfResult ? pdfResult.url  : '';

    // Guardar link de la Nota de Entrega en col W para cada ítem despachado
    if (pdfUrl) {
      for (var fi = 0; fi < filasDesp.length; fi++) {
        ped.hoja.getRange(filasDesp[fi], COL.NE_URL + 1).setValue(pdfUrl);
      }
    }

    // ── Email al reseller: misma estructura visual que Portal Reseller ─
    // Tabla de ítems estilo Portal Reseller
    var tbodyRows = '';
    for (var ri = 0; ri < itemsDesp.length; ri++) {
      var itR = itemsDesp[ri];
      var serRow = itR.seriales
        ? "<tr><td colspan='4' style='padding:3px 10px 8px 10px;font-size:10px;color:#777;border-bottom:1px solid #eee'>" +
          "<span style='font-weight:600;color:#555'>N\xba serie:</span> " + itR.seriales + "</td></tr>"
        : '';
      tbodyRows +=
        "<tr>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";font-family:Consolas,monospace;font-size:12px;color:#00a3e0'>" + (itR.sku  || '—') + "</td>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";font-size:12px;color:#333'>"                                   + (itR.desc || '—') + "</td>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";text-align:center;font-weight:700;color:#333'>"                + itR.cantDesp       + "</td>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";text-align:right;font-size:12px;color:#555'>"                 + (itR.precio > 0 ? 'USD ' + _formatMoneda(Number(itR.precio)) : '—') + "</td>" +
        "</tr>" + serRow;
    }
    var tablaHtml =
      "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:1px solid #dde3ea;border-radius:8px;overflow:hidden'>" +
      "<thead><tr style='background:#f0f5fa'>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>SKU</th>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>Descripcion</th>" +
        "<th style='padding:9px 10px;text-align:center;font-size:11px;font-weight:700;color:#888;width:60px'>Cant.</th>" +
        "<th style='padding:9px 10px;text-align:right;font-size:11px;font-weight:700;color:#888;width:90px'>P. Unit.</th>" +
      "</tr></thead><tbody>" + tbodyRows + "</tbody></table>";

    // Total
    var totalHtml =
      "<div style='text-align:right;margin-top:10px'>" +
        "<span style='font-size:13px;color:#444'>Total despachado: </span>" +
        "<strong style='font-size:15px;color:#00a3e0'>USD " + _formatMoneda(totalUSD) + "</strong>" +
        "<div style='font-size:10px;color:#999;margin-top:3px'>Precios expresados no incluyen impuestos</div>" +
      "</div>";

    // Chips de logística por bulto
    var chipBase = "background:#f0f5fa;border:1px solid #dde3ea;border-radius:6px;padding:8px 14px;font-size:12px";
    var chipLbl  = "color:#888;font-size:10px;text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:2px";
    var chipsHtml = "<div style='margin-top:14px;display:flex;gap:10px;flex-wrap:wrap'>" +
      (transp ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Transportista</span><strong style='color:#1a1a2e'>" + transp + "</strong></div>" : '') +
      (ped.pago ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Forma de pago</span><strong style='color:#1a1a2e'>" + ped.pago + "</strong></div>" : '') +
      (costo > 0 ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Costo de Envío (ARS)</span><strong style='color:#1a1a2e'>$ " + _formatMoneda(costo) + "</strong></div>" : '') +
      (peso  > 0 ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Peso total</span><strong style='color:#1a1a2e'>" + peso + " kg</strong></div>" : '');
    for (var ci = 0; ci < bultos.length; ci++) {
      var bt = bultos[ci];
      if (!bt.tracking && !(bt.peso > 0)) continue;
      var bLbl = bultos.length > 1 ? 'Bulto ' + (ci + 1) : 'N. de seguimiento';
      var bVal = (bt.tracking || '—') + (bt.peso > 0 ? ' &nbsp;·&nbsp; ' + bt.peso + ' kg' : '');
      chipsHtml += "<div style='" + chipBase + "'><span style='" + chipLbl + "'>" + bLbl + "</span><strong style='color:#1a1a2e;font-family:Consolas,monospace;font-size:11px'>" + bVal + "</strong></div>";
    }
    chipsHtml += "</div>";

    // Botones de rastreo (uno por bulto con tracking)
    var rastrearBtns = '';
    for (var rbi = 0; rbi < bultos.length; rbi++) {
      var rbt = bultos[rbi];
      if (!rbt.tracking) continue;
      var rUrl = _trackingUrl(transp, rbt.tracking);
      if (!rUrl) continue;
      var rLbl = bultos.length > 1 ? 'Rastrear Bulto ' + (rbi + 1) : 'Rastrear envío';
      rastrearBtns +=
        "<a href='" + rUrl + "' target='_blank' style='display:inline-flex;align-items:center;gap:8px;padding:10px 22px;" +
        "background:#00a3e0;color:#fff;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600;margin:4px'>" +
        rLbl + " &rarr;</a>";
    }
    var rastrearBtn = rastrearBtns
      ? "<div style='margin-top:16px;text-align:center;display:flex;flex-wrap:wrap;justify-content:center;gap:6px'>" + rastrearBtns + "</div>"
      : '';

    // Botón Drive (igual al botón PDF del Portal Reseller)
    var driveBtn = pdfUrl
      ? "<div style='margin-top:12px;text-align:center'>" +
          "<a href='" + pdfUrl + "' target='_blank' style='display:inline-flex;align-items:center;gap:8px;padding:10px 22px;" +
          "background:#fff;border:1px solid #dde3ea;color:#00a3e0;border-radius:6px;text-decoration:none;" +
          "font-size:13px;font-weight:600'>Ver / Descargar Nota de Entrega &rarr;</a>" +
        "</div>"
      : '';

    var obsHtml = ped.obs
      ? "<div style='margin-top:14px;background:#fffbe6;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px'>" +
          "<strong style='font-size:11px;color:#7a5800;text-transform:uppercase;letter-spacing:.06em'>Observaciones</strong>" +
          "<p style='margin:5px 0 0;font-size:13px;color:#444'>" + ped.obs + "</p>" +
        "</div>"
      : '';

    var confirmEntregaHtml =
      "<div style='margin-top:18px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:12px 16px'>" +
        "<p style='margin:0;font-size:12px;color:#0369a1;line-height:1.6'>" +
          "Cuando recibas el paquete, <strong>respondé este correo con la palabra \"Recibido\"</strong> " +
          "para confirmar la entrega. Gracias!" +
        "</p>" +
      "</div>";

    // Detectar tipo de envío:
    // hayDespPrevio      = algún ítem ya tenía CANT_DESP > 0 antes de este despacho → 2º envío
    // hayBackorderPostDesp = quedarán ítems en Backorder después de este despacho
    var hayDespPrevio       = false;
    var hayBackorderPostDesp = false;
    for (var bk = 1; bk < ped.datos.length; bk++) {
      if (String(ped.datos[bk][COL.NUMERO] || '').trim() !== numero) continue;
      var oldDespBk  = Number(ped.datos[bk][COL.CANT_DESP]) || 0;
      if (oldDespBk > 0) { hayDespPrevio = true; }
      var dispNowBk  = (despMap[bk + 1] !== undefined) ? (Number(despMap[bk + 1]) || 0) : 0;
      if (dispNowBk > 0) {
        var cantSolBk   = Number(ped.datos[bk][COL.CANT_SOL]) || 0;
        var cantFinalBk = oldDespBk + dispNowBk;
        if (cantFinalBk < cantSolBk) { hayBackorderPostDesp = true; }
      } else {
        var estBk = String(ped.datos[bk][COL.ESTADO] || '').trim();
        if (estBk === EST.BACKORDER) { hayBackorderPostDesp = true; }
        // Ítems Preparado/Parcial no despachados con prevDesp=0 serán reseteados a Backorder
        if (oldDespBk === 0 && (estBk === EST.PREPARADO || estBk === EST.PREP_PARCIAL)) { hayBackorderPostDesp = true; }
      }
    }
    var hayBackorder = hayBackorderPostDesp; // alias para Logger y plain text

    var etiquetaEnvio = hayDespPrevio ? '2\xba env\xedo' : (hayBackorderPostDesp ? '1\xba env\xedo' : 'Despacho');

    var introReseller;
    if (hayDespPrevio) {
      introReseller =
        "<p style='font-size:14px;color:#444;margin:0 0 10px'>" +
          "2\xba env\xedo de tu pedido <strong style='color:#00a3e0'>" + numero + "</strong>: " +
          "despachamos el stock que estaba pendiente. Adjuntamos la Nota de Entrega correspondiente." +
        "</p>" +
        (hayBackorderPostDesp
          ? "<div style='background:#fff8e1;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px;margin-bottom:14px'>" +
              "<p style='margin:0;font-size:12px;color:#7a5800'>A\xfan quedan \xedtems en <strong>backorder</strong>. Los despacharemos en cuanto llegue el stock.</p>" +
            "</div>"
          : "<div style='background:#e8f5e9;border-left:4px solid #43a047;padding:10px 14px;border-radius:4px;margin-bottom:14px'>" +
              "<p style='margin:0;font-size:12px;color:#1b5e20'>Con este env\xedo tu pedido queda <strong>completo</strong>.</p>" +
            "</div>");
    } else if (hayBackorderPostDesp) {
      introReseller =
        "<p style='font-size:14px;color:#444;margin:0 0 10px'>" +
          "1\xba env\xedo de tu pedido <strong style='color:#00a3e0'>" + numero + "</strong>: " +
          "despachamos los \xedtems disponibles. Adjuntamos la Nota de Entrega correspondiente." +
        "</p>" +
        "<div style='background:#fff8e1;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px;margin-bottom:14px'>" +
          "<p style='margin:0;font-size:12px;color:#7a5800'>" +
            "Los \xedtems en <strong>backorder</strong> ser\xe1n despachados en un 2\xba env\xedo " +
            "cuando el stock est\xe9 disponible." +
          "</p>" +
        "</div>";
    } else {
      introReseller =
        "<p style='font-size:14px;color:#444;margin:0 0 16px'>" +
          "Tu pedido <strong style='color:#00a3e0'>" + numero + "</strong> fue <strong>despachado</strong>. " +
          "Adjuntamos la Nota de Entrega correspondiente." +
        "</p>";
    }

    var cuerpoReseller = introReseller +
      tablaHtml + totalHtml + chipsHtml + rastrearBtn + driveBtn + obsHtml + confirmEntregaHtml;

    // ── Sección para administración (incluida al pie del mismo email) ─
    // Forma de pago: OT con DJI aprobado → sin costo; OT sin aprobar → 30 días; PR → campo PAGO
    var _esDJIAprobado = _esNumeroOT(numero) && ped.obs.indexOf('DJI ✓ Aprobado') !== -1;
    var formaPago = _esNumeroOT(numero)
      ? (_esDJIAprobado ? 'Sin costo — DJI aprobado' : 'A 30 días')
      : (ped.pago || 'N/E');
    var tituloAdmin = hayDespPrevio
      ? "Facturar 2\xba env\xedo — Pedido " + numero + (hayBackorderPostDesp ? " <span style='font-weight:400;color:#b8860b'>(a\xfan hay pendientes)</span>" : "")
      : (hayBackorderPostDesp
        ? "Facturar 1\xba env\xedo — Pedido " + numero + " <span style='font-weight:400;color:#b8860b'>(despacho parcial)</span>"
        : "Generar factura / ID de venta — Pedido " + numero);
    var adminHtml =
      "<div style='margin-top:28px;border-top:2px dashed #e0e3e8;padding-top:20px'>" +
        "<p style='margin:0 0 12px;font-size:10px;font-weight:700;text-transform:uppercase;" +
           "letter-spacing:.1em;color:#9ba5b4'>Para administración</p>" +
        "<div style='background:#fffbe6;border:1px solid #f9d95a;border-radius:8px;padding:14px 18px'>" +
          "<p style='margin:0 0 10px;font-size:13px;font-weight:700;color:#7a5800'>" +
            tituloAdmin +
          "</p>" +
          "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
            "<tr><td style='padding:3px 0;color:#888;width:130px'>N\xb0 de env\xedo</td>" +
                "<td style='font-weight:700;color:#00a3e0;font-size:13px'>" + Number(notaNumStr) + "\xaa factura \xb7 <span style='font-family:monospace;font-size:11px;color:#555'>" + notaEntrega + "</span></td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Reseller</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>" + ped.reseller + "</td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Fecha despacho</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>" + fecha + "</td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Forma de pago</td>" +
                "<td style='font-weight:600;color:" + (_esDJIAprobado ? '#1a9e4a' : '#1a1f2e') + "'>" + formaPago + "</td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Transportista</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>" + (transp || 'N/E') + "</td></tr>" +
            (costo > 0 ?
            "<tr><td style='padding:3px 0;color:#888'>Costo de env\xedo</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>$ " + _formatMoneda(costo) + "</td></tr>" : '') +
          "</table>" +
          "<div style='background:#fff;border:1px solid #f0e68c;border-radius:6px;padding:10px 14px'>" +
            "<p style='margin:0;font-size:13px;font-weight:800;color:#1a1f2e'>" +
              "Total: <span style='color:#00a3e0'>USD " + _formatMoneda(totalUSD) + "</span>" +
              "<span style='font-size:10px;font-weight:400;color:#9ba5b4;margin-left:8px'>no incluye impuestos</span>" +
            "</p>" +
          "</div>" +
        "</div>" +
      "</div>";

    var tituloEmail  = etiquetaEnvio + ' — ' + numero;
    var htmlCombinado = _wosPortalHead(tituloEmail) +
      "<p style='font-size:14px;color:#666;margin:0 0 22px;line-height:1.5'>Hola <strong>" + ped.reseller + "</strong>:</p>" +
      cuerpoReseller +
      adminHtml +
      _wosPortalFoot('Pedido ' + numero + ' · ' + ped.reseller + '.');

    var plainIntro = hayDespPrevio
      ? '2\xba env\xedo de tu pedido ' + numero + ': despachamos el stock pendiente.' + (hayBackorderPostDesp ? '\nA\xfan quedan \xedtems en backorder que despacharemos cuando llegue el stock.' : '\nCon este env\xedo tu pedido queda completo.')
      : (hayBackorderPostDesp
        ? '1\xba env\xedo de tu pedido ' + numero + ': despachamos los \xedtems disponibles.\nLos \xedtems en backorder llegar\xe1n en un 2\xba env\xedo cuando el stock est\xe9 disponible.'
        : 'Tu pedido ' + numero + ' fue despachado.');

    var plainAdminTitulo = hayDespPrevio
      ? 'Facturar 2\xba env\xedo' + (hayBackorderPostDesp ? ' (a\xfan hay pendientes)' : '') + ' — Pedido '
      : (hayBackorderPostDesp ? 'Facturar 1\xba env\xedo (despacho parcial) — Pedido ' : 'Generar factura / ID de venta — Pedido ');
    var plainCombinado =
      'Hola ' + ped.reseller + ',\n\n' +
      plainIntro +
      (transp ? '\nTransportista: ' + transp : '') +
      (track  ? '\nC\xf3digo de seguimiento: ' + track : '') + '\n\n' +
      'Adjuntamos la Nota de Entrega.' +
      (pdfUrl ? '\nVer en Drive: ' + pdfUrl : '') +
      '\nCuando recib\xe1s el paquete, respond\xe9 este correo con "Recibido" para confirmar la entrega.\n\n' +
      '---\nPARA ADMINISTRACIÓN\n' +
      plainAdminTitulo + numero + '\n' +
      'N\xb0 de env\xedo: ' + Number(notaNumStr) + '\xaa factura (' + notaEntrega + ')\n' +
      'Reseller: ' + ped.reseller + '\n' +
      'Fecha: ' + fecha + '\n' +
      'Forma de pago: ' + formaPago + '\n' +
      (costo > 0 ? 'Costo de env\xedo: $ ' + _formatMoneda(costo) + '\n' : '') +
      'Total: USD ' + _formatMoneda(totalUSD);

    // ── Un solo reply al hilo original, facturación en CC ─────
    var replyOpts = {
      htmlBody:    htmlCombinado,
      name:        'BIDCOMAGRO · Portal Resellers',
      replyTo:     EMAIL_SOPORTE,
      cc:          EMAIL_FACTURACION
    };
    if (pdfBlob) replyOpts.attachments = [pdfBlob];
    try {
      var despThread = GmailApp.getThreadById(ped.threadId);
      despThread.replyAll(plainCombinado, replyOpts);
    } catch(eThread) {
      Logger.log('WOS_despacharCompleto: thread no disponible (' + ped.threadId + '), enviando email nuevo → ' + email);
      GmailApp.sendEmail(email, tituloEmail + ' — Pedido ' + numero, plainCombinado, replyOpts);
    }

    // ── Estado final por ítem ────────────────────────────────────
    // Si cantFinal >= cantSol → Entregado_Cerrado (despacho completo del ítem)
    // Si cantFinal <  cantSol → Backorder (quedan unidades pendientes para 2º envío)
    // Items sin dispNow → se dejan intactos
    var esRetiro   = String(ped.envio || '').toLowerCase().indexOf('retiro') >= 0;
    var estadoDesp = esRetiro ? EST.LISTO_RETIRO : EST.ENTREGADO;
    var ahora2     = new Date();
    for (var sf = 1; sf < ped.datos.length; sf++) {
      if (String(ped.datos[sf][COL.NUMERO] || '').trim() !== numero) continue;
      var dispNowSf  = (despMap[sf + 1] !== undefined) ? (Number(despMap[sf + 1]) || 0) : 0;
      if (dispNowSf > 0) {
        var cantSolSf   = Number(ped.datos[sf][COL.CANT_SOL])  || 0;
        var oldDespSf   = Number(ped.datos[sf][COL.CANT_DESP]) || 0;
        var cantFinalSf = oldDespSf + dispNowSf;
        var rSf = ped.hoja.getRange(sf + 1, COL.ESTADO + 1);
        rSf.clearDataValidations();
        rSf.setValue(cantFinalSf >= cantSolSf ? estadoDesp : EST.BACKORDER);
        ped.hoja.getRange(sf + 1, COL.FECHA_ESTADO + 1).setValue(ahora2);
      } else {
        // Ítem no despachado: si estaba Preparado/Parcial y nunca tuvo despacho previo → volver a Backorder
        var estActualSf = String(ped.datos[sf][COL.ESTADO] || '').trim();
        var prevDespSf  = Number(ped.datos[sf][COL.CANT_DESP]) || 0;
        if (prevDespSf === 0 && (estActualSf === EST.PREPARADO || estActualSf === EST.PREP_PARCIAL)) {
          var rSfRst = ped.hoja.getRange(sf + 1, COL.ESTADO + 1);
          rSfRst.clearDataValidations();
          rSfRst.setValue(EST.BACKORDER);
          ped.hoja.getRange(sf + 1, COL.FECHA_ESTADO + 1).setValue(ahora2);
        }
        // Ítems ya despachados parcialmente (prevDesp>0) o en Backorder/Cancelado → intactos
      }
    }
    SpreadsheetApp.flush();

    // Actualizar HUB PRO OT: E: en REPUESTOS (col Q) + estado si completo
    if (_esNumeroOT(numero)) {
      try {
        var otNum = numero.replace(/^OT-/, '');
        // SKU → total despachado (previo + este despacho)
        var skuDespMap = {};
        for (var di = 1; di < ped.datos.length; di++) {
          if (String(ped.datos[di][COL.NUMERO] || '').trim() !== numero) continue;
          var dSku = String(ped.datos[di][COL.SKU] || '').trim().toUpperCase();
          if (!dSku) continue;
          var totalDi = (Number(ped.datos[di][COL.CANT_DESP]) || 0) +
                        ((despMap[di + 1] !== undefined) ? (Number(despMap[di + 1]) || 0) : 0);
          skuDespMap[dSku] = (skuDespMap[dSku] || 0) + totalDi;
        }
        var hubHoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Ordenes de trabajo');
        if (hubHoja) {
          var hubD = hubHoja.getDataRange().getValues();
          for (var hi = 1; hi < hubD.length; hi++) {
            if (String(hubD[hi][2] || '').trim() !== otNum) continue;
            // Actualizar E: en campo REPUESTOS (col Q = índice 16, rango = col 17)
            var repStr = String(hubD[hi][16] || '').trim();
            if (repStr) {
              var repPartes = repStr.split(' ; ');
              var repNuevas = [];
              for (var rpi = 0; rpi < repPartes.length; rpi++) {
                var rpCols = repPartes[rpi].split(' | ');
                var rpSku  = String(rpCols[0] || '').trim().toUpperCase();
                if (rpSku && skuDespMap[rpSku] !== undefined) {
                  var pedVal = String(rpCols[2] || '').split(' E:')[0].replace('P:', '') || '0';
                  rpCols[2]  = 'P:' + pedVal + ' E:' + skuDespMap[rpSku];
                }
                repNuevas.push(rpCols.join(' | '));
              }
              hubHoja.getRange(hi + 1, 17).setValue(repNuevas.join(' ; '));
            }
            // Estado: solo pasa a "Repuestos enviados" si el despacho fue completo
            if (!hayBackorder) hubHoja.getRange(hi + 1, 5).setValue('Repuestos enviados');
            break;
          }
        }
      } catch(eHub) { Logger.log('WOS_despacharCompleto HUB update: ' + eHub); }
    }

    var logExtra = (hayBackorder ? 'parcial+backorder' : 'completo') + ' transp=' + transp + ' track=' + track;
    _wosLogAccion('Despacho: ' + notaEntrega + ' → ' + estadoDesp, numero, ped.reseller, operario, logExtra);
    Logger.log('WOS_despacharCompleto OK: ' + numero + '-' + notaNumStr + ' | ' + estadoDesp + (hayBackorder ? '+backorder' : '') + ' | transp=' + transp + ' | track=' + track + ' | driveUrl=' + pdfUrl);
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_despacharCompleto ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}


// ─────────────────────────────────────────────────────────────
//  HITO 2 — Notificar Faltante
//  Responde en el hilo ancla informando el faltante y pidiendo
//  Opción A (esperar) o Opción B (despachar lo disponible).
//
//  faltantes: [{sku, desc, cantSol, cantDisp}]
// ─────────────────────────────────────────────────────────────
function WOS_notificarFaltante(numero, faltantes, operario) {
  try {
    operario = String(operario || '');
    var ped = _wosLeerPedido(numero);
    if (!ped.reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    if (!ped.threadId)  return { ok: false, error: 'Sin Thread_ID en col R. Verificar Hito 1 del Portal.' };

    var email = _wosGetEmailReseller(ped.reseller);
    if (!email) return { ok: false, error: 'Email no encontrado para reseller: ' + ped.reseller };

    // ── Tabla de faltantes ────────────────────────────────────
    var tablaFalt =
      "<table style='width:100%;border-collapse:collapse;font-size:12px;margin:16px 0'>" +
      "<thead><tr style='background:#fdecea'>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>SKU</th>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Solicitado</th>" +
        "<th style='padding:7px 10px;text-align:center;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Disponible</th>" +
      "</tr></thead><tbody>";
    for (var i = 0; i < faltantes.length; i++) {
      var f = faltantes[i];
      tablaFalt +=
        "<tr style='border-bottom:1px solid #f0f2f5'>" +
          "<td style='padding:8px 10px;font-family:monospace;font-weight:700;color:#e74c3c'>" + (f.sku  || '') + "</td>" +
          "<td style='padding:8px 10px;color:#1a202c'>"                                      + (f.desc || '') + "</td>" +
          "<td style='padding:8px 10px;text-align:center;font-weight:700'>"                  + (f.cantSol  || 0) + "</td>" +
          "<td style='padding:8px 10px;text-align:center;font-weight:700;color:" +
            (f.cantDisp > 0 ? '#1a9e4a' : '#e74c3c') + "'>"                                 + (f.cantDisp || 0) + "</td>" +
        "</tr>";
    }
    tablaFalt += "</tbody></table>";

    // ¿Hay algo disponible para despachar?
    var hayDisponible = false;
    for (var fd = 0; fd < faltantes.length; fd++) {
      if ((faltantes[fd].cantDisp || 0) > 0) { hayDisponible = true; break; }
    }

    // ── Cuerpo HTML ───────────────────────────────────────────
    var introDisp = hayDisponible
      ? "<p style='font-size:13px;color:#155724;background:#e8f7ee;border:1px solid #8fd4a8;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
          "&#128666; <strong>Vamos a despachar lo disponible</strong> a la brevedad.<br>" +
          "<span style='font-size:12px;color:#4a5568'>Solo necesitamos saber qué preferís hacer con los ítems faltantes:</span>" +
        "</p>"
      : "<p style='font-size:13px;color:#7c3c00;background:#fff3e0;border:1px solid #ffba7b;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
          "&#9888; <strong>No tenemos stock disponible</strong> para ninguno de los ítems faltantes.<br>" +
          "<span style='font-size:12px;color:#4a5568'>Por favor indicanos cómo querés proceder:</span>" +
        "</p>";

    var _wosUrl = '';
    try { _wosUrl = ScriptApp.getService().getUrl(); } catch(eUrl) { Logger.log('WOS URL error: ' + eUrl); }
    var _itemsStr = '';
    for (var fi = 0; fi < faltantes.length; fi++) {
      if (_itemsStr) _itemsStr += ',';
      _itemsStr += String(faltantes[fi].sku || '').toUpperCase().replace(/[,: ]/g, '') + ':' + (faltantes[fi].cantDisp || 0);
    }
    var _urlA = _wosUrl ? _wosUrl + '?page=resp_faltante&num=' + encodeURIComponent(numero) + '&op=A' : '';
    var _urlB = _wosUrl ? _wosUrl + '?page=resp_faltante&num=' + encodeURIComponent(numero) + '&op=B&items=' + encodeURIComponent(_itemsStr) : '';

    var _botonesHtml = _urlA
      ? "<div style='text-align:center;margin:22px 0 8px'>" +
          "<a href='" + _urlA + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#3730a3;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>Opci\xf3n A — Esperar el faltante</a>" +
          "<a href='" + _urlB + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#92400e;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>Opci\xf3n B — Cancelar el faltante</a>" +
        "</div>" +
        "<p style='font-size:11px;color:#888;text-align:center;margin:4px 0 0'>Si los botones no funcionan, respond\xe9 este correo con <strong>\"Opci\xf3n A\"</strong> o <strong>\"Opci\xf3n B\"</strong>.</p>"
      : "<p style='font-size:12px;color:#555;background:#f5f8fc;border-radius:6px;padding:10px 14px;line-height:1.6'>" +
          "Respond\xe9 este correo indicando tu elecci\xf3n: <strong>\"Opci\xf3n A\"</strong> o <strong>\"Opci\xf3n B\"</strong>." +
        "</p>";

    var opcionesHtml = introDisp +
      "<p style='font-size:13px;color:#1a1f2e;font-weight:700;margin:18px 0 10px'>\xbfQu\xe9 hacemos con el faltante?</p>" +
      "<div style='border:1px solid #c7d2fe;border-radius:8px;padding:14px 18px;margin-bottom:8px;background:#eef2ff'>" +
        "<p style='margin:0 0 4px;font-size:13px;color:#3730a3;font-weight:700'>OPCI\xd3N A — Esperar el faltante (segundo env\xedo)</p>" +
        "<p style='margin:0;font-size:12px;color:#4a5568'>Los \xedtems faltantes quedan pendientes. Cuando ingresen al stock los despachamos en un segundo env\xedo.</p>" +
      "</div>" +
      "<div style='border:1px solid #ffba7b;border-radius:8px;padding:14px 18px;margin-bottom:16px;background:#fff3e0'>" +
        "<p style='margin:0 0 4px;font-size:13px;color:#7c3c00;font-weight:700'>OPCI\xd3N B — Cancelar el faltante</p>" +
        "<p style='margin:0;font-size:12px;color:#4a5568'>Cancelamos definitivamente los \xedtems que no tenemos en stock. Solo se factura lo que se despacha ahora.</p>" +
      "</div>" +
      _botonesHtml;

    var htmlBody = _wosPortalHead('Faltante de stock — ' + numero) +
      "<p style='font-size:14px;color:#666;margin:0 0 22px;line-height:1.5'>Hola <strong>" + ped.reseller + "</strong>:</p>" +
      "<p style='font-size:13px;color:#555;margin:0 0 6px'>" +
        "Al procesar tu pedido <strong style='color:#00a3e0'>" + numero + "</strong> " +
        "detectamos un <strong>faltante de stock</strong>:" +
      "</p>" +
      tablaFalt +
      opcionesHtml +
      (ped.obs ? "<p style='font-size:11px;color:#888;margin-top:12px'><strong>Obs.:</strong> " + ped.obs + "</p>" : '') +
      _wosPortalFoot('Pedido ' + numero + ' · ' + ped.reseller + '.');

    // ── Plain text (incluye WOSDATA para que Hito 3 lo parsee) ──
    var wosDataJson = JSON.stringify({ numero: numero, faltantes: faltantes });
    var plainBody =
      'Hola ' + ped.reseller + ',\n\n' +
      'Al procesar tu pedido ' + numero + ' detectamos un faltante de stock.\n' +
      (hayDisponible ? 'Vamos a despachar lo disponible. ' : '') +
      'Por favor indicanos qué hacer con los ítems faltantes:\n\n' +
      'OPCI\xd3N A: Esperar el faltante y recibirlo en un segundo env\xedo cuando el stock est\xe9 disponible.\n' +
      'OPCI\xd3N B: Cancelar definitivamente los \xedtems faltantes.\n\n' +
      'Respond\xe9 este correo con tu elecci\xf3n (escrib\xed "Opci\xf3n A" u "Opci\xf3n B").\n\n' +
      '===WOSDATA===\n' + wosDataJson + '\n===ENDWOSDATA===';

    // ── Enviar como RESPUESTA en el hilo ancla (fallback: email nuevo) ──
    var faltOpts = {
      htmlBody: htmlBody,
      name:     'BIDCOMAGRO · Portal Resellers',
      replyTo:  EMAIL_SOPORTE
    };
    try {
      var faltThread = GmailApp.getThreadById(ped.threadId);
      faltThread.replyAll(plainBody, faltOpts);
    } catch(eThread) {
      Logger.log('WOS_notificarFaltante: thread no disponible (' + ped.threadId + '), enviando email nuevo → ' + email);
      GmailApp.sendEmail(email, 'Faltante de stock — Pedido ' + numero, plainBody, faltOpts);
    }

    // ── Cambiar estado por ítem: faltantes → En_Espera_Reseller, disponibles → Preparado ──
    var faltSet = {};
    for (var ff = 0; ff < faltantes.length; ff++) {
      faltSet[String(faltantes[ff].sku || '').trim().toUpperCase()] = true;
    }
    _wosSetEstadoPorSku(ped.hoja, ped.datos, numero, faltSet, EST.EN_ESPERA, EST.PREPARADO);
    SpreadsheetApp.flush();

    _wosLogAccion('Faltante notificado', numero, ped.reseller, operario, faltantes.length + ' items faltantes');
    Logger.log('WOS_notificarFaltante OK: ' + numero + ' → ' + EST.EN_ESPERA + ' | email → ' + email);
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_notificarFaltante ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}


// ─────────────────────────────────────────────────────────────
//  HITO 3 — Detector automático de respuestas del reseller
//  Asociar a un trigger de tiempo (cada 10 min).
//  Lee los hilos de pedidos en estado 'En_Espera_Reseller',
//  detecta "Opción A" o "Opción B" y actúa autónomamente.
// ─────────────────────────────────────────────────────────────
function WOS_detectarRespuestasResellers() {
  try {
    var hoja  = _wosHoja();
    var datos = hoja.getDataRange().getValues();

    // Agrupar filas por pedido, solo los que están en espera con threadId
    var enEspera = {}; // numero → { threadId, reseller, rowsIdx: [] }
    for (var i = 1; i < datos.length; i++) {
      var num    = String(datos[i][COL.NUMERO]    || '').trim();
      var estado = String(datos[i][COL.ESTADO]    || '').trim();
      var tId    = String(datos[i][COL.THREAD_ID] || '').trim();
      if (!num || estado !== EST.EN_ESPERA || !tId) continue;
      if (!enEspera[num]) {
        enEspera[num] = {
          threadId: tId,
          reseller: String(datos[i][COL.RESELLER] || ''),
          rows: []
        };
      }
      enEspera[num].rows.push({
        rowIdx:  i,                                           // 0-indexed en datos[]
        rowNum:  i + 1,                                       // 1-indexed para getRange
        sku:     String(datos[i][COL.SKU]      || '').toUpperCase(),
        cantSol: Number(datos[i][COL.CANT_SOL]) || 0
      });
    }

    var numeros   = Object.keys(enEspera);
    var procesados = 0;
    if (!numeros.length) {
      Logger.log('WOS_detectarRespuestasResellers: sin pedidos en espera');
      return { ok: true, procesados: 0 };
    }

    var dominioInterno = '@bidcom.com.ar';

    for (var n = 0; n < numeros.length; n++) {
      var numero = numeros[n];
      var info   = enEspera[numero];

      try {
        var thread   = GmailApp.getThreadById(info.threadId);
        var messages = thread.getMessages();
        if (!messages || messages.length < 2) continue; // solo existe nuestro mensaje, aún no hay respuesta

        // Recorrer mensajes: identificar el nuestro (WOSDATA) y la respuesta del reseller
        var wosDataStr    = null;
        var resellerPlain = null;
        var lastExtMsg    = null;

        for (var m = 0; m < messages.length; m++) {
          var msg    = messages[m];
          var from   = msg.getFrom().toLowerCase();
          var plain  = msg.getPlainBody();

          if (from.indexOf(dominioInterno) >= 0) {
            // Mensaje enviado por nosotros: buscar WOSDATA
            var matchW = plain.match(/===WOSDATA===\s*([\s\S]*?)\s*===ENDWOSDATA===/);
            if (matchW) wosDataStr = matchW[1].trim();
          } else {
            // Mensaje del reseller (o externo)
            lastExtMsg    = msg;
            resellerPlain = plain;
          }
        }

        if (!resellerPlain || !lastExtMsg) continue; // aún no hay respuesta externa

        // Detectar opción en el texto de la respuesta.
        // El reseller ve OPCIÓN A (esperar faltante) u OPCIÓN B (cancelar faltante).
        var rLow = resellerPlain.toLowerCase();
        var esA  = /opci[oó]n\s*a\b/.test(rLow);
        var esB  = /opci[oó]n\s*b\b/.test(rLow);

        if (!esA && !esB) {
          Logger.log('WOS_detectarRespuestasResellers: ' + numero + ' → respuesta no reconocida (se esperaba Opci\xf3n A u Opci\xf3n B), se ignora');
          continue;
        }

        // Parsear WOSDATA para obtener cantidades disponibles (necesario para B y C)
        var faltantesMap = {};
        if (wosDataStr) {
          try {
            var wosObj = JSON.parse(wosDataStr);
            var fArr   = wosObj.faltantes || [];
            for (var fi = 0; fi < fArr.length; fi++) {
              var skuUp = String(fArr[fi].sku || '').trim().toUpperCase();
              if (skuUp) faltantesMap[skuUp] = Number(fArr[fi].cantDisp) || 0;
            }
          } catch(eJ) { Logger.log('WOS_detectarRespuestasResellers JSON parse: ' + eJ); }
        }

        var confTitulo, confMensaje, nuevoEst;

        if (esA) {
          // Opción A (reseller) = esperar faltante en segundo envío
          // Solo los items En_Espera_Reseller pasan a Backorder; los Preparado quedan intactos
          nuevoEst   = EST.BACKORDER;
          confTitulo = 'Recibimos tu respuesta — Pedido ' + numero;
          confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar el faltante</strong>.<br>" +
            "Recibir\xe1s la Nota de Entrega por los \xedtems que se despachan ahora. El faltante llegar\xe1 en un segundo env\xedo cuando el stock est\xe9 disponible.";
          _wosSetEstadoFiltrado(hoja, datos, numero, EST.EN_ESPERA, nuevoEst);
          Logger.log('WOS_detectarRespuestasResellers: ' + numero + ' → OPCI\xd3N A → faltantes → ' + nuevoEst);

        } else {
          // Opción B (reseller) = cancelar faltante → despacho único de lo disponible
          // Solo los items En_Espera_Reseller se modifican: cantDisp>0 → Preparado, cantDisp=0 → Cancelado
          nuevoEst   = EST.PREPARADO;
          confTitulo = 'Recibimos tu respuesta — Pedido ' + numero;
          confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Cancelar el faltante</strong>.<br>" +
            "Estamos preparando el despacho de lo disponible. Recibir\xe1s la Nota de Entrega cuando sea despachado. Los \xedtems faltantes quedan cancelados.";
          var ahoraB = new Date();
          for (var ri = 0; ri < info.rows.length; ri++) {
            var row = info.rows[ri];
            if (String(datos[row.rowNum - 1][COL.ESTADO] || '').trim() !== EST.EN_ESPERA) continue;
            var cantDisp = (faltantesMap[row.sku] !== undefined) ? faltantesMap[row.sku] : 0;
            var cantSolOrig = Number(datos[row.rowNum - 1][COL.CANT_SOL]) || 0;
            var rEstB = hoja.getRange(row.rowNum, COL.ESTADO + 1);
            rEstB.clearDataValidations();
            if (cantDisp >= cantSolOrig && cantSolOrig > 0) {
              rEstB.setValue(EST.PREPARADO);
            } else if (cantDisp > 0) {
              hoja.getRange(row.rowNum, COL.CANT_SOL + 1).setValue(cantDisp);
              rEstB.setValue(EST.PREP_PARCIAL);
            } else {
              rEstB.setValue(EST.CANCELADO);
            }
            hoja.getRange(row.rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraB);
          }
          Logger.log('WOS_detectarRespuestasResellers: ' + numero + ' → OPCI\xd3N B → faltantes ajustados por item');
        }

        // ── Confirmación al reseller en el mismo hilo ─────────
        try {
          var confHtml = _wosPortalHead(confTitulo) +
            "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + info.reseller + "</strong>:</p>" +
            "<p style='font-size:13px;color:#555;line-height:1.6'>" + confMensaje + "</p>" +
            "<p style='font-size:12px;color:#888;margin-top:18px'>Consultas: " +
              "<a href='mailto:" + EMAIL_SOPORTE + "' style='color:#00a3e0'>" + EMAIL_SOPORTE + "</a></p>" +
            _wosPortalFoot('Pedido ' + numero + ' · ' + info.reseller + '.');
          var confPlain = confMensaje.replace(/<[^>]+>/g, '');
          thread.replyAll(confPlain, {
            htmlBody: confHtml,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  EMAIL_SOPORTE
          });
        } catch(eConf) {
          Logger.log('WOS_detectarRespuestasResellers confirmacion [' + numero + ']: ' + eConf);
        }

        thread.markRead();
        procesados++;

      } catch(eT) {
        Logger.log('WOS_detectarRespuestasResellers thread [' + numero + ']: ' + eT);
      }
    }

    // ── Confirmación de entrega: escanear pedidos Entregado_Cerrado ──
    var entregados = {};
    for (var ei = 1; ei < datos.length; ei++) {
      var eNum = String(datos[ei][COL.NUMERO]    || '').trim();
      var eEst = String(datos[ei][COL.ESTADO]    || '').trim();
      var eTid = String(datos[ei][COL.THREAD_ID] || '').trim();
      if (!eNum || eEst !== EST.ENTREGADO || !eTid) continue;
      if (!entregados[eNum]) {
        entregados[eNum] = { threadId: eTid, reseller: String(datos[ei][COL.RESELLER] || '') };
      }
    }

    var eNumeros = Object.keys(entregados);
    for (var en = 0; en < eNumeros.length; en++) {
      var eNumero = eNumeros[en];
      var eInfo   = entregados[eNumero];
      try {
        var eThread   = GmailApp.getThreadById(eInfo.threadId);
        var eMsgs     = eThread.getMessages();
        var eLastExt  = null;
        for (var em = 0; em < eMsgs.length; em++) {
          if (eMsgs[em].getFrom().toLowerCase().indexOf(dominioInterno) < 0) {
            eLastExt = eMsgs[em];
          }
        }
        if (!eLastExt) continue;

        var eBody = eLastExt.getPlainBody().toLowerCase();
        if (!/\brecib[ií]do?\b/.test(eBody)) continue;

        // Confirmación detectada → Entregado_Confirmado
        _wosSetEstado(hoja, datos, eNumero, EST.ENTREGADO_CONF);

        var confHtml = _wosPortalHead('Entrega confirmada — Pedido ' + eNumero) +
          "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + eInfo.reseller + "</strong>:</p>" +
          "<p style='font-size:13px;color:#555;line-height:1.6'>" +
            "Registramos la confirmación de recepción de tu pedido <strong style='color:#00a3e0'>" + eNumero + "</strong>. " +
            "El pedido queda cerrado. Gracias por trabajar con BIDCOMAGRO." +
          "</p>" +
          _wosPortalFoot('Pedido ' + eNumero + ' · ' + eInfo.reseller + '.');
        var confPlain = 'Hola ' + eInfo.reseller + ',\n\nRegistramos la confirmación de recepción del pedido ' + eNumero + '. El pedido queda cerrado. Gracias!';
        eThread.replyAll(confPlain, {
          htmlBody: confHtml,
          name:     'BIDCOMAGRO · Portal Resellers',
          replyTo:  EMAIL_SOPORTE
        });
        eThread.markRead();
        procesados++;
        Logger.log('WOS_detectarRespuestasResellers: ' + eNumero + ' → confirmación entrega → ' + EST.ENTREGADO_CONF);
      } catch(eET) {
        Logger.log('WOS_detectarRespuestasResellers entrega [' + eNumero + ']: ' + eET);
      }
    }

    SpreadsheetApp.flush();
    Logger.log('WOS_detectarRespuestasResellers: procesados=' + procesados + ' de ' + (numeros.length + eNumeros.length));
    return { ok: true, procesados: procesados };

  } catch(e) {
    Logger.log('WOS_detectarRespuestasResellers ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ─────────────────────────────────────────────────────────────
//  UTILIDAD — Recuperar Thread IDs perdidos
//  Ejecutar UNA VEZ desde el editor de Apps Script.
//  Para cada pedido sin Thread ID en col R, busca en Gmail
//  el hilo original por número de pedido y lo restaura.
//  Ver resultado en Ver → Registros de ejecución.
// ─────────────────────────────────────────────────────────────
function WOS_recuperarThreadIds() {
  var hojas = [_wosHoja(), _getHojaPedidosOT()].filter(Boolean);

  // Agrupar filas por pedido, solo los que no tienen threadId; registra la hoja de origen
  var sinThread = {}; // numero → { hoja, filas: [] }
  for (var h = 0; h < hojas.length; h++) {
    var hoja  = hojas[h];
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var num = String(datos[i][COL.NUMERO]    || '').trim();
      var tid = String(datos[i][COL.THREAD_ID] || '').trim();
      if (!num || tid) continue;
      if (!sinThread[num]) sinThread[num] = { hoja: hoja, filas: [] };
      sinThread[num].filas.push(i + 1); // filas 1-indexed
    }
  }

  var numeros = Object.keys(sinThread);
  Logger.log('WOS_recuperarThreadIds: ' + numeros.length + ' pedidos sin Thread ID → ' + numeros.join(', '));
  if (!numeros.length) return { ok: true, recuperados: 0, noEncontrados: [] };

  var recuperados = 0;
  var noEncontrados = [];

  for (var n = 0; n < numeros.length; n++) {
    var numero = numeros[n];
    var hoja   = sinThread[numero].hoja;
    var rows   = sinThread[numero].filas;

    // Buscar en Gmail: el Portal usa asunto "[PEDIDO] PR-XXXXX — ..."
    // Un solo número es suficiente porque los IDs son únicos
    var threads = [];
    try {
      threads = GmailApp.search('subject:"' + numero + '"', 0, 10);
    } catch(eS) {
      Logger.log('WOS_recuperarThreadIds: error buscando ' + numero + ' → ' + eS);
      noEncontrados.push(numero);
      continue;
    }

    if (!threads || threads.length === 0) {
      Logger.log('WOS_recuperarThreadIds: ' + numero + ' → sin resultados en Gmail');
      noEncontrados.push(numero);
      continue;
    }

    // Si hay más de un hilo (raro), usar el más antiguo (el del pedido original)
    var threadId = threads[threads.length - 1].getId();

    for (var r = 0; r < rows.length; r++) {
      hoja.getRange(rows[r], COL.THREAD_ID + 1).setValue(threadId);
    }
    recuperados++;
    Logger.log('WOS_recuperarThreadIds: ' + numero + ' (' + rows.length + ' fila/s) → ' + threadId);
  }

  SpreadsheetApp.flush();
  Logger.log('WOS_recuperarThreadIds RESULTADO: recuperados=' + recuperados +
             ' | no encontrados=' + (noEncontrados.length ? noEncontrados.join(', ') : 'ninguno'));
  return { ok: true, recuperados: recuperados, noEncontrados: noEncontrados };
}

// Instala el trigger de tiempo (ejecutar UNA VEZ desde el editor)
function WOS_instalarTriggerDetector() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'WOS_detectarRespuestasResellers') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('WOS_detectarRespuestasResellers')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('Trigger instalado: WOS_detectarRespuestasResellers cada 10 min');
}


// ─────────────────────────────────────────────────────────────
//  Procesa la respuesta del reseller ingresada manualmente por
//  el operario (gestión telefónica). Replica la lógica del
//  detector automático pero añade "confirmado telefónicamente"
//  en el email de confirmación al reseller.
//
//  opcion:     'A' (esperar faltante) | 'B' (cancelar faltante)
//  cantidades: [{sku, cantDisp}] — requerido solo para opción B
// ─────────────────────────────────────────────────────────────
function WOS_procesarRespuestaManual(numero, opcion, cantidades, operario) {
  try {
    operario  = String(operario || '');
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    var datos = hoja.getDataRange().getValues();

    var threadId = '', reseller = '';
    var rows = [];
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      if (!reseller) {
        reseller = String(datos[i][COL.RESELLER]  || '');
        threadId = String(datos[i][COL.THREAD_ID] || '').trim();
      }
      rows.push({ rowNum: i + 1, sku: String(datos[i][COL.SKU] || '').toUpperCase() });
    }
    if (!reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };

    var cantMap = {};
    if (cantidades) {
      for (var c = 0; c < cantidades.length; c++) {
        cantMap[String(cantidades[c].sku || '').toUpperCase()] = Number(cantidades[c].cantDisp) || 0;
      }
    }

    var op = String(opcion || '').toUpperCase();
    var nuevoEst, confTitulo, confMensaje;

    if (op === 'A') {
      // Opción A = esperar faltante en segundo envío
      // Solo los items En_Espera_Reseller pasan a Backorder; los Preparado quedan intactos
      nuevoEst    = EST.BACKORDER;
      confTitulo  = 'Recibimos tu decisi\xf3n — Pedido ' + numero;
      confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar el faltante en un segundo env\xedo.</strong><br>" +
        "Recibir\xe1s la Nota de Entrega por los \xedtems que se despachan ahora. " +
        "El faltante llegar\xe1 en un segundo env\xedo cuando el stock est\xe9 disponible.";
      _wosSetEstadoFiltrado(hoja, datos, numero, EST.EN_ESPERA, nuevoEst);

    } else {
      // Opción B = cancelar faltante → despacho único de lo disponible
      // Solo los items En_Espera_Reseller se modifican: cantDisp>0 → Preparado, cantDisp=0 → Cancelado
      nuevoEst    = EST.PREPARADO;
      confTitulo  = 'Recibimos tu decisi\xf3n — Pedido ' + numero;
      confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Cancelar el faltante.</strong><br>" +
        "Estamos preparando el despacho de lo disponible. Recibir\xe1s la Nota de Entrega cuando sea despachado. Los \xedtems faltantes quedan cancelados.";
      var ahoraM = new Date();
      for (var r = 0; r < rows.length; r++) {
        if (String(datos[rows[r].rowNum - 1][COL.ESTADO] || '').trim() !== EST.EN_ESPERA) continue;
        var cantDispM = (cantMap[rows[r].sku] !== undefined) ? cantMap[rows[r].sku] : 0;
        var cantSolOrigM = Number(datos[rows[r].rowNum - 1][COL.CANT_SOL]) || 0;
        var rEstM = hoja.getRange(rows[r].rowNum, COL.ESTADO + 1);
        rEstM.clearDataValidations();
        if (cantDispM >= cantSolOrigM && cantSolOrigM > 0) {
          rEstM.setValue(EST.PREPARADO);
        } else if (cantDispM > 0) {
          hoja.getRange(rows[r].rowNum, COL.CANT_CANCEL + 1).setValue(cantSolOrigM - cantDispM);
          rEstM.setValue(EST.PREP_PARCIAL);
        } else {
          hoja.getRange(rows[r].rowNum, COL.CANT_CANCEL + 1).setValue(cantSolOrigM);
          rEstM.setValue(EST.CANCELADO);
        }
        hoja.getRange(rows[r].rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraM);
      }
    }

    SpreadsheetApp.flush();

    if (threadId) {
      try {
        var confHtml = _wosPortalHead(confTitulo) +
          "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + "</strong>:</p>" +
          "<p style='font-size:13px;color:#555;line-height:1.6'>" + confMensaje + "</p>" +
          "<p style='font-size:12px;color:#888;margin-top:18px'>Consultas: " +
            "<a href='mailto:" + EMAIL_SOPORTE + "' style='color:#00a3e0'>" + EMAIL_SOPORTE + "</a></p>" +
          _wosPortalFoot('Pedido ' + numero + ' · ' + reseller + '.');
        var confPlain = confMensaje.replace(/<[^>]+>/g, '');
        GmailApp.getThreadById(threadId).replyAll(confPlain, {
          htmlBody: confHtml,
          name:     'BIDCOMAGRO · Portal Resellers',
          replyTo:  EMAIL_SOPORTE
        });
      } catch(eM) { Logger.log('WOS_procesarRespuestaManual email [' + numero + ']: ' + eM); }
    }

    _wosLogAccion('Respuesta tel\xe9fonica: Opci\xf3n ' + op + ' → ' + nuevoEst, numero, reseller, operario, '');
    Logger.log('WOS_procesarRespustaManual OK: ' + numero + ' op=' + op + ' → ' + nuevoEst);
    return { ok: true, estado: nuevoEst };
  } catch(e) {
    Logger.log('WOS_procesarRespuestaManual ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}


// Reemplazada por WOS_despacharCompleto — mantener como stub para no romper referencias externas
function WOS_cerrarBulto(numero) {
  Logger.log('WOS_cerrarBulto: función deprecada, usar WOS_despacharCompleto para ' + numero);
  return { ok: false, error: 'Función deprecada. Usar WOS_despacharCompleto.' };
}


// Reemplazada por WOS_despacharCompleto — mantener como stub para no romper referencias externas
function WOS_despacharYCerrar(numero) {
  Logger.log('WOS_despacharYCerrar: función deprecada, usar WOS_despacharCompleto para ' + numero);
  return { ok: false, error: 'Función deprecada. Usar WOS_despacharCompleto.' };
}
