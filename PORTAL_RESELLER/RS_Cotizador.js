// ============================================================
// @version 1.11
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
    hoja.getRange(1, 1, 1, 11).setBackground('#3a9e3a').setFontColor('#fff').setFontWeight('bold');
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
    var cant = Math.floor(Number(it.cantidad));
    if (isNaN(cant) || cant < 1) cant = 1;   // cantidad entera >= 1

    // Precio de lista (PVP): del catálogo (priceMap está al 40%, se revierte a base) o el enviado.
    var precioLista;
    if (skuU && priceMap[skuU] !== undefined && priceMap[skuU] > 0) {
      precioLista = Math.round(priceMap[skuU] / 0.60 * 100) / 100;
    } else {
      precioLista = Number(it.precioLista) || 0;
    }
    if (isNaN(precioLista) || precioLista < 0) precioLista = 0;   // precio no negativo

    var desc = Number(it.descuento);
    if (isNaN(desc)) desc = _COT_DESC_DEFECTO;
    if (desc < 0)  desc = 0;
    if (desc > 99) desc = 99;   // tope de descuento por pieza: 99%

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

// Parseo tolerante de precio (acepta "USD 50", "1.234,56", "50", "50.00").
function _cotNum(s) {
  var t = String(s == null ? '' : s).replace(/[^0-9.,]/g, '');
  if (t.indexOf(',') !== -1 && t.indexOf('.') !== -1) t = t.replace(/\./g, '').replace(',', '.');
  else if (t.indexOf(',') !== -1) t = t.replace(',', '.');
  var n = parseFloat(t);
  return isNaN(n) ? 0 : n;
}

// Catálogo de mano de obra recomendada (misma hoja que usa HUB PRO: Precios_mano_obra).
function obtenerManoObraCotizador() {
  try {
    var out = [];
    var d = getSheetValues(SCHEMA.SHEETS.PRECIOS_MANO_OBRA);
    for (var i = 1; i < d.length; i++) {
      var cod = String(d[i][0] || '').trim();
      var dsc = String(d[i][1] || '').trim();
      if (!cod && !dsc) continue;
      if (cod && dsc) out.push({ codigo: cod, descripcion: dsc, precio: _cotNum(d[i][2]) });
    }
    return { ok: true, items: out };
  } catch(e) {
    Logger.log('obtenerManoObraCotizador: ' + e);
    return { ok: false, items: [] };
  }
}

// Normaliza líneas de mano de obra: precio editable (recomendado por defecto), cantidad >= 1.
function _cotNormalizarMO(moItems) {
  var out = [], total = 0;
  moItems = moItems || [];
  for (var i = 0; i < moItems.length; i++) {
    var it   = moItems[i] || {};
    var desc = String(it.descripcion || '').trim();
    var cod  = String(it.codigo || '').trim();
    if (!desc && !cod) continue;
    var precio = Math.round(_cotNum(it.precio) * 100) / 100; // mano de obra sin cantidad
    total += precio;
    out.push({ codigo: cod, descripcion: desc || cod, cantidad: 1, precio: precio, subtotal: precio });
  }
  return { items: out, total: Math.round(total * 100) / 100 };
}

// ── Plantillas de Cotización (predefinidas) ───────────────────────────
// Hoja normalizada 1 fila por línea de ítem (repuesto o mano de obra) — mismo criterio que
// KITS (RS_Pedidos.js): legible/editable a mano si hiciera falta, sin JSON.
// Reseller vacío = plantilla general (la arma BIDCOM desde el LAUNCHER); con nombre = privada
// de ese reseller (la armó y la ve solo él, desde "Guardar como plantilla" en el Cotizador).
function _asegurarHojaPlantillasCotizador() {
  var ss = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.PLANTILLAS_COTIZADOR);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.PLANTILLAS_COTIZADOR);
    hoja.appendRow(['Plantilla', 'Reseller', 'Tipo', 'SKU_Codigo', 'Descripción', 'Cantidad', 'Fecha']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 7).setBackground('#3a9e3a').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(1, 200);
    hoja.setColumnWidth(5, 260);
  }
  return hoja;
}

// params: { token, reseller } — devuelve { generales: [...], mias: [...] }, cada plantilla
// como { nombre, items: [{sku,descripcion,cantidad}], mo: [{sku,descripcion,cantidad}] }.
function RS_listarPlantillasCotizador(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.', generales: [], mias: [] };
    var reseller = _ses.nombre;

    _asegurarHojaPlantillasCotizador();
    var d = getSheetValues(SCHEMA.SHEETS.PLANTILLAS_COTIZADOR);
    var mapa = {}; // "nombre|reseller" → { nombre, reseller, items, mo }
    for (var i = 1; i < d.length; i++) {
      var nombre = String(d[i][0] || '').trim();
      if (!nombre) continue;
      var resFila = String(d[i][1] || '').trim();
      var tipo    = String(d[i][2] || '').trim().toUpperCase();
      var key = nombre + '|' + resFila;
      if (!mapa[key]) mapa[key] = { nombre: nombre, reseller: resFila, items: [], mo: [] };
      var linea = { sku: String(d[i][3] || '').trim(), descripcion: String(d[i][4] || '').trim(), cantidad: Number(d[i][5]) || 1 };
      if (tipo === 'MANO_OBRA') mapa[key].mo.push(linea); else mapa[key].items.push(linea);
    }
    var generales = [], mias = [];
    for (var k in mapa) {
      var p = mapa[k];
      if (!p.reseller) generales.push(p);
      else if (p.reseller === reseller) mias.push(p);
    }
    var porNombre = function(a, b) { return a.nombre.localeCompare(b.nombre); };
    generales.sort(porNombre); mias.sort(porNombre);
    return { ok: true, generales: generales, mias: mias };
  } catch(e) {
    Logger.log('RS_listarPlantillasCotizador: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.', generales: [], mias: [] };
  }
}

// params: { token, reseller, nombre, items: [{sku,descripcion,cantidad}], mo: [{sku,descripcion}] }
// SIEMPRE guarda con el reseller de la sesión — un reseller nunca puede crear una plantilla
// general (Reseller vacío) desde acá, esas solo se cargan desde el panel del LAUNCHER.
function RS_guardarPlantillaCotizador(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.' };
    var reseller = _ses.nombre;
    var nombre = String(params.nombre || '').trim();
    if (!nombre) return { ok: false, error: 'Poné un nombre para la plantilla.' };
    var items = params.items || [];
    var mo    = params.mo || [];
    if (!items.length && !mo.length) return { ok: false, error: 'El presupuesto está vacío.' };

    var hoja = _asegurarHojaPlantillasCotizador();
    var d = hoja.getDataRange().getValues();
    var nombreLower = nombre.toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim() === reseller && String(d[i][0] || '').trim().toLowerCase() === nombreLower) {
        return { ok: false, error: 'Ya tenés una plantilla con ese nombre — borrala primero si querés reemplazarla.' };
      }
    }
    var ahora = new Date();
    var filas = [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      var sku = String(it.sku || '').trim();
      var descIt = String(it.descripcion || '').trim();
      if (!sku && !descIt) continue;
      var cant = Math.floor(Number(it.cantidad)) || 1;
      filas.push([nombre, reseller, 'REPUESTO', sku, descIt, cant, ahora]);
    }
    for (var k = 0; k < mo.length; k++) {
      var m = mo[k];
      var cod = String(m.sku || m.codigo || '').trim();
      var dsc = String(m.descripcion || '').trim();
      if (!cod && !dsc) continue;
      filas.push([nombre, reseller, 'MANO_OBRA', cod, dsc, 1, ahora]);
    }
    if (!filas.length) return { ok: false, error: 'El presupuesto está vacío.' };
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.PLANTILLAS_COTIZADOR);
    return { ok: true };
  } catch(e) {
    Logger.log('RS_guardarPlantillaCotizador: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}

// params: { token, reseller, nombre } — solo borra si la plantilla es de ESE reseller
// (nunca una general ni la de otro reseller).
function RS_eliminarPlantillaCotizador(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.' };
    var reseller = _ses.nombre;
    var nombre = String(params.nombre || '').trim();
    if (!nombre) return { ok: false, error: 'Falta el nombre de la plantilla.' };

    var hoja = _asegurarHojaPlantillasCotizador();
    var d = hoja.getDataRange().getValues();
    var nombreLower = nombre.toLowerCase();
    var filasBorrar = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim() === reseller && String(d[i][0] || '').trim().toLowerCase() === nombreLower) {
        filasBorrar.push(i + 1);
      }
    }
    if (!filasBorrar.length) return { ok: false, error: 'No se encontró esa plantilla.' };
    filasBorrar.sort(function(a, b) { return b - a; });
    for (var f = 0; f < filasBorrar.length; f++) hoja.deleteRow(filasBorrar[f]);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.PLANTILLAS_COTIZADOR);
    return { ok: true };
  } catch(e) {
    Logger.log('RS_eliminarPlantillaCotizador: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}

// ── Confirmar cotización — genera PDF, guarda historial, manda email ──
// params: { reseller, cliente, clienteEmail, observaciones,
//           items: [{ sku, descripcion, cantidad, precioLista, descuento, modelos }] }
function RS_confirmarCotizacion(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    params = params || {};
    var _s = _sesionResolver(params.token, params.reseller);
    if (!_s) return { ok: false, error: 'Sesión inválida o expirada. Volvé a ingresar.' };
    var reseller     = _s.nombre;
    var cliente      = String(params.cliente      || '').trim();
    var clienteEmail = String(params.clienteEmail || '').trim();
    var obs          = String(params.observaciones|| '').trim();
    var itemsIn      = params.items || [];

    var moIn = params.manoObra || [];
    if (!reseller)                        return { ok: false, error: 'Falta el reseller.' };
    if (!itemsIn.length && !moIn.length)  return { ok: false, error: 'El presupuesto está vacío.' };
    if (!cliente)                         return { ok: false, error: 'Ingresá el nombre del cliente.' };

    _asegurarHojaCotizaciones();

    // Precios de catálogo (enforcement) + normalización de repuestos
    var priceMap = _buildPriceMap();
    var norm  = _cotNormalizarItems(itemsIn, priceMap);
    var items = norm.items;
    // Mano de obra
    var moNorm    = _cotNormalizarMO(moIn);
    var manoObra  = moNorm.items;
    var total     = Math.round((norm.total + moNorm.total) * 100) / 100;

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
    var pdfUrl = _generarPdfCotizacion(numero, resellerMeta, cliente, items, manoObra, total, obs);

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
        items.length + manoObra.length,
        JSON.stringify({ repuestos: items, manoObra: manoObra }),
        total > 0 ? total : '',
        pdfUrl || '',
        obs
      ]);
      invalidateSheetValues(SCHEMA.SHEETS.COTIZACIONES);
    }

    // Email (al reseller + cliente si cargó email)
    _enviarEmailCotizacion(numero, reseller, emailReseller, cliente, clienteEmail, items, manoObra, total, pdfUrl, obs);

    return { ok: true, numero: numero, pdfUrl: pdfUrl || '', total: total };

  } catch(e) {
    Logger.log('RS_confirmarCotizacion ERROR: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}

// ── PDF de la cotización (hoja temporal → DriveApp.getAs) ─────────
function _generarPdfCotizacion(numero, resellerMeta, cliente, items, manoObra, total, obs) {
  var tempSs = null;
  try {
    var meta     = resellerMeta || { nombre: '', direccion: '', telefono: '' };
    var fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    tempSs = SpreadsheetApp.create('TEMP_COT_' + numero);
    var sheet = tempSs.getActiveSheet();
    sheet.setName('Cotizacion');

    var ri = 1;

    // 1. Membrete — logo "BidcomAgro" (Bidcom negro + Agro verde) + datos del documento a la derecha
    var logoRange = sheet.getRange(ri, 1, 2, 4).merge();
    logoRange.setBackground('#ffffff').setVerticalAlignment('middle').setHorizontalAlignment('left');
    var stB = SpreadsheetApp.newTextStyle().setForegroundColor('#111111').setBold(true).setFontSize(24).build();
    var stA = SpreadsheetApp.newTextStyle().setForegroundColor('#3a9e3a').setBold(true).setFontSize(24).build();
    var logoRt = SpreadsheetApp.newRichTextValue()
      .setText('BidcomAgro')
      .setTextStyle(0, 6,  stB)   // "Bidcom" → negro
      .setTextStyle(6, 10, stA)   // "Agro"   → verde
      .build();
    logoRange.setRichTextValue(logoRt);
    sheet.getRange(ri + 2, 1, 1, 4).merge().setValue('Servicio Técnico Oficial · Repuestos originales DJI · Mano de obra calificada')
      .setFontSize(9).setFontColor('#7a828a').setHorizontalAlignment('left').setVerticalAlignment('top');
    // Bloque del documento (derecha) — texto sobre blanco, sin relleno
    sheet.getRange(ri, 5, 1, 3).merge().setValue('PRESUPUESTO')
      .setFontSize(10).setFontWeight('bold').setFontColor('#7a828a').setHorizontalAlignment('right').setVerticalAlignment('bottom');
    sheet.getRange(ri + 1, 5, 1, 3).merge().setValue('N\xba ' + numero)
      .setFontSize(13).setFontWeight('bold').setFontColor('#2d3436').setHorizontalAlignment('right').setVerticalAlignment('middle');
    sheet.getRange(ri + 2, 5, 1, 3).merge().setValue('Fecha: ' + fechaStr)
      .setFontSize(9).setFontColor('#7a828a').setHorizontalAlignment('right').setVerticalAlignment('top');
    sheet.setRowHeight(ri, 26); sheet.setRowHeight(ri + 1, 16); sheet.setRowHeight(ri + 2, 18);
    // Regla inferior del membrete — acento verde de marca
    sheet.getRange(ri + 2, 1, 1, 7)
      .setBorder(null, null, true, null, null, null, '#3a9e3a', SpreadsheetApp.BorderStyle.SOLID_THICK);
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

    var repInicio = ri;
    var headers = ['SKU', 'Descripción', 'Cant.', 'PVP USD', 'Desc.', 'Precio USD', 'Subtotal USD'];
    sheet.getRange(ri, 1, 1, 7).setValues([headers])
      .setBackground('#2d3436').setFontColor('#ffffff').setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22); ri++;

    for (var i = 0; i < items.length; i++) {
      var it    = items[i];
      var rowBg = (i % 2 === 0) ? '#ffffff' : '#f6f7f8';
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
      sheet.getRange(ri, 1, 1, 2).setWrap(true);   // SKU + Descripción: no cortar
      sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#2d3436');
      sheet.getRange(ri, 3).setHorizontalAlignment('center');
      sheet.getRange(ri, 5).setHorizontalAlignment('center');
      sheet.getRange(ri, 4, 1, 4).setHorizontalAlignment('right');
      sheet.setRowHeight(ri, 20); ri++;
    }
    sheet.getRange(repInicio, 1, ri - repInicio, 7)
      .setBorder(true, true, true, true, true, true, '#dfe3e6', SpreadsheetApp.BorderStyle.SOLID);

    // 3b. Mano de obra
    if (manoObra && manoObra.length) {
      sheet.setRowHeight(ri, 8); sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff'); ri++;
      var moInicio = ri;
      sheet.getRange(ri, 1, 1, 6).merge().setValue('MANO DE OBRA')
        .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff').setBackground('#2d3436').setVerticalAlignment('middle');
      sheet.getRange(ri, 7).setValue('Importe USD').setFontSize(9).setFontWeight('bold').setFontColor('#ffffff').setBackground('#2d3436').setHorizontalAlignment('right');
      sheet.setRowHeight(ri, 22); ri++;
      for (var mi = 0; mi < manoObra.length; mi++) {
        var mo  = manoObra[mi];
        var mBg = (mi % 2 === 0) ? '#ffffff' : '#f6f7f8';
        sheet.getRange(ri, 1, 1, 6).merge().setValue(mo.descripcion || '—').setFontSize(9).setBackground(mBg).setVerticalAlignment('middle').setWrap(true);
        sheet.getRange(ri, 7).setValue(mo.precio > 0 ? _fmtUsd(mo.precio) : '—').setFontSize(9).setFontWeight('bold').setBackground(mBg).setHorizontalAlignment('right');
        sheet.setRowHeight(ri, 20); ri++;
      }
      sheet.getRange(moInicio, 1, ri - moInicio, 7)
        .setBorder(true, true, true, true, true, true, '#dfe3e6', SpreadsheetApp.BorderStyle.SOLID);
    }

    // 4. Total
    if (total > 0) {
      sheet.getRange(ri, 1, 1, 6).merge().setValue('TOTAL (no incluye impuestos)')
        .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff').setBackground('#2d3436')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 7).setValue(_fmtUsd(total))
        .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff').setBackground('#3a9e3a')
        .setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.setRowHeight(ri, 24); ri++;
    }

    if (obs) {
      sheet.setRowHeight(ri, 8); ri++;
      sheet.getRange(ri, 1, 1, 7).merge().setValue('Observaciones: ' + obs)
        .setFontSize(9).setFontColor('#5e6778').setWrap(true).setVerticalAlignment('top');
      sheet.setRowHeight(ri, 34);
      ri++;
    }

    // Pie legal — quién emite y alcance de la cotización
    sheet.setRowHeight(ri, 10); ri++;
    var legal = 'Presupuesto emitido por ' + (meta.nombre || '—') + ' para ' + (cliente || '—') + '. ' +
      'Los valores corresponden únicamente a los daños observados y/o descritos por el cliente. ' +
      'Si durante la reparación se detectan daños adicionales no contemplados en esta cotización, ' +
      'se emitirá un nuevo presupuesto para su aprobación antes de continuar. ' +
      'Precios en USD, no incluyen impuestos. Válido salvo variación de precios.';
    sheet.getRange(ri, 1, 1, 7).merge().setValue(legal)
      .setFontSize(8).setFontColor('#7a828a').setWrap(true)
      .setHorizontalAlignment('left').setVerticalAlignment('top');
    sheet.setRowHeight(ri, 56);
    var lastRow = ri;

    // Anchos de columna (SKU más ancho para no cortar el código)
    sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 205); sheet.setColumnWidth(3, 45);
    sheet.setColumnWidth(4, 90);  sheet.setColumnWidth(5, 50);  sheet.setColumnWidth(6, 90);
    sheet.setColumnWidth(7, 100);

    // Limpieza: ocultar gridlines de fondo + recortar filas/columnas vacías
    sheet.setHiddenGridlines(true);
    if (sheet.getMaxRows()    > lastRow) sheet.deleteRows(lastRow + 1, sheet.getMaxRows() - lastRow);
    if (sheet.getMaxColumns() > 7)       sheet.deleteColumns(8, sheet.getMaxColumns() - 7);

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
function _enviarEmailCotizacion(numero, reseller, emailReseller, cliente, clienteEmail, items, manoObra, total, pdfUrl, obs) {
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
          '<td style="padding:7px 10px;font-size:11px;color:#3a9e3a;font-weight:700;border-bottom:1px solid #eef2f6">' + (it.sku || '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;color:#333;border-bottom:1px solid #eef2f6">' + (it.descripcion || '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:center;border-bottom:1px solid #eef2f6">' + it.cantidad + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:center;border-bottom:1px solid #eef2f6">' + (it.descuento || 0) + '%</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;border-bottom:1px solid #eef2f6">' + (it.precioFinal > 0 ? _fmtUsd(it.precioFinal) : '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;font-weight:700;border-bottom:1px solid #eef2f6">' + (it.subtotal > 0 ? _fmtUsd(it.subtotal) : '—') + '</td>' +
        '</tr>';
    }

    var pdfBloque = pdfUrl
      ? '<div style="margin-top:16px;text-align:center"><a href="' + pdfUrl + '" target="_blank" style="display:inline-block;padding:11px 26px;background:#3a9e3a;color:#fff;border-radius:7px;text-decoration:none;font-size:13px;font-weight:700"> Descargar presupuesto (PDF)</a></div>'
      : '';

    var repTabla = '';
    if (items.length) {
      repTabla =
        '<p style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 6px">Repuestos</p>' +
        '<table style="width:100%;border-collapse:collapse;border:1px solid #e8e8e8;font-family:Arial,sans-serif;margin-bottom:16px">' +
          '<thead><tr style="background:#3a9e3a">' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:left">SKU</th>' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:left">Descripción</th>' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:center">Cant.</th>' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:center">Desc.</th>' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:right">Precio</th>' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:right">Subtotal</th>' +
          '</tr></thead><tbody>' + filas + '</tbody></table>';
    }

    var moFilas = '';
    for (var m = 0; m < (manoObra || []).length; m++) {
      var mo = manoObra[m];
      moFilas +=
        '<tr style="background:' + (m % 2 === 0 ? '#ffffff' : '#f7f8fa') + '">' +
          '<td style="padding:7px 10px;font-size:12px;color:#333;border-bottom:1px solid #eef2f6">' + (mo.descripcion || '—') + '</td>' +
          '<td style="padding:7px 10px;font-size:12px;text-align:right;font-weight:700;border-bottom:1px solid #eef2f6">' + (mo.precio > 0 ? _fmtUsd(mo.precio) : '—') + '</td>' +
        '</tr>';
    }
    var moTabla = moFilas
      ? '<p style="font-size:11px;font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 6px">Mano de obra</p>' +
        '<table style="width:100%;border-collapse:collapse;border:1px solid #e8e8e8;font-family:Arial,sans-serif;margin-bottom:16px">' +
          '<thead><tr style="background:#3a9e3a">' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:left">Descripción</th>' +
            '<th style="padding:7px 10px;font-size:10px;color:#fff;text-align:right">Importe</th>' +
          '</tr></thead><tbody>' + moFilas + '</tbody></table>'
      : '';

    var cuerpo =
      '<p style="font-size:14px;color:#444;margin:0 0 6px">Presupuesto <strong style="color:#3a9e3a">' + numero + '</strong></p>' +
      '<p style="font-size:13px;color:#555;margin:0 0 18px">Cliente: <strong>' + (cliente || '—') + '</strong></p>' +
      repTabla + moTabla +
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
function obtenerHistorialCotizaciones(token, reseller) {
  try {
    var _s = _sesionResolver(token, reseller);
    if (!_s) return { ok: false, historial: [] };
    _asegurarHojaCotizaciones();
    var datos = getSheetValues(SCHEMA.SHEETS.COTIZACIONES);
    var rLow  = String(_s.nombre || '').trim().toLowerCase();
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
        itemsJson: String(f[7] || ''),   // col 8: { repuestos:[...], manoObra:[...] } — para "Repetir"
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
