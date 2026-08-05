// ============================================================
// @version 1.9
//  PORTAL RESELLER BIDCOM — Repuestos, cotizaciones y catálogo
// ============================================================

function _getAccesoriosCache() {
  var CKEY = 'accesorios_v1';
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CKEY);
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    var sheet = SpreadsheetApp.openById(LISTA_PRECIOS_SS_ID).getSheetByName('ACCESORIOS');
    if (!sheet) return [];
    var rows = sheet.getDataRange().getValues();
    var items = [];
    for (var i = 1; i < rows.length; i++) {
      var sku = String(rows[i][0] || '').trim();
      if (!sku) continue;
      var pvRaw = rows[i][3];
      items.push({
        codigo:  sku,
        nombre:  String(rows[i][1] || '').trim(),
        modelos: String(rows[i][2] || '').trim(),
        pvp:     (pvRaw === '' || pvRaw == null) ? null : Number(pvRaw) || null
      });
    }
    try { cache.put(CKEY, JSON.stringify(items), 3600); } catch(e) {}
    return items;
  } catch(e) { Logger.log('_getAccesoriosCache: ' + e); return []; }
}

function buscarRepuesto(busqueda) {
  try {
    var d   = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var q   = _normText(busqueda);
    var D   = SCHEMA.DB_REPUESTOS;
    var matches = [];
    for (var i = 1; i < d.length; i++) {
      var cod    = String(d[i][D.CODIGO]          || '');
      var nom    = String(d[i][D.DESCRIPCION]     || '');
      var descEs = String(d[i][D.DESCRIPCION_ES]  || '');
      var remplz = String(d[i][D.REEMPLAZADO_POR] || '').trim();
      var nCod   = _normText(cod);
      var nNom   = _normText(nom);
      var nDescEs = _normText(descEs);
      if (nCod.indexOf(q) === -1 && nNom.indexOf(q) === -1 && nDescEs.indexOf(q) === -1) continue;
      var score;
      if (nCod === q)                                            score = 10;
      else if (nCod.indexOf(q) === 0)                           score = 6;
      else if (nNom.indexOf(q) === 0 || nDescEs.indexOf(q) === 0) score = 4;
      else                                                       score = 1;
      matches.push({ codigo: cod, nombre: nom, descripcionEs: descEs, modelos: String(d[i][D.MODELOS] || ''), reemplazadoPor: remplz, _score: score });
    }

    // También buscar en ACCESORIOS (misma spreadsheet que LISTA_PRECIOS)
    var accs = _getAccesoriosCache();
    for (var ai = 0; ai < accs.length; ai++) {
      var ac   = accs[ai];
      var nAco = _normText(ac.codigo);
      var nAno = _normText(ac.nombre);
      if (nAco.indexOf(q) === -1 && nAno.indexOf(q) === -1) continue;
      var aScore;
      if (nAco === q)              aScore = 10;
      else if (nAco.indexOf(q) === 0) aScore = 6;
      else if (nAno.indexOf(q) === 0) aScore = 4;
      else                            aScore = 1;
      matches.push({ codigo: ac.codigo, nombre: ac.nombre, descripcionEs: '', modelos: ac.modelos, reemplazadoPor: '', fuente: 'ACC', _score: aScore });
    }

    matches.sort(function(a, b) { return b._score - a._score; });
    var res = matches.slice(0, 20);
    for (var k = 0; k < res.length; k++) { delete res[k]._score; }
    return res;
  } catch(e) { return []; }
}

function registrarPedidoRepuestos(ot, lista) {
  try {
    var ref = _leerOrdenes();
    var otB = String(ot).trim().toUpperCase();
    for (var i = 1; i < ref.datos.length; i++) {
      if (!ref.datos[i][SCHEMA.OT.OT] || String(ref.datos[i][SCHEMA.OT.OT]).toUpperCase() !== otB) continue;
      var fh = Utilities.formatDate(new Date(), "GMT-3", "dd/MM/yyyy - HH:mm");

      // Nota legible en col M (TRABAJO) — historial para el técnico
      var inf     = String(ref.datos[i][SCHEMA.OT.TRABAJO] || "");
      var legible = lista.map(function(it) {
        return '[' + (it.cantidad || 1) + 'x] ' + it.sku + ' (' + (it.nombre || '') + ')';
      }).join("\n");
      ref.hoja.getRange(i + 1, SCHEMA.OT.TRABAJO + 1).setValue(inf + "\n\n PEDIDO REPUESTOS (" + fh + "):\n" + legible);

      // Merge en col Q (REPUESTOS) con formato Hub: SKU | desc | P:X E:Y
      var repActual = String(ref.datos[i][SCHEMA.OT.REPUESTOS] || "").trim();
      var mapaRep   = {};
      if (repActual && repActual !== 'Sin consumo de repuestos') {
        var lineas = repActual.split(/ ; |\r?\n/);   // tolera separador canónico ' ; ' (Hub) y '\n' legacy
        for (var li = 0; li < lineas.length; li++) {
          var partes = lineas[li].split(' | ');
          if (partes.length >= 3) {
            var k = partes[0].trim().toUpperCase();
            mapaRep[k] = { sku: partes[0].trim(), desc: partes[1].trim(), raw: partes[2].trim() };
          }
        }
      }
      for (var ni = 0; ni < lista.length; ni++) {
        var item   = lista[ni];
        var skuUp  = String(item.sku || '').trim().toUpperCase();
        var cant   = parseInt(item.cantidad) || 1;
        if (mapaRep[skuUp]) {
          var pM = mapaRep[skuUp].raw.match(/P:(\d+)/);
          var eM = mapaRep[skuUp].raw.match(/E:(\d+)/);
          mapaRep[skuUp].raw = 'P:' + ((pM ? parseInt(pM[1]) : 0) + cant) + ' E:' + (eM ? parseInt(eM[1]) : 0);
        } else {
          mapaRep[skuUp] = { sku: String(item.sku || '').trim(), desc: String(item.nombre || '').trim(), raw: 'P:' + cant + ' E:0' };
        }
      }
      var lineasFin = [];
      for (var key in mapaRep) { lineasFin.push(mapaRep[key].sku + ' | ' + mapaRep[key].desc + ' | ' + mapaRep[key].raw); }
      ref.hoja.getRange(i + 1, SCHEMA.OT.REPUESTOS + 1).setValue(lineasFin.join(' ; '));   // separador canónico del HUB (lo lee con split(' ; '))

      if (String(ref.datos[i][SCHEMA.OT.PRIORIDAD]).toUpperCase() !== "URGENTE")
        ref.hoja.getRange(i + 1, SCHEMA.OT.PRIORIDAD + 1).setValue("ALERTA REPUESTOS");
      invalidateSheetValues(SCHEMA.SHEETS.OT);
      return { success: true };
    }
    return { success: false };
  } catch(e) { Logger.log('registrarPedidoRepuestos: ' + e); return { success: false }; }
}

function generarHojaCotizacion(datosOrden, listaRepuestos) {
  try {
    if (!datosOrden || !listaRepuestos || !listaRepuestos.length)
      return { ok: false, error: 'Parametros invalidos o lista vacia' };
    var TEMPLATE_ID = '1lvowSri6sCqpnCM-pnY765Z8lMgnnbtgWQIhGio0fvY';
    var nombre = 'Pedido_Repuestos_OT_' + (datosOrden.ot || 'SIN_OT');
    var copia  = DriveApp.getFileById(TEMPLATE_ID).makeCopy(nombre);
    var sheet  = SpreadsheetApp.openById(copia.getId()).getActiveSheet();
    var filaInicio = 12;
    var skus = [], cantidades = [];
    for (var i = 0; i < listaRepuestos.length; i++) {
      skus.push([listaRepuestos[i].sku || '']);
      cantidades.push([listaRepuestos[i].cantidad || 1]);
    }
    if (skus.length) {
      sheet.getRange(filaInicio, 2, skus.length, 1).setValues(skus);
      sheet.getRange(filaInicio, 6, cantidades.length, 1).setValues(cantidades);
    }
    SpreadsheetApp.flush();
    return { ok: true, url: copia.getUrl() };
  } catch(e) {
    Logger.log('generarHojaCotizacion: ' + e);
    return { ok: false, error: e.toString() };
  }
}

function generarCarritoHTML(items) {
  var rows = '';
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var bg = i % 2 === 0 ? '#ffffff' : '#f7f9fc';
    rows +=
      "<tr style='background:" + bg + ";border-bottom:1px solid #eef2f6'>" +
        "<td style='padding:9px 12px;font-size:12px;font-family:Consolas,monospace;color:#00a3e0;font-weight:600;white-space:nowrap'>" + (it.sku || '—') + "</td>" +
        "<td style='padding:9px 12px;font-size:12px;color:#333'>" + (it.descripcion || '—') + "</td>" +
        "<td style='padding:9px 12px;font-size:13px;font-weight:700;text-align:center;color:#1a1f2e'>" + (it.cantidad || 1) + "</td>" +
      "</tr>";
  }
  return "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:1px solid #ddeef7;border-radius:8px;overflow:hidden;margin-top:4px'>" +
    "<thead><tr style='background:#e8f4fb'>" +
      "<th style='padding:9px 12px;font-size:11px;font-weight:700;color:#00a3e0;text-align:left;text-transform:uppercase;letter-spacing:.06em'>SKU</th>" +
      "<th style='padding:9px 12px;font-size:11px;font-weight:700;color:#00a3e0;text-align:left;text-transform:uppercase;letter-spacing:.06em'>Descripción</th>" +
      "<th style='padding:9px 12px;font-size:11px;font-weight:700;color:#00a3e0;text-align:center;width:80px;text-transform:uppercase;letter-spacing:.06em'>Cant.</th>" +
    "</tr></thead>" +
    "<tbody>" + rows + "</tbody>" +
  "</table>";
}

function enviarGestionRepuestos(data) {
  try {
    var cas      = String(data.cas      || '').trim();
    var modelo   = String(data.modelo   || '').trim();
    var sn       = String(data.sn       || '').trim().toUpperCase();
    var reseller = String(data.reseller || '').trim();
    var items    = data.items || [];

    if (!cas || !modelo || !sn || !items.length)
      return { ok: false, error: 'Datos incompletos.' };

    var destinatario = 'soporteagrasdji@gmail.com';
    var asunto       = 'Gestión de Garantía DJI - ' + cas + ' - ' + reseller;

    var ccEmail = '';
    try {
      var dRes = getSheetValues(SCHEMA.SHEETS.RESELLERS);
      var rLow = reseller.trim().toLowerCase();
      for (var ri = 1; ri < dRes.length; ri++) {
        if (String(dRes[ri][0] || '').trim().toLowerCase() === rLow) {
          ccEmail = String(dRes[ri][SCHEMA.RESELLERS.EMAIL] || '').trim();
          break;
        }
      }
    } catch(eCC) { Logger.log("enviarGestionRepuestos CC lookup: " + eCC); }

    var cuerpo =
      "<p style='font-size:14px;color:#444;margin:0 0 20px'>El reseller <strong>" + reseller + "</strong> solicita gestión de garantía (IW) para el siguiente caso DJI.</p>" +
      "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:20px'>" +
        _filaDetalle("N° CAS / FWR", "<strong style='font-size:13px;color:#00a3e0'>" + cas + "</strong>") +
        _filaDetalle("Modelo", modelo) +
        _filaDetalle("N° de Serie", "<span style='font-family:monospace;font-weight:600'>" + sn + "</span>") +
        _filaDetalle("Tipo de gestión", "<span style='background:rgba(26,158,74,.1);color:#1a9e4a;font-size:11px;font-weight:700;padding:2px 8px;border-radius:4px;border:1px solid rgba(26,158,74,.2)'>IW — In Warranty</span>") +
        _filaDetalle("Reseller", reseller) +
      "</div>" +
      "<div style='font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px'>Repuestos solicitados (" + items.length + " ítem" + (items.length !== 1 ? 's' : '') + ")</div>" +
      generarCarritoHTML(items);

    var html = _construirEmailHTML(
      "Gestión de Garantía DJI — " + cas,
      "Equipo DJI Aftermarket",
      cuerpo,
      "Este email fue generado automáticamente desde el Portal Resellers BIDCOMAGRO · " + reseller + "."
    );

    var opciones = { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR };
    if (ccEmail && ccEmail !== destinatario) opciones.cc = ccEmail;

    GmailApp.sendEmail(destinatario, asunto, '', opciones);

    try {
      var hojaLog = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
      if (hojaLog) hojaLog.appendRow([new Date(), cas, destinatario, 'Gestión DJI (Portal)', asunto, 'OK']);
    } catch(eLog) {
      var payload = JSON.stringify({ modulo: "enviarGestionRepuestos", hoja: "EMAIL_LOGS", error: eLog.toString() });
      Logger.log("ERROR_EMAIL_LOGS_APPEND: " + payload);
      console.log(payload);
    }

    return { ok: true };
  } catch(e) {
    Logger.log("enviarGestionRepuestos: " + e);
    return { ok: false, error: e.toString() };
  }
}

// Convierte cualquier URL de Google Drive al formato directo usable en <img>
function _driveUrlToImg(url) {
  if (!url) return '';
  // /file/d/FILE_ID/view  o  /file/d/FILE_ID?...
  var m = url.match(/\/file\/d\/([A-Za-z0-9_-]+)/);
  if (m) return 'https://lh3.googleusercontent.com/d/' + m[1];
  // ?id=FILE_ID  o  &id=FILE_ID  (links uc, open, etc.)
  m = url.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (m) return 'https://lh3.googleusercontent.com/d/' + m[1];
  return url; // URL directa (CDN, etc.) — devolver tal cual
}

// SKU (mayúsculas) → URL de foto, desde la misma hoja TODO que usa RS_getListaPrecios (misma
// spreadsheet LISTA_PRECIOS_SS_ID, mismo manejo de rich-text/chips en col H). Se separa de
// RS_getListaPrecios porque esa devuelve el catálogo entero (descripciones, precios, etc.) y
// su cache puede no entrar en el límite de 90KB de CacheService para catálogos grandes — este
// mapa es mucho más chico (solo SKU + foto) así siempre cachea. Pedido del usuario: mostrar una
// miniatura en el buscador de "Carrito de Repuestos" y en el carrito.
function _repFotoMapCatalogo() {
  var CKEY  = 'rep_foto_map_v1';
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CKEY);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  var map = {};
  try {
    var ss      = SpreadsheetApp.openById(LISTA_PRECIOS_SS_ID);
    var sheet   = ss.getSheetByName('TODO');
    var lastRow = sheet ? sheet.getLastRow() : 0;
    if (sheet && lastRow > 1) {
      var rows     = sheet.getRange(1, 1, lastRow, 8).getValues();
      var richColH = sheet.getRange(2, 8, lastRow - 1, 1).getRichTextValues();
      for (var i = 1; i < rows.length; i++) {
        var sku = String(rows[i][1] || '').trim().toUpperCase();
        if (!sku) continue;
        var fotoPlain = String(rows[i][7] || '').trim();
        var fotoUrl   = (fotoPlain.indexOf('http') === 0) ? fotoPlain : '';
        if (!fotoUrl) {
          try {
            var richCell = richColH[i - 1] && richColH[i - 1][0];
            if (richCell) {
              fotoUrl = richCell.getLinkUrl() || '';
              if (!fotoUrl) {
                var runs = richCell.getRuns();
                for (var r = 0; r < runs.length; r++) {
                  var ru = runs[r].getLinkUrl();
                  if (ru) { fotoUrl = ru; break; }
                }
              }
            }
          } catch(eFoto) {}
        }
        if (fotoUrl) map[sku] = _driveUrlToImg(fotoUrl);
      }
    }
  } catch(e) { Logger.log('_repFotoMapCatalogo: ' + e); }
  try {
    var payload = JSON.stringify(map);
    if (payload.length < 90000) cache.put(CKEY, payload, 3600);
  } catch(eCa) {}
  return map;
}

function RS_getListaPrecios() {
  Logger.log('RS_getListaPrecios: inicio. LISTA_PRECIOS_SS_ID=' + LISTA_PRECIOS_SS_ID);
  var CKEY = 'lista_precios_v2';
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CKEY);
  if (cached) {
    Logger.log('RS_getListaPrecios: devolviendo cache');
    try { return JSON.parse(cached); } catch(e) {
      Logger.log('RS_getListaPrecios: cache parse error: ' + e);
    }
  }
  try {
    Logger.log('RS_getListaPrecios: abriendo spreadsheet...');
    var ss     = SpreadsheetApp.openById(LISTA_PRECIOS_SS_ID);
    Logger.log('RS_getListaPrecios: spreadsheet abierto. Buscando hoja TODO...');
    var sheet  = ss.getSheetByName('TODO');
    if (!sheet) {
      var nombres = ss.getSheets().map(function(s) { return s.getName(); }).join(', ');
      Logger.log('RS_getListaPrecios: hoja TODO no encontrada. Disponibles: ' + nombres);
      return { ok: false, msg: 'Hoja TODO no encontrada. Disponibles: ' + nombres };
    }
    Logger.log('RS_getListaPrecios: hoja encontrada, leyendo filas...');
    var lastRow  = sheet.getLastRow();
    var lastCol  = sheet.getLastColumn();
    var rows     = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    // getRichTextValues solo en col H (índice 8) para no leer todo el sheet
    var richColH = lastRow > 1
      ? sheet.getRange(2, 8, lastRow - 1, 1).getRichTextValues()
      : [];
    Logger.log('RS_getListaPrecios: ' + rows.length + ' filas');
    var items = [];
    var fotoLog = []; // para diagnóstico
    for (var i = 1; i < rows.length; i++) {
      var sku = String(rows[i][1] || '').trim();
      if (!sku) continue;
      var pvpRaw  = rows[i][5];
      var cantRaw = rows[i][4];

      // Col H: valor plano primero (cubre URLs de texto), luego rich text para chips
      var fotoPlain = String(rows[i][7] || '').trim();
      var fotoUrl   = (fotoPlain.indexOf('http') === 0) ? fotoPlain : '';
      if (!fotoUrl) {
        try {
          var richCell = richColH[i - 1] && richColH[i - 1][0];
          if (richCell) {
            fotoUrl = richCell.getLinkUrl() || '';
            if (!fotoUrl) {
              var runs = richCell.getRuns();
              for (var r = 0; r < runs.length; r++) {
                var ru = runs[r].getLinkUrl();
                if (ru) { fotoUrl = ru; break; }
              }
            }
          }
        } catch(eFoto) {}
      }

      // Convertir URLs de Google Drive al formato de thumbnail directo
      if (fotoUrl) fotoUrl = _driveUrlToImg(fotoUrl);

      if (fotoLog.length < 3) fotoLog.push(sku + ' → ' + (fotoUrl || '(sin foto)'));

      items.push({
        sku:       sku,
        desc:      String(rows[i][2] || '').trim(),
        modelos:   String(rows[i][3] || '').trim(),
        cantUsada: (cantRaw === '' || cantRaw === null || cantRaw === undefined) ? '' : String(cantRaw),
        pvp:       (pvpRaw === '' || pvpRaw === null || pvpRaw === undefined) ? null : Number(pvpRaw) || null,
        foto:      fotoUrl
      });
    }
    // Agregar ítems de hoja ACCESORIOS (misma spreadsheet)
    try {
      var accSheet = ss.getSheetByName('ACCESORIOS');
      if (accSheet) {
        var accRows = accSheet.getDataRange().getValues();
        for (var ai = 1; ai < accRows.length; ai++) {
          var aSku = String(accRows[ai][0] || '').trim();
          if (!aSku) continue;
          var aPvRaw = accRows[ai][3];
          items.push({
            sku:       aSku,
            desc:      String(accRows[ai][1] || '').trim(),
            modelos:   String(accRows[ai][2] || '').trim(),
            cantUsada: '',
            pvp:       (aPvRaw === '' || aPvRaw == null) ? null : Number(aPvRaw) || null,
            foto:      '',
            fuente:    'ACCESORIOS'
          });
        }
        Logger.log('RS_getListaPrecios: ACCESORIOS agregados: ' + (accRows.length - 1));
      }
    } catch(eAcc) { Logger.log('RS_getListaPrecios ACCESORIOS: ' + eAcc); }

    Logger.log('RS_getListaPrecios: muestra fotos: ' + fotoLog.join(' | '));
    Logger.log('RS_getListaPrecios: ' + items.length + ' items procesados. Retornando.');
    var resultado = { ok: true, items: items };
    try {
      var payload = JSON.stringify(resultado);
      if (payload.length < 90000) cache.put(CKEY, payload, 3600);
    } catch(eCa) { Logger.log('RS_getListaPrecios: cache write error: ' + eCa); }
    return resultado;
  } catch(e) {
    Logger.log('RS_getListaPrecios ERROR: ' + e + ' | stack: ' + e.stack);
    return { ok: false, msg: String(e) };
  }
}
