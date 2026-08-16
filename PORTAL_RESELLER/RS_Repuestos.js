// ============================================================
// @version 1.10
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

function registrarPedidoRepuestos(token, ot, lista) {
  try {
    // Sin esto, cualquiera podía agregar repuestos a la OT de OTRO reseller desde la consola
    // del navegador (era la única mutación de OT en el proyecto sin este chequeo — el resto
    // ya lo tiene desde la auditoría de seguridad de agosto 2026, ver agregarComentario).
    var _s = _sesionResolver(token);
    if (!_s) return { success: false, error: 'Sesión inválida o expirada. Volvé a ingresar.' };

    var ref = _leerOrdenes();
    var otB = String(ot).trim().toUpperCase();
    for (var i = 1; i < ref.datos.length; i++) {
      if (!ref.datos[i][SCHEMA.OT.OT] || String(ref.datos[i][SCHEMA.OT.OT]).toUpperCase() !== otB) continue;
      if (!_sesionPoseeReseller(_s, ref.datos[i][SCHEMA.OT.RESELLER])) return { success: false, error: 'No autorizado.' };
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
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
