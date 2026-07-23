// @version 1.0
// ── STOCK MANAGER — Compras internas (herramientas e insumos) ──────────────
// ============================================================
//  Lista viva de "qué hay que comprar" que cargan los operarios.
//  Hoja "Compra de herramientas e insumos" (MASTER). Cols A–I las definió el
//  usuario; J=ID lo gestiona la app para ubicar filas de forma robusta.
//    A Fecha de pedido | B Descripcion | C Cantidad | D Tipo |
//    E Precio unitario | F Total | G ¿Pedido? | H ¿Llego? | I Link | J ID
// ============================================================

// ── HELPERS ────────────────────────────────────────────────
function _ciNuevoId() {
  return 'HRR-' + Utilities.getUuid();
}

function _ciNum(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  s = s.replace(/[^\d.,\-]/g, '');
  if (s.indexOf(',') > -1 && s.indexOf('.') === -1) s = s.replace(',', '.'); // coma decimal
  else s = s.replace(/,/g, '');                                             // coma = miles
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function _ciBool(v) {
  if (v === true) return true;
  var s = String(v || '').trim().toLowerCase();
  return s === 'sí' || s === 'si' || s === 'true' || s === 'verdadero' ||
         s === 'x'  || s === '1'  || s === '✓'    || s === 'ok';
}

// Normaliza el Tipo a las dos opciones fijas (o vacío)
function _ciTipo(v) {
  var s = String(v || '').trim().toLowerCase();
  if (s.indexOf('herram') === 0) return 'Herramienta';
  if (s.indexOf('insumo') === 0) return 'Insumo';
  return '';
}

// Anti-inyección de fórmulas: antepone ' a texto libre que empieza con = + - @
function _ciSafe(v) {
  var s = String(v == null ? '' : v);
  return /^[=+\-@]/.test(s) ? ("'" + s) : s;
}

// Ubica la fila (1-based) por ID en col J. -1 si no existe.
function _ciFindRowById(id, hoja) {
  var C = SCHEMA.COMPRAS_INTERNAS;
  hoja = hoja || getSheet(SCHEMA.SHEETS.COMPRAS_INTERNAS);
  if (!hoja) return -1;
  var target = String(id || '').trim();
  if (!target) return -1;
  var lastRow = hoja.getLastRow();
  if (lastRow < 2) return -1;
  var col = hoja.getRange(2, C.ID + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    if (String(col[i][0] || '').trim() === target) return i + 2;
  }
  return -1;
}

// Idempotente: asegura la hoja + encabezado ID en col J + rellena IDs faltantes
// (para las filas que el usuario haya cargado a mano). Devuelve la hoja.
function _ciSetup() {
  var ss   = getSS();
  var name = SCHEMA.SHEETS.COMPRAS_INTERNAS;
  var C    = SCHEMA.COMPRAS_INTERNAS;
  var headers = ['Fecha de pedido', 'Descripcion', 'Cantidad', 'Tipo',
                 'Precio unitario', 'Total', '¿Pedido?', '¿Llego?', 'Link', 'ID'];
  var hoja = ss.getSheetByName(name);
  if (!hoja) {
    hoja = ss.insertSheet(name);
    hoja.appendRow(headers);
    hoja.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    invalidateSheetValues(name);
    return hoja;
  }
  var lastRow = hoja.getLastRow();
  if (lastRow < 1) {
    hoja.appendRow(headers);
    hoja.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    invalidateSheetValues(name);
    return hoja;
  }
  // Asegurar encabezado "ID" en col J
  var hdrId = hoja.getRange(1, C.ID + 1).getValue();
  if (String(hdrId || '').trim().toUpperCase() !== 'ID') {
    hoja.getRange(1, C.ID + 1).setValue('ID').setFontWeight('bold');
  }
  // Rellenar IDs faltantes en filas con contenido
  if (lastRow >= 2) {
    var rng  = hoja.getRange(2, 1, lastRow - 1, C.ID + 1);
    var vals = rng.getValues();
    var idCol = [], changed = false;
    for (var i = 0; i < vals.length; i++) {
      var row = vals[i];
      var tieneContenido = String(row[C.DESCRIPCION] || '').trim() !== '' ||
                           String(row[C.LINK] || '').trim() !== '' ||
                           String(row[C.CANTIDAD] || '').trim() !== '';
      var idActual = String(row[C.ID] || '').trim();
      if (!idActual && tieneContenido) { idActual = _ciNuevoId(); changed = true; }
      idCol.push([idActual]);
    }
    if (changed) {
      hoja.getRange(2, C.ID + 1, idCol.length, 1).setValues(idCol);
      invalidateSheetValues(name);
    }
  }
  return hoja;
}

// ── LECTURA ────────────────────────────────────────────────
function getComprasInternas() {
  try {
    _ciSetup();
    var C = SCHEMA.COMPRAS_INTERNAS;
    var d = getSheetValues(SCHEMA.SHEETS.COMPRAS_INTERNAS, true); // fresco tras setup
    var items = [];
    for (var i = 1; i < d.length; i++) {
      var f = d[i];
      var id   = String(f[C.ID] || '').trim();
      var desc = String(f[C.DESCRIPCION] || '').trim();
      if (!id && !desc) continue; // fila vacía
      var precio = _ciNum(f[C.PRECIO_UNITARIO]);
      var cant   = _ciNum(f[C.CANTIDAD]);
      var totCel = _ciNum(f[C.TOTAL]);
      items.push({
        id: id,
        fecha:   _fmtFecha(f[C.FECHA_PEDIDO]),
        fechaMs: (f[C.FECHA_PEDIDO] instanceof Date) ? f[C.FECHA_PEDIDO].getTime() : 0,
        descripcion: desc,
        cantidad: cant,
        tipo: String(f[C.TIPO] || '').trim(),
        precioUnitario: precio,
        total: totCel || Math.round(precio * cant * 100) / 100,
        pedido: _ciBool(f[C.PEDIDO]),
        llego:  _ciBool(f[C.LLEGO]),
        link: String(f[C.LINK] || '').trim()
      });
    }
    items.sort(function(a, b) { return b.fechaMs - a.fechaMs; }); // más nuevos primero
    return { ok: true, items: items };
  } catch(e) {
    Logger.log('getComprasInternas: ' + e);
    return { ok: false, msg: e.toString(), items: [] };
  }
}

// ── ALTA ───────────────────────────────────────────────────
function agregarCompraInterna(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    data = data || {};
    var desc = _ciSafe(String(data.descripcion || '').trim());
    if (!desc) return { ok: false, msg: 'Falta la descripción.' };
    var cant = _ciNum(data.cantidad); if (!cant || cant <= 0) cant = 1;
    var tipo   = _ciTipo(data.tipo);
    var precio = _ciNum(data.precioUnitario);
    var total  = precio > 0 ? Math.round(precio * cant * 100) / 100 : '';
    var link   = _ciSafe(String(data.link || '').trim());
    var id     = _ciNuevoId();
    var hoja   = _ciSetup();
    hoja.appendRow([ new Date(), desc, cant, tipo, (precio > 0 ? precio : ''), total, '', '', link, id ]);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    return { ok: true, id: id };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── MARCAR ¿Pedido? / ¿Llegó? ──────────────────────────────
function marcarCompraInterna(id, campo, valor) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var C   = SCHEMA.COMPRAS_INTERNAS;
    var col = (campo === 'pedido') ? C.PEDIDO : (campo === 'llego' ? C.LLEGO : -1);
    if (col === -1) return { ok: false, msg: 'Campo inválido.' };
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    var rowNum = _ciFindRowById(id, hoja);
    if (rowNum === -1) return { ok: false, msg: 'No se encontró el ítem.' };
    hoja.getRange(rowNum, col + 1).setValue(valor ? 'SÍ' : '');
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── EDICIÓN ────────────────────────────────────────────────
function editarCompraInterna(id, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    data = data || {};
    var C = SCHEMA.COMPRAS_INTERNAS;
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    var rowNum = _ciFindRowById(id, hoja);
    if (rowNum === -1) return { ok: false, msg: 'No se encontró el ítem.' };
    var desc = _ciSafe(String(data.descripcion || '').trim());
    if (!desc) return { ok: false, msg: 'Falta la descripción.' };
    var cant = _ciNum(data.cantidad); if (!cant || cant <= 0) cant = 1;
    var tipo   = _ciTipo(data.tipo);
    var precio = _ciNum(data.precioUnitario);
    var total  = precio > 0 ? Math.round(precio * cant * 100) / 100 : '';
    var link   = _ciSafe(String(data.link || '').trim());
    // B..F contiguo (Descripcion, Cantidad, Tipo, Precio, Total)
    hoja.getRange(rowNum, C.DESCRIPCION + 1, 1, 5)
        .setValues([[ desc, cant, tipo, (precio > 0 ? precio : ''), total ]]);
    hoja.getRange(rowNum, C.LINK + 1).setValue(link);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ── BORRADO ────────────────────────────────────────────────
function eliminarCompraInterna(id) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    var rowNum = _ciFindRowById(id, hoja);
    if (rowNum === -1) return { ok: false, msg: 'No se encontró el ítem.' };
    hoja.deleteRow(rowNum);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_INTERNAS);
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
