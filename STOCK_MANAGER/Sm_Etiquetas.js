// @version 1.1
// ══════════════════════════════════════════════════════════════
//  ETIQUETAS SO — códigos internos para productos SIN N° de serie
//  El SO funciona como un SN (uno único por unidad). Formato POR SKU:
//    SO-<SKU>-NNN   (contador incremental por SKU)
//  Se registran en la hoja SO_ETIQUETAS para trazabilidad y para que
//  el próximo NNN nunca se repita. El frontend imprime el código de
//  barras Code128 en la Xprinter (etiquetas 50×25mm, 2 por fila).
// ══════════════════════════════════════════════════════════════

// Columnas de la hoja SO_ETIQUETAS
var SO_COL = { SO: 0, SKU: 1, DESC: 2, FECHA: 3, OPERADOR: 4 };

// Crea (si falta) la hoja SO_ETIQUETAS y la devuelve.
function _getHojaSO() {
  var db = getDb();
  var hoja = db.getSheetByName(SCHEMA.SHEETS.SO_ETIQUETAS);
  if (!hoja) {
    hoja = db.insertSheet(SCHEMA.SHEETS.SO_ETIQUETAS);
    hoja.appendRow(['SO', 'SKU', 'DESCRIPCION', 'FECHA', 'OPERADOR']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#fff');
    hoja.setColumnWidth(1, 200); hoja.setColumnWidth(3, 260);
  }
  return hoja;
}

// Descripción del SKU (STOCK primero, luego DB_REPUESTOS).
function _descDeSku(skuUp) {
  try {
    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    for (var i = 1; i < dStr.length; i++) {
      if (String(dStr[i][0] || '').trim().toUpperCase() === skuUp) return String(dStr[i][1] || '');
    }
  } catch(e) {}
  try {
    var dRep = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    for (var j = 1; j < dRep.length; j++) {
      if (String(dRep[j][1] || '').trim().toUpperCase() === skuUp) return String(dRep[j][2] || '');
    }
  } catch(e) {}
  return '';
}

// Genera `cantidad` códigos SO nuevos para un SKU, los registra y los devuelve.
// Devuelve { ok, labels:[{ so, sku, descripcion }], desde, hasta }
function generarSO(sku, cantidad, operador) {
  sku = String(sku || '').trim().toUpperCase();
  cantidad = parseInt(cantidad, 10) || 0;
  if (!sku) return { ok: false, error: 'Falta el SKU.' };
  if (cantidad <= 0) return { ok: false, error: 'La cantidad debe ser mayor a 0.' };
  if (cantidad > 200) return { ok: false, error: 'Máximo 200 etiquetas por vez.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var hoja = _getHojaSO();
    var datos = hoja.getDataRange().getValues();

    // Último NNN usado para este SKU
    var prefijo = 'SO-' + sku + '-';
    var maxN = 0;
    for (var i = 1; i < datos.length; i++) {
      var val = String(datos[i][SO_COL.SO] || '').trim().toUpperCase();
      if (val.indexOf(prefijo) === 0) {
        var n = parseInt(val.substring(prefijo.length), 10);
        if (!isNaN(n) && n > maxN) maxN = n;
      }
    }

    var desc = _descDeSku(sku);
    var op   = String(operador || '');
    var ahora = new Date();
    var labels = [], filas = [];
    for (var k = 1; k <= cantidad; k++) {
      var num = maxN + k;
      var sNum = String(num);
      var pad = (sNum.length >= 6) ? sNum : ('000000' + sNum).slice(-6);  // 6 dígitos (000041)
      var so  = prefijo + pad;
      labels.push({ so: so, sku: sku, descripcion: desc });
      filas.push([so, sku, desc, ahora, op]);
    }
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 5).setValues(filas);
    SpreadsheetApp.flush();

    return { ok: true, labels: labels, desde: maxN + 1, hasta: maxN + cantidad };
  } catch(e) {
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}
