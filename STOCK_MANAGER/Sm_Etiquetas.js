// @version 1.5
// ══════════════════════════════════════════════════════════════
//  ETIQUETAS SO — códigos internos para productos SIN N° de serie
//  El SO funciona como un SN (uno único por unidad). Formato POR SKU:
//    SO-<SKU>-NNN   (contador incremental por SKU)
//  Se registran en la hoja SO_ETIQUETAS para trazabilidad y para que
//  el próximo NNN nunca se repita. El frontend imprime el código de
//  barras Code128 en la Xprinter (etiquetas 50×25mm, 2 por fila).
//  BOLSAS: un mismo SO puede representar N unidades (consumibles: tornillos,
//  gaskets). Se guarda la cantidad en col CANTIDAD y se marca el SKU como
//  "por bolsa" en el maestro compartido con WOS (MAESTRO_ARTICULOS).
// ══════════════════════════════════════════════════════════════

// Columnas de la hoja SO_ETIQUETAS
var SO_COL = { SO: 0, SKU: 1, DESC: 2, FECHA: 3, OPERADOR: 4, CANTIDAD: 5 };

// Layout del maestro compartido con WOS (hoja MAESTRO_ARTICULOS en WOS_NOTAS_SS_ID).
var _SM_MAESTRO_HOJA = 'MAESTRO_ARTICULOS';
var _SM_MAESTRO_COL  = { SKU: 0, DESC: 1, POR_BOLSA: 2, BULTO: 3, FECHA: 4, OPERADOR: 5 };

// Crea (si falta) la hoja SO_ETIQUETAS y la devuelve. Asegura la col CANTIDAD (F).
function _getHojaSO() {
  var db = getDb();
  var hoja = db.getSheetByName(SCHEMA.SHEETS.SO_ETIQUETAS);
  if (!hoja) {
    hoja = db.insertSheet(SCHEMA.SHEETS.SO_ETIQUETAS);
    hoja.appendRow(['SO', 'SKU', 'DESCRIPCION', 'FECHA', 'OPERADOR', 'CANTIDAD']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#fff');
    hoja.setColumnWidth(1, 200); hoja.setColumnWidth(3, 260);
  } else if (hoja.getMaxColumns() < 6 || !String(hoja.getRange(1, 6).getValue() || '').trim()) {
    // Upgrade de hojas viejas (5 columnas): agregar el encabezado CANTIDAD (F).
    if (hoja.getMaxColumns() < 6) hoja.insertColumnsAfter(hoja.getMaxColumns(), 6 - hoja.getMaxColumns());
    hoja.getRange(1, 6).setValue('CANTIDAD').setFontWeight('bold').setBackground('#00a3e0').setFontColor('#fff');
  }
  return hoja;
}

// Marca un SKU como "por bolsa" (+ bulto) en el maestro que también lee WOS
// (hoja MAESTRO_ARTICULOS en WOS_NOTAS_SS_ID). Nunca rompe la generación de etiquetas.
function _smUpsertMaestro(sku, desc, bulto, operador) {
  try {
    sku = String(sku || '').trim();
    if (!sku) return;
    var ss = SpreadsheetApp.openById(WOS_NOTAS_SS_ID);
    var hoja = ss.getSheetByName(_SM_MAESTRO_HOJA);
    if (!hoja) {
      hoja = ss.insertSheet(_SM_MAESTRO_HOJA);
      hoja.appendRow(['SKU', 'Descripción', 'Por bolsa', 'Bulto x defecto', 'Última actualización', 'Operador']);
      hoja.getRange(1, 1, 1, 6).setFontWeight('bold');
      hoja.setFrozenRows(1);
    }
    if (hoja.getMaxColumns() < 6) hoja.insertColumnsAfter(hoja.getMaxColumns(), 6 - hoja.getMaxColumns());
    var data = hoja.getDataRange().getValues();
    var su = sku.toUpperCase();
    var fila = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][_SM_MAESTRO_COL.SKU] || '').trim().toUpperCase() === su) { fila = i + 1; break; }
    }
    if (!fila) { hoja.appendRow([sku, desc || '', '', '', '', '']); fila = hoja.getLastRow(); }
    hoja.getRange(fila, _SM_MAESTRO_COL.SKU + 1).setValue(sku);
    if (desc) hoja.getRange(fila, _SM_MAESTRO_COL.DESC + 1).setValue(desc);
    hoja.getRange(fila, _SM_MAESTRO_COL.POR_BOLSA + 1).setValue(true);
    var b = parseInt(bulto, 10) || 0;
    if (b > 0) hoja.getRange(fila, _SM_MAESTRO_COL.BULTO + 1).setValue(b);
    hoja.getRange(fila, _SM_MAESTRO_COL.FECHA + 1).setValue(new Date());
    if (operador) hoja.getRange(fila, _SM_MAESTRO_COL.OPERADOR + 1).setValue(operador);
    SpreadsheetApp.flush();
  } catch(e) {
    Logger.log('_smUpsertMaestro: ' + e);
  }
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

// Genera `cantBolsas` códigos SO de BOLSA para un SKU: cada código representa
// `unidadesPorBolsa` unidades (1 etiqueta = 1 bolsa). Los registra con la cantidad
// (col CANTIDAD) y marca el SKU como "por bolsa" + bulto en el maestro compartido con
// WOS, para que la preparación lo reconozca solo y autocomplete la cantidad al escanear.
// Devuelve { ok, labels:[{ so, sku, descripcion, cantidad }], desde, hasta, unidadesPorBolsa }
function generarBolsasSO(sku, unidadesPorBolsa, cantBolsas, operador) {
  sku = String(sku || '').trim().toUpperCase();
  unidadesPorBolsa = parseInt(unidadesPorBolsa, 10) || 0;
  cantBolsas       = parseInt(cantBolsas, 10) || 0;
  if (!sku) return { ok: false, error: 'Falta el SKU.' };
  if (unidadesPorBolsa <= 0) return { ok: false, error: 'Las unidades por bolsa deben ser mayor a 0.' };
  if (cantBolsas <= 0) return { ok: false, error: 'La cantidad de bolsas debe ser mayor a 0.' };
  if (cantBolsas > 200) return { ok: false, error: 'Máximo 200 bolsas por vez.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var hoja = _getHojaSO();
    var datos = hoja.getDataRange().getValues();

    var prefijo = 'SO-' + sku + '-';
    var maxN = 0;
    for (var i = 1; i < datos.length; i++) {
      var val = String(datos[i][SO_COL.SO] || '').trim().toUpperCase();
      if (val.indexOf(prefijo) === 0) {
        var n = parseInt(val.substring(prefijo.length), 10);
        if (!isNaN(n) && n > maxN) maxN = n;
      }
    }

    var desc  = _descDeSku(sku);
    var op    = String(operador || '');
    var ahora = new Date();
    var labels = [], filas = [];
    for (var k = 1; k <= cantBolsas; k++) {
      var num  = maxN + k;
      var sNum = String(num);
      var pad  = (sNum.length >= 6) ? sNum : ('000000' + sNum).slice(-6);
      var so   = prefijo + pad;
      labels.push({ so: so, sku: sku, descripcion: desc, cantidad: unidadesPorBolsa });
      filas.push([so, sku, desc, ahora, op, unidadesPorBolsa]);
    }
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
    SpreadsheetApp.flush();

    // Aprender en la MISMA hoja que WOS: este SKU va por bolsa (bulto = unidadesPorBolsa).
    _smUpsertMaestro(sku, desc, unidadesPorBolsa, op);

    return { ok: true, labels: labels, desde: maxN + 1, hasta: maxN + cantBolsas, unidadesPorBolsa: unidadesPorBolsa };
  } catch(e) {
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}

// Mapa de SKUs marcados "por bolsa" en el maestro compartido con WOS (MAESTRO_ARTICULOS):
// { SKU: unidadesPorBolsa }. Lo usa la recepción para saber qué códigos etiquetar como bulto.
function getMaestroBolsas() {
  var out = {};
  try {
    var ss   = SpreadsheetApp.openById(WOS_NOTAS_SS_ID);
    var hoja = ss.getSheetByName(_SM_MAESTRO_HOJA);
    if (!hoja) return out;
    var d = hoja.getDataRange().getValues();
    var C = _SM_MAESTRO_COL;
    for (var i = 1; i < d.length; i++) {
      var sku = String(d[i][C.SKU] || '').trim().toUpperCase();
      if (!sku) continue;
      var pb  = d[i][C.POR_BOLSA];
      var esBolsa = (pb === true) ||
        String(pb).trim().toLowerCase() === 'true' ||
        String(pb).trim().toUpperCase() === 'SI' ||
        String(pb).trim().toUpperCase() === 'S\xcd';
      if (esBolsa) out[sku] = parseInt(d[i][C.BULTO], 10) || 0;
    }
  } catch(e) { Logger.log('getMaestroBolsas: ' + e); }
  return out;
}

// SKUs marcados "requiere N° de serie" en STOCK_REPUESTOS → { SKU: true }. La recepción los deja
// DESTILDADOS por defecto (llevan SN real, no SO), aunque el operador puede tildarlos igual.
function getReqSNMap() {
  var out = {};
  try {
    var d = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S = SCHEMA.STOCK_REPUESTOS;
    for (var i = 1; i < d.length; i++) {
      var sku = String(d[i][S.CODIGO] || '').trim().toUpperCase();
      if (sku && d[i][S.REQUIERE_SN] === true) out[sku] = true;
    }
  } catch(e) { Logger.log('getReqSNMap: ' + e); }
  return out;
}

// Metadata que la recepción necesita para etiquetar bien, en una sola llamada:
// { bolsas: {SKU:unidadesPorBolsa}, sn: {SKU:true} }.
function getEtiquetaMetaRecepcion() {
  return { bolsas: getMaestroBolsas(), sn: getReqSNMap() };
}

// Genera etiquetas SO para varios SKUs en una sola pasada (una sola escritura, un solo lock).
// Lo usa la recepción de compras para imprimir en automático las etiquetas de los códigos que
// el operador tildó, respetando bulto vs unidad:
//   items = [{ sku, cantidad, modo, unidadesPorBolsa }]
//     · modo 'unidad' (default): `cantidad` = etiquetas, 1 por unidad.
//     · modo 'bolsa': `cantidad` = cantidad de bolsas, cada etiqueta = `unidadesPorBolsa` unidades
//       (badge "BOLSA xN u." al imprimir) y el SKU queda marcado por bolsa en el maestro.
// Todas las filas se escriben con 6 columnas (F=CANTIDAD; vacía en modo unidad). Cachea en 200.
// Devuelve { ok, labels:[{so,sku,descripcion,cantidad?}] }.
function generarSOLote(items, operador) {
  items = items || [];
  var limpio = [];
  for (var a = 0; a < items.length; a++) {
    var sk = String(items[a].sku || '').trim().toUpperCase();
    var cn = parseInt(items[a].cantidad, 10) || 0;
    if (!sk || cn <= 0) continue;
    var esBolsa = String(items[a].modo || '') === 'bolsa';
    var uxb = esBolsa ? (parseInt(items[a].unidadesPorBolsa, 10) || 0) : 0;
    if (esBolsa && uxb <= 0) { esBolsa = false; }  // sin tamaño de bolsa válido → tratar por unidad
    limpio.push({ sku: sk, cantidad: Math.min(cn, 200), esBolsa: esBolsa, uxb: uxb });
  }
  if (!limpio.length) return { ok: false, error: 'No hay códigos tildados para etiquetar.' };

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var hoja  = _getHojaSO();
    var datos = hoja.getDataRange().getValues();
    var op    = String(operador || '');
    var ahora = new Date();

    // maxN por SKU: arranca de lo ya registrado y se incrementa dentro del lote (soporta el mismo
    // SKU más de una vez sin repetir NNN).
    var maxPorSku = {};
    function _maxNPara(prefijo) {
      var mx = 0;
      for (var i = 1; i < datos.length; i++) {
        var val = String(datos[i][SO_COL.SO] || '').trim().toUpperCase();
        if (val.indexOf(prefijo) === 0) {
          var n = parseInt(val.substring(prefijo.length), 10);
          if (!isNaN(n) && n > mx) mx = n;
        }
      }
      return mx;
    }

    var labels = [], filas = [], bolsasMarcadas = {};
    for (var t = 0; t < limpio.length; t++) {
      var it      = limpio[t];
      var prefijo = 'SO-' + it.sku + '-';
      if (maxPorSku[it.sku] === undefined) maxPorSku[it.sku] = _maxNPara(prefijo);
      var desc = _descDeSku(it.sku);
      for (var k = 1; k <= it.cantidad; k++) {
        var num  = maxPorSku[it.sku] + 1; maxPorSku[it.sku] = num;
        var sNum = String(num);
        var pad  = (sNum.length >= 6) ? sNum : ('000000' + sNum).slice(-6);
        var so   = prefijo + pad;
        if (it.esBolsa) {
          labels.push({ so: so, sku: it.sku, descripcion: desc, cantidad: it.uxb });
          filas.push([so, it.sku, desc, ahora, op, it.uxb]);
        } else {
          labels.push({ so: so, sku: it.sku, descripcion: desc });
          filas.push([so, it.sku, desc, ahora, op, '']);
        }
      }
      if (it.esBolsa && bolsasMarcadas[it.sku] === undefined) bolsasMarcadas[it.sku] = it.uxb;
    }

    if (filas.length) {
      hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
      SpreadsheetApp.flush();
    }

    // Aprender en el maestro (por bolsa + bulto) los SKU bagueados en este lote, igual que generarBolsasSO.
    var bk = Object.keys(bolsasMarcadas);
    for (var b = 0; b < bk.length; b++) {
      _smUpsertMaestro(bk[b], _descDeSku(bk[b]), bolsasMarcadas[bk[b]], op);
    }

    return { ok: true, labels: labels };
  } catch(e) {
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}
