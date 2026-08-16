// ============================================================
// @version 1.2
//  WOS — Snapshot de stock read-only (WMS integration)
//
//  Lee Carmen (STOCK + UBICACIONES) y MASTER (TABLA_POSICIONES)
//  en un único barrido por spreadsheet y construye mapas en
//  memoria indexados por SKU.  CERO escrituras a sheets de SM.
//
//  Depende de: Despacho_Env.js (CARMEN_SS_ID, MASTER_SS_ID)
// ============================================================

// ── Índices de columna — solo lectura ───────────────────────
var COL_CSTK = { SKU: 0, STOCK_ACTUAL: 2 };          // Carmen 'STOCK'
var COL_CUBIC = { SKU: 0, BIN_ID: 1, CANTIDAD: 2 };  // Carmen 'UBICACIONES'
var COL_CPOS  = { SKU: 0, BIN_ID: 1, CANTIDAD: 2, TIPO_ALMACEN: 3 }; // MASTER 'TABLA_POSICIONES'

// ── Helpers ─────────────────────────────────────────────────

function _wosParseBin(binId) {
  var p = String(binId || '').trim().split('-');
  return { pasillo: p[0] || '', estante: p[1] || '', nivel: p[2] || '' };
}

// Clave de ordenamiento alfanumérico natural para un BIN_ID "P-EE-NN"
function _wosBinSortKey(binId) {
  var parts = String(binId || '').trim().split('-');
  var pasillo = (parts[0] || '').toUpperCase();
  var estante = parts[1] ? (isNaN(parts[1]) ? parts[1] : ('00' + parts[1]).slice(-3)) : '';
  var nivel   = parts[2] ? (isNaN(parts[2]) ? parts[2] : ('00' + parts[2]).slice(-3)) : '';
  return pasillo + '|' + estante + '|' + nivel;
}

// ── Core: snapshot triple en memoria ────────────────────────

/**
 * Construye un snapshot read-only del inventario:
 *   stockMap[SKU] = {
 *     stock_neto_total : Number,
 *     cajas_disponibles: [
 *       { bin_id, pasillo, estante, nivel, cantidad_en_caja, tipo_almacen }
 *     ]   ← ordenadas PICKING primero, luego pasillo/estante/nivel
 *   }
 *
 * Fuentes (solo .getValues(), sin escrituras):
 *   · Carmen 'STOCK'           → stock_neto_total
 *   · Carmen 'UBICACIONES'     → cajas_disponibles (qty por bin)
 *   · MASTER 'TABLA_POSICIONES'→ tipo_almacen inyectado en cada caja
 *
 * @return {{ stockMap: Object, ts: Date }}
 */
function _wosCargarStockSnapshot() {
  var stockMap = {};

  // ── Paso 1 + 2: Carmen (una sola apertura) ───────────────
  try {
    var carmenSS = SpreadsheetApp.openById(CARMEN_SS_ID);

    // Paso 1 — stock neto total desde 'STOCK'
    try {
      var stkRaw = carmenSS.getSheetByName('STOCK').getDataRange().getValues();
      for (var i = 1; i < stkRaw.length; i++) {
        var sku = String(stkRaw[i][COL_CSTK.SKU] || '').trim().toUpperCase();
        if (!sku) continue;
        stockMap[sku] = {
          stock_neto_total:  parseInt(stkRaw[i][COL_CSTK.STOCK_ACTUAL]) || 0,
          cajas_disponibles: []
        };
      }
    } catch(eStk) { Logger.log('_wosCargarStockSnapshot STOCK: ' + eStk); }

    // Paso 2 — desglose por bin desde 'UBICACIONES'
    try {
      var ubicRaw = carmenSS.getSheetByName(CARMEN_UBICACIONES_TAB).getDataRange().getValues();
      for (var j = 1; j < ubicRaw.length; j++) {
        var uSku = String(ubicRaw[j][COL_CUBIC.SKU]    || '').trim().toUpperCase();
        var bin  = String(ubicRaw[j][COL_CUBIC.BIN_ID] || '').trim();
        var cant = parseInt(ubicRaw[j][COL_CUBIC.CANTIDAD]) || 0;
        if (!uSku || !bin) continue;
        if (!stockMap[uSku]) stockMap[uSku] = { stock_neto_total: 0, cajas_disponibles: [] };
        var coords = _wosParseBin(bin);
        stockMap[uSku].cajas_disponibles.push({
          bin_id:           bin,
          pasillo:          coords.pasillo,
          estante:          coords.estante,
          nivel:            coords.nivel,
          cantidad_en_caja: cant,
          tipo_almacen:     'PICKING'   // default; sobreescrito en Paso 4
        });
      }
    } catch(eUbic) { Logger.log('_wosCargarStockSnapshot UBICACIONES: ' + eUbic); }

  } catch(eCarmen) { Logger.log('_wosCargarStockSnapshot Carmen open: ' + eCarmen); }

  // ── Paso 3: MASTER 'TABLA_POSICIONES' — mapa SKU_BINID → tipo ──
  var typeMap = {};   // clave: "SKU_BIN_ID" → tipo_almacen string
  try {
    var posRaw = SpreadsheetApp.openById(MASTER_SS_ID)
                   .getSheetByName('TABLA_POSICIONES')
                   .getDataRange().getValues();
    for (var k = 1; k < posRaw.length; k++) {
      var pSku  = String(posRaw[k][COL_CPOS.SKU]          || '').trim().toUpperCase();
      var pBin  = String(posRaw[k][COL_CPOS.BIN_ID]        || '').trim();
      var pTipo = String(posRaw[k][COL_CPOS.TIPO_ALMACEN]  || '').trim().toUpperCase();
      if (!pSku || !pBin) continue;
      typeMap[pSku + '_' + pBin] = pTipo || 'PICKING';
    }
  } catch(eMaster) { Logger.log('_wosCargarStockSnapshot TABLA_POSICIONES: ' + eMaster); }

  // ── Paso 4: inyectar tipo_almacen + ordenar por SKU ──────
  var skus = Object.keys(stockMap);
  for (var s = 0; s < skus.length; s++) {
    var entry = stockMap[skus[s]];
    var cajas = entry.cajas_disponibles;

    // Inyección: reemplaza el 'PICKING' default con el valor real de TABLA_POSICIONES
    for (var c = 0; c < cajas.length; c++) {
      var typeKey = skus[s] + '_' + cajas[c].bin_id;
      if (typeMap[typeKey]) cajas[c].tipo_almacen = typeMap[typeKey];
    }

    // Ordenamiento: PICKING primero, luego alfanumérico natural pasillo/estante/nivel
    cajas.sort(function(a, b) {
      var aPicking = a.tipo_almacen === 'PICKING' ? 0 : 1;
      var bPicking = b.tipo_almacen === 'PICKING' ? 0 : 1;
      if (aPicking !== bPicking) return aPicking - bPicking;
      var aKey = _wosBinSortKey(a.bin_id);
      var bKey = _wosBinSortKey(b.bin_id);
      return aKey < bKey ? -1 : aKey > bKey ? 1 : 0;
    });
  }

  return { stockMap: stockMap, ts: new Date() };
}

// ── Helper de consulta — para las funciones de despacho ─────

/**
 * Devuelve la entrada del stockMap para un SKU, o null si no existe.
 * Llamar con el snapshot ya cargado; no re-abre spreadsheets.
 *
 * @param {Object} stockMap  Resultado de _wosCargarStockSnapshot().stockMap
 * @param {string} sku
 * @return {{ stock_neto_total: Number, cajas_disponibles: Array }|null}
 */
function _wosStockBySku(stockMap, sku) {
  return stockMap[String(sku || '').trim().toUpperCase()] || null;
}

// ── Función pública expuesta al frontend via google.script.run ──

/**
 * Construye y devuelve el snapshot completo de stock para que el
 * frontend pueda renderizar los badges WMS sin ninguna escritura.
 * @return {{ ok: boolean, stockMap: Object, ts: string }|{ ok: false, error: string }}
 */
function WOS_obtenerStockSnapshot() {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    var snap = _wosCargarStockSnapshot();
    return { ok: true, stockMap: snap.stockMap, ts: snap.ts.toISOString() };
  } catch(e) {
    Logger.log('WOS_obtenerStockSnapshot: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}
