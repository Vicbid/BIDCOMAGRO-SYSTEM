// ============================================================
// @version 2.12
//  PORTAL RESELLER — Pedidos de Repuestos (sin garantía)
// ============================================================

var _ACCESORIOS_SS_ID = '1DWjX4JxHskP1uHa7YXTPpbgh2MD35hs43SpUvhP9Vn0';

function _asegurarHojasPedidos() {
  var ss = getDb();

  var hojaPed = ss.getSheetByName(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
  if (!hojaPed) {
    hojaPed = ss.insertSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    hojaPed.appendRow(['ID','Fecha','Reseller','Email','Cant. Ítems','Ítems Sin Catálogo','Estado','Observaciones','Items JSON','PDF URL','Total USD']);
    hojaPed.setFrozenRows(1);
    hojaPed.getRange(1, 1, 1, 11).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hojaPed.setColumnWidth(1, 100);
    hojaPed.setColumnWidth(3, 160);
    hojaPed.setColumnWidth(9, 300);
    hojaPed.setColumnWidth(10, 200);
    hojaPed.setColumnWidth(11, 90);
  }

  var hojaDesc = ss.getSheetByName(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);
  if (!hojaDesc) {
    hojaDesc = ss.insertSheet(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);
    hojaDesc.appendRow(['Fecha','Reseller','SKU / Descripción','Cantidad','Pedido Ref.']);
    hojaDesc.setFrozenRows(1);
    hojaDesc.getRange(1, 1, 1, 5).setBackground('#d63031').setFontColor('#fff').setFontWeight('bold');
    hojaDesc.setColumnWidth(3, 260);
  }

  var hojaKits = ss.getSheetByName(SCHEMA.SHEETS.KITS);
  if (!hojaKits) {
    hojaKits = ss.insertSheet(SCHEMA.SHEETS.KITS);
    hojaKits.appendRow(['Kit SKU','Componente SKU','Cantidad por kit','Descripción (opcional)']);
    hojaKits.setFrozenRows(1);
    hojaKits.getRange(1, 1, 1, 4).setBackground('#5c5fc0').setFontColor('#fff').setFontWeight('bold');
    hojaKits.setColumnWidth(1, 150);
    hojaKits.setColumnWidth(2, 150);
    hojaKits.setColumnWidth(4, 260);
  }
}

// ── Explosión de kits ─────────────────────────────────────────
// Un SKU "kit" que el reseller pide se reemplaza por sus componentes.
// Hoja KITS: A=Kit SKU · B=Componente SKU · C=Cantidad por kit · D=Descripción (opcional).
// Ej: pide 10 de "C" (kit = A + B) → 10 de A + 10 de B. Multiplica por la cantidad pedida.
function _explotarKits(items) {
  if (!items || !items.length) return items;

  var kitsMap = {}, hayKits = false;
  try {
    var dK = getSheetValues(SCHEMA.SHEETS.KITS);
    for (var i = 1; i < dK.length; i++) {
      var kit  = String(dK[i][0] || '').trim().toUpperCase();
      var comp = String(dK[i][1] || '').trim();
      if (!kit || !comp) continue;
      var qk = Number(dK[i][2]) || 1;
      if (qk < 1) qk = 1;
      if (!kitsMap[kit]) kitsMap[kit] = [];
      kitsMap[kit].push({ sku: comp, cantidadPorKit: qk, descripcion: String(dK[i][3] || '').trim() });
      hayKits = true;
    }
  } catch(eK) { Logger.log('_explotarKits leer KITS: ' + eK); return items; }
  if (!hayKits) return items;

  // Descripciones de los componentes desde DB_REPUESTOS (no dependen de lo que ordenó el reseller)
  var descMap = {};
  try {
    var dDb = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var D   = SCHEMA.DB_REPUESTOS;
    for (var d = 1; d < dDb.length; d++) {
      var cod = String(dDb[d][D.CODIGO] || '').trim().toUpperCase();
      if (cod) descMap[cod] = String(dDb[d][D.DESCRIPCION] || '').trim();
    }
  } catch(eDb) { Logger.log('_explotarKits descMap: ' + eDb); }

  var out = [];
  for (var p = 0; p < items.length; p++) {
    var it    = items[p];
    var skuU  = String(it.sku || '').trim().toUpperCase();
    var comps = kitsMap[skuU];
    if (comps && comps.length) {
      var qBase = Number(it.cantidad) || 1;
      for (var c = 0; c < comps.length; c++) {
        var cSkuU = comps[c].sku.toUpperCase();
        out.push({
          sku:         comps[c].sku,
          descripcion: descMap[cSkuU] || comps[c].descripcion || comps[c].sku,
          precio:      0,                                  // se recalcula en el enforcement de precios
          cantidad:    qBase * comps[c].cantidadPorKit,
          modelos:     it.modelos || ''
        });
      }
      Logger.log('_explotarKits: ' + it.sku + ' x' + qBase + ' → ' + comps.length + ' componente(s)');
    } else {
      out.push(it);
    }
  }
  return out;
}

// Mapa de kits para el frontend: { KITSKU: [{sku, cantidadPorKit}, ...] }.
// Permite explotar C → A + B en vivo en el carrito (igual que la sustitución 1→1).
function obtenerKitsPortal() {
  try {
    var kits = {};
    var dK = getSheetValues(SCHEMA.SHEETS.KITS);
    for (var i = 1; i < dK.length; i++) {
      var kit  = String(dK[i][0] || '').trim().toUpperCase();
      var comp = String(dK[i][1] || '').trim();
      if (!kit || !comp) continue;
      var qk = Number(dK[i][2]) || 1;
      if (qk < 1) qk = 1;
      if (!kits[kit]) kits[kit] = [];
      kits[kit].push({ sku: comp, cantidadPorKit: qk });
    }
    return { ok: true, kits: kits };
  } catch(e) {
    Logger.log('obtenerKitsPortal: ' + e);
    return { ok: false, kits: {} };
  }
}

function _siguienteNumeroPedido() {
  _asegurarHojasPedidos();
  var datos = getSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
  var max   = 0;
  for (var i = 1; i < datos.length; i++) {
    var id = String(datos[i][SCHEMA.PEDIDOS_REPUESTOS.ID] || '');
    var m  = id.match(/PR-(\d+)/);
    if (m) max = Math.max(max, parseInt(m[1]));
  }
  var num = String(max + 1);
  while (num.length < 4) num = '0' + num;
  return 'PR-' + num;
}

// ── Búsqueda con estado de stock ──────────────────────────────
function buscarRepuestoConStockPortal(query, reseller, pctOverride) {
  try {
    var q = _normText(String(query || '').trim());
    if (q.length < 2) return { ok: true, items: [] };

    var _descInfo = _descInfoResolve(reseller, pctOverride);
    var _factor   = _descInfo.factor;

    var stockMap = {};
    var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var S = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStock.length; s++) {
      var cod = String(dStock[s][S.CODIGO] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = Number(dStock[s][S.STOCK_ACTUAL]) || 0;
    }

    var priceMap = _buildPriceMap(_factor);
    var fotoMap  = _repFotoMapCatalogo();

    var dDb = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var D   = SCHEMA.DB_REPUESTOS;
    var matches = [];
    for (var i = 1; i < dDb.length; i++) {
      var sku    = String(dDb[i][D.CODIGO]          || '').trim();
      var desc   = String(dDb[i][D.DESCRIPCION]     || '').trim();
      var mod    = String(dDb[i][D.MODELOS]          || '').trim();
      var descEs = String(dDb[i][D.DESCRIPCION_ES]  || '').trim();
      var remplz = String(dDb[i][D.REEMPLAZADO_POR] || '').trim();
      if (!sku && !desc) continue;
      var nSku   = _normText(sku);
      var nDesc  = _normText(desc);
      var nDescEs = _normText(descEs);
      if (nSku.indexOf(q) === -1 && nDesc.indexOf(q) === -1 && nDescEs.indexOf(q) === -1) continue;

      var score;
      if (nSku === q)                                           score = 10;
      else if (nSku.indexOf(q) === 0)                          score = 6;
      else if (nDesc.indexOf(q) === 0 || nDescEs.indexOf(q) === 0) score = 4;
      else                                                      score = 1;

      var skuUp = sku.toUpperCase();
      var estado;
      if (stockMap[skuUp] !== undefined) {
        estado = stockMap[skuUp] > 0 ? 'disponible' : 'backorder';
      } else {
        estado = 'consultar_Backorder';
      }

      matches.push({
        sku:            sku,
        descripcion:    desc,
        descripcionEs:  descEs,
        modelos:        mod,
        estado:         estado,
        precio:         priceMap[skuUp] || 0,
        stockActual:    stockMap[skuUp] !== undefined ? stockMap[skuUp] : null,
        reemplazadoPor: remplz,
        foto:           fotoMap[skuUp] || '',
        _score:         score
      });
    }

    // Buscar también en ACCESORIOS
    try {
      var accSheet2 = SpreadsheetApp.openById(_ACCESORIOS_SS_ID).getSheetByName('ACCESORIOS');
      if (accSheet2) {
        var accRows2 = accSheet2.getDataRange().getValues();
        for (var ai2 = 1; ai2 < accRows2.length; ai2++) {
          var aSku2  = String(accRows2[ai2][0] || '').trim();
          var aDesc2 = String(accRows2[ai2][1] || '').trim();
          if (!aSku2 && !aDesc2) continue;
          var nASku2  = _normText(aSku2);
          var nADesc2 = _normText(aDesc2);
          if (nASku2.indexOf(q) === -1 && nADesc2.indexOf(q) === -1) continue;
          var aScore2;
          if (nASku2 === q)               aScore2 = 10;
          else if (nASku2.indexOf(q) === 0)   aScore2 = 6;
          else if (nADesc2.indexOf(q) === 0)  aScore2 = 4;
          else                                aScore2 = 1;
          var aSkuUp2 = aSku2.toUpperCase();
          var aEst2 = stockMap[aSkuUp2] !== undefined ? (stockMap[aSkuUp2] > 0 ? 'disponible' : 'backorder') : 'consultar_Backorder';
          matches.push({
            sku:            aSku2,
            descripcion:    aDesc2,
            descripcionEs:  '',
            modelos:        String(accRows2[ai2][2] || '').trim(),
            estado:         aEst2,
            precio:         Math.round((Number(accRows2[ai2][3]) || 0) * _factor * 100) / 100,
            stockActual:    stockMap[aSkuUp2] !== undefined ? stockMap[aSkuUp2] : null,
            reemplazadoPor: '',
            foto:           fotoMap[aSkuUp2] || '',
            _score:         aScore2
          });
        }
      }
    } catch(eAcc2) { Logger.log('buscarRepuestoConStockPortal ACCESORIOS: ' + eAcc2); }

    matches.sort(function(a, b) { return b._score - a._score; });
    var items = matches.slice(0, 20);
    for (var k = 0; k < items.length; k++) { delete items[k]._score; }

    return { ok: true, items: items };
  } catch(e) {
    Logger.log('buscarRepuestoConStockPortal: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Saldo real por SKU sumando toda la historia de MOVIMIENTOS_STOCK ─
// Event-sourcing: fuente de verdad absoluta del inventario live.
// Valores positivos = entradas; negativos = salidas / reservas.
function obtenerSaldoRealEnMemoria() {
  var saldos = {};
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M     = SCHEMA.MOVIMIENTOS_STOCK;
    for (var i = 1; i < datos.length; i++) {
      var cod  = String(datos[i][M.CODIGO]   || '').trim().toUpperCase();
      var cant = Number(datos[i][M.CANTIDAD]) || 0;
      if (!cod) continue;
      if (saldos[cod] === undefined) saldos[cod] = 0;
      saldos[cod] += cant;
    }
  } catch(e) {
    Logger.log('obtenerSaldoRealEnMemoria: ' + e);
  }
  return saldos;
}

// ── BIN routing: localiza posición WMS con mayor stock por SKU ────
// Carga TABLA_POSICIONES en memoria de una sola vez.
// Para cada SKU retorna { binId, stock }; clave ausente = sin posición.
function _buildBinMap(skusArray) {
  var binMap = {};
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);
    if (!datos || datos.length < 2) return binMap;
    var T      = SCHEMA.TABLA_POSICIONES;
    var skuSet = {};
    for (var k = 0; k < skusArray.length; k++) {
      var sUp = String(skusArray[k] || '').trim().toUpperCase();
      if (sUp) skuSet[sUp] = true;
    }
    for (var i = 1; i < datos.length; i++) {
      var cod   = String(datos[i][T.CODIGO]      || '').trim().toUpperCase();
      var binId = String(datos[i][T.BIN_ID]      || '').trim();
      var stock = Number(datos[i][T.STOCK_EN_BIN]) || 0;
      if (!cod || !binId || !skuSet[cod]) continue;
      // Priorizar el BIN con mayor cantidad disponible
      if (!binMap[cod] || stock > (binMap[cod].stock || 0)) {
        binMap[cod] = { binId: binId, stock: stock };
      }
    }
  } catch(e) {
    Logger.log('_buildBinMap: ' + e);
  }
  return binMap;
}

// factor = precio que paga el reseller / precio de lista (0.60 = 40% dto). Default 0.60 (histórico).
function _buildPriceMap(factor) {
  var f = (typeof factor === 'number' && !isNaN(factor) && factor > 0) ? factor : 0.60;
  var map = {};
  try {
    var dLista = getSheetValues(SCHEMA.SHEETS.LISTA_REPUESTOS);
    var L = SCHEMA.LISTA_REPUESTOS;
    for (var p = 1; p < dLista.length; p++) {
      var pCod = String(dLista[p][L.CODIGO] || '').trim().toUpperCase();
      var pVal = Number(dLista[p][L.PRECIO]) || 0;
      if (pCod && pVal > 0) map[pCod] = Math.round(pVal * f * 100) / 100; // precio reseller (dto por reseller)
    }
  } catch(e) { Logger.log('_buildPriceMap: ' + e); }
  return map;
}

// ── Descuento por reseller ────────────────────────────────────────────────
// Devuelve { pct, factor } para un reseller. pct = % de descuento sobre lista.
// Fuente: columna de la hoja Resellers localizada por encabezado ("descuento"/"dto");
// si no existe el encabezado, usa el índice fijo SCHEMA.RESELLERS.DESCUENTO.
// Celda vacía, reseller no encontrado o valor inválido → 40% (comportamiento histórico).
// factor = (100 - pct) / 100  (0.60 para 40%, 1.0 para 0%).
function _resellerDescuentoInfo(nombreReseller) {
  var DEFAULT_PCT = 40;
  var pct = DEFAULT_PCT;
  try {
    var nombre = String(nombreReseller || '').trim().toLowerCase();
    if (nombre) {
      var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
      if (datos && datos.length >= 2) {
        var colDesc = -1;
        var header = datos[0];
        for (var j = 0; j < header.length; j++) {
          var h = _normText(header[j]);
          if (h.indexOf('descuento') !== -1 || h.indexOf('dto') !== -1) { colDesc = j; break; }
        }
        if (colDesc === -1) colDesc = SCHEMA.RESELLERS.DESCUENTO;
        for (var i = 1; i < datos.length; i++) {
          if (String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() !== nombre) continue;
          var raw = datos[i][colDesc];
          var s = String(raw === null || raw === undefined ? '' : raw).trim().replace('%', '').replace(',', '.');
          if (s !== '') {
            var n = Number(s);
            if (!isNaN(n)) pct = n;
          }
          break;
        }
      }
    }
  } catch(e) { Logger.log('_resellerDescuentoInfo: ' + e); }
  if (isNaN(pct)) pct = DEFAULT_PCT;
  if (pct < 0)   pct = 0;
  if (pct > 100) pct = 100;
  return { pct: pct, factor: (100 - pct) / 100 };
}

// ── Descuento resuelto (reseller normal vs. super-RTV) ────────────────────
// Si pctOverride es un número válido (0-100) → descuento GLOBAL manual del super-RTV
// (0 = PVP / precio de lista). Si no viene, cae al descuento propio del reseller.
function _descInfoResolve(reseller, pctOverride) {
  if (pctOverride !== undefined && pctOverride !== null && String(pctOverride).trim() !== '') {
    var p = Number(pctOverride);
    if (!isNaN(p)) {
      if (p < 0)   p = 0;
      if (p > 100) p = 100;
      return { pct: p, factor: (100 - p) / 100 };
    }
  }
  return _resellerDescuentoInfo(reseller);
}

// ── Verificación en batch para importación Excel ──────────────
function verificarStockItems(skus) {
  try {
    if (!skus || !skus.length) return {};

    // Mapa de stock
    var stockMap = {};
    try {
      var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
      var SK = SCHEMA.STOCK_INVENTARIO;
      for (var s = 1; s < dStock.length; s++) {
        var cod = String(dStock[s][SK.CODIGO] || '').trim().toUpperCase();
        if (cod) stockMap[cod] = Number(dStock[s][SK.STOCK_ACTUAL]) || 0;
      }
    } catch(eSt) { Logger.log('verificarStockItems stock: ' + eSt); }

    // Set de DB_REPUESTOS
    var dbSet = {};
    var dDb = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var D   = SCHEMA.DB_REPUESTOS;
    for (var i = 1; i < dDb.length; i++) {
      var sku = String(dDb[i][D.CODIGO] || '').trim().toUpperCase();
      if (sku) dbSet[sku] = true;
    }

    var result = {};
    for (var k = 0; k < skus.length; k++) {
      var skuUp = String(skus[k] || '').trim().toUpperCase();
      if (!skuUp) { result[skuUp] = 'sin_catalogo'; continue; }
      if      (stockMap[skuUp] !== undefined) result[skuUp] = stockMap[skuUp] > 0 ? 'disponible' : 'backorder';
      else if (dbSet[skuUp])                  result[skuUp] = 'consultar_Backorder';
      else                                    result[skuUp] = 'sin_catalogo';
    }
    return result;
  } catch(e) {
    Logger.log('verificarStockItems: ' + e);
    return {};
  }
}

// ── Split helper: disponible / backorder por cantidad ────────────
// Si stock > 0 pero < cantPedida → dos filas: stk disponible + resto backorder.
function _splitItemsConStock(items, stockMap, dbSet) {
  var result = [];
  for (var k = 0; k < items.length; k++) {
    var it    = items[k];
    var skuUp = String(it.sku || '').trim().toUpperCase();
    var cant  = Number(it.cantidad) || 1;
    var base  = { sku: it.sku || '', descripcion: it.descripcion || '', precio: Number(it.precio) || 0, modelos: it.modelos || '' };
    if (!skuUp) {
      result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: cant, estado: 'sin_catalogo' });
    } else if (stockMap[skuUp] !== undefined) {
      var stk = stockMap[skuUp];
      if (stk >= cant) {
        result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: cant, estado: 'disponible' });
      } else if (stk > 0) {
        result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: stk,        estado: 'disponible' });
        result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: cant - stk, estado: 'backorder'  });
      } else {
        result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: cant, estado: 'backorder' });
      }
    } else if (dbSet[skuUp]) {
      result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: cant, estado: 'consultar_Backorder' });
    } else {
      result.push({ sku: base.sku, descripcion: base.descripcion, precio: base.precio, modelos: base.modelos, cantidad: cant, estado: 'sin_catalogo' });
    }
  }
  return result;
}

// ── Re-verificación de borrador (retomar pedido) ─────────────────
// Recibe items con cantidad, devuelve itemsConEstado con splits.
function revisarItemsConStock(items) {
  try {
    if (!items || !items.length) return [];
    var stockMap = {};
    try {
      var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
      var SK = SCHEMA.STOCK_INVENTARIO;
      for (var s = 1; s < dStock.length; s++) {
        var cod = String(dStock[s][SK.CODIGO] || '').trim().toUpperCase();
        if (cod) stockMap[cod] = Number(dStock[s][SK.STOCK_ACTUAL]) || 0;
      }
    } catch(eSt) { Logger.log('revisarItemsConStock stock: ' + eSt); }
    var dbSet = {};
    var dDb = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var D   = SCHEMA.DB_REPUESTOS;
    for (var i = 1; i < dDb.length; i++) {
      var sku = String(dDb[i][D.CODIGO] || '').trim().toUpperCase();
      if (sku) dbSet[sku] = true;
    }
    items = _explotarKits(items); // C → A + B antes de chequear stock, para que la revisión muestre los componentes reales
    return _splitItemsConStock(items, stockMap, dbSet);
  } catch(e) {
    Logger.log('revisarItemsConStock: ' + e);
    return [];
  }
}

// ── Paso 1 (legacy stub) — pasthrough sin stock check ni borrador ─
// Mantenido solo por compatibilidad. El flujo activo llama confirmarPedidoPortal directo.
function verificarStockYCrearBorrador(params) {
  try {
    var reseller = String(params.reseller || '').trim();
    var items    = params.items || [];
    if (!reseller || !items.length) return { ok: false, error: 'Datos incompletos.' };
    var itemsConEstado = [];
    for (var i = 0; i < items.length; i++) {
      itemsConEstado.push({
        sku:         items[i].sku         || '',
        descripcion: items[i].descripcion || '',
        cantidad:    items[i].cantidad    || 1,
        precio:      items[i].precio      || 0,
        modelos:     items[i].modelos     || '',
        estado:      'pendiente'
      });
    }
    return { ok: true, numero: 'TMP', itemsConEstado: itemsConEstado };
  } catch(e) {
    Logger.log('verificarStockYCrearBorrador: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Formatea un número en layout regional AR: punto miles, coma decimales ─
function _fmtUsd(n) {
  var num   = Math.abs(Number(n) || 0);
  var fixed = num.toFixed(2);
  var parts = fixed.split('.');
  var intS  = parts[0];
  var decS  = parts[1];
  var result = '';
  var len    = intS.length;
  for (var k = 0; k < len; k++) {
    if (k > 0 && (len - k) % 3 === 0) result += '.';
    result += intS[k];
  }
  return '$' + result + ',' + decS + ' USD';
}

// ── Obtiene datos comerciales del Reseller desde la hoja RESELLERS ────────
function _lookupResellerMeta(nombreReseller) {
  var meta = { nombre: String(nombreReseller || '').trim(), direccion: '', telefono: '', emailRtv: '' };
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var RS    = SCHEMA.RESELLERS;
    var rLow  = meta.nombre.toLowerCase();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][RS.NOMBRE] || '').trim().toLowerCase() !== rLow) continue;
      var dir = String(datos[i][RS.DIRECCION] || '').trim();
      var cp  = String(datos[i][RS.CP]        || '').trim();
      var loc = String(datos[i][RS.LOCALIDAD] || '').trim();
      meta.nombre    = String(datos[i][RS.NOMBRE]   || nombreReseller).trim();
      meta.direccion = dir + (cp  ? ', CP ' + cp  : '') + (loc ? ' — ' + loc : '');
      meta.telefono  = String(datos[i][RS.TELEFONO] || '').trim();
      meta.emailRtv  = String(datos[i][RS.EMAIL_RTV] || '').trim();
      break;
    }
  } catch(e) { Logger.log('_lookupResellerMeta: ' + e); }
  return meta;
}

// ── Confirmar pedido portal — flujo directo, un paso ─────────────
// Parámetros: { reseller, items: [{sku, descripcion, cantidad, precio?, modelos?}],
//              observaciones, formaPago, envio }
// Crea 1 fila por ítem en SOLICITUDES_DESPACHO (estado='Pendiente') y
// 1 fila en PEDIDOS_REPUESTOS (estado='Recibido'). Sin stock check, sin split.
function confirmarPedidoPortal(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    // ── Modo super-RTV: carga a nombre de cualquier cliente con descuento GLOBAL manual ──
    // Autorización server-side: NO se confía en el cliente; se re-verifica el email de sesión.
    var modoSuper = (params && params.modoSuper === true);
    if (modoSuper) {
      var _superEmail = '';
      try { _superEmail = Session.getActiveUser().getEmail(); } catch(eSU) {}
      if (!_esRTVSuper(_superEmail)) return { ok: false, error: 'No autorizado para carga super-RTV.' };
    }

    var reseller  = String((modoSuper ? params.cliente : params.reseller) || '').trim();
    var items     = params.items                || [];
    var obs       = String(params.observaciones || '').trim();
    var formaPago = String(params.formaPago     || '').trim();
    var envio     = String(params.envio         || '').trim();

    if (!reseller || !items.length) return { ok: false, error: 'Datos incompletos.' };

    _asegurarHojasPedidos();

    // ── A0. Sustituir SKUs reemplazados ───────────────────────────
    try {
      var dDbR = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
      var DR   = SCHEMA.DB_REPUESTOS;
      var dbMapR = {};
      for (var di0 = 1; di0 < dDbR.length; di0++) {
        var skuR = String(dDbR[di0][DR.CODIGO] || '').trim();
        if (!skuR) continue;
        dbMapR[skuR.toUpperCase()] = {
          sku:          skuR,
          descripcion:  String(dDbR[di0][DR.DESCRIPCION]     || '').trim(),
          reemplazadoPor: String(dDbR[di0][DR.REEMPLAZADO_POR] || '').trim()
        };
      }
      for (var pi0 = 0; pi0 < items.length; pi0++) {
        var entry = dbMapR[String(items[pi0].sku || '').trim().toUpperCase()];
        if (!entry || !entry.reemplazadoPor) continue;
        var newEntry = dbMapR[entry.reemplazadoPor.toUpperCase()];
        Logger.log('confirmarPedidoPortal: sustituyendo ' + items[pi0].sku + ' → ' + entry.reemplazadoPor);
        items[pi0].sku         = newEntry ? newEntry.sku         : entry.reemplazadoPor;
        items[pi0].descripcion = newEntry ? newEntry.descripcion : items[pi0].descripcion;
      }
    } catch(eSub) { Logger.log('confirmarPedidoPortal sustitucion: ' + eSub); }

    // ── A0b. Explosión de kits (ej: C → A + B, misma cantidad) ────
    items = _explotarKits(items);

    // ── A. Precio enforcement desde Lista_Repuestos (dto del reseller o global del super) ─
    var descInfo = modoSuper ? _descInfoResolve(null, params.descuentoPct) : _resellerDescuentoInfo(reseller);
    var priceMap = _buildPriceMap(descInfo.factor);
    var total = 0;
    for (var pi = 0; pi < items.length; pi++) {
      var itPi   = items[pi];
      var skuKey = String(itPi.sku || '').trim().toUpperCase();
      if (skuKey && priceMap[skuKey] !== undefined) {
        itPi.precio = priceMap[skuKey];
      }
      if ((itPi.precio || 0) > 0) total += itPi.precio * (Number(itPi.cantidad) || 1);
    }

    // ── B. Metadatos del reseller / cliente ───────────────────────
    // Super-RTV: el cliente puede ser externo (no está en Resellers) → usar el email tipeado (si lo hay).
    var emailReseller = '';
    if (modoSuper) {
      emailReseller = String(params.emailCliente || '').trim();
    } else {
      try {
        var dRes = getSheetValues(SCHEMA.SHEETS.RESELLERS);
        var rLow = reseller.toLowerCase();
        for (var ri = 1; ri < dRes.length; ri++) {
          if (String(dRes[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rLow) {
            emailReseller = String(dRes[ri][SCHEMA.RESELLERS.EMAIL] || '').trim();
            break;
          }
        }
      } catch(eR) { Logger.log('confirmarPedidoPortal email: ' + eR); }
    }
    var resellerMeta = _lookupResellerMeta(reseller);

    // Super-RTV: dirección de envío tipeada (el cliente externo no tiene perfil de reseller).
    // Se refleja en el PDF (bloque del cliente) y en observaciones (planilla + mail a logística).
    if (modoSuper && envio !== 'Retiro') {
      var _dirEnv = String(params.direccionEnvio || '').trim();
      if (_dirEnv) {
        resellerMeta.direccion = _dirEnv;
        obs = (obs ? obs + '\n' : '') + 'Dirección de envío: ' + _dirEnv;
      }
    }

    var numero = _siguienteNumeroPedido();

    // ── C. Registrar ítems en Notas de Entrega (nuevo spreadsheet) ──
    var NOTAS_SS_ID = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
    var notasHoja = null;
    try {
      // IMPORTANTE: usar getSheetByName (NO getActiveSheet). getActiveSheet devuelve
      // la última pestaña que alguien dejó seleccionada en el archivo → si no era
      // "Pedidos_resellers", el pedido se escribía en la pestaña equivocada.
      notasHoja = SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName('Pedidos_resellers');
    } catch(eNS) { Logger.log('confirmarPedidoPortal openNotasSS: ' + eNS); }
    // FIX reportado por el usuario: "la fila no aparece en la hoja Pedidos_resellers". Causa
    // más probable: si getSheetByName no encuentra la pestaña, devuelve null SIN tirar
    // excepción (no cae en el catch de arriba) — y nada de lo que sigue chequeaba notasHoja,
    // así que el pedido se "creaba" igual (PDF, mail, fila en PEDIDOS_REPUESTOS) sin dejar
    // NINGÚN rastro en Pedidos_resellers, que es la hoja que usa WOS para despachar.
    // DECISIÓN EXPLÍCITA DEL USUARIO (no abortar): a diferencia de otros casos similares en
    // este sistema (ej. WOS con el descuento de stock en Carmen), acá se prefiere que el
    // pedido se siga creando igual — mail y PDF al reseller sin que note nada — aunque la fila
    // en Pedidos_resellers quede pendiente de reconstruir a mano después. Lo único que se
    // agrega es una alerta a soporte para que el problema no pase inadvertido.
    if (!notasHoja) {
      Logger.log('confirmarPedidoPortal: NO se pudo abrir "Pedidos_resellers" en ' + NOTAS_SS_ID + ' — el pedido ' +
        '(' + reseller + ') sigue su curso normal, pero esta fila NO va a quedar registrada ahí. Revisar RS_debugNotasHoja().');
      try {
        GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR,
          '[Portal Reseller] Pedido sin registrar en Pedidos_resellers',
          'No se pudo abrir la hoja "Pedidos_resellers" (planilla ' + NOTAS_SS_ID + ') al confirmar un pedido de "' + reseller + '".\n\n' +
          'El pedido se creó igual (mail y PDF salieron normal), pero no va a aparecer en esa hoja hasta que se reconstruya a mano.\n' +
          'Los datos completos del pedido quedan respaldados en PEDIDOS_REPUESTOS (columna "Items JSON").\n\n' +
          'Ejecutá RS_debugNotasHoja() desde el editor de Apps Script para ver el detalle del error.');
      } catch(eAlert) { Logger.log('confirmarPedidoPortal alerta notasHoja: ' + eAlert); }
    }

    var hojaDesc = getSheet(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);
    var sinCatalogo = [];

    // Stock en tiempo real de Carmen para columna H
    var stockMapCarmen = {};
    try {
      var dStk = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
      var SK   = SCHEMA.STOCK_INVENTARIO;
      for (var s = 1; s < dStk.length; s++) {
        var skcod = String(dStk[s][SK.CODIGO] || '').trim().toUpperCase();
        if (skcod) stockMapCarmen[skcod] = Number(dStk[s][SK.STOCK_ACTUAL]) || 0;
      }
    } catch(eStk) { Logger.log('confirmarPedidoPortal stockMapCarmen: ' + eStk); }

    var dbSet = {};
    try {
      var dDb = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
      var D   = SCHEMA.DB_REPUESTOS;
      for (var di = 1; di < dDb.length; di++) {
        var dbSku = String(dDb[di][D.CODIGO] || '').trim().toUpperCase();
        if (dbSku) dbSet[dbSku] = true;
      }
    } catch(eDb) {}

    var notasFilas = []; // filas escritas en Notas de Entrega para link de email

    for (var i = 0; i < items.length; i++) {
      var it    = items[i];
      var skuUp = String(it.sku || '').trim().toUpperCase();
      var cant  = Number(it.cantidad) || 1;
      var prec  = Number(it.precio)   || 0;
      var stkH  = stockMapCarmen[skuUp] !== undefined ? stockMapCarmen[skuUp] : '';

      if (skuUp && !dbSet[skuUp]) {
        sinCatalogo.push(it);
        if (hojaDesc) hojaDesc.appendRow([new Date(), reseller, it.sku || it.descripcion, cant, numero]);
      }

      if (notasHoja) {
        var newRow   = notasHoja.getLastRow() + 1;
        notasFilas.push(newRow);
        var stkActual = stockMapCarmen[skuUp] !== undefined ? stockMapCarmen[skuUp] : -1;
        var estadoNota = (stkActual >= 0 && stkActual >= cant) ? 'Confirmado' : 'Pendiente_Revision';
        notasHoja.appendRow([
          numero,                            // A: Número de pedido
          reseller,                          // B: Reseller
          it.sku         || '',              // C: SKU
          it.descripcion || '',              // D: Descripción
          cant,                              // E: Cantidad solicitada
          0,                                 // F: 0 inicial
          '=E' + newRow + '-F' + newRow + '-Z' + newRow,  // G: Fórmula pendiente (SOL - DESP - CANCEL)
          prec,                              // H: Precio con descuento (reseller)
          stkH,                              // I: Stock Carmen al momento de carga
          estadoNota,                        // J: Estado
          new Date(),                        // K: Timestamp
          envio,                             // L: Método de envío
          formaPago,                         // M: Método de pago
          obs                                // N: Observaciones
        ]);
      }

    }

    if (sinCatalogo.length) invalidateSheetValues(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);

    // ── D. PDF ────────────────────────────────────────────────────
    var pdfUrl = _generarPdfPedido(numero, reseller, items, obs, total, resellerMeta, formaPago, envio, descInfo);

    // ── E. Registrar en PEDIDOS_REPUESTOS (estado='Recibido') ─────
    var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var P       = SCHEMA.PEDIDOS_REPUESTOS;
    if (hojaPed) {
      hojaPed.appendRow([
        numero,
        new Date(),
        reseller,
        emailReseller,
        items.length,
        sinCatalogo.length,
        'Recibido',
        obs,
        JSON.stringify(items),
        pdfUrl || '',
        total > 0 ? total : '',
        formaPago,
        envio
      ]);
      invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    }

    // ── F. Email a logística ──────────────────────────────────────
    var threadId = _enviarEmailPedidoPortal(numero, reseller, emailReseller, items, obs, total, pdfUrl, resellerMeta, envio, formaPago, descInfo);

    // ── G. Guardar Thread_ID (string crudo) en columna R ─────────
    // WOS_GmailFlow.js lo lee con COL.THREAD_ID para reply en hilo.
    if (threadId && notasHoja && notasFilas.length) {
      for (var fi = 0; fi < notasFilas.length; fi++) {
        try { notasHoja.getRange(notasFilas[fi], 18).setValue(threadId); } catch(eHL) {}
      }
    }

    return { ok: true, numero: numero, pdfUrl: pdfUrl || '', total: total };

  } catch(e) {
    Logger.log('confirmarPedidoPortal ERROR: ' + e + ' | reseller=' + (params ? String(params.reseller || '') : 'null'));
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eL) { Logger.log('confirmarPedidoPortal releaseLock: ' + eL); }
  }
}

// ─────────────────────────────────────────────────────────────
//  DEBUG — diagnóstico manual reportado por el usuario: "RS_Pedidos no está escribiendo
//  en Pedidos_resellers" (última fila real: 3/8/2026 19:33:58, nada desde entonces).
//  Corré esta función desde el editor de Apps Script (seleccionarla → Ejecutar) y mirá
//  Ver → Registros de ejecución. Prueba, en orden: abrir la planilla, listar sus pestañas,
//  encontrar "Pedidos_resellers", leer la última fila real, y por último ESCRIBIR una fila
//  de prueba y borrarla enseguida — para confirmar si el problema es de PERMISOS/ACCESO
//  (la escritura de prueba también fallaría) o de otra cosa (la prueba manual funciona,
//  pero algo puntual en confirmarPedidoPortal no llega a ejecutarla).
// ─────────────────────────────────────────────────────────────
function RS_debugNotasHoja() {
  var NOTAS_SS_ID = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
  var out = { notasSsId: NOTAS_SS_ID, pasos: [] };
  try {
    var ss = SpreadsheetApp.openById(NOTAS_SS_ID);
    out.pasos.push('openById OK: "' + ss.getName() + '"');
    out.tabs = ss.getSheets().map(function(s) { return s.getName(); });

    var hoja = ss.getSheetByName('Pedidos_resellers');
    if (!hoja) {
      out.pasos.push('getSheetByName("Pedidos_resellers") devolvió null. Mirá "tabs" arriba — ¿está el nombre exacto ahí?');
      Logger.log(JSON.stringify(out, null, 2));
      return out;
    }
    out.pasos.push('getSheetByName("Pedidos_resellers") OK');

    var lastRow = hoja.getLastRow();
    out.lastRow = lastRow;
    out.ultimaFila = lastRow > 0 ? hoja.getRange(lastRow, 1, 1, 14).getValues()[0] : '(hoja vacía)';
    out.pasos.push('Lectura OK. Última fila real (con datos) es la ' + lastRow + ' — mirá "ultimaFila" para ver la fecha/pedido.');

    try {
      var testRow = lastRow + 1;
      var marca   = 'TEST_DIAGNOSTICO_' + new Date().getTime();
      hoja.getRange(testRow, 1).setValue(marca);
      SpreadsheetApp.flush();
      var releido = hoja.getRange(testRow, 1).getValue();
      out.escrituraTest = (releido === marca)
        ? 'OK — escribí "' + marca + '" en la fila ' + testRow + ' y lo pude releer sin problema.'
        : 'RARO — escribí "' + marca + '" pero al releer la celda dio "' + releido + '".';
      hoja.getRange(testRow, 1).clearContent();
      SpreadsheetApp.flush();
      out.pasos.push('Prueba de escritura terminada y limpiada (no queda basura en la hoja).');
    } catch (eW) {
      out.escrituraTestError = eW.toString();
      out.pasos.push('LA ESCRITURA DE PRUEBA FALLÓ — esto es lo importante: ' + eW);
    }
  } catch (e) {
    out.errorFatal = e.toString();
    out.pasos.push('ERROR abriendo la planilla (ni siquiera llegó a la pestaña): ' + e);
  }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// ── Generar pedido directo (legacy / sin revisión) ────────────
function generarPedidoRepuestosPortal(params) {
  try {
    var reseller = String(params.reseller || '').trim();
    var items    = params.items || [];
    var obs      = String(params.observaciones || '').trim();
    var total    = Number(params.total) || 0;

    if (!reseller || !items.length) return { ok: false, error: 'Datos incompletos.' };

    _asegurarHojasPedidos();

    var numero = _siguienteNumeroPedido();

    // Email del reseller
    var emailReseller = '';
    try {
      var dRes = getSheetValues(SCHEMA.SHEETS.RESELLERS);
      var rLow = reseller.trim().toLowerCase();
      for (var ri = 1; ri < dRes.length; ri++) {
        if (String(dRes[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rLow) {
          emailReseller = String(dRes[ri][SCHEMA.RESELLERS.EMAIL] || '').trim();
          break;
        }
      }
    } catch(eR) { Logger.log('generarPedidoRepuestosPortal email: ' + eR); }

    // Loguear ítems sin catálogo
    var sinCatalogo = [];
    var hojaDesc = getSheet(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);
    for (var i = 0; i < items.length; i++) {
      if (items[i].estado === 'sin_catalogo') {
        sinCatalogo.push(items[i]);
        if (hojaDesc) {
          hojaDesc.appendRow([
            new Date(),
            reseller,
            items[i].sku || items[i].descripcion,
            items[i].cantidad,
            numero
          ]);
        }
      }
    }
    if (sinCatalogo.length) invalidateSheetValues(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);

    // Generar PDF en Drive
    var pdfUrl = _generarPdfPedido(numero, reseller, items, obs, total);

    // Registrar pedido
    var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    if (hojaPed) {
      hojaPed.appendRow([
        numero,
        new Date(),
        reseller,
        emailReseller,
        items.length,
        sinCatalogo.length,
        'Enviado',
        obs,
        JSON.stringify(items),
        pdfUrl || '',
        total > 0 ? total : ''
      ]);
      invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    }

    // Enviar email
    _enviarEmailPedidoPortal(numero, reseller, emailReseller, items, obs, total, pdfUrl);

    return { ok: true, numero: numero, pdfUrl: pdfUrl || '' };
  } catch(e) {
    Logger.log('generarPedidoRepuestosPortal: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Mapea estado WOS a los 3 estados simples para el reseller ─
function _mapEstadoWosSimple(estadoWos) {
  var e = String(estadoWos || '').trim();
  if (e === 'Cancelado') return 'cancelado';
  if (e === 'Entregado_Confirmado') return 'recibido';   // ya confirmado por el reseller
  if (e === 'Entregado_Cerrado' || e === 'Listo_Retiro') return 'enviado';
  return 'en_proceso';
}

// ── Confirmar recepción de repuestos desde el portal ─────────
function RS_confirmarRecepcion(numero, reseller) {
  // Lock + guarda de idempotencia: 30 clicks seguidos NO cierran 30 veces ni mandan 30 mails.
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch(eL) { return { ok: false, error: 'Otra confirmación en curso. Reintentá en unos segundos.' }; }
  try {
    var ss    = SpreadsheetApp.openById('1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw');
    var hoja  = ss.getSheetByName('Pedidos_resellers');
    if (!hoja) return { ok: false, error: 'Hoja no encontrada' };
    var datos = hoja.getDataRange().getValues();

    // 1. Escanear el pedido: ¿ya está confirmado? ¿hay algo para cerrar? ¿thread para responder?
    var existePedido = false, yaConfirmado = false, threadId = '';
    for (var g = 1; g < datos.length; g++) {
      if (String(datos[g][0] || '').trim() !== numero) continue;
      existePedido = true;
      var eg = String(datos[g][9] || '').trim();
      if (eg === 'Entregado_Confirmado') yaConfirmado = true;
      if (!threadId && datos[g][17]) threadId = String(datos[g][17]).trim(); // col R (18) = Thread_ID
    }
    if (!existePedido) return { ok: false, error: 'Pedido no encontrado' };
    // Ya estaba confirmado → no re-cierra ni reenvía mail (idempotente)
    if (yaConfirmado) return { ok: true, actualizados: 0, yaConfirmado: true };

    // 2. Cerrar las líneas entregadas
    var actualizados = 0;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0] || '').trim() !== numero) continue;
      var est = String(datos[i][9] || '').trim();
      if (est === 'Entregado_Cerrado' || est === 'Listo_Retiro') {
        var rEst = hoja.getRange(i + 1, 10);
        rEst.clearDataValidations();
        rEst.setValue('Entregado_Confirmado');
        hoja.getRange(i + 1, 19).setValue(new Date()); // FECHA_ESTADO col 19 (S)
        actualizados++;
      }
    }
    // 2b. Marcar la confirmación en la hoja PROPIA del Portal (AUTORITATIVA): el historial la respeta
    //     aunque el cruce con WOS cambie/falle → el botón "Confirmar recepción" no reaparece.
    try { _rsMarcarPedidoConfirmado(numero); } catch(ePP) { Logger.log('RS_confirmarRecepcion mark portal: ' + ePP); }
    SpreadsheetApp.flush();

    // Nada entregado para cerrar en WOS → igual quedó marcado en el Portal; no se manda mail de cierre
    if (actualizados === 0) return { ok: true, actualizados: 0 };

    // 3. Mail de cierre — UNA sola vez (protegido por lock + guarda de arriba)
    try { _rsEnviarCierreRecepcion(numero, reseller, threadId); }
    catch(eM) { Logger.log('RS_confirmarRecepcion mail: ' + eM); }

    // 4. Log
    try {
      var logHoja = ss.getSheetByName('WOS_Log');
      if (logHoja) logHoja.appendRow([new Date(), numero, reseller, 'Recepción confirmada por reseller', reseller, '']);
    } catch(eLog) {}

    return { ok: true, actualizados: actualizados };
  } catch(e) {
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eF) {}
  }
}

// Marca el pedido como confirmado por el reseller en la hoja PROPIA del Portal (PEDIDOS_REPUESTOS),
// de forma AUTORITATIVA. El historial (obtenerHistorialPedidosPortal) respeta este estado por encima
// del cruce con WOS, así el botón "Confirmar recepción" no reaparece aunque el estado en WOS cambie o
// el cruce falle. Idempotente. Una fila por pedido en PEDIDOS_REPUESTOS.
function _rsMarcarPedidoConfirmado(numero) {
  var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
  if (!hojaPed) return;
  var P     = SCHEMA.PEDIDOS_REPUESTOS;
  var datos = hojaPed.getDataRange().getValues();
  var num   = String(numero || '').trim();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][P.ID] || '').trim() !== num) continue;
    hojaPed.getRange(i + 1, P.ESTADO + 1).setValue('Entregado_Confirmado');
    break;
  }
  invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
}

// Mail "pedido cerrado" al reseller. Responde en el hilo del pedido si existe; si no, mail directo.
function _rsEnviarCierreRecepcion(numero, reseller, threadId) {
  var cuerpo =
    "<p style='font-size:13px;color:#555;line-height:1.6;margin:0'>" +
      "Registramos la confirmación de recepción de tu pedido <strong style='color:#00a3e0'>" + numero + "</strong>. " +
      "El pedido queda cerrado. Gracias por trabajar con BIDCOMAGRO." +
    "</p>";
  var htmlBody = _construirEmailHTML('Recepción confirmada — Pedido ' + numero, reseller, cuerpo,
                                     'Pedido ' + numero + ' · ' + reseller + '.');
  var plain = 'Hola ' + reseller + ',\n\nRegistramos la confirmación de recepción del pedido ' + numero +
              '. El pedido queda cerrado. Gracias por trabajar con BIDCOMAGRO.';
  var opts = { htmlBody: htmlBody, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR };

  if (threadId) {
    try { GmailApp.getThreadById(threadId).replyAll(plain, opts); return; }
    catch(eT) { Logger.log('_rsEnviarCierreRecepcion thread ' + threadId + ': ' + eT); }
  }
  var email = _emailReseller(reseller);
  if (email) GmailApp.sendEmail(email, 'Recepción confirmada — Pedido ' + numero, plain, opts);
}

// ── Historial de pedidos por reseller ────────────────────────
function obtenerHistorialPedidosPortal(reseller) {
  try {
    _asegurarHojasPedidos();
    var datos = getSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var rLow  = String(reseller || '').trim().toLowerCase();
    var P     = SCHEMA.PEDIDOS_REPUESTOS;
    var out   = [];
    for (var i = datos.length - 1; i >= 1; i--) {
      var f = datos[i];
      if (String(f[P.RESELLER] || '').trim().toLowerCase() !== rLow) continue;
      var fecha = f[P.FECHA];
      out.push({
        id:         String(f[P.ID]               || '').trim(),
        fecha:      fecha instanceof Date
                      ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
                      : String(fecha || ''),
        cantItems:  Number(f[P.CANT_ITEMS])       || 0,
        sinCatalogo:Number(f[P.CANT_SIN_CATALOGO])|| 0,
        estado:     String(f[P.ESTADO]            || ''),
        obs:        String(f[P.OBSERVACIONES]     || ''),
        pdfUrl:     String(f[P.PDF_URL]           || ''),
        totalUsd:   Number(f[P.TOTAL_USD])        || 0,
        itemsJson:  String(f[P.ITEMS_JSON]        || ''),
        estadoSimple: 'en_proceso'
      });
      if (out.length >= 30) break;
    }

    // Cruzar con WOS (Pedidos_resellers) para obtener el estado real
    try {
      var NOTAS_SS_ID_PR = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
      var notas_ss       = SpreadsheetApp.openById(NOTAS_SS_ID_PR);
      var sheetNames     = notas_ss.getSheets().map(function(s){ return s.getName(); });
      console.log('[HIST] Sheets en NOTAS SS: ' + sheetNames.join(', '));

      var wosHoja = notas_ss.getSheetByName('Pedidos_resellers');
      console.log('[HIST] wosHoja encontrada: ' + (wosHoja !== null));

      if (wosHoja) {
        var wosData = wosHoja.getDataRange().getValues();
        console.log('[HIST] wosData filas: ' + wosData.length);

        // IDs que estamos buscando
        var buscandoIds = out.map(function(o){ return o.id; });
        console.log('[HIST] buscando IDs: ' + buscandoIds.join(', '));

        // Agrupar estados WOS por número de pedido (Col 0 = NUMERO, Col 9 = ESTADO)
        var wosEstados = {};
        var tz = Session.getScriptTimeZone();
        for (var j = 1; j < wosData.length; j++) {
          var num = String(wosData[j][0] || '').trim();
          var est = String(wosData[j][9] || '').trim();
          if (!num) continue;
          if (!wosEstados[num]) wosEstados[num] = { total: 0, enviado: 0, recibido: 0, cancelado: 0, tracking: '', notaEntrega: '', neUrl: '', fechaDespacho: '' };
          wosEstados[num].total++;
          var mapped = _mapEstadoWosSimple(est);
          if (mapped === 'enviado')   wosEstados[num].enviado++;
          if (mapped === 'recibido')  wosEstados[num].recibido++;
          if (mapped === 'cancelado') wosEstados[num].cancelado++;
          // Recoger datos de despacho (primera fila no vacía gana)
          if (!wosEstados[num].tracking     && wosData[j][16]) wosEstados[num].tracking     = String(wosData[j][16]).trim();
          if (!wosEstados[num].notaEntrega  && wosData[j][15]) wosEstados[num].notaEntrega  = String(wosData[j][15]).trim();
          if (!wosEstados[num].neUrl        && wosData[j][22]) wosEstados[num].neUrl        = String(wosData[j][22]).trim();
          if (!wosEstados[num].fechaDespacho && wosData[j][14] instanceof Date) {
            wosEstados[num].fechaDespacho = Utilities.formatDate(wosData[j][14], tz, 'dd/MM/yyyy');
          }
        }
        console.log('[HIST] wosEstados: ' + JSON.stringify(wosEstados).substring(0, 800));

        // Calcular estado simple + datos de despacho por pedido
        for (var k = 0; k < out.length; k++) {
          // Confirmación del reseller registrada en la hoja PROPIA del pedido = AUTORITATIVA: gana sobre
          // lo que diga WOS, así el botón "Confirmar recepción" NO reaparece aunque el cruce con WOS cambie.
          if (out[k].estado === 'Entregado_Confirmado') { out[k].estadoSimple = 'recibido'; continue; }
          var ws = wosEstados[out[k].id];
          if (!ws || !ws.total) {
            // Fallback: usar el estado de la hoja propia del portal
            var ep = out[k].estado || '';
            if (ep === 'Despachado' || ep === 'Entregado') out[k].estadoSimple = 'enviado';
            else if (ep === 'Cancelado')                    out[k].estadoSimple = 'cancelado';
            // else queda 'en_proceso' (Recibido / En proceso / vacío)
            console.log('[HIST] ID "' + out[k].id + '" no en WOS — estadoPortal="' + ep + '" → ' + out[k].estadoSimple);
            out[k]._debug = 'fallback portal: ' + (ep || 'vacío');
            continue;
          }
          console.log('[HIST] ID "' + out[k].id + '" → total=' + ws.total + ' enviado=' + ws.enviado + ' cancelado=' + ws.cancelado);
          out[k]._debug = 'total=' + ws.total + ' enviado=' + ws.enviado + ' recibido=' + ws.recibido + ' cancelado=' + ws.cancelado;
          var noCancel = ws.total - ws.cancelado;
          var entregadas = ws.enviado + ws.recibido; // líneas entregadas (confirmadas o no)
          if (ws.cancelado === ws.total) {
            out[k].estadoSimple = 'cancelado';
          } else if (entregadas === noCancel) {
            // Pedido COMPLETO (todo entregado): si ya lo confirmaron → recibido; si no → enviado (muestra botón)
            out[k].estadoSimple = (ws.recibido > 0) ? 'recibido' : 'enviado';
          } else if (entregadas > 0) {
            out[k].estadoSimple = 'enviado_parcial';
          } else {
            out[k].estadoSimple = 'en_proceso';
          }
          out[k].tracking      = ws.tracking;
          out[k].notaEntrega   = ws.notaEntrega;
          out[k].neUrl         = ws.neUrl;
          out[k].fechaDespacho = ws.fechaDespacho;
        }
      }
    } catch(eWos) {
      Logger.log('[HIST] ERROR WOS lookup: ' + eWos);
    }

    return { ok: true, historial: out };
  } catch(e) {
    Logger.log('obtenerHistorialPedidosPortal: ' + e);
    return { ok: false, historial: [] };
  }
}

// ── Tiempo promedio de envío (recibido → despachado) ──────────
// Promedio en días hábiles (Lun–Vie) desde que se recibe el pedido del reseller
// hasta que se despacha. Solo cuenta pedidos de este reseller efectivamente
// enviados, cruzando PEDIDOS_REPUESTOS (fecha de recepción) con WOS
// (Pedidos_resellers, col 14 = fecha de despacho) — el mismo join que el historial.
function obtenerTiempoPromedioEnvioPortal(reseller) {
  try {
    var rLow = String(reseller || '').trim().toLowerCase();
    if (!rLow) return { ok: true, conStock: { muestras: 0, promedio: null }, backorder: { muestras: 0, promedio: null } };
    var P     = SCHEMA.PEDIDOS_REPUESTOS;
    var datos = getSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);

    // id → fecha de recepción (Date)
    var recibido = {}, hayAlguno = false;
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (String(f[P.RESELLER] || '').trim().toLowerCase() !== rLow) continue;
      var id = String(f[P.ID] || '').trim();
      var fe = f[P.FECHA];
      if (id && fe instanceof Date) { recibido[id] = fe; hayAlguno = true; }
    }
    if (!hayAlguno) return { ok: true, conStock: { muestras: 0, promedio: null }, backorder: { muestras: 0, promedio: null } };

    // id → { enviado, fechaDespacho } desde WOS (Pedidos_resellers)
    var desp = {};
    try {
      var NOTAS_SS_ID_PR = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
      var wosHoja = SpreadsheetApp.openById(NOTAS_SS_ID_PR).getSheetByName('Pedidos_resellers');
      if (wosHoja) {
        var wosData = wosHoja.getDataRange().getValues();
        for (var j = 1; j < wosData.length; j++) {
          var num = String(wosData[j][0] || '').trim();
          if (!num || !recibido[num]) continue;
          if (!desp[num]) desp[num] = { enviado: 0, esBackorder: false, primerDespacho: null, boLineas: 0, boResueltas: 0, boUltimoDespacho: null };
          if (_mapEstadoWosSimple(String(wosData[j][9] || '').trim()) === 'enviado') desp[num].enviado++;

          var fd   = wosData[j][14]; // col O — fecha de despacho de ESTA línea
          var fdOk = (fd instanceof Date);
          // Despacho general (para pedidos con stock, que salen juntos): el más temprano
          if (fdOk && (!desp[num].primerDespacho || fd < desp[num].primerDespacho)) desp[num].primerDespacho = fd;

          // ¿Esta línea estaba en backorder al cargar? (stock Carmen col I < cantidad pedida col E)
          var req     = Number(wosData[j][4]) || 0;
          var snap    = wosData[j][8];
          var snapNum = (snap === '' || snap === null || isNaN(Number(snap))) ? null : Number(snap);
          if (snapNum === null || snapNum < req) {
            desp[num].esBackorder = true;
            desp[num].boLineas++;
            if (fdOk) { // la línea en backorder ya se despachó
              desp[num].boResueltas++;
              // espera real: cuándo salió el ÚLTIMO ítem que estaba en backorder
              if (!desp[num].boUltimoDespacho || fd > desp[num].boUltimoDespacho) desp[num].boUltimoDespacho = fd;
            }
          }
        }
      }
    } catch(eWos) { Logger.log('obtenerTiempoPromedioEnvioPortal WOS: ' + eWos); }

    // Dos promedios separados:
    //  · Con stock  → recibido → primer (único) despacho.
    //  · Backorder  → recibido → despacho del ÚLTIMO ítem que estaba en backorder.
    //    Solo cuenta si TODAS las líneas en backorder ya se despacharon (si no, la espera no terminó).
    var conStock = { suma: 0, n: 0 }, backorder = { suma: 0, n: 0 };
    for (var pid in desp) {
      var d = desp[pid];
      if (!recibido[pid]) continue;
      if (d.esBackorder) {
        if (d.boLineas > 0 && d.boResueltas === d.boLineas && d.boUltimoDespacho) {
          backorder.suma += _diasHabilesEntre(recibido[pid], d.boUltimoDespacho); backorder.n++;
        }
      } else if (d.enviado > 0 && d.primerDespacho) {
        conStock.suma += _diasHabilesEntre(recibido[pid], d.primerDespacho); conStock.n++;
      }
    }
    return {
      ok: true,
      conStock:  { muestras: conStock.n,  promedio: conStock.n  ? Math.round((conStock.suma  / conStock.n)  * 10) / 10 : null },
      backorder: { muestras: backorder.n, promedio: backorder.n ? Math.round((backorder.suma / backorder.n) * 10) / 10 : null }
    };
  } catch(e) {
    Logger.log('obtenerTiempoPromedioEnvioPortal: ' + e);
    return { ok: false, conStock: { muestras: 0, promedio: null }, backorder: { muestras: 0, promedio: null } };
  }
}

// Días hábiles (Lun–Vie) transcurridos entre dos fechas.
// Excluye el día inicial e incluye el final; mismo día o fecha invertida = 0.
function _diasHabilesEntre(desde, hasta) {
  var a = new Date(desde.getFullYear(), desde.getMonth(), desde.getDate());
  var b = new Date(hasta.getFullYear(), hasta.getMonth(), hasta.getDate());
  if (b <= a) return 0;
  var count = 0, cur = new Date(a);
  while (cur < b) {
    cur.setDate(cur.getDate() + 1);
    var dow = cur.getDay(); // 0=Domingo, 6=Sábado
    if (dow !== 0 && dow !== 6) count++;
  }
  return count;
}

// ── Índice compacto para búsqueda local en el browser ────────
// Formato: [sku, desc, mod, statusCode, precio, stock, descEs, remplz, normSku, normDesc, normDescEs]
// statusCode: D=disponible, B=backorder, R=consultar_Backorder
// indices 8-10: strings pre-normalizados para búsqueda sin llamar _normText en cada keystroke
function obtenerIndiceRepuestosPortal(reseller, pctOverride) {
  try {
    var _descInfo = _descInfoResolve(reseller, pctOverride);
    var _factor   = _descInfo.factor;

    var stockMap = {};
    var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var S = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStock.length; s++) {
      var cod = String(dStock[s][S.CODIGO] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = Number(dStock[s][S.STOCK_ACTUAL]) || 0;
    }
    var priceMap = _buildPriceMap(_factor);
    var dDb = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var D   = SCHEMA.DB_REPUESTOS;
    var items = [];
    for (var i = 1; i < dDb.length; i++) {
      var sku    = String(dDb[i][D.CODIGO]          || '').trim();
      var desc   = String(dDb[i][D.DESCRIPCION]     || '').trim();
      var mod    = String(dDb[i][D.MODELOS]          || '').trim();
      var descEs = String(dDb[i][D.DESCRIPCION_ES]  || '').trim();
      var remplz = String(dDb[i][D.REEMPLAZADO_POR] || '').trim();
      if (!sku && !desc) continue;
      var skuUp = sku.toUpperCase();
      var e = stockMap[skuUp] !== undefined
              ? (stockMap[skuUp] > 0 ? 'D' : 'B')
              : 'R';
      items.push([
        sku, desc, mod, e,
        priceMap[skuUp] || 0,
        stockMap[skuUp] !== undefined ? stockMap[skuUp] : -1,
        descEs, remplz,
        _normText(sku), _normText(desc), _normText(descEs)
      ]);
    }

    // Agregar ítems de hoja ACCESORIOS
    try {
      var accSheet = SpreadsheetApp.openById(_ACCESORIOS_SS_ID).getSheetByName('ACCESORIOS');
      if (accSheet) {
        var accRows = accSheet.getDataRange().getValues();
        for (var ai = 1; ai < accRows.length; ai++) {
          var aSku  = String(accRows[ai][0] || '').trim();
          var aDesc = String(accRows[ai][1] || '').trim();
          if (!aSku && !aDesc) continue;
          var aSkuUp = aSku.toUpperCase();
          var aMod   = String(accRows[ai][2] || '').trim();
          var aPvp   = Math.round((Number(accRows[ai][3]) || 0) * _factor * 100) / 100;
          var aE     = stockMap[aSkuUp] !== undefined ? (stockMap[aSkuUp] > 0 ? 'D' : 'B') : 'R';
          items.push([
            aSku, aDesc, aMod, aE,
            aPvp,
            stockMap[aSkuUp] !== undefined ? stockMap[aSkuUp] : -1,
            '', '',
            _normText(aSku), _normText(aDesc), ''
          ]);
        }
      }
    } catch(eAcc) { Logger.log('obtenerIndiceRepuestosPortal ACCESORIOS: ' + eAcc); }

    return { ok: true, items: items, descuentoPct: _descInfo.pct };
  } catch(ex) {
    Logger.log('obtenerIndiceRepuestosPortal: ' + ex);
    return { ok: false, items: [] };
  }
}

// ── Trigger: email al reseller cuando cambia el estado del pedido ─
// Registrar con _crearTriggerPedidos() una sola vez desde el editor GAS.
function _onEditPedidos(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== SCHEMA.SHEETS.PEDIDOS_REPUESTOS) return;

    var col = e.range.getColumn();
    var P   = SCHEMA.PEDIDOS_REPUESTOS;
    if (col !== (P.ESTADO + 1)) return; // col 1-indexed = index 0-indexed + 1

    var nuevoEstado = String(e.value || '').trim();
    if (!nuevoEstado) return;

    var row       = e.range.getRow();
    var rowVals   = sheet.getRange(row, 1, 1, 11).getValues()[0];
    var numero    = String(rowVals[P.ID]           || '');
    var reseller  = String(rowVals[P.RESELLER]     || '');
    var email     = String(rowVals[P.EMAIL]        || '');
    var pdfUrl    = String(rowVals[P.PDF_URL]      || '');
    var totalUsd  = Number(rowVals[P.TOTAL_USD])   || 0;

    if (!email || !numero) return;

    var labels = {
      'En proceso': { titulo: 'Pedido en preparación', icono: '🔧', color: '#e67e22', msg: 'Estamos preparando tu pedido. Te avisaremos cuando esté despachado.' },
      'Despachado': { titulo: '¡Tu pedido fue despachado!', icono: '🚚', color: '#5c5fc0', msg: 'Tu pedido está en camino. Pronto lo vas a recibir.' },
      'Entregado':  { titulo: '¡Pedido entregado!', icono: '✅', color: '#1a9e4a', msg: 'Confirmamos la entrega de tu pedido. ¡Gracias por elegirnos!' },
      'Cancelado':  { titulo: 'Pedido cancelado', icono: '❌', color: '#d63031', msg: 'Tu pedido fue cancelado. Contactanos si tenés alguna consulta.' }
    };
    var cfg = labels[nuevoEstado];
    if (!cfg) return; // no enviar email para "Enviado" (estado inicial)

    var totalBloque = totalUsd > 0
      ? '<p style="margin:0 0 2px;font-size:13px;color:#444">Total del pedido: <strong style="color:#00a3e0">$' + totalUsd.toFixed(2) + ' USD</strong></p>' +
        '<p style="margin:0 0 6px;font-size:10px;color:#999">Precios expresados no incluyen impuestos</p>'
      : '';

    var pdfBloque = pdfUrl
      ? '<div style="margin-top:14px;text-align:center"><a href="' + pdfUrl + '" target="_blank" style="display:inline-flex;align-items:center;gap:8px;padding:10px 22px;background:#00a3e0;color:#fff;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600">📄 Ver PDF del pedido</a></div>'
      : '';

    var cuerpo =
      '<div style="text-align:center;padding:20px 0 10px">' +
        '<span style="font-size:40px">' + cfg.icono + '</span>' +
        '<h2 style="margin:10px 0 4px;color:' + cfg.color + ';font-size:18px">' + cfg.titulo + '</h2>' +
        '<p style="margin:0;font-size:12px;color:#9ba5b4">Pedido <strong style="color:#1a1f2e">' + numero + '</strong></p>' +
      '</div>' +
      '<p style="font-size:14px;color:#444;margin:16px 0">' + cfg.msg + '</p>' +
      totalBloque + pdfBloque;

    var html = _construirEmailHTML(
      cfg.titulo + ' — ' + numero,
      reseller,
      cuerpo,
      'Portal Resellers BIDCOMAGRO · ' + reseller
    );

    GmailApp.sendEmail(email, cfg.icono + ' ' + cfg.titulo + ' — ' + numero, '', {
      htmlBody: html,
      name: PORTAL_CONFIG.NOMBRE_REMITENTE,
      replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR
    });

    try {
      var hojaLog = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
      if (hojaLog) hojaLog.appendRow([new Date(), numero, email, 'Estado Pedido: ' + nuevoEstado, cfg.titulo + ' — ' + numero, 'OK', '']);
    } catch(eL) {}

  } catch(e) {
    Logger.log('_onEditPedidos: ' + e);
  }
}

function _crearTriggerPedidos() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === '_onEditPedidos') {
      Logger.log('Trigger ya existe — no se creó duplicado.');
      return;
    }
  }
  ScriptApp.newTrigger('_onEditPedidos')
    .forSpreadsheet(getDb())
    .onEdit()
    .create();
  Logger.log('Trigger _onEditPedidos creado correctamente.');
}

// ── Borradores activos del reseller ──────────────────────────
function obtenerBorradoresActivos(resellerName) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var P     = SCHEMA.PEDIDOS_REPUESTOS;
    var rLow  = String(resellerName || '').trim().toLowerCase();
    var out   = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (String(f[P.RESELLER] || '').trim().toLowerCase() !== rLow) continue;
      if (String(f[P.ESTADO]   || '').trim()                !== 'Borrador') continue;
      var fecha = f[P.FECHA];
      out.push({
        id:        String(f[P.ID]              || ''),
        fecha:     fecha instanceof Date
                     ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
                     : String(fecha || ''),
        cantItems: Number(f[P.CANT_ITEMS])     || 0,
        obs:       String(f[P.OBSERVACIONES]   || ''),
        itemsJson: String(f[P.ITEMS_JSON]      || '')
      });
    }
    return { ok: true, borradores: out };
  } catch(e) {
    Logger.log('obtenerBorradoresActivos: ' + e);
    return { ok: false, borradores: [] };
  }
}

// ── Cancelar borrador (atómico con LockService) ───────────────
function cancelarBorradorPortal(numero) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var P       = SCHEMA.PEDIDOS_REPUESTOS;
    if (!hojaPed) return { ok: false, error: 'Hoja no encontrada.' };
    var rows = hojaPed.getDataRange().getValues();
    for (var r = 1; r < rows.length; r++) {
      if (String(rows[r][P.ID]     || '') !== numero)    continue;
      if (String(rows[r][P.ESTADO] || '') !== 'Borrador') continue;
      hojaPed.getRange(r + 1, P.ESTADO + 1).setValue('Cancelado');
      invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
      return { ok: true };
    }
    return { ok: false, error: 'Borrador no encontrado o ya procesado.' };
  } catch(e) {
    Logger.log('cancelarBorradorPortal: ' + e);
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}

// ── Log de búsquedas sin resultado (demanda no satisfecha) ────
function registrarBusquedaSinResultado(query, reseller) {
  try {
    _asegurarHojasPedidos();
    var hoja = getSheet(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);
    if (hoja && query) {
      hoja.appendRow([new Date(), reseller || '—', query, 0, 'BÚSQUEDA']);
      invalidateSheetValues(SCHEMA.SHEETS.ITEMS_SIN_CATALOGO);
    }
  } catch(e) {
    Logger.log('registrarBusquedaSinResultado: ' + e);
  }
}

// ── Telemetría: demanda perdida por falta de stock ────────────
function registrarDemandaPerdidaPortal(sku, cantidad, resellerName, accion) {
  try {
    sku          = String(sku          || '').trim().toUpperCase();
    resellerName = String(resellerName || '').trim();
    cantidad     = Number(cantidad)    || 0;
    accion       = String(accion       || 'Borrado_Por_Falta_Stock').trim();
    if (!sku || !resellerName) return;

    var emailReseller = resellerName;
    try {
      var dRes = getSheetValues(SCHEMA.SHEETS.RESELLERS);
      var rLow = resellerName.toLowerCase();
      for (var ri = 1; ri < dRes.length; ri++) {
        var em = String(dRes[ri][SCHEMA.RESELLERS.EMAIL] || '').trim();
        if (String(dRes[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rLow && em) {
          emailReseller = em;
          break;
        }
      }
    } catch(eR) {}

    var stockDisp = 0;
    try {
      var dStk = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
      var SK   = SCHEMA.STOCK_INVENTARIO;
      for (var s = 1; s < dStk.length; s++) {
        if (String(dStk[s][SK.CODIGO] || '').trim().toUpperCase() === sku) {
          stockDisp = Number(dStk[s][SK.STOCK_ACTUAL]) || 0;
          break;
        }
      }
    } catch(eSt) {}

    var ss    = getDb();
    var sheet = ss.getSheetByName(SCHEMA.SHEETS.LOG_DEMANDA_PERDIDA);
    if (!sheet) {
      sheet = ss.insertSheet(SCHEMA.SHEETS.LOG_DEMANDA_PERDIDA);
      sheet.appendRow(['Fecha','Reseller_Email','SKU','Cantidad_Intentada','Stock_Disponible_Momento','Accion_Reseller']);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, 6).setBackground('#2d3436').setFontColor('#fff').setFontWeight('bold');
    }

    sheet.appendRow([new Date(), emailReseller, sku, cantidad, stockDisp, accion]);
  } catch(e) {
    Logger.log('registrarDemandaPerdidaPortal: ' + e);
  }
}

// ── Genera PDF del pedido via SpreadsheetApp → DriveApp.getAs ─────
function _generarPdfPedido(numero, reseller, items, obs, total, resellerMeta, formaPago, envio, descInfo) {
  var tempSs = null;
  try {
    var meta     = resellerMeta || { nombre: String(reseller || ''), direccion: '', telefono: '' };
    var fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    // Descuento aplicado (para revertir a PVP de lista y rotular). Default 40% / 0.60.
    var _dPct    = (descInfo && typeof descInfo.pct    === 'number') ? descInfo.pct    : 40;
    var _dFactor = (descInfo && typeof descInfo.factor === 'number' && descInfo.factor > 0) ? descInfo.factor : 0.60;
    var _dLabel  = _dPct > 0 ? ('precio reseller \xb7 ' + _dPct + '% dto.') : 'precio reseller \xb7 sin descuento';

    // Stock en tiempo real de Carmen
    var _pdfStockMap = {};
    try {
      var _pdfDStk = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
      var _pdfSK   = SCHEMA.STOCK_INVENTARIO;
      for (var _ps = 1; _ps < _pdfDStk.length; _ps++) {
        var _psc = String(_pdfDStk[_ps][_pdfSK.CODIGO] || '').trim().toUpperCase();
        if (_psc) _pdfStockMap[_psc] = Number(_pdfDStk[_ps][_pdfSK.STOCK_ACTUAL]) || 0;
      }
    } catch(_pe) { Logger.log('_generarPdfPedido stockMap: ' + _pe); }

    // ── Crear hoja de cálculo temporal ────────────────────────────
    tempSs = SpreadsheetApp.create('TEMP_PED_' + numero);
    var sheet = tempSs.getActiveSheet();
    sheet.setName('Pedido');

    var ri = 1; // cursor de fila

    // ──────────────────────────────────────────────────────────────
    // 1. CABECERA: Banda azul dividida en dos bloques
    // ──────────────────────────────────────────────────────────────
    // Bloque izquierdo: BIDCOMAGRO + subtítulos (cols 1–4)
    sheet.getRange(ri, 1, 3, 4).merge().setValue('BIDCOMAGRO')
      .setFontSize(20).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#00a3e0').setVerticalAlignment('middle').setHorizontalAlignment('left');
    // Bloque derecho: tipo de doc + número + fecha (col 5–7)
    sheet.getRange(ri,     5, 1, 3).merge().setValue('SOLICITUD DE PEDIDO')
      .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('bottom');
    sheet.getRange(ri + 1, 5, 1, 3).merge().setValue('Nº ' + numero)
      .setFontSize(10).setFontColor('#cce8f4')
      .setBackground('#007ab3').setHorizontalAlignment('right');
    sheet.getRange(ri + 2, 5, 1, 3).merge().setValue('Fecha: ' + fechaStr)
      .setFontSize(9).setFontColor('#cce8f4')
      .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('top');
    sheet.setRowHeight(ri,     24);
    sheet.setRowHeight(ri + 1, 18);
    sheet.setRowHeight(ri + 2, 20);
    ri += 3;

    // ──────────────────────────────────────────────────────────────
    // 2. BLOQUE DE DATOS DEL RESELLER
    // ──────────────────────────────────────────────────────────────
    sheet.setRowHeight(ri, 6); // espacio visual
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    // Cabecera del bloque
    sheet.getRange(ri, 1, 1, 7).merge().setValue('DATOS DEL RESELLER SOLICITANTE')
      .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#2d3436').setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22);
    ri++;

    var resellerRows = [
      ['Nombre / Razón Social', meta.nombre        || '—'],
      ['Dirección Comercial',   meta.direccion      || '—'],
      ['Teléfono de Contacto',  meta.telefono       || '—'],
      ['Forma de Pago',         formaPago            || '—'],
      ['Método de Envío',       envio                || '—']
    ];
    var rBgs2 = ['#f7f8fa', '#ffffff', '#f7f8fa', '#ffffff', '#f7f8fa'];
    for (var rri = 0; rri < resellerRows.length; rri++) {
      sheet.getRange(ri, 1, 1, 2).merge().setValue(resellerRows[rri][0])
        .setFontSize(9).setFontWeight('bold').setFontColor('#5e6778').setBackground(rBgs2[rri]);
      sheet.getRange(ri, 3, 1, 5).merge().setValue(resellerRows[rri][1])
        .setFontSize(9).setFontWeight(rri === 0 ? 'bold' : 'normal')
        .setFontColor('#1a1a2e').setBackground(rBgs2[rri]);
      sheet.setRowHeight(ri, 20);
      ri++;
    }

    // ──────────────────────────────────────────────────────────────
    // 3. TABLA DE ÍTEMS
    // ──────────────────────────────────────────────────────────────
    sheet.setRowHeight(ri, 8);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    var headers = ['SKU', 'Descripción', 'Cant.', 'Estado', 'PVP USD', 'Precio Res. USD', 'Subtotal USD'];
    sheet.getRange(ri, 1, 1, 7).setValues([headers])
      .setBackground('#00a3e0').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22);
    ri++;

    for (var i = 0; i < items.length; i++) {
      var it     = items[i];
      var p      = Number(it.precio) || 0;
      var cant   = Number(it.cantidad) || 1;
      var sub    = p * cant;
      var rowBg  = (i % 2 === 0) ? '#ffffff' : '#f5f7fa';
      var skuKey = String(it.sku || '').trim().toUpperCase();
      var estadoLabel;
      if (it.estado === 'sin_catalogo') {
        estadoLabel = 'Sin catálogo';
      } else if (skuKey && _pdfStockMap[skuKey] !== undefined) {
        var pdfStk = _pdfStockMap[skuKey];
        if      (pdfStk >= cant) estadoLabel = 'En stock';
        else if (pdfStk > 0)     estadoLabel = pdfStk + ' inm. / ' + (cant - pdfStk) + ' pend.';
        else                     estadoLabel = 'Backorder';
      } else if (it.estado === 'backorder') {
        estadoLabel = 'Backorder';
      } else {
        estadoLabel = 'Backorder';
      }
      var pvp    = (p > 0 && _dFactor > 0) ? Math.round(p / _dFactor * 100) / 100 : p;
      var rowVals = [
        it.sku         || '—',
        it.descripcion || '—',
        cant,
        estadoLabel,
        pvp > 0 ? _fmtUsd(pvp) : '—',
        p   > 0 ? _fmtUsd(p)   : '—',
        p   > 0 ? _fmtUsd(sub) : '—'
      ];
      var rng = sheet.getRange(ri, 1, 1, 7);
      rng.setValues([rowVals]).setFontSize(9).setBackground(rowBg).setVerticalAlignment('middle');
      sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#00a3e0'); // SKU destacado
      sheet.getRange(ri, 3).setHorizontalAlignment('center');              // Cant centrada
      sheet.getRange(ri, 5, 1, 3).setHorizontalAlignment('right');        // PVP/Precio/Subtotal derecha
      sheet.setRowHeight(ri, 20);
      ri++;
    }

    // ──────────────────────────────────────────────────────────────
    // 4. FILA TOTAL
    // ──────────────────────────────────────────────────────────────
    if (total > 0) {
      sheet.getRange(ri, 1, 1, 5).merge().setValue('TOTAL ESTIMADO (' + _dLabel + ' \xb7 no incluye impuestos)')
        .setFontSize(9).setFontWeight('bold').setFontColor('#5e6778')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 6, 1, 2).merge().setValue(_fmtUsd(total))
        .setFontSize(11).setFontWeight('bold').setFontColor('#00a3e0')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.setRowHeight(ri, 24);
      ri++;
    }

    // ──────────────────────────────────────────────────────────────
    // 5. OBSERVACIONES (opcional)
    // ──────────────────────────────────────────────────────────────
    if (obs) {
      sheet.setRowHeight(ri, 6);
      sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
      ri++;
      sheet.getRange(ri, 1, 1, 1).setValue('Obs:')
        .setFontSize(9).setFontWeight('bold').setFontColor('#7a5800').setBackground('#fffbe6');
      sheet.getRange(ri, 2, 1, 6).merge().setValue(obs)
        .setFontSize(9).setFontColor('#444444').setBackground('#fffbe6').setWrap(true);
      sheet.setRowHeight(ri, 32);
      ri++;
    }

    // ──────────────────────────────────────────────────────────────
    // 6. AVISOS LEGALES
    // ──────────────────────────────────────────────────────────────
    sheet.setRowHeight(ri, 6);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    sheet.getRange(ri, 1, 1, 7).merge()
      .setValue('⚠️  ADVERTENCIA: Los precios reflejados en la presente solicitud se expresan en USD y están sujetos a modificaciones y ajustes sin previo aviso debido a regulaciones cambiarias y de comercio internacional.')
      .setFontSize(8).setFontWeight('bold').setFontStyle('italic')
      .setFontColor('#c0392b').setBackground('#fff5f5').setWrap(true).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 38);
    ri++;

    sheet.getRange(ri, 1, 1, 7).merge()
      .setValue('Este documento constituye estrictamente una Solicitud de Pedido de Compra (Orden de Compra Provisoria) emitida por el Reseller. No representa una confirmación de stock físico inmediato, reserva de mercancía, ni una obligación de venta/facturación definitiva por parte de BIDCOMAGRO hasta tanto sea auditada y despachada por el personal administrativo.')
      .setFontSize(8).setFontColor('#5e6778').setBackground('#f7f8fa').setWrap(true).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 38);
    ri++;

    // ──────────────────────────────────────────────────────────────
    // 7. PIE DE PÁGINA
    // ──────────────────────────────────────────────────────────────
    sheet.setRowHeight(ri, 4);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#00a3e0');
    ri++;
    sheet.getRange(ri, 1, 1, 7).merge()
      .setValue('Documento generado automáticamente — Portal Resellers BIDCOMAGRO · ' + meta.nombre)
      .setFontSize(8).setFontStyle('italic').setFontColor('#9ba5b4')
      .setHorizontalAlignment('center').setBackground('#ffffff');
    sheet.setRowHeight(ri, 16);

    // Anchos de columna (puntos aproximados en Sheets)
    sheet.setColumnWidth(1, 110); // SKU
    sheet.setColumnWidth(2, 220); // Descripción
    sheet.setColumnWidth(3, 48);  // Cant.
    sheet.setColumnWidth(4, 80);  // Estado
    sheet.setColumnWidth(5, 80);  // PVP USD
    sheet.setColumnWidth(6, 80);  // Precio Reseller
    sheet.setColumnWidth(7, 80);  // Subtotal

    SpreadsheetApp.flush();

    // ── Exportar como PDF ──────────────────────────────────────────
    var ssFile  = DriveApp.getFileById(tempSs.getId());
    var pdfBlob = ssFile.getAs('application/pdf');
    pdfBlob.setName('Pedido_' + numero + '.pdf');

    var pdfFolder = DriveApp.getFolderById('1BTzjhYxjdV-WWsn1hgkeKEQt0v8TsGiw');
    var pdfFile   = pdfFolder.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = pdfFile.getUrl();

    ssFile.setTrashed(true);
    return url;

  } catch(e) {
    Logger.log('_generarPdfPedido: ' + e);
    if (tempSs) {
      try { DriveApp.getFileById(tempSs.getId()).setTrashed(true); } catch(eT) {}
    }
    return null;
  }
}

// ── Busca Thread ID previo en EMAIL_LOGS para un número de pedido ─
function _buscarThreadIdEnLogs(numero) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.EMAIL_LOGS);
    for (var i = datos.length - 1; i >= 1; i--) {
      var ref      = String(datos[i][1] || '').trim();
      var threadId = String(datos[i][6] || '').trim();
      if (ref === numero && threadId) return threadId;
    }
  } catch(e) { Logger.log('_buscarThreadIdEnLogs: ' + e); }
  return '';
}

// ── Email interno con GmailThread — responde en hilo si ya existe ─
// Validación de formato antes de usar un email como destinatario/CC — ver nota en el
// bloque de armado de ccList más abajo (un solo CC con formato inválido tira abajo TODO
// el envío de Gmail, no solo esa dirección).
function _esEmailValido(s) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim());
}

function _enviarEmailPedidoPortal(numero, reseller, emailReseller, items, obs, total, pdfUrl, resellerMeta, envio, formaPago, descInfo) {
  var _dPct    = (descInfo && typeof descInfo.pct === 'number') ? descInfo.pct : 40;
  var _dTotLbl = _dPct > 0 ? ('precio reseller c/' + _dPct + '% dto.') : 'precio reseller (sin descuento)';
  // Declarados fuera del try para que el log final siempre pueda escribirlos
  var asunto      = '[PEDIDO] ' + numero + ' — ' + reseller + ' — ' + items.length + ' ítem(s)';
  var estadoLog   = 'PENDING';
  var threadIdLog = '';

  try {
    // Stock en tiempo real de Carmen — fuente de verdad para el email
    var _emailStockMap = {};
    try {
      var _dStk = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
      var _SK   = SCHEMA.STOCK_INVENTARIO;
      for (var _s = 1; _s < _dStk.length; _s++) {
        var _sc = String(_dStk[_s][_SK.CODIGO] || '').trim().toUpperCase();
        if (_sc) _emailStockMap[_sc] = Number(_dStk[_s][_SK.STOCK_ACTUAL]) || 0;
      }
    } catch(_eStk) { Logger.log('_enviarEmailPedidoPortal stockMap: ' + _eStk); }

    var disponibles = [], parciales = [], backorders = [], consultas = [], desconocidos = [];
    for (var i = 0; i < items.length; i++) {
      var it     = items[i];
      var cant   = Number(it.cantidad) || 1;
      var skuKey = String(it.sku || '').trim().toUpperCase();
      var stkN   = skuKey && (_emailStockMap[skuKey] !== undefined);
      var stk    = stkN ? _emailStockMap[skuKey] : -1;

      if (stkN) {
        if      (stk >= cant) disponibles.push(it);
        else if (stk > 0)     parciales.push(it);
        else                  backorders.push(it);
      } else if (it.estado === 'backorder') {
        backorders.push(it);
      } else if (it.estado === 'consultar_Backorder') {
        consultas.push(it);
      } else if (it.estado === 'sin_catalogo') {
        desconocidos.push(it);
      } else {
        // SKU no encontrado en Carmen ni en catálogo → backorder por precaución
        backorders.push(it);
      }
    }

    function _filaItem(it, label, color) {
      return "<tr>" +
        "<td style='padding:8px 10px;font-family:Consolas,monospace;font-size:12px;color:#00a3e0;border-bottom:1px solid #eee'>" + (it.sku || '—') + "</td>" +
        "<td style='padding:8px 10px;font-size:12px;color:#333;border-bottom:1px solid #eee'>" + (it.descripcion || '—') + "</td>" +
        "<td style='padding:8px 10px;text-align:center;font-weight:700;color:#333;border-bottom:1px solid #eee'>" + (it.cantidad || 1) + "</td>" +
        "<td style='padding:8px 10px;text-align:center;border-bottom:1px solid #eee'>" +
          "<span style='background:" + color + ";padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;color:#fff'>" + label + "</span>" +
        "</td>" +
      "</tr>";
    }

    function _filaItemParcial(it) {
      var cant   = Number(it.cantidad) || 1;
      var skuKey = String(it.sku || '').trim().toUpperCase();
      var realStk = (_emailStockMap[skuKey] !== undefined) ? _emailStockMap[skuKey] : 0;
      var inm  = Math.min(cant, realStk);
      var pen  = cant - inm;
      var label = inm + ' inm. / ' + pen + ' pend.';
      return "<tr>" +
        "<td style='padding:8px 10px;font-family:Consolas,monospace;font-size:12px;color:#00a3e0;border-bottom:1px solid #eee'>" + (it.sku || '—') + "</td>" +
        "<td style='padding:8px 10px;font-size:12px;color:#333;border-bottom:1px solid #eee'>" + (it.descripcion || '—') + "</td>" +
        "<td style='padding:8px 10px;text-align:center;font-weight:700;color:#333;border-bottom:1px solid #eee'>" + cant + "</td>" +
        "<td style='padding:8px 10px;text-align:center;border-bottom:1px solid #eee'>" +
          "<span style='background:#f39c12;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;color:#fff'>" + label + "</span>" +
        "</td>" +
      "</tr>";
    }

    var tableRows = '';
    for (var d = 0; d < disponibles.length;  d++) tableRows += _filaItem(disponibles[d],  'En stock',     '#1a9e4a');
    for (var p = 0; p < parciales.length;    p++) tableRows += _filaItemParcial(parciales[p]);
    for (var b = 0; b < backorders.length;   b++) tableRows += _filaItem(backorders[b],   'Backorder',    '#e67e22');
    for (var c = 0; c < consultas.length;    c++) tableRows += _filaItem(consultas[c],    'Backorder',    '#d4890a');
    for (var u = 0; u < desconocidos.length; u++) tableRows += _filaItem(desconocidos[u], 'Sin catálogo', '#d63031');

    var tableHtml =
      "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:1px solid #dde3ea;border-radius:8px;overflow:hidden'>" +
      "<thead><tr style='background:#f0f5fa'>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>SKU</th>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>Descripción</th>" +
        "<th style='padding:9px 10px;text-align:center;font-size:11px;font-weight:700;color:#888;width:60px'>Cant.</th>" +
        "<th style='padding:9px 10px;text-align:center;font-size:11px;font-weight:700;color:#888;width:110px'>Estado</th>" +
      "</tr></thead><tbody>" + tableRows + "</tbody></table>";

    var resumen = [];
    if (disponibles.length)  resumen.push(disponibles.length  + ' en stock');
    if (parciales.length)    resumen.push(parciales.length    + ' parcial(es)');
    if (backorders.length)   resumen.push(backorders.length   + ' en backorder');
    if (consultas.length)    resumen.push(consultas.length    + ' a Backorder');
    if (desconocidos.length) resumen.push(desconocidos.length + ' sin catálogo (pedido especial)');

    var obsBloque = obs
      ? "<div style='margin-top:14px;background:#fffbe6;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px'>" +
        "<strong style='font-size:11px;color:#7a5800;text-transform:uppercase;letter-spacing:.06em'>Observaciones</strong>" +
        "<p style='margin:5px 0 0;font-size:13px;color:#444'>" + obs + "</p></div>"
      : '';

    var totalBloque = (total && total > 0)
      ? "<div style='text-align:right;margin-top:10px'>" +
          "<span style='font-size:13px;color:#444'>Total estimado (" + _dTotLbl + "): </span>" +
          "<strong style='font-size:15px;color:#00a3e0'>" + _fmtUsd(total) + "</strong>" +
          "<div style='font-size:10px;color:#999;margin-top:3px'>Precios expresados no incluyen impuestos</div>" +
        "</div>"
      : '';

    var pdfBloque = pdfUrl
      ? "<div style='margin-top:14px;text-align:center'>" +
          "<a href='" + pdfUrl + "' target='_blank' style='display:inline-flex;align-items:center;gap:8px;padding:10px 22px;background:#00a3e0;color:#fff;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600'>" +
          "Ver / Descargar PDF del pedido</a>" +
        "</div>"
      : '';

    var discBloque =
      "<div style='margin-top:20px;padding:10px 14px;background:#fff5f5;border-left:4px solid #c0392b;border-radius:4px'>" +
        "<p style='margin:0;font-size:11px;font-weight:700;font-style:italic;color:#c0392b;line-height:1.6'>" +
          "⚠️ ADVERTENCIA: Los precios reflejados en la presente solicitud se expresan en USD y están sujetos a modificaciones y ajustes sin previo aviso debido a regulaciones cambiarias y de comercio internacional." +
        "</p>" +
      "</div>" +
      "<div style='margin-top:8px;padding:10px 14px;background:#f7f8fa;border-left:4px solid #b2bec3;border-radius:4px'>" +
        "<p style='margin:0;font-size:11px;color:#5e6778;line-height:1.6'>" +
          "Este documento constituye estrictamente una Solicitud de Pedido de Compra (Orden de Compra Provisoria) emitida por el Reseller. No representa una confirmación de stock físico inmediato, reserva de mercancía, ni una obligación de venta/facturación definitiva por parte de BIDCOMAGRO hasta tanto sea auditada y despachada por el personal administrativo." +
        "</p>" +
      "</div>";

    var logisticaBloque = (envio || formaPago)
      ? "<div style='margin-top:12px;margin-bottom:4px;display:flex;gap:12px;flex-wrap:wrap'>" +
          (envio     ? "<div style='background:#f0f5fa;border:1px solid #dde3ea;border-radius:6px;padding:8px 14px;font-size:12px'><span style='color:#888;font-size:10px;text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:2px'>Retiro / Envío</span><strong style='color:#1a1a2e'>" + envio + "</strong></div>" : '') +
          (formaPago ? "<div style='background:#f0f5fa;border:1px solid #dde3ea;border-radius:6px;padding:8px 14px;font-size:12px'><span style='color:#888;font-size:10px;text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:2px'>Forma de pago</span><strong style='color:#1a1a2e'>" + formaPago + "</strong></div>" : '') +
        "</div>"
      : '';

    var cuerpo =
      "<p style='font-size:14px;color:#444;margin:0 0 16px'>" +
        "El reseller <strong>" + reseller + "</strong> generó el pedido <strong>" + numero + "</strong> " +
        "con <strong>" + items.length + " ítem(s)</strong>." +
        (resumen.length ? " Resumen: " + resumen.join(' · ') + "." : '') +
      "</p>" +
      tableHtml + totalBloque + logisticaBloque + obsBloque + pdfBloque + discBloque;

    var html = _construirEmailHTML(
      'Pedido de Repuestos — ' + numero,
      'Equipo de Logística',
      cuerpo,
      'Generado desde el Portal Resellers BIDCOMAGRO · ' + reseller + '.'
    );

    // ── Arquitectura GmailThread: reply si existe, nuevo si no ───────
    var existingThread = _buscarThreadIdEnLogs(numero);

    if (existingThread) {
      try {
        var thread = GmailApp.getThreadById(existingThread);
        thread.replyAll('', { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE });
        estadoLog   = 'OK-THREAD';
        threadIdLog = existingThread;
      } catch(eThread) {
        Logger.log('_enviarEmailPedidoPortal replyAll: ' + eThread);
        existingThread = ''; // hilo expirado o inválido — cae al bloque de nuevo envío
      }
    }

    if (!existingThread) {
      // CC: reseller + RTV del reseller + usuarios internos con columna E = 'si'.
      // Se valida el formato ANTES de sumar cada dirección: un solo CC roto (ej. RTV dado de
      // baja y reemplazado por un nombre en la planilla, no por un mail) tira abajo el envío
      // ENTERO — Gmail rechaza el mail completo si cualquier destinatario/CC es inválido, así
      // que el supervisor y el reseller tampoco lo reciben. Mejor descartar esa dirección
      // puntual y dejar rastro en el log que perder el mail para todos.
      var ccList = [];
      var ccInvalidos = [];
      if (emailReseller && emailReseller !== PORTAL_CONFIG.EMAIL_SUPERVISOR) {
        if (_esEmailValido(emailReseller)) ccList.push(emailReseller);
        else ccInvalidos.push('reseller: ' + emailReseller);
      }
      // RTV: viene en resellerMeta; si no se pasó (call legacy) hacer lookup
      try {
        var _meta   = resellerMeta || _lookupResellerMeta(reseller);
        var _rtv    = String(_meta.emailRtv || '').trim();
        if (_rtv && _rtv !== PORTAL_CONFIG.EMAIL_SUPERVISOR && ccList.indexOf(_rtv) === -1) {
          if (_esEmailValido(_rtv)) ccList.push(_rtv);
          else ccInvalidos.push('RTV: ' + _rtv);
        }
      } catch(eRtv) { Logger.log('_enviarEmailPedidoPortal CC RTV: ' + eRtv); }
      try {
        var dUs = getSheetValues(SCHEMA.SHEETS.USUARIOS);
        for (var ui = 1; ui < dUs.length; ui++) {
          var uEmail = String(dUs[ui][1] || '').trim();
          var uNotif = String(dUs[ui][4] || '').trim().toLowerCase();
          if (uEmail && uNotif === 'si' && uEmail !== PORTAL_CONFIG.EMAIL_SUPERVISOR) {
            if (_esEmailValido(uEmail)) ccList.push(uEmail);
            else ccInvalidos.push('interno: ' + uEmail);
          }
        }
      } catch(eUs) { Logger.log('_enviarEmailPedidoPortal CC internos: ' + eUs); }
      if (ccInvalidos.length) Logger.log('_enviarEmailPedidoPortal CC descartados por email inválido: ' + ccInvalidos.join(' | '));
      var ccNota = ccInvalidos.length ? (' (CC descartado por email inválido: ' + ccInvalidos.join(' | ') + ')') : '';
      var opts = { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR };
      if (ccList.length) opts.cc = ccList.join(',');
      var emailEnviado = false;
      try {
        var draft   = GmailApp.createDraft(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, '', opts);
        var sentMsg = draft.send();
        emailEnviado = true;
        estadoLog    = ('OK' + ccNota).substring(0, 200);
        try {
          threadIdLog = sentMsg.getThread().getId();
        } catch(eThread) {
          Logger.log('_enviarEmailPedidoPortal getThread: ' + eThread);
        }
      } catch(eDraft) {
        Logger.log('_enviarEmailPedidoPortal createDraft: ' + eDraft);
        if (!emailEnviado) {
          try {
            GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, '', opts);
            estadoLog = ('OK-FALLBACK' + ccNota).substring(0, 200);
          } catch(eSend) {
            estadoLog = 'ERROR: ' + String(eSend).substring(0, 100);
            Logger.log('_enviarEmailPedidoPortal sendEmail: ' + eSend);
          }
        }
      }
    }

  } catch(e) {
    estadoLog = 'ERROR: ' + String(e).substring(0, 100);
    Logger.log('_enviarEmailPedidoPortal: ' + e);
  }

  // Siempre loguear — incluso en caso de error, para diagnóstico
  try {
    var hojaLog = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
    if (hojaLog) hojaLog.appendRow([new Date(), numero, PORTAL_CONFIG.EMAIL_SUPERVISOR, 'Pedido Repuestos (Portal)', asunto, estadoLog, threadIdLog]);
  } catch(eL) { Logger.log('_enviarEmailPedidoPortal log: ' + eL); }

  return threadIdLog || '';
}
