// ============================================================
//  PORTAL RESELLER — Pedidos de Repuestos (sin garantía)
// ============================================================

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
function buscarRepuestoConStockPortal(query) {
  try {
    var q = _normText(String(query || '').trim());
    if (q.length < 2) return { ok: true, items: [] };

    var stockMap = {};
    var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var S = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStock.length; s++) {
      var cod = String(dStock[s][S.CODIGO] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = Number(dStock[s][S.STOCK_ACTUAL]) || 0;
    }

    var priceMap = _buildPriceMap();

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
        _score:         score
      });
    }

    matches.sort(function(a, b) { return b._score - a._score; });
    var items = matches.slice(0, 15);
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

function _buildPriceMap() {
  var map = {};
  try {
    var dLista = getSheetValues(SCHEMA.SHEETS.LISTA_REPUESTOS);
    var L = SCHEMA.LISTA_REPUESTOS;
    for (var p = 1; p < dLista.length; p++) {
      var pCod = String(dLista[p][L.CODIGO] || '').trim().toUpperCase();
      var pVal = Number(dLista[p][L.PRECIO]) || 0;
      if (pCod && pVal > 0) map[pCod] = Math.round(pVal * 0.60 * 100) / 100; // 40% dto reseller
    }
  } catch(e) { Logger.log('_buildPriceMap: ' + e); }
  return map;
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

    var reseller  = String(params.reseller      || '').trim();
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

    // ── A. Precio enforcement desde Lista_Repuestos ───────────────
    var priceMap = _buildPriceMap();
    var total = 0;
    for (var pi = 0; pi < items.length; pi++) {
      var itPi   = items[pi];
      var skuKey = String(itPi.sku || '').trim().toUpperCase();
      if (skuKey && priceMap[skuKey] !== undefined) {
        itPi.precio = priceMap[skuKey];
      }
      if ((itPi.precio || 0) > 0) total += itPi.precio * (Number(itPi.cantidad) || 1);
    }

    // ── B. Metadatos del reseller ─────────────────────────────────
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
    } catch(eR) { Logger.log('confirmarPedidoPortal email: ' + eR); }
    var resellerMeta = _lookupResellerMeta(reseller);

    var numero = _siguienteNumeroPedido();

    // ── C. Registrar ítems en Notas de Entrega (nuevo spreadsheet) ──
    var NOTAS_SS_ID = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
    var notasHoja = null;
    try {
      notasHoja = SpreadsheetApp.openById(NOTAS_SS_ID).getActiveSheet();
    } catch(eNS) { Logger.log('confirmarPedidoPortal openNotasSS: ' + eNS); }

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
          '=E' + newRow + '-F' + newRow,     // G: Fórmula pendiente
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
    var pdfUrl = _generarPdfPedido(numero, reseller, items, obs, total, resellerMeta, formaPago, envio);

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
    var threadId = _enviarEmailPedidoPortal(numero, reseller, emailReseller, items, obs, total, pdfUrl, resellerMeta, envio, formaPago);

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
  if (e === 'Entregado_Cerrado' || e === 'Listo_Retiro' || e === 'Entregado_Confirmado') return 'enviado';
  return 'en_proceso';
}

// ── Confirmar recepción de repuestos desde el portal ─────────
function RS_confirmarRecepcion(numero, reseller) {
  try {
    var ss    = SpreadsheetApp.openById('1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw');
    var hoja  = ss.getSheetByName('Pedidos_resellers');
    if (!hoja) return { ok: false, error: 'Hoja no encontrada' };
    var datos = hoja.getDataRange().getValues();
    var actualizados = 0;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0] || '').trim() !== numero) continue;
      var est = String(datos[i][9] || '').trim();
      if (est === 'Entregado_Cerrado' || est === 'Listo_Retiro') {
        hoja.getRange(i + 1, 10).setValue('Entregado_Confirmado');
        hoja.getRange(i + 1, 19).setValue(new Date()); // FECHA_ESTADO col 19 (S)
        actualizados++;
      }
    }
    if (actualizados > 0) SpreadsheetApp.flush();
    // Log en WOS_Log
    try {
      var logHoja = ss.getSheetByName('WOS_Log');
      if (logHoja) logHoja.appendRow([new Date(), numero, reseller, 'Recepción confirmada por reseller', reseller, '']);
    } catch(eLog) {}
    return { ok: true, actualizados: actualizados };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
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
          if (!wosEstados[num]) wosEstados[num] = { total: 0, enviado: 0, cancelado: 0, tracking: '', notaEntrega: '', neUrl: '', fechaDespacho: '' };
          wosEstados[num].total++;
          var mapped = _mapEstadoWosSimple(est);
          if (mapped === 'enviado')   wosEstados[num].enviado++;
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
          out[k]._debug = 'total=' + ws.total + ' enviado=' + ws.enviado + ' cancelado=' + ws.cancelado;
          var noCancel = ws.total - ws.cancelado;
          if (ws.cancelado === ws.total) {
            out[k].estadoSimple = 'cancelado';
          } else if (ws.enviado === noCancel) {
            out[k].estadoSimple = 'enviado';
          } else if (ws.enviado > 0) {
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

// ── Índice compacto para búsqueda local en el browser ────────
// Formato: [sku, desc, mod, statusCode, precio, stock, descEs, remplz, normSku, normDesc, normDescEs]
// statusCode: D=disponible, B=backorder, R=consultar_Backorder
// indices 8-10: strings pre-normalizados para búsqueda sin llamar _normText en cada keystroke
function obtenerIndiceRepuestosPortal() {
  try {
    var stockMap = {};
    var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var S = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStock.length; s++) {
      var cod = String(dStock[s][S.CODIGO] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = Number(dStock[s][S.STOCK_ACTUAL]) || 0;
    }
    var priceMap = _buildPriceMap();
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
    return { ok: true, items: items };
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
function _generarPdfPedido(numero, reseller, items, obs, total, resellerMeta, formaPago, envio) {
  var tempSs = null;
  try {
    var meta     = resellerMeta || { nombre: String(reseller || ''), direccion: '', telefono: '' };
    var fechaStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

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
      var pvp    = p > 0 ? Math.round(p / 0.60 * 100) / 100 : 0;
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
      sheet.getRange(ri, 1, 1, 6).merge().setValue('TOTAL ESTIMADO (precio reseller · 40% dto. · no incluye impuestos)')
        .setFontSize(10).setFontWeight('bold').setFontColor('#5e6778')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 7).setValue(_fmtUsd(total))
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
function _enviarEmailPedidoPortal(numero, reseller, emailReseller, items, obs, total, pdfUrl, resellerMeta, envio, formaPago) {
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
          "<span style='font-size:13px;color:#444'>Total estimado (precio reseller c/40% dto.): </span>" +
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
      // CC: reseller + RTV del reseller + usuarios internos con columna E = 'si'
      var ccList = [];
      if (emailReseller && emailReseller !== PORTAL_CONFIG.EMAIL_SUPERVISOR) ccList.push(emailReseller);
      // RTV: viene en resellerMeta; si no se pasó (call legacy) hacer lookup
      try {
        var _meta   = resellerMeta || _lookupResellerMeta(reseller);
        var _rtv    = String(_meta.emailRtv || '').trim();
        if (_rtv && _rtv !== PORTAL_CONFIG.EMAIL_SUPERVISOR && ccList.indexOf(_rtv) === -1) ccList.push(_rtv);
      } catch(eRtv) { Logger.log('_enviarEmailPedidoPortal CC RTV: ' + eRtv); }
      try {
        var dUs = getSheetValues(SCHEMA.SHEETS.USUARIOS);
        for (var ui = 1; ui < dUs.length; ui++) {
          var uEmail = String(dUs[ui][1] || '').trim();
          var uNotif = String(dUs[ui][4] || '').trim().toLowerCase();
          if (uEmail && uNotif === 'si' && uEmail !== PORTAL_CONFIG.EMAIL_SUPERVISOR) ccList.push(uEmail);
        }
      } catch(eUs) { Logger.log('_enviarEmailPedidoPortal CC internos: ' + eUs); }
      var opts = { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR };
      if (ccList.length) opts.cc = ccList.join(',');
      var emailEnviado = false;
      try {
        var draft   = GmailApp.createDraft(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, '', opts);
        var sentMsg = draft.send();
        emailEnviado = true;
        estadoLog    = 'OK';
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
            estadoLog = 'OK-FALLBACK';
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
