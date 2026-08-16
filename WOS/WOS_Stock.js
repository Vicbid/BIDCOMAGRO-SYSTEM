// @version 1.3
// ============================================================
//  WOS — Stock: consulta para operarios, despacho parcial/batch,
//  ubicaciones de picking. (Distinto de WOS_StockSnapshot.js, que
//  es el snapshot read-only para integración WMS.)
//  Extraído de Despacho_Código.js 3.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// Datos para el overlay "Despacho parcial": lotes DJI en camino disponibles por SKU (netos de
// reservas activas) + reservas activas con su ETA, para proyectar con qué lote y qué fecha se
// cumple cada línea pendiente y mostrar lo ya bloqueado 🔒.
function WOS_despachoParcialData() {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    var ec = WOS_getEnCaminoMap();
    if (!ec || !ec.ok) return { ok: false, error: (ec && ec.error) || 'en camino' };
    var bySku = {};
    for (var s in ec.map) {
      var em = ec.map[s];
      bySku[s] = { batchesDisp: em.batchesDisp || [], disponible: em.disponible || 0, reservado: em.reservado || 0 };
    }
    var reservas = [];
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (h) {
      var d = h.getDataRange().getValues();
      var R = _WOS_RES;
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][R.ESTADO] || '').trim() !== 'Activa') continue;
        var q = Number(d[i][R.CANTIDAD]) || 0;
        if (q <= 0) continue;
        reservas.push({
          pedido: String(d[i][R.PEDIDO] || '').trim(),
          sku:    String(d[i][R.SKU]    || '').trim().toUpperCase(),
          eta:    _wosEtaFmt(d[i][R.ETA]),
          qty:    q,
          cas:    String(d[i][R.CAS] || '').trim(),
          air:    String(d[i][R.AIR] || '').trim()
        });
      }
    }
    return { ok: true, bySku: bySku, reservas: reservas };
  } catch(e) { Logger.log('WOS_despachoParcialData: ' + e); return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' }; }
}


// Bloquea (reserva) las unidades DJI en camino para TODOS los pedidos con deuda, en orden FIFO
// global (pedido más viejo primero) — la misma asignación que muestra el overlay Despacho parcial.
// Solo reserva la parte NO cubierta por stock actual (la "falta"); lo cubierto por depósito no
// necesita reserva. Idempotente: descuenta lo ya reservado por (pedido, SKU) antes de asignar,
// así el botón se puede tocar las veces que haga falta sin duplicar.
function WOS_bloquearEnCaminoParciales() {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(eL) { Logger.log('WOS_bloquearEnCaminoParciales lock: ' + eL); }
  try {
    _wosInvalidarReservasCache();                 // asignar siempre con datos frescos
    var ec = WOS_getEnCaminoMap();                // batchesDisp ya descuenta reservas activas
    if (!ec || !ec.ok) return { ok: false, error: (ec && ec.error) || 'en camino' };
    var ecMap    = ec.map || {};
    var stockMap = ec.stockMap || {};
    var resv     = _wosReservasActivas(true);

    // Líneas con deuda de ambas hojas (pend>0, estado no cerrado), FIFO por fecha de pedido
    var lineas = [];
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    for (var hh = 0; hh < hojas.length; hh++) {
      var pd = hojas[hh].getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        var num = String(pd[i][COL.NUMERO] || '').trim();
        var sku = String(pd[i][COL.SKU]    || '').trim().toUpperCase();
        if (!num || !sku) continue;
        if (_WOS_DP_CERRADOS[String(pd[i][COL.ESTADO] || '').trim()]) continue;
        var pend = (Number(pd[i][COL.CANT_SOL]) || 0) - (Number(pd[i][COL.CANT_DESP]) || 0) - (Number(pd[i][COL.CANT_CANCEL]) || 0);
        if (pend <= 0) continue;
        var fR = pd[i][COL.FECHA];
        lineas.push({ num: num, res: String(pd[i][COL.RESELLER] || '').trim(), sku: sku,
                      pend: pend, f: (fR instanceof Date) ? fR.getTime() : 0 });
      }
    }
    lineas.sort(function(a, b) { return a.f - b.f; });

    // 1º el stock del depósito cubre lo que puede (sin reserva), 2º la falta toma lotes en camino
    var pool = {};
    var nuevas = [], ahora = new Date(), totalRes = 0, pedSet = {};
    for (var l = 0; l < lineas.length; l++) {
      var ln = lineas[l];
      if (pool[ln.sku] === undefined) pool[ln.sku] = Math.max(0, Number(stockMap[ln.sku]) || 0);
      var deStock = Math.min(ln.pend, pool[ln.sku]);
      pool[ln.sku] -= deStock;
      var falta = ln.pend - deStock;
      if (falta <= 0) continue;
      var yaPed = resv.byPedidoSku[ln.num];
      if (yaPed && yaPed[ln.sku] > 0) {          // ya bloqueado para este pedido → no duplicar
        var usa = Math.min(yaPed[ln.sku], falta);
        yaPed[ln.sku] -= usa;
        falta -= usa;
      }
      if (falta <= 0) continue;
      var em = ecMap[ln.sku];
      if (!em || !em.batchesDisp) continue;       // sin lotes en camino → queda "a confirmar"
      for (var b = 0; b < em.batchesDisp.length && falta > 0; b++) {
        var bt = em.batchesDisp[b];
        if (bt.qty <= 0) continue;
        var take = Math.min(falta, bt.qty);
        bt.qty -= take;                           // consumir el lote para las líneas siguientes
        nuevas.push([ahora, ln.num, ln.res, ln.sku, bt.cas, bt.air || '', bt.eta || '', take, 'Activa']);
        falta -= take; totalRes += take; pedSet[ln.num] = true;
      }
    }
    if (nuevas.length) {
      var h = _wosResSheet();
      h.getRange(h.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
      SpreadsheetApp.flush();
      _wosInvalidarReservasCache();
    }
    var nPed = 0; for (var k in pedSet) nPed++;
    return { ok: true, reservas: nuevas.length, cantidad: totalRes, pedidos: nPed };
  } catch(e) { Logger.log('WOS_bloquearEnCaminoParciales: ' + e); return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' }; }
  finally { try { lock.releaseLock(); } catch(eR) {} }
}


// Carga las ubicaciones WMS de un conjunto de SKUs en una sola lectura.
// Devuelve { ok, map: { SKU: [{ubicacion, cantidad}] } } con locs ordenadas desc por cantidad.
function WOS_cargarUbicacionesPedido(skus) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    var hojaUbic = SpreadsheetApp.openById(CARMEN_SS_ID).getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: true, map: {} };
    var d   = hojaUbic.getDataRange().getValues();
    var set = {};
    for (var s = 0; s < skus.length; s++) set[String(skus[s]).trim().toUpperCase()] = true;
    var map = {};
    for (var i = 1; i < d.length; i++) {
      var sku  = String(d[i][0] || '').trim().toUpperCase();
      var ubic = String(d[i][1] || '').trim();
      var cant = parseFloat(d[i][2]) || 0;
      if (!sku || !ubic || !set[sku]) continue;
      if (!map[sku]) map[sku] = [];
      map[sku].push({ ubicacion: ubic, cantidad: cant });
    }
    // ASC por cantidad: primero los bins con menos stock para vaciarlos antes
    for (var k in map) map[k].sort(function(a, b) { return a.cantidad - b.cantidad; });
    return { ok: true, map: map };
  } catch(e) {
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Lista principal desde CARMEN (hoja STOCK: A=SKU, B=nombre, C=stock).
// Metadatos adicionales (min, ubicacion, modelos) desde STOCK_REPUESTOS en MASTER.
// q: filtro de búsqueda (SKU o descripción), vacío = todos.
// Fotos del catálogo unificado (hoja TODO, misma spreadsheet que usa Portal Reseller
// como LISTA_PRECIOS_SS_ID y Stock Manager para cargar las fotos faltantes — col B =
// código corto, col H = link de Drive). Cache 5 min, mismo criterio que el stockMap
// de Carmen en WOS_getEnCaminoMap — el catálogo de fotos no cambia a cada minuto.
function _wosFotoMapCatalogo() {
  var map = {};
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('wos_fotos_catalogo_v1');
    if (cached) return JSON.parse(cached);
    var hoja = SpreadsheetApp.openById(CATALOGO_REPUESTOS_ID).getSheetByName('TODO');
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var cod  = String(d[i][1] || '').trim().toUpperCase(); // col B = código corto
        var foto = String(d[i][7] || '').trim();               // col H = imagen
        if (cod && foto) map[cod] = foto;
      }
    }
    try { cache.put('wos_fotos_catalogo_v1', JSON.stringify(map), 300); } catch(eCp) {}
  } catch(e) { Logger.log('_wosFotoMapCatalogo: ' + e); }
  return map;
}


// Total de unidades despachadas históricamente por SKU, desde el log "Entregados" de Carmen
// (cada despacho de WOS_despacharCompleto agrega ahí 1 fila por SKU con la cantidad — ver
// WOS_GmailFlow.js). Se usa como "más vendido" para ordenar la vista Fotos del Stock (pedido
// del usuario: "los primeros 20 artículos que muestre quiero que sean los más vendidos").
// Cache 30 min — no hace falta al segundo, solo decide qué mostrar primero por default.
function _wosVentasMap() {
  var map = {};
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('wos_ventas_map_v1');
    if (cached) return JSON.parse(cached);
    var hoja = SpreadsheetApp.openById(CARMEN_SS_ID).getSheetByName('Entregados');
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var sku  = String(d[i][0] || '').trim().toUpperCase();
        var cant = Number(d[i][2]) || 0;
        if (!sku || cant <= 0) continue;
        map[sku] = (map[sku] || 0) + cant;
      }
    }
    try { cache.put('wos_ventas_map_v1', JSON.stringify(map), 1800); } catch(eCp) {}
  } catch(e) { Logger.log('_wosVentasMap: ' + e); }
  return map;
}


function WOS_cargarStock(q) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    // Lista principal: CARMEN hoja STOCK (A=SKU, B=nombre, C=stock actual)
    var carmenSS    = SpreadsheetApp.openById(CARMEN_SS_ID);
    var hojaCarmen  = carmenSS.getSheetByName('STOCK');
    if (!hojaCarmen) return { ok: false, error: 'Hoja STOCK no encontrada en Carmen.' };

    // Abrir MASTER una sola vez para repMap + enCamino
    var master = SpreadsheetApp.openById(MASTER_SS_ID);

    // Ubicaciones WMS desde Carmen UBICACIONES tab: SKU → [{ubicacion, cantidad}]
    var ubicMap = {};
    try {
      var hojaUbic = carmenSS.getSheetByName(CARMEN_UBICACIONES_TAB);
      if (hojaUbic) {
        var dUbic = hojaUbic.getDataRange().getValues();
        for (var u = 1; u < dUbic.length; u++) {
          var uSku  = String(dUbic[u][0] || '').trim().toUpperCase();
          var uUbic = String(dUbic[u][1] || '').trim();
          var uCant = parseFloat(dUbic[u][2]) || 0;
          if (!uSku || !uUbic) continue;
          if (!ubicMap[uSku]) ubicMap[uSku] = [];
          ubicMap[uSku].push({ ubicacion: uUbic, cantidad: uCant });
        }
      }
    } catch(eUbic) { Logger.log('WOS_cargarStock ubicMap: ' + eUbic); }

    // Metadatos extra desde STOCK_REPUESTOS: min(D), categoria(E), ubicacion(F), modelos(G)
    var repMap = {};
    try {
      var repData = master.getSheetByName('STOCK_REPUESTOS').getDataRange().getValues();
      for (var r = 1; r < repData.length; r++) {
        var rCod = String(repData[r][0] || '').trim().toUpperCase();
        if (!rCod) continue;
        repMap[rCod] = {
          minimo:    parseInt(repData[r][3]) || 0,
          categoria: String(repData[r][4] || '').trim(),
          ubicacion: String(repData[r][5] || '').trim(),
          modelos:   String(repData[r][6] || '').trim()
        };
      }
    } catch(eR) { Logger.log('WOS_cargarStock repMap: ' + eR); }

    // Planificación: clasificación (col N) y mínimo (col O × 4) desde PLANILLA DE PLANIFICACION
    var planifMap = {};
    try {
      var cache = CacheService.getScriptCache();
      var cachedP = cache.get('wos_planif_map_v6');
      if (cachedP) {
        planifMap = JSON.parse(cachedP);
      } else {
        var ssPlanif = SpreadsheetApp.openById(PLANIF_SS_ID);
        var hojaPlanif = null;
        var hojas = ssPlanif.getSheets();
        for (var h = 0; h < hojas.length; h++) {
          if (hojas[h].getName().trim().toUpperCase() === 'PLANILLA DE PLANIFICACION') {
            hojaPlanif = hojas[h]; break;
          }
        }
        if (!hojaPlanif) throw new Error('Hoja planif no encontrada');
        var planifData = hojaPlanif.getDataRange().getValues();
        // Filas 1-2 son cabeceras (advertencia + nombres de col); datos desde índice 3
        for (var p = 3; p < planifData.length; p++) {
          var pCod   = String(planifData[p][0]  || '').trim().toUpperCase(); // col A: PN Corto
          var pClase = String(planifData[p][13] || '').trim().toUpperCase(); // col N: Clasificacion
          var pMin   = parseFloat(planifData[p][14]) || 0;                   // col O: Stock mínimo
          if (pCod) planifMap[pCod] = { clase: pClase, minimo: Math.round(pMin) };
        }
        if (Object.keys(planifMap).length > 0) {
          try { cache.put('wos_planif_map_v6', JSON.stringify(planifMap), 300); } catch(eCp) {}
        }
      }
    } catch(ePl) { Logger.log('WOS_cargarStock planifMap ERROR: ' + ePl); }

    // Unidades en camino — reutiliza WOS_getEnCaminoMap para evitar duplicar lógica
    var enCaminoMap = {};
    try {
      var ecRes = WOS_getEnCaminoMap();
      if (ecRes.ok) enCaminoMap = ecRes.map;
    } catch(eEC) { Logger.log('WOS_cargarStock enCamino: ' + eEC); }

    // Fotos del catálogo (hoja TODO) — para mostrar/ver la imagen de los ítems que ya tienen
    var fotoMap = _wosFotoMapCatalogo();

    // Unidades despachadas históricamente por SKU — "más vendido" para la vista Fotos
    var ventasMap = _wosVentasMap();

    var datos  = hojaCarmen.getDataRange().getValues();
    var filtro = String(q || '').trim().toLowerCase();
    var out    = [];

    for (var i = 1; i < datos.length; i++) {
      var cod  = String(datos[i][0] || '').trim();
      var desc = String(datos[i][1] || '').trim();
      if (!cod) continue;
      if (filtro && cod.toLowerCase().indexOf(filtro) < 0 && desc.toLowerCase().indexOf(filtro) < 0) continue;

      var codKey  = cod.toUpperCase();
      var actual  = parseInt(datos[i][2]) || 0;
      var meta    = repMap[codKey] || { minimo: 0, categoria: '', ubicacion: '', modelos: '' };
      var planif  = planifMap[codKey] || { clase: '', minimo: 0 };
      var minimo  = planif.minimo > 0 ? planif.minimo : meta.minimo;
      var estado  = actual <= 0 ? 'CRITICO' : (actual <= minimo ? 'BAJO' : 'OK');
      var ecData  = enCaminoMap[codKey];
      var ecTotal = ecData ? ecData.total : 0;
      var ecOcs   = [];
      if (ecData && ecData.batches) {
        for (var _bi = 0; _bi < ecData.batches.length; _bi++) {
          var _bt = ecData.batches[_bi];
          ecOcs.push(_bt.cas + ' (' + _bt.qty + 'u.)' + (_bt.eta ? ' · llega ~' + _bt.eta : ''));
        }
      } else if (ecData) {
        for (var ocId in ecData.ocs) ecOcs.push(ocId + ' (' + ecData.ocs[ocId] + 'u.)');
      }

      out.push({
        codigo:      cod,
        descripcion: desc,
        stockActual: actual,
        stockMinimo: minimo,
        clase:       planif.clase,
        estado:      estado,
        categoria:   meta.categoria,
        ubicacion:   (function() {
          var arr = ubicMap[codKey];
          if (arr && arr.length) {
            return arr.map(function(u) { return u.ubicacion + ' (' + u.cantidad + 'u.)'; }).join(' · ');
          }
          return meta.ubicacion; // fallback: ítem aún sin mapear en WMS
        })(),
        modelos:     String(datos[i][4] || '').trim(),
        enCamino:    ecTotal,
        enCaminoOcs: ecOcs,
        enCaminoETA: ecData ? (ecData.etaMin || '') : '',
        enCaminoReservado: ecData ? (ecData.reservado || 0) : 0,
        foto:        fotoMap[codKey] || '',
        ventas:      ventasMap[codKey] || 0
      });
    }

    // CRÍTICO → BAJO → OK; dentro de cada estado por código alfanuméricamente
    var orden = { 'CRITICO': 0, 'BAJO': 1, 'OK': 2 };
    out.sort(function(a, b) {
      var od = orden[a.estado] - orden[b.estado];
      return od !== 0 ? od : (a.codigo < b.codigo ? -1 : 1);
    });

    return { ok: true, items: out };
  } catch(e) {
    Logger.log('WOS_cargarStock ERROR: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// ── Despacho en batch (múltiples pedidos, mismos bultos/tracking) ─
// batchItems: [{numero, despachos:[{row,cantDesp}]}]
// El costo se divide en partes iguales entre los pedidos del batch.
function WOS_despacharBatch(batchItems, transportista, bultos, costoEnvio, operario, reqToken) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    if (!batchItems || !batchItems.length) return { ok: false, error: 'Sin pedidos en el batch.' };
    operario = String(operario || '');
    var costoPorPedido = batchItems.length > 1
      ? Math.round((costoEnvio || 0) / batchItems.length * 100) / 100
      : (Number(costoEnvio) || 0);
    var resultados = [];
    for (var i = 0; i < batchItems.length; i++) {
      var item = batchItems[i];
      try {
        // token de idempotencia por pedido (deriva del token del batch) → un reintento no duplica
        var itemToken = reqToken ? (reqToken + '::' + item.numero) : '';
        var res = WOS_despacharCompleto(item.numero, item.despachos, transportista, bultos, costoPorPedido, operario, itemToken);
        resultados.push({ numero: item.numero, ok: res.ok, error: res.error || '', desactualizado: !!res.desactualizado });
      } catch(eI) {
        Logger.log('WOS_despacharBatch item ' + item.numero + ': ' + eI);
        resultados.push({ numero: item.numero, ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' });
      }
    }
    var errores = [];
    for (var j = 0; j < resultados.length; j++) { if (!resultados[j].ok) errores.push(resultados[j]); }
    Logger.log('WOS_despacharBatch: ' + resultados.length + ' pedidos, ' + errores.length + ' errores');
    return { ok: true, resultados: resultados, errores: errores };
  } catch(e) {
    Logger.log('WOS_despacharBatch ERROR: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}
