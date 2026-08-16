// @version 1.2
// ============================================================
//  WOS — Reservas de unidades en camino (compras DJI) + ETA.
//  Extraído de Despacho_Código.js 3.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// ── Consulta de stock para operarios ─────────────────────────
// Parsea "dd/MM/yyyy" o "dd/MM" (año actual) → Date; cualquier otra cosa → null.
function _wosEtaToDate(s) {
  if (s instanceof Date) return isNaN(s.getTime()) ? null : s;
  s = String(s || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (!m) return null;
  var dd = parseInt(m[1], 10), mm = parseInt(m[2], 10) - 1;
  var yy = m[3] ? parseInt(m[3], 10) : (new Date()).getFullYear();
  if (yy < 100) yy += 2000;
  var dt = new Date(yy, mm, dd);
  return isNaN(dt.getTime()) ? null : dt;
}


// Formatea una ETA (Date o texto) a "dd/MM/yyyy" limpio para mostrar. '0'/vacío → ''.
// Robustez: Sheets puede convertir el texto de la celda FECHA_ETA en una fecha real; sin esto
// se mostraría "Tue Aug 04 2026 00:00:00 GMT-0300 (…)".
function _wosEtaFmt(v) {
  var dt = _wosEtaToDate(v);
  if (dt) { try { return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch(e) {} }
  var s = String(v == null ? '' : v).trim();
  return (s === '0') ? '' : s;
}


// Retorna el mapa de unidades en camino { SKU: { total, ocs, batches, etaMin } }.
//   ocs     = { CAS: qty }               (compat con consumidores previos)
//   batches = [{ cas, air, eta, qty }]   ordenados por fecha asc (sin fecha al final)
//   etaMin  = fecha (string) del lote más próximo con ETA
// Versión ligera usada por Lista de Compras (evita cargar todo el stock).
function WOS_getEnCaminoMap() {
  try {
    var master      = SpreadsheetApp.openById(MASTER_SS_ID);
    var enCaminoMap = {};

    var dCAS = master.getSheetByName('COMPRAS_DJI');
    var dDET = master.getSheetByName('COMPRAS_DETALLE');
    if (!dCAS || !dDET) return { ok: true, map: {} };

    var casData = dCAS.getDataRange().getValues();
    var casActivos = {};
    for (var c = 1; c < casData.length; c++) {
      var casId  = String(casData[c][0] || '').trim().toUpperCase();
      var estado = String(casData[c][2] || '').trim();
      if (casId && estado !== 'En depósito' && estado.indexOf('Borrador') < 0) {
        casActivos[casId] = true;
      }
    }

    var detData = dDET.getDataRange().getValues();
    for (var d = 1; d < detData.length; d++) {
      var dCas = String(detData[d][0] || '').trim().toUpperCase();
      var dSku = String(detData[d][1] || '').trim().toUpperCase();
      var dPed = Number(detData[d][3]) || 0;
      var dRec = Number(detData[d][4]) || 0;
      if (!dSku || !casActivos[dCas]) continue;
      var pend = Math.max(0, dPed - dRec);
      if (pend > 0) {
        var dEta = _wosEtaFmt(detData[d][6]);            // FECHA_ETA (col G) → dd/MM/yyyy limpio
        var dAir = String(detData[d][7] || '').trim();   // N_AIR    (col H)
        if (!enCaminoMap[dSku]) enCaminoMap[dSku] = { total: 0, ocs: {}, _bmap: {} };
        enCaminoMap[dSku].total += pend;
        enCaminoMap[dSku].ocs[dCas] = (enCaminoMap[dSku].ocs[dCas] || 0) + pend;
        var _bKey = dCas + '|' + dEta + '|' + dAir;
        var _bm   = enCaminoMap[dSku]._bmap;
        if (!_bm[_bKey]) _bm[_bKey] = { cas: dCas, air: dAir, eta: dEta, qty: 0 };
        _bm[_bKey].qty += pend;
      }
    }

    // Reservas activas (unidades ya prometidas a un reseller que aceptó esperar)
    var _resv = _wosReservasActivas();

    // Consolidar lotes por SKU: ordenar por fecha (sin fecha al final) + etaMin,
    // y calcular disponibilidad restando las reservas activas por CAS.
    for (var _sk in enCaminoMap) {
      var _em  = enCaminoMap[_sk];
      var _tmp = [];
      for (var _bk in _em._bmap) { var _bb = _em._bmap[_bk]; _tmp.push({ b: _bb, dt: _wosEtaToDate(_bb.eta) }); }
      _tmp.sort(function(a, b) {   // fecha precalculada (no re-parsear en cada comparación)
        if (a.dt && b.dt) return a.dt - b.dt;
        if (a.dt) return -1;
        if (b.dt) return 1;
        return 0;
      });
      var _arr = [];
      for (var _ti = 0; _ti < _tmp.length; _ti++) _arr.push(_tmp[_ti].b);
      _em.batches = _arr;
      _em.etaMin  = (_arr.length && _arr[0].eta) ? _arr[0].eta : '';

      // Disponible = lotes menos reservas activas (repartidas por CAS, lote más próximo primero)
      var _rc = _resv.byCasSku[_sk] || {};
      var _rem = {}; for (var _c in _rc) _rem[_c] = _rc[_c];
      var _bDisp = [], _disp = 0, _etaMinDisp = '';
      for (var _bi3 = 0; _bi3 < _arr.length; _bi3++) {
        var _b3   = _arr[_bi3];
        var _take = Math.min(_rem[_b3.cas] || 0, _b3.qty);
        _rem[_b3.cas] = (_rem[_b3.cas] || 0) - _take;
        var _av = _b3.qty - _take;
        if (_av > 0) {
          _bDisp.push({ cas: _b3.cas, air: _b3.air, eta: _b3.eta, qty: _av });
          _disp += _av;
          if (!_etaMinDisp && _b3.eta) _etaMinDisp = _b3.eta;
        }
      }
      _em.batchesDisp = _bDisp;
      _em.disponible  = _disp;
      _em.reservado   = _em.total - _disp;
      _em.etaMinDisp  = _etaMinDisp;
      delete _em._bmap;
    }

    // Stock actual desde CARMEN — cache 5 min, clave compartida con WOS_cargarPedidos
    var stockMap = {};
    try {
      var cache = CacheService.getScriptCache();
      var cached = cache.get('wos_carmen_stock_v1');
      if (cached) {
        stockMap = JSON.parse(cached);
      } else {
        var carmenStock = SpreadsheetApp.openById(CARMEN_SS_ID)
                            .getSheetByName('STOCK').getDataRange().getValues();
        for (var sc = 1; sc < carmenStock.length; sc++) {
          var sCod = String(carmenStock[sc][0] || '').trim().toUpperCase();
          if (sCod) stockMap[sCod] = parseInt(carmenStock[sc][2]) || 0;
        }
        try { cache.put('wos_carmen_stock_v1', JSON.stringify(stockMap), 300); } catch(eCp) {}
      }
    } catch(eSC) { Logger.log('WOS_getEnCaminoMap stockMap: ' + eSC); }

    return { ok: true, map: enCaminoMap, stockMap: stockMap };
  } catch(e) {
    Logger.log('WOS_getEnCaminoMap ERROR: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Wrapper gateado para el único caller directo desde el cliente (panel de en-camino). El resto de
// los callers son internos — algunos desde flujos de cron (WOS_notificarIngresos, WOS_GmailFlow.js)
// donde Session.getActiveUser() no se comporta igual que en una request real — así que
// WOS_getEnCaminoMap en sí queda sin gate (mismo criterio que WOS_procesarRespuestaManual/
// WOS_reporteBackorder, ver HUB/plan de esta ronda).
function WOS_getEnCaminoMapCliente() {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  return WOS_getEnCaminoMap();
}


// ═══ RESERVAS DE UNIDADES EN CAMINO (compras DJI) ═══════════════════════
// Cuando un reseller acepta "Esperar" (Opción A) se reservan las unidades del/los
// lote(s) DJI más próximos, para no prometer las mismas unidades a dos resellers.
// Hoja RESERVAS_EN_CAMINO (MASTER): FECHA·PEDIDO·RESELLER·SKU·CAS·N_AIR·ETA·CANTIDAD·ESTADO
//   ESTADO: Activa (comprometida) · Cumplida (llegó y se reactivó) · Cancelada (Opción B / liberada)
var _WOS_RES_SHEET  = 'RESERVAS_EN_CAMINO';
var _WOS_RES        = { FECHA:0, PEDIDO:1, RESELLER:2, SKU:3, CAS:4, AIR:5, ETA:6, CANTIDAD:7, ESTADO:8 };
var _WOS_RES_HEADER = ['FECHA','PEDIDO','RESELLER','SKU','CAS','N_AIR','ETA','CANTIDAD','ESTADO'];


function _wosResSheet() {
  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var h  = ss.getSheetByName(_WOS_RES_SHEET);
  if (!h) {
    h = ss.insertSheet(_WOS_RES_SHEET);
    h.appendRow(_WOS_RES_HEADER);
    h.setFrozenRows(1);
  }
  return h;
}


var _WOS_RES_CACHE = 'wos_reservas_activas_v1';
function _wosInvalidarReservasCache() {
  try { CacheService.getScriptCache().remove(_WOS_RES_CACHE); } catch(e) {}
}


// Lee reservas activas → { byCasSku:{SKU:{CAS:qty}}, byPedidoSku:{PEDIDO:{SKU:qty}}, total }
// Cacheada 60s (la invalidan reservar/cerrar). bypass=true fuerza lectura fresca (al asignar).
function _wosReservasActivas(bypass) {
  if (!bypass) {
    try {
      var c0 = CacheService.getScriptCache().get(_WOS_RES_CACHE);
      if (c0) return JSON.parse(c0);
    } catch(eC) {}
  }
  var out = { byCasSku: {}, byPedidoSku: {}, total: 0 };
  try {
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (h) {
      var d = h.getDataRange().getValues();
      var R = _WOS_RES;
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][R.ESTADO] || '').trim() !== 'Activa') continue;
        var sku = String(d[i][R.SKU]    || '').trim().toUpperCase();
        var cas = String(d[i][R.CAS]    || '').trim().toUpperCase();
        var ped = String(d[i][R.PEDIDO] || '').trim();
        var q   = Number(d[i][R.CANTIDAD]) || 0;
        if (!sku || q <= 0) continue;
        if (!out.byCasSku[sku])    out.byCasSku[sku]    = {};
        if (!out.byPedidoSku[ped]) out.byPedidoSku[ped] = {};
        out.byCasSku[sku][cas]    = (out.byCasSku[sku][cas]    || 0) + q;
        out.byPedidoSku[ped][sku] = (out.byPedidoSku[ped][sku] || 0) + q;
        out.total += q;
      }
    }
  } catch(e) { Logger.log('_wosReservasActivas: ' + e); }
  if (!bypass) {
    try {
      var pl = JSON.stringify(out);
      if (pl.length < 90000) CacheService.getScriptCache().put(_WOS_RES_CACHE, pl, 60);
    } catch(eP) {}
  }
  return out;
}


// Reserva, para un pedido que aceptó esperar, las unidades pendientes (E−F−Z de sus líneas
// en Backorder) contra los lotes DJI disponibles más próximos. Idempotente: no re-reserva lo ya
// reservado para ese pedido. Debe llamarse DESPUÉS de que las líneas queden en Backorder.
function _wosReservarEnCamino(numero, reseller) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(eL) { Logger.log('_wosReservarEnCamino lock: ' + eL); }
  try {
    numero = String(numero || '').trim();
    if (!numero) return { ok: false };
    _wosInvalidarReservasCache();                 // asignar siempre con datos frescos (evita doble reserva)
    var ec    = WOS_getEnCaminoMap();             // batchesDisp ya descuenta reservas activas
    var ecMap = (ec && ec.ok) ? ec.map : {};
    var resv  = _wosReservasActivas(true);
    var yaRes = resv.byPedidoSku[numero] || {};

    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false };
    var datos = hoja.getDataRange().getValues();
    var needBySku = {};
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      if (String(datos[i][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
      var sku = String(datos[i][COL.SKU] || '').trim().toUpperCase();
      if (!sku) continue;
      var pend = (Number(datos[i][COL.CANT_SOL]) || 0) - (Number(datos[i][COL.CANT_DESP]) || 0) - (Number(datos[i][COL.CANT_CANCEL]) || 0);
      if (pend > 0) needBySku[sku] = (needBySku[sku] || 0) + pend;
    }

    var nuevas = [], ahora = new Date(), totalRes = 0, etaProx = '', etaProxDt = null;
    var porEtaMap = {};   // eta (string) → cantidad reservada con esa fecha — un pedido con
                           // varios SKUs faltantes puede reservar de lotes con ETAs distintas;
                           // "etaProx" (la más próxima) solo alcanza para avisar de UNA fecha,
                           // hace falta el desglose para no decirle al reseller que TODO llega
                           // en la fecha más próxima cuando en realidad una parte llega después.
    for (var s in needBySku) {
      var need = needBySku[s] - (yaRes[s] || 0);   // descuenta lo ya reservado para este pedido
      if (need <= 0) continue;
      var em = ecMap[s];
      if (!em || !em.batchesDisp) continue;         // sin lotes disponibles → no se reserva (queda "a confirmar")
      for (var b = 0; b < em.batchesDisp.length && need > 0; b++) {
        var bt   = em.batchesDisp[b];
        var take = Math.min(need, bt.qty);
        if (take <= 0) continue;
        nuevas.push([ahora, numero, reseller || '', s, bt.cas, bt.air || '', bt.eta || '', take, 'Activa']);
        need -= take; totalRes += take;
        var etaKey = bt.eta || '(sin ETA)';
        porEtaMap[etaKey] = (porEtaMap[etaKey] || 0) + take;
        var _dt = _wosEtaToDate(bt.eta);
        if (_dt && (!etaProxDt || _dt < etaProxDt)) { etaProxDt = _dt; etaProx = bt.eta; }
      }
    }
    if (nuevas.length) {
      var h = _wosResSheet();
      h.getRange(h.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
      SpreadsheetApp.flush();
      _wosInvalidarReservasCache();
    }
    var porEta = Object.keys(porEtaMap).map(function(k) {
      return { eta: k, cantidad: porEtaMap[k], _dt: _wosEtaToDate(k) };
    }).sort(function(a, b) {
      var ta = a._dt ? a._dt.getTime() : Infinity;
      var tb = b._dt ? b._dt.getTime() : Infinity;
      return ta - tb;
    }).map(function(x) { return { eta: x.eta, cantidad: x.cantidad }; });
    return { ok: true, reservas: nuevas.length, cantidad: totalRes, etaProx: etaProx, porEta: porEta };
  } catch(e) { Logger.log('_wosReservarEnCamino: ' + e); return { ok: false, error: e.toString() }; }
  finally { try { lock.releaseLock(); } catch(eR) {} }
}

// Arma el fragmento HTML de confirmación de reservas para el mail al reseller (usado por
// WOS_detectarRespuestasResellers y WOS_procesarRespuestaManual, Opción A). BUG que arregla:
// antes se mostraba una sola fecha (la más próxima, "etaProx") como si TODO lo reservado
// llegara ese día — con un pedido de varios SKUs faltantes, cada uno reservado de un lote
// distinto, eso era falso (ej: 1 u. llega ~04/08 pero otras 5 u. llegan ~16/08 y ~24/08).
// Con 1 sola ETA entre todo lo reservado se muestra 1 línea igual que antes; con más de una,
// se desglosa por fecha (r.porEta ya viene ordenado ascendente por _wosReservarEnCamino).
function _wosMsgReservas(r) {
  if (!r || !r.reservas || !r.porEta || !r.porEta.length) return '';
  if (r.porEta.length === 1) {
    return "<br><strong style='color:#3730a3'>Reservamos " + r.cantidad + " u. del lote que llega ~" + r.porEta[0].eta + ".</strong>";
  }
  var out = "<br><strong style='color:#3730a3'>Reservamos " + r.cantidad + " u., en distintos lotes:</strong>";
  for (var i = 0; i < r.porEta.length; i++) {
    out += "<br>&nbsp;&nbsp;• " + r.porEta[i].cantidad + " u. llega ~" + r.porEta[i].eta;
  }
  return out;
}


// Cambia el estado de las reservas ACTIVAS de un pedido (opcionalmente filtrando por SKU).
// estadoNuevo: 'Cancelada' (Opción B / liberar) | 'Cumplida' (llegó y se reactivó).
function _wosCerrarReservas(numero, skuOpt, estadoNuevo) {
  try {
    numero = String(numero || '').trim();
    if (!numero) return { ok: false };
    var skuF = skuOpt ? String(skuOpt).trim().toUpperCase() : '';
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (!h) return { ok: true, cambiadas: 0 };
    var d = h.getDataRange().getValues();
    var R = _WOS_RES, chg = 0;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][R.ESTADO] || '').trim() !== 'Activa') continue;
      if (String(d[i][R.PEDIDO] || '').trim() !== numero) continue;
      if (skuF && String(d[i][R.SKU] || '').trim().toUpperCase() !== skuF) continue;
      h.getRange(i + 1, R.ESTADO + 1).setValue(estadoNuevo);
      chg++;
    }
    if (chg) { SpreadsheetApp.flush(); _wosInvalidarReservasCache(); }
    return { ok: true, cambiadas: chg };
  } catch(e) { Logger.log('_wosCerrarReservas: ' + e); return { ok: false, error: e.toString() }; }
}


// Estados de línea "cerrados" (ya no se le debe nada al reseller) — mismo criterio que el
// overlay Despacho parcial del front (DP_ESTADOS_CERRADOS). Cualquier otro estado con pend>0
// es deuda viva y justifica mantener una reserva de unidades en camino.
var _WOS_DP_CERRADOS = { 'Entregado_Cerrado':1, 'Entregado_Confirmado':1, 'Listo_Retiro':1, 'Cancelado':1 };

// Libera/cierra las reservas ACTIVAS cuyo pedido ya NO debe unidades de ese SKU (se despachó,
// se canceló, o el reseller nunca volvió). Evita que retengan unidades fantasma. Cuenta como
// deuda viva cualquier línea con pendiente>0 en estado no cerrado (no solo Backorder: el bloqueo
// global del overlay Despacho parcial reserva también para líneas En_Espera/Confirmado/Parcial).
// Idempotente y segura de correr seguido (la llama el detector cada 10 min).
// Cumplida si ese SKU tuvo despacho en el pedido; si no, Cancelada.
function WOS_reconciliarReservas() {
  try {
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (!h) return { ok: true, revisadas: 0, cerradas: 0 };
    var d = h.getDataRange().getValues();
    var R = _WOS_RES;

    // Necesidad viva por (pedido, sku): pendiente en Backorder + si hubo despacho de ese SKU
    var need  = {};
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    for (var hh = 0; hh < hojas.length; hh++) {
      var pd = hojas[hh].getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        var num = String(pd[i][COL.NUMERO] || '').trim();
        var sku = String(pd[i][COL.SKU]    || '').trim().toUpperCase();
        if (!num || !sku) continue;
        var est  = String(pd[i][COL.ESTADO] || '').trim();
        var pend = _WOS_DP_CERRADOS[est] ? 0
          : Math.max(0, (Number(pd[i][COL.CANT_SOL]) || 0) - (Number(pd[i][COL.CANT_DESP]) || 0) - (Number(pd[i][COL.CANT_CANCEL]) || 0));
        if (!need[num]) need[num] = {};
        if (!need[num][sku]) need[num][sku] = { pend: 0, desp: 0 };
        need[num][sku].pend += pend;
        need[num][sku].desp += Number(pd[i][COL.CANT_DESP]) || 0;
      }
    }

    var cerradas = 0, revisadas = 0;
    for (var r = 1; r < d.length; r++) {
      if (String(d[r][R.ESTADO] || '').trim() !== 'Activa') continue;
      revisadas++;
      var pnum = String(d[r][R.PEDIDO] || '').trim();
      var psku = String(d[r][R.SKU]    || '').trim().toUpperCase();
      var info = (need[pnum] && need[pnum][psku]) ? need[pnum][psku] : { pend: 0, desp: 0 };
      if (info.pend > 0) continue;   // sigue necesitando ese SKU en backorder → mantener la reserva
      h.getRange(r + 1, R.ESTADO + 1).setValue(info.desp > 0 ? 'Cumplida' : 'Cancelada');
      cerradas++;
    }
    if (cerradas) { SpreadsheetApp.flush(); _wosInvalidarReservasCache(); }
    if (cerradas) Logger.log('WOS_reconciliarReservas: ' + cerradas + ' cerrada(s) de ' + revisadas + ' activas');
    return { ok: true, revisadas: revisadas, cerradas: cerradas };
  } catch(e) { Logger.log('WOS_reconciliarReservas: ' + e); return { ok: false, error: e.toString() }; }
}
