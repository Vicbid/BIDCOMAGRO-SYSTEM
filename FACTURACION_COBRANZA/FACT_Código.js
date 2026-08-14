// @version 1.1
// ============================================================
//  FACTURACION_COBRANZA — Backend
//  Lee (solo lectura) los pedidos de repuestos de WOS (NOTAS_SS_ID) y, para los ligados a
//  una OT, el circuito/cliente en HUB_PRO (MASTER_SHEET_ID) — y escribe únicamente en su
//  propia spreadsheet (FACT_SS_ID): 1 fila por pedido con las marcas Facturado/Cobrado.
//  2 roles por email (_CONFIG: FACTURACION_EMAILS / COBRANZA_EMAILS) — cada persona ve y
//  puede accionar solo su propia lista, gateado también del lado servidor (no solo UI).
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('FACT_Index')
    .setTitle('Facturación · Cobranza')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ── Acceso a hojas (memoizado por ejecución) ──────────────── */
var _factSSCache = {};
function _factSS(id) {
  if (!_factSSCache[id]) _factSSCache[id] = SpreadsheetApp.openById(id);
  return _factSSCache[id];
}


/* ── _CONFIG (self-provisioning, mismo patrón que ComandasPedidos) ── */
function _factConfigHoja() {
  var ss = _factSS(FACT_SS_ID);
  var h = ss.getSheetByName(FACT_CONFIG_TAB);
  if (!h) {
    h = ss.insertSheet(FACT_CONFIG_TAB);
    h.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]);
    h.getRange(2, 1, FACT_CONFIG_DEFAULTS.length, 2).setValues(FACT_CONFIG_DEFAULTS);
    h.setFrozenRows(1);
    h.setColumnWidth(1, 210); h.setColumnWidth(2, 460);
  } else {
    var d = h.getDataRange().getValues();
    var existentes = {};
    for (var i = 1; i < d.length; i++) { var k = String(d[i][0] || '').trim(); if (k) existentes[k.toUpperCase()] = true; }
    var faltan = FACT_CONFIG_DEFAULTS.filter(function(kv) { return !existentes[kv[0].toUpperCase()]; });
    if (faltan.length) h.getRange(h.getLastRow() + 1, 1, faltan.length, 2).setValues(faltan);
  }
  return h;
}

function _factConfig() {
  var d = _factConfigHoja().getDataRange().getValues();
  var m = {};
  for (var i = 1; i < d.length; i++) { var k = String(d[i][0] || '').trim(); if (k) m[k.toUpperCase()] = String(d[i][1] || '').trim(); }
  return m;
}


/* ── Roles / auth ───────────────────────────────────────────
   Igual espíritu que _cpUsuarioAutorizado() de ComandasPedidos, pero con 2 listas en vez
   de 1: acá NO hay "vacío = todos" — si una lista está vacía, nadie tiene ese rol todavía
   (hay que cargar los mails en _CONFIG antes de repartir la URL). */
function _factUsuarioActual() {
  try { return String(Session.getActiveUser().getEmail() || '').trim(); } catch (e) { return ''; }
}

function _factListaEmails(clave) {
  var v = String(_factConfig()[clave] || '');
  return v.split(',').map(function(x) { return x.trim().toLowerCase(); }).filter(Boolean);
}

// Rol del usuario que abrió la URL — la UI muestra solo la(s) pestaña(s) que le correspondan.
function FACT_getMiRol() {
  var email = _factUsuarioActual();
  var elow = email.toLowerCase();
  var facturacion = _factListaEmails('FACTURACION_EMAILS');
  var cobranza = _factListaEmails('COBRANZA_EMAILS');
  return {
    email: email,
    esFacturacion: !!email && facturacion.indexOf(elow) > -1,
    esCobranza: !!email && cobranza.indexOf(elow) > -1
  };
}

function _factMensajeNoAutorizado(rol, clave) {
  return 'No estás autorizada para esta acción' + (rol.email ? ' (' + rol.email + ')' : '') +
    '. Pedí que agreguen tu mail a ' + clave + ' en _CONFIG.';
}


/* ── Hoja de tracking propia (self-provisioning) ───────────── */
function _factHoja() {
  var ss = _factSS(FACT_SS_ID);
  var h = ss.getSheetByName(FACT_TAB);
  if (!h) {
    h = ss.insertSheet(FACT_TAB);
    h.getRange(1, 1, 1, FACT_HEADERS.length).setValues([FACT_HEADERS]);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, FACT_HEADERS.length).setFontWeight('bold');
  }
  return h;
}

// Mapa { NUMERO_UPPER: {numero, tipo, cliente, monto, fechaDespacho, facturado, facturadoFecha,
//                        facturadoPor, idVenta, montoFacturado, cobrado, cobradoFecha,
//                        cobradoPor, montoCobrado, monedaCobrado, rowIdx} }
function _factMapTracking() {
  var h = _factHoja();
  var d = h.getDataRange().getValues();
  var m = {};
  for (var i = 1; i < d.length; i++) {
    var num = String(d[i][COL_FACT.NUMERO] || '').trim();
    if (!num) continue;
    m[num.toUpperCase()] = {
      numero: num,
      tipo: String(d[i][COL_FACT.TIPO] || ''),
      cliente: String(d[i][COL_FACT.CLIENTE] || ''),
      monto: Number(d[i][COL_FACT.MONTO]) || 0,
      fechaDespacho: d[i][COL_FACT.FECHA_DESPACHO],
      facturado: String(d[i][COL_FACT.FACTURADO] || ''),
      facturadoFecha: d[i][COL_FACT.FACTURADO_FECHA],
      facturadoPor: String(d[i][COL_FACT.FACTURADO_POR] || ''),
      idVenta: String(d[i][COL_FACT.IDVENTA] || ''),
      montoFacturado: Number(d[i][COL_FACT.MONTO_FACTURADO]) || 0,
      cobrado: String(d[i][COL_FACT.COBRADO] || ''),
      cobradoFecha: d[i][COL_FACT.COBRADO_FECHA],
      cobradoPor: String(d[i][COL_FACT.COBRADO_POR] || ''),
      montoCobrado: Number(d[i][COL_FACT.MONTO_COBRADO]) || 0,
      monedaCobrado: String(d[i][COL_FACT.MONEDA_COBRADO] || ''),
      rowIdx: i + 1
    };
  }
  return m;
}


/* ── OTs de HUB_PRO (solo lectura, memoizado por ejecución) ──
   Resuelve circuito (para excluir Taller — consumo interno, nunca se factura) y cliente de
   los pedidos ligados a una OT. El número de pedido de WOS es "OT-<id>" (_esNumeroOT en
   WOS/Despacho_Env.js); el id de OT en HUB_PRO (SCHEMA.OT.OT) no lleva ese prefijo. */
var _factOtMapCache = null;
function _factOtMap() {
  if (_factOtMapCache) return _factOtMapCache;
  var m = {};
  try {
    var h = _factSS(MASTER_SHEET_ID).getSheetByName(HOJA_OT);
    if (h) {
      var d = h.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var ot = String(d[i][COL_OT.OT] || '').trim();
        if (!ot) continue;
        m[ot.toUpperCase()] = {
          circuito: String(d[i][COL_OT.CIRCUITO] || '').trim(),
          cliente: String(d[i][COL_OT.RESELLER] || '').trim() || String(d[i][COL_OT.CLIENTE] || '').trim()
        };
      }
    }
  } catch (e) { Logger.log('_factOtMap: ' + e); }
  _factOtMapCache = m;
  return m;
}

function _factEsNumeroOT(numero) {
  return String(numero || '').trim().toUpperCase().indexOf('OT-') === 0;
}

function _factInfoOT(numero) {
  var id = String(numero || '').trim().toUpperCase().replace(/^OT-/, '');
  return _factOtMap()[id] || null;
}


/* ── Lectura de pedidos de WOS ──────────────────────────────
   Agrupa por NUMERO (1 pedido = varias líneas/SKU en la hoja de WOS). Un pedido entra acá
   apenas tiene ALGO despachado (CANT_DESP>0 en al menos 1 línea) — no espera a que esté
   100% entregado; el monto refleja lo despachado hasta el momento. Excluye OTs de circuito
   Taller (repuesto usado en el propio taller, nunca se factura a nadie). */
function _factLeerPedidosWOS() {
  var out = {};
  [{ hoja: HOJA_PEDIDOS, tipo: 'Reseller' }, { hoja: HOJA_PEDIDOS_OT, tipo: 'OT' }].forEach(function(cfg) {
    var h = _factSS(NOTAS_SS_ID).getSheetByName(cfg.hoja);
    if (!h) return;
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var f = d[i];
      var numero = String(f[COL.NUMERO] || '').trim();
      if (!numero) continue;
      var cantDesp = Number(f[COL.CANT_DESP]) || 0;
      if (cantDesp <= 0) continue; // nada despachado todavía en esta línea

      var otInfo = null;
      if (cfg.tipo === 'OT') {
        otInfo = _factInfoOT(numero);
        if (otInfo && otInfo.circuito === CIRCUITO_TALLER) continue; // consumo interno, no se factura
      }

      var key = numero.toUpperCase();
      if (!out[key]) {
        out[key] = {
          numero: numero,
          tipo: cfg.tipo,
          cliente: cfg.tipo === 'OT' ? ((otInfo && otInfo.cliente) || String(f[COL.RESELLER] || '')) : String(f[COL.RESELLER] || ''),
          monto: 0,
          fechaDespacho: null
        };
      }
      var precio = Number(f[COL.PRECIO]) || 0;
      out[key].monto += precio * cantDesp;
      var fd = f[COL.FECHA_DESPACHO];
      if (fd instanceof Date && (!out[key].fechaDespacho || fd < out[key].fechaDespacho)) out[key].fechaDespacho = fd;
    }
  });
  return out; // { NUMERO_UPPER: {numero, tipo, cliente, monto, fechaDespacho} }
}


// Upsert: agrega a la hoja de tracking los pedidos de WOS que todavía no tienen fila, y
// refresca tipo/cliente/monto de los que ya la tienen (sin tocar Facturado/Cobrado).
function _factSincronizarTracking() {
  var wos = _factLeerPedidosWOS();
  var track = _factMapTracking();
  var h = _factHoja();
  var nuevas = [];
  Object.keys(wos).forEach(function(k) {
    var w = wos[k];
    var t = track[k];
    if (!t) {
      var fila = new Array(FACT_HEADERS.length).fill('');
      fila[COL_FACT.NUMERO] = w.numero;
      fila[COL_FACT.TIPO] = w.tipo;
      fila[COL_FACT.CLIENTE] = w.cliente;
      fila[COL_FACT.MONTO] = w.monto;
      fila[COL_FACT.FECHA_DESPACHO] = w.fechaDespacho || '';
      nuevas.push(fila);
    } else if (t.monto !== w.monto || t.cliente !== w.cliente || t.tipo !== w.tipo) {
      // No tocar esto si ya se facturó: el monto BASE es solo el punto de partida antes de
      // facturar (por si un pedido sigue recibiendo despachos parciales) — una vez que
      // facturación ya cargó el monto real (con impuesto), ese no se debe pisar solo.
      if (!t.facturado) h.getRange(t.rowIdx, COL_FACT.TIPO + 1, 1, 3).setValues([[w.tipo, w.cliente, w.monto]]);
    }
  });
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
}


/* ── Formato para la UI ─────────────────────────────────────── */
function _factFmtMoneda(n) {
  n = Number(n) || 0;
  return '$ ' + n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function _factFmtFecha(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
}

function _factFmtItem(t) {
  return {
    numero: t.numero, tipo: t.tipo, cliente: t.cliente,
    monto: t.monto, montoStr: _factFmtMoneda(t.monto) + ' (base, sin impuesto)',
    fechaDespachoStr: _factFmtFecha(t.fechaDespacho),
    idVenta: t.idVenta,
    montoFacturado: t.montoFacturado, montoFacturadoStr: _factFmtMoneda(t.montoFacturado),
    facturadoPor: t.facturadoPor,
    facturadoFechaStr: _factFmtFecha(t.facturadoFecha),
    montoCobrado: t.montoCobrado, monedaCobrado: t.monedaCobrado || 'ARS'
  };
}


/* ── Listas para la UI ──────────────────────────────────────── */
function FACT_getPendientesFacturar() {
  var rol = FACT_getMiRol();
  if (!rol.esFacturacion) return { ok: false, noAutorizado: true, mensaje: _factMensajeNoAutorizado(rol, 'FACTURACION_EMAILS') };
  try {
    _factSincronizarTracking();
    var track = _factMapTracking();
    var out = [];
    Object.keys(track).forEach(function(k) { var t = track[k]; if (!t.facturado) out.push(t); });
    out.sort(function(a, b) {
      var fa = (a.fechaDespacho instanceof Date) ? a.fechaDespacho.getTime() : 0;
      var fb = (b.fechaDespacho instanceof Date) ? b.fechaDespacho.getTime() : 0;
      return fa - fb;
    });
    return { ok: true, pedidos: out.map(_factFmtItem) };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}

function FACT_getPendientesCobrar() {
  var rol = FACT_getMiRol();
  if (!rol.esCobranza) return { ok: false, noAutorizado: true, mensaje: _factMensajeNoAutorizado(rol, 'COBRANZA_EMAILS') };
  try {
    _factSincronizarTracking();
    var track = _factMapTracking();
    var out = [];
    Object.keys(track).forEach(function(k) { var t = track[k]; if (t.facturado && !t.cobrado) out.push(t); });
    out.sort(function(a, b) {
      var fa = (a.facturadoFecha instanceof Date) ? a.facturadoFecha.getTime() : 0;
      var fb = (b.facturadoFecha instanceof Date) ? b.facturadoFecha.getTime() : 0;
      return fa - fb;
    });
    return { ok: true, pedidos: out.map(_factFmtItem) };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}


/* ── Acciones (marcar) ─────────────────────────────────────────────────────
   Ni el monto facturado ni el cobrado se calculan solos (los impuestos varían por reseller
   y hoy no están sistematizados en ningún lado — ver conversación con el usuario). Cada
   persona carga a mano el número/monto real; el "monto base" (sin impuesto) queda solo como
   punto de partida sugerido en la UI. */

// idVenta: ID de venta/factura que se generó en Masterchief (texto libre, para trazabilidad).
// montoFacturado: monto real facturado en USD (ya con el impuesto que corresponda aplicado).
function FACT_marcarFacturado(numero, idVenta, montoFacturado) {
  var rol = FACT_getMiRol();
  if (!rol.esFacturacion) return { ok: false, noAutorizado: true, mensaje: _factMensajeNoAutorizado(rol, 'FACTURACION_EMAILS') };
  numero = String(numero || '').trim();
  if (!numero) return { ok: false, mensaje: 'Falta el número de pedido.' };
  idVenta = String(idVenta || '').trim();
  if (!idVenta) return { ok: false, mensaje: 'Falta el ID de venta (Masterchief).' };
  montoFacturado = Number(montoFacturado);
  if (!montoFacturado || montoFacturado <= 0) return { ok: false, mensaje: 'El monto facturado tiene que ser mayor a 0.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var track = _factMapTracking();
    var t = track[numero.toUpperCase()];
    if (!t) return { ok: false, mensaje: 'No se encontró el pedido ' + numero + ' en la lista.' };
    if (t.facturado) return { ok: false, mensaje: 'Este pedido ya estaba marcado como facturado.' };
    var h = _factHoja();
    h.getRange(t.rowIdx, COL_FACT.FACTURADO + 1, 1, 3).setValues([['SÍ', new Date(), rol.email]]);
    h.getRange(t.rowIdx, COL_FACT.IDVENTA + 1, 1, 2).setValues([[idVenta, montoFacturado]]);
    return { ok: true };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

// montoCobrado / monedaCobrado: lo que REALMENTE entró (puede diferir del facturado por tipo
// de cambio, pago parcial en pesos, etc.) — se registra tal cual, sin intentar convertir nada.
function FACT_marcarCobrado(numero, montoCobrado, monedaCobrado) {
  var rol = FACT_getMiRol();
  if (!rol.esCobranza) return { ok: false, noAutorizado: true, mensaje: _factMensajeNoAutorizado(rol, 'COBRANZA_EMAILS') };
  numero = String(numero || '').trim();
  if (!numero) return { ok: false, mensaje: 'Falta el número de pedido.' };
  montoCobrado = Number(montoCobrado);
  if (!montoCobrado || montoCobrado <= 0) return { ok: false, mensaje: 'El monto cobrado tiene que ser mayor a 0.' };
  monedaCobrado = String(monedaCobrado || '').trim().toUpperCase();
  if (monedaCobrado !== 'ARS' && monedaCobrado !== 'USD') return { ok: false, mensaje: 'La moneda tiene que ser ARS o USD.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var track = _factMapTracking();
    var t = track[numero.toUpperCase()];
    if (!t) return { ok: false, mensaje: 'No se encontró el pedido ' + numero + ' en la lista.' };
    if (!t.facturado) return { ok: false, mensaje: 'Este pedido todavía no está facturado.' };
    if (t.cobrado) return { ok: false, mensaje: 'Este pedido ya estaba marcado como cobrado.' };
    var h = _factHoja();
    h.getRange(t.rowIdx, COL_FACT.COBRADO + 1, 1, 3).setValues([['SÍ', new Date(), rol.email]]);
    h.getRange(t.rowIdx, COL_FACT.MONTO_COBRADO + 1, 1, 2).setValues([[montoCobrado, monedaCobrado]]);
    return { ok: true };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
  finally { try { lock.releaseLock(); } catch (e) {} }
}
