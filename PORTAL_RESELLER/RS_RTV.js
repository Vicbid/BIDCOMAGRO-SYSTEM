// @version 1.5
// ============================================================
//  PORTAL RESELLER BIDCOM — Vista RTV (solo lectura)
// ============================================================

// Super-RTV: emails (dominio bidcom) que ven TODOS los resellers, no solo los de su EMAIL_RTV.
// Solo lectura, igual que cualquier RTV. Autorización server-side por identidad de Google.
var _RTV_SUPER = ['soporteagrasdji@bidcom.com.ar'];
function _esRTVSuper(email) {
  return _RTV_SUPER.indexOf(String(email || '').toLowerCase().trim()) !== -1;
}

// Mapa {nombreLower → nombre} de resellers que este email puede ver como RTV.
// Super-RTV → todos los resellers; RTV normal → los que lo tienen en la columna EMAIL_RTV.
function _resellersAutorizadosRTV(email) {
  var emailLow = String(email || '').toLowerCase().trim();
  var sup      = _esRTVSuper(emailLow);
  var datos    = getSheetValues(SCHEMA.SHEETS.RESELLERS);
  var RS       = SCHEMA.RESELLERS;
  var out      = {};
  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][RS.NOMBRE] || '').trim();
    if (!nombre) continue;
    if (sup) { out[nombre.toLowerCase()] = nombre; continue; }
    var rtvEmail = String(datos[i][RS.EMAIL_RTV] || '').trim().toLowerCase();
    if (rtvEmail && rtvEmail === emailLow) out[nombre.toLowerCase()] = nombre;
  }
  return out;
}

// Detecta si el usuario logueado es un RTV (o super-RTV) verificando su email de sesión.
// Retorna { ok, email, resellers: [nombres], esSuper }
function obtenerDatosRTV() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch(e) {}
    if (!email) return { ok: false };
    var emailLow = email.toLowerCase().trim();

    var mapa = _resellersAutorizadosRTV(emailLow);
    var resellers = [];
    for (var k in mapa) resellers.push(mapa[k]);
    if (!resellers.length) return { ok: false };
    resellers.sort();

    // Mapa reseller → email de su RTV (solo los autorizados). El super lo usa para
    // filtrar por RTV en el front; para un RTV normal apunta todo a su propio email.
    var datosR = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var RS2    = SCHEMA.RESELLERS;
    var resellerRtv = {};
    for (var r = 1; r < datosR.length; r++) {
      var nom = String(datosR[r][RS2.NOMBRE] || '').trim();
      if (!nom || !mapa[nom.toLowerCase()]) continue;
      resellerRtv[nom] = String(datosR[r][RS2.EMAIL_RTV] || '').trim().toLowerCase();
    }

    return { ok: true, email: email, resellers: resellers, esSuper: _esRTVSuper(emailLow), resellerRtv: resellerRtv };
  } catch(e) {
    Logger.log('obtenerDatosRTV: ' + e);
    return { ok: false };
  }
}

// Lee Pedidos_resellers (Notas de Entrega — fuente de verdad de lo despachado) y arma
// un mapa numero → cumplimiento en CANTIDADES (unidades), con detalle por SKU.
// Columnas de Pedidos_resellers: 0=NUMERO, 2=SKU, 3=DESC, 4=SOLICITADA(E), 5=DESPACHADA(F),
// 7=PRECIO(H), 9=ESTADO(J), 25=CANCELADA(Z). WOS actualiza col F al despachar y col Z al cancelar.
// Cancelación (mismo criterio que WOS): fuente primaria = CANT_CANCEL (col Z); fallback datos
// viejos = línea con estado "Cancelado" sin CANT_CANCEL → se cancela toda la cantidad solicitada.
function _mapaCumplimientoPedidos() {
  var mapa = {};
  try {
    var NOTAS_SS_ID_PR = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
    var wosHoja = SpreadsheetApp.openById(NOTAS_SS_ID_PR).getSheetByName('Pedidos_resellers');
    if (!wosHoja) return mapa;
    var d = wosHoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var num = String(d[i][0] || '').trim();
      if (!num) continue;
      var sku  = String(d[i][2] || '').trim();
      var desc = String(d[i][3] || '').trim();
      var sol  = Number(d[i][4])  || 0;
      var des  = Number(d[i][5])  || 0;
      var pre  = Number(d[i][7])  || 0;
      var est  = String(d[i][9] || '').trim().toLowerCase();
      var can  = Number(d[i][25]) || 0;
      if (can <= 0 && est === 'cancelado') can = sol;  // dato viejo: estado cancelado sin CANT_CANCEL
      if (!mapa[num]) mapa[num] = { solicitado: 0, entregado: 0, cancelado: 0, items: {}, orden: [] };
      var reg = mapa[num];
      reg.solicitado += sol;
      reg.entregado  += des;
      reg.cancelado  += can;
      var key = sku.toUpperCase() || ('__' + desc);
      if (!reg.items[key]) { reg.items[key] = { sku: sku, descripcion: desc, precio: pre, solicitado: 0, entregado: 0, cancelado: 0 }; reg.orden.push(key); }
      var it = reg.items[key];
      it.solicitado += sol;
      it.entregado  += des;
      it.cancelado  += can;
      if (!it.precio && pre) it.precio = pre;
    }
  } catch(e) { Logger.log('_mapaCumplimientoPedidos: ' + e); }
  return mapa;
}

// Convierte un registro del mapa en el resumen listo para el front:
// entregado / pendiente / cancelado / porcentaje de cumplimiento + detalle por ítem.
// % = entregado / (solicitado - cancelado); lo cancelado no cuenta como "a entregar".
function _resumenCumplimiento(reg) {
  var solicitado = reg ? reg.solicitado : 0;
  var entregado  = reg ? reg.entregado  : 0;
  var cancelado  = reg ? reg.cancelado  : 0;
  var base       = solicitado - cancelado;
  var pendiente  = base - entregado; if (pendiente < 0) pendiente = 0;
  // Pedido totalmente cancelado (todos los ítems anulados) → no es "0% pendiente", es "Cancelado".
  var anulado    = (solicitado > 0 && base <= 0 && cancelado > 0);
  var porcentaje;
  if (base <= 0) porcentaje = (entregado > 0 ? 100 : 0);
  else           porcentaje = Math.round(entregado / base * 100);
  if (porcentaje > 100) porcentaje = 100;
  if (porcentaje < 0)   porcentaje = 0;
  var items = [];
  if (reg && reg.orden) {
    for (var i = 0; i < reg.orden.length; i++) {
      var it    = reg.items[reg.orden[i]];
      var iBase = it.solicitado - it.cancelado;
      var iPend = iBase - it.entregado; if (iPend < 0) iPend = 0;
      items.push({
        sku: it.sku, descripcion: it.descripcion, precio: it.precio,
        solicitado: it.solicitado, entregado: it.entregado, cancelado: it.cancelado, pendiente: iPend
      });
    }
  }
  return {
    solicitado: solicitado, entregado: entregado, cancelado: cancelado,
    pendiente: pendiente, porcentaje: porcentaje, anulado: anulado,
    tieneDatos: !!(reg && solicitado > 0), items: items
  };
}

// Retorna los pedidos de los resellers asignados al RTV llamante.
// La verificación se hace server-side: no se acepta la lista del cliente,
// se re-calcula desde el email de la sesión activa.
function obtenerPedidosRTV() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e) {}
    if (!email) return { ok: false, error: 'Sin sesión', pedidos: [] };

    // Determinar resellers autorizados para este RTV (super-RTV ve todos)
    var autorizado = _resellersAutorizadosRTV(email);
    if (!Object.keys(autorizado).length) return { ok: false, error: 'Sin resellers asignados', pedidos: [] };

    var cumpl = _mapaCumplimientoPedidos();  // numero → entregado/pendiente/% (fuente: Pedidos_resellers)

    var dPed  = getSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var P     = SCHEMA.PEDIDOS_REPUESTOS;
    var pedidos = [];

    for (var j = dPed.length - 1; j >= 1; j--) {
      var reseller = String(dPed[j][P.RESELLER] || '').trim();
      if (!autorizado[reseller.toLowerCase()]) continue;

      var fecha = dPed[j][P.FECHA];
      var fechaStr = (fecha instanceof Date)
        ? Utilities.formatDate(fecha, 'GMT-3', 'dd/MM/yyyy HH:mm')
        : String(fecha || '');

      var id = String(dPed[j][P.ID] || '').trim();

      var itemsJson = String(dPed[j][P.ITEMS_JSON] || '[]');
      var items = [];
      try { items = JSON.parse(itemsJson); } catch(e) {}

      pedidos.push({
        id:        id,
        fecha:     fechaStr,
        reseller:  reseller,
        cantItems: Number(dPed[j][P.CANT_ITEMS]   || 0),
        estado:    String(dPed[j][P.ESTADO]        || ''),
        total:     Number(dPed[j][P.TOTAL_USD]     || 0),
        obs:       String(dPed[j][P.OBSERVACIONES] || ''),
        pdfUrl:    String(dPed[j][P.PDF_URL]       || ''),
        formaPago: String(dPed[j][P.FORMA_PAGO]    || ''),
        envio:     String(dPed[j][P.ENVIO]         || ''),
        items:     items,
        cumplimiento: _resumenCumplimiento(cumpl[id])
      });

      if (pedidos.length >= 200) break;
    }

    return { ok: true, pedidos: pedidos };
  } catch(e) {
    Logger.log('obtenerPedidosRTV: ' + e);
    return { ok: false, error: e.toString(), pedidos: [] };
  }
}

// Retorna las órdenes de trabajo (OTs) de los resellers asignados al RTV llamante.
// Solo lectura. La autorización se re-calcula server-side desde el email de la
// sesión (mismo criterio que obtenerPedidosRTV): no se confía en el cliente.
function obtenerOrdenesRTV() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e) {}
    if (!email) return { ok: false, error: 'Sin sesión', ordenes: [] };

    // Resellers autorizados para este RTV (super-RTV ve todos)
    var autorizado = _resellersAutorizadosRTV(email);
    if (!Object.keys(autorizado).length) return { ok: false, error: 'Sin resellers asignados', ordenes: [] };

    var ref = _leerOrdenes();
    var tz  = Session.getScriptTimeZone();
    var hoy = new Date();
    var out = [];
    for (var j = 1; j < ref.datos.length; j++) {
      var f = ref.datos[j];
      if (!f[SCHEMA.OT.OT]) continue;
      var reseller = String(f[SCHEMA.OT.RESELLER] || '').trim();
      if (!autorizado[reseller.toLowerCase()]) continue;

      var estado    = String(f[SCHEMA.OT.ESTADO] || '');
      var esCerrada = (estado === 'Finalizado' || estado === 'Entregado' || estado === 'CANCELADO');
      var fechaFin  = (f[SCHEMA.OT.FECHA_CIERRE] instanceof Date) ? f[SCHEMA.OT.FECHA_CIERRE] : hoy;
      var dias      = (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? Math.floor((fechaFin - f[SCHEMA.OT.FECHA_INGRESO]) / 86400000) : 0;

      out.push({
        ot:           String(f[SCHEMA.OT.OT]),
        reseller:     reseller,
        equipo:       String(f[SCHEMA.OT.EQUIPO] || ''),
        sn:           String(f[SCHEMA.OT.SN] || ''),
        estado:       estado,
        garantia:     String(f[SCHEMA.OT.GARANTIA] || ''),
        tecnico:      String(f[SCHEMA.OT.TECNICO] || ''),
        cliente:      String(f[SCHEMA.OT.CLIENTE] || ''),
        prioridad:    String(f[SCHEMA.OT.PRIORIDAD] || '').toUpperCase() === 'URGENTE',
        dias:         dias,
        cerrada:      esCerrada,
        fechaIngreso: (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? Utilities.formatDate(f[SCHEMA.OT.FECHA_INGRESO], tz, 'dd/MM/yyyy') : '',
        fechaCierre:  (esCerrada && f[SCHEMA.OT.FECHA_CIERRE] instanceof Date) ? Utilities.formatDate(f[SCHEMA.OT.FECHA_CIERRE], tz, 'dd/MM/yyyy') : ''
      });
      if (out.length >= 400) break;
    }
    // Activas primero; dentro de cada grupo, las de más días arriba
    out.sort(function(a, b) {
      return (a.cerrada === b.cerrada) ? (b.dias - a.dias) : (a.cerrada ? 1 : -1);
    });
    return { ok: true, ordenes: out };
  } catch(e) {
    Logger.log('obtenerOrdenesRTV: ' + e);
    return { ok: false, error: e.toString(), ordenes: [] };
  }
}
