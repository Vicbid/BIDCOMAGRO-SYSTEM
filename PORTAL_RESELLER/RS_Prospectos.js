// @version 1.3
// ============================================================
//  PORTAL RESELLER BIDCOM — Venta a prospectos (RTV) con autorización manual
// ============================================================
//
// Un RTV puede cargar un pedido a nombre de un "posible reseller" (prospecto) que
// todavía no es parte de la red. A diferencia de un pedido normal, este NO entra al
// circuito de despacho de WOS hasta que una persona designada (hoja CONFIG_PROSPECTOS,
// editada desde LAUNCHER) autorice manualmente cuántas unidades de cada ítem se
// entregan — vía un link por mail, sin login. Mientras está pendiente, no queda
// ningún rastro en Pedidos_resellers: WOS no lo ve y no compromete stock de la red.
// Esto también hace que, en la práctica, quede en "prioridad baja": cualquier pedido
// real de un reseller que se cargue mientras tanto ya le gana el turno en la cola
// FIFO de WOS cuando el de prospecto recién se cree.
//
// Precio: el RTV PROPONE un descuento global al cargar el pedido (por defecto, el % de
// CONFIG_PROSPECTOS, pero puede cambiarlo) — igual que las cantidades, esa propuesta NO es
// definitiva: el autorizador la ve junto con el stock y puede confirmarla o cambiarla antes
// de que el pedido se despache. Cantidades: las define el autorizador, no el RTV ni el
// prospecto.

// ── ¿Puede este email usar el flujo de prospectos? Cualquier RTV real (o el super) ──
function _esRTVValido(email) {
  var e = String(email || '').trim();
  if (!e) return false;
  if (_esRTVSuper(e)) return true;
  return Object.keys(_resellersAutorizadosRTV(e)).length > 0;
}

function _tokenProspecto(numero) {
  // Igual criterio que _tokenAprobacion (RS_Main.js): sin APPROVAL_SECRET seteado, no se
  // genera token con un secreto de respaldo fijo en el código — falla cerrado.
  var secret = PropertiesService.getScriptProperties().getProperty('APPROVAL_SECRET');
  if (!secret) { Logger.log('_tokenProspecto: falta APPROVAL_SECRET en Script Properties'); return null; }
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'prospecto|' + String(numero) + '|' + secret);
  return bytes.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 40);
}

function _prospEsc(s) { return _htmlEsc(s); }

// Config actual — usada por LAUNCHER (vía Launcher_Código.js) y acá mismo.
function _configProspectosInterno() {
  var out = { emailAutorizador: '', descuentoPct: 0 };
  try {
    var hoja = getSheet(SCHEMA.SHEETS.CONFIG_PROSPECTOS);
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var clave = String(d[i][0] || '').trim().toUpperCase();
        if (clave === 'EMAIL_AUTORIZADOR') out.emailAutorizador = String(d[i][1] || '').trim();
        else if (clave === 'DESCUENTO_PCT') out.descuentoPct = Number(d[i][1]) || 0;
      }
    }
  } catch(e) { Logger.log('_configProspectosInterno: ' + e); }
  return out;
}

// Solo lectura, para el cliente (Index.html): trae el % vigente para mostrar precio en la búsqueda.
function obtenerConfigProspectosPortal() {
  var cfg = _configProspectosInterno();
  return { ok: true, descuentoPct: cfg.descuentoPct };
}

// Reemplazados vigentes — misma regla que confirmarPedidoPortal (RS_Pedidos.js), pero
// duplicada acá a propósito: no se toca ese flujo en vivo por esta feature secundaria.
function _sustituirReemplazadosProspecto(items) {
  try {
    var dDbR = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var DR   = SCHEMA.DB_REPUESTOS;
    var dbMapR = {};
    for (var di0 = 1; di0 < dDbR.length; di0++) {
      var skuR = String(dDbR[di0][DR.CODIGO] || '').trim();
      if (!skuR) continue;
      dbMapR[skuR.toUpperCase()] = {
        sku: skuR,
        descripcion: String(dDbR[di0][DR.DESCRIPCION] || '').trim(),
        reemplazadoPor: String(dDbR[di0][DR.REEMPLAZADO_POR] || '').trim()
      };
    }
    for (var pi0 = 0; pi0 < items.length; pi0++) {
      var entry = dbMapR[String(items[pi0].sku || '').trim().toUpperCase()];
      if (!entry || !entry.reemplazadoPor) continue;
      var newEntry = dbMapR[entry.reemplazadoPor.toUpperCase()];
      items[pi0].sku         = newEntry ? newEntry.sku         : entry.reemplazadoPor;
      items[pi0].descripcion = newEntry ? newEntry.descripcion : items[pi0].descripcion;
    }
  } catch(eSub) { Logger.log('_sustituirReemplazadosProspecto: ' + eSub); }
  return items;
}

// ── Paso 1: el RTV carga el pedido — queda pendiente de autorización, WOS no lo ve ──
function crearPedidoProspecto(params) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    var emailRtv = '';
    try { emailRtv = Session.getActiveUser().getEmail(); } catch(eU) {}
    if (!_esRTVValido(emailRtv)) return { ok: false, error: 'No autorizado.' };

    var nombre = String((params && params.nombre) || '').trim();
    var items  = (params && params.items) || [];
    if (!nombre || !items.length) return { ok: false, error: 'Datos incompletos.' };

    var emailProspecto = String((params && params.email)    || '').trim();
    var telefono        = String((params && params.telefono) || '').trim();
    var obs              = String((params && params.observaciones) || '').trim();
    var formaPago        = String((params && params.formaPago) || '').trim();
    var envio            = String((params && params.envio)     || '').trim();
    if (telefono) obs = (obs ? obs + '\n' : '') + 'Teléfono prospecto: ' + telefono;

    var cfg = _configProspectosInterno();
    if (!cfg.emailAutorizador) {
      return { ok: false, error: 'Todavía no se configuró quién autoriza estos pedidos (Launcher → Venta a prospectos). Avisá a soporte.' };
    }

    // El RTV propone un descuento (default = el configurado), pero no es definitivo — el
    // autorizador lo confirma o lo cambia junto con las cantidades (ver _procesarAutorizacionProspecto).
    var dtoPropuesto = Number(params && params.descuentoPct);
    if (isNaN(dtoPropuesto) || dtoPropuesto < 0 || dtoPropuesto > 100) dtoPropuesto = cfg.descuentoPct;

    _asegurarHojasPedidos();

    items = _sustituirReemplazadosProspecto(items);
    items = _explotarKits(items);

    var descInfo = _descInfoResolve(null, dtoPropuesto);
    var priceMap = _buildPriceMap(descInfo.factor);
    var total = 0;
    for (var pi = 0; pi < items.length; pi++) {
      var skuKey = String(items[pi].sku || '').trim().toUpperCase();
      if (skuKey && priceMap[skuKey] !== undefined) items[pi].precio = priceMap[skuKey];
      if ((items[pi].precio || 0) > 0) total += items[pi].precio * (Number(items[pi].cantidad) || 1);
    }

    var numero  = _siguienteNumeroPedido();
    var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    if (!hojaPed) return { ok: false, error: 'No se encontró la hoja de pedidos.' };
    hojaPed.appendRow([
      numero, new Date(), _antiFormula(nombre), _antiFormula(emailProspecto),
      items.length, 0, 'Pendiente Autorización RTV', _antiFormula(obs),
      JSON.stringify(items), '', total > 0 ? total : '',
      formaPago, envio, emailRtv, dtoPropuesto
    ]);
    invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);

    try { _enviarMailAutorizacionProspecto(numero, nombre, telefono, emailProspecto, emailRtv, items, cfg.emailAutorizador, descInfo); }
    catch(eMail) { Logger.log('crearPedidoProspecto mail autorizacion: ' + eMail); }

    return { ok: true, numero: numero };
  } catch(e) {
    Logger.log('crearPedidoProspecto ERROR: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}

function _enviarMailAutorizacionProspecto(numero, nombre, telefono, emailProspecto, emailRtv, items, emailAutorizador, descInfo) {
  var stockMap = {};
  try {
    var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var S = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStock.length; s++) {
      var cod = String(dStock[s][S.CODIGO] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = Number(dStock[s][S.STOCK_ACTUAL]) || 0;
    }
  } catch(e) { Logger.log('_enviarMailAutorizacionProspecto stock: ' + e); }

  var filas = '';
  for (var i = 0; i < items.length; i++) {
    var it    = items[i];
    var skuUp = String(it.sku || '').trim().toUpperCase();
    var stk   = (skuUp && stockMap[skuUp] !== undefined) ? stockMap[skuUp] : '—';
    filas += "<tr>" +
      "<td style='padding:8px 10px;font-family:Consolas,monospace;font-size:12px;color:#00a3e0;border-bottom:1px solid #eee'>" + _prospEsc(it.sku) + "</td>" +
      "<td style='padding:8px 10px;font-size:12px;color:#333;border-bottom:1px solid #eee'>" + _prospEsc(it.descripcion) + "</td>" +
      "<td style='padding:8px 10px;text-align:center;font-weight:700;color:#333;border-bottom:1px solid #eee'>" + (it.cantidad || 1) + "</td>" +
      "<td style='padding:8px 10px;text-align:center;font-size:12px;color:#555;border-bottom:1px solid #eee'>" + stk + "</td>" +
    "</tr>";
  }
  var tabla = "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:1px solid #dde3ea;border-radius:8px;overflow:hidden'>" +
    "<thead><tr style='background:#f0f5fa'>" +
      "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888'>SKU</th>" +
      "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888'>Descripción</th>" +
      "<th style='padding:9px 10px;text-align:center;font-size:11px;font-weight:700;color:#888;width:80px'>Cant. pedida</th>" +
      "<th style='padding:9px 10px;text-align:center;font-size:11px;font-weight:700;color:#888;width:80px'>Stock actual</th>" +
    "</tr></thead><tbody>" + filas + "</tbody></table>";

  var token = _tokenProspecto(numero);
  if (!token) { Logger.log('_enviarMailAutorizacionProspecto: falta APPROVAL_SECRET, no se manda el mail (link quedaría roto)'); return; }
  var url   = ScriptApp.getService().getUrl() + '?action=autorizar-prospecto&numero=' + encodeURIComponent(numero) + '&token=' + token;

  var dPct = (descInfo && typeof descInfo.pct === 'number') ? descInfo.pct : 0;
  var precioLabel = dPct > 0 ? (dPct + '% de descuento') : 'precio de lista (PVP)';

  var cuerpo =
    "<p style='font-size:14px;color:#444;margin:0 0 14px'>El RTV <strong>" + _prospEsc(emailRtv) + "</strong> cargó un pedido para el posible reseller " +
    "<strong>" + _prospEsc(nombre) + "</strong>" + (telefono ? ' (' + _prospEsc(telefono) + ')' : '') + (emailProspecto ? ' — ' + _prospEsc(emailProspecto) : '') +
    ", pedido <strong>" + numero + "</strong>, con un descuento <strong>propuesto</strong> de " + precioLabel + ".</p>" +
    "<p style='font-size:13px;color:#666;margin:0 0 14px'>Este pedido <strong>no se despacha</strong> hasta que autorices las cantidades — y el descuento tampoco es definitivo, lo podés confirmar o cambiar ahí mismo. Revisá el stock actual antes de decidir — la red de resellers siempre tiene prioridad.</p>" +
    tabla +
    "<div style='margin-top:20px;text-align:center'>" +
      "<a href='" + url + "' target='_blank' style='display:inline-flex;align-items:center;gap:8px;padding:12px 28px;background:#7c3aed;color:#fff;border-radius:6px;text-decoration:none;font-size:13px;font-weight:700'>Revisar y autorizar cantidades</a>" +
    "</div>";

  var html = _construirEmailHTML('Autorización pendiente — Prospecto ' + numero, 'Equipo BIDCOMAGRO', cuerpo, 'Portal Resellers BIDCOMAGRO · Venta a prospectos.');
  GmailApp.sendEmail(emailAutorizador, '[Autorización requerida] Pedido de prospecto ' + numero + ' — ' + nombre, '', { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE });
}

// ── Paso 2a: GET del link — muestra el formulario para autorizar cantidades ──
function _paginaAutorizarProspecto(numero, token) {
  numero = String(numero || '').trim();
  token  = String(token  || '').trim();

  if (!numero || !token || token !== _tokenProspecto(numero)) {
    return _prospectoPaginaResultado('Link inválido', '#e74c3c', '⚠️', 'Este link no es válido. Contactá a BIDCOMAGRO si necesitás ayuda.', numero);
  }

  var P = SCHEMA.PEDIDOS_REPUESTOS;
  var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
  var datos = hojaPed ? hojaPed.getDataRange().getValues() : [];
  var row = null;
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][P.ID] || '').trim() === numero) { row = datos[i]; break; }
  }
  if (!row) return _prospectoPaginaResultado('Pedido no encontrado', '#e74c3c', '⚠️', 'No encontramos este pedido.', numero);

  var estado = String(row[P.ESTADO] || '').trim();
  if (estado !== 'Pendiente Autorización RTV') {
    var yaMsg = estado === 'Rechazado'
      ? 'Este pedido ya fue rechazado anteriormente.'
      : 'Este pedido ya fue autorizado anteriormente (estado actual: ' + estado + ').';
    return _prospectoPaginaResultado('Ya procesado', '#f39c12', 'ℹ️', yaMsg, numero);
  }

  var items = [];
  try { items = JSON.parse(String(row[P.ITEMS_JSON] || '[]')); } catch(eJ) {}

  var stockMap = {};
  try {
    var dStock = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var S = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStock.length; s++) {
      var cod = String(dStock[s][S.CODIGO] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = Number(dStock[s][S.STOCK_ACTUAL]) || 0;
    }
  } catch(eS) { Logger.log('_paginaAutorizarProspecto stock: ' + eS); }

  var filas = '';
  for (var k = 0; k < items.length; k++) {
    var it    = items[k];
    var skuUp = String(it.sku || '').trim().toUpperCase();
    var stk   = (skuUp && stockMap[skuUp] !== undefined) ? stockMap[skuUp] : '—';
    var cant  = Number(it.cantidad) || 0;
    filas +=
      "<tr>" +
        "<td style='padding:10px;font-family:Consolas,monospace;font-size:13px;color:#00a3e0;border-bottom:1px solid #eee'>" + _prospEsc(it.sku) + "</td>" +
        "<td style='padding:10px;font-size:13px;color:#333;border-bottom:1px solid #eee'>" + _prospEsc(it.descripcion) + "</td>" +
        "<td style='padding:10px;text-align:center;font-size:13px;color:#333;border-bottom:1px solid #eee'>" + cant + "</td>" +
        "<td style='padding:10px;text-align:center;font-size:13px;color:#555;border-bottom:1px solid #eee'>" + stk + "</td>" +
        "<td style='padding:10px;text-align:center;border-bottom:1px solid #eee'>" +
          "<input type='number' name='cant' value='" + cant + "' min='0' max='" + cant + "' style='width:70px;padding:6px 8px;border:1px solid #dde3ea;border-radius:6px;font-size:13px;text-align:center'>" +
        "</td>" +
      "</tr>";
  }

  var nombre       = String(row[P.RESELLER] || '').trim();
  var dtoPropuesto = Number(row[P.DTO_PROPUESTO_PCT]) || 0;
  var accionUrl    = ScriptApp.getService().getUrl();
  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
    "<meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>Autorizar pedido " + numero + " — BIDCOMAGRO</title>" +
    "<style>body{margin:0;padding:24px 12px;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;display:flex;justify-content:center}" +
    ".card{background:#fff;border-radius:12px;box-shadow:0 4px 24px rgba(0,0,0,.08);max-width:640px;width:100%;padding:32px}" +
    "h1{font-size:19px;font-weight:700;margin:0 0 6px;color:#1a1a2e}" +
    "p.sub{font-size:13px;color:#666;margin:0 0 20px}" +
    "table{width:100%;border-collapse:collapse;border:1px solid #dde3ea;border-radius:8px;overflow:hidden;margin-bottom:20px}" +
    "th{padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;background:#f0f5fa}" +
    ".btns{display:flex;gap:10px;justify-content:flex-end;flex-wrap:wrap}" +
    ".btn{border:none;border-radius:8px;padding:11px 22px;font-size:13px;font-weight:700;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center}" +
    "</style></head><body>" +
    "<div class='card'>" +
      "<h1>Pedido de prospecto — " + numero + "</h1>" +
      "<p class='sub'>Para <strong>" + _prospEsc(nombre) + "</strong>. Ajustá la cantidad que autorizás para cada ítem (podés reducirla o dejarla en 0) — lo que autorices acá pasa a despacho tal cual, la red de resellers no se ve afectada.</p>" +
      "<form method='GET' action='" + accionUrl + "'>" +
        "<input type='hidden' name='action' value='autorizar-prospecto'>" +
        "<input type='hidden' name='numero' value='" + _prospEsc(numero) + "'>" +
        "<input type='hidden' name='token' value='" + _prospEsc(token) + "'>" +
        "<table><thead><tr><th>SKU</th><th>Descripción</th><th style='text-align:center'>Pedido</th><th style='text-align:center'>Stock</th><th style='text-align:center'>Autorizás</th></tr></thead><tbody>" + filas + "</tbody></table>" +
        "<div style='margin-bottom:20px'>" +
          "<label style='font-size:11px;font-weight:700;color:#5e6778;text-transform:uppercase;letter-spacing:.06em;display:block;margin-bottom:5px'>Descuento a confirmar (%)</label>" +
          "<input type='number' name='descuento' value='" + dtoPropuesto + "' min='0' max='100' step='1' style='width:100px;padding:8px 10px;border:1px solid #dde3ea;border-radius:6px;font-size:13px;text-align:center'>" +
          "<span style='font-size:11px;color:#999;margin-left:10px'>0 = precio de lista (PVP). El RTV propuso " + dtoPropuesto + "% — lo podés dejar así o cambiarlo.</span>" +
        "</div>" +
        "<div class='btns'>" +
          "<button type='submit' class='btn' style='background:#1a9e4a;color:#fff'>Autorizar cantidades y descuento</button>" +
        "</div>" +
      "</form>" +
      "<form method='GET' action='" + accionUrl + "' style='margin-top:10px;text-align:right'>" +
        "<input type='hidden' name='action' value='autorizar-prospecto'>" +
        "<input type='hidden' name='numero' value='" + _prospEsc(numero) + "'>" +
        "<input type='hidden' name='token' value='" + _prospEsc(token) + "'>" +
        "<input type='hidden' name='rechazar' value='1'>" +
        "<button type='submit' class='btn' style='background:none;color:#e74c3c;text-decoration:underline;padding:6px'>Rechazar todo el pedido</button>" +
      "</form>" +
    "</div></body></html>";

  return HtmlService.createHtmlOutput(html)
    .setTitle('Autorizar pedido ' + numero + ' — BIDCOMAGRO')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Paso 2b: submit del formulario (o "Rechazar todo") — procesa la autorización ──
function _procesarAutorizacionProspecto(numero, token, cantidadesRaw, rechazarTodo, descuentoRaw) {
  numero = String(numero || '').trim();
  token  = String(token  || '').trim();
  if (!numero || !token || token !== _tokenProspecto(numero)) {
    return _prospectoPaginaResultado('Link inválido', '#e74c3c', '⚠️', 'Este link no es válido. Contactá a BIDCOMAGRO si necesitás ayuda.', numero);
  }

  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch(eL) { return _prospectoPaginaResultado('Ocupado', '#f39c12', 'ℹ️', 'Otra persona está procesando este pedido en este momento. Volvé a abrir el link en un rato.', numero); }

  try {
    var P = SCHEMA.PEDIDOS_REPUESTOS;
    var hojaPed = getSheet(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var datos = hojaPed ? hojaPed.getDataRange().getValues() : [];
    var fila = -1, row = null;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][P.ID] || '').trim() === numero) { fila = i; row = datos[i]; break; }
    }
    if (!row) return _prospectoPaginaResultado('Pedido no encontrado', '#e74c3c', '⚠️', 'No encontramos este pedido.', numero);

    var estado = String(row[P.ESTADO] || '').trim();
    if (estado !== 'Pendiente Autorización RTV') {
      var yaMsg = estado === 'Rechazado'
        ? 'Este pedido ya fue rechazado anteriormente — no se procesó de nuevo.'
        : 'Este pedido ya fue autorizado anteriormente (estado actual: ' + estado + ') — no se procesó de nuevo.';
      return _prospectoPaginaResultado('Ya procesado', '#f39c12', 'ℹ️', yaMsg, numero);
    }

    var items = [];
    try { items = JSON.parse(String(row[P.ITEMS_JSON] || '[]')); } catch(eJ) {}

    // El descuento tampoco es definitivo hasta acá: el autorizador puede dejar la propuesta
    // del RTV o cambiarla — se recalcula el precio de cada ítem con el % CONFIRMADO, no con
    // el que se guardó al cargar el pedido.
    var dtoConfirmado = Number(descuentoRaw);
    if (isNaN(dtoConfirmado) || dtoConfirmado < 0 || dtoConfirmado > 100) {
      dtoConfirmado = Number(row[P.DTO_PROPUESTO_PCT]) || 0;
    }
    var descInfo = _descInfoResolve(null, dtoConfirmado);
    var priceMapConfirmado = _buildPriceMap(descInfo.factor);

    var itemsAutorizados = [];
    var totalAutorizado  = 0;
    for (var k = 0; k < items.length; k++) {
      var it = items[k];
      var cantSolicitada = Number(it.cantidad) || 0;
      var cantAut = rechazarTodo ? 0 : (Number(cantidadesRaw && cantidadesRaw[k]) || 0);
      if (cantAut < 0) cantAut = 0;
      if (cantAut > cantSolicitada) cantAut = cantSolicitada; // nunca autoriza más de lo pedido
      if (cantAut <= 0) continue;
      var skuKeyAut  = String(it.sku || '').trim().toUpperCase();
      var precioFinal = (skuKeyAut && priceMapConfirmado[skuKeyAut] !== undefined) ? priceMapConfirmado[skuKeyAut] : (Number(it.precio) || 0);
      itemsAutorizados.push({ sku: it.sku, descripcion: it.descripcion, precio: precioFinal, cantidad: cantAut, modelos: it.modelos || '' });
      totalAutorizado += precioFinal * cantAut;
    }

    var nombre         = String(row[P.RESELLER]      || '').trim();
    var emailProspecto = String(row[P.EMAIL]         || '').trim();
    var obs             = String(row[P.OBSERVACIONES] || '').trim();
    var formaPago       = String(row[P.FORMA_PAGO]    || '').trim();
    var envio           = String(row[P.ENVIO]         || '').trim();
    var emailRtv        = String(row[P.RTV_EMAIL]     || '').trim();

    if (!itemsAutorizados.length) {
      hojaPed.getRange(fila + 1, P.ESTADO + 1).setValue('Rechazado');
      invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
      try { _avisarRtvResultadoProspecto(emailRtv, numero, nombre, [], true); } catch(eAv) { Logger.log('_procesarAutorizacionProspecto aviso rechazo: ' + eAv); }
      return _prospectoPaginaResultado('Pedido rechazado', '#e74c3c', '❌', 'Marcamos el pedido ' + numero + ' como rechazado. Avisamos al RTV.', numero);
    }

    // ── A partir de acá, mismo tramo final que un pedido normal: Notas de Entrega (WOS) + PDF + mail ──
    // descInfo ya quedó calculado arriba con el % CONFIRMADO (no el propuesto) — se reusa tal cual
    // para el PDF y el mail (rótulo "X% dto." / PVP).
    var resellerMeta = { nombre: nombre, direccion: '', telefono: '', emailRtv: emailRtv };

    _escribirNotasDeEntregaProspecto(numero, nombre, itemsAutorizados, formaPago, envio, obs);

    var pdfUrl = _generarPdfPedido(numero, nombre, itemsAutorizados, obs, totalAutorizado, resellerMeta, formaPago, envio, descInfo);

    hojaPed.getRange(fila + 1, 1, 1, 15).setValues([[
      numero, row[P.FECHA], nombre, emailProspecto,
      itemsAutorizados.length, 0, 'Recibido', obs,
      JSON.stringify(itemsAutorizados), pdfUrl || '', totalAutorizado > 0 ? totalAutorizado : '',
      formaPago, envio, emailRtv, dtoConfirmado
    ]]);
    invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);

    try { _enviarEmailPedidoPortal(numero, nombre, emailProspecto, itemsAutorizados, obs, totalAutorizado, pdfUrl, resellerMeta, envio, formaPago, descInfo); }
    catch(eMail) { Logger.log('_procesarAutorizacionProspecto email logistica: ' + eMail); }

    try { _avisarRtvResultadoProspecto(emailRtv, numero, nombre, itemsAutorizados, false, dtoConfirmado); }
    catch(eAv2) { Logger.log('_procesarAutorizacionProspecto aviso autorizacion: ' + eAv2); }

    return _prospectoPaginaResultado('Pedido autorizado', '#1a9e4a', '✅', 'Autorizaste ' + itemsAutorizados.length + ' ítem(s) del pedido ' + numero + '. Ya entró al circuito normal de despacho.', numero);

  } catch(e) {
    Logger.log('_procesarAutorizacionProspecto ERROR: ' + e);
    return _prospectoPaginaResultado('Error', '#e74c3c', '⚠️', 'Ocurrió un error procesando la autorización. Contactá a soporte.', numero);
  } finally {
    try { lock.releaseLock(); } catch(eF) {}
  }
}

// Notas de Entrega (Pedidos_resellers) — mismo layout de columnas que confirmarPedidoPortal
// (RS_Pedidos.js), duplicado a propósito acá: es la fila que hace que WOS vea el pedido.
function _escribirNotasDeEntregaProspecto(numero, nombre, items, formaPago, envio, obs) {
  var NOTAS_SS_ID = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
  var notasHoja = null;
  try { notasHoja = SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName('Pedidos_resellers'); }
  catch(eNS) { Logger.log('_escribirNotasDeEntregaProspecto openNotasSS: ' + eNS); }
  if (!notasHoja) {
    Logger.log('_escribirNotasDeEntregaProspecto: no se pudo abrir "Pedidos_resellers" — el pedido de prospecto ' + numero + ' no va a quedar visible para WOS.');
    try {
      GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR,
        '[Portal Reseller] Pedido de prospecto sin registrar en Pedidos_resellers',
        'No se pudo abrir la hoja "Pedidos_resellers" al autorizar el pedido de prospecto ' + numero + ' (' + nombre + '). ' +
        'Quedó autorizado en PEDIDOS_REPUESTOS pero WOS no lo va a ver hasta reconstruir la fila a mano.');
    } catch(eAlert) {}
    return;
  }

  var stockMapCarmen = {};
  try {
    var dStk = getStockSheetValues(SCHEMA.SHEETS.STOCK_INVENTARIO);
    var SK   = SCHEMA.STOCK_INVENTARIO;
    for (var s = 1; s < dStk.length; s++) {
      var skcod = String(dStk[s][SK.CODIGO] || '').trim().toUpperCase();
      if (skcod) stockMapCarmen[skcod] = Number(dStk[s][SK.STOCK_ACTUAL]) || 0;
    }
  } catch(eStk) { Logger.log('_escribirNotasDeEntregaProspecto stockMapCarmen: ' + eStk); }

  for (var i = 0; i < items.length; i++) {
    var it    = items[i];
    var skuUp = String(it.sku || '').trim().toUpperCase();
    var cant  = Number(it.cantidad) || 1;
    var prec  = Number(it.precio)   || 0;
    var stkH  = stockMapCarmen[skuUp] !== undefined ? stockMapCarmen[skuUp] : '';
    var newRow    = notasHoja.getLastRow() + 1;
    var stkActual = stockMapCarmen[skuUp] !== undefined ? stockMapCarmen[skuUp] : -1;
    var estadoNota = (stkActual >= 0 && stkActual >= cant) ? 'Confirmado' : 'Pendiente_Revision';
    notasHoja.appendRow([
      numero, _antiFormula(nombre), _antiFormula(it.sku) || '', _antiFormula(it.descripcion) || '', cant, 0,
      '=E' + newRow + '-F' + newRow + '-Z' + newRow,  // fórmula intencional, no tocar
      prec, stkH, estadoNota, new Date(), envio, formaPago, _antiFormula(obs)
    ]);
  }
}

function _avisarRtvResultadoProspecto(emailRtv, numero, nombre, items, rechazado, dtoConfirmado) {
  if (!emailRtv) return;
  var cuerpo;
  if (rechazado) {
    cuerpo = "<p style='font-size:14px;color:#444'>El pedido <strong>" + numero + "</strong> para <strong>" + _prospEsc(nombre) + "</strong> fue <strong style='color:#e74c3c'>rechazado</strong> por quien autoriza. No se despachó nada.</p>";
  } else {
    var filas = '';
    for (var i = 0; i < items.length; i++) {
      filas += "<li>" + _prospEsc(items[i].sku) + " — " + _prospEsc(items[i].descripcion) + " × " + items[i].cantidad + "</li>";
    }
    var dtoLabel = (Number(dtoConfirmado) > 0) ? (Number(dtoConfirmado) + '% de descuento') : 'precio de lista (PVP)';
    cuerpo = "<p style='font-size:14px;color:#444'>El pedido <strong>" + numero + "</strong> para <strong>" + _prospEsc(nombre) + "</strong> fue autorizado a <strong>" + dtoLabel + "</strong> y ya entró al circuito normal de despacho:</p>" +
      "<ul style='font-size:13px;color:#333'>" + filas + "</ul>";
  }
  var html = _construirEmailHTML((rechazado ? 'Pedido rechazado' : 'Pedido autorizado') + ' — ' + numero, 'RTV', cuerpo, 'Portal Resellers BIDCOMAGRO · Venta a prospectos.');
  try { GmailApp.sendEmail(emailRtv, (rechazado ? '[Rechazado] ' : '[Autorizado] ') + 'Pedido de prospecto ' + numero, '', { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE }); }
  catch(e) { Logger.log('_avisarRtvResultadoProspecto: ' + e); }
}

// Tarjeta de resultado inline — mismo estilo visual que _paginaAprobacion (RS_Main.js).
function _prospectoPaginaResultado(titulo, color, icono, mensaje, numero) {
  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
    "<meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>" + titulo + " — BIDCOMAGRO</title>" +
    "<style>body{margin:0;padding:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh}" +
    ".card{background:#fff;border-radius:12px;box-shadow:0 4px 24px rgba(0,0,0,.08);max-width:480px;width:90%;padding:40px 36px;text-align:center}" +
    ".ico{font-size:52px;margin-bottom:16px}" +
    ".ot{display:inline-block;background:#f0f7ff;color:#00a3e0;font-weight:700;font-size:13px;padding:4px 12px;border-radius:20px;margin-bottom:20px}" +
    "h1{font-size:22px;font-weight:700;margin:0 0 12px;color:" + color + "}" +
    "p{font-size:14px;color:#666;line-height:1.6;margin:0 0 24px}" +
    ".logo{font-size:11px;color:#bbb;margin-top:32px;border-top:1px solid #f0f0f0;padding-top:16px}" +
    "</style></head><body>" +
    "<div class='card'>" +
      "<div class='ico'>" + icono + "</div>" +
      (numero ? "<div class='ot'>Pedido " + _prospEsc(numero) + "</div>" : "") +
      "<h1>" + titulo + "</h1>" +
      "<p>" + mensaje + "</p>" +
      "<p style='font-size:12px;color:#aaa'>Podés cerrar esta ventana.</p>" +
      "<div class='logo'>BIDCOMAGRO · Soporte Técnico DJI Agriculture</div>" +
    "</div></body></html>";
  return HtmlService.createHtmlOutput(html)
    .setTitle(titulo + ' — BIDCOMAGRO')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
