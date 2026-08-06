// ============================================================
// @version 1.5
//  PORTAL RESELLER BIDCOM — Configuración, entry point y cachés
// ============================================================

// Se bumpea a mano junto con "<!-- @version X.Y -->" de Index.html — el cliente
// la trae embebida al cargar la página y la vuelve a consultar cada tanto
// (RS_obtenerVersionActual) para avisar si quedó una pestaña vieja abierta.
// Ver _chequearVersionNueva en Index.html.
var PORTAL_VERSION = '1.77';

var PORTAL_CONFIG = {
  EMAIL_SUPERVISOR:     "soporteagrasdji@bidcom.com.ar",
  NOMBRE_REMITENTE:     "BIDCOMAGRO · Soporte Técnico DJI Agriculture",
  COL_EMAIL_RESELLER:   9,
  COL_PIN_RESELLER:     10,
  COL_FECHA_ACTIVACION: 13,
  DIAS_AVISO_GARANTIA:  30,
  DIAS_ESTIMADOS: {
    "Taller-IW": 30, "Taller-OOW": 30,
    "Reseller-IW": 30, "Reseller-OOW": 30,
    "Reseller Propio-IW": 30, "Reseller Propio-OOW": 30
  }
};


// ── CACHÉ DE ÓRDENES — evita lecturas duplicadas por ejecución ─
var _datosOT = null;
var _hojaOT  = null;
function _leerOrdenes() {
  if (!_datosOT) {
    _hojaOT  = getSheet(SCHEMA.SHEETS.OT);
    _datosOT = getSheetValues(_hojaOT);
  }
  return { hoja: _hojaOT, datos: _datosOT };
}

// ── CACHÉ DE EQUIPOS con meses de garantía (col D = índice 3) ─
var _mapaGarantias = null;
function _mapaGarantiaEquipos() {
  if (!_mapaGarantias) {
    _mapaGarantias = {};
    var d = getSheetValues(SCHEMA.SHEETS.EQUIPOS);
    for (var i = 1; i < d.length; i++) {
      var nombre = String(d[i][SCHEMA.EQUIPOS.NOMBRE] || "").trim();
      var meses  = parseInt(d[i][SCHEMA.EQUIPOS.MESES] || "12") || 12;
      if (nombre) _mapaGarantias[nombre.toLowerCase()] = meses;
    }
  }
  return _mapaGarantias;
}

function obtenerEquipos() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.EQUIPOS);
    var lista = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][SCHEMA.EQUIPOS.NOMBRE]) lista.push({
        nombre:  String(d[i][SCHEMA.EQUIPOS.NOMBRE]).trim(),
        prefijo: String(d[i][SCHEMA.EQUIPOS.PREFIJO] || "").trim().toUpperCase(),
        meses:   parseInt(d[i][SCHEMA.EQUIPOS.MESES] || "12") || 12
      });
    }
    return lista;
  } catch(e) { return []; }
}

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = String(params.action || '');

  // Guía del portal para resellers
  if (action === 'guia') {
    return HtmlService.createHtmlOutputFromFile('Guia_Reseller')
      .setTitle('Guía del Portal — BIDCOMAGRO')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // Registro de nuevos resellers (aprobación/rechazo vía link 1-click)
  if (action === 'reg-ok' || action === 'reg-no') {
    return _REG_procesarDecision(params.token, action);
  }

  // Aprobación de presupuestos OT (flujo preexistente)
  if (action === 'aprobar' || action === 'rechazar') {
    return _paginaAprobacion(params.ot, action, params.token);
  }

  // Panel interno de inscripciones a cursos/eventos (link con token; correr urlResumenCurso() en el editor)
  if (action === 'resumen-curso') {
    return _paginaResumenCurso(params.t, params.ev);
  }
  var tmpl = HtmlService.createTemplateFromFile('Index');
  tmpl.DEPLOY_URL     = ScriptApp.getService().getUrl();
  tmpl.PORTAL_VERSION = PORTAL_VERSION;
  return tmpl.evaluate()
    .setTitle('Portal Resellers — BIDCOMAGRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Consultada por el cliente (Index.html: _chequearVersionNueva) para detectar
// pestañas viejas abiertas hace rato: si difiere de la versión con la que esa
// pestaña cargó, se avisa con un banner en vez de seguir operando desactualizado.
function RS_obtenerVersionActual() {
  return PORTAL_VERSION;
}

function _tokenAprobacion(ot, action) {
  var secret = PropertiesService.getScriptProperties().getProperty('APPROVAL_SECRET') || 'bidcomagro-default';
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(ot) + '|' + String(action) + '|' + secret
  );
  return bytes.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 40);
}

function _paginaAprobacion(ot, action, token) {
  var ot = String(ot || '').trim().toUpperCase();
  var esAprobar = action === 'aprobar';
  var tokenEsperado = _tokenAprobacion(ot, action);

  var titulo, color, icono, mensaje;

  if (!ot || !token || token !== tokenEsperado) {
    titulo  = 'Link inválido';
    color   = '#e74c3c';
    icono   = '⚠️';
    mensaje = 'Este link no es válido o ya fue utilizado. Contactá a BIDCOMAGRO si necesitás ayuda.';
  } else {
    var decision = esAprobar ? 'aceptado' : 'rechazado';
    var res = aprobarPresupuestoPortal(ot, decision, 'Aprobación vía email 1-click');
    if (res.ok) {
      titulo  = esAprobar ? 'Presupuesto aprobado' : 'Presupuesto rechazado';
      color   = esAprobar ? '#1a9e4a' : '#e74c3c';
      icono   = esAprobar ? '✅' : '❌';
      mensaje = esAprobar
        ? 'Recibimos tu aprobación. Nuestro equipo iniciará la reparación a la brevedad.'
        : 'Recibimos tu decisión. La orden quedará pausada. Podés contactarnos si querés reevaluar.';
    } else {
      titulo  = 'No se pudo procesar';
      color   = '#f39c12';
      icono   = 'ℹ️';
      mensaje = res.msg || 'Es posible que el presupuesto ya haya sido procesado anteriormente.';
    }
  }

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
      (ot ? "<div class='ot'>OT " + ot + "</div>" : "") +
      "<h1>" + titulo + "</h1>" +
      "<p>" + mensaje + "</p>" +
      "<p style='font-size:12px;color:#aaa'>Podés cerrar esta ventana.</p>" +
      "<div class='logo'>BIDCOMAGRO · Soporte Técnico DJI Agriculture</div>" +
    "</div></body></html>";

  return HtmlService.createHtmlOutput(html)
    .setTitle(titulo + ' — BIDCOMAGRO')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
