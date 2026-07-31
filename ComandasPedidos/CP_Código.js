// @version 1.44
// ============================================================
//  COMANDAS · CARGA MASTERCHIEF — Backend · modelo N-envíos por venta
//
//  Router (doGet) + utilidades base compartidas por todo el proyecto:
//  sesión/acceso (_cpUsuarioActual/_cpUsuarioAutorizado), auditoría,
//  envío de mail genérico (rate-limit + Gmail), formatters (_num/_fmtUSD/
//  _fmtARS/_fmtCant/_norm/_fmtTs/etc.) y acceso a hojas (_cpSS/_cpHoja).
//
//  El resto se reorganizó (2026-07-30, sin cambios funcionales) en:
//    CP_Kits.js      — detección de kits + CP_getComandas
//    CP_Envios.js    — CRUD de envíos/pendientes, trigger de edición, snapshot
//    CP_Datos.js     — PDFs, config, mapeo de resellers/RTV
//    CP_Mail.js      — templates + envío/reenvío de mails, recordatorios, auto-mail
//    CP_Reportes.js  — reporte de tiempos + CP_debug*/CP_diag*
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('CP_Index')
    .setTitle('Comandas · Carga Masterchief')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/* ── HELPERS ─────────────────────────────────────────────── */
// Memoiza los spreadsheets abiertos DENTRO de una misma ejecución. Los globals se reinician
// por request en Apps Script, así que no hay datos viejos: solo evita reabrir el mismo archivo
// 3-5 veces por refresh (getRange/getValues siguen leyendo en vivo). Usa notación de corchetes
// a propósito para no ser reemplazado por el replace de openById.
var _ssCache = {};
function _cpSS(id) {
  if (!_ssCache[id]) _ssCache[id] = SpreadsheetApp['openById'](id);
  return _ssCache[id];
}


function _cpHoja() {
  var ss = _cpSS(CP_SS_ID);
  var h  = ss.getSheetByName(CP_TAB);
  if (!h) throw new Error('No se encontró la hoja "' + CP_TAB + '".');
  return h;
}


function _s(v)  { return String(v == null ? '' : v).trim(); }


// Anti-inyección de fórmulas: si un texto libre del usuario empieza con = + - @ (o tab),
// le antepone ' para que Sheets lo guarde como TEXTO y no como fórmula (evita =IMPORTDATA,
// =HYPERLINK, etc. corriendo en la hoja del dueño cuando la abre).
function _cpSafeCell(s) {
  var v = _s(s);
  return /^[=+\-@\t\r]/.test(v) ? ("'" + v) : v;
}


/* ── SEGURIDAD: usuario, autorización, auditoría, rate-limit ── */
// Email del usuario logueado (con acceso por dominio, ahora sí viene poblado).
function _cpUsuarioActual() {
  try { return _s(Session.getActiveUser().getEmail()); } catch (_) { return ''; }
}


// ¿El usuario puede ejecutar acciones que escriben o mandan mail?
// OPERADORES_AUTORIZADOS vacío = todos (opt-in). Con lista, sólo esos mails.
function _cpUsuarioAutorizado() {
  var lista = _s(_cpConfig()['OPERADORES_AUTORIZADOS']);
  var email = _cpUsuarioActual();
  if (!lista) return { ok: true, email: email };
  var permit = lista.split(',').map(function(x) { return x.trim().toLowerCase(); }).filter(Boolean);
  if (email && permit.indexOf(email.toLowerCase()) > -1) return { ok: true, email: email };
  return { ok: false, noAutorizado: true,
    mensaje: 'No estás autorizado para esta acción' + (email ? ' (' + email + ')' : '') + '. Pedí que agreguen tu mail a OPERADORES_AUTORIZADOS en _CONFIG.' };
}


// Hoja de auditoría (en el sheet de log). Se crea si no existe.
function _cpAuditHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_AUDIT_TAB);
  if (!h) {
    h = ss.insertSheet(CP_AUDIT_TAB);
    h.getRange(1, 1, 1, 6).setValues([['Fecha', 'Usuario', 'Acción', 'ID_Venta', 'Envío', 'Detalle']]);
    h.setFrozenRows(1);
    h.getRange('A1:F1').setFontWeight('bold');
    [150, 220, 150, 130, 60, 380].forEach(function(w, i) { h.setColumnWidth(i + 1, w); });
  }
  return h;
}

// Registra una acción sensible: quién + cuándo + qué + sobre qué. Nunca interrumpe la operación.
function _cpAuditar(accion, idVenta, envio, detalle) {
  try {
    _cpAuditHoja().appendRow([new Date(), _cpUsuarioActual(), _cpSafeCell(accion),
      _cpSafeCell(_s(idVenta)), (envio || envio === 0 ? envio : ''), _cpSafeCell(_s(detalle))]);
  } catch (e) { Logger.log('_cpAuditar: ' + e); }
}


// Rate-limit de mails: contador por ventana de 10 min (CacheService). Protege contra loops
// o abuso y evita agotar la cuota de Gmail. Si el cache falla, NO bloquea (fail-open).
var _mailCapMemo = null;
function _cpMailCap() {
  if (_mailCapMemo == null) { var c = _num(_cpConfig()['MAIL_MAX_POR_10MIN']); _mailCapMemo = c > 0 ? c : 60; }
  return _mailCapMemo;
}

function _cpRateBucketKey() { return 'mailrate_' + Math.floor(Date.now() / (10 * 60000)); }

function _cpRateLimitOk() {
  try { return _num(CacheService.getScriptCache().get(_cpRateBucketKey())) < _cpMailCap(); }
  catch (_) { return true; }
}

function _cpRateLimitInc() {
  try {
    var cache = CacheService.getScriptCache(), k = _cpRateBucketKey();
    cache.put(k, String(_num(cache.get(k)) + 1), 700);  // TTL > ventana de 10 min
  } catch (_) {}
}

// Único punto de envío: aplica rate-limit y manda. Lanza si se pasó el tope (los callers
// ya cachean el error como MAIL ERROR y reintentan cuando se libera la ventana).
// Convierte a ASCII puro toda la parte no-ASCII de un HTML usando referencias numéricas
// (&#NNN;): emojis y acentos viajan como ASCII y el cliente los renderiza igual, sin depender
// de cómo el transporte declare/maneje el charset. FIX de los emojis que llegaban como "������"
// (el transporte de GmailApp mangla los caracteres de plano astral —emojis de 4 bytes—).
// Seguro para nuestro HTML de mail: usa estilos inline (atributos), sin <style>/<script> donde
// las entidades no se decodifican. No toca entidades ya escapadas (&amp; &nbsp; &#233;) porque
// son ASCII. Incluye pares suplentes (flag u) para los emojis de plano astral (📦 → &#128230;).
function _cpHtmlAscii(s) {
  return String(s == null ? '' : s).replace(/[^\x00-\x7F]/gu, function(ch) {
    return '&#' + ch.codePointAt(0) + ';';
  });
}


function _cpGmailSend(to, subject, plain, opts) {
  if (!_cpRateLimitOk()) throw new Error('Límite de envío de mails alcanzado (' + _cpMailCap() + ' por 10 min, anti-abuso). Reintentá en unos minutos.');
  opts = opts || {};
  // MODO PRUEBA: si EMAIL_PRUEBA está seteado en _CONFIG, TODO el correo se redirige a esa
  // dirección (no le llega a resellers/RTV/Sole). Se cortan cc/bcc y el asunto avisa a quién
  // habría ido. Vaciar EMAIL_PRUEBA en _CONFIG para volver a producción.
  var prueba = _s(_cpConfig()['EMAIL_PRUEBA']);
  if (prueba) {
    var destinoReal = [to, opts.cc ? 'cc ' + opts.cc : '', opts.bcc ? 'bcc ' + opts.bcc : '']
      .filter(function(x){ return x; }).join(' · ');
    subject = '[PRUEBA → ' + destinoReal + '] ' + subject;
    to = prueba;
    delete opts.cc;
    delete opts.bcc;
  }
  // Emojis/acentos del cuerpo → referencias numéricas ASCII, así el cliente los renderiza bien
  // (el transporte manglaba los emojis y llegaban como "������"). Se hace acá, en el chokepoint
  // único, para cubrir TODOS los mails (reseller, RTV, Sole/aprobación, recordatorios).
  if (opts.htmlBody) opts.htmlBody = _cpHtmlAscii(opts.htmlBody);
  GmailApp.sendEmail(to, subject, plain, opts);
  _cpRateLimitInc();
}


// Normaliza un encabezado: minúsculas, sin acentos, solo letras/números.
// "ID_Venta" -> "idventa" | "Total USD" -> "totalusd" | "Razon Social" -> "razonsocial"
function _norm(v) {
  return String(v == null ? '' : v)
    .toLowerCase()
    .normalize('NFD')            // separa acentos: "ó" -> "o" + marca combinante
    .replace(/[^a-z0-9]/g, '');  // quita marcas, espacios, guiones, puntos, etc.
}


// Parseo robusto de números (acepta 1.234,56 / 1,234.56 / "USD 1.234")
function _num(v) {
  if (typeof v === 'number') return v;
  var s = _s(v);
  if (!s) return 0;
  s = s.replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
    // el último separador es el decimal
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.indexOf(',') > -1) {
    s = s.replace(',', '.');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}


function _fmtUSD(n) {
  return 'USD ' + n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function _fmtARS(n) {
  return '$ ' + n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// Cantidad: entero sin decimales; fracción con hasta 2 decimales (coma es-AR)
function _fmtCant(n) {
  return n.toLocaleString('es-AR', { maximumFractionDigits: 2 });
}


function _cpTs() {
  return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', "dd/MM/yyyy HH:mm");
}

function _fmtTs(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), 'America/Argentina/Buenos_Aires', "dd/MM/yyyy HH:mm");
}

// Parsea "dd/MM/yyyy HH:mm" (nuestro formato) → Date, o null. Tolera texto alrededor ("SÍ · ...").
function _cpParseFechaAr(s) {
  var m = String(s == null ? '' : s).match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]) : null;
}


/* ════════════════════════════════════════════════════════════
   REGISTRO DEL MOMENTO EN QUE SE MARCÓ "CARGAR"
   - onEdit (instalable) captura el momento EXACTO cuando alguien
     escribe/pega "CARGAR" en ID_Entrega → origen 'edit'.
   - CP_snapshotCargar() (opcional) estampa "detección" aprox. para
     filas que ya estaban en CARGAR antes de instalar el trigger.
   Se guarda en la hoja CP_LOG_TAB, sin tocar la hoja Ventas.
════════════════════════════════════════════════════════════ */

// clave única por línea: ID_Venta + SKU (normalizados)
function _cpKey(idv, sku) {
  return String(idv == null ? '' : idv).trim().toUpperCase() + '||' +
         String(sku == null ? '' : sku).trim().toUpperCase();
}


// Parsea un JSON {SKU:cant} → { SKU: cant>0 }
function _cpParseJson(raw) {
  try {
    var s = _s(raw); if (!s) return {};
    var o = JSON.parse(s), m = {};
    Object.keys(o).forEach(function(k) { var v = _num(o[k]); if (v > 0) m[String(k).toUpperCase()] = v; });
    return m;
  } catch (e) { return {}; }
}
