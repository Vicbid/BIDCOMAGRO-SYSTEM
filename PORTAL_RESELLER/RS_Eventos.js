// @version 1.4
// ══════════════════════════════════════════════════════════════
//  EVENTOS / CURSOS — inscripción de resellers desde el Portal
//  El equipo interno carga eventos en la hoja EVENTOS y los resellers
//  se inscriben desde el portal: ¿asistís? + nombres de los asistentes
//  (email opcional). Reusable: cada evento nuevo = una fila nueva.
//
//  Hoja EVENTOS (config):
//    A=ID · B=Título · C=Fecha · D=Lugar · E=Descripción ·
//    F=Fecha límite inscripción · G=Activo (vacío/SI = visible; NO = oculto)
//  Hoja INSCRIPCIONES_EVENTOS (respuestas, 1 fila por asistente):
//    A=Fecha/hora · B=EventoID · C=Evento · D=Reseller · E=Email reseller ·
//    F=Asiste (SI/NO) · G=Nombre asistente · H=Email asistente · I=Comentario
// ══════════════════════════════════════════════════════════════

var _EV_COL  = { ID: 0, TITULO: 1, FECHA: 2, LUGAR: 3, DESC: 4, LIMITE: 5, ACTIVO: 6 };
var _INS_COL = { TS: 0, EVENTO_ID: 1, EVENTO: 2, RESELLER: 3, EMAIL_RS: 4, ASISTE: 5, NOMBRE: 6, EMAIL: 7, COMENTARIO: 8 };

function _asegurarHojaEventos() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.EVENTOS);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.EVENTOS);
    hoja.appendRow(['ID', 'Título', 'Fecha', 'Lugar', 'Descripción', 'Fecha límite inscripción', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 7).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(2, 260); hoja.setColumnWidth(5, 380);
    // Fila de ejemplo — editala con los datos reales del curso
    hoja.appendRow([
      'CURSO-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd'),
      'Curso presencial DJI Agras',
      'A confirmar',
      'BIDCOMAGRO — Carmen de Areco',
      'Curso presencial de producto y servicio. Anotate indicando cuántas personas van y sus nombres.',
      '',
      'SI'
    ]);
  }
  return hoja;
}

function _asegurarHojaInscripciones() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.INSCRIPCIONES_EVENTOS);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.INSCRIPCIONES_EVENTOS);
    hoja.appendRow(['Fecha/hora', 'EventoID', 'Evento', 'Reseller', 'Email reseller', 'Asiste', 'Nombre asistente', 'Email asistente', 'Comentario']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 9).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(3, 220); hoja.setColumnWidth(7, 200); hoja.setColumnWidth(8, 200);
  }
  return hoja;
}

// Vacío = activo; NO / FALSE / 0 / OCULTO / INACTIVO = oculto.
function _evActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

// Email del reseller desde la hoja Resellers (col J). '' si no se encuentra.
function _emailDeReseller(nombre) {
  try {
    var nn = _normText(nombre);
    if (!nn) return '';
    var d = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var R = SCHEMA.RESELLERS;
    for (var i = 1; i < d.length; i++) {
      if (_normText(d[i][R.NOMBRE]) === nn) return String(d[i][R.EMAIL] || '').trim();
    }
  } catch (e) {}
  return '';
}

// Eventos activos + la inscripción actual de este reseller (para prellenar el formulario).
// token: sin esto, cualquiera podía pedir la lista pasando el nombre de otro reseller y ver
// si está anotado a un curso y con qué asistentes (auditoría de seguridad).
function obtenerEventosPortal(token, reseller) {
  try {
    var _s = _sesionResolver(token, reseller);
    if (!_s) return { ok: false, error: 'Sesión inválida o expirada.', eventos: [] };
    reseller = _s.nombre;
    var hojaEv = _asegurarHojaEventos();
    var dEv    = hojaEv.getDataRange().getValues();
    var tz     = Session.getScriptTimeZone();

    // Inscripciones previas de este reseller, agrupadas por evento
    var misPorEvento = {};
    var nn = _normText(reseller);
    if (nn) {
      var dIns = _asegurarHojaInscripciones().getDataRange().getValues();
      var I = _INS_COL;
      for (var k = 1; k < dIns.length; k++) {
        if (_normText(dIns[k][I.RESELLER]) !== nn) continue;
        var eid = String(dIns[k][I.EVENTO_ID] || '').trim();
        if (!misPorEvento[eid]) {
          misPorEvento[eid] = { asiste: String(dIns[k][I.ASISTE] || '').toUpperCase() === 'SI', asistentes: [], comentario: '' };
        }
        var com = String(dIns[k][I.COMENTARIO] || '');
        if (com) misPorEvento[eid].comentario = com;
        var nom = String(dIns[k][I.NOMBRE] || '').trim();
        if (nom) misPorEvento[eid].asistentes.push({ nombre: nom, email: String(dIns[k][I.EMAIL] || '').trim() });
      }
    }

    var out = [];
    var C = _EV_COL;
    for (var i = 1; i < dEv.length; i++) {
      var f = dEv[i];
      var titulo = String(f[C.TITULO] || '').trim();
      if (!titulo) continue;
      if (!_evActivo(f[C.ACTIVO])) continue;
      var id     = String(f[C.ID] || '').trim() || ('EV-' + i);
      var fecha  = (f[C.FECHA]  instanceof Date) ? Utilities.formatDate(f[C.FECHA],  tz, 'dd/MM/yyyy') : String(f[C.FECHA]  || '');
      var limite = (f[C.LIMITE] instanceof Date) ? Utilities.formatDate(f[C.LIMITE], tz, 'dd/MM/yyyy') : String(f[C.LIMITE] || '');
      out.push({
        id: id, titulo: titulo, fecha: fecha, lugar: String(f[C.LUGAR] || ''),
        descripcion: String(f[C.DESC] || ''), limite: limite,
        miInscripcion: misPorEvento[id] || null
      });
    }
    return { ok: true, eventos: out };
  } catch (e) {
    Logger.log('obtenerEventosPortal: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.', eventos: [] };
  }
}

// Guarda/actualiza la inscripción de un reseller a un evento (upsert por eventoId + reseller).
// data = { asiste:bool, asistentes:[{nombre,email}], comentario }
// token: sin esto, cualquiera podía anotar (o BORRAR — hace upsert, elimina la fila previa
// antes de insertar) la inscripción de cualquier reseller a cualquier evento, sin login
// (auditoría de seguridad).
function guardarInscripcionEvento(token, eventoId, reseller, data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var _s = _sesionResolver(token, reseller);
    if (!_s) return { ok: false, error: 'Sesión inválida o expirada. Volvé a ingresar.' };
    eventoId = String(eventoId || '').trim();
    reseller = _s.nombre;
    if (!eventoId) return { ok: false, error: 'Falta el evento.' };
    if (!reseller) return { ok: false, error: 'No se identificó el reseller. Ingresá con tu PIN.' };

    data = data || {};
    var asiste     = !!data.asiste;
    var comentario = String(data.comentario || '').trim();

    var asistentes = [];
    if (asiste) {
      var arr = data.asistentes || [];
      for (var a = 0; a < arr.length; a++) {
        var nom = String((arr[a] && arr[a].nombre) || '').trim();
        if (!nom) continue;
        asistentes.push({ nombre: nom, email: String((arr[a] && arr[a].email) || '').trim() });
      }
      if (!asistentes.length) return { ok: false, error: 'Agregá al menos un asistente con nombre.' };
      if (asistentes.length > 50) return { ok: false, error: 'Máximo 50 asistentes por reseller.' };
    }

    // Título del evento (para dejar la hoja legible) + email del reseller
    var titulo = eventoId;
    try {
      var dEv = _asegurarHojaEventos().getDataRange().getValues();
      for (var e = 1; e < dEv.length; e++) {
        if (String(dEv[e][_EV_COL.ID] || '').trim() === eventoId) { titulo = String(dEv[e][_EV_COL.TITULO] || eventoId); break; }
      }
    } catch (eT) {}
    var emailRs = _emailDeReseller(reseller);

    var hoja = _asegurarHojaInscripciones();
    var d = hoja.getDataRange().getValues();
    var I = _INS_COL;
    var nn = _normText(reseller);
    // Borrar inscripción previa de este reseller para este evento (de abajo hacia arriba)
    for (var r = d.length - 1; r >= 1; r--) {
      if (String(d[r][I.EVENTO_ID] || '').trim() === eventoId && _normText(d[r][I.RESELLER]) === nn) {
        hoja.deleteRow(r + 1);
      }
    }

    var ahora  = new Date();
    var nuevas = [];
    if (asiste) {
      for (var s = 0; s < asistentes.length; s++) {
        nuevas.push([ahora, eventoId, titulo, reseller, emailRs, 'SI', _antiFormula(asistentes[s].nombre), _antiFormula(asistentes[s].email), _antiFormula(comentario)]);
      }
    } else {
      nuevas.push([ahora, eventoId, titulo, reseller, emailRs, 'NO', '', '', _antiFormula(comentario)]);
    }
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevas.length, 9).setValues(nuevas);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.INSCRIPCIONES_EVENTOS);

    return { ok: true, asiste: asiste, cantidad: asiste ? asistentes.length : 0 };
  } catch (e) {
    Logger.log('guardarInscripcionEvento: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  } finally {
    try { lock.releaseLock(); } catch (eL) {}
  }
}

// Resumen interno (correr desde el editor → Ver registro): quiénes van y cuántos.
// El comentario decía "correr desde el editor" pero, al ser una función normal en un
// proyecto ANYONE_ANONYMOUS, cualquiera podía llamarla desde el navegador y sacar el
// listado completo de asistentes (nombres) de cualquier evento sin login — se restringe a
// staff (auditoría de seguridad).
function resumenInscripcionesEvento(eventoId) {
  var _email = ''; try { _email = Session.getActiveUser().getEmail(); } catch(eA) {}
  if (!_esRTVSuper(_email)) return { evento: eventoId || '(todos)', totalResellers: 0, totalPersonas: 0, lista: [], error: 'No autorizado.' };
  eventoId = String(eventoId || '').trim();
  var d = _asegurarHojaInscripciones().getDataRange().getValues();
  var I = _INS_COL;
  var porReseller = {};
  var totalPersonas = 0;
  for (var i = 1; i < d.length; i++) {
    if (eventoId && String(d[i][I.EVENTO_ID] || '').trim() !== eventoId) continue;
    if (String(d[i][I.ASISTE] || '').toUpperCase() !== 'SI') continue;
    var rs = String(d[i][I.RESELLER] || '');
    if (!porReseller[rs]) porReseller[rs] = [];
    porReseller[rs].push(String(d[i][I.NOMBRE] || ''));
    totalPersonas++;
  }
  var lista = [];
  var keys = Object.keys(porReseller);
  for (var k = 0; k < keys.length; k++) {
    lista.push({ reseller: keys[k], personas: porReseller[keys[k]].length, nombres: porReseller[keys[k]] });
  }
  var res = { evento: eventoId || '(todos)', totalResellers: keys.length, totalPersonas: totalPersonas, lista: lista };
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}

// ── PANEL INTERNO (página web con token) ──────────────────────
// El equipo abre un link tipo  .../exec?action=resumen-curso&t=<token>  y ve, por evento,
// quiénes van y cuántas personas. Correr urlResumenCurso() UNA vez desde el editor para obtenerlo.

function _evEsc(s) { return _htmlEsc(s); }

// Token estable (mismo patrón que _tokenAprobacion): SHA-256 de 'resumen-curso|<secret>'.
function _tokenResumenCurso() {
  // Igual criterio que _tokenAprobacion (RS_Main.js): sin APPROVAL_SECRET seteado, no se
  // genera token con un secreto de respaldo fijo en el código — falla cerrado.
  var secret = PropertiesService.getScriptProperties().getProperty('APPROVAL_SECRET');
  if (!secret) { Logger.log('_tokenResumenCurso: falta APPROVAL_SECRET en Script Properties'); return null; }
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'resumen-curso|' + secret);
  return bytes.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 40);
}

// Devuelve (y loguea) el link interno del panel de inscripciones. Correr desde el editor.
function urlResumenCurso() {
  var t = _tokenResumenCurso();
  if (!t) { Logger.log('urlResumenCurso: falta APPROVAL_SECRET en Script Properties — configurala antes de generar el link.'); return null; }
  var url = ScriptApp.getService().getUrl() + '?action=resumen-curso&t=' + t;
  Logger.log(url);
  return url;
}

// Página HTML del resumen (se sirve desde doGet cuando action === 'resumen-curso').
function _paginaResumenCurso(token, soloEvento) {
  if (!token || token !== _tokenResumenCurso()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:sans-serif;max-width:520px;margin:60px auto;text-align:center;color:#c0392b">' +
      '<div style="font-size:44px">⚠️</div><h2>Link inválido</h2>' +
      '<p style="color:#666">Generá uno nuevo con <code>urlResumenCurso()</code> desde el editor.</p></div>'
    ).setTitle('Resumen curso — BIDCOMAGRO').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  soloEvento = String(soloEvento || '').trim();

  var dEv  = _asegurarHojaEventos().getDataRange().getValues();
  var dIns = _asegurarHojaInscripciones().getDataRange().getValues();
  var I = _INS_COL, C = _EV_COL, tz = Session.getScriptTimeZone();

  // Agrupar inscripciones por evento → reseller
  var porEvento = {};
  for (var k = 1; k < dIns.length; k++) {
    var eid = String(dIns[k][I.EVENTO_ID] || '').trim();
    var rs  = String(dIns[k][I.RESELLER] || '').trim();
    if (!eid || !rs) continue;
    if (!porEvento[eid]) porEvento[eid] = {};
    if (!porEvento[eid][rs]) porEvento[eid][rs] = { asiste: false, asistentes: [] };
    if (String(dIns[k][I.ASISTE] || '').toUpperCase() === 'SI') {
      porEvento[eid][rs].asiste = true;
      var nom = String(dIns[k][I.NOMBRE] || '').trim();
      if (nom) porEvento[eid][rs].asistentes.push({ nombre: nom, email: String(dIns[k][I.EMAIL] || '') });
    }
  }

  var secciones = '';
  for (var e = 1; e < dEv.length; e++) {
    var eId = String(dEv[e][C.ID] || '').trim();
    if (!eId) continue;
    if (soloEvento && eId !== soloEvento) continue;
    var titulo = String(dEv[e][C.TITULO] || eId);
    var fecha  = (dEv[e][C.FECHA] instanceof Date) ? Utilities.formatDate(dEv[e][C.FECHA], tz, 'dd/MM/yyyy') : String(dEv[e][C.FECHA] || '');
    var lugar  = String(dEv[e][C.LUGAR] || '');
    var grupo  = porEvento[eId] || {};
    var resellers = Object.keys(grupo).sort();

    var totalPersonas = 0, vanCount = 0, noCount = 0, filas = '';
    for (var ri = 0; ri < resellers.length; ri++) {
      var g = grupo[resellers[ri]];
      if (!g.asiste) {
        noCount++;
        filas += '<tr style="opacity:.6"><td style="padding:7px 10px;border-bottom:1px solid #eee">' + _evEsc(resellers[ri]) + '</td>' +
          '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:center">—</td>' +
          '<td style="padding:7px 10px;border-bottom:1px solid #eee;color:#c0392b">No asiste</td></tr>';
        continue;
      }
      vanCount++;
      totalPersonas += g.asistentes.length;
      var nombres = g.asistentes.map(function(a) {
        return _evEsc(a.nombre) + (a.email ? ' <span style="color:#999;font-size:12px">&lt;' + _evEsc(a.email) + '&gt;</span>' : '');
      }).join('<br>');
      filas += '<tr><td style="padding:7px 10px;border-bottom:1px solid #eee;font-weight:600">' + _evEsc(resellers[ri]) + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #eee;text-align:center;font-weight:700;color:#00a3e0">' + g.asistentes.length + '</td>' +
        '<td style="padding:7px 10px;border-bottom:1px solid #eee">' + nombres + '</td></tr>';
    }
    if (!filas) filas = '<tr><td colspan="3" style="padding:14px;text-align:center;color:#999">Sin inscriptos todavía.</td></tr>';

    secciones +=
      '<div style="background:#fff;border:1px solid #e5e7eb;border-radius:12px;padding:18px 20px;margin-bottom:18px;box-shadow:0 1px 4px rgba(0,0,0,.04)">' +
        '<div style="font-size:18px;font-weight:800;color:#1a2533">' + _evEsc(titulo) + '</div>' +
        '<div style="font-size:13px;color:#667085;margin:4px 0 14px">' + _evEsc(fecha) + (lugar ? ' · ' + _evEsc(lugar) : '') + '</div>' +
        '<div style="display:flex;gap:14px;margin-bottom:14px;flex-wrap:wrap">' +
          '<div style="background:#f0f9ff;border-radius:10px;padding:10px 18px"><div style="font-size:26px;font-weight:800;color:#00a3e0">' + totalPersonas + '</div><div style="font-size:11px;color:#667085;text-transform:uppercase;letter-spacing:.05em">Personas</div></div>' +
          '<div style="background:#f0fdf4;border-radius:10px;padding:10px 18px"><div style="font-size:26px;font-weight:800;color:#16a34a">' + vanCount + '</div><div style="font-size:11px;color:#667085;text-transform:uppercase;letter-spacing:.05em">Resellers que van</div></div>' +
          (noCount ? '<div style="background:#fef2f2;border-radius:10px;padding:10px 18px"><div style="font-size:26px;font-weight:800;color:#dc2626">' + noCount + '</div><div style="font-size:11px;color:#667085;text-transform:uppercase;letter-spacing:.05em">No asisten</div></div>' : '') +
        '</div>' +
        '<table style="width:100%;border-collapse:collapse;font-size:13px">' +
          '<thead><tr style="background:#f9fafb;text-align:left"><th style="padding:7px 10px">Reseller</th><th style="padding:7px 10px;text-align:center">Personas</th><th style="padding:7px 10px">Asistentes</th></tr></thead>' +
          '<tbody>' + filas + '</tbody>' +
        '</table>' +
      '</div>';
  }
  if (!secciones) secciones = '<p style="text-align:center;color:#999;padding:40px">No hay eventos cargados.</p>';

  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Resumen inscripciones — BIDCOMAGRO</title></head>' +
    '<body style="margin:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif">' +
    '<div style="max-width:820px;margin:0 auto;padding:24px 16px">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:18px;flex-wrap:wrap;gap:10px">' +
        '<h1 style="font-size:22px;color:#1a2533;margin:0">🎓 Inscripciones a cursos</h1>' +
        '<a href="javascript:location.reload()" style="font-size:13px;color:#00a3e0;text-decoration:none">↻ Actualizar</a>' +
      '</div>' + secciones +
      '<div style="text-align:center;font-size:11px;color:#aaa;margin-top:10px">BIDCOMAGRO · actualizado ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + '</div>' +
    '</div></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('Resumen inscripciones — BIDCOMAGRO').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
