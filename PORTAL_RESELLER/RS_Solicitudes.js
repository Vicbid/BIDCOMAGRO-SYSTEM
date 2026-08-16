// @version 1.2
// ══════════════════════════════════════════════════════════════
//  "Pedir algo a BidcomAgro" — solicitudes libres de resellers, sin OT/repuesto/
//  cotización de por medio (para eso ya está "Mensajes", atado a una OT puntual).
//  El reseller manda Asunto (obligatorio) + Detalle (opcional) desde "Contacto y
//  soporte"; BIDCOM recibe un mail (RS_Email.js:_notificarNuevaSolicitud) y
//  responde/marca resuelta desde el panel del LAUNCHER (Launcher_Código.js:
//  LAUNCH_getSolicitudes/LAUNCH_responderSolicitud) — el reseller ve el estado
//  y la respuesta la próxima vez que abre "Mis solicitudes" en el Portal.
//
//  Hoja SOLICITUDES_RESELLER (self-provisioning, vive en el MASTER):
//    A=ID · B=Fecha · C=Reseller · D=Asunto · E=Detalle · F=Estado
//    (Pendiente/Resuelto) · G=Respuesta · H=FechaRespuesta
// ══════════════════════════════════════════════════════════════

function _asegurarHojaSolicitudes() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.SOLICITUDES_RESELLER);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.SOLICITUDES_RESELLER);
    hoja.appendRow(['ID', 'Fecha', 'Reseller', 'Asunto', 'Detalle', 'Estado', 'Respuesta', 'FechaRespuesta']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 8).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(4, 220);
    hoja.setColumnWidth(5, 320);
    hoja.setColumnWidth(7, 320);
  }
  return hoja;
}

// params: { token, reseller, asunto, detalle } — reseller SIEMPRE resuelto por sesión.
function RS_crearSolicitud(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.' };
    var reseller = _ses.nombre;
    var asunto = String(params.asunto || '').trim();
    if (!asunto) return { ok: false, error: 'Poné un asunto para la solicitud.' };
    var detalle = String(params.detalle || '').trim();

    var hoja = _asegurarHojaSolicitudes();
    var ahora = new Date();
    var id = 'S' + ahora.getTime() + Math.floor(100 + Math.random() * 900);
    hoja.appendRow([id, ahora, reseller, _antiFormula(asunto), _antiFormula(detalle), 'Pendiente', '', '']);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.SOLICITUDES_RESELLER);

    try { _notificarNuevaSolicitud(id, reseller, asunto, detalle); } catch(e) { Logger.log('RS_crearSolicitud (notificación): ' + e); }

    return { ok: true };
  } catch(e) {
    Logger.log('RS_crearSolicitud: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}

// params: { token, reseller } — devuelve solo las solicitudes del reseller de la sesión,
// más recientes primero, tope 20.
function RS_listarMisSolicitudes(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.', items: [] };
    var reseller = _ses.nombre;

    _asegurarHojaSolicitudes();
    var d = getSheetValues(SCHEMA.SHEETS.SOLICITUDES_RESELLER);
    var out = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][2] || '').trim() !== reseller) continue;
      out.push({
        id:        String(d[i][0] || ''),
        fecha:     d[i][1] instanceof Date ? Utilities.formatDate(d[i][1], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(d[i][1] || ''),
        asunto:    String(d[i][3] || ''),
        detalle:   String(d[i][4] || ''),
        estado:    String(d[i][5] || 'Pendiente'),
        respuesta: String(d[i][6] || '')
      });
    }
    out.reverse(); // se insertan al final de la hoja → invertir para "más reciente primero"
    return { ok: true, items: out.slice(0, 20) };
  } catch(e) {
    Logger.log('RS_listarMisSolicitudes: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.', items: [] };
  }
}
