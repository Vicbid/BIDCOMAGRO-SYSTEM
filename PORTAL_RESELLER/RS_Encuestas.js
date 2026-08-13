// @version 1.0
// ══════════════════════════════════════════════════════════════
//  Encuesta de satisfacción postventa — corta y periódica, se pregunta después del
//  login (ver Index.html:_mostrarEncuestaPostLogin, encadenada a _mostrarAvisoPostLogin).
//  3 puntuaciones 1-5: Portal Reseller, Casos en garantía, Compra de repuestos.
//  Se vuelve a pedir cada ENCUESTA_PERIODO_MESES desde la última respuesta del reseller;
//  "Recordar después" no escribe nada, así que la próxima vez vuelve a estar pendiente.
//
//  Hoja ENCUESTA_POSTVENTA (self-provisioning, vive en el MASTER):
//    A=ID · B=Fecha · C=Reseller · D=PuntPortal · E=PuntGarantia · F=PuntRepuesto
// ══════════════════════════════════════════════════════════════

var ENCUESTA_PERIODO_MESES = 3;

function _asegurarHojaEncuestas() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.ENCUESTA_POSTVENTA);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.ENCUESTA_POSTVENTA);
    hoja.appendRow(['ID', 'Fecha', 'Reseller', 'PuntPortal', 'PuntGarantia', 'PuntRepuesto']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 6).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
  }
  return hoja;
}

// params: { token, reseller } — reseller SIEMPRE resuelto por sesión.
function RS_getEncuestaPendiente(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.', pendiente: false };
    var reseller = _ses.nombre;

    _asegurarHojaEncuestas();
    var d = getSheetValues(SCHEMA.SHEETS.ENCUESTA_POSTVENTA);
    var ultima = null;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][2] || '').trim() !== reseller) continue;
      var f = d[i][1] instanceof Date ? d[i][1] : new Date(d[i][1]);
      if (!ultima || f > ultima) ultima = f;
    }
    if (!ultima) return { ok: true, pendiente: true };
    var limite = new Date();
    limite.setMonth(limite.getMonth() - ENCUESTA_PERIODO_MESES);
    return { ok: true, pendiente: ultima < limite };
  } catch(e) {
    Logger.log('RS_getEncuestaPendiente: ' + e);
    return { ok: false, error: e.toString(), pendiente: false };
  }
}

// params: { token, reseller, puntPortal, puntGarantia, puntRepuesto } — 1-5 cada una.
function RS_responderEncuesta(params) {
  try {
    params = params || {};
    var _ses = _sesionResolver(params.token, params.reseller);
    if (!_ses) return { ok: false, error: 'Sesión inválida o expirada.' };
    var reseller = _ses.nombre;

    var p1 = parseInt(params.puntPortal, 10);
    var p2 = parseInt(params.puntGarantia, 10);
    var p3 = parseInt(params.puntRepuesto, 10);
    var esValida = function(n) { return n >= 1 && n <= 5; };
    if (!esValida(p1) || !esValida(p2) || !esValida(p3)) {
      return { ok: false, error: 'Faltan puntuaciones por completar.' };
    }

    var hoja = _asegurarHojaEncuestas();
    var ahora = new Date();
    var id = 'E' + ahora.getTime() + Math.floor(100 + Math.random() * 900);
    hoja.appendRow([id, ahora, reseller, p1, p2, p3]);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.ENCUESTA_POSTVENTA);

    return { ok: true };
  } catch(e) {
    Logger.log('RS_responderEncuesta: ' + e);
    return { ok: false, error: e.toString() };
  }
}
