// ============================================================
// @version 1.1
//  PORTAL RESELLER — Avisos / Comunicados post-login
//  Almacenado en MASTER hoja PORTAL_CONFIG (clave | valor)
// ============================================================

var _AVISO_SHEET = 'PORTAL_CONFIG';

function RS_getAviso() {
  try {
    var hoja = getDb().getSheetByName(_AVISO_SHEET);
    if (!hoja) return { activo: false, titulo: '', cuerpo: '' };
    var data = hoja.getDataRange().getValues();
    var map = {};
    for (var i = 0; i < data.length; i++) {
      map[String(data[i][0] || '').trim()] = String(data[i][1] || '').trim();
    }
    return {
      activo: String(map['AVISO_ACTIVO'] || '').toUpperCase() === 'TRUE',
      titulo: map['AVISO_TITULO'] || '',
      cuerpo: map['AVISO_CUERPO'] || ''
    };
  } catch(e) {
    Logger.log('RS_getAviso: ' + e);
    return { activo: false, titulo: '', cuerpo: '' };
  }
}

function RS_registrarVistaAviso(nombre) {
  try {
    if (!nombre) return;
    var ss   = getDb();
    var hoja = ss.getSheetByName('PORTAL_AVISO_VISTAS');
    if (!hoja) hoja = ss.insertSheet('PORTAL_AVISO_VISTAS');
    var data = hoja.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nombre).trim()) return;
    }
    hoja.appendRow([nombre, new Date()]);
  } catch(e) { Logger.log('RS_registrarVistaAviso: ' + e); }
}

// Sin uso desde la UI del Portal hoy (se corre a mano o queda para un panel futuro), pero
// al ser Apps Script "ANYONE_ANONYMOUS" cualquier función del proyecto es invocable desde
// la consola del navegador — sin este chequeo, cualquiera podía reescribir el banner que
// ve toda la red de resellers al loguearse (auditoría de seguridad).
function RS_setAviso(data) {
  try {
    var _email = ''; try { _email = Session.getActiveUser().getEmail(); } catch(eA) {}
    if (!_esRTVSuper(_email)) return { ok: false, error: 'No autorizado.' };
    var ss   = getDb();
    var hoja = ss.getSheetByName(_AVISO_SHEET);
    if (!hoja) hoja = ss.insertSheet(_AVISO_SHEET);
    hoja.clearContents();
    hoja.getRange(1, 1, 3, 2).setValues([
      ['AVISO_ACTIVO', data.activo ? 'TRUE' : 'FALSE'],
      ['AVISO_TITULO', String(data.titulo || '').trim()],
      ['AVISO_CUERPO', String(data.cuerpo || '').trim()]
    ]);
    return { ok: true };
  } catch(e) {
    Logger.log('RS_setAviso: ' + e);
    return { ok: false, error: e.toString() };
  }
}
