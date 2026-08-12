// @version 1.0
// ══════════════════════════════════════════════════════════════
//  MATERIAL DE CURSOS — pedido del usuario: "tengo información valiosa
//  de cursos que hemos hecho, ¿dónde podría colocarla?". Modal buscable
//  dedicado (tarjeta en Recursos → Documentación técnica), mismo patrón
//  que Videos de armado y desarmado (RS_VideosModelos.js): hoja self-
//  provisioning en el MASTER, gestión desde el LAUNCHER
//  (LAUNCH_getCursosMaterial/LAUNCH_saveCursoMaterial, Launcher_Código.js),
//  solo lectura desde acá — no se administra a mano en la hoja.
//
//  Hoja CURSOS_MATERIAL (self-provisioning, vive en el MASTER):
//    A=Título · B=Categoría · C=Descripción · D=Link · E=Fecha (yyyy-MM-dd)
//    F=Activo (vacío/SI = visible; NO = oculto)
// ══════════════════════════════════════════════════════════════

var _CM_COL = { TITULO: 0, CATEGORIA: 1, DESCRIPCION: 2, LINK: 3, FECHA: 4, ACTIVO: 5 };

function _asegurarHojaCursosMaterial() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.CURSOS_MATERIAL);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.CURSOS_MATERIAL);
    hoja.appendRow(['Título', 'Categoría', 'Descripción', 'Link', 'Fecha', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 6).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(3, 320); hoja.setColumnWidth(4, 320);
  }
  return hoja;
}

// Vacío = activo; NO / FALSE / 0 / OCULTO / INACTIVO = oculto. Mismo criterio que _vmActivo.
function _cmActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

// Cursos activos con link cargado, para el Portal — más reciente primero (sin fecha, al final).
function obtenerCursosMaterialPortal() {
  try {
    var d = _asegurarHojaCursosMaterial().getDataRange().getValues();
    var C = _CM_COL, out = [];
    for (var i = 1; i < d.length; i++) {
      var titulo = String(d[i][C.TITULO] || '').trim();
      if (!titulo) continue;
      if (!_cmActivo(d[i][C.ACTIVO])) continue;
      var link = String(d[i][C.LINK] || '').trim();
      if (!link) continue;
      var fechaRaw = d[i][C.FECHA];
      out.push({
        titulo: titulo,
        categoria: String(d[i][C.CATEGORIA] || '').trim(),
        descripcion: String(d[i][C.DESCRIPCION] || '').trim(),
        link: link,
        fecha: (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fechaRaw || '').trim()
      });
    }
    out.sort(function(a, b) { return (b.fecha || '').localeCompare(a.fecha || ''); });
    return { ok: true, cursos: out };
  } catch (e) {
    return { ok: false, error: e.toString(), cursos: [] };
  }
}
