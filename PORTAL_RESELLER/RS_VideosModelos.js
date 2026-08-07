// @version 1.0
// ══════════════════════════════════════════════════════════════
//  VIDEOS DE ARMADO Y DESARMADO — links directos que comparte DJI,
//  2 por modelo (armado / desarmado). Solo lectura desde el Portal:
//  el equipo interno carga/edita los modelos y links desde el LAUNCHER
//  (LAUNCH_getVideosModelos/LAUNCH_saveVideoModelo, Launcher_Código.js),
//  mismo patrón que EVENTOS — no se administra a mano en la hoja.
//
//  Hoja VIDEOS_ARMADO_DESARMADO (self-provisioning, vive en el MASTER):
//    A=Modelo · B=Link armado · C=Link desarmado · D=Activo (vacío/SI = visible; NO = oculto)
// ══════════════════════════════════════════════════════════════

var _VM_COL = { MODELO: 0, ARMADO: 1, DESARMADO: 2, ACTIVO: 3 };

function _asegurarHojaVideosModelos() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.VIDEOS_MODELOS);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.VIDEOS_MODELOS);
    hoja.appendRow(['Modelo', 'Link armado', 'Link desarmado', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 4).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(2, 320); hoja.setColumnWidth(3, 320);
  }
  return hoja;
}

// Vacío = activo; NO / FALSE / 0 / OCULTO / INACTIVO = oculto. Mismo criterio que _evActivo.
function _vmActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

// Modelos con al menos un link cargado, activos, para mostrar en el Portal.
function obtenerVideosModelosPortal() {
  try {
    var d = _asegurarHojaVideosModelos().getDataRange().getValues();
    var C = _VM_COL, out = [];
    for (var i = 1; i < d.length; i++) {
      var modelo = String(d[i][C.MODELO] || '').trim();
      if (!modelo) continue;
      if (!_vmActivo(d[i][C.ACTIVO])) continue;
      var armado    = String(d[i][C.ARMADO]    || '').trim();
      var desarmado = String(d[i][C.DESARMADO] || '').trim();
      if (!armado && !desarmado) continue;
      out.push({ modelo: modelo, armado: armado, desarmado: desarmado });
    }
    out.sort(function(a, b) { return a.modelo.localeCompare(b.modelo); });
    return { ok: true, modelos: out };
  } catch (e) {
    return { ok: false, error: e.toString(), modelos: [] };
  }
}
