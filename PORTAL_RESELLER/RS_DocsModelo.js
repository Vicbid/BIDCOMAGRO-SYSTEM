// @version 1.1
// ══════════════════════════════════════════════════════════════
//  DOCUMENTACIÓN POR MODELO — 2 listas simples pedidas por el usuario, cada
//  una "1 link por modelo de dron" (sin la complejidad de accesorios
//  compartidos entre varios modelos que sí tiene VIDEOS_ARMADO_DESARMADO —
//  decisión confirmada con el usuario, estos documentos son siempre
//  específicos de un solo modelo):
//    - Lista de repuestos / compra de repuestos recomendados según DJI.
//    - Mantenimiento sugerido por DJI.
//  Mismo patrón que Material de Cursos / Videos armado y desarmado: hoja
//  self-provisioning en el MASTER, gestión desde el LAUNCHER
//  (LAUNCH_getRepuestosRecomendados/LAUNCH_saveRepuestoRecomendado y los
//  equivalentes de Mantenimiento, Launcher_Código.js) — solo lectura acá.
//
//  Hoja REPUESTOS_RECOMENDADOS_DJI (self-provisioning, vive en el MASTER):
//    A=Modelo · B=Link · C=Activo (vacío/SI = visible; NO = oculto)
//  Hoja MANTENIMIENTO_DJI (self-provisioning, vive en el MASTER): mismo esquema.
// ══════════════════════════════════════════════════════════════

var _DM_COL = { MODELO: 0, LINK: 1, ACTIVO: 2 };

// Vacío = activo; NO / FALSE / 0 / OCULTO / INACTIVO = oculto. Mismo criterio que _cmActivo/_vmActivo.
function _dmActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

function _dmAsegurarHoja(nombreTab) {
  var ss   = getDb();
  var hoja = ss.getSheetByName(nombreTab);
  if (!hoja) {
    hoja = ss.insertSheet(nombreTab);
    hoja.appendRow(['Modelo', 'Link', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 3).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(2, 360);
  }
  return hoja;
}

// Filas activas con link cargado, ordenadas por modelo — misma forma para las 2 listas.
function _dmListarPortal(nombreTab) {
  try {
    var d = _dmAsegurarHoja(nombreTab).getDataRange().getValues();
    var C = _DM_COL, out = [];
    for (var i = 1; i < d.length; i++) {
      var modelo = String(d[i][C.MODELO] || '').trim();
      if (!modelo) continue;
      if (!_dmActivo(d[i][C.ACTIVO])) continue;
      var link = String(d[i][C.LINK] || '').trim();
      if (!link) continue;
      out.push({ modelo: modelo, link: link });
    }
    out.sort(function(a, b) { return a.modelo.localeCompare(b.modelo); });
    return { ok: true, items: out };
  } catch (e) {
    Logger.log('_dmListarPortal ' + nombreTab + ': ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.', items: [] };
  }
}

function obtenerRepuestosRecomendadosPortal() { return _dmListarPortal(SCHEMA.SHEETS.REPUESTOS_RECOMENDADOS); }
function obtenerMantenimientoDjiPortal()      { return _dmListarPortal(SCHEMA.SHEETS.MANTENIMIENTO_DJI); }
