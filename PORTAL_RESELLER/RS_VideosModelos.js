// @version 1.2
// ══════════════════════════════════════════════════════════════
//  VIDEOS DE ARMADO Y DESARMADO — links directos que comparte DJI.
//  Un modelo puede tener VARIOS componentes (ej: T100 = Aeronave +
//  Sistema de siembra + Sistema de elevación), cada uno con su propio
//  par armado/desarmado — de ahí la columna Componente (vacía cuando
//  el modelo es de una sola pieza, ej. una batería o un control).
//  Cada fila es una "pieza de dron" (esDron=true) o un "accesorio"
//  (esDron=false) que puede servir para varios drones a la vez
//  (dronesAsociados) — el Portal agrupa por dron al dibujar
//  (_renderVideosModelos, Index.html), mostrando cada accesorio
//  compartido debajo de cada dron al que sirve.
//  Solo lectura desde el Portal: el equipo interno carga/edita los
//  modelos y links desde el LAUNCHER (LAUNCH_getVideosModelos/
//  LAUNCH_saveVideoModelo, Launcher_Código.js), mismo patrón que
//  EVENTOS — no se administra a mano en la hoja.
//
//  Hoja VIDEOS_ARMADO_DESARMADO (self-provisioning, vive en el MASTER):
//    A=Modelo · B=Componente · C=Link armado · D=Link desarmado · E=Activo (vacío/SI = visible; NO = oculto)
//    F=Es dron (SI/NO) · G=Drones asociados (solo si F=NO, separado por ';')
// ══════════════════════════════════════════════════════════════

var _VM_COL = { MODELO: 0, COMPONENTE: 1, ARMADO: 2, DESARMADO: 3, ACTIVO: 4, ES_DRON: 5, DRONES: 6 };

function _asegurarHojaVideosModelos() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.VIDEOS_MODELOS);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.VIDEOS_MODELOS);
    hoja.appendRow(['Modelo', 'Componente', 'Link armado', 'Link desarmado', 'Activo', 'Es dron', 'Drones asociados']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 7).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(3, 320); hoja.setColumnWidth(4, 320);
  }
  return hoja;
}

// Vacío = activo; NO / FALSE / 0 / OCULTO / INACTIVO = oculto. Mismo criterio que _evActivo.
function _vmActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

// Modelos (con sus componentes) con al menos un link cargado, activos, para el Portal.
// Devuelve la lista PLANA (una fila por componente) ordenada por modelo y luego por
// componente — el cliente (_renderVideosModelos, Index.html) agrupa por modelo al
// dibujar: si un modelo tiene un solo componente se ve igual que siempre (compacto,
// sin etiqueta), si tiene varios se agrupan bajo el mismo encabezado.
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
      out.push({
        modelo: modelo,
        componente: String(d[i][C.COMPONENTE] || '').trim(),
        armado: armado, desarmado: desarmado,
        esDron: String(d[i][C.ES_DRON] || '').trim().toUpperCase() === 'SI',
        dronesAsociados: String(d[i][C.DRONES] || '').split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s; })
      });
    }
    out.sort(function(a, b) {
      var c = a.modelo.localeCompare(b.modelo);
      return c !== 0 ? c : a.componente.localeCompare(b.componente);
    });
    return { ok: true, modelos: out };
  } catch (e) {
    return { ok: false, error: e.toString(), modelos: [] };
  }
}
