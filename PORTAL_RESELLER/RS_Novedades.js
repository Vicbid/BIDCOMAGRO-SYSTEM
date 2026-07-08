// @version 1.0
// ══════════════════════════════════════════════════════════════
//  NOVEDADES — CMS simple para el Portal Reseller
//  El equipo interno agrega filas en la hoja NOVEDADES y el portal
//  las muestra automáticamente (sin tocar código).
//  Hoja NOVEDADES:
//    A=Fecha · B=Título · C=Descripción · D=Categoría ·
//    E=Link URL · F=Texto del botón · G=Activo (vacío/SI = visible; NO = oculto)
// ══════════════════════════════════════════════════════════════

function _asegurarHojaNovedades() {
  var ss   = getDb();
  var hoja = ss.getSheetByName(SCHEMA.SHEETS.NOVEDADES);
  if (!hoja) {
    hoja = ss.insertSheet(SCHEMA.SHEETS.NOVEDADES);
    hoja.appendRow(['Fecha', 'Título', 'Descripción', 'Categoría', 'Link URL', 'Texto del botón', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 7).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(2, 240);
    hoja.setColumnWidth(3, 380);
    hoja.setColumnWidth(5, 320);
    // Fila de ejemplo — pegá el link de la PPT en la columna E (Link URL)
    hoja.appendRow([
      new Date(),
      'Capacitación: nuevo producto T55',
      'Se viene la capacitación del nuevo DJI Agras T55. Mirá la presentación con todas las novedades del equipo.',
      'Capacitación',
      '',
      'Ver presentación',
      'SI'
    ]);
  }
  return hoja;
}

// Vacío = activo; NO / FALSE / 0 / OCULTO / INACTIVO = oculto.
function _novActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

function obtenerNovedadesPortal() {
  try {
    var hoja  = _asegurarHojaNovedades();
    var datos = hoja.getDataRange().getValues();
    var tz    = Session.getScriptTimeZone();
    var out   = [];
    for (var i = 1; i < datos.length; i++) {
      var f      = datos[i];
      var titulo = String(f[1] || '').trim();
      if (!titulo) continue;
      if (!_novActivo(f[6])) continue;

      var fechaRaw = f[0];
      var esFecha  = (fechaRaw instanceof Date);
      out.push({
        fecha:       esFecha ? Utilities.formatDate(fechaRaw, tz, 'dd/MM/yyyy') : String(fechaRaw || ''),
        ts:          esFecha ? fechaRaw.getTime() : 0,
        titulo:      titulo,
        descripcion: String(f[2] || '').trim(),
        categoria:   String(f[3] || '').trim(),
        url:         String(f[4] || '').trim(),
        urlTexto:    String(f[5] || '').trim() || 'Ver más'
      });
    }
    // Más nuevas primero (por fecha); las sin fecha quedan al final
    out.sort(function(a, b) { return b.ts - a.ts; });
    return { ok: true, novedades: out };
  } catch(e) {
    Logger.log('obtenerNovedadesPortal: ' + e);
    return { ok: false, novedades: [] };
  }
}
