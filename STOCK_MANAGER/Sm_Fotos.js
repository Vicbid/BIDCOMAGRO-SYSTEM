// ── STOCK MANAGER — Fotos faltantes del catálogo (hoja TODO) ──
// @version 1.0
// ============================================================
// Catálogo unificado (misma spreadsheet que Portal Reseller usa como
// LISTA_PRECIOS_SS_ID / RS_getListaPrecios): CATALOGO_REPUESTOS_ID (Env.js),
// pestaña 'TODO'. Columnas: A=Código Largo · B=Código Corto · C=Descripción ·
// D=Modelo · E=Cant/equipo · F=Precio PVP · G=Nivel rep · H=Imagen · I=Tipo.
// El operador ve de a N códigos sin foto (col H vacía) y sube la imagen desde
// el celular/PC; se comprime a JPEG en el front antes de mandarla acá.
// ============================================================
var _SM_FOTOS_COL = { COD_LARGO: 0, COD_CORTO: 1, DESC: 2, MODELO: 3, IMAGEN: 7 };

function _smHojaTodo() {
  return SpreadsheetApp.openById(CATALOGO_REPUESTOS_ID).getSheetByName('TODO');
}

// Devuelve hasta `n` repuestos sin foto + el total pendiente en todo el catálogo
// (para mostrar progreso). `fila` = nº de fila real (1-based) en la hoja, se manda
// de vuelta al subir para escribir directo sin tener que re-buscar por código.
function SM_obtenerRepuestosSinFoto(n) {
  try {
    n = Number(n) || 10;
    var hoja = _smHojaTodo();
    if (!hoja) return { ok: false, error: 'No se encontró la pestaña "TODO" en el catálogo.' };
    var datos = hoja.getDataRange().getValues();
    var C = _SM_FOTOS_COL;
    var items = [];
    var totalPendiente = 0;
    for (var i = 1; i < datos.length; i++) {
      var codigo = String(datos[i][C.COD_CORTO] || '').trim();
      if (!codigo) continue;
      if (String(datos[i][C.IMAGEN] || '').trim()) continue;  // ya tiene foto
      totalPendiente++;
      if (items.length < n) {
        items.push({
          fila:        i + 1,
          codigo:      codigo,
          codigoLargo: String(datos[i][C.COD_LARGO] || '').trim(),
          descripcion: String(datos[i][C.DESC]      || '').trim(),
          modelo:      String(datos[i][C.MODELO]    || '').trim()
        });
      }
    }
    return { ok: true, items: items, totalPendiente: totalPendiente };
  } catch(e) {
    Logger.log('SM_obtenerRepuestosSinFoto: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Carpeta fija de Drive para las fotos del catálogo. Self-provisioning: la busca
// por nombre, la crea si no existe, y guarda el ID en Script Properties para no
// tener que listar Drive en cada subida.
function _smCarpetaFotos() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SM_FOTOS_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  var nombre = 'Fotos Catálogo Repuestos (Stock Manager)';
  var it = DriveApp.getFoldersByName(nombre);
  var folder = it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
  props.setProperty('SM_FOTOS_FOLDER_ID', folder.getId());
  return folder;
}

// Recibe la foto en base64 (JPEG, ya comprimida por el front), la sube a Drive y
// pega el link en la col H de la fila indicada, en el mismo formato que ya usa
// el catálogo (https://drive.google.com/file/d/FILEID/view?usp=drive_link).
// Antes de escribir re-valida fila+código (por si la hoja cambió debajo) y que
// la celda H siga vacía, para no pisar una foto que otro operador subió recién.
function SM_subirFotoRepuesto(fila, codigoEsperado, base64Jpeg) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); }
  catch(eL) { return { ok: false, error: 'Otra subida en curso, reintentá en unos segundos.' }; }
  try {
    fila = Number(fila) || 0;
    if (!fila || fila < 2) return { ok: false, error: 'Fila inválida.' };
    if (!base64Jpeg) return { ok: false, error: 'No llegó ninguna imagen.' };

    var hoja = _smHojaTodo();
    if (!hoja) return { ok: false, error: 'No se encontró la pestaña "TODO" en el catálogo.' };

    var C = _SM_FOTOS_COL;
    var fRow = hoja.getRange(fila, 1, 1, C.IMAGEN + 1).getValues()[0];
    var codigoActual = String(fRow[C.COD_CORTO] || '').trim();
    if (codigoActual !== String(codigoEsperado || '').trim()) {
      return { ok: false, error: 'La lista cambió — refrescá y probá de nuevo.' };
    }
    if (String(fRow[C.IMAGEN] || '').trim()) {
      return { ok: false, error: 'Este repuesto ya tiene una foto cargada (otro operador la subió recién).' };
    }

    var bytes = Utilities.base64Decode(base64Jpeg);
    var blob  = Utilities.newBlob(bytes, 'image/jpeg', codigoActual + '.jpg');
    var file  = _smCarpetaFotos().createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=drive_link';

    hoja.getRange(fila, C.IMAGEN + 1).setValue(url);
    return { ok: true, url: url };
  } catch(e) {
    Logger.log('SM_subirFotoRepuesto: ' + e);
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch(eF) {}
  }
}
