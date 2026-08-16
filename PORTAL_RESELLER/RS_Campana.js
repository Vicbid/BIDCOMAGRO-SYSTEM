// ============================================================
// @version 1.2
//  PORTAL RESELLER — Pedidos de Campaña (estimados sin compromiso)
// ============================================================
//  Script Properties requeridas (configurar desde Apps Script editor):
//    CAMPANA_ACTIVA  → ID interno de la campaña, e.g. "ago-2026"
//                      Vaciar para desactivar la campaña.
//    CAMPANA_LABEL   → Nombre visible, e.g. "Campaña Agosto 2026"
//    CAMPANA_VENCE   → Fecha límite en formato DD/MM/YYYY, e.g. "31/07/2026"
// ============================================================

function _getCampanaConfig() {
  try {
    var hoja = getDb().getSheetByName('PORTAL_CAMPANA');
    if (!hoja) return { activa: '', label: '', vence: '' };
    var data = hoja.getDataRange().getValues();
    var map  = {};
    for (var i = 0; i < data.length; i++) {
      map[String(data[i][0] || '').trim()] = String(data[i][1] || '').trim();
    }
    return {
      activa: map['CAMPANA_ACTIVA'] || '',
      label:  map['CAMPANA_LABEL']  || '',
      vence:  map['CAMPANA_VENCE']  || ''
    };
  } catch(e) { return { activa: '', label: '', vence: '' }; }
}

function RS_abrirCampana(token, reseller) {
  try {
    var _s = _sesionResolver(token, reseller);
    if (!_s) return { ok: false, msg: 'Sesión inválida o expirada. Volvé a ingresar.' };
    reseller = _s.nombre;

    var cfg     = _getCampanaConfig();
    var campana = cfg.activa;
    var label   = cfg.label;
    var vence   = cfg.vence;

    if (!campana) return { ok: true, activa: false };

    var misItems            = [];
    var ultimaActualizacion = '';
    var hoja = getSheet(SCHEMA.SHEETS.PEDIDOS_CAMPANA);

    if (hoja && hoja.getLastRow() > 0) {
      var datos      = hoja.getDataRange().getValues();
      var C          = SCHEMA.PEDIDOS_CAMPANA;
      var resellerLw = reseller.toLowerCase();

      for (var i = 0; i < datos.length; i++) {
        var rowCamp = String(datos[i][C.CAMPANA]  || '').trim();
        var rowRes  = String(datos[i][C.RESELLER] || '').trim();
        if (rowCamp !== campana || rowRes.toLowerCase() !== resellerLw) continue;

        misItems.push({
          codigo:      String(datos[i][C.CODIGO]      || '').trim(),
          descripcion: String(datos[i][C.DESCRIPCION] || '').trim(),
          cantidad:    parseInt(datos[i][C.CANTIDAD]   || '1') || 1
        });

        var ts = datos[i][C.TIMESTAMP];
        if (ts) {
          try {
            ultimaActualizacion = Utilities.formatDate(
              (ts instanceof Date ? ts : new Date(ts)),
              Session.getScriptTimeZone(),
              'dd/MM/yyyy HH:mm'
            );
          } catch(te) {}
        }
      }
    }

    return {
      ok:                  true,
      activa:              true,
      campana:             campana,
      label:               label || campana,
      vence:               vence,
      misItems:            misItems,
      ultimaActualizacion: ultimaActualizacion
    };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function RS_guardarCampana(token, reseller, campana, items) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000);
    var _s = _sesionResolver(token, reseller);
    if (!_s) return { ok: false, msg: 'Sesión inválida o expirada. Volvé a ingresar.' };
    reseller = _s.nombre;
    if (!reseller || !campana)   return { ok: false, msg: 'Faltan datos.' };
    if (!items || !items.length) return { ok: false, msg: 'La lista está vacía.' };

    var campanaActiva = _getCampanaConfig().activa;
    if (campanaActiva !== campana) {
      return { ok: false, msg: 'La campaña ya no está activa. Recargá la página.' };
    }

    var hoja = getSheet(SCHEMA.SHEETS.PEDIDOS_CAMPANA);
    if (!hoja) {
      return { ok: false, msg: 'Hoja "Pedidos_Campaña" no encontrada. Creala en el Spreadsheet.' };
    }

    var C          = SCHEMA.PEDIDOS_CAMPANA;
    var ahora      = new Date();
    var resellerLw = reseller.toLowerCase();

    // Borrar filas previas de este reseller+campaña de abajo hacia arriba.
    // Arrancar desde i=0 para no saltear la primera fila en sheets sin encabezado.
    var lastRow = hoja.getLastRow();
    if (lastRow > 0) {
      var datos = hoja.getDataRange().getValues();
      for (var i = datos.length - 1; i >= 0; i--) {
        var rowCamp = String(datos[i][C.CAMPANA]  || '').trim();
        var rowRes  = String(datos[i][C.RESELLER] || '').trim();
        if (rowCamp === campana && rowRes.toLowerCase() === resellerLw) {
          hoja.deleteRow(i + 1);
        }
      }
    }

    // Insertar filas nuevas
    var filas = [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      filas.push([
        ahora,
        campana,
        reseller,
        String(it.codigo      || '').trim(),
        String(it.descripcion || '').trim(),
        parseInt(it.cantidad  || '1') || 1
      ]);
    }

    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 6).setValues(filas);
    invalidateSheetValues(SCHEMA.SHEETS.PEDIDOS_CAMPANA);

    return { ok: true, total: filas.length };
  } catch(e) {
    return { ok: false, msg: e.message };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}
