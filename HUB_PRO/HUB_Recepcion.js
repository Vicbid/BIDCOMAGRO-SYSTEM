// @version 1.2
// ============================================================
//  HUB PRO — Recepción obligatoria de equipos (circuito Taller)
//  Registro permanente de con qué accesorios ingresó un equipo al
//  taller central. 5 tipos (ver _normalizarTipoRecepcion en
//  HUB_Sistema.js, que preselecciona el correcto según la hoja
//  EQUIPOS): Dron y Mavic entran CON accesorios (baterías/control
//  remoto/cooling box/cargador/hélices/otros); Batería, Control
//  remoto y Generador entran solos, sin accesorios. El gate que
//  hace obligatorio este paso vive en actualizarOrden()
//  (HUB_OTs.js) — acá solo la acción atómica que lo confirma:
//  guarda el detalle en RECEPCIONES_EQUIPO, marca la OT como
//  "Recepcionado" y manda el mail al hilo del reseller.
// ============================================================

var RECEPCION_HEADERS = ["Fecha", "OT", "Usuario", "Tipo Equipo", "Baterías", "Control Remoto", "Cooling Box", "Cargador", "Hélices", "Otros", "Observaciones"];

function _recepcionHoja() {
  var hoja = getSheet(SCHEMA.SHEETS.RECEPCIONES_EQUIPO);
  if (!hoja) {
    hoja = getDb().insertSheet(SCHEMA.SHEETS.RECEPCIONES_EQUIPO);
    hoja.appendRow(RECEPCION_HEADERS);
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// Última recepción registrada para una OT (en la práctica hay a lo sumo una — el gate impide
// una segunda — pero se lee la última por las dudas de haber más de una fila histórica).
function _obtenerRecepcionEquipo(otNum) {
  try {
    if (!otNum) return null;
    var d = getSheetValues(SCHEMA.SHEETS.RECEPCIONES_EQUIPO);
    for (var i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1] || "").trim() === otNum) {
        return {
          fecha:         (d[i][0] instanceof Date) ? d[i][0].getTime() : null,
          usuario:       String(d[i][2] || ""),
          tipoEquipo:    String(d[i][3] || ""),
          baterias:      Number(d[i][4]) || 0,
          controlRemoto: String(d[i][5] || "") === "SI",
          coolingBox:    String(d[i][6] || "") === "SI",
          cargador:      String(d[i][7] || "") === "SI",
          helices:       String(d[i][8] || "") === "SI",
          otros:         String(d[i][9] || ""),
          observaciones: String(d[i][10] || "")
        };
      }
    }
    return null;
  } catch(e) { Logger.log("_obtenerRecepcionEquipo: " + e); return null; }
}

function _armarEmailRecepcion(data, rec, usuarioNombre) {
  var items = [];
  if (rec.tipoEquipo === 'Dron' || rec.tipoEquipo === 'Mavic') {
    items.push(rec.tipoEquipo === 'Mavic' ? 'Mavic' : 'Drone');
    items.push((rec.baterias || 0) + ' batería' + (rec.baterias === 1 ? '' : 's'));
    if (rec.controlRemoto) items.push('Control remoto');
    if (rec.coolingBox)    items.push('Cooling Box');
    if (rec.cargador)      items.push('Cargador');
    if (rec.helices)       items.push('Hélices');
  } else if (rec.tipoEquipo === 'Bateria') {
    items.push('Batería');
  } else if (rec.tipoEquipo === 'Generador') {
    items.push('Generador');
  } else {
    items.push('Control remoto');
  }
  if (rec.otros) items.push(_htmlEsc(rec.otros));

  var listaHtml = "<ul style='margin:0 0 16px;padding-left:18px'>" +
    items.map(function(x) { return "<li style='font-size:13px;color:#333;padding:2px 0'>" + x + "</li>"; }).join('') +
    "</ul>";

  var ficha =
    filaDetalle("Orden de trabajo", "<strong>" + data.ot + "</strong>") +
    (data.equipo ? filaDetalle("Equipo", _htmlEsc(data.equipo)) : "") +
    filaDetalle("Fecha", Utilities.formatDate(new Date(rec.fecha), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm')) +
    filaDetalle("Registrado por", _htmlEsc(usuarioNombre));

  var obsBloque = rec.observaciones
    ? bloqueCard("Observaciones", _htmlEsc(rec.observaciones), "#888")
    : "";

  return construirEmailHTML(
    "Equipo recepcionado",
    "Estimado/a " + _htmlEsc(data.reseller),
    "<p style='font-size:13px;color:#555;margin:0 0 14px'>Su equipo ingresó a nuestro taller. Se recibió:</p>" +
    listaHtml +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:4px'>" + ficha + "</div>" +
    obsBloque,
    "Ante consultas comuníquese con su representante en BIDCOMAGRO."
  );
}

// payload: { fila, ot, tipoEquipo:'Dron'|'Bateria'|'Control'|'Generador'|'Mavic', baterias,
//            controlRemoto, coolingBox, cargador, helices, otros, observaciones }
function confirmarRecepcionEquipo(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var fila = parseInt(payload.fila);
    if (!fila || fila < 2) return { resultado: 'Error: fila inválida.' };

    // getRange NO auto-expande columnas: asegurar que exista hasta AD (RECEPCION_FECHA).
    var _maxCol = hoja.getMaxColumns();
    if (_maxCol < SCHEMA.OT.RECEPCION_FECHA + 1) hoja.insertColumnsAfter(_maxCol, (SCHEMA.OT.RECEPCION_FECHA + 1) - _maxCol);

    var old = hoja.getRange(fila, 1, 1, SCHEMA.OT.RECEPCION_FECHA + 1).getValues()[0];
    var circuito = String(old[SCHEMA.OT.CIRCUITO] || '').trim();
    if (circuito !== 'Taller') return { resultado: 'Error: la recepción solo aplica a órdenes de circuito Taller.' };
    if (old[SCHEMA.OT.RECEPCION_FECHA]) return { resultado: 'Error: esta orden ya fue recepcionada.' };

    var estadoAnterior = String(old[SCHEMA.OT.ESTADO] || '');
    var ot       = String(old[SCHEMA.OT.OT] || payload.ot || '').trim();
    var equipo   = String(old[SCHEMA.OT.EQUIPO] || '');
    var reseller = String(old[SCHEMA.OT.RESELLER] || '');
    var tecnico  = String(old[SCHEMA.OT.TECNICO] || '').trim();
    var ahora    = new Date();
    var usuario  = Session.getActiveUser().getEmail();
    var _u       = identificarUsuario();
    var usuarioNombre = (_u && _u.nombre) ? _u.nombre : usuario;

    var tipoEquipo    = (['Dron', 'Bateria', 'Control', 'Generador', 'Mavic'].indexOf(payload.tipoEquipo) !== -1) ? payload.tipoEquipo : 'Dron';
    var conAccesorios = (tipoEquipo === 'Dron' || tipoEquipo === 'Mavic'); // Batería/Control/Generador entran solos
    var baterias      = conAccesorios ? (parseInt(payload.baterias) || 0) : 0;
    var controlRemoto = conAccesorios && !!payload.controlRemoto;
    var coolingBox    = conAccesorios && !!payload.coolingBox;
    var cargador      = conAccesorios && !!payload.cargador;
    var helices       = conAccesorios && !!payload.helices;
    var otros         = String(payload.otros || '').trim();
    var observaciones = String(payload.observaciones || '').trim();

    // 1. Registro permanente
    _recepcionHoja().appendRow([
      ahora, ot, usuario, tipoEquipo, baterias,
      controlRemoto ? 'SI' : '', coolingBox ? 'SI' : '', cargador ? 'SI' : '', helices ? 'SI' : '',
      otros, observaciones
    ]);
    invalidateSheetValues(SCHEMA.SHEETS.RECEPCIONES_EQUIPO);

    // 2. Flag + estado + historial en la OT (mismo patrón que actualizarOrden)
    hoja.getRange(fila, SCHEMA.OT.ESTADO + 1).setValue('Recepcionado');
    hoja.getRange(fila, SCHEMA.OT.RECEPCION_FECHA + 1).setValue(ahora);
    hoja.getRange(fila, SCHEMA.OT.FECHA_ESTADO + 1).setValue(ahora);
    var celHist = hoja.getRange(fila, SCHEMA.OT.HISTORIAL_ESTADOS + 1);
    var histRaw = celHist.getValue();
    var hist = [];
    try { if (histRaw) hist = JSON.parse(histRaw); } catch(e2) {}
    hist.push({ f: ahora.getTime(), ant: estadoAnterior, nvo: 'Recepcionado', tec: tecnico });
    celHist.setValue(JSON.stringify(hist));
    // Bump del sello de concurrencia optimista (igual que actualizarOrden) — esta OT cambió
    // server-side, así que el cliente necesita el timestamp nuevo para no chocar con un
    // CONFLICT falso en el próximo guardado.
    hoja.getRange(fila, SCHEMA.OT.ULTIMA_MODIFICACION + 1).setValue(ahora);

    // 3. Log de auditoría
    var detalle = 'Recepción: ' + tipoEquipo +
      (conAccesorios ? ' · ' + baterias + ' batería(s)' +
        (controlRemoto ? ' · control remoto' : '') +
        (coolingBox    ? ' · cooling box'    : '') +
        (cargador      ? ' · cargador'       : '') +
        (helices       ? ' · hélices'        : '') : '') +
      (otros ? ' · otros: ' + otros : '');
    registrarLog(ot, tecnico, usuario, 'RECEPCIÓN', estadoAnterior, 'Recepcionado', detalle);

    // 4. Mail al mismo hilo del caso (reseller)
    var rec = {
      fecha: ahora.getTime(), usuario: usuario, tipoEquipo: tipoEquipo, baterias: baterias,
      controlRemoto: controlRemoto, coolingBox: coolingBox, cargador: cargador, helices: helices,
      otros: otros, observaciones: observaciones
    };
    var emailR = obtenerEmailReseller(reseller);
    if (emailR) {
      var asunto = 'Equipo recepcionado — Orden ' + ot;
      var html   = _armarEmailRecepcion({ ot: ot, equipo: equipo, reseller: reseller }, rec, usuarioNombre);
      try {
        var tid = _enviarConHilo(ot, emailR, asunto, html);
        registrarEmailLog(ot, emailR, 'Reseller', asunto, 'OK', tid || '');
      } catch(eMail) {
        registrarEmailLog(ot, emailR, 'Reseller', asunto, 'ERROR: ' + eMail.message, '');
      }
    }

    invalidateSheetValues(SCHEMA.SHEETS.OT);
    return { resultado: 'OK', ot: ot, recepcion: rec, ultimaModificacion: ahora.getTime() };

  } catch(e) {
    Logger.log('confirmarRecepcionEquipo: ' + e);
    return { resultado: 'Error: No se pudo procesar la solicitud. Intentá de nuevo.' };
  } finally {
    SpreadsheetApp.flush();
    if (lock.hasLock()) lock.releaseLock();
  }
}
