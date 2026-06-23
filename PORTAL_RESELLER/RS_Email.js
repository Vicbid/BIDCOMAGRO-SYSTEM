// ============================================================
// @version 1.0
//  PORTAL RESELLER BIDCOM — Helpers de email y notificaciones
// ============================================================

// Espera indexación de Gmail y devuelve el Thread ID del hilo recién creado.
// Usa el número de OT del asunto para acotar la búsqueda y reintenta hasta 3 veces.
function _capturarThreadId(para, asunto) {
  try {
    var otMatch = asunto ? asunto.match(/OT\s+([\w\-\/]+)/) : null;
    var query   = 'in:sent to:(' + para + ') newer_than:1d';
    if (otMatch) query += ' subject:("OT ' + otMatch[1] + '")';
    var delays  = [3000, 5000, 7000];
    for (var i = 0; i < delays.length; i++) {
      Utilities.sleep(delays[i]);
      var hilos = GmailApp.search(query, 0, 1);
      if (hilos.length) return hilos[0].getId();
    }
  } catch(e) { Logger.log('_capturarThreadId: ' + e); }
  return null;
}

// Devuelve el Thread ID ancla almacenado en EMAIL_LOGS para esta OT + destinatario.
function _obtenerThreadId(ot, destinatario) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.EMAIL_LOGS);
    var otS   = String(ot);
    var desS  = String(destinatario);
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][1] || '') === otS &&
          String(datos[i][2] || '') === desS &&
          datos[i][6]) {
        return String(datos[i][6]);
      }
    }
  } catch(e) { Logger.log('_obtenerThreadId: ' + e); }
  return null;
}

// Registra un email en EMAIL_LOGS (col G = MessageID), compatible con HUB
function _logEmail(ot, destinatario, rol, asunto, estado, messageId) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
    if (!hoja) {
      hoja = getDb().insertSheet('EMAIL_LOGS');
      hoja.appendRow(['Fecha','OT','Destinatario','Rol','Asunto','Estado','ThreadID']);
    }
    hoja.appendRow([new Date(), ot, destinatario, rol, asunto, estado, messageId || '']);
    invalidateSheetValues(SCHEMA.SHEETS.EMAIL_LOGS);
  } catch(e) { Logger.log('_logEmail: ' + e); }
}

function _emailReseller(nombre) {
  try {
    var datos   = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var nombreB = String(nombre || '').trim().toLowerCase();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === nombreB)
        return String(datos[i][SCHEMA.RESELLERS.EMAIL] || '').trim();
    }
  } catch(e) { Logger.log('_emailReseller: ' + e); }
  return '';
}

function _notificarNuevaOT(nOT, data, cotizUrl) {
  try {
    var clienteStr  = (data.cliente && String(data.cliente).trim()) ? String(data.cliente).trim() : '';
    var asunto      = "OT " + nOT + " — " + data.equipo + (clienteStr ? " · " + clienteStr : "");
    var repuestosRow = '';
    if (data.repuestos) {
      try {
        var reps = JSON.parse(data.repuestos);
        if (reps && reps.length) {
          var lista = reps.map(function(r){ return r.sku + ' | ' + (r.descripcion || '') + ' | P:' + r.cantidad + ' E:0'; }).join('<br>');
          repuestosRow = _filaDetalle("Repuestos solicitados", lista);
        }
      } catch(e) {}
    }
    var cotizBtn = cotizUrl
      ? "<div style='margin-top:18px;text-align:center'><a href='" + cotizUrl + "' target='_blank' style='display:inline-block;background:#1a9e4a;color:#fff;padding:10px 24px;border-radius:6px;text-decoration:none;font-size:14px;font-weight:600'>Ver Pedido de Repuestos</a></div>"
      : '';

    var html = _construirEmailHTML(
      "Nueva solicitud de servicio", "Supervisor",
      "<p style='font-size:14px;color:#444;margin:0 0 20px'>El reseller <strong>" + data.reseller + "</strong> registró una nueva solicitud desde el Portal.</p>" +
      "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px'>" +
        _filaDetalle("OT generada", "<strong>" + nOT + "</strong>") +
        _filaDetalle("Reseller", data.reseller) +
        _filaDetalle("Equipo", data.equipo) +
        _filaDetalle("N° de Serie", data.sn || "—") +
        (clienteStr ? _filaDetalle("Cliente", clienteStr) : '') +
        _filaDetalle("Garantía", data.garantia) +
        _filaDetalle("Tipo de gestión", data.circuito) +
        (data.cas ? _filaDetalle("N° CAS / FWRC", data.cas) : '') +
        (data.aftEstado ? _filaDetalle("Estado reparación", data.aftEstado === 'repuesto' ? 'Ya reparado — Repuesto para reposición' : 'Pendiente — Necesita repuestos para reparar') : '') +
        _filaDetalle("Fecha activación", data.fechaActivacion || "No indicada") +
        _filaDetalle("Descripción", data.falla || "—") +
        repuestosRow +
      "</div>" + cotizBtn,
      "Revisá y aprobá esta solicitud en el DJI HUB PRO."
    );

    var mailOpts = { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR };
    var emailCC  = _emailReseller(data.reseller);
    if (emailCC) mailOpts.cc = emailCC;

    GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, "", mailOpts);
    var threadId = _capturarThreadId(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto);
    _logEmail(nOT, PORTAL_CONFIG.EMAIL_SUPERVISOR, "Supervisor (Portal)", asunto, "OK", threadId || "");
    if (emailCC) {
      _logEmail(nOT, emailCC, "Reseller (Portal)", asunto, "OK", threadId || "");
    }
  } catch(e) { Logger.log("_notificarNuevaOT: " + e); }
}

function _notificarLoteOT(reseller, items, ots, ss) {
  try {
    var asunto         = "Nuevo ingreso por LOTE desde Portal — " + reseller + " (" + ots.length + " equipos)";
    var detalleEquipos = "";

    for (var i = 0; i < items.length; i++) {
      detalleEquipos +=
        "<div style='background:#fff;border:1px solid #ddeef7;border-radius:6px;padding:10px 14px;margin-bottom:10px'>" +
          "<div style='font-size:12px;font-weight:bold;color:#00a3e0;margin-bottom:4px'>" + ots[i] + "</div>" +
          "<div style='font-size:12px;color:#444'><strong>Modelo:</strong> " + items[i].equipo + " | <strong>S/N:</strong> " + (items[i].sn || "—") + "</div>" +
          "<div style='font-size:12px;color:#444'><strong>Garantía:</strong> " + items[i].garantia + " | <strong>Gestión:</strong> " + items[i].circuito + "</div>" +
          "<div style='font-size:12px;color:#666;margin-top:4px'><em>Falla:</em> " + items[i].falla + "</div>" +
        "</div>";
    }

    var html = _construirEmailHTML(
      "Ingreso de Lote de Equipos", "Supervisor",
      "<p style='font-size:14px;color:#444;margin:0 0 20px'>El reseller <strong>" + reseller + "</strong> registró un lote de <strong>" + ots.length + " equipos</strong> desde el Portal.</p>" +
      "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:14px 16px'>" + detalleEquipos + "</div>",
      "Revisá y aprobá estas solicitudes en el DJI HUB PRO."
    );

    GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, "", {
      htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR
    });

    var hojaLog = ss.getSheetByName("EMAIL_LOGS");
    if (!hojaLog) { hojaLog = ss.insertSheet("EMAIL_LOGS"); hojaLog.appendRow(["Fecha","OT","Destinatario","Rol","Asunto","Estado"]); }
    hojaLog.appendRow([new Date(), "LOTE", PORTAL_CONFIG.EMAIL_SUPERVISOR, "Supervisor (Portal)", asunto, "OK"]);
  } catch(e) { Logger.log("_notificarLoteOT: " + e); }
}

function obtenerEmailLogsPorReseller(nombreReseller) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
    if (!hoja) return [];
    var datos = getSheetValues(hoja);
    var rB    = String(nombreReseller).trim().toLowerCase();

    var mapaOT = {};
    var hojaOT = getSheet(SCHEMA.SHEETS.OT);
    if (hojaOT) {
      var dOT = getSheetValues(hojaOT);
      for (var j = 1; j < dOT.length; j++) {
        var otStr  = String(dOT[j][2] || "").trim();
        var resStr = String(dOT[j][7] || "").trim().toLowerCase();
        if (otStr) mapaOT[otStr] = resStr;
      }
    }

    var out = [];
    for (var i = datos.length - 1; i >= 1; i--) {
      var f      = datos[i];
      var ot     = String(f[1] || "").trim();
      var rol    = String(f[3] || "");
      var asunto = String(f[4] || "");
      if (rol === "Reseller" && asunto.indexOf("Presupuesto") !== -1) {
        if (mapaOT[ot] && mapaOT[ot] === rB) {
          out.push({
            fecha:  f[0] instanceof Date ? Utilities.formatDate(f[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm") : String(f[0]),
            ot:     ot,
            asunto: asunto,
            estado: String(f[5] || "")
          });
          if (out.length >= 20) break;
        }
      }
    }
    return out;
  } catch(e) { Logger.log("obtenerEmailLogsPorReseller: " + e); return []; }
}

function _filaDetalle(label, valor) {
  return "<div style='display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #eef2f6'>" +
    "<span style='font-size:12px;color:#888;font-weight:500;flex-shrink:0;padding-right:16px'>" + label + "</span>" +
    "<span style='font-size:12px;color:#333;text-align:right'>" + valor + "</span></div>";
}

function _construirEmailHTML(titulo, saludo, cuerpo, footer) {
  return "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>" +
    "<body style='margin:0;padding:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif'>" +
    "<table width='100%' cellpadding='0' cellspacing='0' style='background:#f2f4f7;padding:28px 12px'><tr><td>" +
    "<table width='600' align='center' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%'>" +
    "<tr><td style='background:#00a3e0;border-radius:10px 10px 0 0;padding:24px 32px'>" +
      "<table width='100%' cellpadding='0' cellspacing='0'><tr>" +
        "<td><div style='background:#fff;border-radius:5px;padding:4px 10px;display:inline-block'><span style='font-size:14px;font-weight:700;color:#00a3e0;letter-spacing:.1em'>DJI</span></div></td>" +
        "<td align='right'><span style='color:rgba(255,255,255,.85);font-size:11px'>BIDCOMAGRO · Portal Resellers</span></td>" +
      "</tr></table>" +
      "<h1 style='color:#fff;font-size:18px;font-weight:600;margin:18px 0 0;line-height:1.35'>" + titulo + "</h1>" +
    "</td></tr>" +
    "<tr><td style='background:#fff;padding:28px 32px'>" +
      "<p style='font-size:14px;color:#666;margin:0 0 22px;line-height:1.5'>" + saludo + ":</p>" + cuerpo +
    "</td></tr>" +
    "<tr><td style='background:#f9f9f9;border-top:1px solid #eee;border-radius:0 0 10px 10px;padding:16px 32px'>" +
      "<p style='font-size:11px;color:#aaa;margin:0;line-height:1.7'>" + (footer || "") + "<br>Generado automáticamente por el Portal Resellers BIDCOMAGRO.</p>" +
    "</td></tr>" +
    "</table></td></tr></table></body></html>";
}
