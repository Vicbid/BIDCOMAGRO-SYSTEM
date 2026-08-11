// @version 1.2
// ============================================================
//  HUB PRO — Sistema: detección/reposición de batería, triggers
//  (SLA diario + reporte mensual + instalación), diagnóstico general,
//  mensajería HUB↔Portal Reseller.
//  Extraído de HUB_Código.js 2.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================



// ============================================================
//  REPOSICIÓN DE BATERÍA — disparo automático en "Aprobacion DJI"
// ============================================================
function esBateria(nombreEquipo) {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.EQUIPOS);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]||"").trim().toLowerCase() === String(nombreEquipo||"").trim().toLowerCase())
        return String(d[i][1]||"").trim().toLowerCase() === "bateria";
    }
    return false;
  } catch(e) { return false; }
}

// Normaliza la columna TIPO de EQUIPOS (valores libres cargados a mano: "Drone", "Generador",
// "Mavic", "Control Remoto", "bateria"...) a uno de los 5 tipos que maneja la Recepción de equipo
// (HUB_Recepcion.js): 'Dron' | 'Bateria' | 'Control' | 'Generador' | 'Mavic'. Con esto el modal
// preselecciona el tipo correcto según el equipo de la OT — antes solo distinguía "es batería"
// (esBateria) y todo lo demás cae en "Dron" por defecto, que es como un Generador terminaba
// mandando el mail de "recibimos el Drone".
function _normalizarTipoRecepcion(tipoRaw) {
  var t = String(tipoRaw || '').trim().toLowerCase();
  if (t === 'bateria' || t === 'batería') return 'Bateria';
  if (t === 'generador') return 'Generador';
  if (t === 'mavic') return 'Mavic';
  if (t === 'control remoto' || t === 'control') return 'Control';
  return 'Dron'; // "drone"/"dron" y cualquier valor no reconocido explícitamente
}


function obtenerEmailGestionLogistica() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.USUARIOS);
    for (var i = 1; i < d.length; i++) {
      var nombre = String(d[i][0]||"").trim().toLowerCase();
      if (nombre === "gestion logistica") {
        var em = String(d[i][1]||"").trim();
        return (em && em.indexOf("@") !== -1) ? em : null;
      }
    }
    return null;
  } catch(e) { return null; }
}


function enviarEmailReposicionBateria(data) {
  try {
    var emailLog = obtenerEmailGestionLogistica();
    if (!emailLog) { Logger.log("Sin email para Gestion Logistica"); return; }
    var emailReseller = obtenerEmailReseller(data.reseller);
    // Pedido del usuario: Administración también tiene que enterarse de este caso apenas se
    // aprueba, porque antes de que Logística pueda despachar la batería hace falta que carguen
    // un IdVenta (requisito para el envío) — sin eso Logística queda trabada esperando.
    var emailAdmin = getMailAdministracion();
    var asunto = "[BIDCOMAGRO] Reposición de batería — OT " + data.ot + " · " + data.equipo;

    var cuerpoDetalle =
      filaDetalle("Orden de Trabajo", "<strong>" + data.ot + "</strong>") +
      filaDetalle("Modelo de batería", "<strong style='color:#00a3e0'>" + data.equipo + "</strong>") +
      filaDetalle("Nº de Serie", data.sn || "—") +
      (data.cas ? filaDetalle("Caso DJI (CAS/FWR)", "<strong>" + data.cas + "</strong>") : "") +
      filaDetalle("Reseller", data.reseller) +
      filaDetalle("Garantía", "IW — En garantía");

    var htmlEmail = construirEmailHTML(
      "Reposición de Batería — " + data.ot,
      "Estimado equipo de Logística,",
      bloqueCard("Caso aprobado — Reposición requerida",
        "El reseller <strong>" + data.reseller + "</strong> completó los dos pasos requeridos: " +
        "carga del caso reconocido por DJI como en garantía y envío del scrap. " +
        "Se requiere reponer la batería al reseller.",
        "#00a3e0") +
      "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:16px'>" +
      cuerpoDetalle + "</div>" +
      bloqueCard(" Acción requerida",
        "Coordinar envío de batería de reemplazo al reseller <strong>" + data.reseller + "</strong>. " +
        "Requiere que Administración haya cargado el IdVenta correspondiente — sin eso no se puede despachar.",
        "#27ae60"),
      "Aviso automático generado por DJI HUB PRO."
    );

    var tidLog = _enviarConHilo(data.ot, emailLog, asunto, htmlEmail);
    registrarEmailLog(data.ot, emailLog, "Logística", asunto, "OK", tidLog || "");

    if (emailAdmin) {
      var asuntoA = "[BIDCOMAGRO] Cargar IdVenta — Reposición batería OT " + data.ot;
      var htmlAdmin = construirEmailHTML(
        "Reposición de Batería — " + data.ot,
        "Área de Administración,",
        bloqueCard("🧾 Acción requerida — Cargar IdVenta",
          "Este caso de batería en garantía ya fue <strong>aprobado por DJI</strong> y está listo para reponerse al " +
          "reseller <strong>" + data.reseller + "</strong>. Antes de que Logística pueda despachar la batería de " +
          "reemplazo necesitamos que se cargue el <strong>IdVenta</strong> correspondiente — es un requisito para el envío.",
          "#e67e22") +
        "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:16px'>" +
        cuerpoDetalle + "</div>",
        "Aviso automático generado por DJI HUB PRO."
      );
      try {
        var tidAdmin = _enviarConHilo(data.ot, emailAdmin, asuntoA, htmlAdmin);
        registrarEmailLog(data.ot, emailAdmin, "Administración", asuntoA, "OK", tidAdmin || "");
      } catch(eAdmin) {
        registrarEmailLog(data.ot, emailAdmin, "Administración", asuntoA, "ERROR: " + eAdmin.message, "");
      }
    } else {
      Logger.log("Sin email para Administración (reposición batería)");
    }

    if (emailReseller) {
      var asuntoR = "OT " + data.ot + " — " + data.equipo;
      var tidRes = _enviarConHilo(data.ot, emailReseller, asuntoR, htmlEmail);
      registrarEmailLog(data.ot, emailReseller, "Reseller", asuntoR, "OK", tidRes || "");
    }
    Logger.log("✓ Reposición batería → " + emailLog + (emailAdmin ? " + Administración" : ""));
  } catch(e) { Logger.log("enviarEmailReposicionBateria: " + e); }
}


// ============================================================
function verificarSLAs() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var datos = getSheetValues(hoja);
    var hoy   = new Date();
    var alertas = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (!f[2]) continue;
      var est = String(f[4]||"");
      if (est === "Finalizado" || est === "CANCELADO" || est === "Rechazado DJI" || est === "Sin respuesta · Cerrado") continue;
      var dias = (f[0] instanceof Date) ? Math.floor((hoy - f[0]) / 86400000) : 0;
      var urg  = String(f[17]).toUpperCase() === "URGENTE";
      if (!((urg && dias > 1) || (!urg && dias > 7))) continue;
      alertas.push({ ot:String(f[2]), reseller:String(f[7]||"—"), equipo:String(f[5]||"—"),
                      estado:est, dias:dias, urgente:urg, tecnico:String(f[9]||"Sin asignar") });
    }
    if (!alertas.length) { Logger.log("SLA: sin alertas hoy."); return; }
    alertas.sort(function(a,b){ return b.dias - a.dias; });
    var filas = "";
    alertas.forEach(function(a) {
      var col = a.dias > 7 ? "#e74c3c" : "#e67e22";
      filas += "<tr style='background:" + (a.urgente?"#fdf0f0":"#fff") + "'>" +
        "<td style='padding:8px 10px;font-size:12px;font-weight:600'>" + a.ot + "</td>" +
        "<td style='padding:8px 10px;font-size:12px'>" + a.reseller + "</td>" +
        "<td style='padding:8px 10px;font-size:12px'>" + a.equipo + "</td>" +
        "<td style='padding:8px 10px;font-size:12px'>" + a.estado + "</td>" +
        "<td style='padding:8px 10px;font-size:12px;font-weight:700;color:" + col + "'>" + a.dias + "d</td>" +
        "<td style='padding:8px 10px;font-size:12px'>" + a.tecnico + "</td></tr>";
    });
    var tabla = "<table style='width:100%;border-collapse:collapse;border:1px solid #e8e8e8'>" +
      "<thead><tr style='background:#f5f5f5'>" +
      "<th style='padding:8px 10px;font-size:10px;color:#888;text-align:left'>OT</th>" +
      "<th style='padding:8px 10px;font-size:10px;color:#888;text-align:left'>Reseller</th>" +
      "<th style='padding:8px 10px;font-size:10px;color:#888;text-align:left'>Equipo</th>" +
      "<th style='padding:8px 10px;font-size:10px;color:#888;text-align:left'>Estado</th>" +
      "<th style='padding:8px 10px;font-size:10px;color:#888;text-align:left'>Días</th>" +
      "<th style='padding:8px 10px;font-size:10px;color:#888;text-align:left'>Técnico</th>" +
      "</tr></thead><tbody>" + filas + "</tbody></table>";
    enviarEmail(
      CONFIG.EMAIL_SUPERVISOR,
      "[HUB] " + alertas.length + " orden(es) sin movimiento — " + Utilities.formatDate(hoy, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      construirEmailHTML("Reporte diario — Órdenes sin movimiento", "Supervisor",
        "<p style='font-size:14px;color:#444;margin:0 0 20px'>Hay <strong style='color:#e74c3c'>" + alertas.length + " orden(es)</strong> que superaron el tiempo máximo sin actualización.</p>" + tabla,
        "Reporte automático diario a las 8:00 AM.")
    );
  } catch(e) { Logger.log("verificarSLAs: " + e); }
}



// ============================================================
//  TRIGGER MENSUAL — REPORTE
// ============================================================
function reporteMensual() {
  try {
    var hoja   = getSheet(SCHEMA.SHEETS.OT);
    var datos  = getSheetValues(hoja);
    var hoy    = new Date();
    var primerDiaMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    var totAb=0, totCerradas=0, sumasDias=0, countC=0;
    var porReseller={}, porTecnico={};

    for (var i = 1; i < datos.length; i++) {
      var f = datos[i]; if(!f[2]) continue;
      var fechaCierre = f[1] instanceof Date ? f[1] : null;
      var est = String(f[4]||"");
      var res = String(f[7]||"Particular");
      var tec = String(f[9]||"—");
      var dias = (f[0] instanceof Date) ? Math.floor((hoy-f[0])/86400000) : 0;
      if (est !== "Finalizado") {
        totAb++;
        if (!porReseller[res]) porReseller[res]={ab:0,fin:0};
        porReseller[res].ab++;
      }
      if (est === "Finalizado" && fechaCierre && fechaCierre >= primerDiaMes) {
        totCerradas++;
        if (!porReseller[res]) porReseller[res]={ab:0,fin:0};
        porReseller[res].fin++;
        if (tec && tec !== "Gestión Reseller" && tec !== "—") {
          if (!porTecnico[tec]) porTecnico[tec]={fin:0,dias:0};
          porTecnico[tec].fin++; porTecnico[tec].dias += dias;
        }
        sumasDias += dias; countC++;
      }
    }
    var prom = countC > 0 ? Math.round(sumasDias/countC) : 0;
    var mes  = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"][hoy.getMonth()];

    function statBox(num, label, color) {
      return "<div style='flex:1;background:#f5f9fc;border-top:3px solid "+color+";border-radius:6px;padding:14px;text-align:center'>" +
        "<div style='font-size:28px;font-weight:700;color:"+color+"'>"+num+"</div>" +
        "<div style='font-size:11px;color:#888;margin-top:4px'>"+label+"</div></div>";
    }

    var filasRes = "", filasTec = "";
    Object.keys(porReseller).sort().forEach(function(r){
      var d = porReseller[r];
      filasRes += "<tr><td style='padding:7px 10px;font-size:12px'>"+r+"</td><td style='padding:7px 10px;font-size:12px;text-align:center'>"+d.ab+"</td><td style='padding:7px 10px;font-size:12px;text-align:center;color:#27ae60;font-weight:600'>"+d.fin+"</td></tr>";
    });
    Object.keys(porTecnico).sort().forEach(function(t){
      var d = porTecnico[t];
      var p = d.fin>0?Math.round(d.dias/d.fin):0;
      var col = p<=5?"#27ae60":(p<=10?"#f39c12":"#e74c3c");
      filasTec += "<tr><td style='padding:7px 10px;font-size:12px'>"+t+"</td><td style='padding:7px 10px;font-size:12px;text-align:center'>"+d.fin+"</td><td style='padding:7px 10px;font-size:12px;text-align:center;font-weight:600;color:"+col+"'>"+p+"d</td></tr>";
    });

    var cuerpo =
      "<div style='display:flex;gap:12px;margin-bottom:20px'>" +
        statBox(String(totAb), "Órdenes activas", "#00a3e0") +
        statBox(String(totCerradas), "Cerradas en " + mes, "#27ae60") +
        statBox(prom+"d", "Promedio resolución", "#f39c12") +
      "</div>" +
      (filasRes ? "<p style='font-size:12px;font-weight:700;color:#444;margin:16px 0 8px;text-transform:uppercase;letter-spacing:.06em'>Por Reseller</p><table style='width:100%;border-collapse:collapse;border:1px solid #e8e8e8'><thead><tr style='background:#f5f5f5'><th style='padding:7px 10px;font-size:10px;color:#888;text-align:left'>Reseller</th><th style='padding:7px 10px;font-size:10px;color:#888;text-align:center'>Abiertas</th><th style='padding:7px 10px;font-size:10px;color:#888;text-align:center'>Cerradas</th></tr></thead><tbody>"+filasRes+"</tbody></table>" : "") +
      (filasTec ? "<p style='font-size:12px;font-weight:700;color:#444;margin:16px 0 8px;text-transform:uppercase;letter-spacing:.06em'>Por Técnico</p><table style='width:100%;border-collapse:collapse;border:1px solid #e8e8e8'><thead><tr style='background:#f5f5f5'><th style='padding:7px 10px;font-size:10px;color:#888;text-align:left'>Técnico</th><th style='padding:7px 10px;font-size:10px;color:#888;text-align:center'>Finalizadas</th><th style='padding:7px 10px;font-size:10px;color:#888;text-align:center'>Prom. días</th></tr></thead><tbody>"+filasTec+"</tbody></table>" : "");

    enviarEmail(
      CONFIG.EMAIL_SUPERVISOR,
      "[HUB] Reporte mensual " + mes + " " + hoy.getFullYear(),
      construirEmailHTML("Reporte Mensual — " + mes + " " + hoy.getFullYear(), "Supervisor", cuerpo,
        "Este reporte se genera automáticamente el 1º de cada mes.")
    );
    Logger.log("✓ Reporte mensual enviado");
  } catch(e) { Logger.log("reporteMensual: " + e); }
}



// ============================================================
//  INSTALACIÓN DE TRIGGERS (ejecutar UNA VEZ desde el editor)
// ============================================================
function instalarTodosTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var fn = triggers[i].getHandlerFunction();
    if (fn === "verificarSLAs" || fn === "reporteMensual") ScriptApp.deleteTrigger(triggers[i]);
  }
  ScriptApp.newTrigger("verificarSLAs").timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger("reporteMensual").timeBased().onMonthDay(1).atHour(8).create();
  Logger.log("✓ Trigger diario (SLA) y mensual instalados.");
}


function diagnosticarSistema() {
  Logger.log("=== DIAGNÓSTICO DJI HUB PRO ===");
  var hojaRes = getSheet(SCHEMA.SHEETS.RESELLERS);
  if (!hojaRes) { Logger.log("✘ Hoja 'Resellers' no encontrada"); return; }
  var d = getSheetValues(hojaRes);
  Logger.log("Resellers: " + (d.length-1) + " filas | Col email: " + (CONFIG.COL_EMAIL_RESELLER+1) + " (" + String.fromCharCode(65+CONFIG.COL_EMAIL_RESELLER) + ")");
  Logger.log("Headers: " + d[0].join(" | "));
  var ok=0, sin=0;
  for (var i=1; i<d.length; i++) {
    var n=String(d[i][0]||"").trim(), em=String(d[i][CONFIG.COL_EMAIL_RESELLER]||"").trim();
    if (!n) continue;
    if (em && em.indexOf("@")!==-1) { ok++; Logger.log("✓ "+n+" → "+em); }
    else { sin++; Logger.log("✘ "+n+" → SIN EMAIL ('"+em+"')"); }
  }
  Logger.log("Con email: "+ok+" | Sin email: "+sin);
  try { Logger.log("Gmail: " + MailApp.getRemainingDailyQuota() + " emails/día"); } catch(e) { Logger.log("✘ Gmail: "+e); }
  Logger.log("==============================");
}


function marcarMensajesLeidos(fila) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var cell = hoja.getRange(fila, SCHEMA.OT.MENSAJES + 1);
    var actual = String(cell.getValue() || "").trim();
    if (!actual) return { ok: true };
    var lastMsg   = actual.lastIndexOf("💬");
    var lastLeido = actual.lastIndexOf("[LEIDO]");
    if (lastMsg === -1 || lastLeido > lastMsg) return { ok: true };
    cell.setValue(actual + "\n\n[LEIDO]");
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);
    return { ok: true };
  } catch(e) {
    Logger.log("marcarMensajesLeidos: " + e);
    return { ok: false };
  }
}


function enviarMensajeHUB(fila, texto) {
  try {
    var filaNum = parseInt(fila);
    if (isNaN(filaNum) || filaNum < 2) return { ok: false, msg: 'Fila inválida' };
    if (!texto || !String(texto).trim()) return { ok: false, msg: 'Mensaje vacío' };
    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var cell = hoja.getRange(filaNum, SCHEMA.OT.MENSAJES + 1);
    var actual = String(cell.getValue() || "").trim();
    var autor  = Session.getActiveUser().getEmail() || 'BIDCOMAGRO';
    var fecha  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var bloque = "💬 [" + fecha + "] — " + autor + ":\n" + String(texto).trim();
    var nuevo  = actual ? actual + "\n\n" + bloque : bloque;
    cell.setValue(nuevo);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);

    // Notificar al reseller por email
    try {
      var row      = hoja.getRange(filaNum, 1, 1, SCHEMA.OT.RESELLER + 2).getValues()[0];
      var ot       = String(row[SCHEMA.OT.OT]       || '');
      var equipo   = String(row[SCHEMA.OT.EQUIPO]   || '');
      var reseller = String(row[SCHEMA.OT.RESELLER] || '');
      var emailR   = obtenerEmailReseller(reseller);
      if (emailR) {
        var asunto  = 'OT ' + ot + (equipo ? ' — ' + equipo : '');
        var textoHtml = String(texto).trim().replace(/\n/g, '<br>');
        var cuerpo =
          "<div style='background:#f0f7ff;border-left:4px solid #00a3e0;border-radius:4px;padding:14px 18px;margin:0 0 18px'>" +
            "<p style='font-size:13px;color:#333;margin:0;line-height:1.7'>" + textoHtml + "</p>" +
          "</div>" +
          "<p style='font-size:13px;color:#555;margin:0'>Podés responder desde tu portal de reseller o contestando este email.</p>";
        var html = construirEmailHTML(
          'Nuevo mensaje en tu orden de trabajo',
          'Hola, el equipo de soporte te dejó un mensaje en la OT <strong>' + ot + '</strong>' + (equipo ? ' (' + equipo + ')' : ''),
          cuerpo,
          'OT: ' + ot + ' · Reseller: ' + reseller
        );
        var tidMsg = _enviarConHilo(ot, emailR, asunto, html);
        registrarEmailLog(ot, emailR, 'Reseller', asunto, 'OK', tidMsg || '');
      }
    } catch(em) { Logger.log('enviarMensajeHUB notif: ' + em); }

    return { ok: true };
  } catch(e) {
    Logger.log("enviarMensajeHUB: " + e);
    return { ok: false, msg: e.toString() };
  }
}
