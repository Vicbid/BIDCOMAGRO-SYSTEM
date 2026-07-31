// ============================================================
//  DJI HUB PRO v14.1 — Codigo.gs
//  Proyecto: DJI HUB PRO
//  Sheet ID: MASTER_SHEET_ID (Env.js) — ver getDb()/getSheet() ahí.
// @version 2.28
//
//  Router (doGet) + utilidades base compartidas por todo el proyecto:
//  sesión (identificarUsuario), logs (registrarLog/registrarEmailLog),
//  hilos de Gmail (_enviarConHilo/enviarEmail) y los primitivos HTML de
//  los mails (bloqueCard/filaDetalle/construirEmailHTML/construirTablaRepuestos).
//  El acceso a hojas (getSheet/getSheetValues) vive en Env.js — antes había
//  un getSheet() SEGUNDO acá mismo (apuntaba a getActiveSpreadsheet() en vez
//  de MASTER_SHEET_ID); se sacó por ambiguo y quedó uno solo (fix 2026-07-30).
//
//  El resto se reorganizó (2026-07-30, sin cambios funcionales) en:
//    HUB_OTs.js            — CRUD/listado de órdenes de trabajo + pedido de repuestos
//    HUB_Notificaciones.js — motor de mails por cambio de estado
//    HUB_Presupuestos.js   — armado y envío de presupuestos
//    HUB_Supervisor.js     — Command Center, métricas, búsquedas/historial
//    HUB_Facturacion.js    — solicitud de factura (XLS + mail)
//    HUB_Sistema.js        — batería, triggers/SLA/reporte mensual, mensajería
// ============================================================

function _fmtNum(n) {
  var num   = Math.abs(Number(n) || 0);
  var parts = num.toFixed(2).split('.');
  var intS  = parts[0];
  var result = '';
  for (var k = 0; k < intS.length; k++) {
    if (k > 0 && (intS.length - k) % 3 === 0) result += '.';
    result += intS[k];
  }
  return result + ',' + parts[1];
}

//  CONFIGURACIÓN — EDITÁ SOLO ESTA SECCIÓN

var CONFIG = {
  EMAIL_SUPERVISOR:   "soporteagrasdji@bidcom.com.ar",
  EMAIL_FACTURACION:  "Cecilia.f@bidcom.com.ar,lucia.c@bidcom.com.ar",  // Administración — solicitud de factura al finalizar OOW (igual que WOS)
  NOMBRE_REMITENTE:   "BIDCOMAGRO · Soporte Técnico DJI Agriculture",
  COL_EMAIL_RESELLER: 9,   // Columna J en hoja Resellers (A=0…J=9)
  PORTAL_URL:         "https://script.google.com/macros/s/TU_DEPLOYMENT_ID_PORTAL/exec",

  DIAS_ESTIMADOS: {
    "Taller-IW":10, "Taller-OOW":15,
    "Reseller-IW":7, "Reseller-OOW":10,
    "Reseller Propio-IW":5, "Reseller Propio-OOW":7
  },

  ESTADOS_NOTIFICAR_RESELLER: [
    "Abierto","Presupuesto rechazado",
    "Presupuesto aceptado","Espera de repuestos",
    "Repuestos enviados","En reparacion","Rechazado DJI","Sin respuesta · Cerrado","Finalizado",
    "Caso Enviado","Aprobado por DJI","Bateria enviada a reseller"
  ],
  ESTADOS_NOTIFICAR_TECNICO:    ["Abierto","Presupuesto aceptado","Repuestos enviados"],
  // Vacío: al finalizar ya se avisa al reseller ("Actualización de su Orden de Servicio").
  // El aviso "Alerta —" al supervisor en cada finalización era redundante (el supervisor es
  // quien finaliza) → se quitó "Finalizado". El supervisor sigue recibiendo URGENTE/Backorder,
  // y la solicitud de facturación (OOW) se manda igual desde el bloque 4.
  ESTADOS_NOTIFICAR_SUPERVISOR: [],
  SUPERVISOR_RECIBE_URGENTES:   true,
  SUPERVISOR_RECIBE_BACKORDER:  true
};

// ============================================================


// ── ENTRY POINT ─────────────────────────────────────────────
function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'guia') {
    return HtmlService.createHtmlOutputFromFile('GUIA_FLUJOS')
      .setTitle('HUB PRO — Guía de Flujos')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('DJI HUB PRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function getGuiaUrl() {
  return ScriptApp.getService().getUrl() + '?page=guia';
}



// ============================================================
//  SSO — IDENTIFICAR USUARIO INTERNO
//  Lee email activo, busca en Usuarios_Internos col B
//  Devuelve: { nombre, email, rol, esTecnico } o null
// ============================================================
function identificarUsuario() {
  try {
    var emailActivo = Session.getActiveUser().getEmail();
    if (!emailActivo) return null;
    var hoja = getSheet(SCHEMA.SHEETS.USUARIOS);
    if (!hoja) return null;
    var datos = getSheetValues(hoja);
    for (var i = 1; i < datos.length; i++) {
      var f        = datos[i];
      var nombre   = String(f[0]||"").trim();
      var email    = String(f[1]||"").trim().toLowerCase();
      var rol      = String(f[2]||"operador").trim().toLowerCase();
      var esTecnico = String(f[3]||"").trim().toLowerCase() === "si";
      if (!nombre || !email) continue;
      if (email === emailActivo.toLowerCase()) {
        return { nombre: nombre, email: email, rol: rol, esTecnico: esTecnico };
      }
    }
    return { nombre: emailActivo, email: emailActivo, rol: "operador", esTecnico: false };
  } catch(e) {
    Logger.log("identificarUsuario: " + e);
    return null;
  }
}


function registrarLog(ot, tec, em, acc, ant, nue, det) {
  try { getSheet(SCHEMA.SHEETS.LOGS).appendRow([new Date(), ot, tec, em, acc, ant, nue, det]); } catch(e) {
    var payload = JSON.stringify({ modulo: "registrarLog", hoja: "LOGS", ot: ot, accion: acc, error: e.toString() });
    Logger.log("ERROR_LOGS_APPEND: " + payload);
    console.log(payload);
  }
}




function registrarEmailLog(ot, destinatario, rol, asunto, estado, threadId) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
    if (!hoja) {
      hoja = getDb().insertSheet("EMAIL_LOGS");
      hoja.appendRow(["Fecha","OT","Destinatario","Rol","Asunto","Estado","ThreadID"]);
    }
    hoja.appendRow([new Date(), ot, destinatario, rol, asunto, estado, threadId || '']);
    invalidateSheetValues(SCHEMA.SHEETS.EMAIL_LOGS);
  } catch(e) { Logger.log("registrarEmailLog: " + e); }
}


function bloqueCard(titulo, texto, color) {
  return "<div style='background:rgba(0,0,0,.03);border-left:3px solid " + color + ";border-radius:0 6px 6px 0;padding:10px 14px;margin-bottom:12px'>" +
    "<p style='font-size:12px;font-weight:700;color:" + color + ";margin:0 0 3px'>" + titulo + "</p>" +
    "<p style='font-size:13px;color:#333;margin:0'>" + texto + "</p></div>";
}


function filaDetalle(label, valor) {
  return "<div style='display:flex;justify-content:space-between;align-items:center;padding:8px 0;border-bottom:1px solid #eef2f6'>" +
    "<span style='font-size:12px;color:#888;font-weight:500;flex-shrink:0;padding-right:16px'>" + label + "</span>" +
    "<span style='font-size:12px;color:#333;text-align:right'>" + valor + "</span></div>";
}


// cerrada=true (OT finalizada): los repuestos se muestran como entregados, sin marcar pendientes.
function construirTablaRepuestos(rep, cerrada) {
  if (!rep || rep === "Sin consumo de repuestos") return "";
  var ls = rep.split(" ; "), filas = "", hayBack = false;
  for (var i = 0; i < ls.length; i++) {
    var p = ls[i].split(" | ");
    if (p.length < 3) continue;
    var cod  = p[0].trim();
    var des  = (p[1]||"").replace("("+cod+")","").replace(/\(\s*\)/g,"").trim();
    var ped  = parseInt(p[2].split(" E:")[0].replace("P:",""))||0;
    var env  = parseInt(p[2].split(" E:")[1])||0;
    if (cerrada) env = ped; // OT cerrada: no hay pendientes, todo entregado
    var back = ped - env;
    if (back > 0) hayBack = true;
    var est = back > 0
      ? "<span style='color:#e67e22;font-weight:600'>Pendiente (" + back + ")</span>"
      : "<span style='color:#27ae60;font-weight:600'>OK</span>";
    filas += "<tr style='background:" + (back>0?"#fffaf5":"#fff") + "'>" +
      "<td style='padding:7px 10px;font-size:11px;color:#666;border-bottom:1px solid #f0f0f0'>" + cod + "</td>" +
      "<td style='padding:7px 10px;font-size:12px;color:#333;border-bottom:1px solid #f0f0f0'>" + des + "</td>" +
      "<td style='padding:7px 10px;font-size:12px;text-align:center;border-bottom:1px solid #f0f0f0'>" + env + "/" + ped + "</td>" +
      "<td style='padding:7px 10px;text-align:center;border-bottom:1px solid #f0f0f0'>" + est + "</td></tr>";
  }
  if (!filas) return "";
  return "<div style='margin-top:20px'>" +
    "<p style='font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 8px'>Repuestos</p>" +
    "<table style='width:100%;border-collapse:collapse;border:1px solid #e8e8e8'>" +
    "<thead><tr style='background:#f5f5f5'>" +
    "<th style='padding:7px 10px;font-size:10px;color:#888;text-align:left'>Código</th>" +
    "<th style='padding:7px 10px;font-size:10px;color:#888;text-align:left'>Descripción</th>" +
    "<th style='padding:7px 10px;font-size:10px;color:#888;text-align:center'>Env/Ped</th>" +
    "<th style='padding:7px 10px;font-size:10px;color:#888;text-align:center'>Estado</th>" +
    "</tr></thead><tbody>" + filas + "</tbody></table>" +
    (hayBack ? "<p style='font-size:11px;color:#e67e22;margin-top:8px'>Algunos repuestos están pendientes de envío.</p>" : "") +
    "</div>";
}


function construirEmailHTML(titulo, saludo, cuerpo, footer) {
  return "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>" +
    "<body style='margin:0;padding:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Helvetica,Arial,sans-serif'>" +
    "<table width='100%' cellpadding='0' cellspacing='0' style='background:#f2f4f7;padding:28px 12px'><tr><td>" +
    "<table width='600' align='center' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%'>" +
    "<tr><td style='background:#00a3e0;border-radius:10px 10px 0 0;padding:24px 32px'>" +
      "<table width='100%' cellpadding='0' cellspacing='0'><tr>" +
        "<td><div style='background:#fff;border-radius:5px;padding:4px 10px;display:inline-block'><span style='font-size:14px;font-weight:700;color:#00a3e0;letter-spacing:.1em'>DJI</span></div></td>" +
        "<td align='right'><span style='color:rgba(255,255,255,0.85);font-size:11px'>BIDCOMAGRO · Servicio Técnico Oficial</span></td>" +
      "</tr></table>" +
      "<h1 style='color:#fff;font-size:18px;font-weight:600;margin:18px 0 0;line-height:1.35'>" + titulo + "</h1>" +
    "</td></tr>" +
    "<tr><td style='background:#fff;padding:28px 32px'>" +
      "<p style='font-size:14px;color:#666;margin:0 0 22px;line-height:1.5'>" + saludo + ":</p>" + cuerpo +
    "</td></tr>" +
    "<tr><td style='background:#f9f9f9;border-top:1px solid #eee;border-radius:0 0 10px 10px;padding:16px 32px'>" +
      "<p style='font-size:11px;color:#aaa;margin:0;line-height:1.7'>" + (footer||"") + "<br>Generado automáticamente por DJI HUB PRO. No responda este email.</p>" +
    "</td></tr>" +
    "</table></td></tr></table></body></html>";
}


// Espera indexación y devuelve el Thread ID del último hilo enviado a este destinatario.
function _capturarThreadId(para, asunto) {
  try {
    var base   = asunto.replace(/^\[URGENTE\]\s*/i, '').replace(/^\[COPIA\]\s*/i, '');
    var query  = 'in:sent to:(' + para + ') subject:("' + base + '") newer_than:1d';
    var delays = [2000, 4000];
    for (var i = 0; i < delays.length; i++) {
      Utilities.sleep(delays[i]);
      var hilos = GmailApp.search(query, 0, 1);
      if (hilos.length) return hilos[0].getId();
    }
  } catch(e) { Logger.log('_capturarThreadId: ' + e); }
  return null;
}


// Devuelve el Thread ID ancla almacenado en EMAIL_LOGS para esta OT + destinatario.
function _obtenerThreadIdLog(ot, destinatario) {
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
  } catch(e) { Logger.log('_obtenerThreadIdLog: ' + e); }
  return null;
}


// Envía dentro del hilo existente de la OT (replyHtml), o crea uno nuevo (sendEmail).
// Retorna el Thread ID para persistir en EMAIL_LOGS.
function _enviarConHilo(ot, para, asunto, html) {
  if (!para || para.indexOf('@') === -1) { Logger.log('Email inválido: ' + para); return null; }
  var threadId = _obtenerThreadIdLog(ot, para);
  if (threadId) {
    try {
      var hilo = GmailApp.getThreadById(threadId);
      if (!hilo) throw new Error('Thread no encontrado');
      hilo.replyAll('', { htmlBody: html, name: CONFIG.NOMBRE_REMITENTE });
      Logger.log('✓ replyAll → ' + para + ' [' + asunto + '] [thread=' + threadId + ']');
      return threadId;
    } catch(eRep) {
      Logger.log('_enviarConHilo replyHtml: ' + eRep + ' — fallback a sendEmail');
    }
  }
  GmailApp.sendEmail(para, asunto, '', { htmlBody: html, name: CONFIG.NOMBRE_REMITENTE, replyTo: CONFIG.EMAIL_SUPERVISOR });
  Logger.log('✓ sendEmail → ' + para + ' [' + asunto + ']');
  return _capturarThreadId(para, asunto);
}


// Usado solo para emails de reporte sin OT asociada (SLA diario, reporte mensual).
function enviarEmail(para, asunto, html) {
  if (!para || para.indexOf('@') === -1) { Logger.log('Email inválido: ' + para); return; }
  GmailApp.sendEmail(para, asunto, '', { htmlBody: html, name: CONFIG.NOMBRE_REMITENTE, replyTo: CONFIG.EMAIL_SUPERVISOR });
  Logger.log('✓ Email → ' + para + ' [' + asunto + ']');
}
