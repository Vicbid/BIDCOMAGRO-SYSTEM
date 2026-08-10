// @version 1.2
// ============================================================
//  HUB PRO — Motor de notificaciones por mail ante cambio de estado
//  de una OT (reseller/técnico/supervisor/facturación) + las 2
//  herramientas de recuperación manual (HUB_notificarOT,
//  HUB_diagnosticoBateria) — genéricas, no atadas a un incidente
//  puntual, se mantienen por si hace falta reenviar/diagnosticar un
//  aviso puntual desde el editor.
//  Extraído de HUB_Código.js 2.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// ============================================================
//  RECUPERAR MAIL DE APERTURA — OT creadas desde el HUB antes del
//  fix v2.14 (crearNuevaOT no llamaba a enviarNotificaciones).
//  Lee la fila ACTUAL de la OT y dispara el aviso de estado.
//  NO modifica la OT: solo manda el mail y lo registra en EMAIL_LOGS.
//    HUB_notificarOT("WH/REP/00123")        → mail del estado ACTUAL de la OT
//    HUB_notificarOT("WH/REP/00123", true)  → fuerza el texto de "Abierto"
// ============================================================
function HUB_notificarOT(numeroOT, forzarAbierto) {
  try {
    var buscado = String(numeroOT || "").trim();
    if (!buscado) return { ok: false, error: "Falta el número de OT" };

    var datos = getSheetValues(SCHEMA.SHEETS.OT, true);  // force: leer del sheet, no del cache
    var O = SCHEMA.OT;
    var f = null;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][O.OT] || "").trim() === buscado) { f = datos[i]; break; }
    }
    if (!f) return { ok: false, error: "OT no encontrada: " + buscado };

    var estadoActual = String(f[O.ESTADO] || "").trim();
    var estadoNotif  = forzarAbierto ? "Abierto" : estadoActual;
    var tecnico      = String(f[O.TECNICO] || "").trim();

    var data = {
      ot:        buscado,
      estado:    estadoNotif,
      reseller:  String(f[O.RESELLER]  || "").trim(),
      equipo:    String(f[O.EQUIPO]    || "").trim(),
      sn:        String(f[O.SN]        || "").trim(),
      garantia:  String(f[O.GARANTIA]  || "OOW").trim(),
      circuito:  String(f[O.CIRCUITO]  || "Taller").trim(),
      cas:       String(f[O.CAS]       || "").trim(),
      cliente:   String(f[O.CLIENTE]   || "").trim(),
      trabajo:   String(f[O.TRABAJO]   || "").trim(),
      informeTecnico: String(f[O.INFORME_TECNICO] || "").trim(),
      repuestos: String(f[O.REPUESTOS] || "").trim(),
      prioridad: (String(f[O.PRIORIDAD] || "").toUpperCase() === "URGENTE")
    };
    data._fechaEstimada = calcularFechaEstimada(data.circuito, data.garantia, f[O.FECHA_INGRESO]);

    // estadoAnterior "-" → enviarNotificaciones lo trata como aviso nuevo (estado != anterior)
    enviarNotificaciones(data, "-", tecnico);

    var avisaReseller = estaEnLista(estadoNotif, CONFIG.ESTADOS_NOTIFICAR_RESELLER);
    return {
      ok: true,
      ot: buscado,
      estadoNotificado: estadoNotif,
      reseller: data.reseller,
      nota: avisaReseller
        ? "Aviso disparado — revisá EMAIL_LOGS para confirmar el envío."
        : "El estado '" + estadoNotif + "' no está en ESTADOS_NOTIFICAR_RESELLER: no se manda mail al reseller (usá forzarAbierto=true si querés el mail de apertura)."
    };
  } catch(e) {
    Logger.log("HUB_notificarOT: " + e);
    return { ok: false, error: e.toString() };
  }
}


// Wrapper editable para correr desde el editor de Apps Script.
// Poné uno o varios números separados por coma. FORZAR_ABIERTO=true manda el
// texto de "Abierto" aunque la OT ya haya avanzado de estado.
function HUB_notificarOTTest() {
  var NUMEROS        = "WH/REP/00000";   // ej: "WH/REP/00123, WH/REP/00124"
  var FORZAR_ABIERTO = false;

  var lista = String(NUMEROS).split(",");
  var res   = [];
  for (var i = 0; i < lista.length; i++) {
    var n = lista[i].trim();
    if (!n) continue;
    res.push(HUB_notificarOT(n, FORZAR_ABIERTO));
  }
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}


// ============================================================
//  DIAGNÓSTICO DE MAILS DE BATERÍA — read-only, no manda nada.
//  Dice por qué una OT de batería no está enviando mails.
//    HUB_diagnosticoBateria("WH/REP/00123")
//  o editá HUB_diagnosticoBateriaTest() y corré desde el editor.
// ============================================================
function HUB_diagnosticoBateria(numeroOT) {
  try {
    var buscado = String(numeroOT || "").trim();
    if (!buscado) return { ok: false, error: "Falta el número de OT" };

    var datos = getSheetValues(SCHEMA.SHEETS.OT, true);
    var O = SCHEMA.OT, f = null;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][O.OT] || "").trim() === buscado) { f = datos[i]; break; }
    }
    if (!f) return { ok: false, error: "OT no encontrada: " + buscado };

    var equipo   = String(f[O.EQUIPO]   || "").trim();
    var estado   = String(f[O.ESTADO]   || "").trim();
    var reseller = String(f[O.RESELLER] || "").trim();

    // Tipo del equipo tal como figura en la hoja EQUIPOS (col B)
    var tipoEquipo = "(equipo no encontrado en EQUIPOS)";
    var dEq = getSheetValues(SCHEMA.SHEETS.EQUIPOS);
    for (var e = 1; e < dEq.length; e++) {
      if (String(dEq[e][0] || "").trim().toLowerCase() === equipo.toLowerCase()) {
        tipoEquipo = String(dEq[e][1] || "").trim(); break;
      }
    }

    var esBat          = esBateria(equipo);
    var emailReseller  = obtenerEmailReseller(reseller);
    var emailLogistica = obtenerEmailGestionLogistica();
    var enLista        = estaEnLista(estado, CONFIG.ESTADOS_NOTIFICAR_RESELLER);

    var problemas = [];
    if (!esBat) problemas.push("esBateria('" + equipo + "')=false → NO dispara la reposición ni los textos de batería. Tipo en EQUIPOS col B = '" + tipoEquipo + "' (debe ser exactamente 'bateria', minúscula y sin tilde).");
    if (!emailReseller) problemas.push("El reseller '" + reseller + "' no tiene email válido en Resellers (col J) → no le llega nada.");
    if (!emailLogistica) problemas.push("No hay usuario 'Gestion Logistica' con email válido en Usuarios_Internos → el mail de reposición (estado 'Scrap Enviado (Evidencias)') se ABORTA entero: ni logística ni la copia al reseller, y no queda registro en EMAIL_LOGS.");
    if (!enLista) problemas.push("El estado actual '" + estado + "' NO está en ESTADOS_NOTIFICAR_RESELLER → en este estado no sale mail de cambio de estado (es esperable si el estado es 'Scrap Enviado (Evidencias)', que usa el mail de reposición aparte).");

    var r = {
      ok: true,
      ot: buscado,
      equipo: equipo,
      tipoEnEquipos: tipoEquipo,
      esBateria: esBat,
      estadoActual: estado,
      estadoAvisaAlReseller: enLista,
      reseller: reseller,
      emailReseller: emailReseller || "(SIN EMAIL)",
      emailGestionLogistica: emailLogistica || "(SIN USUARIO / SIN EMAIL)",
      problemas: problemas.length ? problemas : ["Sin problemas de config detectados: en este estado la OT debería mandar mail. Si igual no llega, revisá EMAIL_LOGS por filas 'ERROR:' de esta OT."]
    };
    Logger.log(JSON.stringify(r, null, 2));
    return r;
  } catch(e) {
    Logger.log("HUB_diagnosticoBateria: " + e);
    return { ok: false, error: e.toString() };
  }
}


function HUB_diagnosticoBateriaTest() {
  var NUMERO = "WH/REP/00000";   // poné acá una OT de batería que no esté mandando mails
  return HUB_diagnosticoBateria(NUMERO);
}


function calcularFechaEstimada(circuito, garantia, fechaApertura) {
  var key  = circuito + "-" + garantia;
  var dias = CONFIG.DIAS_ESTIMADOS[key] || 10;
  var base = (fechaApertura instanceof Date) ? fechaApertura : new Date();
  var est  = new Date(base.getTime() + dias * 86400000);
  while (est.getDay() === 0 || est.getDay() === 6) est.setDate(est.getDate() + 1);
  return Utilities.formatDate(est, Session.getScriptTimeZone(), "dd/MM/yyyy");
}



// ============================================================
//  SISTEMA DE NOTIFICACIONES POR EMAIL
// ============================================================
function enviarNotificaciones(data, estadoAnterior, tecnico) {
  try {
    var estadoNuevo    = data.estado;
    var tieneBackorder = detectarBackorder(data.repuestos);
    // El aviso de "Backorder / Repuestos pendientes de envío" habla de un ENVÍO (repuesto que
    // viaja de BIDCOMAGRO al reseller) — no aplica a Taller: ahí el repuesto se usa en el
    // propio taller, no se "envía" a ningún lado. Se filtra acá para no mandarlo en OTs Taller.
    var tieneBackorderEnvio = tieneBackorder && data.circuito !== "Taller";
    Logger.log("=== NOTIF " + data.ot + " | " + estadoAnterior + " → " + estadoNuevo + " ===");

    // 0. REPOSICIÓN BATERÍA — se dispara al confirmar el envío del scrap a DJI
    if (estadoNuevo === "Scrap Enviado (Evidencias)" && estadoAnterior !== "Scrap Enviado (Evidencias)") {
      if (esBateria(data.equipo)) {
        enviarEmailReposicionBateria(data);
      }
    }

    // 1. RESELLER — solo cuando CAMBIA de estado (no en cada guardado ni al re-guardar el mismo estado)
    if (estadoNuevo !== estadoAnterior && estaEnLista(estadoNuevo, CONFIG.ESTADOS_NOTIFICAR_RESELLER)) {
      var emailR = obtenerEmailReseller(data.reseller);
      if (emailR) {
        var asuntoR = armarAsunto(data);
        try {
          var tidR = _enviarConHilo(data.ot, emailR, asuntoR, armarEmailReseller(data, estadoAnterior, estadoNuevo, tecnico));
          registrarEmailLog(data.ot, emailR, "Reseller", asuntoR, "OK", tidR || "");
        } catch(e) {
          registrarEmailLog(data.ot, emailR, "Reseller", asuntoR, "ERROR: " + e.message, "");
        }
      } else {
        registrarEmailLog(data.ot, data.reseller, "Reseller", "—", "SIN EMAIL CONFIGURADO", "");
        Logger.log("✘ Sin email para: " + data.reseller);
      }
    }

    // 2. TÉCNICO
    if (tecnico && tecnico !== "Gestión Reseller" && estaEnLista(estadoNuevo, CONFIG.ESTADOS_NOTIFICAR_TECNICO)) {
      var emailT = obtenerEmailTecnico(tecnico);
      if (emailT) {
        var asuntoT = "[ASIGNADO] OT " + data.ot + " — " + data.equipo;
        try {
          var tidT = _enviarConHilo(data.ot, emailT, asuntoT, armarEmailTecnico(data, estadoNuevo, tecnico));
          registrarEmailLog(data.ot, emailT, "Técnico", asuntoT, "OK", tidT || "");
        } catch(e) {
          registrarEmailLog(data.ot, emailT, "Técnico", asuntoT, "ERROR: " + e.message, "");
        }
      }
    }

    // 3. SUPERVISOR
    var motivoSup = [];
    if (estaEnLista(estadoNuevo, CONFIG.ESTADOS_NOTIFICAR_SUPERVISOR)) motivoSup.push(estadoNuevo);
    if (CONFIG.SUPERVISOR_RECIBE_URGENTES  && data.prioridad)   motivoSup.push("URGENTE");
    if (CONFIG.SUPERVISOR_RECIBE_BACKORDER && tieneBackorderEnvio) motivoSup.push("Backorder");
    if (motivoSup.length > 0) {
      var asuntoS = "[HUB] " + motivoSup.join(" · ") + " — " + data.ot;
      try {
        var tidS = _enviarConHilo(data.ot, CONFIG.EMAIL_SUPERVISOR, asuntoS, armarEmailSupervisor(data, estadoAnterior, estadoNuevo, tecnico, tieneBackorderEnvio));
        registrarEmailLog(data.ot, CONFIG.EMAIL_SUPERVISOR, "Supervisor", asuntoS, "OK", tidS || "");
      } catch(e) {
        registrarEmailLog(data.ot, CONFIG.EMAIL_SUPERVISOR, "Supervisor", asuntoS, "ERROR: " + e.message, "");
      }
    }

    // 4. FACTURACIÓN (Administración) — solo al FINALIZAR una OT fuera de garantía (con presupuesto).
    //    Se dispara únicamente en la transición a "Finalizado" (estadoAnterior distinto) para no duplicar la solicitud.
    if (estadoNuevo === "Finalizado" && estadoAnterior !== "Finalizado"
        && String(data.garantia || "").toUpperCase() === "OOW") {
      var asuntoF = "[FACTURAR] OT " + data.ot + " — " + data.reseller + " · " + data.equipo;
      try {
        var tidF = _enviarConHilo(data.ot, CONFIG.EMAIL_FACTURACION, asuntoF, armarEmailFacturacion(data, tecnico));
        registrarEmailLog(data.ot, CONFIG.EMAIL_FACTURACION, "Facturación", asuntoF, "OK", tidF || "");
      } catch(e) {
        registrarEmailLog(data.ot, CONFIG.EMAIL_FACTURACION, "Facturación", asuntoF, "ERROR: " + e.message, "");
      }
    }
  } catch(e) { Logger.log("enviarNotificaciones: " + e); }
}


// Email a Administración solicitando la facturación de una OT finalizada fuera de garantía (OOW).
// Reutiliza los datos del presupuesto: mano de obra (manoObraGuardada) + repuestos consumidos.
function armarEmailFacturacion(data, tecnico) {
  // Mano de obra: manoObraGuardada es un JSON [{codigo,descripcion,precio}]
  var mo = [];
  try { if (data.manoObraGuardada) mo = JSON.parse(data.manoObraGuardada); } catch(e) {}
  var tablaMOHTML = "", totalMO = 0;
  if (mo && mo.length) {
    var filasMO = "";
    for (var i = 0; i < mo.length; i++) {
      var pMO = parseFloat(String(mo[i].precio).replace(/[^0-9.]/g, "")) || 0;
      totalMO += pMO;
      filasMO += "<tr>" +
        "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + mo[i].codigo + "</td>" +
        "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + mo[i].descripcion + "</td>" +
        "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + mo[i].precio + "</td></tr>";
    }
    tablaMOHTML = "<div style='margin-top:20px'>" +
      "<p style='font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 8px'>Mano de Obra</p>" +
      "<table style='width:100%;border-collapse:collapse;font-size:12px'>" +
      "<thead><tr style='background:#f5f5f5'>" +
      "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
      "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
      "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Precio</th>" +
      "</tr></thead><tbody>" + filasMO +
      (totalMO > 0 ? "<tr style='background:#f5f5f5;font-weight:700'>" +
        "<td colspan='2' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL MANO DE OBRA</td>" +
        "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalMO) + "</td></tr>" : "") +
      "</tbody></table></div>";
  }

  var cuerpoDetalle =
    filaDetalle("Orden de Trabajo", data.ot) +
    filaDetalle("Reseller", data.reseller) +
    filaDetalle("Equipo / Modelo", data.equipo) +
    filaDetalle("Nº de Serie", data.sn || "—") +
    (data.cas ? filaDetalle("Caso DJI (CAS/FWR)", data.cas) : "") +
    filaDetalle("Garantía", "Fuera de garantía (OOW)") +
    (tecnico && tecnico !== "Gestión Reseller" ? filaDetalle("Técnico", tecnico) : "");

  var repFact = construirTablaRepuestosFacturacion(data.repuestos);
  var totalGeneral = Math.round((totalMO + repFact.total) * 100) / 100;
  var totalFilaHTML = totalGeneral > 0
    ? "<table style='width:100%;border-collapse:collapse;margin-top:14px'>" +
      "<tr style='background:#111;color:#fff;font-weight:700;font-size:14px'>" +
      "<td style='padding:10px 14px;text-align:right'>TOTAL A FACTURAR (USD, sin impuestos)</td>" +
      "<td style='padding:10px 14px;text-align:right;white-space:nowrap'>USD " + _fmtNum(totalGeneral) + "</td></tr></table>"
    : "";

  var bloques =
    bloqueCard("🧾 Solicitud de facturación",
      "Esta orden fue <strong>finalizada fuera de garantía</strong>. Solicitamos generar la factura correspondiente al reseller.", "#e67e22") +
    "<div style='background:rgba(0,0,0,.02);border:1px solid #eef2f6;border-radius:8px;padding:6px 16px;margin-bottom:12px'>" + cuerpoDetalle + "</div>" +
    repFact.html +
    tablaMOHTML +
    totalFilaHTML +
    "<p style='font-size:11px;color:#888;margin-top:14px'>Repuestos con precio neto reseller (40% de descuento sobre PVP), sin impuestos. Los importes finales son los del presupuesto enviado al reseller para esta orden.</p>";

  return construirEmailHTML(
    "Solicitud de facturación — OT " + data.ot,
    "Área de Administración,<br>Solicitamos la facturación de la siguiente orden finalizada fuera de garantía:",
    bloques,
    "Cualquier duda, respondé este correo."
  );
}


function estaEnLista(v, lista) {
  for (var i = 0; i < lista.length; i++) if (lista[i] === v) return true;
  return false;
}


function detectarBackorder(rep) {
  if (!rep || rep === "Sin consumo de repuestos") return false;
  var ls = rep.split(" ; ");
  for (var i = 0; i < ls.length; i++) {
    var p = ls[i].split(" | ");
    if (p.length < 3) continue;
    if ((parseInt(p[2].split(" E:")[0].replace("P:",""))||0) > (parseInt(p[2].split(" E:")[1])||0)) return true;
  }
  return false;
}


function obtenerEmailReseller(nombre) {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toLowerCase() === String(nombre).trim().toLowerCase()) {
        var em = String(d[i][CONFIG.COL_EMAIL_RESELLER]||"").trim();
        return (em && em.indexOf("@") !== -1) ? em : null;
      }
    }
    return null;
  } catch(e) { return null; }
}


function obtenerEmailTecnico(nombre) {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.LOGS);
    for (var i = d.length - 1; i >= 1; i--) {
      if (String(d[i][2]).trim() === String(nombre).trim()) {
        var em = String(d[i][3]||"").trim();
        if (em && em.indexOf("@") !== -1) return em;
      }
    }
    return null;
  } catch(e) { return null; }
}


function armarAsunto(data) {
  var base = (data.prioridad ? "[URGENTE] " : "") + "OT " + data.ot + " — " + data.equipo;
  return (data.cliente && data.cliente.trim()) ? base + " · " + data.cliente.trim() : base;
}


function armarEmailReseller(data, ant, nvo, tec) {
  var esReseller = data.circuito === "Reseller" || data.circuito === "Reseller Propio";
  var msgsTaller = {
    "Abierto":              "Hemos recibido el equipo y abierto la orden de trabajo. Nuestro equipo lo revisará a la brevedad.",
    "Presupuesto enviado":  "Hemos completado el diagnóstico. A continuación le enviamos el presupuesto para su aprobación.",
    "Presupuesto rechazado":"Hemos recibido su decisión de no proceder. Coordinaremos la devolución del equipo.",
    "Presupuesto aceptado": "Hemos recibido su aprobación. Iniciamos la reparación de inmediato.",
    "Espera de repuestos":  "Necesitamos repuestos específicos. Estamos gestionando el pedido.",
    "Repuestos enviados":   "Los repuestos fueron despachados. Retomamos la reparación al recibirlos.",
    "En reparacion":        "Su equipo se encuentra en proceso de reparación activa.",
    "Finalizado":           "La reparación fue completada exitosamente. Coordine el retiro con nosotros."
  };
  var msgsReseller = {
    "Abierto":              "La orden de trabajo fue registrada en el sistema. Revisaremos los detalles del caso y le informaremos sobre los próximos pasos.",
    "Presupuesto enviado":  "El diagnóstico del equipo fue completado. A continuación encontrará el presupuesto para su revisión y aprobación.",
    "Presupuesto rechazado":"Hemos registrado su decisión de no proceder con la reparación. La orden quedará cerrada en el sistema.",
    "Presupuesto aceptado": "Hemos recibido su aprobación. La reparación puede iniciarse.",
    "Espera de repuestos":  "Se requieren repuestos específicos para completar la reparación. Estamos gestionando el pedido.",
    "Repuestos enviados":   "Los repuestos fueron despachados. Una vez recibidos podrán retomar la reparación.",
    "En reparacion":        "El caso se encuentra en proceso de reparación.",
    "Finalizado":           "La reparación fue completada exitosamente. La orden queda cerrada en el sistema."
  };
  var msgs = esReseller ? msgsReseller : msgsTaller;
  // Estados exclusivos del flujo de batería — textos propios (solo si el equipo ES batería,
  // porque "Rechazado DJI" también existe en el circuito Reseller Propio y no debe usar esta redacción)
  if (esBateria(data.equipo)) {
    var msgsBateria = {
      "Caso Enviado":               "Enviamos su caso a DJI para la evaluación de la batería en garantía. Le avisaremos apenas tengamos la respuesta.",
      "Aprobado por DJI":           "DJI aprobó el caso: la batería quedó reconocida en garantía. Coordinaremos la reposición y le informaremos los próximos pasos.",
      "Rechazado DJI":              "DJI no aprobó el caso de la batería en garantía. Nos comunicaremos con usted para informarle las opciones disponibles.",
      "Bateria enviada a reseller": "Despachamos su batería de reemplazo. En breve la recibirá."
    };
    for (var kB in msgsBateria) { if (msgsBateria.hasOwnProperty(kB)) msgs[kB] = msgsBateria[kB]; }
  }
  var estimada = data._fechaEstimada ? filaDetalle("Fecha estimada de entrega", "<strong style='color:#00a3e0'>" + data._fechaEstimada + "</strong>") : "";
  var urgBanner = data.prioridad ? bloqueCard("⚡ Prioridad URGENTE", "Esta orden tiene prioridad máxima.", "#e74c3c") : "";
  var ficha =
    filaDetalle("Orden de trabajo", "<strong>" + data.ot + "</strong>") +
    filaDetalle("Equipo", data.equipo) +
    filaDetalle("Nº de Serie", data.sn||"—") +
    (data.cliente ? filaDetalle("Cliente", data.cliente) : "") +
    filaDetalle("Garantía", data.garantia==="IW"?"In Warranty (IW)":"Out of Warranty (OOW)") +
    (data.cas ? filaDetalle("Caso DJI", data.cas) : "") +
    filaDetalle("Estado anterior", ant||"—") +
    filaDetalle("Estado actual", "<strong style='color:#00a3e0'>" + nvo + "</strong>") +
    filaDetalle("Técnico asignado", (tec&&tec!=="Gestión Reseller") ? tec : "Equipo técnico BIDCOMAGRO") +
    estimada;
  // Informe técnico REAL (lo completa el técnico) — si todavía no existe (OTs viejas,
  // anteriores a que este campo se separara de "Falla reportada"), cae a data.trabajo
  // para no mandar el mail de cierre sin ningún contenido.
  var informeTxt = data.informeTecnico || data.trabajo;
  var informe = (nvo==="Finalizado" && informeTxt)
    ? "<div style='margin-top:20px;background:#f5f9fc;border-left:3px solid #00a3e0;border-radius:0 6px 6px 0;padding:14px 16px'><p style='font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 8px'>Informe técnico</p><p style='font-size:13px;color:#444;line-height:1.7;margin:0'>" + informeTxt + "</p></div>"
    : "";
  return construirEmailHTML(
    "Actualización de su Orden de Servicio", "Estimado/a " + data.reseller,
    urgBanner +
    "<p style='font-size:14px;color:#444;line-height:1.7;margin:0 0 22px'>" + (msgs[nvo]||"Su orden fue actualizada.") + "</p>" +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:4px'>" + ficha + "</div>" +
    construirTablaRepuestos(data.repuestos, nvo === "Finalizado") + informe,
    "Ante consultas comuníquese con su representante en BIDCOMAGRO."
  );
}


function armarEmailTecnico(data, estado, tec) {
  var msgs = {
    "Abierto":"Se te asignó una nueva OT. Ingresá al sistema para ver los detalles.",
    "Presupuesto aceptado":"El reseller aprobó el presupuesto. Podés iniciar la reparación.",
    "Repuestos enviados":"Los repuestos llegaron. Podés retomar la reparación."
  };
  return construirEmailHTML(
    "OT asignada: " + data.ot, "Hola, " + tec,
    "<p style='font-size:14px;color:#444;line-height:1.7;margin:0 0 22px'>" + (msgs[estado]||"Hay una actualización en tu OT.") + "</p>" +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px'>" +
      filaDetalle("OT", "<strong>" + data.ot + "</strong>") +
      filaDetalle("Reseller", data.reseller) +
      filaDetalle("Equipo", data.equipo) +
      filaDetalle("S/N", data.sn||"—") +
      (data.cliente ? filaDetalle("Cliente", data.cliente) : "") +
      filaDetalle("Estado", "<strong style='color:#00a3e0'>" + estado + "</strong>") +
      filaDetalle("Garantía", data.garantia) +
      (data.prioridad ? filaDetalle("Prioridad","<strong style='color:#e74c3c'>URGENTE</strong>") : "") +
    "</div>",
    "Ingresá a DJI HUB PRO para actualizar el estado."
  );
}


function armarEmailSupervisor(data, ant, nvo, tec, back) {
  var alertas = "";
  if (data.prioridad) alertas += bloqueCard("⚡ URGENTE","OT marcada como urgente.","#e74c3c");
  if (back)            alertas += bloqueCard(" Backorder","Repuestos pendientes de envío.","#e67e22");
  if (nvo==="Finalizado") alertas += bloqueCard("✓ Finalizada","Orden cerrada correctamente.","#27ae60");
  return construirEmailHTML(
    "Alerta — " + data.ot, "Supervisor",
    alertas +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-top:16px'>" +
      filaDetalle("OT","<strong>" + data.ot + "</strong>") +
      filaDetalle("Reseller", data.reseller) +
      filaDetalle("Equipo", data.equipo + (data.sn?" · S/N: "+data.sn:"")) +
      (data.cliente ? filaDetalle("Cliente", data.cliente) : "") +
      filaDetalle("Estado anterior", ant||"—") +
      filaDetalle("Estado actual","<strong style='color:#00a3e0'>" + nvo + "</strong>") +
      filaDetalle("Técnico", tec||"Sin asignar") +
      filaDetalle("Garantía", data.garantia) +
      (data.cas ? filaDetalle("CAS", data.cas) : "") +
      (data._fechaEstimada ? filaDetalle("Fecha estimada", "<strong>" + data._fechaEstimada + "</strong>") : "") +
    "</div>" + construirTablaRepuestos(data.repuestos),
    "Aviso automático por DJI HUB PRO."
  );
}


// Tabla de repuestos CON PRECIOS para la solicitud de facturación.
// Usa la misma lógica de precios que el presupuesto: PVP (sin imp.) desde
// DB_REPUESTOS × descuento reseller (40%). Devuelve { html, total }.
function construirTablaRepuestosFacturacion(rep) {
  if (!rep || rep === "Sin consumo de repuestos") return { html: "", total: 0 };

  // Mapa de precios PVP (col F / índice 5) por código (col B / índice 1) — igual que obtenerPresupuestoHTML
  var precioMap = {};
  var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
  if (hojaRep) {
    var dRep = getSheetValues(hojaRep);
    for (var pr = 1; pr < dRep.length; pr++) {
      var codR = String(dRep[pr][1] || "").trim().toUpperCase();
      if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5] || "0").replace(",", ".")) || 0;
    }
  }

  var DESC_RESELLER = 0.40; // precio reseller = PVP × (1 − 40%)
  var ls = rep.split(" ; "), filas = "", total = 0;
  for (var i = 0; i < ls.length; i++) {
    var p = ls[i].split(" | ");
    if (p.length < 3) continue;
    var cod = p[0].trim();
    var des = (p[1] || "").replace("(" + cod + ")", "").replace(/\(\s*\)/g, "").trim();
    var ped = parseInt((p[2].split(" E:")[0] || "").replace("P:", "")) || 0;
    var pvpBase = precioMap[cod.toUpperCase()] || 0;
    var pNeto = pvpBase > 0 ? Math.round(pvpBase * (1 - DESC_RESELLER) * 100) / 100 : 0;
    var sub = Math.round(pNeto * ped * 100) / 100;
    total += sub;
    filas += "<tr>" +
      "<td style='padding:7px 10px;font-size:11px;color:#00a3e0;font-weight:600;border:1px solid #e0e0e0'>" + cod + "</td>" +
      "<td style='padding:7px 10px;font-size:12px;color:#333;border:1px solid #e0e0e0'>" + des + "</td>" +
      "<td style='padding:7px 10px;font-size:12px;text-align:center;border:1px solid #e0e0e0'>" + ped + "</td>" +
      "<td style='padding:7px 10px;font-size:12px;text-align:right;border:1px solid #e0e0e0'>" + (pNeto ? "USD " + _fmtNum(pNeto) : "—") + "</td>" +
      "<td style='padding:7px 10px;font-size:12px;text-align:right;font-weight:700;border:1px solid #e0e0e0'>" + (sub ? "USD " + _fmtNum(sub) : "—") + "</td></tr>";
  }
  if (!filas) return { html: "", total: 0 };

  var html = "<div style='margin-top:20px'>" +
    "<p style='font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 8px'>Repuestos</p>" +
    "<table style='width:100%;border-collapse:collapse;font-size:12px'>" +
    "<thead><tr style='background:#f5f5f5'>" +
    "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
    "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
    "<th style='padding:7px 10px;text-align:center;border:1px solid #e0e0e0'>Cant.</th>" +
    "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>P.Unit neto (USD)</th>" +
    "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Subtotal (USD)</th>" +
    "</tr></thead><tbody>" + filas +
    (total > 0 ? "<tr style='background:#f5f5f5;font-weight:700'>" +
      "<td colspan='4' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL REPUESTOS</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(total) + "</td></tr>" : "") +
    "</tbody></table></div>";

  return { html: html, total: Math.round(total * 100) / 100 };
}
