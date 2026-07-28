// ============================================================
//  PORTAL RESELLER BIDCOM — Gestión de órdenes de trabajo
// @version 1.12
// ============================================================

function enviarCasoAlHub(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000);

    // Conexión directa para garantizar escritura atómica
    var ss   = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var hoja = ss.getSheetByName("Ordenes de trabajo");

    // Verificar duplicado S/N antes de crear.
    // Mismo criterio de "cerrada" que HUB PRO (_esCerrada / _ESTADOS_CERRADOS): una OT en
    // cualquiera de estos estados NO bloquea abrir un caso nuevo para el mismo equipo. Debe
    // quedar sincronizado con HUB_Código.js para que Portal y HUB decidan idéntico.
    var snNorm = String(data.sn || '').trim().toUpperCase();
    if (snNorm) {
      var ESTADOS_TERMINAL = {
        'Finalizado': true, 'Entregado': true, 'CANCELADO': true,
        'Rechazado DJI': true, 'Sin respuesta · Cerrado': true
      };
      var datosOT = hoja.getDataRange().getValues();
      for (var di = 1; di < datosOT.length; di++) {
        var snFila    = String(datosOT[di][SCHEMA.OT.SN]    || '').trim().toUpperCase();
        var estadoFila = String(datosOT[di][SCHEMA.OT.ESTADO] || '').trim();
        if (snFila !== snNorm) continue;
        if (!ESTADOS_TERMINAL[estadoFila]) {
          var otExistente = String(datosOT[di][SCHEMA.OT.OT] || '');
          lock.releaseLock();
          return { ok: false, error: 'Ya existe la OT ' + otExistente + ' abierta para ese S/N (' + snNorm + ')', otExistente: otExistente };
        }
      }
    }

    var num = String(hoja.getLastRow() + 1);
    while (num.length < 5) num = "0" + num;
    var nOT = "WH/REP/" + num;

    var fechaActual  = new Date();
    var repuestosStr = 'Sin consumo de repuestos';
    if (data.repuestos) {
      try {
        var reps = JSON.parse(data.repuestos);
        if (reps && reps.length) {
          repuestosStr = reps.map(function(r) {
            return r.sku + ' | ' + (r.descripcion || '') + ' | P:' + r.cantidad + ' E:0';
          }).join(' ; ');   // separador canónico del HUB (lo lee con split(' ; ')); NO usar '\n' o el Hub toma solo el primero
        }
      } catch(eRep) {}
    }

    var fallaFinal = data.falla || '';
    if (data.aftEstado) {
      var estadoLabel = data.aftEstado === 'repuesto'
        ? '[YA REPARADO — Repuesto para reposición de stock]'
        : '[PENDIENTE — Necesita repuestos para completar la reparación]';
      fallaFinal = estadoLabel + '\n' + fallaFinal;
    }

    var fechaStr = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    var msgBody = "Caso abierto desde el Portal Reseller.";
    if (data.driveUrl) msgBody += "\n\n📁 Drive: " + data.driveUrl;
    var mensajeInicial = "💬 [" + fechaStr + "] — " + data.reseller + ":\n" + msgBody + "\n\n[LEIDO]";

    var fila = [
      fechaActual, "", nOT, data.garantia, "Pendiente de Aprobación",
      data.equipo, data.sn, data.reseller, "", "", data.cliente || "", mensajeInicial,
      fallaFinal, "", data.cas || "", "", repuestosStr, "NORMAL", data.circuito, "", fechaActual
    ];

    if (data.fechaActivacion) {
      var partes = data.fechaActivacion.split("-");
      if (partes.length === 3) {
        fila[SCHEMA.OT.FECHA_ACTIVACION] = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
      }
    }

    hoja.appendRow(fila);
    SpreadsheetApp.flush();

    // Origen del repuesto → col AA (27) del HUB: precarga el badge "quién pone la pieza".
    //   aftEstado 'repuesto' = el reseller reparó con su stock (pide reposición) → "Stock reseller"
    //   cualquier otro       = necesita que le adelantemos el repuesto            → "Adelantado"
    if (data.aftEstado) {
      try {
        var origenVal = (data.aftEstado === 'repuesto') ? 'Stock reseller' : 'Adelantado';
        var nuevaFila = hoja.getLastRow();
        if (hoja.getMaxColumns() < 27) hoja.insertColumnsAfter(hoja.getMaxColumns(), 27 - hoja.getMaxColumns());
        hoja.getRange(nuevaFila, 27).setValue(origenVal);
      } catch(eOrig) { Logger.log('enviarCasoAlHub origen: ' + eOrig); }
    }

    lock.releaseLock();

    var cotizUrl = '';
    if (data.repuestos) {
      try {
        var repsArr = JSON.parse(data.repuestos);
        if (repsArr && repsArr.length) {
          var cotizResult = generarHojaCotizacion({ ot: nOT }, repsArr);
          if (cotizResult && cotizResult.ok) cotizUrl = cotizResult.url;
        }
      } catch(eCotiz) { Logger.log('enviarCasoAlHub cotiz: ' + eCotiz); }
    }

    _notificarNuevaOT(nOT, data, cotizUrl);
    return { nOT: nOT, cotizUrl: cotizUrl };
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    Logger.log("enviarCasoAlHub: " + e);
    return "Error: " + e.toString();
  }
}

function enviarLoteAlHub(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);

    // Conexión directa para garantizar escritura atómica
    var ss   = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var hoja = ss.getSheetByName("Ordenes de trabajo");

    var reseller       = data.reseller;
    var items          = data.items;
    var otsGeneradas   = [];
    var filasAInsertar = [];
    var fechaActual    = new Date();
    var lastRow        = hoja.getLastRow();
    var fechaStr       = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var num  = String(lastRow + 1 + i);
      while (num.length < 5) num = "0" + num;
      var nOT = "WH/REP/" + num;
      otsGeneradas.push(nOT);
      var msgLote = "💬 [" + fechaStr + "] — " + reseller + ":\n" + (item.falla || "Solicitud ingresada desde Portal") + "\n\n[LEIDO]";
      filasAInsertar.push([
        fechaActual, "", nOT, item.garantia, "Pendiente de Aprobación",
        item.equipo, item.sn, reseller, "", "", item.cliente || "", msgLote,
        item.falla, "", "", "", "Sin consumo de repuestos", "NORMAL", item.circuito, "", fechaActual
      ]);
    }

    if (filasAInsertar.length > 0) {
      hoja.getRange(lastRow + 1, 1, filasAInsertar.length, filasAInsertar[0].length).setValues(filasAInsertar);
    }

    SpreadsheetApp.flush();
    lock.releaseLock();

    _notificarLoteOT(reseller, items, otsGeneradas, ss);
    return { ok: true, ots: otsGeneradas };
  } catch(e) {
    if (lock.hasLock()) lock.releaseLock();
    Logger.log("enviarLoteAlHub: " + e);
    return { ok: false, error: e.toString() };
  }
}

function consultarEstado(ot, sn) {
  try {
    var ref  = _leerOrdenes();
    var otB  = String(ot).trim().toUpperCase();
    var snB  = String(sn).trim().toUpperCase();
    var mapaG = _mapaGarantiaEquipos();

    for (var i = 1; i < ref.datos.length; i++) {
      var f = ref.datos[i];
      if (!f[SCHEMA.OT.OT] || String(f[SCHEMA.OT.OT]).toUpperCase() !== otB) continue;
      if (!f[SCHEMA.OT.SN] || String(f[SCHEMA.OT.SN]).toUpperCase() !== snB) continue;

      var est      = String(f[SCHEMA.OT.ESTADO] || "").toUpperCase();
      var circRaw  = String(f[SCHEMA.OT.CIRCUITO] || "").trim().toUpperCase();
      var flujo    = "Taller";
      if (circRaw === "RESELLER" || circRaw === "SI") flujo = "Reseller";
      else if (circRaw === "RESELLER PROPIO") flujo = "Reseller Propio";

      var notaTec = String(f[SCHEMA.OT.TRABAJO] || "");
      var paso = 1, quePasa = "", accion = null;

      var DIR_CARMEN = "📍 EUCAWOODS — SIPCA, Carmen de Areco, BS.AS. (CP 6725)\n⏰ Lun a Vie 08:00-12:00 / 13:00-17:00\n👤 Mariano Castronuevo (+54 9 2325 65-6826)";

      // Normalizar nombres viejos para backward compat con filas existentes en el sheet
      var EST_ALIAS = {
        "REPARADO Y APROBADO EN EL AFTERSALES": "APROBADO DJI · REPUESTOS EN PREPARACIÓN",
        "PARTES DAÑADAS SCRAPEADAS":            "EN CIERRE",
        "RECEPCION PIEZAS DAÑADAS":             "PIEZAS DAÑADAS RECIBIDAS"
      };
      var estKey = EST_ALIAS[est] || est;

      // Detectar batería para usar el flujo unificado independientemente del circuito
      var esBateriaOT = (function() {
        try {
          var equipoNom = String(f[SCHEMA.OT.EQUIPO] || "").trim().toLowerCase();
          var hEq = getSheet(SCHEMA.SHEETS.EQUIPOS);
          if (!hEq) return false;
          var dEq = getSheetValues(hEq);
          for (var eb = 1; eb < dEq.length; eb++) {
            if (String(dEq[eb][0] || "").trim().toLowerCase() === equipoNom)
              return String(dEq[eb][1] || "").trim().toLowerCase() === "bateria";
          }
          return false;
        } catch(e) { return false; }
      })();

      if (estKey === "CANCELADO" || est.indexOf("CANCELADO") !== -1) {
        paso = 5; quePasa = "La orden fue anulada.";
      } else if (estKey === "FINALIZADO" || estKey === "ENTREGADO") {
        paso = 4; quePasa = "El caso está cerrado. ¡Gracias!";
      } else if (esBateriaOT) {
        if (estKey === "ABIERTO") {
          paso = 1; quePasa = "Registramos tu caso. Bidcom está preparando el expediente ante DJI.";
        } else if (estKey === "CASO ENVIADO") {
          paso = 2; quePasa = "Cargamos el caso en el portal DJI. Estamos esperando su respuesta.";
        } else if (estKey === "APROBADO POR DJI") {
          paso = 3; quePasa = "DJI aprobó el reemplazo de batería. Estamos coordinando el proceso para enviarte la batería de reposición.";
        } else if (estKey === "SCRAP ENVIADO (EVIDENCIAS)") {
          paso = 3; quePasa = "Enviamos el scrap a DJI como evidencia. Tu batería de reemplazo está siendo preparada por nuestro equipo de logística.";
        } else if (estKey === "BATERIA ENVIADA A RESELLER") {
          paso = 4; quePasa = "Tu batería de reemplazo fue despachada. Vas a recibirla en breve.";
        } else if (estKey === "RECHAZADO DJI") {
          paso = 5; quePasa = "DJI rechazó el caso de batería. Contactanos si tenés alguna consulta.";
        } else if (estKey === "SIN RESPUESTA · CERRADO") {
          paso = 5; quePasa = "El caso fue cerrado por inactividad. Si querés retomarlo, contactanos y lo reabrimos.";
        } else {
          paso = 2; quePasa = "El caso está siendo gestionado ante DJI. Te avisamos ante cualquier novedad.";
        }
      } else if (flujo === "Taller") {
        if (estKey === "ABIERTO" || estKey === "EN REVISION") {
          paso = 1; quePasa = "Tu solicitud fue registrada. Estamos revisando los datos del equipo para autorizar el ingreso.";
        } else if (estKey === "PRESUPUESTO ENVIADO") {
          paso = 2; quePasa = "Te enviamos un presupuesto para tu aprobación.";
          accion = "Revisá y aprobá el presupuesto desde este portal.";
        } else if (estKey === "PRESUPUESTO RECHAZADO") {
          paso = 5; quePasa = "El presupuesto fue rechazado. Contactanos si querés avanzar con una reparación fuera de garantía.";
        } else if (estKey === "PRESUPUESTO ACEPTADO") {
          paso = 2; quePasa = "Aceptaste el presupuesto. Coordiná el envío del equipo a nuestro taller.";
          accion = "Enviá el equipo a Carmen de Areco.\n" + DIR_CARMEN;
        } else if (estKey === "ESPERA DE REPUESTOS") {
          paso = 3; quePasa = "El equipo está en nuestro taller, esperando repuestos para completar la reparación.";
        } else if (estKey === "EN REPARACION") {
          paso = 3; quePasa = "El equipo está siendo reparado en nuestro taller.";
        } else {
          paso = 3; quePasa = "El caso está en proceso.";
        }
      } else if (flujo === "Reseller") {
        if (estKey === "ABIERTO") {
          paso = 1; quePasa = "Tu solicitud fue registrada. Estamos revisando el caso para autorizarte la reparación.";
        } else if (estKey === "EN REPARACION") {
          paso = 2; quePasa = "Tu caso fue aprobado. Ya podés ejecutar la reparación.";
          accion = "Realizá la reparación y envianos el informe técnico una vez finalizada.";
        } else if (estKey === "PEDIDO DE REPUESTOS") {
          paso = 2; quePasa = "Solicitaste repuestos. Los estamos preparando en almacén para enviártelos.";
        } else if (estKey === "REPUESTOS ENVIADOS") {
          paso = 3; quePasa = "Te despachamos los repuestos. En breve los vas a recibir.";
          accion = "Cuando los recibas, confirmá la recepción y realizá la reparación.";
        } else if (estKey === "INFORME DE REPARACION") {
          paso = 3; quePasa = "Recibimos tu informe de reparación. Lo estamos presentando ante DJI para su aprobación.";
          accion = "Envianos las piezas dañadas siguiendo las instrucciones de más abajo.";
        } else if (estKey === "PIEZAS DAÑADAS RECIBIDAS") {
          paso = 3; quePasa = "Recibimos las piezas dañadas. Estamos gestionando la aprobación ante DJI.";
        } else if (estKey === "APROBACION DJI") {
          paso = 4; quePasa = "Presentamos el caso ante DJI. Estamos esperando su resolución.";
        } else if (estKey === "RECHAZADO DJI") {
          paso = 5; quePasa = "DJI rechazó el caso. Los repuestos enviados quedan facturados a tu cuenta. Contactanos si tenés alguna consulta.";
        } else if (estKey === "SIN RESPUESTA · CERRADO") {
          paso = 5; quePasa = "El caso fue cerrado por inactividad. Si querés retomarlo, contactanos y lo reabrimos.";
        } else {
          paso = 3; quePasa = "El caso está en proceso.";
        }
      } else if (flujo === "Reseller Propio") {
        if (estKey === "ABIERTO") {
          paso = 1; quePasa = "Tu caso fue registrado y aprobado. Podés proceder con la reparación.";
          accion = "Realizá la reparación, cargá el caso en el portal DJI (AfterSales) y solicitá los repuestos de reposición desde aquí cuando tengas la aprobación.";
        } else if (estKey === "PEDIDO DE REPUESTO PARA REPARAR") {
          paso = 2; quePasa = "Solicitaste un adelanto de repuesto. Lo estamos preparando y te lo enviamos en breve.";
        } else if (estKey === "EN REPARACION") {
          paso = 2; quePasa = "Estás ejecutando la reparación. Una vez finalizada, cargá el caso en el portal DJI AfterSales y avisanos por la mensajería de la orden cuando tengas la aprobación.";
        } else if (estKey === "APROBADO DJI · REPUESTOS EN PREPARACIÓN") {
          paso = 3; quePasa = "Verificamos la aprobación de DJI. Tus repuestos de reposición están siendo preparados en almacén.";
        } else if (estKey === "EN CIERRE") {
          paso = 3; quePasa = "El caso está siendo procesado por nuestro equipo para el cierre final.";
        } else if (estKey === "RECHAZADO DJI") {
          paso = 5; quePasa = "DJI rechazó el caso. Los repuestos enviados quedan facturados a tu cuenta. Contactanos si tenés alguna consulta.";
        } else if (estKey === "SIN RESPUESTA · CERRADO") {
          paso = 5; quePasa = "El caso fue cerrado por inactividad. Si querés retomarlo, contactanos y lo reabrimos.";
        } else {
          paso = 3; quePasa = "El caso está en proceso.";
        }
      } else {
        paso = 3; quePasa = "El caso está en proceso. Te avisamos ante cualquier novedad.";
      }

      var inst = quePasa + (accion ? "\n\n" + accion : "") + (notaTec ? "\n\n📝 NOTA TÉCNICA:\n" + notaTec : "");

      var fechaAct     = f[SCHEMA.OT.FECHA_ACTIVACION];
      var equipo       = String(f[SCHEMA.OT.EQUIPO] || "").trim();
      var mesesGar     = mapaG[equipo.toLowerCase()] || 12;
      var infoGarantia = null;

      if (fechaAct instanceof Date) {
        var venc = new Date(fechaAct.getTime());
        venc.setMonth(venc.getMonth() + mesesGar);
        var hoy  = new Date();
        var dias = Math.floor((venc - hoy) / 86400000);
        infoGarantia = {
          fechaActivacion: Utilities.formatDate(fechaAct, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          vencimiento:     Utilities.formatDate(venc, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          diasParaVencer:  dias,
          vencida:         dias < 0,
          meses:           mesesGar,
          tieneFecha:      true
        };
      } else {
        infoGarantia = { tieneFecha: false };
      }

      var mensajesRaw = String(f[SCHEMA.OT.MENSAJES] || "").trim();
      return {
        encontrado:    true,
        ot:            String(f[SCHEMA.OT.OT]),
        cliente:       String(f[SCHEMA.OT.CLIENTE] || ""),
        equipo:        equipo,
        sn:            String(f[SCHEMA.OT.SN] || ""),
        estado:        String(f[SCHEMA.OT.ESTADO] || ""),
        garantia:      String(f[SCHEMA.OT.GARANTIA] || ""),
        cas:           String(f[SCHEMA.OT.CAS] || ""),
        trabajo:       inst,
        quePasa:       quePasa,
        accion:        accion,
        notaTec:       notaTec,
        pasoActual:    paso,
        flujo:         flujo,
        tecnico:       String(f[9] || ""),
        fechaEstimada: calcularFechaEstimada(flujo, String(f[3] || "OOW"), f[0]),
        infoGarantia:  infoGarantia,
        mensajes:      mensajesRaw,
        esBateria:     esBateriaOT
      };
    }
    return { encontrado: false };
  } catch(e) { Logger.log("consultarEstado: " + e); return { encontrado: false }; }
}

function buscarPorSN(sn) {
  try {
    var ref = _leerOrdenes();
    var snB = String(sn).trim().toUpperCase();
    var hoy = new Date();
    var out = [];
    for (var i = 1; i < ref.datos.length; i++) {
      var f   = ref.datos[i];
      var snH = String(f[SCHEMA.OT.SN] || "").trim().toUpperCase();
      if (!snH || snH !== snB) continue;
      var fechaFin = (f[SCHEMA.OT.FECHA_CIERRE] instanceof Date) ? f[SCHEMA.OT.FECHA_CIERRE] : hoy;
      var dias     = (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? Math.floor((fechaFin - f[SCHEMA.OT.FECHA_INGRESO]) / 86400000) : 0;
      out.push({
        ot:       String(f[SCHEMA.OT.OT]),
        equipo:   String(f[SCHEMA.OT.EQUIPO] || ""),
        sn:       String(f[SCHEMA.OT.SN] || ""),
        estado:   String(f[SCHEMA.OT.ESTADO] || ""),
        garantia: String(f[SCHEMA.OT.GARANTIA] || ""),
        reseller: String(f[SCHEMA.OT.RESELLER] || ""),
        tecnico:  String(f[SCHEMA.OT.TECNICO] || ""),
        circuito: String(f[SCHEMA.OT.CIRCUITO] || ""),
        dias:     dias
      });
    }
    return out;
  } catch(e) { return []; }
}

function buscarOTsReseller(nombreReseller, busqueda) {
  try {
    var ref = _leerOrdenes();
    var rB  = String(nombreReseller).trim().toLowerCase();
    var qB  = busqueda ? String(busqueda).trim().toLowerCase() : "";
    var hoy = new Date();
    var out = [];
    for (var i = 1; i < ref.datos.length; i++) {
      var f = ref.datos[i];
      if (!f[SCHEMA.OT.OT] || String(f[SCHEMA.OT.RESELLER] || "").trim().toLowerCase() !== rB) continue;
      if (qB) {
        var otStr      = String(f[SCHEMA.OT.OT]      || "").toLowerCase();
        var clienteStr = String(f[SCHEMA.OT.CLIENTE]  || "").toLowerCase();
        if (otStr.indexOf(qB) === -1 && clienteStr.indexOf(qB) === -1) continue;
      }
      var fechaFin     = (f[SCHEMA.OT.FECHA_CIERRE] instanceof Date) ? f[SCHEMA.OT.FECHA_CIERRE] : hoy;
      var dias         = (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? Math.floor((fechaFin - f[SCHEMA.OT.FECHA_INGRESO]) / 86400000) : 0;
      var tieneFechaAct = (f[SCHEMA.OT.FECHA_ACTIVACION] instanceof Date);
      var mensajes     = String(f[SCHEMA.OT.MENSAJES] || "").trim();
      var lastMsg      = mensajes.lastIndexOf("💬");
      var lastLeido    = mensajes.lastIndexOf("[LEIDO]");
      var msgNoLeido   = lastMsg !== -1 && (lastLeido === -1 || lastMsg > lastLeido);
      var esCerrada = (String(f[SCHEMA.OT.ESTADO]||"") === "Finalizado" ||
                       String(f[SCHEMA.OT.ESTADO]||"") === "Entregado"  ||
                       String(f[SCHEMA.OT.ESTADO]||"") === "CANCELADO");
      out.push({
        ot:           String(f[SCHEMA.OT.OT]),
        equipo:       String(f[SCHEMA.OT.EQUIPO] || ""),
        sn:           String(f[SCHEMA.OT.SN] || ""),
        estado:       String(f[SCHEMA.OT.ESTADO] || ""),
        garantia:     String(f[SCHEMA.OT.GARANTIA] || ""),
        tecnico:      String(f[SCHEMA.OT.TECNICO] || ""),
        cliente:      String(f[SCHEMA.OT.CLIENTE] || ""),
        circuito:     String(f[SCHEMA.OT.CIRCUITO] || ""),
        prioridad:    String(f[SCHEMA.OT.PRIORIDAD]).toUpperCase() === "URGENTE",
        tieneFechaAct: tieneFechaAct,
        msgNoLeido:   msgNoLeido,
        mensajes:     mensajes,
        dias:         dias,
        fechaIngreso: (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date)
          ? Utilities.formatDate(f[SCHEMA.OT.FECHA_INGRESO], Session.getScriptTimeZone(), "dd/MM/yyyy")
          : "",
        fechaCierre: (esCerrada && f[SCHEMA.OT.FECHA_CIERRE] instanceof Date)
          ? Utilities.formatDate(f[SCHEMA.OT.FECHA_CIERRE], Session.getScriptTimeZone(), "dd/MM/yyyy")
          : ""
      });
    }
    out.sort(function(a, b) {
      var af = (a.estado === "Finalizado" || a.estado === "Entregado" || a.estado === "CANCELADO") ? 1 : 0;
      var bf = (b.estado === "Finalizado" || b.estado === "Entregado" || b.estado === "CANCELADO") ? 1 : 0;
      return af !== bf ? af - bf : b.dias - a.dias;
    });
    return out;
  } catch(e) { return []; }
}

// Escritura única — rechaza si la fecha ya fue registrada
function agregarFechaActivacion(ot, fechaStr) {
  try {
    var ref = _leerOrdenes();
    var otB = String(ot).trim().toUpperCase();
    for (var i = 1; i < ref.datos.length; i++) {
      if (String(ref.datos[i][SCHEMA.OT.OT] || "").trim().toUpperCase() !== otB) continue;

      var actual = ref.datos[i][SCHEMA.OT.FECHA_ACTIVACION];
      if (actual instanceof Date || (actual && String(actual).trim() !== "")) {
        return { ok: false, msg: "La fecha de activación ya fue registrada y no puede modificarse." };
      }

      var partes = String(fechaStr).split("-");
      if (partes.length !== 3) return { ok: false, msg: "Formato de fecha inválido." };
      var fecha = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, parseInt(partes[2]));
      if (isNaN(fecha.getTime())) return { ok: false, msg: "Fecha inválida." };

      ref.hoja.getRange(i + 1, SCHEMA.OT.FECHA_ACTIVACION + 1).setValue(fecha);
      return { ok: true };
    }
    return { ok: false, msg: "OT no encontrada." };
  } catch(e) {
    Logger.log("agregarFechaActivacion: " + e);
    return { ok: false, msg: e.toString() };
  }
}

function obtenerGarantias(nombreReseller) {
  try {
    var ref    = _leerOrdenes();
    var rB     = String(nombreReseller).trim().toLowerCase();
    var mapaG  = _mapaGarantiaEquipos();
    var hoy    = new Date();
    var limite = new Date(hoy.getTime() + PORTAL_CONFIG.DIAS_AVISO_GARANTIA * 86400000);
    var out    = [];

    for (var i = 1; i < ref.datos.length; i++) {
      var f = ref.datos[i];
      if (!f[SCHEMA.OT.OT]) continue;
      if (String(f[SCHEMA.OT.RESELLER] || "").trim().toLowerCase() !== rB) continue;
      var estG = String(f[SCHEMA.OT.ESTADO] || "");
      if (estG === "Finalizado" || estG === "CANCELADO" || estG === "Entregado") continue;

      var fechaAct = f[SCHEMA.OT.FECHA_ACTIVACION];
      if (!(fechaAct instanceof Date)) continue;

      var equipo = String(f[SCHEMA.OT.EQUIPO] || "").trim();
      var meses  = mapaG[equipo.toLowerCase()] || 12;
      var venc   = new Date(fechaAct.getTime());
      venc.setMonth(venc.getMonth() + meses);

      var diasParaVencer = Math.floor((venc - hoy) / 86400000);
      if (diasParaVencer > PORTAL_CONFIG.DIAS_AVISO_GARANTIA) continue;

      out.push({
        ot:              String(f[SCHEMA.OT.OT]),
        equipo:          equipo,
        sn:              String(f[SCHEMA.OT.SN] || ""),
        estado:          String(f[SCHEMA.OT.ESTADO] || ""),
        fechaActivacion: Utilities.formatDate(fechaAct, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        vencimiento:     Utilities.formatDate(venc, Session.getScriptTimeZone(), "dd/MM/yyyy"),
        diasParaVencer:  diasParaVencer,
        vencida:         diasParaVencer < 0,
        mesesGarantia:   meses
      });
    }

    out.sort(function(a, b) { return a.diasParaVencer - b.diasParaVencer; });
    return out;
  } catch(e) { Logger.log("obtenerGarantias: " + e); return []; }
}

function solicitarRevisionTaller(ot, sn) {
  try {
    var ref = _leerOrdenes();
    var otB = String(ot).toUpperCase();
    var snB = String(sn).toUpperCase();
    for (var i = 1; i < ref.datos.length; i++) {
      if (!ref.datos[i][SCHEMA.OT.OT] || String(ref.datos[i][SCHEMA.OT.OT]).toUpperCase() !== otB) continue;
      if (!ref.datos[i][SCHEMA.OT.SN] || String(ref.datos[i][SCHEMA.OT.SN]).toUpperCase() !== snB) continue;
      if (String(ref.datos[i][SCHEMA.OT.PRIORIDAD]).toUpperCase() !== "URGENTE")
        ref.hoja.getRange(i + 1, SCHEMA.OT.PRIORIDAD + 1).setValue("ALERTA CLIENTE");
      return true;
    }
    return false;
  } catch(e) { return false; }
}

function agregarComentario(ot, comentario, autor) {
  try {
    var ref = _leerOrdenes();
    var otB = String(ot).trim().toUpperCase();

    for (var i = 1; i < ref.datos.length; i++) {
      if (String(ref.datos[i][SCHEMA.OT.OT] || "").trim().toUpperCase() !== otB) continue;

      var actual = String(ref.datos[i][SCHEMA.OT.MENSAJES] || "").replace(/\n\n\[LEIDO\]/g, "").replace(/\[LEIDO\]/g, "").trim();
      var fecha  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      var bloque = (actual ? actual + "\n\n" : "") +
                   "💬 [" + fecha + "] — " + String(autor || "Portal") + ":\n" + comentario;

      ref.hoja.getRange(i + 1, 12).setValue(bloque);

      if (String(ref.datos[i][SCHEMA.OT.PRIORIDAD]).toUpperCase() !== "URGENTE")
        ref.hoja.getRange(i + 1, SCHEMA.OT.PRIORIDAD + 1).setValue("COMENTARIO RESELLER");

      try {
        var asunto = "[PORTAL] Comentario reseller — " + ot;
        var htmlComent = _construirEmailHTML(
          "Comentario del reseller — " + ot, "Supervisor",
          "<div style='background:#fff8e6;border-left:3px solid #f39c12;border-radius:0 6px 6px 0;padding:12px 16px;margin-bottom:16px'>" +
            "<p style='font-size:12px;font-weight:700;color:#e67e22;margin:0 0 6px'>💬 " + (autor || "Reseller") + "</p>" +
            "<p style='font-size:13px;color:#333;margin:0'>" + comentario + "</p>" +
          "</div>" +
          "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px'>" +
            _filaDetalle("Orden", "<strong>" + ot + "</strong>") +
            _filaDetalle("Estado actual", String(ref.datos[i][4] || "—")) +
          "</div>",
          "El reseller dejó un comentario Revisá el HUB PRO."
        );
        var supThreadId = _obtenerThreadId(ot, PORTAL_CONFIG.EMAIL_SUPERVISOR);
        if (supThreadId) {
          try {
            var hilo = GmailApp.getThreadById(supThreadId);
            if (!hilo) throw new Error('Thread no encontrado');
            hilo.replyAll('', { htmlBody: htmlComent, name: PORTAL_CONFIG.NOMBRE_REMITENTE });
          } catch(eRep) {
            Logger.log('agregarComentario replyAll: ' + eRep + ' — fallback');
            GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, "", { htmlBody: htmlComent, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR });
          }
        } else {
          GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, "", { htmlBody: htmlComent, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR });
        }
      } catch(e2) { Logger.log("Email comentario: " + e2); }

      return { ok: true };
    }
    return { ok: false, msg: "OT no encontrada" };
  } catch(e) { Logger.log("agregarComentario: " + e); return { ok: false, msg: e.toString() }; }
}

// Aprobación o rechazo de presupuesto desde el Portal, con registro firmado en col L
function aprobarPresupuestoPortal(ot, decision, observaciones) {
  try {
    var ref   = _leerOrdenes();
    var otB   = String(ot).trim().toUpperCase();
    var dec   = String(decision||"").toLowerCase();
    if (dec !== "aceptado" && dec !== "rechazado") return { ok: false, msg: "Decisión inválida." };

    for (var i = 1; i < ref.datos.length; i++) {
      var f = ref.datos[i];
      if (String(f[SCHEMA.OT.OT]||"").trim().toUpperCase() !== otB) continue;

      var estadoActual = String(f[SCHEMA.OT.ESTADO]||"");
      if (estadoActual !== "Presupuesto enviado") {
        return { ok: false, msg: "La OT no está en estado 'Presupuesto enviado' (actual: " + estadoActual + ")." };
      }

      var nuevoEstado = dec === "aceptado" ? "Presupuesto aceptado" : "Presupuesto rechazado";
      var fecha       = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
      var firma       = "✅ PRESUPUESTO " + dec.toUpperCase() + " VÍA PORTAL · " + fecha;
      if (observaciones) firma += " — " + String(observaciones).trim();

      // Registrar en MENSAJES (col L)
      var mensajesActual = String(f[SCHEMA.OT.MENSAJES]||"").trim();
      var bloque = (mensajesActual ? mensajesActual + "\n\n" : "") +
                   "💬 [" + fecha + "] — Portal Reseller:\n" + firma;
      ref.hoja.getRange(i + 1, SCHEMA.OT.MENSAJES + 1).setValue(bloque);

      // Cambiar estado en col E
      ref.hoja.getRange(i + 1, SCHEMA.OT.ESTADO + 1).setValue(nuevoEstado);
      ref.hoja.getRange(i + 1, SCHEMA.OT.FECHA_ESTADO + 1).setValue(new Date());

      // Notificar al supervisor
      try {
        var asunto = "[PORTAL] Presupuesto " + dec + " — " + otB + " · " + String(f[SCHEMA.OT.EQUIPO]||"");
        var html = _construirEmailHTML(
          "Presupuesto " + dec + " por el reseller", "Supervisor",
          "<div style='background:" + (dec==="aceptado"?"#eafaf1":"#fdf2f2") + ";border-left:3px solid " + (dec==="aceptado"?"#1a9e4a":"#e74c3c") + ";border-radius:0 6px 6px 0;padding:12px 16px;margin-bottom:16px'>" +
            "<p style='font-size:13px;font-weight:700;color:" + (dec==="aceptado"?"#1a9e4a":"#e74c3c") + ";margin:0 0 6px'>" +
            (dec==="aceptado"?"✅ Aprobado":"❌ Rechazado") + " desde el Portal</p>" +
            (observaciones?"<p style='font-size:12px;color:#555;margin:0'>\"" + observaciones + "\"</p>":"") +
          "</div>" +
          "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px'>" +
            _filaDetalle("OT", "<strong>" + otB + "</strong>") +
            _filaDetalle("Reseller", String(f[SCHEMA.OT.RESELLER]||"")) +
            _filaDetalle("Equipo", String(f[SCHEMA.OT.EQUIPO]||"")) +
            _filaDetalle("Nuevo estado", "<strong>" + nuevoEstado + "</strong>") +
            _filaDetalle("Fecha decisión", fecha) +
          "</div>",
          "El estado fue actualizado automáticamente en el sistema."
        );
        var supThreadId = _obtenerThreadId(otB, PORTAL_CONFIG.EMAIL_SUPERVISOR);
        if (supThreadId) {
          try {
            var hilo = GmailApp.getThreadById(supThreadId);
            if (!hilo) throw new Error('Thread no encontrado');
            hilo.replyAll('', { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE });
            _logEmail(otB, PORTAL_CONFIG.EMAIL_SUPERVISOR, "Supervisor", asunto, "OK", supThreadId);
          } catch(eRep) {
            Logger.log('aprobarPresupuestoPortal replyAll: ' + eRep + ' — fallback');
            GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, "", { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR });
            _logEmail(otB, PORTAL_CONFIG.EMAIL_SUPERVISOR, "Supervisor", asunto, "OK", "");
          }
        } else {
          GmailApp.sendEmail(PORTAL_CONFIG.EMAIL_SUPERVISOR, asunto, "", { htmlBody: html, name: PORTAL_CONFIG.NOMBRE_REMITENTE, replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR });
          _logEmail(otB, PORTAL_CONFIG.EMAIL_SUPERVISOR, "Supervisor", asunto, "OK", "");
        }
      } catch(e2) { Logger.log("aprobarPresupuestoPortal email: " + e2); }

      invalidateSheetValues(SCHEMA.SHEETS.OT);
      return { ok: true, nuevoEstado: nuevoEstado };
    }
    return { ok: false, msg: "OT no encontrada." };
  } catch(e) { Logger.log("aprobarPresupuestoPortal: " + e); return { ok: false, msg: e.toString() }; }
}

function confirmarRecepcionRepuestos(ot, sn) {
  try {
    var ref = _leerOrdenes();
    var otB = String(ot).trim().toUpperCase();
    var snB = String(sn).trim().toUpperCase();
    for (var i = 1; i < ref.datos.length; i++) {
      var f = ref.datos[i];
      if (String(f[SCHEMA.OT.OT] || "").trim().toUpperCase() !== otB) continue;
      if (String(f[SCHEMA.OT.SN] || "").trim().toUpperCase() !== snB) continue;
      var estadoAct = String(f[SCHEMA.OT.ESTADO] || "").trim();
      if (estadoAct !== "Repuestos enviados") {
        return { ok: false, msg: "El estado actual no es 'Repuestos enviados'." };
      }
      ref.hoja.getRange(i + 1, 5).setValue("En reparacion");
      invalidateSheetValues(SCHEMA.SHEETS.OT);   // sin esto el cache (getSheetValues, TTL 60s) sigue devolviendo "Repuestos enviados" y el botón reaparece por ~1 min
      return { ok: true };
    }
    return { ok: false, msg: "OT no encontrada." };
  } catch(e) {
    Logger.log("confirmarRecepcionRepuestos: " + e);
    return { ok: false, msg: e.toString() };
  }
}

// Polling liviano: devuelve sólo el snapshot mínimo de cada OT del reseller
// para detectar cambios de estado o mensajes nuevos sin cargar todo el detalle.
// snapshotAnterior: objeto { ot: {estado, lastMsg} } serializado desde el cliente.
function obtenerSnapshotOTs(nombreReseller, snapshotJson) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.OT);
    var rB    = String(nombreReseller).trim().toLowerCase();
    var prev  = {};
    try { if (snapshotJson) prev = JSON.parse(snapshotJson); } catch(e2) {}

    var cambios = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      var ot = String(f[SCHEMA.OT.OT]||"").trim();
      if (!ot) continue;
      if (String(f[SCHEMA.OT.RESELLER]||"").trim().toLowerCase() !== rB) continue;

      var estado   = String(f[SCHEMA.OT.ESTADO]||"");
      var mensajes = String(f[SCHEMA.OT.MENSAJES]||"").trim();
      var lastMsg  = mensajes.lastIndexOf("💬");
      var lastLeido= mensajes.lastIndexOf("[LEIDO]");
      var msgNuevo = lastMsg !== -1 && (lastLeido === -1 || lastMsg > lastLeido);

      var p = prev[ot] || {};
      if (p.estado !== estado || (msgNuevo && !p.msgVisto)) {
        cambios.push({
          ot:       ot,
          equipo:   String(f[SCHEMA.OT.EQUIPO]||""),
          estado:   estado,
          msgNuevo: msgNuevo,
          estadoCambio: p.estado && p.estado !== estado
        });
      }
    }
    return cambios;
  } catch(e) { Logger.log("obtenerSnapshotOTs: " + e); return []; }
}

function guardarClienteOT(ot, cliente) {
  try {
    var sheet = getSheet(SCHEMA.SHEETS.OT);
    var datos = getSheetValues(SCHEMA.SHEETS.OT);
    var otB   = String(ot).trim().toUpperCase();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][SCHEMA.OT.OT] || '').toUpperCase() !== otB) continue;
      sheet.getRange(i + 1, SCHEMA.OT.CLIENTE + 1).setValue(String(cliente || '').trim());
      invalidateSheetValues(SCHEMA.SHEETS.OT);
      return { ok: true };
    }
    return { ok: false, msg: 'OT no encontrada' };
  } catch(e) {
    Logger.log('guardarClienteOT: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

// ── GRUPO: obtener OTs de múltiples resellers ─────────────────
// Retorna el mismo formato que buscarOTsReseller pero agrega campo `sucursal`.
function RS_buscarOTsGrupo(resellers) {
  try {
    var ref = _leerOrdenes();
    var hoy = new Date();
    var out = [];

    // Construir mapa nombre → nombre normalizado para lookup rápido
    var mapaResellers = {};
    for (var r = 0; r < resellers.length; r++) {
      mapaResellers[String(resellers[r]).trim().toLowerCase()] = String(resellers[r]).trim();
    }

    for (var i = 1; i < ref.datos.length; i++) {
      var f           = ref.datos[i];
      var resellerFila = String(f[SCHEMA.OT.RESELLER] || '').trim();
      var resellerKey  = resellerFila.toLowerCase();
      if (!f[SCHEMA.OT.OT] || !mapaResellers[resellerKey]) continue;

      var fechaFin     = (f[SCHEMA.OT.FECHA_CIERRE] instanceof Date) ? f[SCHEMA.OT.FECHA_CIERRE] : hoy;
      var dias         = (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? Math.floor((fechaFin - f[SCHEMA.OT.FECHA_INGRESO]) / 86400000) : 0;
      var tieneFechaAct = (f[SCHEMA.OT.FECHA_ACTIVACION] instanceof Date);
      var mensajes     = String(f[SCHEMA.OT.MENSAJES] || '').trim();
      var lastMsg      = mensajes.lastIndexOf('💬');
      var lastLeido    = mensajes.lastIndexOf('[LEIDO]');
      var msgNoLeido   = lastMsg !== -1 && (lastLeido === -1 || lastMsg > lastLeido);
      var esCerrada    = (String(f[SCHEMA.OT.ESTADO]||'') === 'Finalizado' ||
                          String(f[SCHEMA.OT.ESTADO]||'') === 'Entregado'  ||
                          String(f[SCHEMA.OT.ESTADO]||'') === 'CANCELADO');
      out.push({
        ot:           String(f[SCHEMA.OT.OT]),
        equipo:       String(f[SCHEMA.OT.EQUIPO] || ''),
        sn:           String(f[SCHEMA.OT.SN] || ''),
        estado:       String(f[SCHEMA.OT.ESTADO] || ''),
        garantia:     String(f[SCHEMA.OT.GARANTIA] || ''),
        tecnico:      String(f[SCHEMA.OT.TECNICO] || ''),
        cliente:      String(f[SCHEMA.OT.CLIENTE] || ''),
        circuito:     String(f[SCHEMA.OT.CIRCUITO] || ''),
        prioridad:    String(f[SCHEMA.OT.PRIORIDAD]).toUpperCase() === 'URGENTE',
        tieneFechaAct: tieneFechaAct,
        msgNoLeido:   msgNoLeido,
        mensajes:     mensajes,
        dias:         dias,
        sucursal:     mapaResellers[resellerKey],
        fechaIngreso: (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date)
          ? Utilities.formatDate(f[SCHEMA.OT.FECHA_INGRESO], Session.getScriptTimeZone(), 'dd/MM/yyyy')
          : '',
        fechaCierre: (esCerrada && f[SCHEMA.OT.FECHA_CIERRE] instanceof Date)
          ? Utilities.formatDate(f[SCHEMA.OT.FECHA_CIERRE], Session.getScriptTimeZone(), 'dd/MM/yyyy')
          : ''
      });
    }

    out.sort(function(a, b) {
      var af = (a.estado === 'Finalizado' || a.estado === 'Entregado' || a.estado === 'CANCELADO') ? 1 : 0;
      var bf = (b.estado === 'Finalizado' || b.estado === 'Entregado' || b.estado === 'CANCELADO') ? 1 : 0;
      if (af !== bf) return af - bf;
      if (a.sucursal !== b.sucursal) return a.sucursal < b.sucursal ? -1 : 1;
      return b.dias - a.dias;
    });
    return out;
  } catch(e) { Logger.log('RS_buscarOTsGrupo: ' + e); return []; }
}

function calcularFechaEstimada(circuito, garantia, fechaApertura) {
  var key  = circuito + "-" + garantia;
  var dias = PORTAL_CONFIG.DIAS_ESTIMADOS[key] || 10;
  var base = (fechaApertura instanceof Date) ? fechaApertura : new Date();
  var est  = new Date(base.getTime() + dias * 86400000);
  while (est.getDay() === 0 || est.getDay() === 6) est.setDate(est.getDate() + 1);
  return Utilities.formatDate(est, Session.getScriptTimeZone(), "dd/MM/yyyy");
}
