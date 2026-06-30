// ============================================================
//  DJI HUB PRO v14.1 — Codigo.gs
//  Proyecto: DJI HUB PRO
//  Sheet ID: el spreadsheet activo (SS)
// @version 2.6
//
//  Funciones exclusivas del HUB interno:
//  cargarTodo, actualizarOrden, crearNuevaOT,
//  obtenerLogs, obtenerEmailLogs, obtenerMetricasTecnicos,
//  obtenerPendientesEnvio, obtenerInfoCliente,
//  sistema completo de emails + triggers SLA + mensual
// ============================================================

// Formato numérico regional AR: punto miles, coma decimales (ej. 4.150,00)
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

// Token HMAC-SHA256 para aprobación de presupuesto 1-click.
// El mismo secreto debe estar en las Script Properties de HUB y Portal con clave APPROVAL_SECRET.
function _tokenAprobacion(ot, action) {
  var secret = PropertiesService.getScriptProperties().getProperty('APPROVAL_SECRET') || 'bidcomagro-default';
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(ot) + '|' + String(action) + '|' + secret
  );
  return bytes.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 40);
}

var WOS_SS_ID    = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
var WOS_HOJA_PED = 'Pedidos_OTs';

function HUB_generarPedidoRepuestos(data) {
  try {
    var hoja  = SpreadsheetApp.openById(WOS_SS_ID).getSheetByName(WOS_HOJA_PED);
    if (!hoja) return { ok: false, error: 'Hoja Pedidos_OTs no encontrada en el WOS.' };
    var todos = hoja.getDataRange().getValues();
    var numero = 'OT-' + String(data.ot || '').trim();

    // Recopilar SKUs ya pedidos y threadId existente para este NUMERO
    var existingSKUs = {}, existingThreadId = '';
    for (var ci = 1; ci < todos.length; ci++) {
      if (String(todos[ci][0] || '').trim() !== numero) continue;
      var eSku = String(todos[ci][2] || '').trim().toUpperCase();
      if (eSku) existingSKUs[eSku] = true;
      if (!existingThreadId) existingThreadId = String(todos[ci][17] || '').trim();
    }

    var items = data.items || [];
    if (!items.length) return { ok: false, error: 'No hay ítems para pedir.' };

    // Solo agregar SKUs que no estén ya en el pedido
    var itemsNuevos = [];
    for (var fi = 0; fi < items.length; fi++) {
      var fSku = String(items[fi].cod || '').trim().toUpperCase();
      if (!existingSKUs[fSku]) itemsNuevos.push(items[fi]);
    }
    if (!itemsNuevos.length) return { ok: false, error: 'Todos los ítems ya están en el pedido ' + numero + '.' };
    items = itemsNuevos;

    var reseller  = String(data.reseller  || '').trim();
    var idVenGar  = String(data.cas || '').trim();
    var envio     = String(data.envio     || 'Retiro').trim();
    var circuito  = String(data.circuito  || 'Taller').trim();
    var esTaller  = !circuito || circuito === 'Taller' || circuito.toLowerCase() === 'taller';
    var tecnico   = String(data.tecnico   || '').trim();
    var aprobado  = data.aprobadoDJI ? 'DJI ✓ Aprobado' : '⚠ SIN APROBACIÓN DJI';
    var obs = aprobado +
      (data.garantia ? ' | ' + data.garantia : '') +
      (idVenGar ? ' | CAS: ' + idVenGar : '') +
      (esTaller ? ' | Taller · entrega técnico' + (tecnico ? ' ' + tecnico : '') : '');
    var fecha    = new Date();
    var operario = Session.getActiveUser().getEmail();

    // Abrir MASTER una sola vez para precios y email del reseller
    var masterSS = null;
    try { masterSS = SpreadsheetApp.openById(MASTER_SHEET_ID); } catch(eMSS) { Logger.log('MASTER open: ' + eMSS); }

    // Mapa de precios desde Lista_Repuestos (precio reseller = lista × 0.60)
    var priceMap = {};
    if (masterSS) try {
      var listaData = masterSS.getSheetByName('Lista_Repuestos').getDataRange().getValues();
      for (var li = 1; li < listaData.length; li++) {
        var lSku = String(listaData[li][0] || '').trim().toUpperCase();
        var lPre = Number(listaData[li][4]) || 0;
        if (lSku && lPre > 0) priceMap[lSku] = Math.round(lPre * 0.60 * 100) / 100;
      }
    } catch(ePM) { Logger.log('HUB_generarPedidoRepuestos priceMap: ' + ePM); }

    // Buscar email del reseller en MASTER
    var emailReseller = '';
    if (masterSS) try {
      var rsData = masterSS.getSheetByName(SCHEMA.SHEETS.RESELLERS).getDataRange().getValues();
      var rNom = reseller.toLowerCase();
      for (var ri = 1; ri < rsData.length; ri++) {
        if (String(rsData[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rNom) {
          emailReseller = String(rsData[ri][SCHEMA.RESELLERS.EMAIL] || '').trim();
          break;
        }
      }
    } catch(eRS) { Logger.log('HUB_generarPedidoRepuestos emailReseller: ' + eRS); }

    // Notificar al reseller — continúa en el hilo de la OT si ya existía, nuevo si es el primer pedido
    var threadId = existingThreadId;
    try {
      var ot     = String(data.ot || numero).trim();
      var asunto = '[' + numero + '] Repuestos — ' + reseller;
      var itemsText = '';
      for (var ii = 0; ii < items.length; ii++) {
        var itI = items[ii];
        itemsText += '• ' + String(itI.cod || '').trim() + ' · ' + String(itI.desc || '').trim() + ' (x' + (Number(itI.qty) || 1) + ')\n';
      }
      var bodyHtml = existingThreadId
        ? '<p>Ítems adicionales agregados a <strong>' + numero + '</strong>:</p>' +
          '<pre style="font-size:12px;background:#f5f5f5;padding:10px;border-radius:6px">' + itemsText + '</pre>'
        : '<p><strong>' + numero + '</strong> — Solicitud de repuestos de reparación</p>' +
          '<p>Reseller: <strong>' + reseller + '</strong>' + (idVenGar ? ' · Ref: <em>' + idVenGar + '</em>' : '') + '</p>' +
          '<pre style="font-size:12px;background:#f5f5f5;padding:10px;border-radius:6px">' + itemsText + '</pre>';
      var toAddr = emailReseller || CONFIG.EMAIL_SUPERVISOR;
      threadId = _enviarConHilo(ot, toAddr, asunto, bodyHtml) || threadId;
      if (threadId && !existingThreadId) registrarEmailLog(ot, toAddr, 'Repuestos', asunto, 'OK', threadId);
    } catch(eGm) { Logger.log('HUB_generarPedidoRepuestos Gmail: ' + eGm); }

    // Escribir filas con el schema de 26 cols (igual a Pedidos_resellers / COL)
    var primeraFila = hoja.getLastRow() + 1;
    for (var j = 0; j < items.length; j++) {
      var it  = items[j];
      if (!it.cod || !it.desc) continue;
      var qty    = Number(it.qty) || 1;
      var skuUp  = String(it.cod || '').trim().toUpperCase();
      var precio = priceMap[skuUp] || Number(it.precio || it.price || 0);
      // 26 posiciones: índice 0=A … 25=Z
      var fila = ['','','','','','','','','','','','','','','','','','','','','','','','','',''];
      fila[0]  = numero;                                   // A NUMERO
      fila[1]  = reseller;                                 // B RESELLER
      fila[2]  = String(it.cod  || '').trim().toUpperCase(); // C SKU
      fila[3]  = String(it.desc || '').trim();             // D DESC
      fila[4]  = qty;                                      // E CANT_SOL
      fila[5]  = 0;                                        // F CANT_DESP
      // G(6) CANT_PEND: se pone fórmula después (appendRow no soporta fórmulas)
      fila[7]  = precio;                                   // H PRECIO
      fila[8]  = '';                                       // I STOCK_ORI (desconocido al crear)
      fila[9]  = 'Confirmado';                             // J ESTADO
      fila[10] = fecha;                                    // K FECHA
      fila[11] = envio;                                     // L ENVIO (Envío / Retiro)
      fila[12] = '';                                       // M PAGO
      fila[13] = obs;                                      // N OBS
      fila[17] = threadId;                                 // R THREAD_ID
      fila[23] = operario;                                 // X OPERARIO
      fila[25] = 0;                                        // Z CANT_CANCEL
      hoja.appendRow(fila);
    }

    // Escribir fórmula CANT_PEND (=E-F-Z) en col G para cada fila recién agregada
    var ultimaFila = hoja.getLastRow();
    for (var f = primeraFila; f <= ultimaFila; f++) {
      hoja.getRange(f, 7).setFormula('=E' + f + '-F' + f + '-Z' + f);
    }

    // Backfill threadId en filas preexistentes que no lo tenían
    if (threadId && !existingThreadId) {
      for (var bi = 1; bi < todos.length; bi++) {
        if (String(todos[bi][0] || '').trim() === numero && !String(todos[bi][17] || '').trim()) {
          hoja.getRange(bi + 1, 18).setValue(threadId);
        }
      }
    }

    SpreadsheetApp.flush();

    registrarLog(data.ot, '', operario,
      'PEDIDO REP', data.estado || '', data.estado || '',
      'Generado ' + numero + ' — ' + items.length + ' ítem(s) nuevo(s)' + (threadId ? ' [hilo Gmail OK]' : ' [SIN hilo Gmail]'));

    return { ok: true, numero: numero, threadId: threadId };
  } catch(e) {
    Logger.log('HUB_generarPedidoRepuestos ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

var _SS = null;
function SS() {
  if (!_SS) _SS = SpreadsheetApp.getActiveSpreadsheet();
  return _SS;
}
function getSheet(nombre) { return SS().getSheetByName(nombre); }

//  CONFIGURACIÓN — EDITÁ SOLO ESTA SECCIÓN

var CONFIG = {
  EMAIL_SUPERVISOR:   "soporteagrasdji@bidcom.com.ar",
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
    "Caso Enviado","Bateria enviada a reseller"
  ],
  ESTADOS_NOTIFICAR_TECNICO:    ["Abierto","Presupuesto aceptado","Repuestos enviados"],
  ESTADOS_NOTIFICAR_SUPERVISOR: ["Finalizado"],
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
//  CARGA UNIFICADA — una sola llamada desde el frontend
//  Devuelve: ordenes + repuestos + tecnicos + resellers
// ============================================================
function cargarTodo() {
  try {
    var hoy     = new Date();
    var hojaOT  = getSheet(SCHEMA.SHEETS.OT);
    var datosOT = getSheetValues(hojaOT);

    // Mapa de equipos tipo Bateria + meses de garantía desde hoja EQUIPOS
    var mapaBaterias = {}, mesesMap = {};
    var hojaEqBat = getSheet(SCHEMA.SHEETS.EQUIPOS);
    if (hojaEqBat) {
      var dEqBat = getSheetValues(hojaEqBat);
      for (var eb = 1; eb < dEqBat.length; eb++) {
        var nomEq = String(dEqBat[eb][0]||"").trim();
        if (!nomEq) continue;
        var nomLow = nomEq.toLowerCase();
        if (String(dEqBat[eb][1]||"").trim().toLowerCase() === "bateria") {
          mapaBaterias[nomLow] = true;
        }
        var mes = parseInt(dEqBat[eb][SCHEMA.EQUIPOS.MESES], 10);
        if (!isNaN(mes) && mes > 0) mesesMap[nomLow] = mes;
      }
    }
    Logger.log("Baterías mapeadas: " + Object.keys(mapaBaterias).length);

    var ordenes = [], tecSet = {};
    for (var i = 1; i < datosOT.length; i++) {
      var f = datosOT[i];
      if (!f[2] || String(f[2]).trim() === "") continue;

      var otStr = String(f[2]).trim();

      // 1. Días Totales (col A → col B si finalizado, sino hoy)
      var dias = (f[0] instanceof Date) ? Math.floor((hoy - f[0]) / 86400000) : 0;

      // 2. Días en Estado (col U = sello de último cambio de estado)
      var diasEstado = dias;
      var selloEstado = f[20];
      if (selloEstado instanceof Date) {
        var dEst = Math.floor((hoy - selloEstado) / 86400000);
        if (!isNaN(dEst)) diasEstado = dEst;
      }

      var circ = String(f[18] || "").trim();
      var circUp = circ.toUpperCase();
      var tipo = "Taller";
      if (circUp === "SI" || circUp === "RESELLER") tipo = "Reseller";
      else if (circUp === "RESELLER PROPIO")        tipo = "Reseller Propio";
      
      var tec = String(f[9] || "").trim();
      if (tec && tec !== "Gestión Reseller") tecSet[tec] = 1;

      var mensajesRaw  = String(f[11]||"").trim();
      var lastMsg      = mensajesRaw.lastIndexOf("💬");
      var lastLeido    = mensajesRaw.lastIndexOf("[LEIDO]");
      var msgNoLeido   = lastMsg !== -1 && (lastLeido === -1 || lastMsg > lastLeido);

      var rawUM = f[SCHEMA.OT.ULTIMA_MODIFICACION];
      ordenes.push({
        fila:      i+1,
        ot:        otStr,
        garantia:  String(f[3]||"OOW"),
        reseller:  String(f[7]||"Particular"),
        cliente:   String(f[SCHEMA.OT.CLIENTE]||"").trim(),
        estado:    String(f[4]||"Abierto"),
        equipo:    String(f[5]||"S/D"),
        sn:        String(f[6]||""),
        tecnico:    (f[9] && f[9] !== 'undefined') ? String(f[9]) : "",
        trabajo:   String(f[12]||""),
        factura:   (f[13] instanceof Date) ? Utilities.formatDate(f[13], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(f[13] || ""),
        vencimientoGar: (function() {
          var eqKey = String(f[SCHEMA.OT.EQUIPO]||'').trim().toLowerCase();
          var mes   = mesesMap[eqKey] || 0;
          if (mes > 0 && f[13] instanceof Date) {
            var vd = new Date(f[13].getTime());
            vd.setMonth(vd.getMonth() + mes);
            return Utilities.formatDate(vd, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          return '';
        })(),
        cas:       String(f[14]||""),
        repuestos: f[16] ? String(f[16]).replace(/\r/g, "").trim() : "Sin consumo de repuestos",
        manoObraGuardada: f[22] ? String(f[22]).trim() : "",
        notasInternas:   f[23] ? String(f[23]).trim() : "",
        mensajes:   String(f[11]||"").trim(),
        msgNoLeido: msgNoLeido,
        fechaIngreso: (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? f[SCHEMA.OT.FECHA_INGRESO].getTime() : null,
        prioridad: String(f[17]).toUpperCase() === "URGENTE",
        circuito:  tipo,
        esBateria: mapaBaterias[String(f[5]||'').trim().toLowerCase()] === true,
        dias:      dias,
        diasEstado: diasEstado,
        umMod:     (rawUM instanceof Date) ? rawUM.getTime() : 0
      });
    }

    var tecnicosDisp = obtenerTecnicosDisponibles();
    var catalogo = cargarCatalogo();
    
    return {
      ordenes:             ordenes,
      tecnicos:            Object.keys(tecSet).sort(),
      tecnicosDisponibles: tecnicosDisp,
      repuestos:           catalogo.repuestos,
      resellers:           catalogo.resellers,
      equipos:             catalogo.equipos,
      manoObra:            catalogo.manoObra || [],
      equiposBateria:      catalogo.equiposBateria || []
    };

  } catch(e) {
    Logger.log("cargarTodo: " + e);
    return { ordenes:[], tecnicos:[], error: e.toString() };
  }
}

// ============================================================
//  CATÁLOGO SEPARADO — llamado una vez post-render
// ============================================================
function cargarCatalogo() {
  try {
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    var hojaRes = getSheet(SCHEMA.SHEETS.RESELLERS);
    var hojaEq  = getSheet(SCHEMA.SHEETS.EQUIPOS);

    var repuestos = [];
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      for (var j = 1; j < dRep.length; j++) {
        if (dRep[j][1] && dRep[j][2]) repuestos.push({ codigo: String(dRep[j][1]), nombre: String(dRep[j][2]), modelos: String(dRep[j][SCHEMA.DB_REPUESTOS.MODELOS] || '') });
      }
    }

    var resellersList = [];
    if (hojaRes) {
      var dRes = getSheetValues(hojaRes);
      for (var k = 1; k < dRes.length; k++) {
        var nombre = String(dRes[k][0]||"").trim();
        if (!nombre) continue;
        resellersList.push({
          nombre: nombre, email: String(dRes[k][CONFIG.COL_EMAIL_RESELLER]||"").trim(),
          cuit: String(dRes[k][1]||""), direccion: String(dRes[k][2]||""),
          cp: String(dRes[k][3]||""), localidad: String(dRes[k][4]||""), telefono: String(dRes[k][6]||"")
        });
      }
    }

    var equipos = [];
    if (hojaEq) {
      var dEq = getSheetValues(hojaEq);
      for (var e = 1; e < dEq.length; e++) {
        var nom = String(dEq[e][0]||"").trim();
        if (nom) equipos.push(nom);
      }
    }

    var manoObra = [];
var hojaMO = getSheet(SCHEMA.SHEETS.PRECIOS_MANO_OBRA);
if (hojaMO) {
  var dMO = getSheetValues(hojaMO);
  for (var mo = 1; mo < dMO.length; mo++) {
    var codMO = String(dMO[mo][0]||"").trim();
    var dscMO = String(dMO[mo][1]||"").trim();
    var prcMO = String(dMO[mo][2]||"").trim();
    if (codMO && dscMO) manoObra.push({ codigo: codMO, descripcion: dscMO, precio: prcMO });
  }
}
return { repuestos: repuestos, resellers: resellersList, equipos: equipos, manoObra: manoObra };
  } catch(e) {
    Logger.log("cargarCatalogo: " + e);
    return { repuestos: [], resellers: [], equipos: [], error: e.toString() };
  }
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


// ============================================================
//  TÉCNICOS DISPONIBLES — desde Usuarios_Internos col D
// ============================================================
function obtenerTecnicosDisponibles() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.USUARIOS);
    if (!hoja) return [];
    var datos = getSheetValues(hoja);
    var lista = [];
    for (var i = 1; i < datos.length; i++) {
      var f      = datos[i];
      var nombre = String(f[0]||"").trim();
      var email  = String(f[1]||"").trim();
      var rol    = String(f[2]||"operador").trim().toLowerCase();
      var esTec  = String(f[3]||"").trim().toLowerCase() === "si";
      if (nombre && esTec) lista.push({ nombre: nombre, email: email, rol: rol });
    }
    Logger.log("Técnicos disponibles: " + lista.length);
    return lista;
  } catch(e) {
    Logger.log("obtenerTecnicosDisponibles: " + e);
    return [];
  }
}


// ============================================================
//  DETALLE DE OT — campos pesados al abrir una OT
// ============================================================
function obtenerDetalleOT(fila) {
  try {
    var filaNum = parseInt(fila);
    if (isNaN(filaNum) || filaNum < 2) return { trabajo: "", repuestos: "", mensajes: "", stockPorCodigo: {} };
    var datos = getSheetValues(SCHEMA.SHEETS.OT, true);  // force=true: siempre leer del sheet, nunca del cache
    var f = datos[filaNum - 1];
    if (!f) return { trabajo: "", repuestos: "", mensajes: "", stockPorCodigo: {} };

    // Build stock map for repuestos in this OT
    var repStr = String(f[16] || "").trim();
    var stockPorCodigo = {};
    if (repStr && repStr !== 'Sin consumo de repuestos') {
      var dStock = getSheetValues(SCHEMA.SHEETS.STOCK);
      var ST = SCHEMA.STOCK_REPUESTOS;
      var stockIdx = {};
      for (var s = 1; s < dStock.length; s++) {
        stockIdx[String(dStock[s][ST.CODIGO] || '').trim().toUpperCase()] = parseInt(dStock[s][ST.STOCK_ACTUAL]) || 0;
      }
      var ls = repStr.split(' ; ');
      for (var r = 0; r < ls.length; r++) {
        var p = ls[r].split(' | ');
        if (p.length < 1) continue;
        var cod = String(p[0]).trim().toUpperCase();
        stockPorCodigo[cod] = stockIdx.hasOwnProperty(cod) ? stockIdx[cod] : null;
      }
    }

    var histRaw = String(f[SCHEMA.OT.HISTORIAL_ESTADOS]||"").trim();
    var historial = [];
    try { if (histRaw) historial = JSON.parse(histRaw); } catch(eh) {}

    var rawUM = f[SCHEMA.OT.ULTIMA_MODIFICACION];

    return {
      trabajo:             String(f[12]||""),
      repuestos:           String(f[16]||""),
      mensajes:            String(f[11]||""),
      stockPorCodigo:      stockPorCodigo,
      historialEstados:    historial,
      ultimaModificacion:  (rawUM instanceof Date) ? rawUM.getTime() : 0
    };
  } catch(e) {
    Logger.log("obtenerDetalleOT ERROR fila=" + fila + " : " + e);
    return { trabajo: "", repuestos: "", mensajes: "", stockPorCodigo: {} };
  }
}


// ============================================================
//  ACTUALIZAR ORDEN
// ============================================================
// ============================================================
//  ACTUALIZAR ORDEN (CORREGIDA Y MAPEADA COLUMNA POR COLUMNA)
// ============================================================
function actualizarOrden(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var fila = parseInt(data.fila);
    var old  = hoja.getRange(fila, 1, 1, 26).getValues()[0];
    var estadoAnterior = String(old[SCHEMA.OT.ESTADO] || "");

    // ── CONTROL DE CONCURRENCIA OPTIMISTA ──────────────────────
    var rawUM = old[SCHEMA.OT.ULTIMA_MODIFICACION];
    var currentUMms = (rawUM instanceof Date) ? rawUM.getTime() : 0;
    var sentUMms    = parseInt(data.ultimaModificacion) || 0;
    if (currentUMms !== sentUMms) {
      return {
        resultado: 'CONFLICT',
        ot:        data.ot || '',
        msg:       'La orden fue modificada por otro usuario. Recargá los datos antes de guardar.'
      };
    }

    // Sello de tiempo en Columna U (21) si el estado cambia
    if (data.estado !== estadoAnterior) {
      hoja.getRange(fila, SCHEMA.OT.FECHA_ESTADO + 1).setValue(new Date());
      // Historial de estados — columna Y (25): JSON array de transiciones
      var celHist = hoja.getRange(fila, SCHEMA.OT.HISTORIAL_ESTADOS + 1);
      var histRaw = celHist.getValue();
      var hist = [];
      try { if (histRaw) hist = JSON.parse(histRaw); } catch(e2) {}
      hist.push({ f: new Date().getTime(), ant: estadoAnterior, nvo: data.estado, tec: data.tecnico || "" });
      celHist.setValue(JSON.stringify(hist));
    }

    // Regla Pikie: limpiar técnico si no es Taller
    var tecnico = String(data.tecnico || "").trim();
    if (data.circuito !== "Taller") {
      hoja.getRange(fila, SCHEMA.OT.TECNICO + 1).setValue("");
      tecnico = "Gestión Reseller";
    } else {
      hoja.getRange(fila, SCHEMA.OT.TECNICO + 1).setValue(tecnico || "");
    }

    // Escritura inequívoca por columna (evita desfasajes)
    hoja.getRange(fila, SCHEMA.OT.GARANTIA         + 1).setValue(data.garantia);
    hoja.getRange(fila, SCHEMA.OT.ESTADO           + 1).setValue(data.estado);
    hoja.getRange(fila, SCHEMA.OT.EQUIPO           + 1).setValue(data.equipo);
    hoja.getRange(fila, SCHEMA.OT.SN               + 1).setValue(data.sn);
    hoja.getRange(fila, SCHEMA.OT.RESELLER         + 1).setValue(data.reseller);
    if (data.mensajes !== undefined) hoja.getRange(fila, SCHEMA.OT.MENSAJES + 1).setValue(data.mensajes || "");
    hoja.getRange(fila, SCHEMA.OT.TRABAJO          + 1).setValue(data.trabajo);
    hoja.getRange(fila, SCHEMA.OT.FECHA_ACTIVACION + 1).setValue(data.factura || "");
    hoja.getRange(fila, SCHEMA.OT.CAS              + 1).setValue(data.cas);
    hoja.getRange(fila, SCHEMA.OT.REPUESTOS        + 1).setValue(data.repuestos);
    hoja.getRange(fila, SCHEMA.OT.PRIORIDAD        + 1).setValue(data.prioridad ? "URGENTE" : "NORMAL");
    hoja.getRange(fila, SCHEMA.OT.CIRCUITO         + 1).setValue(data.circuito);
    if (data.manoObraGuardada !== undefined) {
      hoja.getRange(fila, SCHEMA.OT.MANO_OBRA      + 1).setValue(data.manoObraGuardada);
      if (data.notasInternas !== undefined) {
        hoja.getRange(fila, SCHEMA.OT.NOTAS_INTERNAS + 1).setValue(data.notasInternas || "");
      }
    }

    var esEstadoTerminal = (data.estado === "Finalizado" || data.estado === "Entregado" || data.estado === "CANCELADO" || data.estado === "Rechazado DJI" || data.estado === "Sin respuesta · Cerrado");

    if (data.estado === "Finalizado") {
      hoja.getRange(fila, SCHEMA.OT.FECHA_CIERRE + 1).setValue(new Date());
      if (data.repuestosItems && data.repuestosItems.length) {
        procesarSalidaRepuestos(data.ot, data.repuestosItems, Session.getActiveUser().getEmail());
      }
      _cerrarSolicitudesPendientes(data.ot, Session.getActiveUser().getEmail());
    } else if (data.estado === "Entregado") {
      // Taller central: el equipo fue retirado o despachado — cierra la OT
      if (!old[SCHEMA.OT.FECHA_CIERRE]) hoja.getRange(fila, SCHEMA.OT.FECHA_CIERRE + 1).setValue(new Date());
      _cerrarSolicitudesPendientes(data.ot, Session.getActiveUser().getEmail());
    } else if (data.estado === "CANCELADO") {
      _cerrarSolicitudesPendientes(data.ot, Session.getActiveUser().getEmail());
    }

    sincronizarDeudaReseller(data);

    var cambios = [];
    if (estadoAnterior          !== data.estado)    cambios.push("Estado: " + data.estado);
    if (String(old[SCHEMA.OT.REPUESTOS]) !== data.repuestos) cambios.push("Repuestos");
    if (String(old[SCHEMA.OT.TECNICO]).trim() !== tecnico)   cambios.push("Técnico: " + tecnico);
    var detalle = cambios.length ? "Modificó: " + cambios.join(", ") : "Guardado manual";

    registrarLog(data.ot, tecnico, Session.getActiveUser().getEmail(),
                 "ACTUALIZACIÓN", estadoAnterior, data.estado, detalle);

    data._fechaEstimada = calcularFechaEstimada(data.circuito, data.garantia, old[SCHEMA.OT.FECHA_INGRESO]);
    data.cliente        = String(old[SCHEMA.OT.CLIENTE] || '').trim();

    enviarNotificaciones(data, estadoAnterior, tecnico);
    // No re-registrar despachos si la OT ya cerró — evita recrear filas Pendiente
    if (!esEstadoTerminal) {
      registrarSolicitudDespacho(data.ot, data.reseller, data.repuestos);
    }

    // Sello de concurrencia: siempre al final, tras todas las escrituras
    hoja.getRange(fila, SCHEMA.OT.ULTIMA_MODIFICACION + 1).setValue(new Date());

    invalidateSheetValues(SCHEMA.SHEETS.OT);
    return { resultado: "OK", ot: data.ot };

  } catch(e) {
    Logger.log("actualizarOrden: " + e);
    return { resultado: "Error: " + e.toString(), ot: data.ot };
  } finally {
    SpreadsheetApp.flush();
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ============================================================
//  CANCELAR CASO
//  1. Marca la OT como CANCELADO en 'Ordenes de trabajo'.
//  2. Cancela todas las SOLICITUDES_DESPACHO Pendientes de esa OT.
//  3. Registra el evento en LOGS.
// ============================================================
function cancelarCaso(idOT) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var otBusc = String(idOT).trim();
    if (!otBusc) throw new Error('ID de OT inválido');

    // ── 1. Actualizar estado en 'Ordenes de trabajo' ──────────
    var hojaOT   = getSheet(SCHEMA.SHEETS.OT);
    var datosOT  = hojaOT.getDataRange().getValues();
    var O        = SCHEMA.OT;
    var filaOT   = -1;
    var estadoAnterior = '';

    for (var i = 1; i < datosOT.length; i++) {
      if (String(datosOT[i][O.OT] || '').trim() === otBusc) {
        filaOT = i + 1;
        estadoAnterior = String(datosOT[i][O.ESTADO] || '');
        break;
      }
    }
    if (filaOT < 0) throw new Error('OT no encontrada: ' + otBusc);
    if (estadoAnterior === 'CANCELADO') return { ok: false, msg: 'La OT ya está cancelada.' };

    hojaOT.getRange(filaOT, O.ESTADO       + 1).setValue('CANCELADO');
    hojaOT.getRange(filaOT, O.FECHA_CIERRE + 1).setValue(new Date());
    hojaOT.getRange(filaOT, O.FECHA_ESTADO + 1).setValue(new Date());

    // ── 2. Cancelar solicitudes de despacho Pendientes ────────
    var hojaSol  = getSheet(SCHEMA.SHEETS.SOLICITUDES);
    var canceladas = 0;
    if (hojaSol) {
      var datosSol = hojaSol.getDataRange().getValues();
      var SD       = SCHEMA.SOLICITUDES_DESPACHO;
      var solChanged = false;
      for (var s = 1; s < datosSol.length; s++) {
        if (String(datosSol[s][SD.OT] || '').trim() !== otBusc) continue;
        var estSol = String(datosSol[s][SD.ESTADO] || '').trim();
        if (estSol === 'Pendiente') {
          datosSol[s][SD.ESTADO]        = 'Cancelado';
          datosSol[s][SD.OBSERVACIONES] = 'Cancelado por anulación de OT';
          canceladas++;
          solChanged = true;
        }
      }
      if (solChanged) {
        hojaSol.getDataRange().setValues(datosSol);
        invalidateSheetValues(SCHEMA.SHEETS.SOLICITUDES);
      }
    }

    // ── 3. Registrar en LOGS ──────────────────────────────────
    var hojaLog = getSheet(SCHEMA.SHEETS.LOGS);
    if (hojaLog) {
      hojaLog.appendRow([
        new Date(), otBusc, '', '', '',
        estadoAnterior, 'CANCELADO',
        'Caso cancelado · ' + canceladas + ' solicitudes liberadas',
        Session.getActiveUser().getEmail()
      ]);
      invalidateSheetValues(SCHEMA.SHEETS.LOGS);
    }

    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);

    return { ok: true, ot: otBusc, solicitudesCanceladas: canceladas };
  } catch(e) {
    Logger.log('cancelarCaso: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

// ============================================================
//  SALIDA DE REPUESTOS — descuento atómico al cerrar OT
//  items: [{ sku, cantidad }]
//  Escribe directo en STOCK_REPUESTOS y MOVIMIENTOS_STOCK usando
//  las funciones de Env.js de este mismo contenedor GAS.
//  generarValeMovimiento() vive en SM; el operario lo emite desde allí.
// ============================================================
function procesarSalidaRepuestos(otId, items, operador) {
  var S  = SCHEMA.STOCK_REPUESTOS;
  var SM = SCHEMA.MOVIMIENTOS_STOCK;
  var errores = [];

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    var dStock    = hojaStock.getDataRange().getValues();
    var hojaMov   = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
    var hoy       = new Date();
    var stockChanged = false;

    for (var i = 0; i < items.length; i++) {
      var item    = items[i];
      var codBusc = String(item.sku).trim().toUpperCase();
      var cant    = parseInt(item.cantidad) || 0;
      if (!codBusc || cant <= 0) continue;

      var hallado = false;
      for (var j = 1; j < dStock.length; j++) {
        if (String(dStock[j][S.CODIGO]).trim().toUpperCase() !== codBusc) continue;

        var stockActual = parseInt(dStock[j][S.STOCK_ACTUAL]) || 0;
        if (stockActual < cant) {
          errores.push(codBusc + ': stock insuficiente (' + stockActual + ' disponible)');
          hallado = true;
          break;
        }

        var nuevoStock = stockActual - cant;
        dStock[j][S.STOCK_ACTUAL]  = nuevoStock;
        dStock[j][S.ULTIMA_SALIDA] = hoy;
        stockChanged = true;

        hojaMov.appendRow([
          hoy,
          'EGRESO',
          codBusc,
          String(dStock[j][S.DESCRIPCION] || ''),
          -cant,
          nuevoStock,
          'OT #' + otId,
          operador || '',
          ''
        ]);

        hallado = true;
        break;
      }
      if (!hallado) errores.push(codBusc + ': SKU no encontrado en stock');
    }

    if (stockChanged) hojaStock.getDataRange().setValues(dStock);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);

    return errores.length
      ? { success: false, errores: errores }
      : { success: true };

  } catch(e) {
    Logger.log('procesarSalidaRepuestos [OT=' + otId + ']: ' + e);
    return { success: false, error: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// Al finalizar una OT marca todas sus solicitudes Pendiente como Despachado,
// con cant_despachada = cant_solicitada (los repuestos se usaron en la reparación).
function _cerrarSolicitudesPendientes(idOT, operador) {
  var hojaSol = getSheet(SCHEMA.SHEETS.SOLICITUDES);
  if (!hojaSol) return;
  var SD      = SCHEMA.SOLICITUDES_DESPACHO;
  var otKey   = String(idOT).trim().toUpperCase();
  var datos   = hojaSol.getDataRange().getValues();
  var hoy     = new Date();
  var changed = false;
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][SD.OT] || '').trim().toUpperCase() !== otKey) continue;
    if (String(datos[i][SD.ESTADO] || '') !== 'Pendiente') continue;
    datos[i][SD.CANT_DESPACHADA] = parseInt(datos[i][SD.CANT_SOLICITADA]) || 0;
    datos[i][SD.ESTADO]          = 'Despachado';
    datos[i][SD.FECHA_DESPACHO]  = hoy;
    datos[i][SD.OPERADOR]        = operador || '';
    changed = true;
  }
  if (changed) {
    hojaSol.getDataRange().setValues(datos);
    invalidateSheetValues(SCHEMA.SHEETS.SOLICITUDES);
  }
}

// ============================================================
//  CREAR NUEVA OT (CON CANDADO DE CONCURRENCIA)
// ============================================================
function crearNuevaOT(datos) {
  datos = datos || {};
  var lock = LockService.getScriptLock();
  var nOT  = "";
  try {
    lock.waitLock(10000);

    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var num  = String(hoja.getLastRow() + 1);
    while (num.length < 5) num = "0" + num;
    nOT = "WH/REP/" + num;

    var fechaAct = "";
    if (datos.fechaActivacion) {
      var parts = String(datos.fechaActivacion).split("-");
      if (parts.length === 3) fechaAct = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    }

    var row = new Array(21).fill("");
    row[0]  = new Date();
    row[2]  = nOT;
    row[3]  = datos.garantia  || "OOW";
    row[4]  = "Abierto";
    row[5]  = datos.equipo    || "";
    row[6]  = datos.sn        || "";
    row[7]  = datos.reseller  || "";
    row[9]  = datos.tecnico   || "";
    row[12] = datos.trabajo   || "";
    row[13] = fechaAct        || "";
    row[17] = datos.prioridad || "NORMAL";
    row[18] = datos.circuito  || "Taller";
    row[20] = new Date();
    hoja.appendRow(row);

    registrarLog(nOT, datos.reseller || "Sin asignar", Session.getActiveUser().getEmail(), "CREACIÓN", "-", "Abierto", "Nueva orden");

    return { resultado: nOT, ot: nOT };

  } catch(e) {
    Logger.log("crearNuevaOT: " + e);
    return { resultado: "Error: " + e.toString(), ot: "" };
  } finally {
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ============================================================
//  LOGS DE UNA OT (historial visible en el formulario)
// ============================================================
function obtenerLogs(ot) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.LOGS);
    if (!hoja) return [];
    var datos = getSheetValues(hoja);
    var out   = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (String(f[1]) !== String(ot)) continue;
      var fecha = f[0] instanceof Date
        ? Utilities.formatDate(f[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : String(f[0]);
      out.push({
        fecha:   fecha,
        tecnico: String(f[2]||"—"),
        accion:  String(f[4]||"—"),
        estAnt:  String(f[5]||"—"),
        estNvo:  String(f[6]||"—"),
        detalle: String(f[7]||"—")
      });
    }
    return out.reverse();
  } catch(e) { return []; }
}


// ============================================================
//  HISTORIAL DE EMAILS ENVIADOS (pestaña del HUB)
//  Hoja EMAIL_LOGS: Fecha|OT|Destinatario|Rol|Asunto|Estado
// ============================================================
function obtenerEmailLogs(limite) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.EMAIL_LOGS);
    if (!hoja) return [];
    var datos = getSheetValues(hoja);
    var max   = limite || 200;
    var out   = [];
    for (var i = datos.length - 1; i >= 1 && out.length < max; i--) {
      var f = datos[i];
      var fecha = f[0] instanceof Date
        ? Utilities.formatDate(f[0], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
        : String(f[0]);
      out.push({
        fecha:        fecha,
        ot:           String(f[1]||""),
        destinatario: String(f[2]||""),
        rol:          String(f[3]||""),
        asunto:       String(f[4]||""),
        estado:       String(f[5]||"")
      });
    }
    return out;
  } catch(e) { return []; }
}


// ============================================================
//  MÉTRICAS POR TÉCNICO
// ============================================================
function obtenerMetricasTecnicos() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.OT);
    var datos = getSheetValues(hoja);
    var hoy   = new Date();
    var mapa  = {};

    // Mapa de reincidentes por S/N: cuántos técnicos distintos atendieron el mismo equipo
    var snCount = {};
    for (var s = 1; s < datos.length; s++) {
      var sn = String(datos[s][SCHEMA.OT.SN]||"").trim();
      var tec0 = String(datos[s][SCHEMA.OT.TECNICO]||"").trim();
      if (sn && tec0 && tec0 !== "Gestión Reseller") {
        if (!snCount[sn]) snCount[sn] = {};
        snCount[sn][tec0] = true;
      }
    }

    for (var i = 1; i < datos.length; i++) {
      var f      = datos[i];
      var tec    = String(f[SCHEMA.OT.TECNICO]||"").trim();
      if (!tec || tec === "Gestión Reseller" || !f[SCHEMA.OT.OT]) continue;
      var estado = String(f[SCHEMA.OT.ESTADO]||"");
      if (estado === "CANCELADO") continue;
      if (!mapa[tec]) mapa[tec] = { tecnico:tec, abiertas:0, finalizadas:0,
                                     diasAbiertasTotal:0, diasAbiertasCount:0,
                                     diasCierreTotal:0, diasCierreCount:0,
                                     urgentes:0, reincidentesCount:0 };
      if (String(f[SCHEMA.OT.PRIORIDAD]||"").toUpperCase() === "URGENTE") mapa[tec].urgentes++;

      // Reincidente: S/N con más de una OT en todo el historial
      var snF = String(f[SCHEMA.OT.SN]||"").trim();
      if (snF && snCount[snF] && Object.keys(snCount[snF]).length > 1) mapa[tec].reincidentesCount++;

      if (estado === "Finalizado") {
        mapa[tec].finalizadas++;
        // Tiempo real de cierre: FECHA_CIERRE - FECHA_INGRESO
        if (f[SCHEMA.OT.FECHA_CIERRE] instanceof Date && f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) {
          var dc = Math.floor((f[SCHEMA.OT.FECHA_CIERRE] - f[SCHEMA.OT.FECHA_INGRESO]) / 86400000);
          if (dc >= 0) { mapa[tec].diasCierreTotal += dc; mapa[tec].diasCierreCount++; }
        }
      } else {
        mapa[tec].abiertas++;
        var da = (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? Math.floor((hoy - f[SCHEMA.OT.FECHA_INGRESO]) / 86400000) : 0;
        mapa[tec].diasAbiertasTotal += da;
        mapa[tec].diasAbiertasCount++;
      }
    }
    var out = [], keys = Object.keys(mapa);
    for (var k = 0; k < keys.length; k++) {
      var m = mapa[keys[k]];
      out.push({
        tecnico:      m.tecnico,
        abiertas:     m.abiertas,
        finalizadas:  m.finalizadas,
        urgentes:     m.urgentes,
        promDias:     m.diasAbiertasCount > 0 ? Math.round(m.diasAbiertasTotal / m.diasAbiertasCount) : 0,
        promCierre:   m.diasCierreCount   > 0 ? Math.round(m.diasCierreTotal   / m.diasCierreCount)   : null,
        reincidentes: m.reincidentesCount
      });
    }
    out.sort(function(a,b){ return b.abiertas - a.abiertas; });
    return out;
  } catch(e) { Logger.log("obtenerMetricasTecnicos: " + e); return []; }
}


// ============================================================
//  ADEUDOS (cruce DEUDA_RESELLERS con estado real de OTs)
// ============================================================
function obtenerPendientesEnvio() {
  try {
    var hojaDeuda = getSheet(SCHEMA.SHEETS.DEUDA_RESELLERS);
    var hojaOT    = getSheet(SCHEMA.SHEETS.OT);
    if (!hojaDeuda || !hojaOT) return [];
    var datosOT = getSheetValues(hojaOT);
    var estados = {};
    for (var i = 1; i < datosOT.length; i++) estados[datosOT[i][2]] = datosOT[i][4];
    var deuda = getSheetValues(hojaDeuda);
    var res   = {};
    for (var j = 1; j < deuda.length; j++) {
      var f = deuda[j];
      var k = f[1] + f[3];
      if (!res[k]) res[k] = {
        ot: f[1], reseller: f[2], codigo: f[3], repuesto: f[4],
        pedido: parseInt(f[5])||0, enviado: 0, estado: estados[f[1]]||"Sin estado"
      };
      res[k].enviado += parseInt(f[6])||0;
    }
    var out = [], keys = Object.keys(res);
    for (var k2 = 0; k2 < keys.length; k2++) {
      var p = res[keys[k2]];
      if (p.pedido - p.enviado > 0)
        out.push({ ot:p.ot, reseller:p.reseller, codigo:p.codigo, repuesto:p.repuesto, falta:p.pedido-p.enviado, estado:p.estado });
    }
    return out;
  } catch(e) { return []; }
}


// ============================================================
//  INFO CLIENTE (para remito PDF)
// ============================================================
function obtenerInfoCliente(nombre) {
  try {
    var hoja  = getSheet(SCHEMA.SHEETS.RESELLERS);
    var datos = getSheetValues(hoja);
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toLowerCase() === String(nombre).trim().toLowerCase())
        return { cuit:datos[i][1], direccion:datos[i][2], cp:datos[i][3], localidad:datos[i][4], telefono:datos[i][6] };
    }
    return null;
  } catch(e) { return null; }
}


// ============================================================
//  HELPERS INTERNOS
// ============================================================

// ============================================================
//  SINCRONIZAR DEUDA RESELLER (OPTIMIZACIÓN EN MEMORIA - ES5)
// ============================================================
function sincronizarDeudaReseller(data) {
  var hoja = getSheet(SCHEMA.SHEETS.DEUDA_RESELLERS);
  if (!hoja) return;
  
  var reg = getSheetValues(hoja);
  if (reg.length < 1) return;

  var cabecera = reg[0];
  var filasLimpias = [cabecera];

  // 1. FILTRADO EN MEMORIA (Ultrarrápido)
  // Guardamos todas las filas EXCEPTO las que pertenecen a la OT que estamos editando
  for (var i = 1; i < reg.length; i++) {
    if (reg[i][1] !== data.ot) {
      filasLimpias.push(reg[i]);
    }
  }

  // 2. REESCRITURA MASIVA (Evitamos el deleteRow iterativo)
  if (filasLimpias.length < reg.length) {
    hoja.clearContents();
    hoja.getRange(1, 1, filasLimpias.length, filasLimpias[0].length).setValues(filasLimpias);
  }

  // Si es Taller Central o no hay consumo, no hay backorder que registrar
  if (data.circuito === "Taller") return;
  if (!data.repuestos || data.repuestos === "Sin consumo de repuestos") return;

  // 3. AGREGAR NUEVAS DEUDAS (En bloque)
  var ls = data.repuestos.split(" ; ");
  var filasAgregar = [];
  var fechaHoy = new Date();

  for (var j = 0; j < ls.length; j++) {
    var p = ls[j].split(" | ");
    if (p.length < 3) continue;
    
    var ped = parseInt(p[2].split(" E:")[0].replace("P:","")) || 0;
    var env = parseInt(p[2].split(" E:")[1]) || 0;
    
    if (ped > env) {
      var cod  = p[0].trim();
      var desc = p[1].replace("(" + cod + ")", "").trim();
      // Preparamos la fila en memoria
      filasAgregar.push([fechaHoy, data.ot, data.reseller, cod, desc, ped, env, ped - env, "Pendiente", data.estado]);
    }
  }

  // Pegamos todas las deudas nuevas de una sola vez
  if (filasAgregar.length > 0) {
    hoja.getRange(hoja.getLastRow() + 1, 1, filasAgregar.length, filasAgregar[0].length).setValues(filasAgregar);
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
      hoja = SS().insertSheet("EMAIL_LOGS");
      hoja.appendRow(["Fecha","OT","Destinatario","Rol","Asunto","Estado","ThreadID"]);
    }
    hoja.appendRow([new Date(), ot, destinatario, rol, asunto, estado, threadId || '']);
    invalidateSheetValues(SCHEMA.SHEETS.EMAIL_LOGS);
  } catch(e) { Logger.log("registrarEmailLog: " + e); }
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
    Logger.log("=== NOTIF " + data.ot + " | " + estadoAnterior + " → " + estadoNuevo + " ===");

    // 0. REPOSICIÓN BATERÍA — se dispara al confirmar el envío del scrap a DJI
    if (estadoNuevo === "Scrap Enviado (Evidencias)" && estadoAnterior !== "Scrap Enviado (Evidencias)") {
      if (esBateria(data.equipo)) {
        enviarEmailReposicionBateria(data);
      }
    }

    // 1. RESELLER
    if (estaEnLista(estadoNuevo, CONFIG.ESTADOS_NOTIFICAR_RESELLER)) {
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
    if (CONFIG.SUPERVISOR_RECIBE_BACKORDER && tieneBackorder)   motivoSup.push("Backorder");
    if (motivoSup.length > 0) {
      var asuntoS = "[HUB] " + motivoSup.join(" · ") + " — " + data.ot;
      try {
        var tidS = _enviarConHilo(data.ot, CONFIG.EMAIL_SUPERVISOR, asuntoS, armarEmailSupervisor(data, estadoAnterior, estadoNuevo, tecnico, tieneBackorder));
        registrarEmailLog(data.ot, CONFIG.EMAIL_SUPERVISOR, "Supervisor", asuntoS, "OK", tidS || "");
      } catch(e) {
        registrarEmailLog(data.ot, CONFIG.EMAIL_SUPERVISOR, "Supervisor", asuntoS, "ERROR: " + e.message, "");
      }
    }
  } catch(e) { Logger.log("enviarNotificaciones: " + e); }
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
  var informe = (nvo==="Finalizado" && data.trabajo)
    ? "<div style='margin-top:20px;background:#f5f9fc;border-left:3px solid #00a3e0;border-radius:0 6px 6px 0;padding:14px 16px'><p style='font-size:11px;font-weight:600;color:#888;text-transform:uppercase;letter-spacing:.06em;margin:0 0 8px'>Informe técnico</p><p style='font-size:13px;color:#444;line-height:1.7;margin:0'>" + data.trabajo + "</p></div>"
    : "";
  return construirEmailHTML(
    "Actualización de su Orden de Servicio", "Estimado/a " + data.reseller,
    urgBanner +
    "<p style='font-size:14px;color:#444;line-height:1.7;margin:0 0 22px'>" + (msgs[nvo]||"Su orden fue actualizada.") + "</p>" +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:4px'>" + ficha + "</div>" +
    construirTablaRepuestos(data.repuestos) + informe,
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
  if (back)            alertas += bloqueCard("📦 Backorder","Repuestos pendientes de envío.","#e67e22");
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

function construirTablaRepuestos(rep) {
  if (!rep || rep === "Sin consumo de repuestos") return "";
  var ls = rep.split(" ; "), filas = "", hayBack = false;
  for (var i = 0; i < ls.length; i++) {
    var p = ls[i].split(" | ");
    if (p.length < 3) continue;
    var cod  = p[0].trim();
    var des  = (p[1]||"").replace("("+cod+")","").replace(/\(\s*\)/g,"").trim();
    var ped  = parseInt(p[2].split(" E:")[0].replace("P:",""))||0;
    var env  = parseInt(p[2].split(" E:")[1])||0;
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


// ============================================================
//  TRIGGER DIARIO — SLA

function _bloquesBotonesPresupuesto(ot) {
  var urlBase   = CONFIG.PORTAL_URL;
  var tokAprob  = _tokenAprobacion(ot, 'aprobar');
  var tokRechaz = _tokenAprobacion(ot, 'rechazar');
  var urlAprob  = urlBase + '?action=aprobar&ot='  + encodeURIComponent(ot) + '&token=' + tokAprob;
  var urlRechaz = urlBase + '?action=rechazar&ot=' + encodeURIComponent(ot) + '&token=' + tokRechaz;
  return "<div style='background:#f9f9f9;border:1px solid #e0e0e0;border-radius:8px;padding:20px 24px;margin-bottom:16px;text-align:center'>" +
    "<p style='font-size:13px;color:#555;margin:0 0 16px;line-height:1.5'>Por favor revisá el detalle y tomá una decisión:</p>" +
    "<a href='" + urlAprob + "' style='display:inline-block;background:#1a9e4a;color:#fff;font-weight:700;font-size:14px;padding:12px 28px;border-radius:6px;text-decoration:none;margin:0 8px'>✅ Aprobar presupuesto</a>" +
    "<a href='" + urlRechaz + "' style='display:inline-block;background:#e74c3c;color:#fff;font-weight:700;font-size:14px;padding:12px 28px;border-radius:6px;text-decoration:none;margin:0 8px'>❌ Rechazar presupuesto</a>" +
    "<p style='font-size:11px;color:#aaa;margin:14px 0 0'>Este link es de uso único y está asociado a tu cuenta. Ante dudas respondé este email.</p>" +
  "</div>";
}

// ============================================================
//  ENVIAR PRESUPUESTO AL RESELLER — manual desde el HUB
// ============================================================
function enviarPresupuesto(data) {
  try {
    var emailReseller = obtenerEmailReseller(data.reseller);
    if (!emailReseller) return { ok: false, msg: "El reseller no tiene email registrado." };

    // Precios repuestos desde DB_REPUESTOS col F
    var precioMap = {};
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      for (var pr = 1; pr < dRep.length; pr++) {
        var codR = String(dRep[pr][1]||"").trim().toUpperCase();
        if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5]||"0").replace(",",".")) || 0;
      }
    }

    // Tabla repuestos con precios
    var tablaRepHTML = "";
    var totalRep = 0;
    if (data.repuestos && data.repuestos !== "Sin consumo de repuestos") {
      var ls = data.repuestos.split(" ; ");
      tablaRepHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border:1px solid #e0e0e0'>Cant.</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>P.Unit (USD)</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Subtotal (USD)</th>" +
        "</tr></thead><tbody>";
      for (var ri = 0; ri < ls.length; ri++) {
        var p = ls[ri].split(" | ");
        if (p.length < 3) continue;
        var cod     = p[0].trim();
        var desc    = p[1].trim();
        var ped     = parseInt((p[2].split(" E:")[0] || "").replace("P:","")) || 0;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var descPct = (data.tipoDestinatario === "cliente") ? 0.25 : 0.40;
        var pvp     = pvpBase > 0 ? pvpBase * (1 - descPct) : 0;
        var sub     = pvp * ped;
        totalRep   += sub;
        var labelDesc = (data.tipoDestinatario === "cliente") ? "25% desc." : "40% desc.";
        tablaRepHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + cod + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + desc + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:center'>" + ped + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + (pvpBase ? "USD " + _fmtNum(pvpBase) + " <span style='color:#888;font-size:10px'>(-" + labelDesc + ")</span>" : "—") + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right;font-weight:700'>" + (pvp ? "USD " + _fmtNum(pvp) : "—") + "</td>" +
          "</tr>";
      }
      if (totalRep > 0) {
        tablaRepHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='4' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL REPUESTOS</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalRep) + "</td></tr>";
      }
      tablaRepHTML += "</tbody></table>";
    }

    // Tabla mano de obra
    var tablaMOHTML = "";
    var totalMO = 0;
    if (data.manoObra && data.manoObra.length) {
      tablaMOHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Precio</th>" +
        "</tr></thead><tbody>";
      for (var mi = 0; mi < data.manoObra.length; mi++) {
        var mo = data.manoObra[mi];
        var pMO = parseFloat(String(mo.precio).replace(/[^0-9.]/g,"")) || 0;
        totalMO += pMO;
        tablaMOHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + mo.codigo + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + mo.descripcion + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + mo.precio + "</td></tr>";
      }
      if (totalMO > 0) {
        tablaMOHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='2' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL MANO DE OBRA</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalMO) + "</td></tr>";
      }
      tablaMOHTML += "</tbody></table>";
    }

    var cuerpoDetalle =
      filaDetalle("Orden de Trabajo", data.ot) +
      filaDetalle("Equipo / Modelo",  data.equipo) +
      filaDetalle("Nº de Serie",      data.sn || "—") +
      (data.cas ? filaDetalle("Caso DJI (CAS/FWR)", data.cas) : "") +
      filaDetalle("Garantía", "Fuera de garantía (OOW)");

    var bloques =
      bloqueCard("📋 Detalle de la Orden", cuerpoDetalle, "#00a3e0") +
      (data.trabajo ? bloqueCard("🔧 Diagnóstico Técnico",
        "<p style='margin:0;font-size:13px;line-height:1.6;color:#555'>" + data.trabajo.replace(/\n/g,"<br>") + "</p>",
        "#27ae60") : "") +
      (tablaRepHTML ? "<div style='margin-bottom:16px'><p style='font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase'>Repuestos</p>" + tablaRepHTML + "</div>" : "") +
      (tablaMOHTML  ? "<div style='margin-bottom:16px'><p style='font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase'>Mano de Obra</p>" + tablaMOHTML + "</div>" : "") +
      _bloquesBotonesPresupuesto(data.ot);

    var htmlEmail = construirEmailHTML(
      "Presupuesto de Reparación — " + data.ot,
      "Estimado equipo de <strong>" + data.reseller + "</strong>,<br>Le enviamos el presupuesto para la siguiente orden:",
      bloques,
      "Ante cualquier consulta no dude en contactarnos."
    );

    var asunto = "OT " + data.ot + " — " + data.equipo;
    var tidPres = _enviarConHilo(data.ot, emailReseller, asunto, htmlEmail);
    registrarEmailLog(data.ot, emailReseller, "Reseller", asunto, "OK", tidPres || "");
    var tidCopia = _enviarConHilo(data.ot, CONFIG.EMAIL_SUPERVISOR, "[COPIA] " + asunto, htmlEmail);
    registrarEmailLog(data.ot, CONFIG.EMAIL_SUPERVISOR, "Supervisor", "[COPIA] " + asunto, "OK", tidCopia || "");

    // Actualizar estado a "Presupuesto enviado" automáticamente
    try {
      var hoja = getSheet(SCHEMA.SHEETS.OT);
      var todos = getSheetValues(hoja);
      for (var ri = 1; ri < todos.length; ri++) {
        if (String(todos[ri][SCHEMA.OT.OT] || "").trim() === String(data.ot).trim()) {
          var fila = ri + 1;
          hoja.getRange(fila, SCHEMA.OT.ESTADO + 1).setValue("Presupuesto enviado");
          hoja.getRange(fila, SCHEMA.OT.FECHA_ESTADO + 1).setValue(new Date());
          hoja.getRange(fila, SCHEMA.OT.ULTIMA_MODIFICACION + 1).setValue(new Date());
          invalidateSheetValues(SCHEMA.SHEETS.OT);
          registrarLog(data.ot, "", Session.getActiveUser().getEmail(), "ACTUALIZACIÓN", "En Revision", "Presupuesto enviado", "Presupuesto enviado via botón");
          break;
        }
      }
    } catch(eUpd) { Logger.log("enviarPresupuesto: no se pudo actualizar estado: " + eUpd); }

    return { ok: true, emailEnviado: emailReseller };

  } catch(e) {
    Logger.log("enviarPresupuesto ERROR: " + e);
    return { ok: false, msg: e.toString() };
  }
}

function _getPresupNum() {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty('PRESUP_COUNTER') || '1000') + 1;
  props.setProperty('PRESUP_COUNTER', String(n));
  return '0002-' + (Array(8).join('0') + n).slice(-7);
}

// Devuelve el HTML del presupuesto sin enviarlo (para vista previa / descarga)
function obtenerPresupuestoHTML(data) {
  try {
    // Reseller data
    var rsData = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var rLow = String(data.reseller || '').trim().toLowerCase();
    var rCuit = '', rDir = '', rCp = '', rLoc = '', rTel = '';
    for (var ri = 1; ri < rsData.length; ri++) {
      if (String(rsData[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rLow) {
        rCuit = String(rsData[ri][SCHEMA.RESELLERS.CUIT]      || '').trim();
        rDir  = String(rsData[ri][SCHEMA.RESELLERS.DIRECCION] || '').trim();
        rCp   = String(rsData[ri][SCHEMA.RESELLERS.CP]        || '').trim();
        rLoc  = String(rsData[ri][SCHEMA.RESELLERS.LOCALIDAD] || '').trim();
        rTel  = String(rsData[ri][SCHEMA.RESELLERS.TELEFONO]  || '').trim();
        break;
      }
    }

    // Presupuesto number
    var presupNum = data.presupNum || (data.modoPreview ? 'BORRADOR' : _getPresupNum());

    // Prices from DB_REPUESTOS
    var precioMap = {};
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      for (var pr = 1; pr < dRep.length; pr++) {
        var codR = String(dRep[pr][1]||"").trim().toUpperCase();
        if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5]||"0").replace(",",".")) || 0;
      }
    }

    var tablaRepHTML = "", totalRep = 0;
    if (data.repuestos && data.repuestos !== "Sin consumo de repuestos") {
      var ls = data.repuestos.split(" ; ");
      tablaRepHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border:1px solid #e0e0e0'>Cant.</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>P.Unit (USD)</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Subtotal (USD)</th>" +
        "</tr></thead><tbody>";
      for (var ri = 0; ri < ls.length; ri++) {
        var p = ls[ri].split(" | ");
        if (p.length < 3) continue;
        var cod  = p[0].trim();
        var desc = p[1].trim();
        var ped  = parseInt((p[2].split(" E:")[0]||"").replace("P:","")) || 0;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var descPct = (data.tipoDestinatario === "cliente") ? 0.25 : 0.40;
        var pvp  = pvpBase > 0 ? pvpBase * (1 - descPct) : 0;
        var sub  = pvp * ped;
        totalRep += sub;
        var labelDesc = (data.tipoDestinatario === "cliente") ? "25% desc." : "40% desc.";
        tablaRepHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + cod + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + desc + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:center'>" + ped + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + (pvpBase ? "USD " + _fmtNum(pvpBase) + " <span style='color:#888;font-size:10px'>(-" + labelDesc + ")</span>" : "—") + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right;font-weight:700'>" + (pvp ? "USD " + _fmtNum(pvp) : "—") + "</td></tr>";
      }
      if (totalRep > 0) {
        tablaRepHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='4' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL REPUESTOS</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalRep) + "</td></tr>";
      }
      tablaRepHTML += "</tbody></table>";
    }

    var tablaMOHTML = "", totalMO = 0;
    if (data.manoObra && data.manoObra.length) {
      tablaMOHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Precio</th>" +
        "</tr></thead><tbody>";
      for (var mi = 0; mi < data.manoObra.length; mi++) {
        var mo = data.manoObra[mi];
        var pMO = parseFloat(String(mo.precio).replace(/[^0-9.]/g,"")) || 0;
        totalMO += pMO;
        tablaMOHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + mo.codigo + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + mo.descripcion + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + mo.precio + "</td></tr>";
      }
      if (totalMO > 0) {
        tablaMOHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='2' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL MANO DE OBRA</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalMO) + "</td></tr>";
      }
      tablaMOHTML += "</tbody></table>";
    }

    var totalGeneral = totalRep + totalMO;
    var totalFila = totalGeneral > 0
      ? "<table style='width:100%;border-collapse:collapse;margin-top:8px'><tr style='background:#00a3e0;color:#fff;font-weight:700;font-size:14px'>" +
        "<td style='padding:10px 14px;text-align:right'>TOTAL GENERAL</td>" +
        "<td style='padding:10px 14px;text-align:right;white-space:nowrap'>USD " + _fmtNum(totalGeneral) + "</td></tr></table>"
      : "";

    var cuerpoDetalle =
      filaDetalle("Orden de Trabajo", data.ot) +
      filaDetalle("Equipo / Modelo",  data.equipo) +
      filaDetalle("Nº de Serie",      data.sn || "—") +
      (data.cas ? filaDetalle("Caso DJI", data.cas) : "") +
      filaDetalle("Reseller",         data.reseller) +
      filaDetalle("Garantía",         "Fuera de garantía (OOW)") +
      filaDetalle("Fecha",            Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));

    // Build items table rows
    var descPct   = (data.tipoDestinatario === 'cliente') ? 0.25 : 0.40;
    var descLabel = (data.tipoDestinatario === 'cliente') ? '25,00%' : '40,00%';
    var tablaFilas = '', totalGeneral = 0;

    if (data.repuestos && data.repuestos !== 'Sin consumo de repuestos') {
      var ls = data.repuestos.split(' ; ');
      for (var li = 0; li < ls.length; li++) {
        var p = ls[li].split(' | ');
        if (p.length < 3) continue;
        var cod  = p[0].trim();
        var desc = p[1].trim();
        var ped  = parseInt((p[2].split(' E:')[0]||'').replace('P:','')) || 0;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var pvpNeto = pvpBase > 0 ? pvpBase * (1 - descPct) : 0;
        var sub  = pvpNeto * ped;
        totalGeneral += sub;
        tablaFilas +=
          '<tr>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px;color:#00a3e0;font-weight:700">' + cod + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px">' + desc + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px">' + (pvpBase ? _fmtNum(pvpBase) : '—') + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:center;font-size:10px">' + descLabel + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px">' + (pvpNeto ? _fmtNum(pvpNeto) : '—') + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:center;font-size:10px">' + ped + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px;font-weight:700">' + (sub ? _fmtNum(sub) : '—') + '</td>' +
          '</tr>';
      }
    }

    if (data.manoObra && data.manoObra.length) {
      for (var mi2 = 0; mi2 < data.manoObra.length; mi2++) {
        var mo = data.manoObra[mi2];
        var pMO = parseFloat(String(mo.precio).replace(/[^0-9.,]/g,'').replace(',','.')) || 0;
        totalGeneral += pMO;
        tablaFilas +=
          '<tr style="background:#f8fff8">' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px;color:#27ae60;font-weight:700">' + mo.codigo + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px">' + mo.descripcion + '</td>' +
          '<td colspan="3" style="padding:4px 7px;border:1px solid #ccc;font-size:9px;color:#888;text-align:center">Revisión y mano de obra</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:center;font-size:10px">1</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px;font-weight:700">' + _fmtNum(pMO) + '</td>' +
          '</tr>';
      }
    }

    // Blank filler rows
    for (var ei = 0; ei < 4; ei++) {
      tablaFilas += '<tr><td style="border:1px solid #ccc;height:18px" colspan="7">&nbsp;</td></tr>';
    }

    // Date parts
    var tz  = Session.getScriptTimeZone();
    var now = new Date();
    var dia  = Utilities.formatDate(now, tz, 'dd');
    var mes2 = Utilities.formatDate(now, tz, 'MM');
    var anio = Utilities.formatDate(now, tz, 'yyyy');

    // Technician initials
    var tecNombre = data.tecnico || '';
    var tecIni = tecNombre ? tecNombre.split(' ').map(function(w){ return w.charAt(0).toUpperCase(); }).join('') : '—';

    // Borrador watermark
    var wm = data.modoPreview
      ? '<div style="position:fixed;top:45%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:90px;font-weight:900;color:rgba(200,0,0,0.06);pointer-events:none;white-space:nowrap;z-index:999">BORRADOR</div>'
      : '';

    var html =
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>' +
      'body{font-family:Arial,sans-serif;margin:0;padding:14px 18px;font-size:11px;color:#222}' +
      'table{border-collapse:collapse}' +
      '.it td,.it th{border:1px solid #ccc}' +
      '.it th{background:#e6e6e6;padding:5px 7px;font-size:10px}' +
      '.total-r td{background:#111;color:#fff;font-weight:700;font-size:13px;padding:6px 9px;text-align:right;border:1px solid #111}' +
      '.fbox{border:1px solid #444;display:inline-block;min-width:28px;padding:2px 6px;text-align:center;font-size:17px;font-weight:700}' +
      '.lbl{font-size:9px;color:#888}' +
      '.nota{font-size:8.5px;color:#555;line-height:1.5}' +
      '.footer-bar{text-align:center;font-size:10px;color:#00a3e0;font-weight:700;padding-top:7px;border-top:2px solid #00a3e0;margin-top:8px}' +
      '@media print{body{padding:8px 10px}.print-bar{display:none}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
      '</style></head><body>' +
      wm +
      // ── HEADER ──
      '<table style="width:100%;margin-bottom:6px"><tr>' +
        '<td style="width:38%;vertical-align:top">' +
          '<div style="font-size:24px;font-weight:900;color:#00a3e0;line-height:1">Bidcom<span style="color:#222">Agro</span></div>' +
          '<div style="font-size:8px;color:#666;line-height:1.5;margin-top:3px">Parque Industrial Carmen de Areco Lote 28<br>B6725 Carmen de Areco, Prov. Buenos Aires<br>Tel: 2325-65-6826</div>' +
        '</td>' +
        '<td style="width:24%;text-align:center;vertical-align:middle">' +
          '<div style="border:3px solid #333;display:inline-block;padding:4px 12px;font-size:34px;font-weight:900;line-height:1">P</div><br>' +
          '<div style="border:2px solid #e74c3c;color:#e74c3c;font-size:8px;font-weight:700;padding:3px 6px;margin-top:4px;display:inline-block;line-height:1.3;text-align:center">DOCUMENTO NO VALIDO<br>COMO FACTURA</div>' +
        '</td>' +
        '<td style="width:38%;text-align:right;vertical-align:top">' +
          '<div style="font-size:19px;font-weight:900;letter-spacing:.04em">PRESUPUESTO</div>' +
          '<div style="font-size:12px;font-weight:700;color:#555;margin-bottom:6px">N° ' + presupNum + '</div>' +
          '<span class="fbox">' + dia + '</span>' +
          '<span style="font-size:11px;color:#aaa;margin:0 3px">|</span>' +
          '<span class="fbox">' + mes2 + '</span>' +
          '<span style="font-size:11px;color:#aaa;margin:0 3px">|</span>' +
          '<span class="fbox">' + anio + '</span>' +
          '<div style="font-size:8px;color:#aaa;margin-top:2px">Día &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mes &nbsp;&nbsp;&nbsp;&nbsp; Año</div>' +
        '</td>' +
      '</tr></table>' +
      // ── RESELLER INFO ──
      '<table class="it" style="width:100%;margin-bottom:5px"><tr>' +
        '<td style="width:12%;padding:4px 7px" class="lbl">Señores:</td>' +
        '<td style="width:50%;padding:4px 7px;font-weight:700;font-size:12px">' + data.reseller + '</td>' +
        '<td style="width:38%;padding:4px 7px" class="lbl">CUIT N°: <strong style="color:#222">' + (rCuit||'—') + '</strong></td>' +
      '</tr><tr>' +
        '<td style="padding:4px 7px" class="lbl">Domicilio:</td>' +
        '<td style="padding:4px 7px;font-size:11px">' + (rDir||'—') + '</td>' +
        '<td style="padding:4px 7px" class="lbl">Localidad: <strong style="color:#222">' + (rLoc||'—') + '</strong></td>' +
      '</tr><tr>' +
        '<td style="padding:4px 7px" class="lbl">Provincia:</td>' +
        '<td style="padding:4px 7px;font-size:10px">Prov. de Buenos Aires &nbsp; <span class="lbl">Cód. Postal:</span> ' + (rCp||'—') + ' &nbsp; <span class="lbl">Descuentos (Repuestos):</span> <strong>' + descLabel + '</strong></td>' +
        '<td style="padding:4px 7px" class="lbl">Técnico Asignado: <strong style="color:#222">' + tecIni + '</strong></td>' +
      '</tr><tr>' +
        '<td style="padding:4px 7px" class="lbl">Teléfono:</td>' +
        '<td style="padding:4px 7px;font-size:11px">' + (rTel||'—') + '</td>' +
        '<td style="padding:4px 7px" class="lbl">Referencia: <strong style="color:#222">' + data.ot + '</strong>' + (data.cas ? ' &nbsp; <span class="lbl">CAS:</span> <strong style="color:#222">' + data.cas + '</strong>' : '') + '</td>' +
      '</tr>' +
      (data.trabajo ? '<tr><td style="padding:4px 7px" class="lbl">Diagnóstico:</td><td colspan="2" style="padding:4px 7px;font-size:10px;color:#2d6a3f;background:#f5fff8">' + data.trabajo.replace(/\n/g,' ') + '</td></tr>' : '') +
      '</table>' +
      // ── ITEMS TABLE ──
      '<table class="it" style="width:100%;margin-bottom:5px">' +
        '<thead><tr>' +
          '<th style="width:14%">Cód</th>' +
          '<th style="width:30%">Descripción</th>' +
          '<th style="width:13%;text-align:right">Precio (USD) s/i</th>' +
          '<th style="width:9%;text-align:center">Dto.</th>' +
          '<th style="width:13%;text-align:right">P. Neto (USD)</th>' +
          '<th style="width:8%;text-align:center">Cantidad</th>' +
          '<th style="width:13%;text-align:right">Subtotal (USD)</th>' +
        '</tr></thead>' +
        '<tbody>' + tablaFilas + '</tbody>' +
        '<tfoot><tr class="total-r"><td colspan="6">Total USD (Sin impuestos)</td><td>' + (totalGeneral ? _fmtNum(totalGeneral) : '—') + '</td></tr></tfoot>' +
      '</table>' +
      // ── NOTAS ──
      '<table style="width:100%;border:1px solid #ccc"><tr><td style="padding:6px 9px">' +
        '<div class="nota">NOTA: Este presupuesto se basa en los datos visibles e informados por el cliente durante la inspección. No incluye costos por daños ocultos o no visibles. Si se descubren daños adicionales durante la reparación, se procederá a informar un nuevo presupuesto. Los costos finales pueden variar según los datos identificados durante el proceso de reparación. Este presupuesto tiene una validez de 5 días hábiles a partir de la fecha de su realización.</div>' +
        '<div class="nota" style="margin-top:3px">*Cotización expresada en dólares estadounidenses; los pagos en pesos argentinos se tomarán al TC Banco Nación Venta Billetes al momento de recepción de valores, efectivo o depósito.<br>' +
        '*La presente oferta no constituye un compromiso y está sujeta a aprobación y evaluación de riesgo crediticio.<br>' +
        '*Lugar de Entrega: Depósito Parque Industrial Carmen de Areco. No incluye envío.<br>' +
        '*Condición de pago: 10 días a partir del efectivo pago.</div>' +
      '</td></tr></table>' +
      // ── FOOTER ──
      '<div class="footer-bar">www.dji.bidcomagro.com.ar &nbsp;|&nbsp; www.brumby.com.ar &nbsp;|&nbsp; 0810-345-9002</div>' +
      '</body></html>';

    return { ok: true, html: html };
  } catch(e) {
    Logger.log("obtenerPresupuestoHTML ERROR: " + e);
    return { ok: false, msg: e.toString() };
  }
}


// ============================================================
//  VISTA SUPERVISOR — backorder predictivo + SLA por estado
// ============================================================
function obtenerDatosSupervisor() {
  try {
    var hoja   = getSheet(SCHEMA.SHEETS.OT);
    var hojaLog = getSheet(SCHEMA.SHEETS.LOGS);
    var datos  = getSheetValues(hoja);
    var hoy    = new Date();

    // ── Backorder predictivo ─────────────────────────────────
    var EST_BACK_R  = ["Aprobacion DJI","Pedido de repuestos"];
    var EST_BACK_RP = ["Pedido de repuesto para reparar","Reparado y aprobado en el aftersales"];
    var mapaRep = {};

    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (!f[2]) continue;
      var estado  = String(f[4]||"");
      if (estado === "CANCELADO") continue;
      var circUp  = String(f[18]||"").trim().toUpperCase();
      var esR      = (circUp === "RESELLER" || circUp === "SI");
      var esRP    = (circUp === "RESELLER PROPIO");
      var rep     = String(f[16]||"").trim();
      if (!rep || rep === "Sin consumo de repuestos") continue;

      var esBack = false;
      if (esR  && EST_BACK_R.indexOf(estado)  !== -1) esBack = true;
      if (esRP && EST_BACK_RP.indexOf(estado) !== -1) esBack = true;
      if (!esBack) continue;

      var ls = rep.split(" ; ");
      for (var j = 0; j < ls.length; j++) {
        var p   = ls[j].split(" | ");
        if (p.length < 3) continue;
        var cod = p[0].trim().toUpperCase();
        var nom = p[1].trim();
        var ped = parseInt(p[2].split("E:")[0].replace("P:","")) || 0;
        var env = parseInt(p[2].split("E:")[1]) || 0;
        var pend = ped - env;
        if (pend <= 0) continue;
        if (!mapaRep[cod]) mapaRep[cod] = { codigo: cod, nombre: nom, cantOTs: 0, pendiente: 0, ots: [] };
        mapaRep[cod].cantOTs++;
        mapaRep[cod].pendiente += pend;
        mapaRep[cod].ots.push(String(f[2]));
      }
    }
    var backorderPred = [];
    var keys = Object.keys(mapaRep);
    for (var k = 0; k < keys.length; k++) backorderPred.push(mapaRep[keys[k]]);
    backorderPred.sort(function(a,b){ return b.cantOTs - a.cantOTs; });

    // ── SLA por estado (últimos 30 días desde LOGS) ──────────
    var slaData = { Taller: {}, Reseller: {}, "Reseller Propio": {} };
    var hace30  = new Date(hoy.getTime() - 30 * 86400000);

    if (hojaLog) {
      var logs = getSheetValues(hojaLog);
      for (var l = 1; l < logs.length; l++) {
        var fl = logs[l];
        if (!(fl[0] instanceof Date) || fl[0] < hace30) continue;
        var otLog    = String(fl[1]||"").trim();
        var estAnt    = String(fl[5]||"").trim();
        var estNvo    = String(fl[6]||"").trim();
        if (!estAnt || !estNvo || estAnt === estNvo || estAnt === "-") continue;

        // Buscar circuito de la OT
        var circOT = "Taller";
        for (var d = 1; d < datos.length; d++) {
          if (String(datos[d][2]).trim() === otLog) {
            var cUp = String(datos[d][18]||"").toUpperCase();
            if (cUp === "RESELLER" || cUp === "SI") circOT = "Reseller";
            else if (cUp === "RESELLER PROPIO") circOT = "Reseller Propio";
            break;
          }
        }

        // Buscar cuánto duró en estAnt
        var duracion = 0;
        for (var l2 = l - 1; l2 >= 1; l2--) {
          if (String(logs[l2][1]).trim() === otLog && String(logs[l2][6]).trim() === estAnt) {
            if (logs[l2][0] instanceof Date) {
              duracion = Math.floor((fl[0] - logs[l2][0]) / 86400000);
            }
            break;
          }
        }
        if (duracion <= 0) continue;
        if (!slaData[circOT][estAnt]) slaData[circOT][estAnt] = { total: 0, count: 0 };
        slaData[circOT][estAnt].total += duracion;
        slaData[circOT][estAnt].count++;
      }
    }

    // Convertir a arrays ordenados
    var slaFinal = {};
    var circs = ["Taller","Reseller","Reseller Propio"];
    for (var ci = 0; ci < circs.length; ci++) {
      var circ = circs[ci];
      var arr  = [];
      var est  = Object.keys(slaData[circ]);
      for (var ei = 0; ei < est.length; ei++) {
        var sd = slaData[circ][est[ei]];
        arr.push({ estado: est[ei], promedio: Math.round(sd.total / sd.count) });
      }
      arr.sort(function(a,b){ return b.promedio - a.promedio; });
      slaFinal[circ] = arr;
    }

    // Tiempo de respuesta de repuestos: solicitud → despacho
    var hojaSolicR = getSheet(SCHEMA.SHEETS.SOLICITUDES);
    var dSolicR    = hojaSolicR ? getSheetValues(hojaSolicR) : [];
    var SD = SCHEMA.SOLICITUDES_DESPACHO;
    var sumDesp = 0, cntDesp = 0;
    for (var si = 1; si < dSolicR.length; si++) {
      var fs = dSolicR[si];
      if (String(fs[SD.ESTADO]).trim() !== 'Despachado') continue;
      var fSol  = fs[SD.FECHA]         instanceof Date ? fs[SD.FECHA]         : null;
      var fDisp = fs[SD.FECHA_DESPACHO] instanceof Date ? fs[SD.FECHA_DESPACHO] : null;
      if (!fSol || !fDisp) continue;
      var diasDisp = Math.floor((fDisp - fSol) / 86400000);
      if (diasDisp >= 0) { sumDesp += diasDisp; cntDesp++; }
    }
    var tiempoDespacho = cntDesp
      ? { promedio: Math.round(sumDesp / cntDesp), count: cntDesp }
      : null;

    // ── Ranking de resellers con más casos abiertos ──────────────
    var EST_CERRADOS_R = ['CANCELADO', 'Finalizado', 'Entregado', 'Partes dañadas scrapeadas'];
    var resMap = {};
    for (var ri = 1; ri < datos.length; ri++) {
      var rf = datos[ri];
      if (!rf[SCHEMA.OT.OT]) continue;
      var estR   = String(rf[SCHEMA.OT.ESTADO]   || '').trim();
      var circR  = String(rf[SCHEMA.OT.CIRCUITO] || '').trim().toUpperCase();
      var esResR = (circR === 'RESELLER' || circR === 'SI' || circR === 'RESELLER PROPIO');
      if (!esResR) continue;
      if (EST_CERRADOS_R.indexOf(estR) !== -1) continue;
      var resellerR = String(rf[SCHEMA.OT.RESELLER] || '').trim() || 'Sin nombre';
      resMap[resellerR] = (resMap[resellerR] || 0) + 1;
    }
    var resellersRanking = [];
    var rkeys = Object.keys(resMap);
    for (var rk = 0; rk < rkeys.length; rk++) {
      resellersRanking.push({ nombre: rkeys[rk], count: resMap[rkeys[rk]] });
    }
    resellersRanking.sort(function(a, b) { return b.count - a.count; });
    if (resellersRanking.length > 10) resellersRanking = resellersRanking.slice(0, 10);

    return { backorderPred: backorderPred, slaData: slaFinal, tiempoDespacho: tiempoDespacho, resellersRanking: resellersRanking };
  } catch(e) {
    Logger.log("obtenerDatosSupervisor: " + e);
    return { backorderPred: [], slaData: {} };
  }
}


// ============================================================
//  MOTOR DE DIAGNÓSTICO — OTs similares por modelo y/o síntoma
// ============================================================
function buscarOTsSimilares(modelo, falla, otActual) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.OT);
    var modeloB = String(modelo||"").trim().toLowerCase();
    var fallaB  = String(falla||"").trim().toLowerCase();
    var otB     = String(otActual||"").trim().toUpperCase();
    var palabras = fallaB.split(/\s+/).filter(function(p){ return p.length > 3; });
    var candidatos = [];

    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (!f[SCHEMA.OT.OT]) continue;
      var otStr = String(f[SCHEMA.OT.OT]).trim().toUpperCase();
      if (otStr === otB) continue;                              // excluir la propia OT
      if (String(f[SCHEMA.OT.ESTADO]||"") !== "Finalizado") continue; // solo cerradas
      var trabajo = String(f[SCHEMA.OT.TRABAJO]||"").trim();
      if (!trabajo) continue;

      var modeloF = String(f[SCHEMA.OT.EQUIPO]||"").trim().toLowerCase();
      var mismoModelo = modeloB && modeloF === modeloB;
      if (!mismoModelo) continue;                               // mismo modelo obligatorio

      // Puntaje por coincidencia de palabras del síntoma en informe técnico
      var trabaLow = trabajo.toLowerCase();
      var score = 0;
      for (var p = 0; p < palabras.length; p++) {
        if (trabaLow.indexOf(palabras[p]) !== -1) score++;
      }

      // Extraer repuestos usados
      var repsRaw = String(f[SCHEMA.OT.REPUESTOS]||"").trim();
      var repsList = [];
      if (repsRaw && repsRaw !== "Sin consumo de repuestos") {
        var ls = repsRaw.split(" ; ");
        for (var r = 0; r < ls.length; r++) {
          var parts = ls[r].split(" | ");
          if (parts.length >= 2) repsList.push(String(parts[1]).trim());
        }
      }

      var fechaStr = (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date)
        ? Utilities.formatDate(f[SCHEMA.OT.FECHA_INGRESO], Session.getScriptTimeZone(), "MM/yyyy")
        : "";

      candidatos.push({
        ot:       otStr,
        equipo:   String(f[SCHEMA.OT.EQUIPO]||""),
        sn:       String(f[SCHEMA.OT.SN]||""),
        garantia: String(f[SCHEMA.OT.GARANTIA]||""),
        trabajo:  trabajo.length > 180 ? trabajo.substring(0, 180) + "…" : trabajo,
        repuestos: repsList.join(", ") || "Sin repuestos",
        fecha:    fechaStr,
        score:    score
      });
    }

    // Ordenar por score descendente; si hay empate, más recientes primero
    candidatos.sort(function(a, b) { return b.score - a.score; });
    return candidatos.slice(0, 5);
  } catch(e) { Logger.log("buscarOTsSimilares: " + e); return []; }
}

// ============================================================
//  HISTORIAL POR S/N — reincidentes
// ============================================================
function buscarHistorialSN(sn) {
  try {
    var hoja  = getSheet(SCHEMA.SHEETS.OT);
    var datos = getSheetValues(hoja);
    var hoy   = new Date();
    var snB   = String(sn).trim().toUpperCase();
    var out   = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (!f[2]) continue;
      if (String(f[6]||"").trim().toUpperCase() !== snB) continue;
      var dias = (f[0] instanceof Date) ? Math.floor((hoy - f[0]) / 86400000) : 0;
      var fechaAp = (f[0] instanceof Date)
        ? Utilities.formatDate(f[0], Session.getScriptTimeZone(), "dd/MM/yyyy")
        : String(f[0]);
      out.push({
        ot:       String(f[2]),
        equipo:   String(f[5]||""),
        reseller: String(f[7]||""),
        estado:   String(f[4]||""),
        dias:     dias,
        fechaAp:  fechaAp,
        trabajo:  String(f[12]||"")
      });
    }
    out.sort(function(a,b){ return b.dias - a.dias; });
    return out;
  } catch(e) {
    Logger.log("buscarHistorialSN: " + e);
    return [];
  }
}


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
      bloqueCard("🔋 Caso aprobado — Reposición requerida",
        "El reseller <strong>" + data.reseller + "</strong> completó los dos pasos requeridos: " +
        "carga del caso reconocido por DJI como en garantía y envío del scrap. " +
        "Se requiere reponer la batería al reseller.",
        "#00a3e0") +
      "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:16px'>" +
      cuerpoDetalle + "</div>" +
      bloqueCard("📦 Acción requerida",
        "Coordinar envío de batería de reemplazo al reseller <strong>" + data.reseller + "</strong>.",
        "#27ae60"),
      "Aviso automático generado por DJI HUB PRO."
    );

    var tidLog = _enviarConHilo(data.ot, emailLog, asunto, htmlEmail);
    registrarEmailLog(data.ot, emailLog, "Logística", asunto, "OK", tidLog || "");
    if (emailReseller) {
      var asuntoR = "OT " + data.ot + " — " + data.equipo;
      var tidRes = _enviarConHilo(data.ot, emailReseller, asuntoR, htmlEmail);
      registrarEmailLog(data.ot, emailReseller, "Reseller", asuntoR, "OK", tidRes || "");
    }
    Logger.log("✓ Reposición batería → " + emailLog);
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

function registrarSolicitudDespacho(idOT, reseller, strRepuestos) {
  var hoja = getSheet(SCHEMA.SHEETS.SOLICITUDES);
  if (!hoja) return;

  var SD    = SCHEMA.SOLICITUDES_DESPACHO;
  var otKey = String(idOT).trim().toUpperCase();

  // Parse new repuesto list into SKU → full item-string map
  // If list is empty/no pipes (e.g. 'Sin consumo de repuestos'), newSkus stays {}
  // which means all Pendiente solicitudes for this OT will be cancelled below
  var newSkus = {};
  if (strRepuestos && strRepuestos.indexOf('|') !== -1) {
    var items = strRepuestos.split(';');
    for (var x = 0; x < items.length; x++) {
      var partes = items[x].split('|');
      if (partes.length >= 2) {
        newSkus[partes[0].trim().toUpperCase()] = items[x];
      }
    }
  }

  // Read SOLICITUDES fresh — bypass cache so we see the real current state
  var dExist    = hoja.getDataRange().getValues();
  var activos   = {};  // SKUs whose Pendiente row we kept (still in the new list)
  var solChanged = false;
  for (var e = 1; e < dExist.length; e++) {
    var rowOT  = String(dExist[e][SD.OT]    || '').trim().toUpperCase();
    var rowCod = String(dExist[e][SD.CODIGO] || '').trim().toUpperCase();
    var rowEst = String(dExist[e][SD.ESTADO] || '');
    if (rowOT  !== otKey)     continue;
    if (rowEst === 'Cancelado' || rowEst === 'Despachado') continue;
    if (!newSkus[rowCod]) {
      // SKU was removed from the OT — cancel its pending solicitud
      dExist[e][SD.ESTADO] = 'Cancelado';
      solChanged = true;
    } else {
      activos[rowCod] = true;  // still present; keep it
    }
  }
  if (solChanged) hoja.getDataRange().setValues(dExist);

  // Append new SKUs that don't already have an active solicitud
  var skuKeys = Object.keys(newSkus);
  for (var k = 0; k < skuKeys.length; k++) {
    var codigo = skuKeys[k];
    if (activos[codigo]) continue;

    var itemStr   = newSkus[codigo];
    var partes2   = itemStr.split('|');
    var desc      = partes2[1].trim();
    var cantMatch = itemStr.match(/P:(\d+)/);
    var cantidad  = cantMatch ? parseInt(cantMatch[1]) : 1;
    var idSoli    = 'SD-' + new Date().getTime() + '-' + Math.floor(Math.random()*1000);

    hoja.appendRow([
      idSoli, new Date(), idOT, reseller, codigo, desc, cantidad, 0,
      'Pendiente', 'NORMAL', '', Session.getActiveUser().getEmail(), 'Auto-HUB'
    ]);
  }

  invalidateSheetValues(SCHEMA.SHEETS.SOLICITUDES);
}

// ============================================================
//  RESERVAS DE STOCK EN TRÁNSITO — HUB
// ============================================================

function obtenerCASEnTransito() {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var out   = [];
    for (var i = 1; i < datos.length; i++) {
      var f      = datos[i];
      var estado = String(f[2] || '').trim();
      if (estado !== 'En vuelo' && estado !== 'En aduana') continue;
      out.push({ cas: String(f[0] || ''), estado: estado });
    }
    return out;
  } catch(e) { return []; }
}

function crearReservaDesdeHUB(sku, descripcion, cantidad, otNombre, casRef, operador) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (!hoja) return { ok: false, msg: 'Hoja RESERVAS_STOCK no encontrada. Creá la hoja ejecutando crearHojaReservas() desde Stock Manager.' };
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var id = 'RES-' + new Date().getTime();
      var R  = SCHEMA.RESERVAS_STOCK;
      var row = new Array(11);
      row[R.ID]            = id;
      row[R.FECHA]         = new Date();
      row[R.SKU]           = String(sku).trim().toUpperCase();
      row[R.DESCRIPCION]   = String(descripcion || '');
      row[R.CANTIDAD]      = parseInt(cantidad) || 1;
      row[R.ORIGEN]        = 'TALLER';
      row[R.ID_REFERENCIA] = String(otNombre || '');
      row[R.ESTADO]        = 'Activa';
      row[R.CAS_REF]       = String(casRef || '');
      row[R.OPERADOR]      = String(operador || '');
      row[R.OBSERVACIONES] = '';
      hoja.appendRow(row);
      invalidateSheetValues(SCHEMA.SHEETS.RESERVAS);
      return { ok: true, id: id };
    } finally {
      if (lock.hasLock()) lock.releaseLock();
    }
  } catch(e) { return { ok: false, msg: e.toString() }; }
}


// ============================================================
//  FACTURACIÓN — FLUJO COMPLETO
//  1. getMailAdministracion()        → busca email en Usuarios_Internos
//  2. _obtenerDescuentoReseller()    → lee descuento por reseller (fallback 40%)
//  3. _generarXLSFactura()           → crea XLSX temporal en Drive y lo devuelve como Blob
//  4. _construirEmailFacturacion()   → arma el HTML del cuerpo
//  5. solicitarFactura(data)         → punto de entrada desde el frontend
// ============================================================

function getMailAdministracion() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.USUARIOS);
    if (!hoja) return null;
    var datos = getSheetValues(hoja);
    for (var i = 1; i < datos.length; i++) {
      var nombre = String(datos[i][0]||'').trim().toLowerCase();
      var email  = String(datos[i][1]||'').trim();
      if (nombre === 'administracion' || nombre === 'administración') {
        return (email && email.indexOf('@') !== -1) ? email : null;
      }
    }
    return null;
  } catch(e) {
    Logger.log('getMailAdministracion: ' + e);
    return null;
  }
}

function _obtenerDescuentoReseller(nombreReseller) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    if (!datos || datos.length < 2) return 0.40;
    var header = datos[0].map(function(c){ return String(c||'').toLowerCase().trim(); });
    var colDesc = -1;
    for (var j = 0; j < header.length; j++) {
      if (header[j].indexOf('descuento') !== -1 || header[j].indexOf('discount') !== -1) {
        colDesc = j; break;
      }
    }
    if (colDesc < 0) return 0.40;
    var nb = String(nombreReseller||'').trim().toLowerCase();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]||'').trim().toLowerCase() === nb) {
        var val = parseFloat(String(datos[i][colDesc]||'').replace(',','.'));
        if (!isNaN(val) && val > 0 && val <= 1) return val;
      }
    }
    return 0.40;
  } catch(e) { return 0.40; }
}

function _generarXLSFactura(ot, items, totalGeneral, data) {
  var nombre = 'FACTURA_' + ot + '_' + new Date().getTime();
  var tmpSS = SpreadsheetApp.create(nombre);
  var ws = tmpSS.getSheets()[0];
  ws.setName('Detalle');

  var headers = ['CÓDIGO', 'DESCRIPCIÓN', 'CANTIDAD', 'PVP UNITARIO (USD)', 'TOTAL ITEM (USD)'];
  ws.getRange(1, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#00a3e0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  if (items.length > 0) {
    var rows = items.map(function(item) {
      return [
        item.codigo,
        item.descripcion,
        item.cantidad,
        item.pvp > 0 ? item.pvp : '',
        item.total > 0 ? item.total : ''
      ];
    });
    var rango = ws.getRange(2, 1, rows.length, 5);
    rango.setValues(rows);
    ws.getRange(2, 4, rows.length, 2).setNumberFormat('[$$-es-AR]#,##0.00');
  }

  var filaTotal = items.length + 3;
  ws.getRange(filaTotal, 4).setValue('TOTAL GENERAL').setFontWeight('bold').setHorizontalAlignment('right');
  ws.getRange(filaTotal, 5).setValue(totalGeneral > 0 ? totalGeneral : 0)
    .setFontWeight('bold').setNumberFormat('[$$-es-AR]#,##0.00').setBackground('#e8f4f9');

  var filaInfo = filaTotal + 2;
  ws.getRange(filaInfo,     1).setValue('OT: ' + ot);
  ws.getRange(filaInfo + 1, 1).setValue('Reseller: ' + (data.reseller || ''));
  ws.getRange(filaInfo + 2, 1).setValue('Descuento aplicado: ' + Math.round(_obtenerDescuentoReseller(data.reseller) * 100) + '% (Precio Reseller)');
  ws.getRange(filaInfo + 3, 1).setValue('Generado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  ws.getRange(filaInfo, 1, 4, 1).setFontColor('#888888').setFontSize(10);

  ws.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();

  var blob = tmpSS.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  blob.setName('Factura_' + ot + '.xlsx');
  DriveApp.getFileById(tmpSS.getId()).setTrashed(true);
  return blob;
}

function _construirEmailFacturacion(data, items, totalGeneral, infoCliente, descuento) {
  var tablaHTML =
    "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px'>" +
    "<thead><tr style='background:#00a3e0'>" +
    "<th style='padding:8px 10px;color:#fff;text-align:left;border:1px solid #0088bb'>CÓDIGO</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:left;border:1px solid #0088bb'>DESCRIPCIÓN</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:center;border:1px solid #0088bb'>CANTIDAD</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:right;border:1px solid #0088bb'>PVP UNITARIO</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:right;border:1px solid #0088bb'>TOTAL ITEM</th>" +
    "</tr></thead><tbody>";

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var bg = i % 2 === 0 ? '#ffffff' : '#f7f9fc';
    tablaHTML +=
      "<tr style='background:" + bg + "'>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;font-weight:700;color:#00a3e0'>" + item.codigo + "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0'>" + item.descripcion + "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:center'>" + item.cantidad + "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>" +
        (item.pvp > 0 ? "USD " + _fmtNum(item.pvp) + " <span style='color:#888;font-size:10px'>(-" + Math.round(descuento*100) + "% desc.)</span>" : "—") +
      "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right;font-weight:700'>" +
        (item.total > 0 ? "USD " + _fmtNum(item.total) : "—") +
      "</td></tr>";
  }

  if (totalGeneral > 0) {
    tablaHTML +=
      "<tr style='background:#e8f4f9;font-weight:700'>" +
      "<td colspan='4' style='padding:8px 10px;border:1px solid #e0e0e0;text-align:right'>TOTAL GENERAL</td>" +
      "<td style='padding:8px 10px;border:1px solid #e0e0e0;text-align:right;color:#00a3e0;font-size:14px'>" +
        "USD " + _fmtNum(totalGeneral) +
      "</td></tr>";
  }
  tablaHTML += "</tbody></table>";

  var resellerInfo =
    filaDetalle('Reseller / Empresa', '<strong>' + data.reseller + '</strong>') +
    (infoCliente ? (
      filaDetalle('CUIT', infoCliente.cuit || '—') +
      filaDetalle('Dirección', (infoCliente.direccion||'—') + (infoCliente.localidad ? ', ' + infoCliente.localidad : '')) +
      filaDetalle('Código Postal', infoCliente.cp || '—') +
      filaDetalle('Teléfono', infoCliente.telefono || '—')
    ) : '');

  var cuerpo =
    bloqueCard('📋 Detalle de la Orden',
      filaDetalle('OT', '<strong>' + data.ot + '</strong>') +
      filaDetalle('Equipo / Modelo', data.equipo || '—') +
      filaDetalle('N° de Serie', data.sn || '—') +
      (data.cas ? filaDetalle('Caso DJI (CAS/FWR)', data.cas) : '') +
      filaDetalle('Garantía', 'OOW — Fuera de garantía'),
      '#00a3e0') +
    "<div style='margin-bottom:16px'>" +
      "<p style='font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase;letter-spacing:.04em'>Repuestos a facturar (" + Math.round(descuento*100) + "% dto. reseller)</p>" +
      tablaHTML +
    "</div>" +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:16px'>" +
      "<p style='font-size:11px;font-weight:700;color:#888;text-transform:uppercase;margin:10px 0 4px'>Datos del reseller</p>" +
      resellerInfo +
    "</div>" +
    bloqueCard('📎 Adjunto XLS', 'Se incluye el detalle en formato Excel (.xlsx). Generá la factura y enviala al reseller.', '#27ae60');

  return construirEmailHTML(
    'Solicitud de Facturación — OT ' + data.ot,
    'Estimado equipo de Administración,<br>Se solicita la emisión de factura por repuestos despachados en la siguiente orden:',
    cuerpo,
    'Procesá este pedido a la brevedad y emití la factura correspondiente al reseller.'
  );
}

function solicitarFactura(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var emailAdmin = getMailAdministracion();
    if (!emailAdmin) return { ok: false, msg: 'No se encontró "Administracion" en Usuarios_Internos (Columna A) con email en Columna B.' };

    var infoCliente = obtenerInfoCliente(data.reseller);
    var descuento   = _obtenerDescuentoReseller(data.reseller);

    // Leer precios desde DB_REPUESTOS col F (índice 5)
    var precioMap = {};
    var hojaDB = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaDB) {
      var dRep = getSheetValues(hojaDB);
      for (var pr = 1; pr < dRep.length; pr++) {
        var codR = String(dRep[pr][1]||'').trim().toUpperCase();
        if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5]||'0').replace(',','.')) || 0;
      }
    }

    // Parsear repuestos — se factura lo despachado (E:)
    var items = [];
    var totalGeneral = 0;
    if (data.repuestos && data.repuestos !== 'Sin consumo de repuestos') {
      var ls = data.repuestos.split(' ; ');
      for (var i = 0; i < ls.length; i++) {
        var p = ls[i].split(' | ');
        if (p.length < 3) continue;
        var cod  = p[0].trim();
        var desc = p[1].trim();
        var cant = parseInt(p[2].split('E:')[1]) || 0;
        if (cant <= 0) continue;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var pvp     = pvpBase > 0 ? pvpBase * (1 - descuento) : 0;
        var total   = pvp * cant;
        totalGeneral += total;
        items.push({ codigo: cod, descripcion: desc, cantidad: cant, pvp: pvp, total: total });
      }
    }
    if (!items.length) return { ok: false, msg: 'No hay repuestos despachados (Env. > 0) para facturar en esta OT.' };

    // Generar XLSX como Blob
    var xlsBlob = _generarXLSFactura(data.ot, items, totalGeneral, data);

    // Enviar email con adjunto
    var htmlEmail = _construirEmailFacturacion(data, items, totalGeneral, infoCliente, descuento);
    var asunto    = 'SOLICITUD DE FACTURACIÓN - OT: ' + data.ot + ' - ' + data.reseller;

    GmailApp.sendEmail(emailAdmin, asunto, '', {
      htmlBody:    htmlEmail,
      name:        CONFIG.NOMBRE_REMITENTE,
      replyTo:     CONFIG.EMAIL_SUPERVISOR,
      attachments: [xlsBlob]
    });
    registrarEmailLog(data.ot, emailAdmin, 'Administración', asunto, 'OK');

    // Actualizar estado de la OT en la hoja
    var fila    = parseInt(data.fila);
    var hojaOT  = getSheet(SCHEMA.SHEETS.OT);
    var estadoAnt = String(data.estado || '');
    hojaOT.getRange(fila, SCHEMA.OT.ESTADO      + 1).setValue('ESPERANDO FACTURA');
    hojaOT.getRange(fila, SCHEMA.OT.FECHA_ESTADO + 1).setValue(new Date());

    var operador = Session.getActiveUser().getEmail();
    registrarLog(data.ot, operador, operador, 'FACTURACIÓN', estadoAnt, 'ESPERANDO FACTURA',
                 'Solicitud de factura enviada · mail: ' + emailAdmin);

    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);

    return { ok: true, mailEnviado: emailAdmin };

  } catch(e) {
    Logger.log('solicitarFactura: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
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