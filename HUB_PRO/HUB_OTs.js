// @version 1.0
// ============================================================
//  HUB PRO — Órdenes de trabajo: CRUD/listado, catálogo, pedido de
//  repuestos para una OT, validación de duplicados CAS/FWRC/SN.
//  Extraído de HUB_Código.js 2.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


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



// ============================================================
//  CARGA UNIFICADA — una sola llamada desde el frontend
//  Devuelve: ordenes + repuestos + tecnicos + resellers
// ============================================================
// Estados "cerrados"/terminales (una OT ya terminada). Debe coincidir con el front.
var _ESTADOS_CERRADOS = { 'Finalizado':1, 'Entregado':1, 'CANCELADO':1, 'Rechazado DJI':1, 'Sin respuesta · Cerrado':1 };
function _esCerrada(estado) { return _ESTADOS_CERRADOS[String(estado||'').trim()] === 1; }


// Origen del repuesto (quién lo pone): col AA si tiene valor; si no, se deriva del texto
// que el Portal deja al inicio del informe técnico (col M) en casos viejos.
//   "Stock reseller" = el reseller reparó/repara con su stock (pide reposición)
//   "Adelantado"     = nosotros adelantamos el repuesto
function _origenRepuestoDe(f) {
  var v = String(f[SCHEMA.OT.ORIGEN_REPUESTO] || "").trim();
  if (v) return v;
  var informe = String(f[SCHEMA.OT.TRABAJO] || "").toUpperCase();
  if (informe.indexOf("YA REPARADO") !== -1) return "Stock reseller";
  if (informe.indexOf("PENDIENTE") !== -1 && informe.indexOf("NECESITA REPUESTOS") !== -1) return "Adelantado";
  return "";
}


// Refresco liviano: solo la lista de órdenes (viva), sin el catálogo estático.
// incluirCerradas=false (default) → solo OTs activas + cerradasCount (para no mandar
// cientos de OTs cerradas en cada refresh). incluirCerradas=true → todas (vista Finalizados).
function cargarOrdenes(incluirCerradas) { return cargarTodo(true, incluirCerradas); }


function cargarTodo(soloOrdenes, incluirCerradas) {
  try {
    var hoy     = new Date();
    var hojaOT  = getSheet(SCHEMA.SHEETS.OT);
    // force=true: la lista SIEMPRE lee la hoja OT en vivo. Si se leyera del cache,
    // tras cambiar un estado se veía el valor viejo (CacheService.remove tiene
    // propagación diferida entre ejecuciones y el refresh inmediato ganaba la carrera → obligaba a F5).
    var datosOT = getSheetValues(hojaOT, true);

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

    var ordenes = [], tecSet = {}, cerradasCount = 0;
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

      // B (performance): no mandar las OTs cerradas salvo que se pidan explícitamente
      if (_esCerrada(f[4])) { cerradasCount++; if (!incluirCerradas) continue; }

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
        fwrc:      String(f[SCHEMA.OT.FWRC]||"").trim(),
        repuestos: f[16] ? String(f[16]).replace(/\r/g, "").trim() : "Sin consumo de repuestos",
        manoObraGuardada: f[22] ? String(f[22]).trim() : "",
        notasInternas:   f[23] ? String(f[23]).trim() : "",
        mensajes:   String(f[11]||"").trim(),
        msgNoLeido: msgNoLeido,
        fechaIngreso: (f[SCHEMA.OT.FECHA_INGRESO] instanceof Date) ? f[SCHEMA.OT.FECHA_INGRESO].getTime() : null,
        prioridad: String(f[17]).toUpperCase() === "URGENTE",
        circuito:  tipo,
        origenRepuesto: _origenRepuestoDe(f),               // AA: quién pone el repuesto (badge)
        cierreTipo:     String(f[SCHEMA.OT.CIERRE_TIPO]||"").trim(), // AB: reposición vs NC
        esBateria: mapaBaterias[String(f[5]||'').trim().toLowerCase()] === true,
        dias:      dias,
        diasEstado: diasEstado,
        umMod:     (rawUM instanceof Date) ? rawUM.getTime() : 0
      });
    }

    var tecnicosDisp = obtenerTecnicosDisponibles();

    // soloOrdenes: refresco liviano — devuelve la lista viva SIN el catálogo
    // (el catálogo es estático, se trae una sola vez y se cachea en el cliente).
    if (soloOrdenes) {
      return {
        ordenes:             ordenes,
        tecnicos:            Object.keys(tecSet).sort(),
        tecnicosDisponibles: tecnicosDisp,
        cerradasCount:       cerradasCount
      };
    }

    var catalogo = cargarCatalogo();
    return {
      ordenes:             ordenes,
      tecnicos:            Object.keys(tecSet).sort(),
      tecnicosDisponibles: tecnicosDisp,
      cerradasCount:       cerradasCount,
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
    if (isNaN(filaNum) || filaNum < 2) return { trabajo: "", repuestos: "", mensajes: "" };
    var datos = getSheetValues(SCHEMA.SHEETS.OT, true);  // force=true: siempre leer del sheet, nunca del cache
    var f = datos[filaNum - 1];
    if (!f) return { trabajo: "", repuestos: "", mensajes: "" };

    var histRaw = String(f[SCHEMA.OT.HISTORIAL_ESTADOS]||"").trim();
    var historial = [];
    try { if (histRaw) historial = JSON.parse(histRaw); } catch(eh) {}

    var rawUM = f[SCHEMA.OT.ULTIMA_MODIFICACION];

    return {
      trabajo:             String(f[12]||""),
      repuestos:           String(f[16]||""),
      mensajes:            String(f[11]||""),
      historialEstados:    historial,
      ultimaModificacion:  (rawUM instanceof Date) ? rawUM.getTime() : 0
    };
  } catch(e) {
    Logger.log("obtenerDetalleOT ERROR fila=" + fila + " : " + e);
    return { trabajo: "", repuestos: "", mensajes: "" };
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

    // getRange NO auto-expande columnas: asegurar que existan AA/AB antes de escribirlas.
    var _maxCol = hoja.getMaxColumns();
    if (_maxCol < SCHEMA.OT.CIERRE_TIPO + 1) hoja.insertColumnsAfter(_maxCol, (SCHEMA.OT.CIERRE_TIPO + 1) - _maxCol);

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

    // Unicidad de CAS / FWRC contra las DEMÁS órdenes (excluye la actual por su nº de OT).
    var _dupU = _otBuscarDuplicadoCasFwrc(data.cas, data.fwrc, data.ot);
    if (_dupU) {
      return { resultado: 'El ' + _dupU.campo + ' ' + _dupU.valor + ' ya está usado en la orden ' + _dupU.ot + '.', ot: data.ot || '' };
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
    if (data.fwrc !== undefined) hoja.getRange(fila, SCHEMA.OT.FWRC + 1).setValue(data.fwrc || "");
    hoja.getRange(fila, SCHEMA.OT.REPUESTOS        + 1).setValue(data.repuestos);
    hoja.getRange(fila, SCHEMA.OT.PRIORIDAD        + 1).setValue(data.prioridad ? "URGENTE" : "NORMAL");
    hoja.getRange(fila, SCHEMA.OT.CIRCUITO         + 1).setValue(data.circuito);
    // Origen del repuesto (AA) y tipo de cierre (AB) — informativos, editables en HUB
    if (data.origenRepuesto !== undefined) hoja.getRange(fila, SCHEMA.OT.ORIGEN_REPUESTO + 1).setValue(data.origenRepuesto || "");
    if (data.cierreTipo     !== undefined) hoja.getRange(fila, SCHEMA.OT.CIERRE_TIPO     + 1).setValue(data.cierreTipo || "");
    if (data.manoObraGuardada !== undefined) {
      hoja.getRange(fila, SCHEMA.OT.MANO_OBRA      + 1).setValue(data.manoObraGuardada);
      if (data.notasInternas !== undefined) {
        hoja.getRange(fila, SCHEMA.OT.NOTAS_INTERNAS + 1).setValue(data.notasInternas || "");
      }
    }

    if (data.estado === "Finalizado") {
      hoja.getRange(fila, SCHEMA.OT.FECHA_CIERRE + 1).setValue(new Date());
      // HUB ya no descuenta stock al finalizar: la baja de repuestos se registra fuera de HUB.
      // (Antes: procesarSalidaRepuestos escribía STOCK_REPUESTOS + MOVIMIENTOS_STOCK — removido.)
    } else if (data.estado === "Entregado") {
      // Taller central: el equipo fue retirado o despachado — cierra la OT
      if (!old[SCHEMA.OT.FECHA_CIERRE]) hoja.getRange(fila, SCHEMA.OT.FECHA_CIERRE + 1).setValue(new Date());
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
//  2. Registra el evento en LOGS.
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

    // ── 2. Registrar en LOGS ──────────────────────────────────
    var hojaLog = getSheet(SCHEMA.SHEETS.LOGS);
    if (hojaLog) {
      hojaLog.appendRow([
        new Date(), otBusc, '', '', '',
        estadoAnterior, 'CANCELADO',
        'Caso cancelado',
        Session.getActiveUser().getEmail()
      ]);
      invalidateSheetValues(SCHEMA.SHEETS.LOGS);
    }

    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);

    return { ok: true, ot: otBusc };
  } catch(e) {
    Logger.log('cancelarCaso: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}


// ============================================================
//  CREAR NUEVA OT (CON CANDADO DE CONCURRENCIA)
// ============================================================
// Normaliza un CAS/FWRC SOLO para comparar unicidad: minúsculas y sin guiones ni espacios.
// Así "AB-123", "ab 123" y "AB123" cuentan como el mismo número (para que no se pasen de vivos
// metiendo un guion o cambiando mayúsculas). El valor se GUARDA tal cual lo escribió el operador;
// esta normalización es solo para el chequeo de duplicados.
function _normCasFwrc(v) {
  return String(v == null ? "" : v).toLowerCase().replace(/[\s\-]/g, "");
}


// Unicidad de CAS / FWRC: devuelve { campo, valor, ot } del primer conflicto, o null.
// cas/fwrc vacíos se ignoran (son opcionales). otExcluir = nº de OT a saltear (al editar,
// para no chocar consigo misma). Como ahora CAS/FWRC es UN solo caso DJI (uno u otro), el
// valor se compara contra AMBAS columnas (CAS col O y FWRC col P) de las demás OTs, así no
// puede colisionar ni con un CAS ni con un FWRC viejo. Normalizado (sin guiones/espacios,
// case-insensitive). Lee del sheet (force), no del cache.
function _otBuscarDuplicadoCasFwrc(cas, fwrc, otExcluir) {
  var casN  = _normCasFwrc(cas);
  var fwrcN = _normCasFwrc(fwrc);
  if (!casN && !fwrcN) return null;
  var O = SCHEMA.OT;
  var datos = getSheetValues(SCHEMA.SHEETS.OT, true);
  var excl  = String(otExcluir || "").trim();
  for (var i = 1; i < datos.length; i++) {
    var otRow = String(datos[i][O.OT] || "").trim();
    if (!otRow || (excl && otRow === excl)) continue;
    var colCas = _normCasFwrc(datos[i][O.CAS]), colFwrc = _normCasFwrc(datos[i][O.FWRC]);
    if (casN  && (colCas === casN  || colFwrc === casN))  return { campo: "CAS/FWRC", valor: String(cas).trim(),  ot: otRow };
    if (fwrcN && (colFwrc === fwrcN || colCas === fwrcN)) return { campo: "CAS/FWRC", valor: String(fwrc).trim(), ot: otRow };
  }
  return null;
}


// Unicidad de N° de serie: no puede haber DOS casos ABIERTOS con el mismo S/N.
// Cubre también los casos abiertos desde el Portal Reseller, porque ambos sistemas
// escriben en la MISMA hoja 'Ordenes de trabajo'. "Abierto" = estado NO cerrado
// (usa _esCerrada, misma definición que el front y que el resto del HUB), así una OT
// ya cerrada (Finalizado/Entregado/CANCELADO/Rechazado DJI/Sin respuesta·Cerrado) no
// bloquea abrir un caso nuevo para el mismo equipo. S/N vacío se ignora. Lee del sheet
// (force). otExcluir = nº de OT a saltear (para no chocar consigo misma).
// Devuelve { sn, ot, estado } del primer conflicto, o null.
function _otBuscarSnAbierto(sn, otExcluir) {
  var snN = String(sn == null ? "" : sn).trim().toUpperCase();
  if (!snN) return null;
  var O = SCHEMA.OT;
  var datos = getSheetValues(SCHEMA.SHEETS.OT, true);
  var excl  = String(otExcluir || "").trim();
  for (var i = 1; i < datos.length; i++) {
    var otRow = String(datos[i][O.OT] || "").trim();
    if (!otRow || (excl && otRow === excl)) continue;
    if (String(datos[i][O.SN] || "").trim().toUpperCase() !== snN) continue;
    var estadoFila = String(datos[i][O.ESTADO] || "").trim();
    if (!_esCerrada(estadoFila)) return { sn: snN, ot: otRow, estado: estadoFila };
  }
  return null;
}


function crearNuevaOT(datos) {
  datos = datos || {};
  var lock = LockService.getScriptLock();
  var nOT  = "";
  try {
    lock.waitLock(10000);

    // Unicidad: no permitir dos OTs con el mismo CAS o FWRC (con el lock tomado → sin carreras).
    var _dup = _otBuscarDuplicadoCasFwrc(datos.cas, datos.fwrc, null);
    if (_dup) return { resultado: "Error: El " + _dup.campo + " " + _dup.valor + " ya está usado en la orden " + _dup.ot + ".", ot: "", duplicado: _dup };

    // Unicidad: no puede haber dos casos ABIERTOS con el mismo N° de serie (incluye los
    // abiertos desde el Portal Reseller, misma hoja). Solo bloquea contra OTs no cerradas.
    var _dupSn = _otBuscarSnAbierto(datos.sn, null);
    if (_dupSn) return { resultado: "Error: Ya existe la orden " + _dupSn.ot + " abierta (" + _dupSn.estado + ") con el N° de serie " + _dupSn.sn + ". No puede haber dos casos abiertos para el mismo equipo.", ot: "", duplicado: _dupSn };

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
    row[14] = datos.cas       || "";   // O: CAS (opcional, único)
    row[15] = datos.fwrc      || "";   // P: FWRC (opcional, único)
    row[17] = datos.prioridad || "NORMAL";
    row[18] = datos.circuito  || "Taller";
    row[20] = new Date();
    hoja.appendRow(row);

    registrarLog(nOT, datos.reseller || "Sin asignar", Session.getActiveUser().getEmail(), "CREACIÓN", "-", "Abierto", "Nueva orden");

    // Notificar la apertura — igual criterio que actualizarOrden: "Abierto" está en
    // ESTADOS_NOTIFICAR_RESELLER/TÉCNICO. Sin esto, las OT creadas desde el HUB (no desde
    // el Portal Reseller) no avisaban al reseller/técnico ni dejaban registro en EMAIL_LOGS.
    try {
      var dataNotif = {
        ot:        nOT,
        estado:    "Abierto",
        reseller:  datos.reseller || "",
        equipo:    datos.equipo   || "",
        sn:        datos.sn       || "",
        garantia:  datos.garantia || "OOW",
        circuito:  datos.circuito || "Taller",
        cas:       datos.cas      || "",
        fwrc:      datos.fwrc     || "",
        trabajo:   datos.trabajo  || "",
        repuestos: "",
        cliente:   "",
        prioridad: (String(datos.prioridad || "").toUpperCase() === "URGENTE")
      };
      dataNotif._fechaEstimada = calcularFechaEstimada(dataNotif.circuito, dataNotif.garantia, row[0]);
      enviarNotificaciones(dataNotif, "-", datos.tecnico || "");
    } catch(eN) { Logger.log("crearNuevaOT notif: " + eN); }

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
