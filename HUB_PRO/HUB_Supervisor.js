// @version 1.0
// ============================================================
//  HUB PRO — Command Center del supervisor: métricas, logs,
//  pendientes de envío, deuda de resellers, búsquedas/historial de
//  OTs similares, y el cruce de estado de repuestos de OT con WOS
//  (stock/ETA — helpers _hub*, backorderPred, chips Reservado_Consolidar).
//  Extraído de HUB_Código.js 2.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


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
    // Fuente: Pedidos_OTs (WOS) — el mismo dato real que alimenta el Command Center de
    // la orden, no una heurística por texto/estado de la OT. Antes esto se inferí­a de
    // f[16] (texto de repuestos) cruzado con una tabla de estados fija por circuito que
    // excluía por completo Taller (la mayoría de las OTs de este taller) — ver VERSIONS.md.
    var pedidoRowsSup = _hubPedidosOTsRows();
    var mapaRep = {};
    var vistoOtSku = {};
    for (var pi = 1; pi < pedidoRowsSup.length; pi++) {
      var fp = pedidoRowsSup[pi];
      var numeroSup = String(fp[0] || '').trim();
      if (numeroSup.toUpperCase().indexOf('OT-') !== 0) continue;
      if (String(fp[9] || '').trim() !== 'Backorder') continue;
      var otIdSup = numeroSup.substring(3);
      var cod     = String(fp[2] || '').trim().toUpperCase();
      if (!cod) continue;
      var dupKey = otIdSup + '|' + cod;
      if (vistoOtSku[dupKey]) continue;   // no consolidar duplicados del mismo OT+SKU
      vistoOtSku[dupKey] = true;
      var pend = Math.max(0, (Number(fp[4]) || 0) - (Number(fp[5]) || 0) - (Number(fp[25]) || 0));
      if (pend <= 0) continue;
      var nom = String(fp[3] || '').trim();
      if (!mapaRep[cod]) mapaRep[cod] = { codigo: cod, nombre: nom, cantOTs: 0, pendiente: 0, ots: [] };
      mapaRep[cod].cantOTs++;
      mapaRep[cod].pendiente += pend;
      mapaRep[cod].ots.push(otIdSup);
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

    // Tiempo de respuesta de repuestos: pedido a WOS → despacho.
    // Fuente: Pedidos_OTs (WOS) FECHA/FECHA_DESPACHO — antes leía la hoja shadow
    // SOLICITUDES_DESPACHO, que HUB_PRO ya no escribe (ver VERSIONS.md).
    var sumDesp = 0, cntDesp = 0;
    for (var sj = 1; sj < pedidoRowsSup.length; sj++) {
      var fj = pedidoRowsSup[sj];
      if (String(fj[0] || '').trim().toUpperCase().indexOf('OT-') !== 0) continue;
      var fSol  = fj[10] instanceof Date ? fj[10] : null;
      var fDisp = fj[14] instanceof Date ? fj[14] : null;
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
//  COMMAND CENTER DE REPUESTOS — helpers de dominio
//  Cruzan, en una sola lectura por hoja (sin N+1), el estado real de un repuesto de OT:
//  · Pedidos_OTs (WOS)      → pedido a WOS, su estado, fecha, operario.
//  · RESERVAS_STOCK (MASTER)→ reserva directa "de importación" armada desde HUB_PRO.
//  · COMPRAS_DJI+DETALLE    → ETA del lote DJI más próximo, si el pedido está en Backorder.
//  · Carmen STOCK           → stock actual real (ver fix en obtenerDetalleOT).
// ============================================================

function _hubCarmenStockMap() {              // { SKU: stockActual } — Carmen col A(código)/C(stock)
  var d = _hubExtSheetValues(CARMEN_SS_ID, 'STOCK', 60);
  var m = {};
  for (var i = 1; i < d.length; i++) {
    var c = String(d[i][0] || '').trim().toUpperCase();
    if (c) m[c] = parseInt(d[i][2]) || 0;
  }
  return m;
}


function _hubPedidosOTsRows() {
  return _hubExtSheetValues(WOS_SS_ID, WOS_HOJA_PED, 30);
}


function _hubReservasStockRows() {
  return _hubExtSheetValues(MASTER_SHEET_ID, SCHEMA.SHEETS.RESERVAS, 30);
}


// Parsea "dd/MM/yyyy" o "dd/MM" (año actual) → Date; Date → Date; cualquier otra cosa → null.
// Réplica de _wosEtaToDate (WOS/Despacho_Código.js) — mismo formato de FECHA_ETA en COMPRAS_DETALLE.
function _hubEtaToDate(s) {
  if (s instanceof Date) return isNaN(s.getTime()) ? null : s;
  s = String(s || '').trim();
  if (!s) return null;
  var m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (!m) return null;
  var dd = parseInt(m[1], 10), mm = parseInt(m[2], 10) - 1;
  var yy = m[3] ? parseInt(m[3], 10) : (new Date()).getFullYear();
  if (yy < 100) yy += 2000;
  var dt = new Date(yy, mm, dd);
  return isNaN(dt.getTime()) ? null : dt;
}


// Réplica simplificada de WOS_getEnCaminoMap (WOS/Despacho_Código.js): solo la ETA más
// próxima por SKU entre CAS activos — no necesitamos acá la lógica de reservas-netas de WOS.
// { SKU: 'dd/MM/yyyy' }
function _hubEnCaminoEtaPorSku() {
  var dCAS = _hubExtSheetValues(MASTER_SHEET_ID, 'COMPRAS_DJI', 60);
  var dDET = _hubExtSheetValues(MASTER_SHEET_ID, 'COMPRAS_DETALLE', 60);
  var casActivos = {};
  for (var c = 1; c < dCAS.length; c++) {
    var casId  = String(dCAS[c][0] || '').trim().toUpperCase();
    var estado = String(dCAS[c][2] || '').trim();
    if (casId && estado !== 'En depósito' && estado.indexOf('Borrador') < 0) casActivos[casId] = true;
  }
  var out = {}, bestDt = {};
  for (var d = 1; d < dDET.length; d++) {
    var dCas = String(dDET[d][0] || '').trim().toUpperCase();
    var dSku = String(dDET[d][1] || '').trim().toUpperCase();
    var pend = (Number(dDET[d][3]) || 0) - (Number(dDET[d][4]) || 0);
    if (!dSku || !casActivos[dCas] || pend <= 0) continue;
    var dt = _hubEtaToDate(dDET[d][6]);
    if (dt && (!bestDt[dSku] || dt < bestDt[dSku])) {
      bestDt[dSku] = dt;
      out[dSku] = Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    }
  }
  return out;
}


// ── Detalle de UNA OT (se llama al abrir la orden) ───────────────────────────────────
// otRaw: número crudo de la OT (actual.ot, ej. "WH/REP/00123", SIN el prefijo "OT-").
// skus: array de códigos ya parseados en el cliente desde det.repuestos.
function obtenerEstadoRepuestosOT(otRaw, skus) {
  try {
    var numero     = 'OT-' + String(otRaw || '').trim();
    var otRawTrim  = String(otRaw || '').trim();
    var pedidoRows = _hubPedidosOTsRows();
    var resRows    = _hubReservasStockRows();
    var stockMap   = _hubCarmenStockMap();
    var etaMap     = _hubEnCaminoEtaPorSku();
    var R = SCHEMA.RESERVAS_STOCK;
    var out = {};

    for (var s = 0; s < (skus || []).length; s++) {
      var sku = String(skus[s]).trim().toUpperCase();
      if (!sku || out[sku]) continue;
      var item = { pedidoWOS: null, reserva: null, stockActual: stockMap.hasOwnProperty(sku) ? stockMap[sku] : null };

      for (var p = 1; p < pedidoRows.length; p++) {
        if (String(pedidoRows[p][0] || '').trim() !== numero) continue;
        if (String(pedidoRows[p][2] || '').trim().toUpperCase() !== sku) continue;
        var est = String(pedidoRows[p][9] || '').trim();
        item.pedidoWOS = {
          numero:   numero,
          estado:   est,
          fecha:    (pedidoRows[p][10] instanceof Date) ? pedidoRows[p][10].getTime() : null,
          operario: String(pedidoRows[p][23] || ''),
          cantPend: Math.max(0, (Number(pedidoRows[p][4]) || 0) - (Number(pedidoRows[p][5]) || 0) - (Number(pedidoRows[p][25]) || 0)),
          eta:      (est === 'Backorder' && etaMap[sku]) ? etaMap[sku] : ''
        };
        break;   // 1ª fila que matchea numero+sku (ver Riesgos: duplicados del mismo SKU no se consolidan)
      }

      for (var r = 1; r < resRows.length; r++) {
        if (String(resRows[r][R.ID_REFERENCIA] || '').trim() !== otRawTrim) continue;
        if (String(resRows[r][R.SKU] || '').trim().toUpperCase() !== sku) continue;
        if (String(resRows[r][R.ESTADO] || '').trim() !== 'Activa') continue;
        item.reserva = {
          id:       String(resRows[r][R.ID] || ''),
          casRef:   String(resRows[r][R.CAS_REF] || ''),
          fecha:    (resRows[r][R.FECHA] instanceof Date) ? resRows[r][R.FECHA].getTime() : null,
          operador: String(resRows[r][R.OPERADOR] || '')
        };
        break;
      }
      out[sku] = item;
    }
    return { ok: true, items: out };
  } catch(e) {
    Logger.log('obtenerEstadoRepuestosOT ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}


// ── Resumen BULK de TODAS las OTs abiertas (se llama al cargar la lista) ─────────────
// Una sola pasada por Pedidos_OTs y RESERVAS_STOCK — NO abre "Ordenes de trabajo" (el
// cliente ya tiene c.repuestos de cargarOrdenes()). otId sin el prefijo "OT-".
function obtenerEstadoRepuestosOTs() {
  try {
    var pedidoRows = _hubPedidosOTsRows();
    var resRows    = _hubReservasStockRows();
    var stockMap   = _hubCarmenStockMap();
    var R = SCHEMA.RESERVAS_STOCK;

    var pedidosPorOT = {};   // { otId: { SKU: {estado, cantPend} } }
    for (var p = 1; p < pedidoRows.length; p++) {
      var numero = String(pedidoRows[p][0] || '').trim();
      if (numero.toUpperCase().indexOf('OT-') !== 0) continue;
      var otId = numero.substring(3);
      var sku  = String(pedidoRows[p][2] || '').trim().toUpperCase();
      if (!sku) continue;
      if (!pedidosPorOT[otId]) pedidosPorOT[otId] = {};
      pedidosPorOT[otId][sku] = {
        estado:   String(pedidoRows[p][9] || '').trim(),
        cantPend: Math.max(0, (Number(pedidoRows[p][4]) || 0) - (Number(pedidoRows[p][5]) || 0) - (Number(pedidoRows[p][25]) || 0))
      };
    }

    var reservasPorOT = {};  // { otId: { SKU: true } } — solo Activas
    for (var r = 1; r < resRows.length; r++) {
      if (String(resRows[r][R.ESTADO] || '').trim() !== 'Activa') continue;
      var otRef = String(resRows[r][R.ID_REFERENCIA] || '').trim();
      var skuR  = String(resRows[r][R.SKU] || '').trim().toUpperCase();
      if (!otRef || !skuR) continue;
      if (!reservasPorOT[otRef]) reservasPorOT[otRef] = {};
      reservasPorOT[otRef][skuR] = true;
    }

    return { ok: true, pedidos: pedidosPorOT, reservas: reservasPorOT, stockMap: stockMap };
  } catch(e) {
    Logger.log('obtenerEstadoRepuestosOTs ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}
