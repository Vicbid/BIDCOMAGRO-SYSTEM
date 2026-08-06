// @version 1.6
// ============================================================
//  COMANDAS — Reporte de tiempos + herramientas CP_debug*/CP_diag*
//  (diagnóstico genérico, no atado a un incidente puntual).
//  Extraído de CP_Código.js 1.43 el 2026-07-30 — reorganización sin
//  cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


/* ════════════════════════════════════════════════════════════
   REPORTE DE TIEMPOS — 4 tramos, de punta a punta:
     1) CARGAR → comanda cargada   (el operador registra el N° de comanda)
     2) comanda cargada → autorizada (Comandas Master le asigna guía — proxy:
        el mail #1 "autorizado" al reseller, que sale apenas se detecta)
     3) autorizada → despachada   (Comandas Master marca DESPACHADO — proxy:
        el mail #2 "despachado" al reseller)
     4) TOTAL: CARGAR → despachada (punta a punta, ciclo completo)
   Los tramos 2 y 3 tienen hasta 30 min de rezago (cadencia del cron
   CP_autoMailEnvios), igual que ya tenía el tramo 1 con origen 'deteccion'.
   origen 'edit' = exacto | 'deteccion' = aproximado.
════════════════════════════════════════════════════════════ */

// Mapa liviano leyendo Ventas una sola vez: { IDVENTA: {idVenta, idNegocio, reseller, cliente} }
function _cpInfoVentasMap() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) return {};
  var col = det.col, m = {};
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var idv = _s(col.idVenta > -1 ? rows[i][col.idVenta] : ''); if (!idv) continue;
    var k = idv.toUpperCase();
    if (!m[k]) m[k] = {
      idVenta:   idv,
      idNegocio: _s(col.idNegocio > -1 ? rows[i][col.idNegocio] : ''),
      reseller:  _s(col.reseller > -1 ? rows[i][col.reseller] : ''),
      cliente:   _s(col.razonSocial > -1 ? rows[i][col.razonSocial] : '')
    };
  }
  return m;
}


// Estadística simple de un array de números: n, promedio, mediana, min, max.
function _cpStats(arr) {
  if (!arr || !arr.length) return { n: 0, prom: 0, mediana: 0, min: 0, max: 0 };
  var s = arr.slice().sort(function(a, b) { return a - b; });
  var n = s.length, sum = s.reduce(function(a, b) { return a + b; }, 0), mid = Math.floor(n / 2);
  var med = (n % 2) ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  var r1 = function(x) { return Math.round(x * 10) / 10; };
  return { n: n, prom: r1(sum / n), mediana: r1(med), min: r1(s[0]), max: r1(s[n - 1]) };
}


// Diferencia en horas (1 decimal) entre 2 timestamps (ms), o null si negativa/faltante.
function _cpHoras(msIni, msFin) {
  if (msIni == null || msFin == null) return null;
  var h = Math.round(((msFin - msIni) / 3600000) * 10) / 10;
  return h >= 0 ? h : null;
}


// Lunes 00:00 (hora local del script) de la semana que contiene ts — clave para agrupar la
// tendencia semana a semana. Semana = lunes a domingo.
function _cpSemanaKey(ts) {
  var d = new Date(ts);
  var dia = d.getDay(); // 0=domingo..6=sábado
  var diffALunes = (dia === 0 ? -6 : 1) - dia;
  var lunes = new Date(d.getFullYear(), d.getMonth(), d.getDate() + diffALunes);
  return lunes.getTime();
}


// Cálculo compartido por CP_reporteTiempos() (UI/CSV) y CP_actualizarHojaTiempos() (hoja TIEMPOS)
// — una sola fuente de verdad para no tener 2 versiones de la misma cuenta que puedan divergir.
// Devuelve { rows:[por envío], pendientes:[marcadas CARGAR sin envío], resumen }.
function _cpTiemposCalc() {
  var logMap = _cpLogMap();
  // CARGAR más antiguo por venta (+ si es exacto)
  var cargarPorVenta = {};
  Object.keys(logMap).forEach(function(key) {
    var idv = key.split('||')[0], reg = logMap[key], cur = cargarPorVenta[idv];
    if (!cur || reg.ts < cur.ts) cargarPorVenta[idv] = { ts: reg.ts, origen: reg.origen };
  });
  var info = _cpInfoVentasMap();
  var enviosMap = _cpEnviosMap();
  var rows = [], dPrim = [], dTodos = [], dAut = [], dDesp = [], dTotal = [];
  Object.keys(enviosMap).forEach(function(idvU) {
    var arr = enviosMap[idvU], carg = cargarPorVenta[idvU];
    var cargTs = carg ? carg.ts.getTime() : null, exacto = carg ? (carg.origen === 'edit') : false;
    var iv = info[idvU] || {};
    arr.forEach(function(e) {
      if (!e.fechaTs) return;
      // Corte de "medición limpia" pedido por el usuario: si la autorización es anterior a
      // CP_TIEMPOS_DESDE, se excluye el envío ENTERO del reporte (ver comentario en Env.js) —
      // los nombres de estado en Comandas Master cambiaron con el tiempo y ensuciaban lo viejo.
      if (e.mailAutorizadoTs != null && e.mailAutorizadoTs < CP_TIEMPOS_DESDE.getTime()) return;
      var hCarga = _cpHoras(cargTs, e.fechaTs);
      if (hCarga != null) { dTodos.push(hCarga); if (e.envio === 1) dPrim.push(hCarga); }
      var hAut = _cpHoras(e.fechaTs, e.mailAutorizadoTs);
      if (hAut != null) dAut.push(hAut);
      var hDesp = _cpHoras(e.mailAutorizadoTs, e.mailResellerTs);
      if (hDesp != null) dDesp.push(hDesp);
      var hTotal = _cpHoras(cargTs, e.mailResellerTs);
      if (hTotal != null) dTotal.push(hTotal);
      rows.push({
        idVenta: iv.idVenta || idvU, idNegocio: iv.idNegocio || '', reseller: iv.reseller || '', cliente: iv.cliente || '',
        envio: e.envio, comanda: e.comanda, operador: e.operador || '',
        cargarTs: cargTs, cargarStr: cargTs != null ? _fmtTs(new Date(cargTs)) : '',
        origen: carg ? (exacto ? 'exacto' : 'aprox') : 'sin registro',
        cargadaTs: e.fechaTs, cargadaStr: e.fechaStr || '',
        autorizadoTs: e.mailAutorizadoTs || null, autorizadoStr: e.mailAutorizadoTs ? _fmtTs(new Date(e.mailAutorizadoTs)) : '',
        despachadoTs: e.mailResellerTs || null, despachadoStr: e.mailResellerTs ? _fmtTs(new Date(e.mailResellerTs)) : '',
        horasCarga: hCarga, horasAutorizar: hAut, horasDespachar: hDesp, horasTotal: hTotal,
        primero: (e.envio === 1)
      });
    });
  });
  rows.sort(function(a, b) { return a.idVenta === b.idVenta ? a.envio - b.envio : (a.idVenta < b.idVenta ? -1 : 1); });
  // pendientes: marcadas CARGAR pero sin ningún envío todavía
  var now = Date.now(), pend = [];
  Object.keys(cargarPorVenta).forEach(function(idvU) {
    if (enviosMap[idvU]) return;
    var carg = cargarPorVenta[idvU], iv = info[idvU] || {};
    pend.push({ idVenta: iv.idVenta || idvU, idNegocio: iv.idNegocio || '', reseller: iv.reseller || '',
      cargarTs: carg.ts.getTime(),
      horasAbierto: _cpHoras(carg.ts.getTime(), now),
      origen: (carg.origen === 'edit' ? 'exacto' : 'aprox') });
  });
  pend.sort(function(a, b) { return b.horasAbierto - a.horasAbierto; });
  // Top 5 comandas más lentas de punta a punta — para que el jefe pueda ir directo a preguntar
  // por casos concretos, no solo mirar un promedio general.
  var peores = rows.filter(function(x) { return x.horasTotal != null; })
    .slice().sort(function(a, b) { return b.horasTotal - a.horasTotal; })
    .slice(0, 5)
    .map(function(x) {
      return { idVenta: x.idVenta, idNegocio: x.idNegocio, reseller: x.reseller, comanda: x.comanda,
        horasCarga: x.horasCarga, horasAutorizar: x.horasAutorizar, horasDespachar: x.horasDespachar, horasTotal: x.horasTotal };
    });
  // Tendencia semana a semana (lunes a domingo, ancladas en la fecha de CARGAR de cada venta)
  // por etapa — para ver si mejora o empeora, no solo el promedio general acumulado. Últimas 10
  // semanas con datos (una semana entra si tiene AL MENOS una etapa con dato).
  var semanas = {};
  rows.forEach(function(x) {
    if (x.cargarTs == null) return;
    var wk = _cpSemanaKey(x.cargarTs);
    if (!semanas[wk]) semanas[wk] = { carga: [], autorizar: [], despachar: [] };
    if (x.horasCarga != null) semanas[wk].carga.push(x.horasCarga);
    if (x.horasAutorizar != null) semanas[wk].autorizar.push(x.horasAutorizar);
    if (x.horasDespachar != null) semanas[wk].despachar.push(x.horasDespachar);
  });
  var _prom1 = function(arr) { return arr.length ? Math.round((arr.reduce(function(a, b) { return a + b; }, 0) / arr.length) * 10) / 10 : null; };
  var tendencia = Object.keys(semanas).map(Number).sort(function(a, b) { return a - b; })
    .map(function(wk) {
      var s = semanas[wk];
      return {
        semanaTs: wk, semanaStr: Utilities.formatDate(new Date(wk), 'America/Argentina/Buenos_Aires', 'dd/MM'),
        carga: _prom1(s.carga), autorizar: _prom1(s.autorizar), despachar: _prom1(s.despachar),
        nCarga: s.carga.length, nAutorizar: s.autorizar.length, nDespachar: s.despachar.length
      };
    })
    .slice(-10);
  return {
    rows: rows, pendientes: pend, peores: peores, tendencia: tendencia,
    resumen: { primer: _cpStats(dPrim), todos: _cpStats(dTodos), autorizar: _cpStats(dAut), despachar: _cpStats(dDesp), total: _cpStats(dTotal) }
  };
}


// Devuelve { ok, rows:[por envío], pendientes:[marcadas CARGAR sin envío], peores:[top 5 más lentas],
//            tendencia:[por semana], desde:'dd/MM/yyyy', resumen:{primer, todos, autorizar, despachar, total} }.
function CP_reporteTiempos() {
  try {
    var c = _cpTiemposCalc();
    var desde = Utilities.formatDate(CP_TIEMPOS_DESDE, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
    return { ok: true, rows: c.rows, pendientes: c.pendientes, peores: c.peores, tendencia: c.tendencia, desde: desde, resumen: c.resumen };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}


/* ════════════════════════════════════════════════════════════
   HOJA "TIEMPOS" — el mismo cálculo de arriba, pero volcado directo
   a una pestaña del spreadsheet de log (1 fila por envío, + 1 fila
   por venta marcada CARGAR que todavía no tiene comanda cargada).
   Pensada para que el operador la mire/filtre/ordene directo en
   Sheets sin pasar por la app. Se reescribe entera cada vez (full
   refresh, no upsert) — determinístico y sin riesgo de filas
   duplicadas o desincronizadas. Se llama sola: al crear un envío
   (CP_crearEnvio) y en cada corrida del cron de mails (CP_autoMailEnvios,
   cada 30 min) — así queda al día sin que el operador tenga que hacer nada,
   más un botón "Actualizar" en la UI por si quiere forzarlo al toque.
════════════════════════════════════════════════════════════ */

var CP_TIEMPOS_HEADERS = [
  'Estado', 'ID Venta', 'N° Negocio', 'Reseller', 'Razón Social', 'Envío', 'Comanda', 'Operador',
  'Marcado CARGAR', 'Comanda cargada', 'Tiempo en cargar (hs)',
  'Autorizado', 'Tiempo en autorizar (hs)',
  'Despachado', 'Tiempo en despachar (hs)',
  'TIEMPO TOTAL — Cargar→Despacho (hs)'
];
var CP_TIEMPOS_COL = {
  ESTADO: 1, ID_VENTA: 2, ID_NEGOCIO: 3, RESELLER: 4, RAZON_SOCIAL: 5, ENVIO: 6, COMANDA: 7, OPERADOR: 8,
  CARGAR: 9, CARGADA: 10, H_CARGA: 11, AUTORIZADO: 12, H_AUTORIZAR: 13, DESPACHADO: 14, H_DESPACHAR: 15, H_TOTAL: 16
};

function _cpTiemposHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_TIEMPOS_TAB);
  if (!h) {
    h = ss.insertSheet(CP_TIEMPOS_TAB);
    h.getRange(1, 1, 1, CP_TIEMPOS_HEADERS.length).setValues([CP_TIEMPOS_HEADERS]);
    h.setFrozenRows(1);
    h.getRange(1, 1, 1, CP_TIEMPOS_HEADERS.length).setFontWeight('bold').setBackground('#1a1f2e').setFontColor('#ffffff');
    h.setColumnWidths(1, 1, 190);
    h.setColumnWidth(CP_TIEMPOS_COL.ID_VENTA, 110);
    h.setColumnWidth(CP_TIEMPOS_COL.ID_NEGOCIO, 110);
    h.setColumnWidths(CP_TIEMPOS_COL.RESELLER, 2, 170);
    h.setColumnWidth(CP_TIEMPOS_COL.ENVIO, 60);
    h.setColumnWidth(CP_TIEMPOS_COL.COMANDA, 110);
    h.setColumnWidth(CP_TIEMPOS_COL.OPERADOR, 170);
    h.setColumnWidths(CP_TIEMPOS_COL.CARGAR, 1, 140);
    h.setColumnWidths(CP_TIEMPOS_COL.CARGADA, 1, 140);
    h.setColumnWidth(CP_TIEMPOS_COL.H_CARGA, 130);
    h.setColumnWidths(CP_TIEMPOS_COL.AUTORIZADO, 1, 140);
    h.setColumnWidth(CP_TIEMPOS_COL.H_AUTORIZAR, 140);
    h.setColumnWidths(CP_TIEMPOS_COL.DESPACHADO, 1, 140);
    h.setColumnWidth(CP_TIEMPOS_COL.H_DESPACHAR, 145);
    h.setColumnWidth(CP_TIEMPOS_COL.H_TOTAL, 190);
    try { h.setFrozenColumns(1); } catch (fe) {}
    try {
      var FILAS_FMT = 4999; // headroom generoso: no hay que retocar esto cuando crezcan las filas
      h.setConditionalFormatRules([
        _cpTiemposHeatRule(h.getRange(2, CP_TIEMPOS_COL.H_CARGA, FILAS_FMT, 1)),
        _cpTiemposHeatRule(h.getRange(2, CP_TIEMPOS_COL.H_AUTORIZAR, FILAS_FMT, 1)),
        _cpTiemposHeatRule(h.getRange(2, CP_TIEMPOS_COL.H_DESPACHAR, FILAS_FMT, 1)),
        _cpTiemposHeatRule(h.getRange(2, CP_TIEMPOS_COL.H_TOTAL, FILAS_FMT, 1))
      ]);
    } catch (cfe) { Logger.log('TIEMPOS heatmap: ' + cfe); }
  }
  return h;
}

// Heatmap adaptativo (verde=rápido → rojo=lento) para una columna de horas. Se calibra solo
// (min/mediana/max de lo que haya en la columna en cada momento, sin umbrales fijos en el
// código) — pensado para que el jefe (perfil Excel) vea de un vistazo dónde hay problemas
// con solo abrir la hoja, sin pasar por el programa.
function _cpTiemposHeatRule(range) {
  return SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#b7ecc4', SpreadsheetApp.InterpolationType.MIN, '')
    .setGradientMidpointWithValue('#ffe9a8', SpreadsheetApp.InterpolationType.PERCENTILE, '50')
    .setGradientMaxpointWithValue('#f3a6a6', SpreadsheetApp.InterpolationType.MAX, '')
    .setRanges([range])
    .build();
}


// Reescribe TODA la hoja TIEMPOS a partir de _cpTiemposCalc(). Se llama sola desde
// CP_crearEnvio() y CP_autoMailEnvios(); también expuesta para un botón manual "Actualizar".
function CP_actualizarHojaTiempos() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (le) {}
  try {
    var c = _cpTiemposCalc();
    var h = _cpTiemposHoja();
    var filas = [];

    c.rows.forEach(function(x) {
      var estado = x.despachadoTs ? '🟢 Despachado' : x.autorizadoTs ? '🔵 Esperando despacho' : '🟠 Esperando autorización';
      filas.push([
        estado, x.idVenta, x.idNegocio, x.reseller, x.cliente, x.envio, x.comanda, x.operador,
        x.cargarTs ? new Date(x.cargarTs) : '', x.cargadaTs ? new Date(x.cargadaTs) : '',
        x.horasCarga != null ? x.horasCarga : '',
        x.autorizadoTs ? new Date(x.autorizadoTs) : '', x.horasAutorizar != null ? x.horasAutorizar : '',
        x.despachadoTs ? new Date(x.despachadoTs) : '', x.horasDespachar != null ? x.horasDespachar : '',
        x.horasTotal != null ? x.horasTotal : ''
      ]);
    });
    c.pendientes.forEach(function(p) {
      filas.push([
        '🟡 Esperando carga de comanda', p.idVenta, p.idNegocio, p.reseller, '', '', '', '',
        p.cargarTs ? new Date(p.cargarTs) : '', '', p.horasAbierto != null ? p.horasAbierto : '',
        '', '', '', '', ''
      ]);
    });

    // más reciente primero (lo que más le interesa al operador es lo que está pasando ahora)
    filas.sort(function(a, b) {
      var ta = a[CP_TIEMPOS_COL.CARGAR - 1] instanceof Date ? a[CP_TIEMPOS_COL.CARGAR - 1].getTime() : 0;
      var tb = b[CP_TIEMPOS_COL.CARGAR - 1] instanceof Date ? b[CP_TIEMPOS_COL.CARGAR - 1].getTime() : 0;
      return tb - ta;
    });

    var nCols = CP_TIEMPOS_HEADERS.length;
    var last = h.getLastRow();
    if (last > 1) h.getRange(2, 1, last - 1, nCols).clearContent();
    if (filas.length) {
      h.getRange(2, 1, filas.length, nCols).setValues(filas);
      var fmtFecha = 'dd/mm/yyyy hh:mm', fmtHoras = '0.0" hs"';
      [CP_TIEMPOS_COL.CARGAR, CP_TIEMPOS_COL.CARGADA, CP_TIEMPOS_COL.AUTORIZADO, CP_TIEMPOS_COL.DESPACHADO].forEach(function(c1) {
        h.getRange(2, c1, filas.length, 1).setNumberFormat(fmtFecha);
      });
      [CP_TIEMPOS_COL.H_CARGA, CP_TIEMPOS_COL.H_AUTORIZAR, CP_TIEMPOS_COL.H_DESPACHAR, CP_TIEMPOS_COL.H_TOTAL].forEach(function(c1) {
        h.getRange(2, c1, filas.length, 1).setNumberFormat(fmtHoras);
      });
    }
    h.getRange(1, nCols + 2).setValue('Actualizado: ' + _fmtTs(new Date()));
    return { ok: true, filas: filas.length };
  } catch (e) {
    Logger.log('CP_actualizarHojaTiempos: ' + e);
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  } finally { try { lock.releaseLock(); } catch (fe) {} }
}


/* ── DIAGNÓSTICO (correr desde el editor si algo falla) ──── */
// Ejecutá esta función y mirá los logs: muestra la fila de encabezados
// detectada y a qué índice quedó mapeada cada columna.
function CP_debugColumnas() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) { Logger.log('No se detectó fila de encabezados con ID_Venta + ID_Entrega.'); return; }
  Logger.log('Fila de encabezados (0-based): ' + det.headerRow);
  Logger.log('Encabezados: ' + JSON.stringify(det.headers));
  Logger.log('Mapa columnas -> índice: ' + JSON.stringify(det.col));
  var flag = CP_FLAG.toUpperCase(), n = 0;
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    if (_s(rows[i][det.col.idEntrega]).toUpperCase() === flag) n++;
  }
  Logger.log('Filas con ID_Entrega = ' + CP_FLAG + ': ' + n);
}


// Diagnóstico crudo: qué pestañas hay, qué lee la hoja configurada y cómo
// normalizan las primeras filas. Correr desde el editor y mirar los logs.
function CP_debugCrudo() {
  var ss = _cpSS(CP_SS_ID);
  Logger.log('Pestañas en el archivo: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(' | '));
  Logger.log('CP_TAB configurada: "' + CP_TAB + '"');

  var h = ss.getSheetByName(CP_TAB);
  if (!h) { Logger.log('>> getSheetByName("' + CP_TAB + '") devolvió NULL. El nombre no coincide.'); return; }

  var rng = h.getDataRange();
  Logger.log('Dimensiones: ' + rng.getNumRows() + ' filas x ' + rng.getNumColumns() + ' columnas');
  var rows = rng.getValues();

  var maxScan = Math.min(4, rows.length);
  for (var r = 0; r < maxScan; r++) {
    Logger.log('--- Fila ' + r + ' (0-based) ---');
    Logger.log('  CRUDO: ' + JSON.stringify(rows[r]));
    Logger.log('  NORM : ' + JSON.stringify(rows[r].map(_norm)));
  }
}


// Diagnóstico de mails: cuántos resellers/RTV mapeó y ejemplos.
function CP_debugMails() {
  var rm = _cpResellerMap(), rt = _cpRtvMailMap();
  var conMail = Object.keys(rm).filter(function(k){ return rm[k].mail; });
  Logger.log('Resellers leídos: ' + Object.keys(rm).length + ' — con mail: ' + conMail.length);
  Object.keys(rm).slice(0, 6).forEach(function(k) { Logger.log('  ' + k + ' → mail:' + (rm[k].mail||'—') + ' | RTV:' + (rm[k].rtv||'—')); });
  Logger.log('RTV con mail (hoja RTV): ' + Object.keys(rt).length);
  Object.keys(rt).slice(0, 6).forEach(function(k) { Logger.log('  ' + k + ' → ' + rt[k]); });
}


// Diagnóstico de KITS: cuántos kits leyó y la composición de un par de ejemplos.
function CP_debugKits() {
  var map = _cpKitMap();
  var kits = Object.keys(map);
  Logger.log('KITS leídos: ' + kits.length);
  var ejemplos = ['T20COMBO', 'KITRTK3'].concat(kits.slice(0, 3));
  ejemplos.forEach(function(k) {
    var m = map[_kitKey(k)];
    if (!m) { Logger.log('  ' + k + ' → (no está)'); return; }
    Logger.log('  ' + k + ' → ' + m.orden.map(function(cu) {
      var comp = m.comps[cu];
      return comp.cant + 'x ' + comp.sku + ' (' + comp.desc + ')';
    }).join(' | '));
  });
}


/* ════════════════════════════════════════════════════════════
   DIAGNÓSTICO: ¿por qué una venta con envío no pasa a parcial/completo?
   Corré CP_diagEnvios() desde el editor y mirá el registro (Ver → Registro).
   Compara los ID_Venta de la hoja ENVIOS contra los ID_Venta que hoy
   tienen el flag CARGAR en Ventas → revela si el problema es:
     (a) el ID_Venta guardado no matchea (tipo/formato, ceros a la izquierda), o
     (b) la venta ya no está en CARGAR (desapareció de la lista).
════════════════════════════════════════════════════════════ */
function CP_diagEnvios() {
  var out = [];
  var norm = function(v) { return _s(v).toUpperCase(); };          // igual que las claves reales
  var normNum = function(v) { var s = _s(v).replace(/^0+/, ''); return s; }; // sin ceros a la izquierda

  // 1) ID_Venta con envíos (hoja ENVIOS)
  var envMap = _cpEnviosMap();
  var envKeys = Object.keys(envMap);
  out.push('ENVIOS: ' + envKeys.length + ' venta(s) con envío(s) registrados.');

  // 2) ID_Venta con flag CARGAR (hoja Ventas)
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) { out.push('❌ No detecté encabezados en Ventas.'); Logger.log(out.join('\n')); return out.join('\n'); }
  var col = det.col, flag = CP_FLAG.toUpperCase();
  var cargar = {};      // clave normal → {raw, tipo}
  var cargarNum = {};   // clave sin ceros → clave normal (para detectar near-match)
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    if (norm(rows[i][col.idEntrega]) !== flag) continue;
    var raw = rows[i][col.idVenta];
    var k = norm(raw); if (!k) continue;
    cargar[k] = { raw: raw, tipo: typeof raw };
    cargarNum[normNum(raw)] = k;
  }
  out.push('CARGAR: ' + Object.keys(cargar).length + ' venta(s) marcadas hoy.');
  out.push('');

  // 3) Por cada venta con envíos, ¿matchea?
  envKeys.forEach(function(k) {
    if (cargar[k]) {
      out.push('✅ "' + k + '" (' + envMap[k].length + ' env) → coincide con una venta en CARGAR.');
    } else if (cargarNum[normNum(k)]) {
      out.push('⚠️ "' + k + '" (' + envMap[k].length + ' env) → NO coincide EXACTO, pero sí sin ceros a la izquierda con "' + cargarNum[normNum(k)] + '" (Ventas tipo ' + cargar[cargarNum[normNum(k)]].tipo + '). ⇒ mismatch de formato/tipo.');
    } else {
      out.push('❌ "' + k + '" (' + envMap[k].length + ' env) → NO está en CARGAR (ni exacto ni por número). ⇒ la venta salió del flag CARGAR.');
    }
  });

  var txt = out.join('\n');
  Logger.log(txt);
  return txt;
}


/* ════════════════════════════════════════════════════════════
   DIAGNÓSTICO: ¿por qué NO sale el mail automático al reseller + RTV?
   Corré CP_diagAutoMail() desde el editor y mirá el registro (Ver → Registro).
   Chequea, en orden, TODO lo que tiene que estar bien para que el trigger
   horario CP_autoMailEnvios mande el mail:
     1) EMAIL_PRUEBA vacío (si tiene un mail, TODO se redirige a esa dirección
        y el reseller/RTV NO reciben nada → causa #1 después de usar el modo prueba).
     2) AUTO_MAIL_DESPACHO = SI (si no, el trigger corre pero no manda nada).
     3) El trigger CP_autoMailEnvios instalado (CP_setupAutoMail).
     4) Por cada envío sin mail: ¿tiene guía? ¿tiene destinatarios?
   Es SOLO LECTURA: no manda ningún correo.
════════════════════════════════════════════════════════════ */
function CP_diagAutoMail() {
  var out = [];
  var cfg = _cpConfig();

  // 1) MODO PRUEBA (la causa más común después de testear)
  var prueba = _s(cfg['EMAIL_PRUEBA']);
  if (prueba) {
    out.push('🔴 EMAIL_PRUEBA = "' + prueba + '"  → MODO PRUEBA ACTIVO.');
    out.push('   TODOS los correos (reseller, RTV, Sole) se redirigen a esa dirección.');
    out.push('   El reseller/RTV NO reciben nada. Para volver a producción: VACIÁ EMAIL_PRUEBA en _CONFIG.');
  } else {
    out.push('✅ EMAIL_PRUEBA vacío (producción: los mails van a los destinatarios reales).');
  }

  // 2) Interruptor del automático
  var auto = _s(cfg['AUTO_MAIL_DESPACHO']).toUpperCase();
  var autoOn = (auto.indexOf('SI') === 0 || auto.indexOf('SÍ') === 0);
  out.push((autoOn ? '✅' : '🔴') + ' AUTO_MAIL_DESPACHO = "' + (_s(cfg['AUTO_MAIL_DESPACHO']) || '(vacío)') + '"' +
           (autoOn ? '' : '  → el envío automático está DESACTIVADO. Poné AUTO_MAIL_DESPACHO=SI en _CONFIG.'));

  // 3) Trigger instalado
  var trg = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'CP_autoMailEnvios'; });
  out.push((trg.length ? '✅' : '🔴') + ' Trigger CP_autoMailEnvios: ' +
           (trg.length ? trg.length + ' instalado(s) (corre cada 30 min).' : 'NO instalado → corré CP_setupAutoMail() una vez.'));

  // 4) Estado de los envíos pendientes de mail
  out.push('');
  var mapAll = _cpEnviosMap();
  var master = _cpMasterMap();
  // Pre-lecturas ÚNICAS: antes se releía Ventas/Resellers/RTV por CADA envío con guía →
  // con un backlog grande de envíos sin mail, el diagnóstico se colgaba y no mostraba nada.
  var ventaMap    = _cpVentaResellerRtvMap();   // { IDVENTA: {reseller, rtv} } (1 lectura de Ventas)
  var resellerMap = _cpResellerMap();           // 1 lectura de Resellers
  var rtvMap      = _cpRtvMailMap();            // 1 lectura de RTV
  var total = 0, yaMail = 0, sinComanda = 0, sinDespachar = 0, sinDest = 0, listos = 0;
  var detListos = [], detSinDespachar = [], detSinDest = [];

  Object.keys(mapAll).forEach(function(k) {
    mapAll[k].forEach(function(e) {
      total++;
      if (e.mailReseller) { yaMail++; return; }
      var parts = _s(e.comanda).split('/').map(function(s){ return s.trim(); }).filter(Boolean);
      if (!parts.length) { sinComanda++; return; }
      var listo = _cpEnvioListoDespacho(parts, master).listo;
      if (!listo) { sinDespachar++; if (detSinDespachar.length < 8) detSinDespachar.push(k + ' env' + e.envio + ' (' + e.comanda + ')'); return; }
      var vi = ventaMap[k.toUpperCase()] || {};
      var det = { reseller: vi.reseller || '', rtv: vi.rtv || '' };
      var dest = _cpDestinatariosEnvio(det, cfg, resellerMap, rtvMap);
      if (!dest.to.length) { sinDest++; if (detSinDest.length < 8) detSinDest.push(k + ' env' + e.envio + ' — ' + (det.reseller || '?')); return; }
      listos++; if (detListos.length < 8) detListos.push(k + ' env' + e.envio + ' → ' + dest.to.join(', '));
    });
  });

  out.push('ENVÍOS: ' + total + ' total · ' + yaMail + ' con mail ya enviado · ' + (total - yaMail) + ' sin mail.');
  out.push('  De los que están sin mail:');
  out.push('    ⏳ ' + sinDespachar + ' esperan que Comandas Master (col F) diga DESPACHADO (tener guía no alcanza, solo significa autorizado).');
  out.push('    🔴 ' + sinDest + ' listos pero SIN destinatarios (falta mail del reseller/RTV o MAIL_DESTINATARIOS).');
  out.push('    ⚠️ ' + sinComanda + ' sin comanda cargada.');
  out.push('    ✅ ' + listos + ' LISTOS para enviar (DESPACHADO y con destinatarios).');
  if (detListos.length)  out.push('       listos: ' + detListos.join(' | '));
  if (detSinDespachar.length) out.push('       sin despachar: ' + detSinDespachar.join(' | '));
  if (detSinDest.length) out.push('       sin dest: ' + detSinDest.join(' | '));

  // Veredicto
  out.push('');
  if (prueba)           out.push('👉 CAUSA MÁS PROBABLE: MODO PRUEBA activo (EMAIL_PRUEBA). Vacialo en _CONFIG y probá de nuevo.');
  else if (!autoOn)     out.push('👉 CAUSA: AUTO_MAIL_DESPACHO no está en SI.');
  else if (!trg.length) out.push('👉 CAUSA: falta instalar el trigger (corré CP_setupAutoMail una vez).');
  else if (listos > 0)  out.push('👉 Config OK y hay ' + listos + ' envío(s) listos. Corré CP_autoMailEnvios() a mano para forzar el envío ahora, o esperá al próximo ciclo (≤30 min). Si aun así no llegan, revisá spam/cuota de Gmail.');
  else if (sinDest > 0) out.push('👉 CAUSA: hay envíos con guía pero SIN destinatarios (cargá el mail del reseller en Resellers o del RTV en la hoja RTV).');
  else                  out.push('👉 No hay envíos listos: los pendientes esperan la guía. Cuando aparezca el número de seguimiento el mail sale solo.');

  var txt = out.join('\n');
  Logger.log(txt);
  return txt;
}
