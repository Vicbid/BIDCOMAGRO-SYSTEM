// @version 1.1
// ============================================================
//  COMANDAS — Reporte de tiempos + herramientas CP_debug*/CP_diag*
//  (diagnóstico genérico, no atado a un incidente puntual).
//  Extraído de CP_Código.js 1.43 el 2026-07-30 — reorganización sin
//  cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


/* ════════════════════════════════════════════════════════════
   REPORTE DE TIEMPOS — CARGAR → comanda cargada
   delta = fecha del envío (cuando el operador registró la comanda)
           menos el momento MÁS ANTIGUO en que se marcó CARGAR en la venta.
   origen 'edit' = exacto | 'deteccion' = aproximado.
════════════════════════════════════════════════════════════ */

// Mapa liviano leyendo Ventas una sola vez: { IDVENTA: {idVenta, reseller, cliente} }
function _cpInfoVentasMap() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) return {};
  var col = det.col, m = {};
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var idv = _s(col.idVenta > -1 ? rows[i][col.idVenta] : ''); if (!idv) continue;
    var k = idv.toUpperCase();
    if (!m[k]) m[k] = {
      idVenta:  idv,
      reseller: _s(col.reseller > -1 ? rows[i][col.reseller] : ''),
      cliente:  _s(col.razonSocial > -1 ? rows[i][col.razonSocial] : '')
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


// Devuelve { ok, rows:[por envío], pendientes:[marcadas CARGAR sin envío], resumen:{primer, todos} }.
function CP_reporteTiempos() {
  try {
    var logMap = _cpLogMap();
    // CARGAR más antiguo por venta (+ si es exacto)
    var cargarPorVenta = {};
    Object.keys(logMap).forEach(function(key) {
      var idv = key.split('||')[0], reg = logMap[key], cur = cargarPorVenta[idv];
      if (!cur || reg.ts < cur.ts) cargarPorVenta[idv] = { ts: reg.ts, origen: reg.origen };
    });
    var info = _cpInfoVentasMap();
    var enviosMap = _cpEnviosMap();
    var rows = [], dPrim = [], dTodos = [], dDesp = [];
    Object.keys(enviosMap).forEach(function(idvU) {
      var arr = enviosMap[idvU], carg = cargarPorVenta[idvU];
      var cargTs = carg ? carg.ts.getTime() : null, exacto = carg ? (carg.origen === 'edit') : false;
      var iv = info[idvU] || {};
      arr.forEach(function(e) {
        if (!e.fechaTs) return;
        var horas = cargTs != null ? Math.round(((e.fechaTs - cargTs) / 3600000) * 10) / 10 : null;
        if (horas != null && horas >= 0) { dTodos.push(horas); if (e.envio === 1) dPrim.push(horas); }
        // tramo despacho: comanda cargada (fecha del envío) → mail al reseller (proxy del despacho)
        var hDesp = e.mailResellerTs ? Math.round(((e.mailResellerTs - e.fechaTs) / 3600000) * 10) / 10 : null;
        if (hDesp != null && hDesp >= 0) dDesp.push(hDesp);
        rows.push({
          idVenta: iv.idVenta || idvU, reseller: iv.reseller || '', cliente: iv.cliente || '',
          envio: e.envio, comanda: e.comanda,
          cargarStr: cargTs != null ? _fmtTs(new Date(cargTs)) : '',
          origen: carg ? (exacto ? 'exacto' : 'aprox') : 'sin registro',
          cargadaStr: e.fechaStr || '', operador: e.operador || '',
          despachoStr: e.mailResellerTs ? _fmtTs(new Date(e.mailResellerTs)) : '',
          horas: horas, horasDespacho: hDesp, primero: (e.envio === 1)
        });
      });
    });
    rows.sort(function(a, b) { return a.idVenta === b.idVenta ? a.envio - b.envio : (a.idVenta < b.idVenta ? -1 : 1); });
    // pendientes: marcadas CARGAR pero sin ningún envío todavía
    var now = Date.now(), pend = [];
    Object.keys(cargarPorVenta).forEach(function(idvU) {
      if (enviosMap[idvU]) return;
      var carg = cargarPorVenta[idvU], iv = info[idvU] || {};
      pend.push({ idVenta: iv.idVenta || idvU, reseller: iv.reseller || '',
        horasAbierto: Math.round(((now - carg.ts.getTime()) / 3600000) * 10) / 10,
        origen: (carg.origen === 'edit' ? 'exacto' : 'aprox') });
    });
    pend.sort(function(a, b) { return b.horasAbierto - a.horasAbierto; });
    return { ok: true, rows: rows, pendientes: pend, resumen: { primer: _cpStats(dPrim), todos: _cpStats(dTodos), despacho: _cpStats(dDesp) } };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
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
