// @version 1.2
// ============================================================
//  COMANDAS — Envíos: CRUD/pendientes de entrega, trigger onEdit,
//  snapshot, log de estado.
//  Extraído de CP_Código.js 1.43 el 2026-07-30 — reorganización sin
//  cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// Devuelve/crea la hoja de log (en el sheet SEPARADO, nunca en Ventas).
function _cpLogHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_LOG_TAB);
  if (!h) {
    h = ss.insertSheet(CP_LOG_TAB);
    h.getRange(1, 1, 1, 5).setValues([['ID_Venta', 'SKU', 'Marcado CARGAR', 'Origen', 'Usuario']]);
    h.setFrozenRows(1);
  }
  return h;
}


// Lee el log a un mapa { clave: {ts:Date, origen:string} } (se queda con el más antiguo).
function _cpLogMap() {
  try {
    var ss = _cpSS(CP_LOG_SS_ID);
    var h = ss.getSheetByName(CP_LOG_TAB);
    if (!h) return {};
    var d = h.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var idv = d[i][0], sku = d[i][1], fecha = d[i][2], origen = d[i][3];
      if (idv === '' && sku === '') continue;
      var ts = (fecha instanceof Date) ? fecha : (fecha ? new Date(fecha) : null);
      if (!ts || isNaN(ts.getTime())) continue;
      var key = _cpKey(idv, sku);
      if (!m[key] || ts < m[key].ts) m[key] = { ts: ts, origen: String(origen || 'edit') };
    }
    return m;
  } catch (e) {
    Logger.log('_cpLogMap error: ' + e);
    return {};
  }
}


// Agrega filas [idv, sku, Date, origen, email] al log, evitando duplicados por clave.
function _cpLogStamp(filas) {
  if (!filas || !filas.length) return;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { /* seguimos igual */ }
  try {
    var h = _cpLogHoja();
    var d = h.getDataRange().getValues();
    var seen = {};
    for (var i = 1; i < d.length; i++) seen[_cpKey(d[i][0], d[i][1])] = true;

    var nuevas = [];
    filas.forEach(function(f) {
      var key = _cpKey(f[0], f[1]);
      if (!seen[key]) { seen[key] = true; nuevas.push(f); }
    });
    if (nuevas.length) {
      h.getRange(h.getLastRow() + 1, 1, nuevas.length, 5).setValues(nuevas);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


// TRIGGER onEdit (instalable). Se dispara con cada edición manual/pegado.
function CP_onEditVentas(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CP_TAB) return;

    var lastCol = sh.getLastColumn();
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(_norm);
    var cEnt = headers.indexOf('identrega');
    var cVen = headers.indexOf('idventa');
    var cSku = headers.indexOf('sku');
    if (cEnt === -1 || cVen === -1) return;

    var r0 = e.range.getRow(), c0 = e.range.getColumn();
    var nR = e.range.getNumRows(), nC = e.range.getNumColumns();

    // ¿el rango editado incluye la columna ID_Entrega?
    var colEnt1 = cEnt + 1; // 1-based
    if (colEnt1 < c0 || colEnt1 > c0 + nC - 1) return;

    var block = sh.getRange(r0, 1, nR, lastCol).getValues();
    var flag = CP_FLAG.toUpperCase();
    var now = new Date();
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}

    var filas = [];
    for (var i = 0; i < nR; i++) {
      if (r0 + i === 1) continue; // saltar encabezado
      var row = block[i];
      if (String(row[cEnt] || '').trim().toUpperCase() !== flag) continue;
      var idv = String(row[cVen] || '').trim();
      var sku = cSku > -1 ? String(row[cSku] || '').trim() : '';
      filas.push([idv, sku, now, 'edit', email]);
    }
    _cpLogStamp(filas);
  } catch (err) {
    // nunca interrumpir la edición del usuario
    Logger.log('CP_onEditVentas error: ' + err);
  }
}


// SETUP: correr UNA vez desde el editor. Instala el trigger onEdit y crea la hoja de log.
function CP_setupTrigger() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'CP_onEditVentas';
  });
  if (ya) { Logger.log('El trigger onEdit ya estaba instalado. Nada que hacer.'); }
  else {
    ScriptApp.newTrigger('CP_onEditVentas').forSpreadsheet(CP_SS_ID).onEdit().create();
    Logger.log('✅ Trigger onEdit instalado sobre el spreadsheet.');
  }
  _cpLogHoja();
  _cpEnviosHoja();
  _cpConfigHoja();
  _cpRtvHoja();
  _cpPendHoja();
  _cpAuditHoja();
  try { CP_poblarRtvDesdeResellers(); } catch (e) { Logger.log('poblar RTV: ' + e); }
  Logger.log('✅ Hojas listas (CARGAR_LOG, ENVIOS, _CONFIG, RTV, PENDIENTES_ENTREGA, AUDITORIA) en el sheet de log.');
}


// SELF-CHECK de read-only: verifica que el sheet de ESCRITURA (log) sea distinto del de
// Ventas, y lista las hojas donde el programa escribe. Garantía extra de que Ventas no se toca.
// Correr desde el editor y mirar los logs / el objeto devuelto.
function CP_selfCheckReadOnly() {
  var ok = (CP_LOG_SS_ID !== CP_SS_ID);
  var r = {
    ok: ok,
    ventas_solo_lectura: CP_SS_ID,
    escribe_en: CP_LOG_SS_ID,
    hojas_de_escritura: [CP_LOG_TAB, CP_ENVIOS_TAB, CP_CONFIG_TAB, CP_RTV_TAB, CP_PEND_TAB, CP_AUDIT_TAB],
    mensaje: ok ? 'OK: el archivo de Ventas y el de escritura son distintos; nada se escribe sobre Ventas.'
                : '⚠️ PELIGRO: CP_LOG_SS_ID == CP_SS_ID — la app escribiría sobre la hoja de Ventas.'
  };
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}


// OPCIONAL: estampa una marca de "detección" (aprox = ahora) para todas las filas
// que HOY están en CARGAR y todavía no tienen registro. Útil como línea de base
// para los pedidos que ya estaban marcados antes de instalar el trigger.
function CP_snapshotCargar() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) { Logger.log('No se detectaron encabezados.'); return; }
  var col = det.col, flag = CP_FLAG.toUpperCase();
  var existentes = _cpLogMap();
  var now = new Date();
  var filas = [];
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[col.idEntrega] || '').trim().toUpperCase() !== flag) continue;
    var idv = String(r[col.idVenta] || '').trim();
    var sku = col.sku > -1 ? String(r[col.sku] || '').trim() : '';
    if (existentes[_cpKey(idv, sku)]) continue; // ya registrado
    filas.push([idv, sku, now, 'deteccion', '']);
  }
  _cpLogStamp(filas);
  Logger.log('Snapshot: ' + filas.length + ' fila(s) nuevas marcadas como detección (aprox.).');
}


/* ════════════════════════════════════════════════════════════
   ENVÍOS — una venta puede tener VARIOS envíos. Cada envío tiene su
   propia comanda, su aprobación (mail a Sole) y su guía (mail al
   reseller). Se guardan en CP_ENVIOS_TAB (sheet de log, nunca en Ventas).
   Columnas: ID_Venta | Envío | Comanda | Fecha | Operador |
     Productos(JSON {SKU:cant}) | Guía | Transportista | Estado |
     Mail Aprobador | Mail Reseller | Nota Aprobador | Nota Reseller
════════════════════════════════════════════════════════════ */
function _cpEnviosHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_ENVIOS_TAB);
  if (!h) {
    h = ss.insertSheet(CP_ENVIOS_TAB);
    h.getRange(1, 1, 1, 15).setValues([[
      'ID_Venta', 'Envío', 'Comanda', 'Fecha', 'Operador', 'Productos',
      'Guía', 'Transportista', 'Estado', 'Mail Aprobador', 'Mail Reseller', 'Nota Aprobador', 'Nota Reseller',
      'Thread ID Reseller', 'Mail Autorizado Reseller'
    ]]);
    h.setFrozenRows(1);
    h.getRange('A1:O1').setFontWeight('bold');
  } else {
    // Migración de hojas ya existentes: agrega las columnas que se fueron sumando sin romper
    // lo que ya había (ver CP_Mail.js: col N = hilo de Gmail, col O = mail #1 "autorizado").
    if (!_s(h.getRange(1, 14).getValue())) {
      if (h.getMaxColumns() < 14) h.insertColumnsAfter(h.getMaxColumns(), 14 - h.getMaxColumns());
      h.getRange(1, 14).setValue('Thread ID Reseller').setFontWeight('bold');
    }
    if (!_s(h.getRange(1, 15).getValue())) {
      if (h.getMaxColumns() < 15) h.insertColumnsAfter(h.getMaxColumns(), 15 - h.getMaxColumns());
      h.getRange(1, 15).setValue('Mail Autorizado Reseller').setFontWeight('bold');
    }
  }
  return h;
}

// Normaliza objeto {SKU:cant} → JSON (sólo > 0) o '' si no hay.
function _cpProductosJson(obj) {
  var limpio = {};
  if (obj && typeof obj === 'object') {
    Object.keys(obj).forEach(function(k) { var v = _num(obj[k]); if (v > 0) limpio[String(k).toUpperCase()] = v; });
  }
  return Object.keys(limpio).length ? JSON.stringify(limpio) : '';
}


// Mapa { IDVENTA: [ {envio, comanda, fecha, fechaStr, fechaTs, operador, productos, guia,
//                    transportista, estado, mailAprob, mailReseller, notaAprob, notaReseller,
//                    threadIdReseller, mailAutorizado, rowIdx} ] }
function _cpEnviosMap() {
  try {
    var ss = _cpSS(CP_LOG_SS_ID);
    var h = ss.getSheetByName(CP_ENVIOS_TAB);
    if (!h) return {};
    var d = h.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var idv = _s(d[i][0]); if (!idv) continue;
      var key = idv.toUpperCase();
      var fecha = (d[i][3] instanceof Date) ? d[i][3] : (d[i][3] ? new Date(d[i][3]) : null);
      var e = {
        envio:         _num(d[i][1]) || 0,
        comanda:       _s(d[i][2]),
        fecha:         fecha,
        fechaStr:      fecha ? _fmtTs(fecha) : '',
        fechaTs:       (fecha && !isNaN(fecha.getTime())) ? fecha.getTime() : null,
        operador:      _s(d[i][4]),
        productos:     _cpParseJson(d[i][5]),
        guia:          _s(d[i][6]),
        transportista: _s(d[i][7]),
        estado:        _s(d[i][8]),
        mailAprob:     _s(d[i][9]),
        mailReseller:  _s(d[i][10]),
        mailResellerTs:(function() { var dt = _cpParseFechaAr(d[i][10]); return dt ? dt.getTime() : null; })(),
        notaAprob:     _s(d[i][11]),
        notaReseller:  _s(d[i][12]),
        threadIdReseller: _s(d[i][13]),
        mailAutorizado:   _s(d[i][14]),
        rowIdx:        i + 1
      };
      if (!m[key]) m[key] = [];
      m[key].push(e);
    }
    Object.keys(m).forEach(function(k) { m[k].sort(function(a, b) { return a.envio - b.envio; }); });
    return m;
  } catch (e) { Logger.log('_cpEnviosMap error: ' + e); return {}; }
}


// Total enviado por SKU sumando todos los envíos de una venta.
function _cpEnviadoTotal(enviosArr) {
  var t = {};
  (enviosArr || []).forEach(function(e) {
    Object.keys(e.productos || {}).forEach(function(sk) { t[sk] = (t[sk] || 0) + _num(e.productos[sk]); });
  });
  return t;
}


// Lo que falta enviar de una venta: [{sku, descripcion, cantidad, cantNum}] (pedido - enviado > 0).
function _cpPendingVenta(det, enviosArr) {
  var enviado = _cpEnviadoTotal(enviosArr);
  var pend = [];
  (det.cargar || []).forEach(function(it) {
    var falta = Math.round(((it.cantNum || 0) - _num(enviado[String(it.sku).toUpperCase()])) * 100) / 100;
    if (falta > 0) pend.push({ sku: it.sku, descripcion: it.desc || it.descripcion, cantidad: _fmtCant(falta), cantNum: falta });
  });
  return pend;
}


// Detalle de los productos de un envío: [{sku, desc, cant}] (para los mails).
function _cpDetalleEnvio(det, prodMap) {
  var descBy = {};
  (det.cargar || []).forEach(function(it) { descBy[String(it.sku).toUpperCase()] = it.desc || it.descripcion; });
  return Object.keys(prodMap || {}).map(function(sk) { return { sku: sk, desc: descBy[sk] || '', cant: _fmtCant(_num(prodMap[sk])) }; });
}


// ── PENDIENTES DE ENTREGA (lo que falta enviar: 1 producto por línea) ──
function _cpPendHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_PEND_TAB);
  if (!h) {
    h = ss.insertSheet(CP_PEND_TAB);
    h.getRange(1, 1, 1, 8).setValues([['ID_Venta', 'SKU', 'Descripción', 'Cantidad', 'Reseller', 'Razón Social', 'Fecha', 'Estado']]);
    h.setFrozenRows(1);
    h.getRange('A1:H1').setFontWeight('bold');
    [110, 130, 300, 80, 170, 210, 130, 120].forEach(function(w, i) { h.setColumnWidth(i + 1, w); });
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['PENDIENTE', 'ENTREGADO'], true).setAllowInvalid(true).build();
    h.getRange('H2:H2000').setDataValidation(rule);
  }
  return h;
}


// Reescribe las líneas pendiente-de-entrega de una venta (1 fila por SKU que falta enviar),
// PRESERVANDO el Estado y la Fecha que ya tuviera cada SKU (para no pisar ediciones manuales).
function _cpSyncPendientesEntrega(idVenta, det, pendArr) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var h = _cpPendHoja();
    var d = h.getDataRange().getValues();
    var key = _s(idVenta).toUpperCase();
    var prev = {};   // SKU -> {estado, fecha}
    for (var i = d.length - 1; i >= 1; i--) {
      if (_s(d[i][0]).toUpperCase() === key) {
        prev[_s(d[i][1]).toUpperCase()] = { estado: _s(d[i][7]), fecha: d[i][6] };
        h.deleteRow(i + 1);
      }
    }
    if (!pendArr || !pendArr.length) return;
    var now = new Date();
    var filas = pendArr.map(function(p) {
      var pv = prev[String(p.sku).toUpperCase()] || {};
      return [_cpSafeCell(idVenta), _cpSafeCell(p.sku), _cpSafeCell(p.descripcion || ''), p.cantNum,
              _cpSafeCell(det.reseller || ''), _cpSafeCell(det.razonSocial || ''),
              (pv.fecha instanceof Date ? pv.fecha : now), (_s(pv.estado) || 'PENDIENTE')];
    });
    h.getRange(h.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}


// Lee PENDIENTES_ENTREGA → { 'IDVENTA||SKU': {estado, entregado, fechaStr} }.
// Permite reflejar en la app lo que se marca ENTREGADO manualmente en la hoja (sync inverso).
function _cpEntregaMap() {
  try {
    var ss = _cpSS(CP_LOG_SS_ID);
    var h = ss.getSheetByName(CP_PEND_TAB);
    if (!h) return {};
    var d = h.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var idv = _s(d[i][0]), sku = _s(d[i][1]); if (!idv) continue;
      var estado = _s(d[i][7]).toUpperCase();
      var fecha = (d[i][6] instanceof Date) ? d[i][6] : null;
      m[_cpKey(idv, sku)] = { estado: estado, entregado: (estado === 'ENTREGADO'), fechaStr: fecha ? _fmtTs(fecha) : '' };
    }
    return m;
  } catch (e) { Logger.log('_cpEntregaMap error: ' + e); return {}; }
}


// El operador crea un ENVÍO: N° de comanda + cuántas unidades de cada SKU manda en este envío.
// productos = { SKU: cantidad }. Dispara el mail de aprobación a Sole (con PDF adjunto).
function CP_crearEnvio(idVenta, comanda, productos, notaAprob, notaReseller, force) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    idVenta = _s(idVenta); comanda = _s(comanda);
    notaAprob = _s(notaAprob); notaReseller = _s(notaReseller);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    if (!comanda) return { ok: false, mensaje: 'Ingresá el número de comanda.' };
    if (!/^\d{5,}(\/\d{5,})*$/.test(comanda)) return { ok: false, mensaje: 'N° de comanda inválido (solo números, o dos separados por /).' };

    var det = _cpDetalleVenta(idVenta);
    var enviosArr = _cpEnviosMap()[idVenta.toUpperCase()] || [];

    // comanda ya usada en otro envío de esta u otra venta → aviso salvo force
    if (!force) {
      var usada = false, dondeV = '';
      var mapAll = _cpEnviosMap();
      Object.keys(mapAll).forEach(function(k) {
        mapAll[k].forEach(function(e) {
          e.comanda.split('/').forEach(function(cp) {
            comanda.split('/').forEach(function(cn) { if (cp.trim() && cp.trim() === cn.trim()) { usada = true; dondeV = k; } });
          });
        });
      });
      if (usada) return { ok: false, yaUsada: true, donde: dondeV, mensaje: 'La comanda ' + comanda + ' ya está usada en un envío (' + dondeV + ').' };
    }

    // lo pendiente antes de este envío
    var pendBy = {};
    _cpPendingVenta(det, enviosArr).forEach(function(p) { pendBy[String(p.sku).toUpperCase()] = p.cantNum; });
    // productos a enviar ahora, topeados a lo pendiente
    var envProd = {};
    Object.keys(productos || {}).forEach(function(k) {
      var ku = String(k).toUpperCase();
      var q = Math.min(_num(productos[k]), pendBy[ku] || 0);
      if (q > 0) envProd[ku] = Math.round(q * 100) / 100;
    });
    if (!Object.keys(envProd).length) return { ok: false, mensaje: 'No hay productos para enviar (revisá las cantidades).' };

    var nextEnvio = 1;
    enviosArr.forEach(function(e) { if (e.envio >= nextEnvio) nextEnvio = e.envio + 1; });
    var email = ''; try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}
    var now = new Date();

    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch (e) {}
    try {
      _cpEnviosHoja().appendRow([_cpSafeCell(idVenta), nextEnvio, comanda, now, email, _cpProductosJson(envProd),
        '', '', '', '', '', _cpSafeCell(notaAprob), _cpSafeCell(notaReseller)]);
    } finally { try { lock.releaseLock(); } catch (e) {} }

    // recomputar pendiente y espejar (1 producto por línea)
    var pend2 = _cpPendingVenta(det, enviosArr.concat([{ productos: envProd }]));
    try { _cpSyncPendientesEntrega(idVenta, det, pend2); } catch (pe) { Logger.log('syncPend: ' + pe); }

    // mail a Sole (aprobación de ESTE envío) con PDF adjunto
    var res = { ok: true, envio: nextEnvio, completo: pend2.length === 0 };
    var soleRes = _cpEnviarMailSoleCore(idVenta, nextEnvio);
    if (soleRes && soleRes.ok) { res.mailSole = true; if (soleRes.sinPdf) res.mailSoleSinPdf = true; }
    else { res.mailSole = false; if (soleRes && soleRes.mensaje) res.mailSoleError = soleRes.mensaje; }
    _cpAuditar('Crear envío', idVenta, nextEnvio, 'comanda ' + comanda + ' · ' + Object.keys(envProd).length + ' ítem(s)' + (res.mailSole ? ' · mail Sole' : ''));
    return res;
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


// Envía (o reenvía) a Sole el mail de aprobación de UN envío (con PDF adjunto). Marca col "Mail Aprobador".
function _cpEnviarMailSoleCore(idVenta, envio) {
  idVenta = _s(idVenta); envio = _num(envio);
  var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
  var e = null; arr.forEach(function(x) { if (x.envio === envio) e = x; });
  if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };
  var cfg = _cpConfig();
  var sole = cfg['MAIL_APROBACION'];
  if (!sole) return { ok: false, sinDestino: true, mensaje: 'Falta MAIL_APROBACION en _CONFIG.' };
  var det = _cpDetalleVenta(idVenta);
  var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
  var pdfs = [];
  parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
  var detEnv = _cpDetalleEnvio(det, e.productos);
  var pend = _cpPendingVenta(det, arr);   // faltantes de la venta tras este envío
  var asunto = (cfg['ASUNTO_APROBACION'] || 'APROBAR MC · {COMANDA} · {IDVENTA}')
    .replace('{COMANDA}', parts.join('/')).replace('{IDVENTA}', idVenta);
  var opts = {
    htmlBody: _cpMailAprobacionHtml(idVenta, parts, det, pdfs, e.notaAprob, detEnv, e.envio, pend),
    name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO'
  };
  if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];
  var adj = _cpPdfBlobs(parts);
  if (adj.length) opts.attachments = adj;
  try {
    _cpGmailSend(sole, asunto, 'Autorizar comanda ' + e.comanda + ' en Masterchief.', opts);
  } catch (se) {
    return { ok: false, mailError: true, mensaje: 'No se pudo enviar el mail a Sole: ' + String(se && se.message ? se.message : se) };
  }
  _cpMarcarMailAprob(e.rowIdx);
  return { ok: true, sinPdf: !pdfs.length, destinatario: sole };
}


function _cpMarcarMailAprob(rowIdx) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try { _cpEnviosHoja().getRange(rowIdx, 10).setValue('SÍ · ' + _fmtTs(new Date())); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}


// Reenvía a Sole el mail de aprobación de un envío (por si no lo vio).
function CP_reenviarAprobacion(idVenta, envio) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    var r = _cpEnviarMailSoleCore(_s(idVenta), _num(envio));
    if (r && r.ok) { _cpAuditar('Reenviar aprobación', idVenta, envio, 'a ' + (r.destinatario || '')); return { ok: true, destinatario: r.destinatario, sinPdf: !!r.sinPdf }; }
    return { ok: false, mensaje: (r && r.mensaje) || 'No se pudo reenviar.' };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


// Chequea si están en Drive los PDF de una comanda (una o varias separadas por /).
// Devuelve { ok, hay:[...], falta:[...], todos:bool }.
function CP_checkPdf(comanda) {
  try {
    var parts = _s(comanda).split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    var hay = [], falta = [];
    parts.forEach(function(p) { if (_cpBuscarPdf(p)) hay.push(p); else falta.push(p); });
    return { ok: true, hay: hay, falta: falta, todos: (parts.length > 0 && falta.length === 0) };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}


// Borra un envío (solo si NO se mandó el mail al reseller). Recalcula pendientes.
function CP_borrarEnvio(idVenta, envio) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch (e) {}
    try {
      var h = _cpEnviosHoja();
      var d = h.getDataRange().getValues();
      var key = idVenta.toUpperCase();
      for (var i = d.length - 1; i >= 1; i--) {
        if (_s(d[i][0]).toUpperCase() === key && (_num(d[i][1]) === envio)) {
          if (_s(d[i][10])) return { ok: false, mensaje: 'Ya se envió el mail al reseller de este envío; no se puede borrar.' };
          h.deleteRow(i + 1);
        }
      }
    } finally { try { lock.releaseLock(); } catch (e) {} }
    var det = _cpDetalleVenta(idVenta);
    var enviosArr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    try { _cpSyncPendientesEntrega(idVenta, det, _cpPendingVenta(det, enviosArr)); } catch (pe) { Logger.log('syncPend: ' + pe); }
    _cpAuditar('Borrar envío', idVenta, envio, '');
    return { ok: true };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


// Edita la nota para el reseller de un envío puntual.
function CP_setNotaResellerEnvio(idVenta, envio, texto) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    idVenta = _s(idVenta); envio = _num(envio); texto = _s(texto);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch (e) {}
    try {
      var h = _cpEnviosHoja();
      var d = h.getDataRange().getValues();
      var key = idVenta.toUpperCase();
      for (var i = 1; i < d.length; i++) {
        if (_s(d[i][0]).toUpperCase() === key && (_num(d[i][1]) === envio)) { h.getRange(i + 1, 13).setValue(_cpSafeCell(texto)); break; }
      }
    } finally { try { lock.releaseLock(); } catch (e) {} }
    return { ok: true };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}
