// @version 1.3
// ============================================================
//  COMANDAS — Datos de apoyo: PDFs adjuntos, configuración,
//  detalle de venta, mapeo de resellers/RTV.
//  Extraído de CP_Código.js 1.43 el 2026-07-30 — reorganización sin
//  cambios funcionales (más archivos, mismo comportamiento).
// ============================================================



/* ════════════════════════════════════════════════════════════
   DESPACHO — lee "Comandas Master" (col A idComprobante → estado F,
   guía K, transportista B). Habilita el envío de mail manual.
════════════════════════════════════════════════════════════ */

// Mapa { COMANDA(idComprobante): {estado, guia, transportista, fechaEntrega, idVenta} }
function _cpMasterMap() {
  try {
    var ss = _cpSS(CP_SS_ID);
    var h = ss.getSheetByName(CP_MASTER_TAB);
    if (!h) { Logger.log('Master tab no encontrada: ' + CP_MASTER_TAB); return {}; }
    var d = h.getDataRange().getValues();
    if (d.length < 2) return {};
    var H = d[0].map(_norm);
    var cId    = H.indexOf('idcomprobante');            if (cId    === -1) cId    = 0;  // A
    var cTrans = H.indexOf('nombredespachomedio');      if (cTrans === -1) cTrans = 1;  // B
    var cEst   = H.indexOf('nombredespachoestado');     if (cEst   === -1) cEst   = 5;  // F
    var cFecha = H.indexOf('fechaentregadespacho');     if (cFecha === -1) cFecha = 6;  // G
    var cGuia  = H.indexOf('codigoseguimientodespacho');if (cGuia  === -1) cGuia  = 10; // K
    var cIdV   = H.indexOf('idventa');                  if (cIdV   === -1) cIdV   = 11; // L
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var id = _s(d[i][cId]); if (!id) continue;
      m[id.toUpperCase()] = {
        estado:        _s(d[i][cEst]).toUpperCase(),
        guia:          _s(d[i][cGuia]),
        transportista: _s(d[i][cTrans]),
        fechaEntrega:  (d[i][cFecha] instanceof Date) ? _fmtTs(d[i][cFecha]) : _s(d[i][cFecha]),
        idVenta:       _s(d[i][cIdV])
      };
    }
    return m;
  } catch (e) { Logger.log('_cpMasterMap error: ' + e); return {}; }
}


// Un envío está "listo" para el mail de despacho al reseller cuando TODAS sus comandas
// (parts, ej. "123/456") están marcadas DESPACHADO en Comandas Master col F (estado).
// OJO: que ya tengan guía/código de seguimiento (col K) NO alcanza — eso significa que
// Masterchief ya autorizó y va a despachar pronto, no que ya salió físicamente. El código
// viejo usaba la guía como proxy de "despachado" y mandaba el mail antes de tiempo.
function _cpEnvioListoDespacho(parts, master) {
  var guias = [], transs = [], listo = !!(parts && parts.length);
  (parts || []).forEach(function(p) {
    var m = master[p.toUpperCase()];
    if (!m || m.estado !== 'DESPACHADO') listo = false;
    if (m) {
      if (m.guia) guias.push(m.guia);
      if (m.transportista && transs.indexOf(m.transportista) === -1) transs.push(m.transportista);
    }
  });
  return { listo: listo, guias: guias, transs: transs };
}


// Un envío está "autorizado" (Masterchief ya lo procesó y le asignó código de seguimiento)
// cuando TODAS sus comandas tienen guía en Comandas Master col K — esto es ANTERIOR al
// despacho físico (ver _cpEnvioListoDespacho arriba). Dispara el mail #1 al reseller
// ("autorizado, en breve se despacha"); el #2 ("despachado") sigue siendo el de col F.
function _cpEnvioAutorizado(parts, master) {
  var guias = [], transs = [], autorizado = !!(parts && parts.length);
  (parts || []).forEach(function(p) {
    var m = master[p.toUpperCase()];
    if (!m || !m.guia) autorizado = false;
    if (m) {
      if (m.guia) guias.push(m.guia);
      if (m.transportista && transs.indexOf(m.transportista) === -1) transs.push(m.transportista);
    }
  });
  return { autorizado: autorizado, guias: guias, transs: transs };
}


// Busca el PDF de una comanda en la carpeta de Drive (el nombre empieza con el N°).
// Devuelve {name, url, id} o null.
// CACHÉ (positive-only): sólo se cachean los ACIERTOS — una comanda que ya tiene su PDF no
// cambia, así que evitamos re-buscar en Drive en cada refresh. Los "no encontrado" NUNCA se
// cachean, por lo que un PDF recién subido aparece en el siguiente refresh (no queda pegado).
function _cpBuscarPdf(comanda) {
  comanda = _s(comanda);
  if (!comanda || !CP_PDF_FOLDER_ID) return null;
  var cache = null, ckey = 'pdf_' + comanda.replace(/[^0-9A-Za-z]/g, '');
  try {
    cache = CacheService.getScriptCache();
    var hit = cache.get(ckey);
    if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  } catch (_) {}
  var res = _cpBuscarPdfRaw(comanda);
  if (res && cache) { try { cache.put(ckey, JSON.stringify(res), 3600); } catch (_) {} }  // sólo aciertos, TTL 1 h
  return res;
}


// Búsqueda real en Drive (sin caché).
// Convención: el nombre del PDF EMPIEZA con el N° de comanda (ej. "15861196 - Flo Agro.pdf").
function _cpBuscarPdfRaw(comanda) {
  try {
    var q = _s(comanda).replace(/[^0-9A-Za-z]/g, '');  // sólo alfanumérico → no se puede inyectar en la query de Drive
    if (!q) return null;
    var folder = DriveApp.getFolderById(CP_PDF_FOLDER_ID);
    var norm = function(n) { return _s(n).replace(/[^0-9A-Za-z]/g, ''); };
    var prefijo = function(list) { return list.filter(function(f) { return norm(f.getName()).indexOf(q) === 0; }); };

    // 1) Índice de Drive (rápido). Preferimos los que EMPIEZAN con el número de comanda.
    var idx = [];
    try {
      var it = folder.searchFiles('title contains "' + q + '"');
      while (it.hasNext()) idx.push(it.next());
    } catch (eS) { Logger.log('_cpBuscarPdf searchFiles: ' + eS); }
    var hit = prefijo(idx)[0];

    // 2) Fallback: el índice de Drive TARDA en ver un archivo recién subido → searchFiles no lo
    //    encuentra aunque ya esté en la carpeta. Recorremos la carpeta en vivo (getFiles, listado
    //    directo) y matcheamos por prefijo de nombre. Así el PDF recién subido aparece al instante.
    if (!hit) {
      var it2 = folder.getFiles(), live = [];
      while (it2.hasNext()) live.push(it2.next());
      hit = prefijo(live)[0];
    }
    if (!hit) return null;
    return { name: hit.getName(), url: hit.getUrl(), id: hit.getId() };
  } catch (e) {
    Logger.log('_cpBuscarPdf error: ' + e);
    return null;
  }
}


// Olvida el PDF cacheado de una comanda (por si reemplazaste el archivo y querés que se
// re-resuelva ya, sin esperar el TTL de 1 h). Opcional, se corre desde el editor.
function CP_olvidarPdf(comanda) {
  try {
    var c = _s(comanda);
    if (!c) return { ok: false, mensaje: 'Falta la comanda.' };
    CacheService.getScriptCache().remove('pdf_' + c.replace(/[^0-9A-Za-z]/g, ''));
    return { ok: true };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}


// Devuelve los blobs PDF (para adjuntar) de una lista de comandas.
function _cpPdfBlobs(comandas) {
  var blobs = [];
  (comandas || []).forEach(function(p) {
    try {
      var pdf = _cpBuscarPdf(p);
      if (pdf && pdf.id) blobs.push(DriveApp.getFileById(pdf.id).getBlob());
    } catch (e) { Logger.log('_cpPdfBlobs ' + p + ': ' + e); }
  });
  return blobs;
}


// "Documentos definidos" = TODOS los archivos de la carpeta CP_DOCS_FOLDER_ID. Se adjuntan al
// reseller SOLO en el primer envío de la venta. Carpeta sin configurar o vacía → [] (no adjunta
// nada ni falla el mail).
// _cpDocsDefinidosArchivos: lista liviana [{id, name}] (para nombrarlos en el cuerpo del mail).
function _cpDocsDefinidosArchivos() {
  var out = [];
  try {
    var fid = _s(typeof CP_DOCS_FOLDER_ID !== 'undefined' ? CP_DOCS_FOLDER_ID : '');
    if (!fid) return out;   // carpeta sin configurar
    var it = DriveApp.getFolderById(fid).getFiles();
    while (it.hasNext()) { var f = it.next(); out.push({ id: f.getId(), name: f.getName() }); }
  } catch (e) { Logger.log('_cpDocsDefinidosArchivos: ' + e); }
  return out;
}

// _cpDocsDefinidosBlobs: los blobs (para adjuntar).
function _cpDocsDefinidosBlobs() {
  var blobs = [];
  _cpDocsDefinidosArchivos().forEach(function(a) {
    try { blobs.push(DriveApp.getFileById(a.id).getBlob()); } catch (e) { Logger.log('_cpDocsDefinidosBlobs ' + a.name + ': ' + e); }
  });
  return blobs;
}


/* ════════════════════════════════════════════════════════════
   REMITO (AGRAS Y BRUMBY) — planilla aparte, ajena a Masterchief, SOLO LECTURA.
   Col E = IDVenta | Col AJ = N° de remito | Col AK = link al remito.
════════════════════════════════════════════════════════════ */

// { numero, url } del remito de una venta, o null si esa venta todavía no tiene
// remito cargado en esa hoja (o la hoja/pestaña no está disponible).
function _cpRemitoInfo(idVenta) {
  try {
    var key = _s(idVenta).toUpperCase();
    if (!key) return null;
    var ss = _cpSS(CP_REMITOS_SS_ID);
    var h = ss.getSheetByName(CP_REMITOS_TAB);
    if (!h) { Logger.log('Tab remitos no encontrada: ' + CP_REMITOS_TAB); return null; }
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (_s(d[i][4]).toUpperCase() !== key) continue;   // E = IDVenta
      var numero = _s(d[i][35]);  // AJ = N° de remito
      var url    = _s(d[i][36]);  // AK = link al remito
      if (numero || url) return { numero: numero, url: url };
    }
    return null;
  } catch (e) { Logger.log('_cpRemitoInfo error: ' + e); return null; }
}


// ── _CONFIG (parámetros del mail) ──
var CP_CONFIG_DEFAULTS = [
  ['MAIL_APROBACION',       ''],  // ← mail de Sole (aprobación de comanda en Masterchief)
  ['ASUNTO_APROBACION',     'APROBAR MC · {COMANDA} · {IDVENTA}'],
  ['MAIL_DESTINATARIOS',    ''],  // despacho: destinatarios fijos (separá varios con coma)
  ['MAIL_CC',               ''],
  ['MAIL_BCC',              ''],  // copia oculta (reseller + aprobación); separá varios con coma
  ['MAIL_ASUNTO',           'Despacho {IDVENTA} · Comanda {COMANDA} — {CLIENTE}'],
  ['MAIL_REMITENTE_NOMBRE', 'BIDCOMAGRO'],
  ['EMAIL_PRUEBA',          ''],  // MODO PRUEBA: si ponés un mail acá, TODOS los correos (reseller, RTV, Sole, recordatorios) se mandan SOLO a esa dirección (no a los reales). Vaciar para volver a producción.
  ['OCA_TRACKING_URL',      'https://www1.oca.com.ar/OEPTrackingWeb/trackingdetalle.aspx?numero={GUIA}'],
  ['AUTO_MAIL_DESPACHO',    'NO'],  // SI = manda el mail final al reseller+RTV automáticamente al detectar despacho
  ['RECORDATORIO_HORAS',    '24'],  // cada cuántas horas recordarle a Sole una comanda sin despachar
  ['SLA_WARN_HORAS',        '4'],   // semáforo SLA: a partir de cuántas horas pasa a amarillo
  ['SLA_DANGER_HORAS',      '24'],  // semáforo SLA: a partir de cuántas horas pasa a rojo
  ['OPERADORES_AUTORIZADOS','']  ,  // mails (coma) que pueden crear/borrar envíos y mandar mail; VACÍO = todos
  ['MAIL_MAX_POR_10MIN',    '60']   // tope anti-abuso de mails enviados por ventana de 10 min
];

function _cpConfigHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_CONFIG_TAB);
  if (!h) {
    h = ss.insertSheet(CP_CONFIG_TAB);
    h.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]);
    h.getRange(2, 1, CP_CONFIG_DEFAULTS.length, 2).setValues(CP_CONFIG_DEFAULTS);
    h.setFrozenRows(1);
    h.setColumnWidth(1, 210); h.setColumnWidth(2, 460);
  } else {
    // agregar claves nuevas que falten (para hojas ya creadas)
    var d = h.getDataRange().getValues();
    var existentes = {};
    for (var i = 1; i < d.length; i++) { var k = _s(d[i][0]); if (k) existentes[k.toUpperCase()] = true; }
    var faltan = CP_CONFIG_DEFAULTS.filter(function(kv) { return !existentes[kv[0].toUpperCase()]; });
    if (faltan.length) h.getRange(h.getLastRow() + 1, 1, faltan.length, 2).setValues(faltan);
  }
  return h;
}

function _cpConfig() {
  var d = _cpConfigHoja().getDataRange().getValues();
  var m = {};
  for (var i = 1; i < d.length; i++) { var k = _s(d[i][0]); if (k) m[k.toUpperCase()] = _s(d[i][1]); }
  return m;
}


// Detalle de una venta (para el cuerpo del mail): cliente, totales, qué cargar, pedido.
function _cpDetalleVenta(idVenta) {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) return {};
  var col = det.col;
  function get(r, f) { var i = col[f]; return i === -1 ? '' : r[i]; }
  var kitMap = _cpKitMap();
  var key = _s(idVenta).toUpperCase();
  var razonSocial = '', reseller = '', operacion = '', rtv = '', totalUSD = 0, totalARS = 0;
  var lin = {}, linOrd = [];
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var r = rows[i];
    if (_s(get(r, 'idVenta')).toUpperCase() !== key) continue;
    if (!razonSocial) razonSocial = _s(get(r, 'razonSocial'));
    if (!reseller)    reseller    = _s(get(r, 'reseller'));
    if (!operacion)   operacion   = _s(get(r, 'operacion'));
    if (!rtv)         rtv         = _s(get(r, 'rtv'));
    totalUSD += _num(get(r, 'totalUSD'));
    totalARS += _num(get(r, 'totalARS'));
    var sku = _s(get(r, 'sku')), desc = _s(get(r, 'descripcion'));
    var lk = (sku || desc || ('__' + i)).toUpperCase();
    if (!lin[lk]) { lin[lk] = { sku: sku, desc: desc, cant: 0 }; linOrd.push(lk); }
    if (!lin[lk].sku && sku)   lin[lk].sku = sku;
    if (!lin[lk].desc && desc) lin[lk].desc = desc;
    lin[lk].cant += _num(get(r, 'cantidad'));
  }
  var cargarMap = {}, cargarOrden = [];
  linOrd.forEach(function(lk) {
    var lg = lin[lk], cant = Math.round(lg.cant * 100) / 100;
    var kit = kitMap[_kitKey(lg.sku)];
    if (kit && kit.orden.length) {
      kit.orden.forEach(function(cu) { var comp = kit.comps[cu]; _addCargar(cargarMap, cargarOrden, comp.sku, comp.desc, comp.cant * cant, false); });
    } else { _addCargar(cargarMap, cargarOrden, lg.sku, lg.desc, cant, true); }
  });
  return {
    razonSocial: razonSocial, reseller: reseller, operacion: operacion, rtv: rtv,
    totalUSDStr: _fmtUSD(totalUSD), totalARSStr: _fmtARS(totalARS),
    cargar: cargarOrden.map(function(cu) { var it = cargarMap[cu]; var q = Math.round(it.cant*100)/100; return { sku: it.sku, desc: it.desc, cant: _fmtCant(q), cantNum: q }; }),
    pedido: linOrd.map(function(lk) { var lg = lin[lk]; return { sku: lg.sku, desc: lg.desc, cant: _fmtCant(Math.round(lg.cant*100)/100) }; })
  };
}


// Mapa LIVIANO { IDVENTA: {reseller, rtv} } de UNA sola lectura de Ventas.
// Sirve para chequear destinatarios de muchos envíos sin releer toda la hoja Ventas
// por cada uno (lo que colgaba el diagnóstico cuando había muchos envíos con guía
// pendientes de mail). Usa la misma detección de columnas que _cpDetalleVenta.
function _cpVentaResellerRtvMap() {
  var out = {};
  try {
    var rows = _cpHoja().getDataRange().getValues();
    var det = _cpDetectar(rows);
    if (!det) return out;
    var col = det.col;
    function get(r, f) { var i = col[f]; return i === -1 ? '' : r[i]; }
    for (var i = det.headerRow + 1; i < rows.length; i++) {
      var r = rows[i];
      var key = _s(get(r, 'idVenta')).toUpperCase();
      if (!key) continue;
      if (!out[key]) out[key] = { reseller: '', rtv: '' };
      if (!out[key].reseller) out[key].reseller = _s(get(r, 'reseller'));
      if (!out[key].rtv)      out[key].rtv      = _s(get(r, 'rtv'));
    }
  } catch (e) { Logger.log('_cpVentaResellerRtvMap error: ' + e); }
  return out;
}


// Mapa nombre_reseller(normalizado) → {mail, rtv}, leído de la pestaña "Resellers".
//   B = Reseller (nombre) | C = RTV | J = Email  (se detectan por encabezado)
function _cpResellerMap() {
  try {
    var ss = _cpSS(CP_SS_ID);
    var h = ss.getSheetByName(CP_RESELLERS_TAB);
    if (!h) { Logger.log('Tab "' + CP_RESELLERS_TAB + '" no encontrada.'); return {}; }
    var d = h.getDataRange().getValues();
    if (d.length < 2) return {};
    var H = d[0].map(_norm);
    var cName = _cpFindCol(H, ['reseller', 'nombrereseller', 'empresa', 'nombre']);
    var cMail = _cpFindCol(H, ['email', 'mail', 'correo', 'correoelectronico']);
    var cRtv  = _cpFindCol(H, ['rtv']);
    if (cName === -1) cName = 1;  // col B por defecto
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var nom = _s(d[i][cName]); if (!nom) continue;
      var mail = cMail > -1 ? _s(d[i][cMail]) : '';
      var rtv  = cRtv  > -1 ? _s(d[i][cRtv])  : '';
      m[_kitKey(nom)] = { mail: (mail.indexOf('@') > -1 ? mail : ''), rtv: rtv };
    }
    return m;
  } catch (e) { Logger.log('_cpResellerMap error: ' + e); return {}; }
}


// Hoja RTV (Nombre | Mail) en el sheet de log; se crea si no existe.
function _cpRtvHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_RTV_TAB);
  if (!h) {
    h = ss.insertSheet(CP_RTV_TAB);
    h.getRange(1, 1, 1, 2).setValues([['Nombre RTV', 'Mail']]);
    h.setFrozenRows(1);
    h.setColumnWidth(1, 240); h.setColumnWidth(2, 300);
  }
  return h;
}

function _cpRtvMailMap() {
  try {
    var d = _cpRtvHoja().getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var nom = _s(d[i][0]), mail = _s(d[i][1]);
      if (!nom || !mail || mail.indexOf('@') === -1) continue;
      m[_kitKey(nom)] = mail;
    }
    return m;
  } catch (e) { Logger.log('_cpRtvMailMap error: ' + e); return {}; }
}


// Pre-carga la hoja RTV con los nombres distintos de RTV que hay en "Resellers"
// (col C), dejando el mail en blanco para completar a mano. Correr una vez.
function CP_poblarRtvDesdeResellers() {
  var rm = _cpResellerMap();
  var nombres = {};
  Object.keys(rm).forEach(function(k) { var n = _s(rm[k].rtv); if (n) nombres[n.toUpperCase()] = n; });
  var h = _cpRtvHoja();
  var d = h.getDataRange().getValues();
  var exist = {};
  for (var i = 1; i < d.length; i++) { var n = _s(d[i][0]); if (n) exist[n.toUpperCase()] = true; }
  var nuevas = [];
  Object.keys(nombres).forEach(function(u) { if (!exist[u]) nuevas.push([nombres[u], '']); });
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, 2).setValues(nuevas);
  // resumen claro (es idempotente: NO re-agrega los que ya están)
  var dd = h.getDataRange().getValues();
  var total = 0, conMail = 0;
  for (var j = 1; j < dd.length; j++) {
    var nn = _s(dd[j][0]); if (!nn) continue;
    total++; if (_s(dd[j][1]).indexOf('@') > -1) conMail++;
  }
  Logger.log('RTV distintos encontrados en Resellers (col C): ' + Object.keys(nombres).length +
    ' · agregados nuevos: ' + nuevas.length +
    ' · total en la hoja "' + CP_RTV_TAB + '": ' + total +
    ' · con mail cargado: ' + conMail +
    (total ? (conMail < total ? ' → faltan ' + (total - conMail) + ' mails por completar (col B)' : ' ✅ todos con mail') : ' → ⚠️ la hoja está vacía y no se encontraron RTV en Resellers'));
}


// Busca el primer encabezado (normalizado) que coincida con alguno de los alias.
function _cpFindCol(H, alias) {
  for (var a = 0; a < alias.length; a++) { var i = H.indexOf(alias[a]); if (i > -1) return i; }
  // por si el header contiene el alias (ej. "Mail Reseller")
  for (var j = 0; j < H.length; j++) {
    for (var b = 0; b < alias.length; b++) { if (H[j].indexOf(alias[b]) > -1) return j; }
  }
  return -1;
}
