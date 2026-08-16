// @version 2.5
// ============================================================
//  COMANDAS — Mail: templates de envío/aprobación, envío/reenvío,
//  recordatorios a Sole, auto-mail + sus triggers de setup.
//  Extraído de CP_Código.js 1.43 el 2026-07-30 — reorganización sin
//  cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// Destinatarios del mail al reseller de un envío: reseller (Resellers col J) + RTV (hoja RTV) +
// fijos (_CONFIG MAIL_DESTINATARIOS). Devuelve { to:[...], detalle:[...] } deduplicado.
// resellerMap / rtvMap: opcionales. Si se pasan (ya leídos una vez), evita releer
// las hojas Resellers/RTV por cada envío (clave para no colgar el diagnóstico en lote).
function _cpDestinatariosEnvio(det, cfg, resellerMap, rtvMap) {
  var toList = [], detalle = [];
  var rinfo = (resellerMap || _cpResellerMap())[_kitKey(det.reseller)] || {};
  if (rinfo.mail) { toList.push(rinfo.mail); detalle.push('reseller'); }
  var rtvName = rinfo.rtv || det.rtv;
  var mailRtv = rtvName ? (rtvMap || _cpRtvMailMap())[_kitKey(rtvName)] : '';
  if (mailRtv) { toList.push(mailRtv); detalle.push('RTV (' + rtvName + ')'); }
  if (cfg['MAIL_DESTINATARIOS']) { toList.push(cfg['MAIL_DESTINATARIOS']); detalle.push('fijos'); }
  var seen = {}, to = [];
  toList.join(',').split(',').forEach(function(x) { x = x.trim(); var k = x.toLowerCase(); if (x && !seen[k]) { seen[k] = true; to.push(x); } });
  return { to: to, detalle: detalle };
}


// VISTA PREVIA del mail al reseller de un envío (NO envía, NO marca nada).
// Devuelve { ok, html, asunto, destinatarios, a, noDespachado, sinDestino }.
function CP_previewMailEnvio(idVenta, envio) {
  try {
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    var e = null; arr.forEach(function(x) { if (x.envio === envio) e = x; });
    if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };

    var master = _cpMasterMap();
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    var chk = _cpEnvioListoDespacho(parts, master);
    var guias = chk.guias, transs = chk.transs;

    var cfg = _cpConfig();
    var det = _cpDetalleVenta(idVenta);
    var dest = _cpDestinatariosEnvio(det, cfg);
    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);
    var pend = _cpPendingVenta(det, arr);
    var ocaBase = cfg['OCA_TRACKING_URL'] || '';
    // Mismos adjuntos que el envío real: comanda(s) siempre; documentos definidos solo en el 1er envío.
    var esPrimerEnvio = !arr.some(function(x) { return x.envio !== e.envio && x.mailReseller; });
    var docNames = esPrimerEnvio ? _cpDocsDefinidosArchivos().map(function(a) { return a.name; }) : [];
    var remito = _cpRemitoInfo(idVenta);
    var asunto = (cfg['MAIL_ASUNTO'] || 'Despacho {IDVENTA} · Comanda {COMANDA} — {CLIENTE}')
      .replace('{IDVENTA}', idVenta).replace('{COMANDA}', parts.join('/')).replace('{CLIENTE}', det.razonSocial || det.reseller || '');
    var html = _cpMailHtml(idVenta, parts, det, guias, transs.join(', '), ocaBase, pdfs, detEnv, e.notaReseller, pend, envio, docNames, '', remito);

    var bcc = _s(cfg['MAIL_BCC']), cc = _s(cfg['MAIL_CC']);
    return {
      ok: true, html: html, asunto: asunto,
      destinatarios: dest.to.join(', '), a: dest.detalle.join(' + '),
      cc: cc, bcc: bcc, noDespachado: !chk.listo, sinDestino: !dest.to.length
    };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


// Mail al reseller de UN envío. Manual: force=true → reenvía aunque ya se haya mandado.
function CP_enviarMailEnvio(idVenta, envio) {
  var _auth = _cpUsuarioAutorizado();
  if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
  return _cpEnviarEnvioCore(_s(idVenta), _num(envio), true);
}


// Core del mail al reseller (por envío). force=false → NO reenvía si ya se mandó.
function _cpEnviarEnvioCore(idVenta, envio, force) {
  try {
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    var e = null;
    arr.forEach(function(x) { if (x.envio === envio) e = x; });
    if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };
    if (!force && e.mailReseller) return { ok: false, yaEnviado: true, mensaje: 'El mail de este envío ya fue enviado (' + e.mailReseller + ').' };

    var master = _cpMasterMap();
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    if (!parts.length) return { ok: false, mensaje: 'El envío no tiene comanda.' };

    var chk = _cpEnvioListoDespacho(parts, master);
    var guias = chk.guias, transs = chk.transs;
    if (!chk.listo) return { ok: false, mensaje: 'Este envío todavía no está marcado como DESPACHADO en Comandas Master (col F) — la guía ya cargada solo significa que está autorizado.' };

    var cfg = _cpConfig();
    var det = _cpDetalleVenta(idVenta);

    // destinatarios: reseller (Resellers col J) + RTV (col C → hoja RTV) + fijos de _CONFIG
    var dest = _cpDestinatariosEnvio(det, cfg);
    var to = dest.to, detalleDest = dest.detalle;
    if (!to.length) return { ok: false, mensaje: 'No hay destinatarios: falta el mail del reseller (Resellers), del RTV (hoja RTV) o MAIL_DESTINATARIOS en _CONFIG.' };

    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);          // lo que se envió en ESTE envío
    var pend = _cpPendingVenta(det, arr);                    // lo que todavía falta de la venta
    var ocaBase = cfg['OCA_TRACKING_URL'] || '';

    // Si esta venta ya tiene un hilo (del mail #1 "autorizado" de este mismo envío, o de
    // cualquier mail de OTRO envío de la venta), este va como REPLY ahí — así el reseller ve
    // toda la conversación de una venta en un solo lugar en vez de un mail nuevo por cada uno.
    // Si no hay hilo previo, sale como mail nuevo y su hilo queda guardado para encadenar los próximos.
    var hiloPrevio = '';
    arr.forEach(function(x) { if (!hiloPrevio && x.threadIdReseller) hiloPrevio = x.threadIdReseller; });

    // Adjuntos: el PDF de la comanda de ESTE envío siempre; los "documentos definidos"
    // (carpeta CP_DOCS_FOLDER_ID) SOLO en el primer mail de TODA la venta (sea el #1
    // "autorizado" o este #2 "despachado", lo que se haya mandado primero) — si ya hay hilo
    // previo es porque ya se adjuntaron antes.
    var docsArch = !hiloPrevio ? _cpDocsDefinidosArchivos() : [];   // 1 solo listado de la carpeta
    var docNames = docsArch.map(function(a) { return a.name; });      // nombres para el cuerpo del mail

    var remito = _cpRemitoInfo(idVenta);
    var asunto = (cfg['MAIL_ASUNTO'] || 'Despacho {IDVENTA} · Comanda {COMANDA} — {CLIENTE}')
      .replace('{IDVENTA}', idVenta).replace('{COMANDA}', parts.join('/')).replace('{CLIENTE}', det.razonSocial || det.reseller || '');
    var html = _cpMailHtml(idVenta, parts, det, guias, transs.join(', '), ocaBase, pdfs, detEnv, e.notaReseller, pend, envio, docNames, 'despachado', remito);

    var opts = { htmlBody: html, name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO' };
    if (cfg['MAIL_CC'])  opts.cc  = cfg['MAIL_CC'];
    if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];

    var adjuntos = _cpPdfBlobs(parts);
    docsArch.forEach(function(a) { try { adjuntos.push(DriveApp.getFileById(a.id).getBlob()); } catch (eD) { Logger.log('adjuntar doc ' + a.name + ': ' + eD); } });
    if (adjuntos.length) opts.attachments = adjuntos;

    var threadResultante;
    try {
      threadResultante = _cpEnviarDespachoReseller(hiloPrevio, to.join(','), asunto, 'Envío ' + envio + ' · comanda ' + parts.join('/') + ' — ver versión HTML.', opts);
    } catch (se) {
      // registrar el error en la hoja (col Estado) para que quede visible; el auto-trigger reintentará.
      try { _cpEnviosHoja().getRange(e.rowIdx, 9).setValue('MAIL ERROR · ' + _fmtTs(new Date()) + ' · ' + String(se && se.message ? se.message : se)); } catch (_) {}
      return { ok: false, mailError: true, mensaje: 'No se pudo enviar el mail: ' + String(se && se.message ? se.message : se) };
    }
    _cpMarcarMailEnvio(e.rowIdx, guias.join('/'), transs.join(', '), threadResultante);
    _cpAuditar('Mail reseller', idVenta, envio, 'a ' + to.join(', ') + (hiloPrevio ? ' · reply en hilo del envío anterior' : ''));
    return { ok: true, destinatarios: to.join(', '), a: detalleDest.join(' + '), envio: envio };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


function _cpMarcarMailEnvio(rowIdx, guia, transportista, threadId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var h = _cpEnviosHoja();
    if (guia)          h.getRange(rowIdx, 7).setValue(guia);
    if (transportista) h.getRange(rowIdx, 8).setValue(transportista);
    h.getRange(rowIdx, 9).setValue(CP_DESPACHADO);
    h.getRange(rowIdx, 11).setValue('SÍ · ' + _fmtTs(new Date()));
    if (threadId) h.getRange(rowIdx, 14).setValue(threadId);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}


/* ════════════════════════════════════════════════════════════
   MAIL #1 AL RESELLER — "PEDIDO AUTORIZADO": se manda apenas Comandas Master le asigna
   código de seguimiento (guía) al envío. Todavía NO significa que salió físicamente — eso
   lo avisa el mail #2 (_cpEnviarEnvioCore / "DESPACHADO"), que sigue esperando la col F.
   Mismo contenido completo que el mail #2 (chips, guía, PDF adjunto, documentos definidos,
   nota al reseller) — usa el mismo _cpMailHtml con tipo='autorizado", solo cambia el
   encabezado/intro. Mismo criterio de destinatarios/hilo/rate-limit.
════════════════════════════════════════════════════════════ */

// Core del mail #1. force=false → NO reenvía si ya se mandó.
function _cpEnviarMailAutorizadoCore(idVenta, envio, force) {
  try {
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    var e = null;
    arr.forEach(function(x) { if (x.envio === envio) e = x; });
    if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };
    if (!force && e.mailAutorizado) return { ok: false, yaEnviado: true, mensaje: 'El mail de autorización de este envío ya fue enviado (' + e.mailAutorizado + ').' };

    var master = _cpMasterMap();
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    if (!parts.length) return { ok: false, mensaje: 'El envío no tiene comanda.' };

    var chk = _cpEnvioAutorizado(parts, master);
    var guias = chk.guias, transs = chk.transs;
    if (!chk.autorizado) return { ok: false, mensaje: 'Este envío todavía no tiene código de seguimiento (guía) en Comandas Master.' };

    var cfg = _cpConfig();
    var det = _cpDetalleVenta(idVenta);
    var dest = _cpDestinatariosEnvio(det, cfg);
    var to = dest.to, detalleDest = dest.detalle;
    if (!to.length) return { ok: false, mensaje: 'No hay destinatarios: falta el mail del reseller (Resellers), del RTV (hoja RTV) o MAIL_DESTINATARIOS en _CONFIG.' };

    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);
    var pend = _cpPendingVenta(det, arr);
    var ocaBase = cfg['OCA_TRACKING_URL'] || '';

    // Mismo criterio de hilo que el mail de despacho: si la venta ya tiene uno (de otro
    // envío, por ejemplo), este va como reply ahí. "Documentos definidos" solo si este es
    // el primer mail de TODA la venta (sin hilo previo) — igual que en _cpEnviarEnvioCore.
    var hiloPrevio = '';
    arr.forEach(function(x) { if (!hiloPrevio && x.threadIdReseller) hiloPrevio = x.threadIdReseller; });
    var docsArch = !hiloPrevio ? _cpDocsDefinidosArchivos() : [];
    var docNames = docsArch.map(function(a) { return a.name; });

    var remito = _cpRemitoInfo(idVenta);
    var asunto = (cfg['MAIL_ASUNTO_AUTORIZADO'] || 'Pedido autorizado {IDVENTA} · Comanda {COMANDA} — {CLIENTE}')
      .replace('{IDVENTA}', idVenta).replace('{COMANDA}', parts.join('/')).replace('{CLIENTE}', det.razonSocial || det.reseller || '');
    var html = _cpMailHtml(idVenta, parts, det, guias, transs.join(', '), ocaBase, pdfs, detEnv, e.notaReseller, pend, envio, docNames, 'autorizado', remito);

    var opts = { htmlBody: html, name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO' };
    if (cfg['MAIL_CC'])  opts.cc  = cfg['MAIL_CC'];
    if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];

    var adjuntos = _cpPdfBlobs(parts);
    docsArch.forEach(function(a) { try { adjuntos.push(DriveApp.getFileById(a.id).getBlob()); } catch (eD) { Logger.log('adjuntar doc ' + a.name + ': ' + eD); } });
    if (adjuntos.length) opts.attachments = adjuntos;

    var threadResultante;
    try {
      threadResultante = _cpEnviarDespachoReseller(hiloPrevio, to.join(','), asunto, 'Envío ' + envio + ' · comanda ' + parts.join('/') + ' — autorizado, ver versión HTML.', opts);
    } catch (se) {
      return { ok: false, mailError: true, mensaje: 'No se pudo enviar el mail: ' + String(se && se.message ? se.message : se) };
    }
    _cpMarcarMailAutorizado(e.rowIdx, threadResultante);
    _cpAuditar('Mail autorizado reseller', idVenta, envio, 'a ' + to.join(', ') + (hiloPrevio ? ' · reply en hilo existente' : ''));
    return { ok: true, destinatarios: to.join(', '), a: detalleDest.join(' + '), envio: envio };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


// Mail #1 al reseller de UN envío ("autorizado"). Manual: force=true → reenvía aunque ya se haya mandado.
function CP_enviarMailAutorizado(idVenta, envio) {
  var _auth = _cpUsuarioAutorizado();
  if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
  return _cpEnviarMailAutorizadoCore(_s(idVenta), _num(envio), true);
}


function _cpMarcarMailAutorizado(rowIdx, threadId) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var h = _cpEnviosHoja();
    h.getRange(rowIdx, 15).setValue('SÍ · ' + _fmtTs(new Date()));
    // El hilo (col N) lo fija el PRIMER mail que se manda de la venta, sea el #1 o el #2 —
    // no pisar uno que ya exista (podría venir de otro envío de la misma venta).
    if (threadId && !_s(h.getRange(rowIdx, 14).getValue())) h.getRange(rowIdx, 14).setValue(threadId);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}


// Manda el mail de despacho al reseller encadenado con los envíos previos de la misma venta:
// si `threadIdPrevio` viene y el hilo todavía existe, hace REPLY ahí (mismo asunto/conversación
// para el reseller); si no, manda un mail nuevo (mismo chokepoint _cpGmailSend: rate-limit,
// modo prueba, conversión de emojis a ASCII). Devuelve el threadId a guardar para el PRÓXIMO
// envío de esta venta (el del reply si hubo, o el recién creado si se mandó nuevo).
// En modo prueba SIEMPRE manda nuevo: replyAll no admite redirigir el destinatario como sendEmail,
// así que encadenar ahí mandaría el mail de prueba al reseller real.
function _cpEnviarDespachoReseller(threadIdPrevio, to, subject, plain, opts) {
  var prueba = _s(_cpConfig()['EMAIL_PRUEBA']);
  if (threadIdPrevio && !prueba) {
    try {
      var thread = GmailApp.getThreadById(threadIdPrevio);
      var msgs = thread && thread.getMessages();
      if (msgs && msgs.length) {
        if (!_cpRateLimitOk()) throw new Error('Límite de envío de mails alcanzado (' + _cpMailCap() + ' por 10 min, anti-abuso). Reintentá en unos minutos.');
        var o = {}; for (var k in opts) if (opts.hasOwnProperty(k)) o[k] = opts[k];
        if (o.htmlBody) o.htmlBody = _cpHtmlAscii(o.htmlBody);
        msgs[msgs.length - 1].replyAll(plain, o);
        _cpRateLimitInc();
        return threadIdPrevio;
      }
    } catch (eR) {
      Logger.log('_cpEnviarDespachoReseller: reply en hilo ' + threadIdPrevio + ' falló (' + eR + ') → mando mail nuevo');
    }
  }
  _cpGmailSend(to, subject, plain, opts);
  return _cpBuscarThreadRecienEnviado(subject);
}


// Busca en "Enviados" el hilo del mail que se acaba de mandar (por asunto exacto, que incluye
// IDVENTA+comanda y es único) para guardar su threadId y poder encadenar los próximos envíos
// de la misma venta. Best-effort: si el índice de Gmail todavía no lo indexó, devuelve '' y
// simplemente no se pudo encadenar — el próximo envío manda un mail nuevo en vez de reply.
function _cpBuscarThreadRecienEnviado(subject) {
  try {
    var q = 'in:sent subject:"' + String(subject || '').replace(/"/g, '') + '"';
    var threads = GmailApp.search(q, 0, 3);
    return (threads && threads.length) ? threads[0].getId() : '';
  } catch (e) { Logger.log('_cpBuscarThreadRecienEnviado: ' + e); return ''; }
}


// AUTOMÁTICO (trigger horario). Manda los 2 mails al reseller que correspondan por envío:
// #1 "autorizado" apenas Comandas Master le asigna guía, #2 "despachado" cuando col F dice
// DESPACHADO. Mismo interruptor para los 2 (AUTO_MAIL_DESPACHO=SI) — no tiene sentido activar
// uno sin el otro.
function CP_autoMailEnvios() {
  var cfg = _cpConfig();
  var v = _s(cfg['AUTO_MAIL_DESPACHO']).toUpperCase();
  if (v.indexOf('SI') !== 0 && v.indexOf('SÍ') !== 0) { Logger.log('AUTO_MAIL_DESPACHO desactivado en _CONFIG.'); return; }
  var mapAll = _cpEnviosMap(), enviados = 0, enviadosAutorizado = 0;
  Object.keys(mapAll).forEach(function(k) {
    mapAll[k].forEach(function(e) {
      if (!e.mailAutorizado) {
        var ra = _cpEnviarMailAutorizadoCore(k, e.envio, false);
        if (ra && ra.ok) enviadosAutorizado++;
      }
      if (!e.mailReseller) {
        var r = _cpEnviarEnvioCore(k, e.envio, false);
        if (r && r.ok) enviados++;
      }
    });
  });
  Logger.log('CP_autoMailEnvios: ' + enviadosAutorizado + ' mail(s) de autorización + ' + enviados + ' mail(s) de despacho enviado(s).');
  try { CP_actualizarHojaTiempos(); } catch (te) { Logger.log('actualizarHojaTiempos tras autoMail: ' + te); }
}

// Alias por si quedó instalado el trigger viejo.
function CP_autoMailDespacho() { return CP_autoMailEnvios(); }


// Reintenta manualmente TODOS los envíos con algún mail (autorizado o despacho) pendiente
// (incluye los que quedaron en "MAIL ERROR"). Independiente de AUTO_MAIL_DESPACHO.
// Devuelve { ok, enviados, fallidos:[{idVenta, envio, mensaje}] }.
function CP_reintentarFallidos() {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    var mapAll = _cpEnviosMap(), enviados = 0, fallidos = [];
    Object.keys(mapAll).forEach(function(k) {
      mapAll[k].forEach(function(e) {
        if (!_s(e.comanda)) return;                 // sin comanda → no aplica
        if (!e.mailAutorizado) {
          var ra = _cpEnviarMailAutorizadoCore(k, e.envio, false);
          if (ra && ra.ok) enviados++;
        }
        if (e.mailReseller) return;                 // despacho ya salió
        var r = _cpEnviarEnvioCore(k, e.envio, false);
        if (r && r.ok) enviados++;
        else if (r && r.mailError) fallidos.push({ idVenta: k, envio: e.envio, mensaje: r.mensaje || 'error' });
        // los que aún no tienen guía/destinatario se ignoran (no son "fallidos", sólo no están listos)
      });
    });
    if (enviados || fallidos.length) _cpAuditar('Reintentar fallidos', '', '', enviados + ' enviado(s), ' + fallidos.length + ' fallo(s)');
    return { ok: true, enviados: enviados, fallidos: fallidos };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


// Recordatorio a Sole cada X horas (RECORDATORIO_HORAS, def 24) para comandas que siguen
// SIN autorizar en Masterchief (sin guía todavía) — una vez que tiene guía, ya hizo su parte
// (autorizar), aunque el despacho físico y el mail al reseller recién pasen después (ver
// _cpEnvioListoDespacho). Usa ScriptProperties para no repetir antes de X horas.
function CP_recordatoriosSole() {
  var cfg = _cpConfig();
  var sole = cfg['MAIL_APROBACION'];
  if (!sole) { Logger.log('CP_recordatoriosSole: falta MAIL_APROBACION.'); return; }
  var horas = _num(cfg['RECORDATORIO_HORAS']) || 24;
  var master = _cpMasterMap();
  var props = PropertiesService.getScriptProperties();
  var now = Date.now(), enviados = 0;
  var mapAll = _cpEnviosMap();
  Object.keys(mapAll).forEach(function(k) {
    mapAll[k].forEach(function(e) {
      if (!e.fechaTs) return;
      var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
      if (!parts.length) return;
      var tieneGuia = true;
      parts.forEach(function(p) { var m = master[p.toUpperCase()]; if (!m || !m.guia) tieneGuia = false; });
      if (tieneGuia) return;                              // ya autorizado → no molestar a Sole
      if ((now - e.fechaTs) / 3600000 < horas) return;   // todavía no cumple X horas
      var pk = 'rem_' + k + '_' + e.envio;
      var last = _num(props.getProperty(pk));
      if (last && (now - last) / 3600000 < horas) return; // recordado hace < X horas
      var det = _cpDetalleVenta(k);
      var okr = _cpEnviarRecordatorioSole(sole, k, e, cfg, det, Math.floor((now - e.fechaTs) / 3600000));
      if (okr) { props.setProperty(pk, String(now)); enviados++; }
    });
  });
  Logger.log('CP_recordatoriosSole: ' + enviados + ' recordatorio(s) enviado(s).');
}


function _cpEnviarRecordatorioSole(sole, idVenta, e, cfg, det, horas) {
  try {
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);
    var arr = _cpEnviosMap()[String(idVenta).toUpperCase()] || [];
    var pend = _cpPendingVenta(det, arr);   // faltantes de la venta tras este envío
    var nota = 'RECORDATORIO — esta comanda sigue sin despacharse hace ~' + horas + ' h.' + (e.notaAprob ? ' · ' + e.notaAprob : '');
    var asunto = '⏰ RECORDATORIO · ' + (cfg['ASUNTO_APROBACION'] || 'APROBAR MC · {COMANDA} · {IDVENTA}')
      .replace('{COMANDA}', parts.join('/')).replace('{IDVENTA}', idVenta);
    var opts = {
      htmlBody: _cpMailAprobacionHtml(idVenta, parts, det, pdfs, nota, detEnv, e.envio, pend),
      name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO'
    };
    // Mismo RTV en copia que el mail de aprobación original (_cpEnviarMailSoleCore) —
    // el recordatorio es el mismo pedido, solo que Sole todavía no lo atendió.
    var mailRtv = _cpRtvDeReseller(det.reseller);
    if (mailRtv) opts.cc = mailRtv;
    if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];
    var adj = _cpPdfBlobs(parts);
    if (adj.length) opts.attachments = adj;
    _cpGmailSend(sole, asunto, 'Recordatorio: autorizar comanda ' + e.comanda + ' en Masterchief.', opts);
    _cpAuditar('Recordatorio Sole', idVenta, e.envio, 'a ' + sole + ' · ~' + horas + ' h sin despachar');
    return true;
  } catch (err) { Logger.log('_cpEnviarRecordatorioSole ' + idVenta + '#' + e.envio + ': ' + err); return false; }
}


// SETUP: instala el trigger de recordatorios a Sole (revisa cada 6 h; recuerda cada RECORDATORIO_HORAS). Correr una vez.
function CP_setupRecordatorios() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'CP_recordatoriosSole'; });
  if (ya) { Logger.log('El trigger CP_recordatoriosSole ya existe.'); return; }
  ScriptApp.newTrigger('CP_recordatoriosSole').timeBased().everyHours(6).create();
  Logger.log('✅ Trigger de recordatorios instalado (revisa cada 6 h; recuerda cada ' + (_cpConfig()['RECORDATORIO_HORAS'] || 24) + ' h una comanda sin despachar).');
}


// SETUP: instala el trigger horario del envío automático. Correr una vez.
function CP_setupAutoMail() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'CP_autoMailEnvios'; });
  if (ya) { Logger.log('El trigger CP_autoMailEnvios ya existe.'); return; }
  ScriptApp.newTrigger('CP_autoMailEnvios').timeBased().everyMinutes(30).create();
  Logger.log('✅ Trigger automático instalado (cada 30 min). Activá AUTO_MAIL_DESPACHO=SI en _CONFIG para que envíe.');
}


// ── Cuerpos de mail ──
function _cpMailItemsRows(items, esc) {
  return (items || []).map(function(it) {
    return '<tr><td style="padding:5px 8px;border-bottom:1px solid #eee;font-weight:700;text-align:center">' + esc(it.cant) + '×</td>' +
           '<td style="padding:5px 8px;border-bottom:1px solid #eee;font-family:monospace;font-weight:700">' + esc(it.sku || '—') + '</td>' +
           '<td style="padding:5px 8px;border-bottom:1px solid #eee;color:#555">' + esc(it.desc || '') + '</td></tr>';
  }).join('');
}


// Bloque HTML de FALTANTES (lo que queda pendiente de enviar de la venta).
// Devuelve '' si no hay nada pendiente → así no sale nada en el mail.
function _cpPendBloqueHtml(pendientes, esc, titulo) {
  if (!pendientes || !pendientes.length) return '';
  var rows = pendientes.map(function(f){
    return '<tr><td style="padding:5px 8px;border-bottom:1px solid #f3d9d9;font-weight:700;text-align:center;color:#c0392b">'+esc(f.cantidad)+'×</td>'+
           '<td style="padding:5px 8px;border-bottom:1px solid #f3d9d9;font-family:monospace;font-weight:700">'+esc(f.sku||'—')+'</td>'+
           '<td style="padding:5px 8px;border-bottom:1px solid #f3d9d9;color:#555">'+esc(f.descripcion||'')+'</td></tr>';
  }).join('');
  return '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#c0392b;margin:16px 0 8px">'+esc(titulo)+'</div>'+
    '<table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #e6b3b3;border-radius:8px;overflow:hidden;background:#fff7f7">'+rows+'</table>';
}


// Mail al reseller de un ENVÍO — mismo contenido completo (chips, guía, PDF, adjuntos) para
// los 2 mails de la venta; lo único que cambia es el encabezado/intro según `tipo`:
// 'autorizado' (mail #1, apenas hay guía) o 'despachado' (mail #2, default, cuando ya salió).
function _cpMailHtml(idVenta, comandas, det, guias, transportista, ocaBase, pdfs, itemsEnviados, notaReseller, pendientes, envioNum, docsAdjuntos, tipo, remito) {
  function esc(s){ return String(s==null?'':s).replace(/[&<>]/g,function(m){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[m];}); }
  var esAutorizado = (tipo === 'autorizado');
  var headerColor = esAutorizado ? '#6b7280' : '#00a3e0';
  var headerLabel = esAutorizado ? 'BIDCOMAGRO · PEDIDO AUTORIZADO' : 'BIDCOMAGRO · DESPACHO · EN CAMINO';
  var tituloTxt   = esAutorizado ? ' autorizada' : ' en camino';
  var introTxt    = esAutorizado
    ? 'Tu pedido ya fue autorizado y tiene código de seguimiento asignado. Te va a llegar otro aviso apenas salga físicamente.'
    : 'Tu pedido ya salió y está en camino. Guardá este mail para el seguimiento.';
  // Sección explícita de adjuntos: la(s) comanda(s) de este envío + los documentos definidos (1er envío).
  var adjItems = [];
  (pdfs||[]).forEach(function(p){ adjItems.push('📄 Comanda ' + esc(p.comanda) + (p.name ? ' <span style="color:#888">('+esc(p.name)+')</span>' : '')); });
  (docsAdjuntos||[]).forEach(function(n){ adjItems.push('📎 ' + esc(n)); });
  var adjBloque = adjItems.length
    ? '<div style="background:#f0faf4;border:1px solid #b7e3c8;border-left:3px solid #1a9e4a;border-radius:0 8px 8px 0;padding:12px 15px;margin:0 0 16px;font-size:13px;color:#1a1f2e">'
      + '<div style="font-weight:700;margin-bottom:6px">📎 Documentos adjuntos a este correo</div>'
      + '<ul style="margin:0;padding-left:20px;line-height:1.8">' + adjItems.map(function(x){ return '<li>'+x+'</li>'; }).join('') + '</ul>'
      + '</div>'
    : '';
  var guiasHtml = (guias||[]).map(function(g){
    var url = ocaBase ? ocaBase.replace('{GUIA}', encodeURIComponent(g)) : '';
    return url ? '<a href="'+url+'" style="color:#00a3e0;font-weight:700;text-decoration:none">'+esc(g)+' ↗</a>' : '<b>'+esc(g)+'</b>';
  }).join(' &nbsp;·&nbsp; ') || '—';
  var pdfHtml = (pdfs && pdfs.length)
    ? pdfs.map(function(p){ return '<a href="'+esc(p.url)+'" style="color:#00a3e0;font-weight:700;text-decoration:none">📄 '+esc(p.comanda)+' ↗</a>'; }).join(' &nbsp;·&nbsp; ')
    : '—';
  var remitoHtml = (remito && (remito.numero || remito.url))
    ? (remito.url ? '<a href="'+esc(remito.url)+'" style="color:#00a3e0;font-weight:700;text-decoration:none">'+esc(remito.numero || 'Ver remito')+' ↗</a>' : '<b>'+esc(remito.numero)+'</b>')
    : '';
  var notaBloque = _s(notaReseller)
    ? '<div style="background:#eef7ff;border:1px solid #cfe6fb;border-left:3px solid #00a3e0;border-radius:0 8px 8px 0;padding:12px 15px;margin:0 0 16px;font-size:13px;color:#1a1f2e;white-space:pre-wrap"><b>Mensaje:</b> '+esc(notaReseller)+'</div>'
    : '';
  var chip = function(l,v){ return '<tr><td style="padding:4px 0;color:#888;width:150px">'+esc(l)+'</td><td style="padding:4px 0;color:#1a1f2e;font-weight:600">'+v+'</td></tr>'; };

  // Los faltantes NO van en el mail de despacho (reseller+RTV): solo lo que se envía en esta
  // comanda. El seguimiento de pendientes sigue vivo en la UI y en PENDIENTES_ENTREGA.
  return ''+
  '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:640px;margin:0 auto;color:#1a1f2e">'+
    '<div style="background:'+headerColor+';color:#fff;padding:18px 22px;border-radius:10px 10px 0 0">'+
      '<div style="font-size:12px;opacity:.85;letter-spacing:.05em">'+headerLabel + (envioNum ? ' · ENVÍO ' + esc(envioNum) : '') + '</div>'+
      '<div style="font-size:20px;font-weight:800;margin-top:2px">Comanda '+esc((comandas||[]).join(' / '))+esc(tituloTxt)+'</div>'+
    '</div>'+
    '<div style="border:1px solid #e0e3e8;border-top:none;border-radius:0 0 10px 10px;padding:20px 22px">'+
      '<p style="margin:0 0 14px;font-size:13.5px">'+esc(introTxt)+'</p>'+
      notaBloque+
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px">'+
        chip('ID Venta', esc(idVenta))+
        chip('Comanda(s)', esc((comandas||[]).join(' / ')))+
        chip('Razón Social', esc(det.razonSocial||'—'))+
        chip('Reseller', esc(det.reseller||'—'))+
        chip('Operación', esc(det.operacion||'—'))+
        chip('Transportista', esc(transportista||'—'))+
        chip('Guía de seguimiento', guiasHtml)+
        chip('PDF comanda', pdfHtml)+
        (remitoHtml ? chip('Remito', remitoHtml) : '')+
      '</table>'+
      adjBloque+
      '<div style="margin-top:18px;font-size:11px;color:#aaa">Enviado automáticamente desde Comandas · Carga Masterchief.</div>'+
    '</div>'+
  '</div>';
}


// Mail a Sole: aprobar la comanda de un ENVÍO en Masterchief (con detalle de ese envío + PDF adjunto).
function _cpMailAprobacionHtml(idVenta, comandas, det, pdfs, notaAprob, itemsEnviados, envioNum, pendientes) {
  function esc(s){ return String(s==null?'':s).replace(/[&<>]/g,function(m){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[m];}); }
  var pdfHtml = (pdfs && pdfs.length)
    ? pdfs.map(function(p){ return '<a href="'+esc(p.url)+'" style="color:#00a3e0;font-weight:700;text-decoration:none">📄 '+esc(p.comanda)+' ↗</a>'; }).join(' &nbsp;·&nbsp; ')
    : '—';
  var notaBloque = _s(notaAprob)
    ? '<div style="background:#fff6e6;border:1px solid #f3dca6;border-left:3px solid #d4890a;border-radius:0 8px 8px 0;padding:12px 15px;margin:0 0 16px;font-size:13px;color:#1a1f2e;white-space:pre-wrap"><b>Nota del operador:</b> '+esc(notaAprob)+'</div>'
    : '';
  var chip = function(l,v){ return '<tr><td style="padding:4px 0;color:#888;width:150px">'+esc(l)+'</td><td style="padding:4px 0;color:#1a1f2e;font-weight:600">'+v+'</td></tr>'; };

  return ''+
  '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:640px;margin:0 auto;color:#1a1f2e">'+
    '<div style="background:#d4890a;color:#fff;padding:18px 22px;border-radius:10px 10px 0 0">'+
      '<div style="font-size:12px;opacity:.9;letter-spacing:.05em">BIDCOMAGRO · APROBACIÓN MASTERCHIEF' + (envioNum ? ' · ENVÍO ' + esc(envioNum) : '') + '</div>'+
      '<div style="font-size:20px;font-weight:800;margin-top:2px">Autorizar comanda '+esc((comandas||[]).join(' / '))+'</div>'+
    '</div>'+
    '<div style="border:1px solid #e0e3e8;border-top:none;border-radius:0 0 10px 10px;padding:20px 22px">'+
      '<p style="margin:0 0 14px;font-size:13.5px">Se cargó una comanda en Masterchief y necesita tu <b>autorización</b>.</p>'+
      notaBloque+
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px">'+
        chip('Comanda(s)', esc((comandas||[]).join(' / ')))+
        chip('ID Venta', esc(idVenta))+
        chip('Razón Social', esc(det.razonSocial||'—'))+
        chip('Reseller', esc(det.reseller||'—'))+
        chip('Operación', esc(det.operacion||'—'))+
        chip('PDF comanda', pdfHtml)+
      '</table>'+
      '<div style="margin-top:18px;font-size:11px;color:#aaa">Enviado automáticamente desde Comandas · Carga Masterchief.</div>'+
    '</div>'+
  '</div>';
}
