// ============================================================
// @version 2.35
//  WOS — Gestión de hilos Gmail · V-1.0 (Hitos 2–5)
//
//  Hito 1 vive en PORTAL_RESELLER/RS_Pedidos.js.
//  Este módulo requiere que COL.THREAD_ID (col R) ya esté
//  poblado por el Portal al confirmar el pedido.
//
//  TRIGGER a instalar: WOS_instalarTriggerDetector()
//  (ejecutar UNA VEZ desde el editor de Apps Script)
//
//  El resto se reorganizó (2026-07-30, sin cambios funcionales) en:
//    WOS_PDF.js          — generación del PDF de nota de entrega
//    WOS_Herramientas.js — recuperación de tracking/threadId + diagnóstico Carmen
// ============================================================

function _wosPortalHead(titulo) {
  return "<!DOCTYPE html><html><head><meta charset='UTF-8'></head>" +
    "<body style='margin:0;padding:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,\"Segoe UI\",Helvetica,Arial,sans-serif'>" +
    "<table width='100%' cellpadding='0' cellspacing='0' style='background:#f2f4f7;padding:28px 12px'><tr><td>" +
    "<table width='600' align='center' cellpadding='0' cellspacing='0' style='max-width:600px;width:100%'>" +
    "<tr><td style='background:#00a3e0;border-radius:10px 10px 0 0;padding:24px 32px'>" +
      "<table width='100%' cellpadding='0' cellspacing='0'><tr>" +
        "<td><div style='background:#fff;border-radius:5px;padding:4px 10px;display:inline-block'>" +
          "<span style='font-size:14px;font-weight:700;color:#00a3e0;letter-spacing:.1em'>DJI</span></div></td>" +
        "<td align='right'><span style='color:rgba(255,255,255,.85);font-size:11px'>BIDCOMAGRO · Portal Resellers</span></td>" +
      "</tr></table>" +
      "<h1 style='color:#fff;font-size:18px;font-weight:600;margin:18px 0 0;line-height:1.35'>" + titulo + "</h1>" +
    "</td></tr>" +
    "<tr><td style='background:#fff;padding:28px 32px'>";
}

function _wosPortalFoot(footerTxt) {
  return "</td></tr>" +
    "<tr><td style='background:#f9f9f9;border-top:1px solid #eee;border-radius:0 0 10px 10px;padding:16px 32px'>" +
      "<p style='font-size:11px;color:#aaa;margin:0;line-height:1.7'>" + (footerTxt || '') +
      "<br>Generado automaticamente por WOS · BIDCOMAGRO.</p>" +
    "</td></tr></table></td></tr></table></body></html>";
}


function _formatMoneda(n) {
  n = Number(n) || 0;
  var partes  = n.toFixed(2).split('.');
  var entero  = partes[0];
  var cents   = partes[1];
  var res = '';
  var len = entero.length;
  for (var i = 0; i < len; i++) {
    if (i > 0 && (len - i) % 3 === 0) res += '.';
    res += entero[i];
  }
  return res + ',' + cents;
}


function _trackingUrl(transp, track) {
  if (!track) return '';
  var t = String(transp || '').toLowerCase();
  if (t.indexOf('correo') >= 0)   return 'https://www.correoargentino.com.ar';
  if (t.indexOf('oca') >= 0)      return 'https://www.oca.com.ar';
  if (t.indexOf('andreani') >= 0) return 'https://www.andreani.com';
  if (t.indexOf('liqen') >= 0)    return 'https://www.liqen.com.ar';
  if (t.indexOf('via cargo') >= 0 || t.indexOf('viacargo') >= 0)
                                   return 'https://www.viacargo.com.ar';
  if (t.indexOf('credifin') >= 0) return 'https://www.credifin.com.ar';
  return '';
}


function _wosHoja() {
  return SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName(HOJA_PEDIDOS);
}


function _wosGetEmailReseller(nombre) {
  try {
    var d = SpreadsheetApp.openById(MASTER_SS_ID)
              .getSheetByName('Resellers').getDataRange().getValues();
    var n = String(nombre || '').trim().toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][COL_RS.NOMBRE] || '').trim().toLowerCase() === n)
        return String(d[i][COL_RS.EMAIL] || '').trim();
    }
  } catch(e) { Logger.log('_wosGetEmailReseller: ' + e); }
  return '';
}


// Lee todas las filas de un pedido desde NOTAS — resuelve la hoja por prefijo de número
function _wosLeerPedido(numero) {
  var hoja  = _getHojaPorNumero(numero);
  if (!hoja) return { hoja: null, datos: [], reseller: '', envio: '', pago: '', obs: '', threadId: '', tracking: '', items: [] };
  var datos = hoja.getDataRange().getValues();
  var res   = {
    hoja: hoja, datos: datos,
    reseller: '', envio: '', pago: '', obs: '',
    threadId: '', tracking: '',
    items: []
  };
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (!res.reseller) {
      res.reseller = String(datos[i][COL.RESELLER]     || '');
      res.envio    = String(datos[i][COL.ENVIO]        || '');
      res.pago     = String(datos[i][COL.PAGO]         || '');
      res.obs      = String(datos[i][COL.OBS]          || '');
      res.threadId = String(datos[i][COL.THREAD_ID]    || '').trim();
      res.tracking = String(datos[i][COL.TRACKING]     || '').trim();
    }
    res.items.push({
      row:      i + 1,                                        // 1-indexed para getRange
      sku:      String(datos[i][COL.SKU]       || ''),
      desc:     String(datos[i][COL.DESC]      || ''),
      cantSol:  Number(datos[i][COL.CANT_SOL]) || 0,
      cantDesp: Number(datos[i][COL.CANT_DESP])|| 0,
      cantPend: Number(datos[i][COL.CANT_PEND])|| 0,         // valor calculado por fórmula
      precio:   Number(datos[i][COL.PRECIO])   || 0,
      estado:   String(datos[i][COL.ESTADO]    || '')
    });
  }
  return res;
}


// Actualiza ESTADO y FECHA_ESTADO para todas las filas del pedido.
// Limpia la validación de celda antes de escribir para que cualquier
// estado definido en el código funcione sin depender del dropdown del sheet.
function _wosSetEstado(hoja, datos, numero, nuevoEstado) {
  var ahora = new Date();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    var rEst = hoja.getRange(i + 1, COL.ESTADO + 1);
    rEst.clearDataValidations();
    rEst.setValue(nuevoEstado);
    hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
  }
}


// Cambia ESTADO solo para filas cuyo estado actual sea estadoFiltro.
// Las demás filas del pedido no se modifican.
function _wosSetEstadoFiltrado(hoja, datos, numero, estadoFiltro, nuevoEstado) {
  var ahora = new Date();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (String(datos[i][COL.ESTADO] || '').trim() !== estadoFiltro) continue;
    var rEst = hoja.getRange(i + 1, COL.ESTADO + 1);
    rEst.clearDataValidations();
    rEst.setValue(nuevoEstado);
    hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
  }
}


// Para la notificación de faltante: fija estado por SKU.
// SKUs en skuFaltSet → estFalt (En_Espera_Reseller)
// SKUs restantes del mismo pedido → estDisp (Preparado)
function _wosSetEstadoPorSku(hoja, datos, numero, skuFaltSet, estFalt, estDisp) {
  var ahora = new Date();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    var sku = String(datos[i][COL.SKU] || '').trim().toUpperCase();
    var nuevoEst = skuFaltSet[sku] ? estFalt : estDisp;
    var rEst = hoja.getRange(i + 1, COL.ESTADO + 1);
    rEst.clearDataValidations();
    rEst.setValue(nuevoEst);
    hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
  }
}


// Pedidos OT con líneas Reservado_Consolidar (ítems ya disponibles, retenidos a propósito
// por la Opción A de WOS_procesarRespuestaManual): una vez que el pedido YA NO tiene
// ninguna fila Backorder (el faltante real se resolvió, por cualquier vía — cron de avisos
// de ingreso, recepción manual, respuesta de ingreso), las filas Reservado_Consolidar pasan
// a Preparado junto — el pedido queda completo para un solo envío. `datos` debe ser una
// lectura fresca (post-flush) del caller, si no puede evaluar estados ya viejos.
function _wosLiberarConsolidarSiSinBackorder(hoja, datos, numero) {
  var quedaBackorder = false;
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (String(datos[i][COL.ESTADO] || '').trim() === EST.BACKORDER) { quedaBackorder = true; break; }
  }
  if (quedaBackorder) return;
  _wosSetEstadoFiltrado(hoja, datos, numero, EST.RESERVADO_CONSOLIDAR, EST.PREPARADO);
}


// Responde DENTRO del hilo pero apuntando a los destinatarios ORIGINALES del pedido, no a quien
// haya quedado en el ÚLTIMO mensaje. Problema que resuelve: thread.replyAll() responde al último
// mensaje del hilo; si alguien tuvo una conversación aparte con el reseller, el 2º envío / mail de
// administración / confirmación le llegaba a ESA gente y NO al reseller/facturación originales.
// Solución: (1) respondemos al PRIMER mensaje (getMessages()[0] → destinatarios originales del hilo),
// (2) forzamos en CC los que SÍ o SÍ deben recibir (reseller + los que pase el caller), sin duplicar,
// así reciben aunque el primer mensaje no los tuviera. Devuelve true si respondió en el hilo, false
// si no hay hilo/mensajes (el caller hace el fallback con sendEmail explícito).
function _wosReplyHiloOriginal(threadId, plainBody, opts, mustCc) {
  var thread = GmailApp.getThreadById(String(threadId || '').trim());
  if (!thread) return false;
  var msgs = thread.getMessages();
  if (!msgs || !msgs.length) return false;
  var seen = {}, ccList = [];
  function _addCc(v) {
    String(v || '').split(',').forEach(function(x) {
      x = x.trim(); var key = x.toLowerCase();
      if (x && !seen[key]) { seen[key] = true; ccList.push(x); }
    });
  }
  _addCc(opts && opts.cc);
  (mustCc || []).forEach(_addCc);
  var o = {};
  for (var k in opts) if (opts.hasOwnProperty(k)) o[k] = opts[k];
  if (ccList.length) o.cc = ccList.join(',');
  msgs[0].replyAll(plainBody, o);   // primer mensaje = destinatarios originales, ignora conversaciones aparte
  return true;
}


// Lookup de metadata del reseller desde hoja master
function _wosGetResellerMeta(nombre) {
  try {
    var d = SpreadsheetApp.openById(MASTER_SS_ID)
              .getSheetByName('Resellers').getDataRange().getValues();
    var n = String(nombre || '').trim().toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() === n) {
        return {
          nombre:    String(d[i][0] || '').trim(),
          cuit:      String(d[i][1] || '').trim(),
          direccion: String(d[i][2] || '').trim(),
          telefono:  String(d[i][6] || '').trim()
        };
      }
    }
  } catch(e) { Logger.log('_wosGetResellerMeta: ' + e); }
  return { nombre: String(nombre || ''), cuit: '', direccion: '', telefono: '' };
}


// Devuelve el próximo número de nota para un pedido ("01", "02", ...) buscando en la carpeta Drive
function _wosNextNotaNum(numero) {
  try {
    var folder = DriveApp.getFolderById(_wosConfig().pdfFolderId);
    var prefix = 'NE_' + numero + '-';
    // searchFiles filtra en Drive — mucho más rápido que iterar getFiles() completo
    var files  = folder.searchFiles('title contains "' + prefix + '"');
    var max    = 0;
    while (files.hasNext()) {
      var fname  = files.next().getName();
      if (fname.indexOf(prefix) === 0) {
        var parsed = parseInt(fname.substring(prefix.length)) || 0;
        if (parsed > max) max = parsed;
      }
    }
    var n = max + 1;
    return n < 10 ? '0' + n : String(n);
  } catch(e) {
    Logger.log('_wosNextNotaNum: ' + e);
    return '01';
  }
}


// ─────────────────────────────────────────────────────────────
// Idempotencia: devuelve el resultado previo si el token ya se procesó, o null.
function _wosIdempotResultado(token) {
  if (!token) return null;
  try {
    var prev = CacheService.getScriptCache().get('wosdesp_' + token);
    return prev ? JSON.parse(prev) : null;
  } catch(e) { return null; }
}

function _wosIdempotGuardar(token, resultado) {
  if (!token) return;
  try { CacheService.getScriptCache().put('wosdesp_' + token, JSON.stringify(resultado), 600); } catch(e) {} // TTL 10 min
}

// Envuelve una acción con lock + idempotencia. Serializa ejecuciones y, si el
// mismo token ya se procesó (doble-click, dos pestañas, reintento), devuelve el
// resultado previo sin re-ejecutar. Sólo cachea resultados ok:true.
// Uso: return _wosLockIdempot(reqToken, function() { ...body...; return {ok:true}; });
function _wosLockIdempot(reqToken, fn) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch (eLock) { return { ok: false, error: 'Otra operación está en curso. Esperá unos segundos y reintentá.' }; }
  try {
    var prev = _wosIdempotResultado(reqToken);
    if (prev) { Logger.log('_wosLockIdempot: token repetido ' + reqToken + ' → resultado previo'); return prev; }
    var res = fn();
    if (res && res.ok) _wosIdempotGuardar(reqToken, res);
    return res;
  } catch (e) {
    Logger.log('_wosLockIdempot ERROR: ' + e);
    return { ok: false, error: e.toString() };
  } finally {
    try { lock.releaseLock(); } catch (eR) {}
  }
}


// Parsea los bins guardados en col AA (UBIC_PREP) al preparar. Tolerante: acepta el
// JSON [{bin,cant}], devuelve [] si está vacío o mal formado. Normaliza cant a número.
function _parseUbicPrep(raw) {
  var s = String(raw || '').trim();
  if (!s) return [];
  try {
    var arr = JSON.parse(s);
    if (Object.prototype.toString.call(arr) !== '[object Array]') return [];
    var out = [];
    for (var i = 0; i < arr.length; i++) {
      var b = arr[i] || {};
      var bin  = String(b.bin || '').trim();
      var cant = Number(b.cant) || 0;
      if (bin && cant > 0) out.push({ bin: bin, cant: cant });
    }
    return out;
  } catch (e) {
    Logger.log('_parseUbicPrep: JSON inválido → ' + s);
    return [];
  }
}


function WOS_despacharCompleto(numero, despachos, transportista, bultos, costoEnvio, operario, reqToken) {
  // Lock de script: serializa los despachos para que dos ejecuciones (doble-click,
  // dos pestañas) no corran en paralelo y dupliquen nota de entrega + mail.
  var _lock = LockService.getScriptLock();
  try { _lock.waitLock(30000); }
  catch (eLock) { return { ok: false, error: 'Otra operación de despacho está en curso. Esperá unos segundos y reintentá.' }; }
  try {
    // ¿este despacho ya se procesó? (mismo token) → devolver el resultado previo, sin re-ejecutar
    var _prevRes = _wosIdempotResultado(reqToken);
    if (_prevRes) { Logger.log('WOS_despacharCompleto: token repetido ' + reqToken + ' → devuelvo resultado previo'); return _prevRes; }

    var ped = _wosLeerPedido(numero);
    if (!ped.reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    var esRetiroDesp = String(ped.envio || '').toLowerCase().indexOf('retiro') >= 0;
    // Sin threadId: el bloque try/catch de más abajo ya envía email nuevo como fallback.

    // Cliente externo (carga super-RTV): puede no figurar en Resellers → sin email propio.
    // El aviso de despacho sale por replyAll al hilo ancla (col R), que incluye al cliente si
    // dejó su mail (va en CC del pedido) + facturación. El email directo es solo fallback (abajo).
    // Solo se aborta si NO hay email NI hilo NI es retiro: no habría ningún canal de aviso.
    var email = _wosGetEmailReseller(ped.reseller);
    if (!email && !ped.threadId && !esRetiroDesp) return { ok: false, error: 'Sin email ni hilo para avisar el despacho de: ' + ped.reseller };

    operario    = String(operario || '');
    var transp  = String(transportista || '').trim();
    bultos      = bultos || [];
    var costo   = Number(costoEnvio) || 0;
    // Computar tracking string y peso total desde los bultos
    var trackParts = [];
    var pesoTotal  = 0;
    for (var bi = 0; bi < bultos.length; bi++) {
      var btTrack = String(bultos[bi].tracking || '').trim();
      if (btTrack) trackParts.push(btTrack);
      pesoTotal += Number(bultos[bi].peso) || 0;
    }
    var track = trackParts.join(' | ');
    var peso  = Math.round(pesoTotal * 1000) / 1000;
    var tz     = Session.getScriptTimeZone();
    var fecha  = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    var ahora  = new Date();

    var despMap    = {};
    var serialMap  = {};
    var cajaMap    = {};
    // Los bins (ubicaciones WMS) ya NO llegan del modal de despacho: se eligieron al
    // preparar y quedaron guardados en col AA (UBIC_PREP) de cada fila. Acá se leen y
    // se aplica el descuento. Ver _parseUbicPrep().
    for (var d = 0; d < despachos.length; d++) {
      despMap[despachos[d].row]   = Number(despachos[d].cantDesp) || 0;
      serialMap[despachos[d].row] = String(despachos[d].seriales || '').trim();
      cajaMap[despachos[d].row]   = despachos[d].cajaIdx !== undefined ? Number(despachos[d].cajaIdx) : 0;
    }

    var carmenSS    = null;
    var carmenHoja  = null;
    var carmenUbicH = null;
    var _carmenErr  = '';
    // Reintento por si openById falla de forma transitoria (blip de Google).
    for (var _tryC = 0; _tryC < 3 && !carmenHoja; _tryC++) {
      try {
        carmenSS    = SpreadsheetApp.openById(CARMEN_SS_ID);
        carmenHoja  = carmenSS.getSheetByName('Entregados');
        carmenUbicH = carmenSS.getSheetByName(CARMEN_UBICACIONES_TAB);
      } catch(eC) { _carmenErr = eC.toString(); Utilities.sleep(600); }
    }
    // El descuento de stock (línea en "Entregados") es OBLIGATORIO. Si no se puede escribir,
    // se ABORTA el despacho ACÁ — antes de tocar el pedido y antes de mandar el mail — para
    // NUNCA despachar sin descontar (antes fallaba en silencio y el mail salía igual).
    if (!carmenSS || !carmenHoja) {
      Logger.log('WOS_despacharCompleto: Carmen inaccesible. err=' + _carmenErr +
        ' abrio=' + (!!carmenSS) + ' Entregados=' + (!!carmenHoja));
      return { ok: false, error: 'NO se despach\xf3 nada: WOS no pudo escribir el descuento de stock en Carmen ' +
        '(' + (carmenSS ? 'no existe la pesta\xf1a "Entregados"' : 'no se pudo abrir la planilla Carmen') + ')' +
        (_carmenErr ? ' \x2014 ' + _carmenErr : '') + '. Revis\xe1 permisos/nombre de pesta\xf1a y reintent\xe1. ' +
        'Ejecut\xe1 WOS_diagnosticoCarmen desde el editor para ver el detalle.' };
    }

    // Calcular número de nota antes del loop para poder escribirlo en col P por ítem
    var notaNumStr   = _wosNextNotaNum(numero);
    var notaEntrega  = 'NE_' + numero + '-' + notaNumStr;

    var itemsDesp = [];
    var filasDesp = [];   // filas (1-indexed) donde dispNow > 0 para escribir NE_URL post-PDF
    var totalUSD  = 0;

    for (var i = 1; i < ped.datos.length; i++) {
      if (String(ped.datos[i][COL.NUMERO] || '').trim() !== numero) continue;

      // Despacho aditivo: el modal envía cuánto se despacha AHORA;
      // el sheet guarda el acumulado total (oldCantDesp + dispNow).
      var oldCantDesp = Number(ped.datos[i][COL.CANT_DESP]) || 0;
      var dispNow     = (despMap[i + 1] !== undefined) ? (Number(despMap[i + 1]) || 0) : 0;
      var cantFinal   = oldCantDesp + dispNow;

      if (dispNow > 0) {
        // CANT_DESP aislado (col 6)
        ped.hoja.getRange(i + 1, COL.CANT_DESP + 1).setValue(cantFinal);
        // Tracking ACUMULATIVO: si la fila ya tenía código(s) de un despacho anterior,
        // se agregan los nuevos sin pisar ni duplicar (antes se perdía el seguimiento previo).
        var mergedTrack = _wosMergeTracking(ped.datos[i][COL.TRACKING], track);
        // FECHA_DESPACHO (col O), NOTA_ENTREGA (col P), TRACKING (col Q) — bloque contiguo
        ped.hoja.getRange(i + 1, COL.FECHA_DESPACHO + 1, 1, 3).setValues([[ahora, notaEntrega, mergedTrack]]);
        // FECHA_ESTADO (col 19), TRANSPORTISTA_DESP (col 20), COSTO_ENVIO (col 21), PESO_ENVIO (col 22) — bloque contiguo
        ped.hoja.getRange(i + 1, COL.FECHA_ESTADO + 1, 1, 4).setValues([[ahora, transp, costo > 0 ? costo : '', peso > 0 ? peso : '']]);
        if (operario) ped.hoja.getRange(i + 1, COL.OPERARIO + 1).setValue(operario);
        var rowSeriales = serialMap[i + 1] || '';
        if (rowSeriales) ped.hoja.getRange(i + 1, COL.SERIALES + 1).setValue(rowSeriales);
        filasDesp.push(i + 1);
      }

      if (dispNow > 0 && carmenHoja) {
        var _skuDesp  = String(ped.datos[i][COL.SKU]  || '').trim().toUpperCase();
        // Bins elegidos al preparar (col AA). Se limpian tras descontar para que un
        // 2° despacho de la misma fila (backorder) no vuelva a descontar del WMS.
        var _ubicBins = _parseUbicPrep(ped.datos[i][COL.UBIC_PREP]);
        var _ubicStr  = _ubicBins.map(function(b){ return b.bin + '\xd7' + b.cant; }).join(', ');
        carmenHoja.appendRow([_skuDesp, String(ped.datos[i][COL.DESC] || ''), dispNow, String(numero || ''), '', '', fecha, _ubicStr]);
        // Descontar de cada bin en multi-bin; leer UBICACIONES una sola vez por item
        if (_ubicBins.length && carmenUbicH) {
          var _dU = carmenUbicH.getDataRange().getValues();
          var _rowsToDelete = [];
          for (var ub = 0; ub < _ubicBins.length; ub++) {
            var _binKey  = String(_ubicBins[ub].bin  || '').trim().toUpperCase();
            var _binCant = Number(_ubicBins[ub].cant) || 0;
            if (!_binKey || _binCant <= 0) continue;
            for (var ui = 1; ui < _dU.length; ui++) {
              if (String(_dU[ui][0] || '').trim().toUpperCase() !== _skuDesp) continue;
              if (String(_dU[ui][1] || '').trim().toUpperCase() !== _binKey)  continue;
              var _nueva = Math.max(0, (parseFloat(_dU[ui][2]) || 0) - _binCant);
              if (_nueva === 0) {
                _rowsToDelete.push(ui + 1);
                _dU[ui][2] = 0;
              } else {
                carmenUbicH.getRange(ui + 1, 3).setValue(_nueva);
                _dU[ui][2] = _nueva;
              }
              break;
            }
          }
          _rowsToDelete.sort(function(a, b) { return b - a; });
          for (var dr = 0; dr < _rowsToDelete.length; dr++) carmenUbicH.deleteRow(_rowsToDelete[dr]);
        }
        // Consumir los bins de esta fila: se limpia col AA para que un despacho
        // posterior de la misma fila no vuelva a descontar del WMS.
        if (_ubicBins.length) ped.hoja.getRange(i + 1, COL.UBIC_PREP + 1).setValue('');
      }

      if (dispNow > 0) {
        var prec = Number(ped.datos[i][COL.PRECIO]) || 0;
        itemsDesp.push({
          sku:      String(ped.datos[i][COL.SKU]  || ''),
          desc:     String(ped.datos[i][COL.DESC] || ''),
          cantDesp: dispNow, precio: prec,
          // Si el modal de despacho no reingresó seriales, usar los ya registrados al preparar (col Y)
          seriales: (serialMap[i + 1] && String(serialMap[i + 1]).length) ? serialMap[i + 1] : String(ped.datos[i][COL.SERIALES] || ''),
          cajaIdx:  cajaMap[i + 1] !== undefined ? cajaMap[i + 1] : 0
        });
        totalUSD += dispNow * prec;
      }
    }

    if (!itemsDesp.length) return { ok: false, error: 'Ningún ítem con cantidad > 0.' };

    var pdfResult  = _wosGenerarPDF(numero, notaNumStr, ped.reseller, itemsDesp, fecha, transp, bultos, costo);
    var pdfBlob    = pdfResult ? pdfResult.blob : null;
    var pdfUrl     = pdfResult ? pdfResult.url  : '';
    var sheetUrl   = pdfResult ? (pdfResult.sheetUrl || '') : '';

    // Guardar link de la Nota de Entrega en col W para cada ítem despachado
    if (pdfUrl) {
      for (var fi = 0; fi < filasDesp.length; fi++) {
        ped.hoja.getRange(filasDesp[fi], COL.NE_URL + 1).setValue(pdfUrl);
      }
    }

    // ── Email al reseller: misma estructura visual que Portal Reseller ─
    // Tabla de ítems estilo Portal Reseller
    var tbodyRows = '';
    for (var ri = 0; ri < itemsDesp.length; ri++) {
      var itR = itemsDesp[ri];
      var serRow = itR.seriales
        ? "<tr><td colspan='4' style='padding:3px 10px 8px 10px;font-size:10px;color:#777;border-bottom:1px solid #eee'>" +
          "<span style='font-weight:600;color:#555'>N\xba serie:</span> " + itR.seriales + "</td></tr>"
        : '';
      tbodyRows +=
        "<tr>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";font-family:Consolas,monospace;font-size:12px;color:#00a3e0'>" + (itR.sku  || '—') + "</td>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";font-size:12px;color:#333'>"                                   + (itR.desc || '—') + "</td>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";text-align:center;font-weight:700;color:#333'>"                + itR.cantDesp       + "</td>" +
        "<td style='padding:8px 10px" + (itR.seriales ? '' : ';border-bottom:1px solid #eee') + ";text-align:right;font-size:12px;color:#555'>"                 + (itR.precio > 0 ? 'USD ' + _formatMoneda(Number(itR.precio)) : '—') + "</td>" +
        "</tr>" + serRow;
    }
    var tablaHtml =
      "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:1px solid #dde3ea;border-radius:8px;overflow:hidden'>" +
      "<thead><tr style='background:#f0f5fa'>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>SKU</th>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>Descripcion</th>" +
        "<th style='padding:9px 10px;text-align:center;font-size:11px;font-weight:700;color:#888;width:60px'>Cant.</th>" +
        "<th style='padding:9px 10px;text-align:right;font-size:11px;font-weight:700;color:#888;width:90px'>P. Unit.</th>" +
      "</tr></thead><tbody>" + tbodyRows + "</tbody></table>";

    // Total
    var totalHtml =
      "<div style='text-align:right;margin-top:10px'>" +
        "<span style='font-size:13px;color:#444'>Total despachado: </span>" +
        "<strong style='font-size:15px;color:#00a3e0'>USD " + _formatMoneda(totalUSD) + "</strong>" +
        "<div style='font-size:10px;color:#999;margin-top:3px'>Precios expresados no incluyen impuestos</div>" +
      "</div>";

    // Chips de logística por bulto
    var chipBase = "background:#f0f5fa;border:1px solid #dde3ea;border-radius:6px;padding:8px 14px;font-size:12px";
    var chipLbl  = "color:#888;font-size:10px;text-transform:uppercase;letter-spacing:.05em;display:block;margin-bottom:2px";
    var chipsHtml = "<div style='margin-top:14px;display:flex;gap:10px;flex-wrap:wrap'>" +
      (transp ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Transportista</span><strong style='color:#1a1a2e'>" + transp + "</strong></div>" : '') +
      (ped.pago ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Forma de pago</span><strong style='color:#1a1a2e'>" + ped.pago + "</strong></div>" : '') +
      (costo > 0 ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Costo de Envío (ARS)</span><strong style='color:#1a1a2e'>$ " + _formatMoneda(costo) + "</strong></div>" : '') +
      (peso  > 0 ? "<div style='" + chipBase + "'><span style='" + chipLbl + "'>Peso total</span><strong style='color:#1a1a2e'>" + peso + " kg</strong></div>" : '');
    for (var ci = 0; ci < bultos.length; ci++) {
      var bt = bultos[ci];
      if (!bt.tracking && !(bt.peso > 0)) continue;
      var bLbl = bultos.length > 1 ? 'Bulto ' + (ci + 1) : 'N. de seguimiento';
      var bVal = (bt.tracking || '—') + (bt.peso > 0 ? ' &nbsp;·&nbsp; ' + bt.peso + ' kg' : '');
      chipsHtml += "<div style='" + chipBase + "'><span style='" + chipLbl + "'>" + bLbl + "</span><strong style='color:#1a1a2e;font-family:Consolas,monospace;font-size:11px'>" + bVal + "</strong></div>";
    }
    chipsHtml += "</div>";

    // Botones de rastreo (uno por bulto con tracking)
    var rastrearBtns = '';
    for (var rbi = 0; rbi < bultos.length; rbi++) {
      var rbt = bultos[rbi];
      if (!rbt.tracking) continue;
      var rUrl = _trackingUrl(transp, rbt.tracking);
      if (!rUrl) continue;
      var rLbl = bultos.length > 1 ? 'Rastrear Bulto ' + (rbi + 1) : 'Rastrear envío';
      rastrearBtns +=
        "<a href='" + rUrl + "' target='_blank' style='display:inline-flex;align-items:center;gap:8px;padding:10px 22px;" +
        "background:#00a3e0;color:#fff;border-radius:6px;text-decoration:none;font-size:13px;font-weight:600;margin:4px'>" +
        rLbl + " &rarr;</a>";
    }
    var rastrearBtn = rastrearBtns
      ? "<div style='margin-top:16px;text-align:center;display:flex;flex-wrap:wrap;justify-content:center;gap:6px'>" + rastrearBtns + "</div>"
      : '';

    // Botón Drive (igual al botón PDF del Portal Reseller)
    var driveBtn = pdfUrl
      ? "<div style='margin-top:12px;text-align:center'>" +
          "<a href='" + pdfUrl + "' target='_blank' style='display:inline-flex;align-items:center;gap:8px;padding:10px 22px;" +
          "background:#fff;border:1px solid #dde3ea;color:#00a3e0;border-radius:6px;text-decoration:none;" +
          "font-size:13px;font-weight:600'>Ver / Descargar Nota de Entrega &rarr;</a>" +
        "</div>"
      : '';

    var obsHtml = ped.obs
      ? "<div style='margin-top:14px;background:#fffbe6;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px'>" +
          "<strong style='font-size:11px;color:#7a5800;text-transform:uppercase;letter-spacing:.06em'>Observaciones</strong>" +
          "<p style='margin:5px 0 0;font-size:13px;color:#444'>" + ped.obs + "</p>" +
        "</div>"
      : '';

    // Confirmación de recepción con botones (mide la "precisión de resultado"). Los links van al
    // webapp ?page=confirma_entrega → _doGetConfirmaEntrega registra ok/problema en WOS_QA.
    var _confUrl = '';
    try { _confUrl = ScriptApp.getService().getUrl(); } catch(eCU) { Logger.log('WOS confirmar URL: ' + eCU); }
    var _urlOk   = _confUrl ? _confUrl + '?page=confirma_entrega&num=' + encodeURIComponent(numero) + '&r=ok'   : '';
    var _urlProb = _confUrl ? _confUrl + '?page=confirma_entrega&num=' + encodeURIComponent(numero) + '&r=prob' : '';
    var confirmEntregaHtml = _confUrl
      ? "<div style='margin-top:18px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:14px 16px;text-align:center'>" +
          "<p style='margin:0 0 12px;font-size:13px;color:#0369a1;line-height:1.6;font-weight:600'>Cuando recibas el paquete, avisanos c\xf3mo lleg\xf3:</p>" +
          "<div style='display:inline-block'>" +
            "<a href='" + _urlOk + "' target='_blank' style='display:inline-block;padding:11px 20px;margin:4px;background:#00875a;color:#fff;border-radius:6px;text-decoration:none;font-size:13px;font-weight:700'>&#x2705; Recib\xed todo bien</a>" +
            "<a href='" + _urlProb + "' target='_blank' style='display:inline-block;padding:11px 20px;margin:4px;background:#ffffff;border:1px solid #e0b400;color:#7a5800;border-radius:6px;text-decoration:none;font-size:13px;font-weight:700'>&#x26A0;&#xFE0F; Reportar un problema</a>" +
          "</div>" +
          "<p style='margin:12px 0 0;font-size:11px;color:#7a8794'>O respond\xe9 este correo con \"Recibido\". \xa1Gracias!</p>" +
        "</div>"
      : "<div style='margin-top:18px;background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:12px 16px'>" +
          "<p style='margin:0;font-size:12px;color:#0369a1;line-height:1.6'>" +
            "Cuando recibas el paquete, <strong>respond\xe9 este correo con la palabra \"Recibido\"</strong> " +
            "para confirmar la entrega. Gracias!" +
          "</p>" +
        "</div>";

    // Detectar tipo de envío:
    // hayDespPrevio      = algún ítem ya tenía CANT_DESP > 0 antes de este despacho → 2º envío
    // hayBackorderPostDesp = quedarán ítems en Backorder después de este despacho
    var hayDespPrevio       = false;
    var hayBackorderPostDesp = false;
    for (var bk = 1; bk < ped.datos.length; bk++) {
      if (String(ped.datos[bk][COL.NUMERO] || '').trim() !== numero) continue;
      var oldDespBk  = Number(ped.datos[bk][COL.CANT_DESP]) || 0;
      if (oldDespBk > 0) { hayDespPrevio = true; }
      var dispNowBk  = (despMap[bk + 1] !== undefined) ? (Number(despMap[bk + 1]) || 0) : 0;
      if (dispNowBk > 0) {
        var cantSolBk    = Number(ped.datos[bk][COL.CANT_SOL])    || 0;
        var cantCancelBk = Number(ped.datos[bk][COL.CANT_CANCEL]) || 0;
        var cantFinalBk  = oldDespBk + dispNowBk;
        // Sólo hay backorder si falta despachar algo NO cancelado (cantSol − Z). Lo cancelado por el
        // reseller (Opción B) no es backorder → no dispara 2º envío ni el aviso de "quedan ítems".
        if (cantFinalBk < (cantSolBk - cantCancelBk)) { hayBackorderPostDesp = true; }
      } else {
        var estBk = String(ped.datos[bk][COL.ESTADO] || '').trim();
        if (estBk === EST.BACKORDER) { hayBackorderPostDesp = true; }
        // Ítems Preparado/Parcial no despachados con prevDesp=0 serán reseteados a Backorder
        if (oldDespBk === 0 && (estBk === EST.PREPARADO || estBk === EST.PREP_PARCIAL)) { hayBackorderPostDesp = true; }
      }
    }
    var hayBackorder = hayBackorderPostDesp; // alias para Logger y plain text

    var etiquetaEnvio = hayDespPrevio ? '2\xba env\xedo' : (hayBackorderPostDesp ? '1\xba env\xedo' : 'Despacho');

    var introReseller;
    if (hayDespPrevio) {
      introReseller =
        "<p style='font-size:14px;color:#444;margin:0 0 10px'>" +
          "2\xba env\xedo de tu pedido <strong style='color:#00a3e0'>" + numero + "</strong>: " +
          "despachamos el stock que estaba pendiente. Adjuntamos la Nota de Entrega correspondiente." +
        "</p>" +
        (hayBackorderPostDesp
          ? "<div style='background:#fff8e1;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px;margin-bottom:14px'>" +
              "<p style='margin:0;font-size:12px;color:#7a5800'>A\xfan quedan \xedtems en <strong>backorder</strong>. Los despacharemos en cuanto llegue el stock.</p>" +
            "</div>"
          : "<div style='background:#e8f5e9;border-left:4px solid #43a047;padding:10px 14px;border-radius:4px;margin-bottom:14px'>" +
              "<p style='margin:0;font-size:12px;color:#1b5e20'>Con este env\xedo tu pedido queda <strong>completo</strong>.</p>" +
            "</div>");
    } else if (hayBackorderPostDesp) {
      introReseller =
        "<p style='font-size:14px;color:#444;margin:0 0 10px'>" +
          "1\xba env\xedo de tu pedido <strong style='color:#00a3e0'>" + numero + "</strong>: " +
          "despachamos los \xedtems disponibles. Adjuntamos la Nota de Entrega correspondiente." +
        "</p>" +
        "<div style='background:#fff8e1;border-left:4px solid #f39c12;padding:10px 14px;border-radius:4px;margin-bottom:14px'>" +
          "<p style='margin:0;font-size:12px;color:#7a5800'>" +
            "Los \xedtems en <strong>backorder</strong> ser\xe1n despachados en un 2\xba env\xedo " +
            "cuando el stock est\xe9 disponible." +
          "</p>" +
        "</div>";
    } else {
      introReseller =
        "<p style='font-size:14px;color:#444;margin:0 0 16px'>" +
          "Tu pedido <strong style='color:#00a3e0'>" + numero + "</strong> fue <strong>despachado</strong>. " +
          "Adjuntamos la Nota de Entrega correspondiente." +
        "</p>";
    }

    var cuerpoReseller = introReseller +
      tablaHtml + totalHtml + chipsHtml + rastrearBtn + driveBtn + obsHtml + confirmEntregaHtml;

    // ── Sección para administración (incluida al pie del mismo email) ─
    // Forma de pago: OT con DJI aprobado → sin costo; OT sin aprobar → 30 días; PR → campo PAGO
    var _esDJIAprobado = _esNumeroOT(numero) && ped.obs.indexOf('DJI ✓ Aprobado') !== -1;
    var formaPago = _esNumeroOT(numero)
      ? (_esDJIAprobado ? 'Sin costo — DJI aprobado' : 'A 30 días')
      : (ped.pago || 'N/E');
    var tituloAdmin = hayDespPrevio
      ? "Facturar 2\xba env\xedo — Pedido " + numero + (hayBackorderPostDesp ? " <span style='font-weight:400;color:#b8860b'>(a\xfan hay pendientes)</span>" : "")
      : (hayBackorderPostDesp
        ? "Facturar 1\xba env\xedo — Pedido " + numero + " <span style='font-weight:400;color:#b8860b'>(despacho parcial)</span>"
        : "Generar factura / ID de venta — Pedido " + numero);
    var adminHtml =
      "<div style='margin-top:28px;border-top:2px dashed #e0e3e8;padding-top:20px'>" +
        "<p style='margin:0 0 12px;font-size:10px;font-weight:700;text-transform:uppercase;" +
           "letter-spacing:.1em;color:#9ba5b4'>Para administración</p>" +
        "<div style='background:#fffbe6;border:1px solid #f9d95a;border-radius:8px;padding:14px 18px'>" +
          "<p style='margin:0 0 10px;font-size:13px;font-weight:700;color:#7a5800'>" +
            tituloAdmin +
          "</p>" +
          "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
            "<tr><td style='padding:3px 0;color:#888;width:130px'>N\xb0 de env\xedo</td>" +
                "<td style='font-weight:700;color:#00a3e0;font-size:13px'>" + Number(notaNumStr) + "\xaa factura \xb7 <span style='font-family:monospace;font-size:11px;color:#555'>" + notaEntrega + "</span></td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Reseller</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>" + ped.reseller + "</td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Fecha despacho</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>" + fecha + "</td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Forma de pago</td>" +
                "<td style='font-weight:600;color:" + (_esDJIAprobado ? '#1a9e4a' : '#1a1f2e') + "'>" + formaPago + "</td></tr>" +
            "<tr><td style='padding:3px 0;color:#888'>Transportista</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>" + (transp || 'N/E') + "</td></tr>" +
            (costo > 0 ?
            "<tr><td style='padding:3px 0;color:#888'>Costo de env\xedo</td>" +
                "<td style='font-weight:600;color:#1a1f2e'>$ " + _formatMoneda(costo) + "</td></tr>" : '') +
          "</table>" +
          "<div style='background:#fff;border:1px solid #f0e68c;border-radius:6px;padding:10px 14px'>" +
            "<p style='margin:0;font-size:13px;font-weight:800;color:#1a1f2e'>" +
              "Total: <span style='color:#00a3e0'>USD " + _formatMoneda(totalUSD) + "</span>" +
              "<span style='font-size:10px;font-weight:400;color:#9ba5b4;margin-left:8px'>no incluye impuestos</span>" +
            "</p>" +
          "</div>" +
          (sheetUrl || pdfUrl ?
          "<p style='margin:12px 0 0;font-size:12px;color:#5e6778'>Nota de Entrega:&nbsp; " +
            (pdfUrl   ? "<a href='" + pdfUrl   + "' target='_blank' style='color:#00a3e0;font-weight:700;text-decoration:none'> PDF</a>" : '') +
            (sheetUrl && pdfUrl ? "&nbsp;&nbsp;\xb7&nbsp;&nbsp;" : '') +
            (sheetUrl ? "<a href='" + sheetUrl + "' target='_blank' style='color:#188038;font-weight:700;text-decoration:none'>Planilla (Google Sheets)</a>" : '') +
          "</p>" : '') +
        "</div>" +
      "</div>";

    var tituloEmail  = etiquetaEnvio + ' — ' + numero;
    var htmlCombinado = _wosPortalHead(tituloEmail) +
      "<p style='font-size:14px;color:#666;margin:0 0 22px;line-height:1.5'>Hola <strong>" + ped.reseller + "</strong>:</p>" +
      cuerpoReseller +
      adminHtml +
      _wosPortalFoot('Pedido ' + numero + ' · ' + ped.reseller + '.');

    var plainIntro = hayDespPrevio
      ? '2\xba env\xedo de tu pedido ' + numero + ': despachamos el stock pendiente.' + (hayBackorderPostDesp ? '\nA\xfan quedan \xedtems en backorder que despacharemos cuando llegue el stock.' : '\nCon este env\xedo tu pedido queda completo.')
      : (hayBackorderPostDesp
        ? '1\xba env\xedo de tu pedido ' + numero + ': despachamos los \xedtems disponibles.\nLos \xedtems en backorder llegar\xe1n en un 2\xba env\xedo cuando el stock est\xe9 disponible.'
        : 'Tu pedido ' + numero + ' fue despachado.');

    var plainAdminTitulo = hayDespPrevio
      ? 'Facturar 2\xba env\xedo' + (hayBackorderPostDesp ? ' (a\xfan hay pendientes)' : '') + ' — Pedido '
      : (hayBackorderPostDesp ? 'Facturar 1\xba env\xedo (despacho parcial) — Pedido ' : 'Generar factura / ID de venta — Pedido ');
    var plainCombinado =
      'Hola ' + ped.reseller + ',\n\n' +
      plainIntro +
      (transp ? '\nTransportista: ' + transp : '') +
      (track  ? '\nC\xf3digo de seguimiento: ' + track : '') + '\n\n' +
      'Adjuntamos la Nota de Entrega.' +
      (pdfUrl ? '\nVer en Drive: ' + pdfUrl : '') +
      '\nCuando recib\xe1s el paquete, respond\xe9 este correo con "Recibido" para confirmar la entrega.\n\n' +
      '---\nPARA ADMINISTRACIÓN\n' +
      plainAdminTitulo + numero + '\n' +
      'N\xb0 de env\xedo: ' + Number(notaNumStr) + '\xaa factura (' + notaEntrega + ')\n' +
      'Reseller: ' + ped.reseller + '\n' +
      'Fecha: ' + fecha + '\n' +
      'Forma de pago: ' + formaPago + '\n' +
      (costo > 0 ? 'Costo de env\xedo: $ ' + _formatMoneda(costo) + '\n' : '') +
      'Total: USD ' + _formatMoneda(totalUSD) +
      (pdfUrl   ? '\nNota de Entrega (PDF): ' + pdfUrl : '') +
      (sheetUrl ? '\nNota de Entrega (planilla): ' + sheetUrl : '');

    // ── Un solo reply al hilo original, facturación en CC ─────
    var replyOpts = {
      htmlBody:    htmlCombinado,
      name:        'BIDCOMAGRO · Portal Resellers',
      replyTo:     EMAIL_SOPORTE,
      cc:          EMAIL_FACTURACION
    };
    if (pdfBlob) replyOpts.attachments = [pdfBlob];
    try {
      // Responder al hilo apuntando a los destinatarios ORIGINALES (no a quien haya quedado en el
      // último mensaje si hubo una conversación aparte), garantizando reseller + facturación en CC.
      var _repOk = _wosReplyHiloOriginal(ped.threadId, plainCombinado, replyOpts, [email, EMAIL_FACTURACION]);
      if (!_repOk) {
        // Sin email de cliente (externo) → al menos facturación recibe la sección administrativa.
        var _destFallback = email || EMAIL_FACTURACION;
        Logger.log('WOS_despacharCompleto: hilo no disponible (' + ped.threadId + '), enviando email nuevo → ' + _destFallback);
        GmailApp.sendEmail(_destFallback, tituloEmail + ' — Pedido ' + numero, plainCombinado, replyOpts);
      }
    } catch(eThread) {
      var _destFallback = email || EMAIL_FACTURACION;
      Logger.log('WOS_despacharCompleto reply hilo error (' + ped.threadId + '): ' + eThread + ' → email nuevo a ' + _destFallback);
      try { GmailApp.sendEmail(_destFallback, tituloEmail + ' — Pedido ' + numero, plainCombinado, replyOpts); } catch(eSE) { Logger.log('WOS_despacharCompleto fallback sendEmail: ' + eSE); }
    }

    // ── Estado final por ítem ────────────────────────────────────
    // Si cantFinal >= cantSol → Entregado_Cerrado (despacho completo del ítem)
    // Si cantFinal <  cantSol → Backorder (quedan unidades pendientes para 2º envío)
    // Items sin dispNow → se dejan intactos
    var esRetiro   = String(ped.envio || '').toLowerCase().indexOf('retiro') >= 0;
    var estadoDesp = esRetiro ? EST.LISTO_RETIRO : EST.ENTREGADO;
    var ahora2     = new Date();
    for (var sf = 1; sf < ped.datos.length; sf++) {
      if (String(ped.datos[sf][COL.NUMERO] || '').trim() !== numero) continue;
      var dispNowSf  = (despMap[sf + 1] !== undefined) ? (Number(despMap[sf + 1]) || 0) : 0;
      if (dispNowSf > 0) {
        var cantSolSf    = Number(ped.datos[sf][COL.CANT_SOL])    || 0;
        var oldDespSf    = Number(ped.datos[sf][COL.CANT_DESP])   || 0;
        var cantCancelSf = Number(ped.datos[sf][COL.CANT_CANCEL]) || 0;
        var cantFinalSf  = oldDespSf + dispNowSf;
        var estPrevSf    = String(ped.datos[sf][COL.ESTADO] || '').trim();
        var rSf = ped.hoja.getRange(sf + 1, COL.ESTADO + 1);
        rSf.clearDataValidations();
        // La línea se da por CERRADA cuando se despachó todo lo que NO fue cancelado (cantSol − Z),
        // no cuando se llega a cantSol: si el reseller canceló el faltante (Opción B), despachar lo
        // disponible completa el ítem → Entregado (antes exigía llegar a cantSol y lo devolvía a
        // Backorder pese a estar cancelado el resto).
        // Si aún queda pendiente y la línea estaba En_Espera_Reseller (faltante esperando la decisión
        // del reseller), se MANTIENE En_Espera aunque despachemos lo disponible ahora — así la
        // respuesta A/B (que filtra por En_Espera) todavía la resuelve.
        var _estFinalSf = (cantFinalSf >= (cantSolSf - cantCancelSf)) ? estadoDesp
                        : (estPrevSf === EST.EN_ESPERA ? EST.EN_ESPERA : EST.BACKORDER);
        rSf.setValue(_estFinalSf);
        ped.hoja.getRange(sf + 1, COL.FECHA_ESTADO + 1).setValue(ahora2);
      } else {
        // Ítem no despachado: si estaba Preparado/Parcial y nunca tuvo despacho previo → volver a Backorder
        var estActualSf = String(ped.datos[sf][COL.ESTADO] || '').trim();
        var prevDespSf  = Number(ped.datos[sf][COL.CANT_DESP]) || 0;
        if (prevDespSf === 0 && (estActualSf === EST.PREPARADO || estActualSf === EST.PREP_PARCIAL)) {
          var rSfRst = ped.hoja.getRange(sf + 1, COL.ESTADO + 1);
          rSfRst.clearDataValidations();
          rSfRst.setValue(EST.BACKORDER);
          ped.hoja.getRange(sf + 1, COL.FECHA_ESTADO + 1).setValue(ahora2);
        }
        // Ítems ya despachados parcialmente (prevDesp>0) o en Backorder/Cancelado → intactos
      }
    }
    SpreadsheetApp.flush();

    // Actualizar HUB PRO OT: E: en REPUESTOS (col Q). WOS NO toca el ESTADO de la OT (col E) —
    // eso lo decide un humano en HUB_PRO, con el stepper de la orden. Si el despacho quedó
    // completo (sin backorder), se deja una marca de texto en col AC (fecha + nota + qué se
    // envió) avisando que se puede pasar el estado a "Repuestos enviados" — ver
    // HUB_PRO/Env.js SCHEMA.OT.REPUESTOS_ENVIO_WOS y el aviso que pinta Index.html.
    if (_esNumeroOT(numero)) {
      try {
        var otNum = numero.replace(/^OT-/, '');
        // SKU → total despachado (previo + este despacho), y SKU → cantidad de ESTE despacho
        var skuDespMap = {}, skuEnviadoAhora = {};
        for (var di = 1; di < ped.datos.length; di++) {
          if (String(ped.datos[di][COL.NUMERO] || '').trim() !== numero) continue;
          var dSku = String(ped.datos[di][COL.SKU] || '').trim().toUpperCase();
          if (!dSku) continue;
          var enviadoAhoraDi = (despMap[di + 1] !== undefined) ? (Number(despMap[di + 1]) || 0) : 0;
          var totalDi = (Number(ped.datos[di][COL.CANT_DESP]) || 0) + enviadoAhoraDi;
          skuDespMap[dSku] = (skuDespMap[dSku] || 0) + totalDi;
          if (enviadoAhoraDi > 0) skuEnviadoAhora[dSku] = (skuEnviadoAhora[dSku] || 0) + enviadoAhoraDi;
        }
        var hubHoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Ordenes de trabajo');
        if (hubHoja) {
          var hubD = hubHoja.getDataRange().getValues();
          for (var hi = 1; hi < hubD.length; hi++) {
            if (String(hubD[hi][2] || '').trim() !== otNum) continue;
            // Actualizar E: en campo REPUESTOS (col Q = índice 16, rango = col 17)
            var repStr = String(hubD[hi][16] || '').trim();
            if (repStr) {
              var repPartes = repStr.split(' ; ');
              var repNuevas = [];
              for (var rpi = 0; rpi < repPartes.length; rpi++) {
                var rpCols = repPartes[rpi].split(' | ');
                var rpSku  = String(rpCols[0] || '').trim().toUpperCase();
                if (rpSku && skuDespMap[rpSku] !== undefined) {
                  var pedVal = String(rpCols[2] || '').split(' E:')[0].replace('P:', '') || '0';
                  rpCols[2]  = 'P:' + pedVal + ' E:' + skuDespMap[rpSku];
                }
                repNuevas.push(rpCols.join(' | '));
              }
              hubHoja.getRange(hi + 1, 17).setValue(repNuevas.join(' ; '));
            }
            // Marca de "envío completo" en col AC (29) — solo aviso, no toca el ESTADO.
            if (!hayBackorder) {
              var _maxColHub = hubHoja.getMaxColumns();
              if (_maxColHub < 29) hubHoja.insertColumnsAfter(_maxColHub, 29 - _maxColHub);
              if (!String(hubHoja.getRange(1, 29).getValue() || '').trim()) {
                hubHoja.getRange(1, 29).setValue('Envío Repuestos (WOS)');
              }
              var _itemsTxt = [];
              for (var _esk in skuEnviadoAhora) _itemsTxt.push(_esk + ' x' + skuEnviadoAhora[_esk]);
              hubHoja.getRange(hi + 1, 29).setValue(
                Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') +
                ' — Nota ' + notaEntrega + (_itemsTxt.length ? ' (' + _itemsTxt.join(', ') + ')' : '')
              );
            }
            break;
          }
        }
      } catch(eHub) { Logger.log('WOS_despacharCompleto HUB update: ' + eHub); }
    }

    var logExtra = (hayBackorder ? 'parcial+backorder' : 'completo') + ' transp=' + transp + ' track=' + track;
    _wosLogAccion('Despacho: ' + notaEntrega + ' → ' + estadoDesp, numero, ped.reseller, operario, logExtra);
    Logger.log('WOS_despacharCompleto OK: ' + numero + '-' + notaNumStr + ' | ' + estadoDesp + (hayBackorder ? '+backorder' : '') + ' | transp=' + transp + ' | track=' + track + ' | driveUrl=' + pdfUrl);
    var resultado = { ok: true, nota: notaEntrega };
    _wosIdempotGuardar(reqToken, resultado);   // marca el token como procesado (retornos repetidos no re-ejecutan)
    return resultado;
  } catch(e) {
    Logger.log('WOS_despacharCompleto ERROR: ' + e);
    return { ok: false, error: e.toString() };
  } finally {
    try { _lock.releaseLock(); } catch(eR) {}
  }
}



// ─────────────────────────────────────────────────────────────
//  HITO 2 — Notificar Faltante
//  Responde en el hilo ancla informando el faltante y pidiendo
//  Opción A (esperar) o Opción B (despachar lo disponible).
//
//  faltantes: [{sku, desc, cantSol, cantDisp}]
// ─────────────────────────────────────────────────────────────
function WOS_notificarFaltante(numero, faltantes, operario, reqToken) {
 return _wosLockIdempot(reqToken, function() {
  try {
    operario = String(operario || '');
    // Pedidos de OT (reparación, HUB_PRO): la pregunta de faltante es distinta — acá no
    // tiene sentido despachar una pieza suelta si falta la principal para terminar el
    // arreglo. Ver EST.RESERVADO_CONSOLIDAR y WOS_procesarRespuestaManual.
    var esOT = _esNumeroOT(numero);
    var ped = _wosLeerPedido(numero);
    if (!ped.reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    // Sin threadId: el bloque try/catch de más abajo ya envía email nuevo como fallback.

    // Cliente externo (super-RTV): puede no estar en Resellers. El aviso va por replyAll al hilo
    // (threadId ya garantizado arriba); el email directo es solo fallback (abajo). No se aborta.
    var email = _wosGetEmailReseller(ped.reseller);

    // ── ETA de reposición por SKU (compras DJI en camino) ──────
    // Para que el reseller sepa CUÁNDO llegaría el faltante si elige "Esperar".
    var _ecMap = {};
    try {
      var _ec = WOS_getEnCaminoMap();
      if (_ec && _ec.ok) _ecMap = _ec.map || {};
    } catch(eEc) { Logger.log('WOS_notificarFaltante enCamino: ' + eEc); }
    // Reposición por SKU: reparte el faltante entre los lotes DJI DISPONIBLES (descuenta lo
    // reservado a otros) → cuántas unidades llegan con cada ETA. { lineas:[{qty,eta}], etaProx, sinFecha }
    function _faltReposicion(sku, faltaQty) {
      var e = _ecMap[String(sku || '').trim().toUpperCase()];
      if (!e || !e.batchesDisp || !e.batchesDisp.length || faltaQty <= 0) {
        return { lineas: [], etaProx: '', sinFecha: faltaQty > 0 ? faltaQty : 0 };
      }
      var rem = faltaQty, lineas = [], etaProx = '';
      for (var b = 0; b < e.batchesDisp.length && rem > 0; b++) {
        var bt = e.batchesDisp[b];
        var take = Math.min(rem, bt.qty);
        if (take <= 0) continue;
        lineas.push({ qty: take, eta: bt.eta });
        if (!etaProx && bt.eta) etaProx = bt.eta;
        rem -= take;
      }
      return { lineas: lineas, etaProx: etaProx, sinFecha: rem };
    }
    var _repMap = {};
    for (var _rf = 0; _rf < faltantes.length; _rf++) {
      var _rfSku   = String(faltantes[_rf].sku || '').trim().toUpperCase();
      var _rfFalta = Math.max(0, (Number(faltantes[_rf].cantSol) || 0) - (Number(faltantes[_rf].cantDisp) || 0));
      _repMap[_rfSku] = _faltReposicion(_rfSku, _rfFalta);
    }
    // ETA de reposición para el bloque OPCIÓN A: antes se tomaba una sola fecha (la más
    // próxima entre TODOS los faltantes), lo cual era engañoso apenas había más de un SKU
    // faltante con lotes de ETAs distintas — mostraba "~04/08" como si TODO llegara ese día
    // cuando en realidad una parte llegaba 16/08 y otra 24/08 (bug reportado). Ahora se arma
    // un desglose por fecha (mismo criterio que _wosMsgReservas para el mail de confirmación).
    var _etaMap = {};
    for (var _pk in _repMap) {
      var _lineas = _repMap[_pk].lineas || [];
      for (var _li = 0; _li < _lineas.length; _li++) {
        var _le = _lineas[_li].eta || '';
        if (!_le) continue;
        _etaMap[_le] = (_etaMap[_le] || 0) + (_lineas[_li].qty || 0);
      }
    }
    var _etaDesglose = Object.keys(_etaMap).map(function(k) {
      return { eta: k, cantidad: _etaMap[k], _dt: _wosEtaToDate(k) };
    }).sort(function(a, b) {
      var ta = a._dt ? a._dt.getTime() : Infinity;
      var tb = b._dt ? b._dt.getTime() : Infinity;
      return ta - tb;
    });
    var _etaHtml = '';
    if (_etaDesglose.length === 1) {
      _etaHtml = "<p style='margin:8px 0 0;font-size:12px;color:#3730a3;font-weight:600'>&#128666; Según las compras en curso, la reposición llegaría aprox. el <strong>~" + _etaDesglose[0].eta + "</strong>.</p>";
    } else if (_etaDesglose.length > 1) {
      _etaHtml = "<p style='margin:8px 0 0;font-size:12px;color:#3730a3;font-weight:600'>&#128666; Según las compras en curso:</p>" +
        "<ul style='margin:2px 0 0;padding-left:18px;font-size:12px;color:#3730a3'>" +
        _etaDesglose.map(function(d) { return "<li>" + d.cantidad + " u. llegarían aprox. el ~" + d.eta + "</li>"; }).join('') +
        "</ul>";
    }
    var _etaPlain = '';
    if (_etaDesglose.length === 1) {
      _etaPlain = ' (reposición estimada ~' + _etaDesglose[0].eta + ')';
    } else if (_etaDesglose.length > 1) {
      _etaPlain = ' (reposición estimada: ' + _etaDesglose.map(function(d) { return d.cantidad + 'u ~' + d.eta; }).join(', ') + ')';
    }

    // ── Tabla de faltantes ────────────────────────────────────
    var tablaFalt =
      "<table style='width:100%;border-collapse:collapse;font-size:12px;margin:16px 0'>" +
      "<thead><tr style='background:#fdecea'>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>SKU</th>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Solicitado</th>" +
        "<th style='padding:7px 10px;text-align:center;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Disponible</th>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Reposici\xf3n estim.</th>" +
      "</tr></thead><tbody>";
    for (var i = 0; i < faltantes.length; i++) {
      var f = faltantes[i];
      var _rep = _repMap[String(f.sku || '').trim().toUpperCase()] || { lineas: [], etaProx: '', sinFecha: 0 };
      var _repHtml;
      if (_rep.lineas.length) {
        _repHtml = '';
        for (var _rl = 0; _rl < _rep.lineas.length; _rl++) {
          _repHtml += "<div style='white-space:nowrap;color:#1a56db'><strong>" + _rep.lineas[_rl].qty + " u.</strong> \xb7 llega ~" + (_rep.lineas[_rl].eta || 'a confirmar') + "</div>";
        }
        if (_rep.sinFecha > 0) _repHtml += "<div style='white-space:nowrap;color:#94a3b8'>" + _rep.sinFecha + " u. \xb7 a confirmar</div>";
      } else {
        _repHtml = "<span style='color:#94a3b8'>a confirmar</span>";
      }
      tablaFalt +=
        "<tr style='border-bottom:1px solid #f0f2f5'>" +
          "<td style='padding:8px 10px;font-family:monospace;font-weight:700;color:#e74c3c;vertical-align:top'>" + (f.sku  || '') + "</td>" +
          "<td style='padding:8px 10px;color:#1a202c;vertical-align:top'>"                                      + (f.desc || '') + "</td>" +
          "<td style='padding:8px 10px;text-align:center;font-weight:700;vertical-align:top'>"                  + (f.cantSol  || 0) + "</td>" +
          "<td style='padding:8px 10px;text-align:center;font-weight:700;vertical-align:top;color:" +
            (f.cantDisp > 0 ? '#1a9e4a' : '#e74c3c') + "'>"                                 + (f.cantDisp || 0) + "</td>" +
          "<td style='padding:8px 10px;text-align:left;font-size:12px;vertical-align:top'>"                     + _repHtml + "</td>" +
        "</tr>";
    }
    tablaFalt += "</tbody></table>";

    // ¿Hay algo disponible para despachar?
    var hayDisponible = false;
    for (var fd = 0; fd < faltantes.length; fd++) {
      if ((faltantes[fd].cantDisp || 0) > 0) { hayDisponible = true; break; }
    }

    // ── Cuerpo HTML ───────────────────────────────────────────
    // Pedidos de OT: TODAVÍA no se despachó nada (a diferencia de reseller, donde lo
    // disponible ya se preparó de inmediato) — el copy tiene que reflejar eso.
    var introDisp = esOT
      ? (hayDisponible
          ? "<p style='font-size:13px;color:#1a1f2e;background:#eef2ff;border:1px solid #c7d2fe;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
              "&#128230; <strong>Ya tenemos parte de este pedido disponible.</strong><br>" +
              "<span style='font-size:12px;color:#4a5568'>Todav\xeda no despachamos nada — necesitamos que nos digas c\xf3mo prefer\xeds recibirlo:</span>" +
            "</p>"
          : "<p style='font-size:13px;color:#7c3c00;background:#fff3e0;border:1px solid #ffba7b;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
              "&#9888; <strong>No tenemos stock disponible</strong> para ninguno de los \xedtems de este pedido.<br>" +
              "<span style='font-size:12px;color:#4a5568'>Por favor indicanos c\xf3mo quer\xe9s proceder:</span>" +
            "</p>")
      : (hayDisponible
          ? "<p style='font-size:13px;color:#155724;background:#e8f7ee;border:1px solid #8fd4a8;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
              "&#128666; <strong>Vamos a despachar lo disponible</strong> a la brevedad.<br>" +
              "<span style='font-size:12px;color:#4a5568'>Solo necesitamos saber qué preferís hacer con los ítems faltantes:</span>" +
            "</p>"
          : "<p style='font-size:13px;color:#7c3c00;background:#fff3e0;border:1px solid #ffba7b;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
              "&#9888; <strong>No tenemos stock disponible</strong> para ninguno de los ítems faltantes.<br>" +
              "<span style='font-size:12px;color:#4a5568'>Por favor indicanos cómo querés proceder:</span>" +
            "</p>");

    var _wosUrl = '';
    try { _wosUrl = ScriptApp.getService().getUrl(); } catch(eUrl) { Logger.log('WOS URL error: ' + eUrl); }
    var _itemsStr = '';
    for (var fi = 0; fi < faltantes.length; fi++) {
      if (_itemsStr) _itemsStr += ',';
      _itemsStr += String(faltantes[fi].sku || '').toUpperCase().replace(/[,: ]/g, '') + ':' + (faltantes[fi].cantDisp || 0);
    }
    // items= va en AMBAS URLs (antes solo en B): WOS_procesarRespuestaManual lo necesita
    // en pedidos OT para distinguir, entre los ítems retenidos, cuáles eran el faltante
    // real (Opción A · consolidar) de los que solo estaban disponibles.
    var _urlA = _wosUrl ? _wosUrl + '?page=resp_faltante&num=' + encodeURIComponent(numero) + '&op=A&items=' + encodeURIComponent(_itemsStr) : '';
    var _urlB = _wosUrl ? _wosUrl + '?page=resp_faltante&num=' + encodeURIComponent(numero) + '&op=B&items=' + encodeURIComponent(_itemsStr) : '';

    // Rótulos de los botones/tarjetas: mismas letras A/B (el parser de Hito 3 depende del
    // texto literal "Opción A"/"Opción B"), pero para OT el SIGNIFICADO es otro — ver
    // WOS_procesarRespuestaManual: acá A = consolidar (retener todo), B = despachar ahora.
    var _lblBotonA = esOT ? 'Opci\xf3n A — Esperar y consolidar' : 'Opci\xf3n A — Esperar el faltante';
    var _lblBotonB = esOT ? 'Opci\xf3n B — Despachar ahora'      : 'Opci\xf3n B — Cancelar el faltante';

    var _botonesHtml = _urlA
      ? "<div style='text-align:center;margin:22px 0 8px'>" +
          "<a href='" + _urlA + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#3730a3;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>" + _lblBotonA + "</a>" +
          "<a href='" + _urlB + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#92400e;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>" + _lblBotonB + "</a>" +
        "</div>" +
        "<p style='font-size:11px;color:#888;text-align:center;margin:4px 0 0'>Si los botones no funcionan, respond\xe9 este correo con <strong>\"Opci\xf3n A\"</strong> o <strong>\"Opci\xf3n B\"</strong>.</p>"
      : "<p style='font-size:12px;color:#555;background:#f5f8fc;border-radius:6px;padding:10px 14px;line-height:1.6'>" +
          "Respond\xe9 este correo indicando tu elecci\xf3n: <strong>\"Opci\xf3n A\"</strong> o <strong>\"Opci\xf3n B\"</strong>." +
        "</p>";

    var opcionesHtml = introDisp +
      "<p style='font-size:13px;color:#1a1f2e;font-weight:700;margin:18px 0 10px'>\xbfQu\xe9 hacemos con " + (esOT ? 'este pedido' : 'el faltante') + "?</p>" +
      (esOT
        ? "<div style='border:1px solid #c7d2fe;border-radius:8px;padding:14px 18px;margin-bottom:8px;background:#eef2ff'>" +
            "<p style='margin:0 0 4px;font-size:13px;color:#3730a3;font-weight:700'>OPCI\xd3N A — Esperar y consolidar en un solo env\xedo</p>" +
            "<p style='margin:0;font-size:12px;color:#4a5568'>Retenemos tambi\xe9n lo que ya est\xe1 disponible. Cuando llegue el resto, se despacha/entrega todo junto — un solo env\xedo.</p>" +
            _etaHtml +
          "</div>" +
          "<div style='border:1px solid #ffba7b;border-radius:8px;padding:14px 18px;margin-bottom:16px;background:#fff3e0'>" +
            "<p style='margin:0 0 4px;font-size:13px;color:#7c3c00;font-weight:700'>OPCI\xd3N B — Despachar ahora lo disponible</p>" +
            "<p style='margin:0;font-size:12px;color:#4a5568'>Te enviamos ya lo que tenemos; el resto llega despu\xe9s, en un segundo env\xedo. No se cancela nada.</p>" +
          "</div>"
        : "<div style='border:1px solid #c7d2fe;border-radius:8px;padding:14px 18px;margin-bottom:8px;background:#eef2ff'>" +
            "<p style='margin:0 0 4px;font-size:13px;color:#3730a3;font-weight:700'>OPCI\xd3N A — Esperar el faltante (segundo env\xedo)</p>" +
            "<p style='margin:0;font-size:12px;color:#4a5568'>Los \xedtems faltantes quedan pendientes. Cuando ingresen al stock los despachamos en un segundo env\xedo.</p>" +
            _etaHtml +
          "</div>" +
          "<div style='border:1px solid #ffba7b;border-radius:8px;padding:14px 18px;margin-bottom:16px;background:#fff3e0'>" +
            "<p style='margin:0 0 4px;font-size:13px;color:#7c3c00;font-weight:700'>OPCI\xd3N B — Cancelar el faltante</p>" +
            "<p style='margin:0;font-size:12px;color:#4a5568'>Cancelamos definitivamente los \xedtems que no tenemos en stock. Solo se factura lo que se despacha ahora.</p>" +
          "</div>") +
      _botonesHtml;

    var htmlBody = _wosPortalHead('Faltante de stock — ' + numero) +
      "<p style='font-size:14px;color:#666;margin:0 0 22px;line-height:1.5'>Hola <strong>" + ped.reseller + "</strong>:</p>" +
      "<p style='font-size:13px;color:#555;margin:0 0 6px'>" +
        "Al procesar tu pedido <strong style='color:#00a3e0'>" + numero + "</strong> " +
        "detectamos un <strong>faltante de stock</strong>:" +
      "</p>" +
      tablaFalt +
      opcionesHtml +
      (ped.obs ? "<p style='font-size:11px;color:#888;margin-top:12px'><strong>Obs.:</strong> " + ped.obs + "</p>" : '') +
      _wosPortalFoot('Pedido ' + numero + ' · ' + ped.reseller + '.');

    // ── Plain text (incluye WOSDATA para que Hito 3 / WOS_detectarRespuestasOT lo parseen) ──
    var wosDataJson = JSON.stringify({ numero: numero, faltantes: faltantes });
    var plainBody = esOT
      ? ('Hola ' + ped.reseller + ',\n\n' +
        'Al procesar el pedido ' + numero + ' detectamos un faltante de stock.\n' +
        (hayDisponible ? 'Todav\xeda no despachamos nada de este pedido. ' : '') +
        'Por favor indicanos c\xf3mo prefer\xeds recibirlo:\n\n' +
        'OPCI\xd3N A: Esperar y consolidar en un solo env\xedo (retenemos tambi\xe9n lo disponible).' +
        _etaPlain + '\n' +
        'OPCI\xd3N B: Despachar ahora lo disponible; el resto llega despu\xe9s (no se cancela nada).\n\n' +
        'Respond\xe9 este correo con tu elecci\xf3n (escrib\xed "Opci\xf3n A" u "Opci\xf3n B").\n\n' +
        '===WOSDATA===\n' + wosDataJson + '\n===ENDWOSDATA===')
      : ('Hola ' + ped.reseller + ',\n\n' +
        'Al procesar tu pedido ' + numero + ' detectamos un faltante de stock.\n' +
        (hayDisponible ? 'Vamos a despachar lo disponible. ' : '') +
        'Por favor indicanos qué hacer con los ítems faltantes:\n\n' +
        'OPCI\xd3N A: Esperar el faltante y recibirlo en un segundo env\xedo cuando el stock est\xe9 disponible.' +
        _etaPlain + '\n' +
        'OPCI\xd3N B: Cancelar definitivamente los \xedtems faltantes.\n\n' +
        'Respond\xe9 este correo con tu elecci\xf3n (escrib\xed "Opci\xf3n A" u "Opci\xf3n B").\n\n' +
        '===WOSDATA===\n' + wosDataJson + '\n===ENDWOSDATA===');

    // ── Enviar como RESPUESTA en el hilo ancla (fallback: email nuevo) ──
    var faltOpts = {
      htmlBody: htmlBody,
      name:     'BIDCOMAGRO · Portal Resellers',
      replyTo:  EMAIL_SOPORTE
    };
    try {
      // Responder al hilo apuntando a los destinatarios ORIGINALES (no a una conversación aparte),
      // garantizando que el reseller reciba la pregunta Opción A/B.
      var _faltOk = _wosReplyHiloOriginal(ped.threadId, plainBody, faltOpts, [email]);
      if (!_faltOk) {
        // Sin email (cliente externo) y sin hilo → no hay a quién preguntar Opción A/B; se omite el envío.
        if (email) {
          Logger.log('WOS_notificarFaltante: hilo no disponible (' + ped.threadId + '), enviando email nuevo → ' + email);
          GmailApp.sendEmail(email, 'Faltante de stock — Pedido ' + numero, plainBody, faltOpts);
        } else {
          Logger.log('WOS_notificarFaltante: hilo no disponible y sin email (cliente externo) → se omite el aviso; los estados igual cambian');
        }
      }
    } catch(eThread) {
      if (email) {
        Logger.log('WOS_notificarFaltante reply hilo error (' + ped.threadId + '): ' + eThread + ' → email nuevo → ' + email);
        try { GmailApp.sendEmail(email, 'Faltante de stock — Pedido ' + numero, plainBody, faltOpts); } catch(eSE) { Logger.log('WOS_notificarFaltante fallback: ' + eSE); }
      } else {
        Logger.log('WOS_notificarFaltante: sin hilo y sin email (cliente externo) → se omite el aviso');
      }
    }

    // ── Cambiar estado por ítem ──
    // Reseller: faltantes → En_Espera_Reseller, disponibles → Preparado (se despachan ya).
    // OT: faltantes → En_Espera_Reseller, disponibles TAMBIÉN → En_Espera_Reseller — nada
    // se despacha todavía. Recién en WOS_procesarRespuestaManual, según la respuesta, los
    // disponibles pasan a Preparado (Opción B) o a Reservado_Consolidar (Opción A).
    var faltSet = {};
    for (var ff = 0; ff < faltantes.length; ff++) {
      faltSet[String(faltantes[ff].sku || '').trim().toUpperCase()] = true;
    }
    _wosSetEstadoPorSku(ped.hoja, ped.datos, numero, faltSet, EST.EN_ESPERA, esOT ? EST.EN_ESPERA : EST.PREPARADO);
    SpreadsheetApp.flush();

    _wosLogAccion('Faltante notificado', numero, ped.reseller, operario, faltantes.length + ' items faltantes');
    Logger.log('WOS_notificarFaltante OK: ' + numero + ' → ' + EST.EN_ESPERA + ' | email → ' + email);
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_notificarFaltante ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
 });
}



// ─────────────────────────────────────────────────────────────
//  HITO 3 — Detector automático de respuestas del reseller
//  Asociar a un trigger de tiempo (cada 10 min).
//  Lee los hilos de pedidos en estado 'En_Espera_Reseller',
//  detecta "Opción A" o "Opción B" y actúa autónomamente.
// ─────────────────────────────────────────────────────────────
function WOS_detectarRespuestasResellers() {
  try {
    var hoja  = _wosHoja();
    var datos = hoja.getDataRange().getValues();

    // Agrupar filas por pedido, solo los que están en espera con threadId
    var enEspera = {}; // numero → { threadId, reseller, rowsIdx: [] }
    for (var i = 1; i < datos.length; i++) {
      var num    = String(datos[i][COL.NUMERO]    || '').trim();
      var estado = String(datos[i][COL.ESTADO]    || '').trim();
      var tId    = String(datos[i][COL.THREAD_ID] || '').trim();
      if (!num || estado !== EST.EN_ESPERA || !tId) continue;
      if (!enEspera[num]) {
        enEspera[num] = {
          threadId: tId,
          reseller: String(datos[i][COL.RESELLER] || ''),
          rows: []
        };
      }
      enEspera[num].rows.push({
        rowIdx:  i,                                           // 0-indexed en datos[]
        rowNum:  i + 1,                                       // 1-indexed para getRange
        sku:     String(datos[i][COL.SKU]      || '').toUpperCase(),
        cantSol: Number(datos[i][COL.CANT_SOL]) || 0
      });
    }

    var numeros   = Object.keys(enEspera);
    var procesados = 0;
    if (!numeros.length) {
      Logger.log('WOS_detectarRespuestasResellers: sin pedidos en espera');
      return { ok: true, procesados: 0 };
    }

    var dominioInterno = '@bidcom.com.ar';

    for (var n = 0; n < numeros.length; n++) {
      var numero = numeros[n];
      var info   = enEspera[numero];

      try {
        var thread   = GmailApp.getThreadById(info.threadId);
        var messages = thread.getMessages();
        if (!messages || messages.length < 2) continue; // solo existe nuestro mensaje, aún no hay respuesta

        // Recorrer mensajes: identificar el nuestro (WOSDATA) y la respuesta del reseller
        var wosDataStr    = null;
        var resellerPlain = null;
        var lastExtMsg    = null;

        for (var m = 0; m < messages.length; m++) {
          var msg    = messages[m];
          var from   = msg.getFrom().toLowerCase();
          var plain  = msg.getPlainBody();

          if (from.indexOf(dominioInterno) >= 0) {
            // Mensaje enviado por nosotros: buscar WOSDATA
            var matchW = plain.match(/===WOSDATA===\s*([\s\S]*?)\s*===ENDWOSDATA===/);
            if (matchW) wosDataStr = matchW[1].trim();
          } else {
            // Mensaje del reseller (o externo)
            lastExtMsg    = msg;
            resellerPlain = plain;
          }
        }

        if (!resellerPlain || !lastExtMsg) continue; // aún no hay respuesta externa

        // Detectar opción en el texto de la respuesta.
        // El reseller ve OPCIÓN A (esperar faltante) u OPCIÓN B (cancelar faltante).
        var rLow = resellerPlain.toLowerCase();
        var esA  = /opci[oó]n\s*a\b/.test(rLow);
        var esB  = /opci[oó]n\s*b\b/.test(rLow);

        if (!esA && !esB) {
          Logger.log('WOS_detectarRespuestasResellers: ' + numero + ' → respuesta no reconocida (se esperaba Opci\xf3n A u Opci\xf3n B), se ignora');
          continue;
        }

        // Parsear WOSDATA para obtener cantidades disponibles (necesario para B y C)
        var faltantesMap = {};
        if (wosDataStr) {
          try {
            var wosObj = JSON.parse(wosDataStr);
            var fArr   = wosObj.faltantes || [];
            for (var fi = 0; fi < fArr.length; fi++) {
              var skuUp = String(fArr[fi].sku || '').trim().toUpperCase();
              if (skuUp) faltantesMap[skuUp] = Number(fArr[fi].cantDisp) || 0;
            }
          } catch(eJ) { Logger.log('WOS_detectarRespuestasResellers JSON parse: ' + eJ); }
        }

        var confTitulo, confMensaje, nuevoEst;

        if (esA) {
          // Opción A (reseller) = esperar faltante en segundo envío
          // Solo los items En_Espera_Reseller pasan a Backorder; los Preparado quedan intactos
          nuevoEst   = EST.BACKORDER;
          confTitulo = 'Recibimos tu respuesta — Pedido ' + numero;
          confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar el faltante</strong>.<br>" +
            "Recibir\xe1s la Nota de Entrega por los \xedtems que se despachan ahora. El faltante llegar\xe1 en un segundo env\xedo cuando el stock est\xe9 disponible.";
          _wosSetEstadoFiltrado(hoja, datos, numero, EST.EN_ESPERA, nuevoEst);
          Logger.log('WOS_detectarRespuestasResellers: ' + numero + ' → OPCI\xd3N A → faltantes → ' + nuevoEst);

        } else {
          // Opción B (reseller) = cancelar faltante → despacho único de lo disponible
          // Solo los items En_Espera_Reseller se modifican: cantDisp>0 → Preparado, cantDisp=0 → Cancelado
          nuevoEst   = EST.PREPARADO;
          confTitulo = 'Recibimos tu respuesta — Pedido ' + numero;
          confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Cancelar el faltante</strong>.<br>" +
            "Estamos preparando el despacho de lo disponible. Recibir\xe1s la Nota de Entrega cuando sea despachado. Los \xedtems faltantes quedan cancelados.";
          var ahoraB = new Date();
          for (var ri = 0; ri < info.rows.length; ri++) {
            var row = info.rows[ri];
            // IDÉNTICO a WOS_procesarRespuestaManual: aplica al faltante pendiente esté En_Espera_Reseller
            // O ya aparcado en Backorder (en pedidos mixtos el 2º no se cancelaba y quedaba en backorder).
            var _estFilaBe = String(datos[row.rowNum - 1][COL.ESTADO] || '').trim();
            if (_estFilaBe !== EST.EN_ESPERA && _estFilaBe !== EST.BACKORDER) continue;
            // Sobre lo REALMENTE pendiente (E−F−Z) y acumulando CANT_CANCEL → no cancela de más una
            // línea con despacho/cancelación parcial previa.
            var _solBe   = Number(datos[row.rowNum - 1][COL.CANT_SOL])    || 0;
            var _despBe  = Number(datos[row.rowNum - 1][COL.CANT_DESP])   || 0;
            var _cancBe0 = Number(datos[row.rowNum - 1][COL.CANT_CANCEL]) || 0;
            var _pendBe  = _solBe - _despBe - _cancBe0;
            if (_pendBe <= 0) continue;
            var cantDisp = (faltantesMap[row.sku] !== undefined) ? faltantesMap[row.sku] : 0;
            if (cantDisp > _pendBe) cantDisp = _pendBe;
            var rEstB = hoja.getRange(row.rowNum, COL.ESTADO + 1);
            rEstB.clearDataValidations();
            if (cantDisp >= _pendBe) {
              // Hay stock para todo lo pendiente → se despacha completo, no se cancela nada
              rEstB.setValue(EST.PREPARADO);
            } else if (cantDisp > 0) {
              // Parcial: se despacha lo disponible y se CANCELA el resto pendiente (acumula sobre Z).
              // Se PRESERVA CANT_SOL → CANT_PEND (=E−F−Z) queda en lo disponible.
              hoja.getRange(row.rowNum, COL.CANT_CANCEL + 1).setValue(_cancBe0 + (_pendBe - cantDisp));
              rEstB.setValue(EST.PREP_PARCIAL);
            } else {
              // Faltante total: se cancela todo lo pendiente (col Z) → CANT_PEND = 0.
              hoja.getRange(row.rowNum, COL.CANT_CANCEL + 1).setValue(_cancBe0 + _pendBe);
              rEstB.setValue(EST.CANCELADO);
            }
            hoja.getRange(row.rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraB);
          }
          Logger.log('WOS_detectarRespuestasResellers: ' + numero + ' → OPCI\xd3N B → faltantes ajustados por item');
        }

        // ── Reservar (A) o liberar (B) las unidades en camino comprometidas ──
        SpreadsheetApp.flush();
        if (esA) {
          try {
            var _rSumD = _wosReservarEnCamino(numero, info.reseller);
            confMensaje += _wosMsgReservas(_rSumD);
          } catch(eRs) { Logger.log('detector reservar [' + numero + ']: ' + eRs); }
        } else if (esB) {
          try { _wosCerrarReservas(numero, '', 'Cancelada'); } catch(eRl) { Logger.log('detector liberar [' + numero + ']: ' + eRl); }
        }

        // ── Confirmación al reseller en el mismo hilo ─────────
        try {
          var confHtml = _wosPortalHead(confTitulo) +
            "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + info.reseller + "</strong>:</p>" +
            "<p style='font-size:13px;color:#555;line-height:1.6'>" + confMensaje + "</p>" +
            "<p style='font-size:12px;color:#888;margin-top:18px'>Consultas: " +
              "<a href='mailto:" + EMAIL_SOPORTE + "' style='color:#00a3e0'>" + EMAIL_SOPORTE + "</a></p>" +
            _wosPortalFoot('Pedido ' + numero + ' · ' + info.reseller + '.');
          var confPlain = confMensaje.replace(/<[^>]+>/g, '');
          thread.replyAll(confPlain, {
            htmlBody: confHtml,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  EMAIL_SOPORTE
          });
        } catch(eConf) {
          Logger.log('WOS_detectarRespuestasResellers confirmacion [' + numero + ']: ' + eConf);
        }

        thread.markRead();
        procesados++;

      } catch(eT) {
        Logger.log('WOS_detectarRespuestasResellers thread [' + numero + ']: ' + eT);
      }
    }

    // Reconciliar reservas: liberar las que ya no tienen backorder pendiente (huérfanas)
    try { WOS_reconciliarReservas(); } catch(eRec) { Logger.log('detector reconciliar: ' + eRec); }

    // Avisos de ingreso: procesar la cola que llena SM al recibir un CAS con unidades bloqueadas
    try { WOS_notificarIngresos(); } catch(eNI) { Logger.log('detector notifIngresos: ' + eNI); }
    // Recordatorio/escalado de avisos de ingreso sin respuesta (3 días / 6 días)
    try { WOS_recordarIngresosPendientes(); } catch(eRI) { Logger.log('detector recordarIngresos: ' + eRI); }
    // Avisos de cambio de ETA: cola que llena SM al sincronizar un CAS contra el planner
    // externo y detectar que la fecha estimada se atrasó (solo informativo, no pide elegir)
    try { WOS_notificarCambiosEta(); } catch(eNE) { Logger.log('detector notifCambiosEta: ' + eNE); }
    // Respuestas de texto libre en Gmail para pedidos de OT (Pedidos_OTs no lo escaneaba)
    try { WOS_detectarRespuestasOT(); } catch(eOT) { Logger.log('detector respuestasOT: ' + eOT); }

    // ── Confirmación de entrega: escanear pedidos Entregado_Cerrado ──
    var entregados = {};
    for (var ei = 1; ei < datos.length; ei++) {
      var eNum = String(datos[ei][COL.NUMERO]    || '').trim();
      var eEst = String(datos[ei][COL.ESTADO]    || '').trim();
      var eTid = String(datos[ei][COL.THREAD_ID] || '').trim();
      if (!eNum || eEst !== EST.ENTREGADO || !eTid) continue;
      if (!entregados[eNum]) {
        entregados[eNum] = { threadId: eTid, reseller: String(datos[ei][COL.RESELLER] || '') };
      }
    }

    var eNumeros = Object.keys(entregados);
    for (var en = 0; en < eNumeros.length; en++) {
      var eNumero = eNumeros[en];
      var eInfo   = entregados[eNumero];
      try {
        var eThread   = GmailApp.getThreadById(eInfo.threadId);
        var eMsgs     = eThread.getMessages();
        var eLastExt  = null;
        for (var em = 0; em < eMsgs.length; em++) {
          if (eMsgs[em].getFrom().toLowerCase().indexOf(dominioInterno) < 0) {
            eLastExt = eMsgs[em];
          }
        }
        if (!eLastExt) continue;

        var eBody = eLastExt.getPlainBody().toLowerCase();
        if (!/\brecib[ií]do?\b/.test(eBody)) continue;

        // Confirmación detectada → Entregado_Confirmado
        _wosSetEstado(hoja, datos, eNumero, EST.ENTREGADO_CONF);

        var confHtml = _wosPortalHead('Entrega confirmada — Pedido ' + eNumero) +
          "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + eInfo.reseller + "</strong>:</p>" +
          "<p style='font-size:13px;color:#555;line-height:1.6'>" +
            "Registramos la confirmación de recepción de tu pedido <strong style='color:#00a3e0'>" + eNumero + "</strong>. " +
            "El pedido queda cerrado. Gracias por trabajar con BIDCOMAGRO." +
          "</p>" +
          _wosPortalFoot('Pedido ' + eNumero + ' · ' + eInfo.reseller + '.');
        var confPlain = 'Hola ' + eInfo.reseller + ',\n\nRegistramos la confirmación de recepción del pedido ' + eNumero + '. El pedido queda cerrado. Gracias!';
        eThread.replyAll(confPlain, {
          htmlBody: confHtml,
          name:     'BIDCOMAGRO · Portal Resellers',
          replyTo:  EMAIL_SOPORTE
        });
        eThread.markRead();
        procesados++;
        Logger.log('WOS_detectarRespuestasResellers: ' + eNumero + ' → confirmación entrega → ' + EST.ENTREGADO_CONF);
      } catch(eET) {
        Logger.log('WOS_detectarRespuestasResellers entrega [' + eNumero + ']: ' + eET);
      }
    }

    SpreadsheetApp.flush();
    Logger.log('WOS_detectarRespuestasResellers: procesados=' + procesados + ' de ' + (numeros.length + eNumeros.length));
    return { ok: true, procesados: procesados };

  } catch(e) {
    Logger.log('WOS_detectarRespuestasResellers ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}


// ─────────────────────────────────────────────────────────────
//  Detector de respuestas de texto libre — pedidos OT
//  Réplica del escaneo de WOS_detectarRespuestasResellers pero sobre Pedidos_OTs (que ese
//  detector nunca miraba — hoy no existía NINGÚN canal de "responder por texto" para OT).
//  A diferencia del detector de resellers, acá NO se reimplementa la lógica de estados:
//  se delega a WOS_procesarRespuestaManual (el mismo punto de convergencia que ya usan el
//  botón del mail y la respuesta telefónica), que ya sabe bifurcar por _esNumeroOT.
// ─────────────────────────────────────────────────────────────
function WOS_detectarRespuestasOT() {
  try {
    var hoja  = _getHojaPedidosOT();
    if (!hoja) return { ok: true, procesados: 0 };
    var datos = hoja.getDataRange().getValues();

    var enEspera = {}; // numero → { threadId, rows: [...] }
    for (var i = 1; i < datos.length; i++) {
      var num    = String(datos[i][COL.NUMERO]    || '').trim();
      var estado = String(datos[i][COL.ESTADO]    || '').trim();
      var tId    = String(datos[i][COL.THREAD_ID] || '').trim();
      if (!num || estado !== EST.EN_ESPERA || !tId) continue;
      if (!enEspera[num]) enEspera[num] = { threadId: tId, rows: [] };
      enEspera[num].rows.push(i);
    }

    var numeros = Object.keys(enEspera);
    if (!numeros.length) return { ok: true, procesados: 0 };

    var dominioInterno = '@bidcom.com.ar';
    var procesados = 0;

    for (var n = 0; n < numeros.length; n++) {
      var numero = numeros[n];
      var info   = enEspera[numero];
      try {
        var thread   = GmailApp.getThreadById(info.threadId);
        var messages = thread.getMessages();
        if (!messages || messages.length < 2) continue; // aún no hay respuesta

        var wosDataStr = null, resellerPlain = null, lastExtMsg = null;
        for (var m = 0; m < messages.length; m++) {
          var msg   = messages[m];
          var from  = msg.getFrom().toLowerCase();
          var plain = msg.getPlainBody();
          if (from.indexOf(dominioInterno) >= 0) {
            var matchW = plain.match(/===WOSDATA===\s*([\s\S]*?)\s*===ENDWOSDATA===/);
            if (matchW) wosDataStr = matchW[1].trim();
          } else {
            lastExtMsg    = msg;
            resellerPlain = plain;
          }
        }
        if (!resellerPlain || !lastExtMsg) continue;

        var rLow = resellerPlain.toLowerCase();
        var esA  = /opci[oó]n\s*a\b/.test(rLow);
        var esB  = /opci[oó]n\s*b\b/.test(rLow);
        if (!esA && !esB) {
          Logger.log('WOS_detectarRespuestasOT: ' + numero + ' → respuesta no reconocida, se ignora');
          continue;
        }

        // WOSDATA trae {numero, faltantes:[{sku,cantDisp}]} — mismo shape que `cantidades`
        // que espera WOS_procesarRespuestaManual (viene de _wosNotificarFaltante).
        var cantidades = [];
        if (wosDataStr) {
          try {
            var wosObj = JSON.parse(wosDataStr);
            var fArr   = wosObj.faltantes || [];
            for (var fi = 0; fi < fArr.length; fi++) {
              cantidades.push({ sku: fArr[fi].sku, cantDisp: Number(fArr[fi].cantDisp) || 0 });
            }
          } catch(eJ) { Logger.log('WOS_detectarRespuestasOT JSON parse: ' + eJ); }
        }

        var res = WOS_procesarRespuestaManual(numero, esA ? 'A' : 'B', cantidades,
          'ot_texto_libre', 'ot_txt_' + numero + '_' + info.threadId);
        if (!res || !res.ok) {
          Logger.log('WOS_detectarRespuestasOT: ' + numero + ' → procesarRespuestaManual error: ' + (res && res.error));
          continue;
        }

        thread.markRead();
        procesados++;
        Logger.log('WOS_detectarRespuestasOT: ' + numero + ' → OPCI\xd3N ' + (esA ? 'A' : 'B') + ' procesada por texto libre');
      } catch(eT) {
        Logger.log('WOS_detectarRespuestasOT thread [' + numero + ']: ' + eT);
      }
    }

    Logger.log('WOS_detectarRespuestasOT: procesados=' + procesados + ' de ' + numeros.length);
    return { ok: true, procesados: procesados };
  } catch(e) {
    Logger.log('WOS_detectarRespuestasOT ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}


// Instala el trigger de tiempo (ejecutar UNA VEZ desde el editor)
function WOS_instalarTriggerDetector() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'WOS_detectarRespuestasResellers') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('WOS_detectarRespuestasResellers')
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log('Trigger instalado: WOS_detectarRespuestasResellers cada 10 min');
}



// ─────────────────────────────────────────────────────────────
//  Procesa la respuesta del reseller ingresada manualmente por
//  el operario (gestión telefónica). Replica la lógica del
//  detector automático pero añade "confirmado telefónicamente"
//  en el email de confirmación al reseller.
//
//  opcion:     'A' (esperar faltante) | 'B' (cancelar faltante)
//  cantidades: [{sku, cantDisp}] — requerido solo para opción B
// ─────────────────────────────────────────────────────────────
function WOS_procesarRespuestaManual(numero, opcion, cantidades, operario, reqToken) {
 return _wosLockIdempot(reqToken, function() {
  try {
    operario  = String(operario || '');
    var esOT  = _esNumeroOT(numero);
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    var datos = hoja.getDataRange().getValues();

    var threadId = '', reseller = '';
    var rows = [];
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      if (!reseller) {
        reseller = String(datos[i][COL.RESELLER]  || '');
        threadId = String(datos[i][COL.THREAD_ID] || '').trim();
      }
      rows.push({ rowNum: i + 1, sku: String(datos[i][COL.SKU] || '').toUpperCase() });
    }
    if (!reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };

    var cantMap = {};
    if (cantidades) {
      for (var c = 0; c < cantidades.length; c++) {
        cantMap[String(cantidades[c].sku || '').toUpperCase()] = Number(cantidades[c].cantDisp) || 0;
      }
    }

    var op = String(opcion || '').toUpperCase();
    var nuevoEst, confTitulo, confMensaje;
    // cantMap trae el SKU:cantDisp de los ítems que Hito 2 marcó como faltante (ver
    // WOS_notificarFaltante) — llega tanto en Opción A como B (ambas URLs incluyen
    // &items=...). Sirve para distinguir, en pedidos OT, cuáles filas eran el faltante
    // REAL (estaban en cantMap) de las que solo quedaron retenidas por estar en el mismo
    // pedido (no estaban en cantMap — eran "disponibles" cuando se detectó el faltante).

    if (!esOT) {
      // ── Pedidos de reseller: comportamiento sin cambios ──────────────────────────
      if (op === 'A') {
        // Opción A = esperar faltante en segundo envío
        // Solo los items En_Espera_Reseller pasan a Backorder; los Preparado quedan intactos
        nuevoEst    = EST.BACKORDER;
        confTitulo  = 'Recibimos tu decisi\xf3n — Pedido ' + numero;
        confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar el faltante en un segundo env\xedo.</strong><br>" +
          "Recibir\xe1s la Nota de Entrega por los \xedtems que se despachan ahora. " +
          "El faltante llegar\xe1 en un segundo env\xedo cuando el stock est\xe9 disponible.";
        _wosSetEstadoFiltrado(hoja, datos, numero, EST.EN_ESPERA, nuevoEst);

      } else {
        // Opción B = cancelar faltante → despacho único de lo disponible
        // Solo los items En_Espera_Reseller se modifican: cantDisp>0 → Preparado, cantDisp=0 → Cancelado
        nuevoEst    = EST.PREPARADO;
        confTitulo  = 'Recibimos tu decisi\xf3n — Pedido ' + numero;
        confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Cancelar el faltante.</strong><br>" +
          "Estamos preparando el despacho de lo disponible. Recibir\xe1s la Nota de Entrega cuando sea despachado. Los \xedtems faltantes quedan cancelados.";
        var ahoraM = new Date();
        for (var r = 0; r < rows.length; r++) {
          var _estFilaB = String(datos[rows[r].rowNum - 1][COL.ESTADO] || '').trim();
          // "Cancelar el faltante" aplica a TODA la línea pendiente del faltante: tanto la que aún espera
          // respuesta (En_Espera_Reseller) COMO la que ya quedó aparcada en Backorder. Antes el filtro era
          // sólo En_Espera → en un pedido mixto la línea en Backorder no se cancelaba y "quedaba en
          // backorder" aunque el operario eligiera la Opción B (bug reportado).
          if (_estFilaB !== EST.EN_ESPERA && _estFilaB !== EST.BACKORDER) continue;
          // Trabajar sobre lo REALMENTE pendiente (E - F - Z), no sobre lo solicitado: una línea de
          // Backorder puede tener unidades ya despachadas (F>0) o canceladas (Z>0) → así no se cancela de más.
          var _solB   = Number(datos[rows[r].rowNum - 1][COL.CANT_SOL])    || 0;
          var _despB  = Number(datos[rows[r].rowNum - 1][COL.CANT_DESP])   || 0;
          var _cancB0 = Number(datos[rows[r].rowNum - 1][COL.CANT_CANCEL]) || 0;
          var _pendB  = _solB - _despB - _cancB0;
          if (_pendB <= 0) continue;   // nada pendiente en esta línea
          var cantDispM = (cantMap[rows[r].sku] !== undefined) ? cantMap[rows[r].sku] : 0;
          if (cantDispM > _pendB) cantDispM = _pendB;   // no preparar más de lo pendiente
          var rEstM = hoja.getRange(rows[r].rowNum, COL.ESTADO + 1);
          rEstM.clearDataValidations();
          if (cantDispM >= _pendB) {
            // Hay stock para cubrir todo lo pendiente → se despacha completo, no se cancela nada
            rEstM.setValue(EST.PREPARADO);
          } else if (cantDispM > 0) {
            // Se despacha lo disponible y se CANCELA el resto pendiente (acumula sobre lo ya cancelado)
            hoja.getRange(rows[r].rowNum, COL.CANT_CANCEL + 1).setValue(_cancB0 + (_pendB - cantDispM));
            rEstM.setValue(EST.PREP_PARCIAL);
          } else {
            // Sin stock → se CANCELA todo lo pendiente
            hoja.getRange(rows[r].rowNum, COL.CANT_CANCEL + 1).setValue(_cancB0 + _pendB);
            rEstM.setValue(EST.CANCELADO);
          }
          hoja.getRange(rows[r].rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraM);
        }
      }

    } else if (op === 'A') {
      // ── OT · Opción A = Esperar y consolidar en un solo envío ────────────────────
      // Los SKUs que SÍ eran el faltante real (están en cantMap) → Backorder (2º envío
      // cuando llegue). Los que estaban disponibles y quedaron retenidos (no están en
      // cantMap) → Reservado_Consolidar: NO es "falta comprar", es "ya está, pero se
      // retiene a propósito" — se libera solo cuando el faltante real se resuelva (ver
      // _wosLiberarConsolidarSiSinBackorder).
      nuevoEst    = EST.RESERVADO_CONSOLIDAR;
      confTitulo  = 'Recibimos tu decisi\xf3n — Pedido ' + numero;
      confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar y consolidar en un solo env\xedo.</strong><br>" +
        "Retenemos tambi\xe9n lo que ya est\xe1 disponible. Cuando llegue el resto, va todo junto en un \xfanico env\xedo.";
      var ahoraOA = new Date();
      for (var ra = 0; ra < rows.length; ra++) {
        var _estFilaOA = String(datos[rows[ra].rowNum - 1][COL.ESTADO] || '').trim();
        if (_estFilaOA !== EST.EN_ESPERA) continue;
        var rEstOA = hoja.getRange(rows[ra].rowNum, COL.ESTADO + 1);
        rEstOA.clearDataValidations();
        rEstOA.setValue(cantMap.hasOwnProperty(rows[ra].sku) ? EST.BACKORDER : EST.RESERVADO_CONSOLIDAR);
        hoja.getRange(rows[ra].rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraOA);
      }

    } else {
      // ── OT · Opción B = Despachar ahora lo disponible + el resto después ─────────
      // A diferencia de la Opción B de reseller, ACÁ NUNCA SE CANCELA nada: lo que no se
      // puede cubrir ahora queda en Backorder (2º envío), no Cancelado — para una
      // reparación, cancelar un repuesto que hace falta no tiene sentido.
      nuevoEst    = EST.PREPARADO;
      confTitulo  = 'Recibimos tu decisi\xf3n — Pedido ' + numero;
      confMensaje = "Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Despachar ahora lo disponible.</strong><br>" +
        "El resto llega despu\xe9s, en un segundo env\xedo. No se cancela nada.";
      var ahoraOB = new Date();
      for (var rb = 0; rb < rows.length; rb++) {
        var _estFilaOB = String(datos[rows[rb].rowNum - 1][COL.ESTADO] || '').trim();
        if (_estFilaOB !== EST.EN_ESPERA) continue;
        var rEstOB = hoja.getRange(rows[rb].rowNum, COL.ESTADO + 1);
        rEstOB.clearDataValidations();
        if (!cantMap.hasOwnProperty(rows[rb].sku)) {
          // Nunca fue faltante — estaba disponible y solo se retuvo por consolidación → despachar ya
          rEstOB.setValue(EST.PREPARADO);
          hoja.getRange(rows[rb].rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraOB);
          continue;
        }
        var _solBO   = Number(datos[rows[rb].rowNum - 1][COL.CANT_SOL])    || 0;
        var _despBO  = Number(datos[rows[rb].rowNum - 1][COL.CANT_DESP])   || 0;
        var _cancBO0 = Number(datos[rows[rb].rowNum - 1][COL.CANT_CANCEL]) || 0;
        var _pendBO  = _solBO - _despBO - _cancBO0;
        if (_pendBO <= 0) continue;
        var cantDispBO = (cantMap[rows[rb].sku] !== undefined) ? cantMap[rows[rb].sku] : 0;
        if (cantDispBO > _pendBO) cantDispBO = _pendBO;
        if (cantDispBO >= _pendBO) {
          rEstOB.setValue(EST.PREPARADO);
        } else if (cantDispBO > 0) {
          rEstOB.setValue(EST.PREP_PARCIAL);   // parcial ahora, resto en Backorder — sin CANT_CANCEL
        } else {
          rEstOB.setValue(EST.BACKORDER);      // nada ahora, sigue esperando — sin cancelar
        }
        hoja.getRange(rows[rb].rowNum, COL.FECHA_ESTADO + 1).setValue(ahoraOB);
      }
    }

    SpreadsheetApp.flush();

    // Reservar (Opción A / Consolidar) o liberar (Opción B reseller) las unidades en
    // camino comprometidas. Para OT-Opción B ("despachar ahora") no se cancela nada, así
    // que no hay reservas que liberar — se deja que WOS_reconciliarReservas() reconcilie.
    if (op === 'A') {
      try {
        var _rSumM = _wosReservarEnCamino(numero, reseller);
        confMensaje += _wosMsgReservas(_rSumM);
      } catch(eRs) { Logger.log('procesarRespuestaManual reservar [' + numero + ']: ' + eRs); }
    } else if (!esOT) {
      try { _wosCerrarReservas(numero, '', 'Cancelada'); } catch(eRl) { Logger.log('procesarRespuestaManual liberar [' + numero + ']: ' + eRl); }
    }

    if (threadId) {
      try {
        var confHtml = _wosPortalHead(confTitulo) +
          "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + "</strong>:</p>" +
          "<p style='font-size:13px;color:#555;line-height:1.6'>" + confMensaje + "</p>" +
          "<p style='font-size:12px;color:#888;margin-top:18px'>Consultas: " +
            "<a href='mailto:" + EMAIL_SOPORTE + "' style='color:#00a3e0'>" + EMAIL_SOPORTE + "</a></p>" +
          _wosPortalFoot('Pedido ' + numero + ' · ' + reseller + '.');
        var confPlain = confMensaje.replace(/<[^>]+>/g, '');
        GmailApp.getThreadById(threadId).replyAll(confPlain, {
          htmlBody: confHtml,
          name:     'BIDCOMAGRO · Portal Resellers',
          replyTo:  EMAIL_SOPORTE
        });
      } catch(eM) { Logger.log('WOS_procesarRespuestaManual email [' + numero + ']: ' + eM); }
    }

    _wosLogAccion('Respuesta tel\xe9fonica: Opci\xf3n ' + op + ' → ' + nuevoEst, numero, reseller, operario, '');
    Logger.log('WOS_procesarRespustaManual OK: ' + numero + ' op=' + op + ' → ' + nuevoEst);
    return { ok: true, estado: nuevoEst };
  } catch(e) {
    Logger.log('WOS_procesarRespuestaManual ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
 });
}


// ═══ AVISOS DE INGRESO — cola NOTIF_INGRESOS_WOS (la llena SM al recibir un CAS) ═══════════
// Por cada pedido con unidades bloqueadas que acaban de ingresar:
//  · Si el stock de HOY cubre TODO su backorder → se reactiva solo (Backorder→Preparado,
//    reservas Cumplida) y se avisa al reseller que se lo despachamos — sin preguntar (default).
//  · Si todavía falta otra cosa → mail con lo que llegó + lo que falta (con su ETA) y links:
//    A = despachar ahora lo disponible · B = esperar y recibir todo junto (un solo envío).
// Idempotente por fila (ESTADO Pendiente→Notificado); un error deja la fila Pendiente y se
// reintenta en la próxima corrida del detector.
var _WOS_NOTIF_ING_SHEET = 'NOTIF_INGRESOS_WOS';
// cols: 0 FECHA · 1 CAS · 2 PEDIDO · 3 RESELLER · 4 SKU · 5 CANTIDAD · 6 ESTADO · 7 RESULTADO

function WOS_notificarIngresos() {
  var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_NOTIF_ING_SHEET);
  if (!hoja) return { ok: true, procesadas: 0 };
  var d = hoja.getDataRange().getValues();
  var porPedido = {};
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][6] || '').trim() !== 'Pendiente') continue;
    var ped = String(d[i][2] || '').trim();
    var sku = String(d[i][4] || '').trim().toUpperCase();
    if (!ped || !sku) continue;
    if (!porPedido[ped]) porPedido[ped] = { rows: [], items: {} };
    porPedido[ped].rows.push(i);
    porPedido[ped].items[sku] = (porPedido[ped].items[sku] || 0) + (Number(d[i][5]) || 0);
  }
  var peds = Object.keys(porPedido);
  if (!peds.length) return { ok: true, procesadas: 0 };

  // Stock FRESCO desde Carmen (sin cache: el ingreso puede ser de hace minutos)
  var stockMap = {};
  try {
    var cs = SpreadsheetApp.openById(CARMEN_SS_ID).getSheetByName('STOCK').getDataRange().getValues();
    for (var c = 1; c < cs.length; c++) {
      var cod = String(cs[c][0] || '').trim().toUpperCase();
      if (cod) stockMap[cod] = parseInt(cs[c][2]) || 0;
    }
  } catch(eCs) { Logger.log('WOS_notificarIngresos stock: ' + eCs); }

  var ecMap = {};
  try { var _ec = WOS_getEnCaminoMap(); if (_ec && _ec.ok) ecMap = _ec.map || {}; } catch(eEc) {}

  var procesadas = 0;
  for (var p = 0; p < peds.length; p++) {
    var numero = peds[p], resultado = '';
    try { resultado = _wosNotificarIngresoPedido(numero, porPedido[numero].items, stockMap, ecMap); }
    catch(eP) { Logger.log('WOS_notificarIngresos [' + numero + ']: ' + eP); }
    if (!resultado) continue;
    for (var rI = 0; rI < porPedido[numero].rows.length; rI++) {
      var fila = porPedido[numero].rows[rI] + 1;
      hoja.getRange(fila, 7).setValue('Notificado');
      hoja.getRange(fila, 8).setValue(resultado);
    }
    procesadas += porPedido[numero].rows.length;
  }
  if (procesadas) SpreadsheetApp.flush();
  return { ok: true, procesadas: procesadas };
}


// Evalúa y notifica UN pedido de la cola de ingresos. Devuelve el resultado ('' = error, reintentar).
function _wosNotificarIngresoPedido(numero, llegoItems, stockMap, ecMap) {
  var hoja = _getHojaPorNumero(numero);
  if (!hoja) return 'pedido_no_encontrado';
  var datos = hoja.getDataRange().getValues();
  var reseller = '', threadId = '', boPorSku = {}, descPorSku = {};
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (!reseller) {
      reseller = String(datos[i][COL.RESELLER]  || '');
      threadId = String(datos[i][COL.THREAD_ID] || '').trim();
    }
    var sku = String(datos[i][COL.SKU] || '').trim().toUpperCase();
    if (sku && !descPorSku[sku]) descPorSku[sku] = String(datos[i][COL.DESC] || '');
    if (String(datos[i][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
    var pend = (Number(datos[i][COL.CANT_SOL]) || 0) - (Number(datos[i][COL.CANT_DESP]) || 0) - (Number(datos[i][COL.CANT_CANCEL]) || 0);
    if (sku && pend > 0) boPorSku[sku] = (boPorSku[sku] || 0) + pend;
  }
  if (!reseller) return 'pedido_no_encontrado';
  var skusBO = Object.keys(boPorSku);
  if (!skusBO.length) return 'sin_backorder';   // se resolvió por otro camino (despacho manual, etc.)

  // ¿El stock de hoy (incluye lo recién ingresado) cubre TODO el backorder del pedido?
  var faltan = [], completo = true;
  for (var s = 0; s < skusBO.length; s++) {
    var skuB = skusBO[s];
    var stk  = Number(stockMap[skuB]) || 0;
    if (stk >= boPorSku[skuB]) continue;
    completo = false;
    var em = ecMap[skuB];
    faltan.push({ sku: skuB, desc: descPorSku[skuB] || '', falta: boPorSku[skuB] - Math.max(0, stk),
                  eta: (em && em.etaMin) ? em.etaMin : '' });
  }

  var llegaron = [];
  for (var lk in llegoItems) llegaron.push({ sku: lk, desc: descPorSku[lk] || '', qty: llegoItems[lk] });

  var email = '';
  try { email = _wosGetEmailReseller(reseller); } catch(eEm) {}

  var filasLlego = '';
  for (var fl = 0; fl < llegaron.length; fl++) {
    filasLlego += '<tr>' +
      "<td style='padding:6px 10px;font-size:12px;font-family:monospace'>" + llegaron[fl].sku + '</td>' +
      "<td style='padding:6px 10px;font-size:12px'>" + llegaron[fl].desc + '</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center;font-weight:700;color:#00875a'>" + llegaron[fl].qty + ' u.</td>' +
      '</tr>';
  }
  var tablaLlego =
    "<table style='width:100%;border-collapse:collapse;border:1px solid #d1e7dd;margin:10px 0'>" +
      "<thead><tr style='background:#e7f5ee'>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:left'>C\xf3digo</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:left'>Descripci\xf3n</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:center'>Ingres\xf3</th>" +
      '</tr></thead><tbody>' + filasLlego + '</tbody></table>';

  if (completo) {
    // ── Default: llegó todo → preparar y despachar sin preguntar nada ──
    var reac = null;
    try { reac = WOS_reactivarBackorder(numero, 'auto_ingreso'); } catch(eRe) { Logger.log('notifIngreso reactivar [' + numero + ']: ' + eRe); }
    if (!reac || !reac.ok) return '';   // reintentar en la próxima corrida

    var htmlC = _wosPortalHead('\xa1Lleg\xf3 lo que faltaba! — Pedido ' + numero) +
      "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + '</strong>:</p>' +
      "<p style='font-size:13px;color:#555;line-height:1.6'>Ingres\xf3 a nuestro dep\xf3sito la mercader\xeda que estabas esperando de tu pedido " +
        "<strong style='color:#00a3e0'>" + numero + '</strong>:</p>' +
      tablaLlego +
      "<p style='font-size:13px;color:#1a7f4f;font-weight:700;line-height:1.6'>Tu pedido queda completo: lo estamos preparando y te lo despachamos. No ten\xe9s que hacer nada.</p>" +
      _wosPortalFoot('Pedido ' + numero + ' \xb7 ' + reseller + '.');
    var plainC = 'Hola ' + reseller + ',\n\nIngres\xf3 la mercader\xeda que estabas esperando de tu pedido ' + numero +
      '. Tu pedido queda completo: lo estamos preparando y te lo despachamos. No ten\xe9s que hacer nada.';
    var optsC = { htmlBody: htmlC, name: 'BIDCOMAGRO \xb7 Portal Resellers', replyTo: _wosConfig().emailSoporte };
    if (threadId || email) {
      try {
        var okC = _wosReplyHiloOriginal(threadId, plainC, optsC, [email]);
        if (!okC && email) GmailApp.sendEmail(email, '\xa1Lleg\xf3 lo que faltaba! — Pedido ' + numero, plainC, optsC);
      } catch(eSc) {
        if (email) { try { GmailApp.sendEmail(email, '\xa1Lleg\xf3 lo que faltaba! — Pedido ' + numero, plainC, optsC); } catch(eS2) { Logger.log('notifIngreso mail completo [' + numero + ']: ' + eS2); } }
      }
      return 'auto_completo';
    }
    // Sin hilo ni email (OT / reseller sin mail cargado): el pedido ya se reactivó solo,
    // pero nadie fue avisado → alertar a soporte para que lo contacte a mano (llamado, WhatsApp).
    _wosAlertarSinContacto(numero, reseller, llegaron, [], true);
    return 'auto_completo_sin_contacto';
  }

  // ── Ingreso PARCIAL → preguntar: ¿despachamos ahora o esperás el resto? ──
  if (!threadId && !email) {
    // OTs / clientes sin mail: no hay a quién preguntarle A/B → alertar a soporte, no queda en silencio.
    _wosAlertarSinContacto(numero, reseller, llegaron, faltan, false);
    return 'sin_contacto_alertado';
  }

  var filasFalta = '';
  for (var ff = 0; ff < faltan.length; ff++) {
    filasFalta += '<tr>' +
      "<td style='padding:6px 10px;font-size:12px;font-family:monospace'>" + faltan[ff].sku + '</td>' +
      "<td style='padding:6px 10px;font-size:12px'>" + faltan[ff].desc + '</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center;font-weight:700;color:#B54708'>" + faltan[ff].falta + ' u.</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center'>" + (faltan[ff].eta ? 'llega ~<strong>' + faltan[ff].eta + '</strong>' : 'a confirmar') + '</td>' +
      '</tr>';
  }
  var tablaFalta =
    "<table style='width:100%;border-collapse:collapse;border:1px solid #ffe1bd;margin:10px 0'>" +
      "<thead><tr style='background:#fff3e0'>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:left'>C\xf3digo</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:left'>Descripci\xf3n</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:center'>Falta</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:center'>Reposici\xf3n estim.</th>" +
      '</tr></thead><tbody>' + filasFalta + '</tbody></table>';

  var urlBase = '';
  try { urlBase = ScriptApp.getService().getUrl(); } catch(eU) { Logger.log('notifIngreso URL: ' + eU); }
  var urlA = urlBase ? urlBase + '?page=resp_ingreso&num=' + encodeURIComponent(numero) + '&op=A' : '';
  var urlB = urlBase ? urlBase + '?page=resp_ingreso&num=' + encodeURIComponent(numero) + '&op=B' : '';
  var botones = urlA
    ? "<div style='text-align:center;margin:22px 0 8px'>" +
        "<a href='" + urlA + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#00875a;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>Despachar AHORA lo disponible</a>" +
        "<a href='" + urlB + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#3730a3;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>Esperar y recibir todo junto</a>" +
      '</div>' +
      "<p style='font-size:11px;color:#888;text-align:center;margin:4px 0 0'>Si los botones no funcionan, respond\xe9 este correo y te ayudamos.</p>"
    : "<p style='font-size:12px;color:#555;background:#f5f8fc;border-radius:6px;padding:10px 14px;line-height:1.6'>Respond\xe9 este correo indic\xe1ndonos si despach\xe1s ahora lo disponible o esper\xe1s a que llegue todo.</p>";

  var htmlP = _wosPortalHead('Lleg\xf3 parte de tu pedido — ' + numero) +
    "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + '</strong>:</p>' +
    "<p style='font-size:13px;color:#555;line-height:1.6'>Ingres\xf3 a nuestro dep\xf3sito parte de la mercader\xeda que estabas esperando de tu pedido " +
      "<strong style='color:#00a3e0'>" + numero + '</strong>:</p>' +
    tablaLlego +
    "<p style='font-size:13px;color:#555;margin:14px 0 6px'>Todav\xeda est\xe1 pendiente:</p>" +
    tablaFalta +
    "<p style='font-size:13px;color:#1a1f2e;font-weight:700;margin:18px 0 10px'>\xbfC\xf3mo prefer\xeds recibirlo?</p>" +
    "<div style='border:1px solid #b7e0cd;border-radius:8px;padding:14px 18px;margin-bottom:8px;background:#e7f5ee'>" +
      "<p style='margin:0 0 4px;font-size:13px;color:#00875a;font-weight:700'>Despachar AHORA lo disponible</p>" +
      "<p style='margin:0;font-size:12px;color:#4a5568'>Te enviamos ya lo que ingres\xf3; el resto va en un segundo env\xedo cuando llegue.</p>" +
    '</div>' +
    "<div style='border:1px solid #c7d2fe;border-radius:8px;padding:14px 18px;margin-bottom:16px;background:#eef2ff'>" +
      "<p style='margin:0 0 4px;font-size:13px;color:#3730a3;font-weight:700'>Esperar y recibir todo junto</p>" +
      "<p style='margin:0;font-size:12px;color:#4a5568'>Tus unidades quedan reservadas y despachamos todo en un solo env\xedo cuando ingrese el resto \x2014 un solo costo de env\xedo.</p>" +
    '</div>' +
    botones +
    _wosPortalFoot('Pedido ' + numero + ' \xb7 ' + reseller + '.');

  var plainP = 'Hola ' + reseller + ',\n\nIngres\xf3 parte de la mercader\xeda de tu pedido ' + numero + '.\n\n' +
    'Ya disponible:\n';
  for (var pl = 0; pl < llegaron.length; pl++) plainP += '\x2022 ' + llegaron[pl].sku + ' \x2014 ' + llegaron[pl].desc + ' \x2014 ' + llegaron[pl].qty + ' u.\n';
  plainP += '\nTodav\xeda pendiente:\n';
  for (var pf = 0; pf < faltan.length; pf++) plainP += '\x2022 ' + faltan[pf].sku + ' \x2014 ' + faltan[pf].desc + ' \x2014 ' + faltan[pf].falta + ' u.' + (faltan[pf].eta ? ' (llega ~' + faltan[pf].eta + ')' : ' (fecha a confirmar)') + '\n';
  plainP += '\n\xbfDespachamos ahora lo disponible o esper\xe1s a que llegue todo y va en un solo env\xedo?\n' +
    (urlA ? 'Despachar ahora: ' + urlA + '\nEsperar todo junto: ' + urlB + '\n' : 'Respond\xe9 este correo con tu preferencia.\n');

  var optsP = { htmlBody: htmlP, name: 'BIDCOMAGRO \xb7 Portal Resellers', replyTo: _wosConfig().emailSoporte };
  try {
    var okP = _wosReplyHiloOriginal(threadId, plainP, optsP, [email]);
    if (!okP && email) GmailApp.sendEmail(email, 'Lleg\xf3 parte de tu pedido — ' + numero, plainP, optsP);
    else if (!okP && !email) { _wosAlertarSinContacto(numero, reseller, llegaron, faltan, false); return 'sin_contacto_alertado'; }
  } catch(eSp) {
    if (email) { try { GmailApp.sendEmail(email, 'Lleg\xf3 parte de tu pedido — ' + numero, plainP, optsP); } catch(eS3) { Logger.log('notifIngreso mail parcial [' + numero + ']: ' + eS3); return ''; } }
    else { _wosAlertarSinContacto(numero, reseller, llegaron, faltan, false); return 'sin_contacto_alertado'; }
  }
  try { _wosLogAccion('Aviso de ingreso parcial enviado (\xbfdespachar ahora o esperar?)', numero, reseller, 'auto_ingreso', ''); } catch(eLg) {}
  return 'pregunta_enviada';
}


// Alerta interna a soporte cuando un pedido con unidades bloqueadas recibe mercadería pero no
// hay forma de avisarle al reseller (OT sin email, o reseller sin mail cargado en el Portal).
// Evita que la recepción quede "resuelta en silencio" sin que nadie lo contacte.
function _wosAlertarSinContacto(numero, reseller, llegaron, faltan, completo) {
  try {
    var llegoTxt = llegaron.map(function(x) { return '- ' + x.sku + ' (' + x.desc + ') — ' + x.qty + ' u.'; }).join('\n');
    var faltaTxt = faltan.length
      ? faltan.map(function(x) { return '- ' + x.sku + ' (' + x.desc + ') — ' + x.falta + ' u. pendientes' + (x.eta ? ' (ETA ~' + x.eta + ')' : ''); }).join('\n')
      : '(nada \x2014 el pedido ya qued\xf3 completo y se prepar\xf3 solo)';
    var asunto = '[WOS] Ingreso sin contacto \x2014 Pedido ' + numero + ' (' + reseller + ')';
    var cuerpo = 'Ingres\xf3 mercader\xeda reservada para el pedido ' + numero + ' de ' + reseller +
      ', pero no hay hilo de Gmail ni email cargado para avisarle autom\xe1ticamente.\n\n' +
      'Ingres\xf3 ahora:\n' + llegoTxt + '\n\n' +
      (completo ? '' : 'Todav\xeda pendiente:\n' + faltaTxt + '\n\n') +
      'Contactalo manualmente (tel\xe9fono / WhatsApp)' +
      (completo ? ' para avisarle que se lo despachamos.' : ' y pregunt\xe1le si despachamos ahora lo disponible o espera a que llegue todo junto.');
    GmailApp.sendEmail(_wosConfig().emailSoporte, asunto, cuerpo);
  } catch(e) { Logger.log('_wosAlertarSinContacto: ' + e); }
}


// ═══ RECORDATORIO / ESCALADO — pedidos que no responden al aviso de ingreso parcial ═══════
// Sin esto, un pedido en "pregunta_enviada" puede quedar esperando indefinidamente si el
// reseller nunca contesta ni con click ni por mail. A los _NI_RECORDAR_DIAS se reenvía el
// aviso; si sigue sin respuesta, a los _NI_ESCALAR_DIAS se avisa a soporte para contacto manual.
// Corre en cada pasada del detector (piggyback, sin trigger nuevo); es idempotente porque cada
// etapa solo dispara una vez (avanza el RESULTADO de la fila en NOTIF_INGRESOS_WOS).
var _NI_RECORDAR_DIAS = 3;
var _NI_ESCALAR_DIAS  = 6;

function WOS_recordarIngresosPendientes() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_NOTIF_ING_SHEET);
    if (!hoja) return { ok: true, recordados: 0, escalados: 0 };
    var d = hoja.getDataRange().getValues();
    var DIA = 24 * 60 * 60 * 1000, ahoraMs = Date.now();

    var porPedido = {};
    for (var i = 1; i < d.length; i++) {
      var res = String(d[i][7] || '').trim();
      if (res !== 'pregunta_enviada' && res !== 'recordado') continue;
      var ped = String(d[i][2] || '').trim();
      if (!ped) continue;
      var fRaw = d[i][0];
      var fMs  = (fRaw instanceof Date) ? fRaw.getTime() : 0;
      if (!porPedido[ped]) porPedido[ped] = { rows: [], fechaMin: 0, etapa: res };
      porPedido[ped].rows.push(i);
      if (fMs && (!porPedido[ped].fechaMin || fMs < porPedido[ped].fechaMin)) porPedido[ped].fechaMin = fMs;
      if (res === 'recordado') porPedido[ped].etapa = 'recordado';   // ya se recordó una vez → evaluar escalado
    }

    var recordados = 0, escalados = 0;
    for (var pedNum in porPedido) {
      var info = porPedido[pedNum];
      var dias = info.fechaMin ? (ahoraMs - info.fechaMin) / DIA : 0;
      if (info.etapa === 'pregunta_enviada' && dias >= _NI_RECORDAR_DIAS) {
        if (_wosRecordarIngresoPedido(pedNum)) {
          for (var r1 = 0; r1 < info.rows.length; r1++) hoja.getRange(info.rows[r1] + 1, 8).setValue('recordado');
          recordados++;
        }
      } else if (info.etapa === 'recordado' && dias >= _NI_ESCALAR_DIAS) {
        if (_wosEscalarIngresoPedido(pedNum)) {
          for (var r2 = 0; r2 < info.rows.length; r2++) hoja.getRange(info.rows[r2] + 1, 8).setValue('escalado');
          escalados++;
        }
      }
    }
    if (recordados || escalados) SpreadsheetApp.flush();
    return { ok: true, recordados: recordados, escalados: escalados };
  } catch(e) { Logger.log('WOS_recordarIngresosPendientes: ' + e); return { ok: false, error: e.toString() }; }
}


// ═══ AVISOS DE CAMBIO DE ETA — cola NOTIF_ETA_CAMBIO_WOS (la llena SM en sincronizarItemsCAS
// al detectar, contra el planner externo, que la fecha estimada de un CAS se atrasó) ═══════
// Puramente informativo — no pide elegir A/B como el aviso de ingreso: solo avisa la nueva
// fecha al reseller que tiene una reserva ACTIVA de ese SKU/CAS (RESERVAS_EN_CAMINO). Agrupa
// por pedido (puede haber varios SKUs atrasados del mismo pedido) para mandar 1 solo mail.
// Idempotente por fila (ESTADO Pendiente→Notificado).
var _WOS_NOTIF_ETA_SHEET = 'NOTIF_ETA_CAMBIO_WOS';
// cols: 0 FECHA · 1 CAS · 2 PEDIDO · 3 RESELLER · 4 SKU · 5 CANTIDAD · 6 ETA_ANTERIOR · 7 ETA_NUEVA · 8 ESTADO · 9 RESULTADO

function WOS_notificarCambiosEta() {
  var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_NOTIF_ETA_SHEET);
  if (!hoja) return { ok: true, procesadas: 0 };
  var d = hoja.getDataRange().getValues();
  var porPedido = {};
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][8] || '').trim() !== 'Pendiente') continue;
    var ped = String(d[i][2] || '').trim();
    if (!ped) continue;
    if (!porPedido[ped]) porPedido[ped] = { rows: [], reseller: String(d[i][3] || ''), items: [] };
    porPedido[ped].rows.push(i);
    porPedido[ped].items.push({
      sku:         String(d[i][4] || '').trim().toUpperCase(),
      cantidad:    Number(d[i][5]) || 0,
      etaAnterior: String(d[i][6] || ''),
      etaNueva:    String(d[i][7] || '')
    });
  }
  var peds = Object.keys(porPedido);
  if (!peds.length) return { ok: true, procesadas: 0 };

  var procesadas = 0;
  for (var p = 0; p < peds.length; p++) {
    var numero = peds[p], res = { ok: false, msg: 'error_desconocido' };
    try { res = _wosNotificarCambioEtaPedido(numero, porPedido[numero]); }
    catch(eP) { res = { ok: false, msg: 'excepcion: ' + eP }; Logger.log('WOS_notificarCambiosEta [' + numero + ']: ' + eP); }
    // Siempre dejamos el motivo en RESULTADO (aunque sea un reintento) para poder
    // diagnosticar sin depender de los logs de ejecución de Apps Script.
    for (var rI = 0; rI < porPedido[numero].rows.length; rI++) {
      var fila = porPedido[numero].rows[rI] + 1;
      if (res.ok) hoja.getRange(fila, 9).setValue('Notificado');
      hoja.getRange(fila, 10).setValue(res.msg);
    }
    if (res.ok) procesadas += porPedido[numero].rows.length;
  }
  if (procesadas) SpreadsheetApp.flush();
  return { ok: true, procesadas: procesadas };
}

// Manda UN mail al reseller de `numero` con todos los SKUs de `info.items` que se atrasaron
// (antes/ahora). Devuelve { ok, msg }: ok=false con msg de diagnóstico se reintenta en la
// próxima corrida (queda visible en RESULTADO sin marcar 'Notificado').
function _wosNotificarCambioEtaPedido(numero, info) {
  var hoja = _getHojaPorNumero(numero);
  if (!hoja) return { ok: false, msg: 'pedido_no_encontrado' };
  var datos = hoja.getDataRange().getValues();
  var reseller = info.reseller, threadId = '';
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    threadId = String(datos[i][COL.THREAD_ID] || '').trim();
    if (!reseller) reseller = String(datos[i][COL.RESELLER] || '');
    break;
  }
  if (!reseller) return { ok: false, msg: 'pedido_no_encontrado' };

  var email = '';
  try { email = _wosGetEmailReseller(reseller); } catch(eEm) {}
  if (!threadId && !email) return { ok: false, msg: 'sin_contacto' };

  var filas = '';
  for (var it = 0; it < info.items.length; it++) {
    var x = info.items[it];
    filas += '<tr>' +
      "<td style='padding:6px 10px;font-size:12px;font-family:monospace'>" + x.sku + '</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center'>" + x.cantidad + ' u.</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center;color:#888;text-decoration:line-through'>" + (x.etaAnterior || '—') + '</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center;font-weight:700;color:#B54708'>" + (x.etaNueva || '—') + '</td>' +
    '</tr>';
  }
  var tabla =
    "<table style='width:100%;border-collapse:collapse;border:1px solid #ffe1bd;margin:10px 0'>" +
      "<thead><tr style='background:#fff3e0'>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:left'>C\xf3digo</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:center'>Reservado</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:center'>Antes</th>" +
        "<th style='padding:6px 10px;font-size:11px;text-align:center'>Ahora</th>" +
      '</tr></thead><tbody>' + filas + '</tbody></table>';

  var html = _wosPortalHead('Cambio de fecha estimada — Pedido ' + numero) +
    "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + '</strong>:</p>' +
    "<p style='font-size:13px;color:#555;line-height:1.6'>Nuestro proveedor actualiz\xf3 la fecha estimada de llegada de lo que est\xe1s esperando en tu pedido " +
      "<strong style='color:#00a3e0'>" + numero + '</strong>. Se atras\xf3:</p>' +
    tabla +
    "<p style='font-size:12px;color:#888;line-height:1.6'>No ten\xe9s que hacer nada — es solo un aviso. Te contactamos apenas ingrese.</p>" +
    _wosPortalFoot('Pedido ' + numero + ' \xb7 ' + reseller + '.');

  var plain = 'Hola ' + reseller + ',\n\nNuestro proveedor actualiz\xf3 la fecha estimada de llegada de tu pedido ' + numero + '. Se atras\xf3:\n\n';
  for (var pl = 0; pl < info.items.length; pl++) {
    var xp = info.items[pl];
    plain += '• ' + xp.sku + ' — ' + xp.cantidad + ' u. — antes ~' + (xp.etaAnterior || '?') + ', ahora ~' + (xp.etaNueva || '?') + '\n';
  }
  plain += '\nNo ten\xe9s que hacer nada, es solo un aviso.';

  var opts = { htmlBody: html, name: 'BIDCOMAGRO \xb7 Portal Resellers', replyTo: _wosConfig().emailSoporte };
  try {
    var ok = _wosReplyHiloOriginal(threadId, plain, opts, [email]);
    if (!ok && email) GmailApp.sendEmail(email, 'Cambio de fecha estimada — Pedido ' + numero, plain, opts);
    else if (!ok && !email) return { ok: false, msg: 'sin_contacto' };
  } catch(eS) {
    if (email) { try { GmailApp.sendEmail(email, 'Cambio de fecha estimada — Pedido ' + numero, plain, opts); } catch(eS2) { Logger.log('notifCambioEta mail [' + numero + ']: ' + eS2); return { ok: false, msg: 'excepcion_mail: ' + eS2 }; } }
    else return { ok: false, msg: 'sin_contacto' };
  }
  return { ok: true, msg: 'avisado' };
}


// Datos mínimos de un pedido para recordatorio/escalado: reseller, contacto, y su backorder vivo.
function _wosPedidoInfoBasico(numero) {
  var hoja = _getHojaPorNumero(numero);
  if (!hoja) return null;
  var datos = hoja.getDataRange().getValues();
  var reseller = '', threadId = '', email = '', backorder = [], descPorSku = {};
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
    if (!reseller) { reseller = String(datos[i][COL.RESELLER] || ''); threadId = String(datos[i][COL.THREAD_ID] || '').trim(); }
    var sku = String(datos[i][COL.SKU] || '').trim().toUpperCase();
    if (sku && !descPorSku[sku]) descPorSku[sku] = String(datos[i][COL.DESC] || '');
    if (String(datos[i][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
    var pend = (Number(datos[i][COL.CANT_SOL]) || 0) - (Number(datos[i][COL.CANT_DESP]) || 0) - (Number(datos[i][COL.CANT_CANCEL]) || 0);
    if (sku && pend > 0) backorder.push({ sku: sku, desc: descPorSku[sku] || '', pend: pend });
  }
  if (!reseller) return null;
  try { email = _wosGetEmailReseller(reseller); } catch(eE) {}
  return { reseller: reseller, threadId: threadId, email: email, backorder: backorder };
}


function _wosRecordarIngresoPedido(numero) {
  var info = _wosPedidoInfoBasico(numero);
  if (!info || !info.reseller || (!info.threadId && !info.email)) return false;

  var filas = '';
  for (var i = 0; i < info.backorder.length; i++) {
    var b = info.backorder[i];
    filas += '<tr>' +
      "<td style='padding:6px 10px;font-size:12px;font-family:monospace'>" + b.sku + '</td>' +
      "<td style='padding:6px 10px;font-size:12px'>" + b.desc + '</td>' +
      "<td style='padding:6px 10px;font-size:12px;text-align:center'>" + b.pend + ' u.</td>' +
      '</tr>';
  }
  var tabla = filas
    ? "<table style='width:100%;border-collapse:collapse;border:1px solid #eee;margin:10px 0'>" +
        "<thead><tr style='background:#f5f5f5'>" +
          "<th style='padding:6px 10px;font-size:11px;text-align:left'>C\xf3digo</th>" +
          "<th style='padding:6px 10px;font-size:11px;text-align:left'>Descripci\xf3n</th>" +
          "<th style='padding:6px 10px;font-size:11px;text-align:center'>Pendiente</th>" +
        '</tr></thead><tbody>' + filas + '</tbody></table>'
    : '';

  var urlBase = '';
  try { urlBase = ScriptApp.getService().getUrl(); } catch(eU) {}
  var urlA = urlBase ? urlBase + '?page=resp_ingreso&num=' + encodeURIComponent(numero) + '&op=A' : '';
  var urlB = urlBase ? urlBase + '?page=resp_ingreso&num=' + encodeURIComponent(numero) + '&op=B' : '';
  var botones = urlA
    ? "<div style='text-align:center;margin:18px 0 6px'>" +
        "<a href='" + urlA + "' style='display:inline-block;margin:6px;padding:12px 20px;background:#00875a;color:#fff;border-radius:8px;font-size:13px;font-weight:700;text-decoration:none'>Despachar AHORA lo disponible</a>" +
        "<a href='" + urlB + "' style='display:inline-block;margin:6px;padding:12px 20px;background:#3730a3;color:#fff;border-radius:8px;font-size:13px;font-weight:700;text-decoration:none'>Esperar y recibir todo junto</a>" +
      '</div>'
    : '';

  var html = _wosPortalHead('\xbfSeguimos esperando? — Pedido ' + numero) +
    "<p style='font-size:14px;color:#666;margin:0 0 18px'>Hola <strong>" + info.reseller + '</strong>:</p>' +
    "<p style='font-size:13px;color:#555;line-height:1.6'>Te escribimos de nuevo porque no tuvimos respuesta sobre tu pedido " +
      "<strong style='color:#00a3e0'>" + numero + '</strong>. Todav\xeda tenemos pendiente:</p>' +
    tabla +
    "<p style='font-size:13px;color:#555'>\xbfDespachamos ya lo que est\xe1 disponible, o segu\xeds esperando a que llegue todo junto?</p>" +
    botones +
    _wosPortalFoot('Pedido ' + numero + ' \xb7 ' + info.reseller + '.');
  var plain = 'Hola ' + info.reseller + ',\n\nTe escribimos de nuevo porque no tuvimos respuesta sobre tu pedido ' + numero +
    '. \xbfDespachamos ya lo disponible o segu\xeds esperando a que llegue todo junto?\n' +
    (urlA ? 'Despachar ahora: ' + urlA + '\nEsperar todo junto: ' + urlB + '\n' : '');

  var opts = { htmlBody: html, name: 'BIDCOMAGRO \xb7 Portal Resellers', replyTo: _wosConfig().emailSoporte };
  try {
    var ok = _wosReplyHiloOriginal(info.threadId, plain, opts, [info.email]);
    if (!ok && info.email) GmailApp.sendEmail(info.email, '\xbfSeguimos esperando? — Pedido ' + numero, plain, opts);
    else if (!ok && !info.email) return false;
  } catch(e) {
    if (info.email) { try { GmailApp.sendEmail(info.email, '\xbfSeguimos esperando? — Pedido ' + numero, plain, opts); } catch(e2) { return false; } }
    else return false;
  }
  try { _wosLogAccion('Recordatorio de ingreso parcial enviado', numero, info.reseller, 'auto_recordatorio', ''); } catch(eLg) {}
  return true;
}


function _wosEscalarIngresoPedido(numero) {
  var info = _wosPedidoInfoBasico(numero);
  if (!info || !info.reseller) return false;
  var filasTxt = info.backorder.map(function(b) { return '- ' + b.sku + ' (' + b.desc + ') — ' + b.pend + ' u. pendientes'; }).join('\n');
  var asunto = '[WOS] Sin respuesta \x2014 Pedido ' + numero + ' (' + info.reseller + ')';
  var cuerpo = 'El pedido ' + numero + ' de ' + info.reseller + ' tiene una reposici\xf3n parcial esperando respuesta desde hace ' +
    _NI_ESCALAR_DIAS + '+ d\xedas y el reseller no contest\xf3 (ni click ni respuesta al mail).\n\n' +
    'Pendiente:\n' + (filasTxt || '(sin detalle)') + '\n\n' +
    'Contactalo manualmente para definir: despachar ahora lo disponible o seguir esperando.';
  try { GmailApp.sendEmail(_wosConfig().emailSoporte, asunto, cuerpo); } catch(e) { Logger.log('_wosEscalarIngresoPedido mail: ' + e); return false; }
  try { _wosLogAccion('Escalado a soporte: sin respuesta al aviso de ingreso', numero, info.reseller, 'auto_escalado', ''); } catch(eLg) {}
  return true;
}



// Reemplazada por WOS_despacharCompleto — mantener como stub para no romper referencias externas
function WOS_cerrarBulto(numero) {
  Logger.log('WOS_cerrarBulto: función deprecada, usar WOS_despacharCompleto para ' + numero);
  return { ok: false, error: 'Función deprecada. Usar WOS_despacharCompleto.' };
}



// Reemplazada por WOS_despacharCompleto — mantener como stub para no romper referencias externas
function WOS_despacharYCerrar(numero) {
  Logger.log('WOS_despacharYCerrar: función deprecada, usar WOS_despacharCompleto para ' + numero);
  return { ok: false, error: 'Función deprecada. Usar WOS_despacharCompleto.' };
}
