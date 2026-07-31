// @version 3.25
function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
  if (page === 'manual') {
    return HtmlService.createHtmlOutputFromFile('WOS_Manual')
      .setTitle('WOS · Manual de Operaci\xf3n')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === 'resp_faltante') {
    return _doGetRespFaltante(e ? e.parameter : {});
  }
  if (page === 'resp_ingreso') {
    return _doGetRespIngreso(e ? e.parameter : {});
  }
  if (page === 'confirma_entrega') {
    return _doGetConfirmaEntrega(e ? e.parameter : {});
  }
  return HtmlService.createHtmlOutputFromFile('Despacho_Index')
    .setTitle('WOS · Despacho Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function _doGetRespFaltante(params) {
  var numero = String(params.num || '').trim();
  var opcion = String(params.op  || '').trim().toUpperCase();
  var items  = String(params.items || '').trim();

  if (!numero || (opcion !== 'A' && opcion !== 'B')) {
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Enlace no v\xe1lido', 'El enlace de respuesta no es v\xe1lido. Comunicate con nuestro equipo.'));
  }

  try {
    // Antes hardcodeado a _getHojaPedidos() (solo Pedidos_resellers) — el botón del mail
    // de faltante para pedidos de OT ('OT-...') siempre daba "Pedido no encontrado".
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) {
      return HtmlService.createHtmlOutput(_rfHtml(false, 'Pedido no encontrado', 'No encontramos el pedido ' + numero + '. Comunicate con nuestro equipo.'));
    }
    var datos = hoja.getDataRange().getValues();
    var reseller = '', hayEnEspera = false;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      if (!reseller) reseller = String(datos[i][COL.RESELLER] || '');
      if (String(datos[i][COL.ESTADO] || '').trim() === EST.EN_ESPERA) { hayEnEspera = true; }
    }
    if (!reseller) {
      return HtmlService.createHtmlOutput(_rfHtml(false, 'Pedido no encontrado', 'No encontramos el pedido ' + numero + '. Comunicate con nuestro equipo.'));
    }
    if (!hayEnEspera) {
      return HtmlService.createHtmlOutput(_rfHtml(true, '\xa1Gracias, ' + reseller + '!', 'Tu respuesta ya fue registrada anteriormente. No se realizaron cambios adicionales.'))
        .setTitle('Ya procesado');
    }

    var cantidades = [];
    if (items) {
      var parts = items.split(',');
      for (var p = 0; p < parts.length; p++) {
        var kv = parts[p].split(':');
        if (kv.length === 2 && kv[0]) cantidades.push({ sku: kv[0], cantDisp: Number(kv[1]) || 0 });
      }
    }

    var res = WOS_procesarRespuestaManual(numero, opcion, cantidades, 'reseller_click');
    if (!res.ok) {
      return HtmlService.createHtmlOutput(_rfHtml(false, 'Error al procesar', 'No pudimos registrar tu respuesta. Por favor comunicate con nuestro equipo.'));
    }

    var esOT = _esNumeroOT(numero);
    var msg;
    if (esOT) {
      msg = opcion === 'A'
        ? 'Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar y consolidar en un solo env\xedo.</strong><br>Retenemos tambi\xe9n lo que ya est\xe1 disponible; cuando llegue el resto va todo junto.'
        : 'Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Despachar ahora lo disponible.</strong><br>El resto llega despu\xe9s, en un segundo env\xedo. No se cancela nada.';
    } else {
      msg = opcion === 'A'
        ? 'Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar el faltante en un segundo env\xedo.</strong><br>Despachamos lo disponible a la brevedad y los \xedtems faltantes llegan cuando haya stock.'
        : 'Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Cancelar el faltante.</strong><br>Despachamos lo disponible y los \xedtems faltantes quedan cancelados.';
    }
    return HtmlService.createHtmlOutput(_rfHtml(true, '\xa1Gracias, ' + reseller + '!', msg))
      .setTitle('Respuesta registrada \xb7 Pedido ' + numero);
  } catch(e) {
    Logger.log('_doGetRespFaltante ERROR [' + numero + ']: ' + e);
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Error interno', 'Ocurri\xf3 un error. Comunicate con nuestro equipo.'));
  }
}

// Respuesta del reseller al aviso de INGRESO PARCIAL ("llegó parte de lo que esperabas"):
// op=A → despachar ahora lo que ya está disponible · op=B → esperar y recibir todo junto.
function _doGetRespIngreso(params) {
  var numero = String(params.num || '').trim();
  var opcion = String(params.op  || '').trim().toUpperCase();
  if (!numero || (opcion !== 'A' && opcion !== 'B')) {
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Enlace no v\xe1lido', 'El enlace de respuesta no es v\xe1lido. Comunicate con nuestro equipo.'));
  }
  try {
    var res = WOS_procesarRespuestaIngreso(numero, opcion, 'reseller_click');
    if (!res.ok) {
      return HtmlService.createHtmlOutput(_rfHtml(false, 'Error al procesar', 'No pudimos registrar tu respuesta. Por favor comunicate con nuestro equipo.'));
    }
    var msg;
    if (opcion === 'A') {
      msg = res.reactivados > 0
        ? 'Registramos tu elecci\xf3n: <strong>Despachar ahora lo disponible.</strong><br>Estamos preparando el env\xedo con lo que ya ingres\xf3; el resto llega en un segundo env\xedo cuando ingrese.'
        : 'Registramos tu elecci\xf3n, pero por el momento no hay unidades listas para despachar. Nuestro equipo lo va a revisar y te contactamos.';
    } else {
      msg = 'Registramos tu elecci\xf3n: <strong>Esperar y recibir todo junto en un solo env\xedo.</strong><br>Tus unidades quedan reservadas y te avisamos cuando ingrese el resto.';
    }
    return HtmlService.createHtmlOutput(_rfHtml(true, '\xa1Gracias, ' + res.reseller + '!', msg))
      .setTitle('Respuesta registrada \xb7 Pedido ' + numero);
  } catch(e) {
    Logger.log('_doGetRespIngreso ERROR [' + numero + ']: ' + e);
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Error interno', 'Ocurri\xf3 un error. Comunicate con nuestro equipo.'));
  }
}

// A: reactiva (Backorder→Preparado) las líneas del pedido cuyo SKU tiene stock suficiente HOY
//    para su pendiente (lee Carmen fresco) y cierra sus reservas (Cumplida). Las líneas con
//    stock parcial o sin stock siguen en Backorder con su reserva (las maneja el operador).
// B: no cambia nada — el pedido sigue esperando con sus reservas activas; cuando llegue el
//    resto, la cola de ingresos vuelve a evaluarlo (y si queda completo se prepara solo).
function WOS_procesarRespuestaIngreso(numero, opcion, operario) {
  try {
    numero = String(numero || '').trim();
    var op = (String(opcion || '').trim().toUpperCase() === 'A') ? 'A' : 'B';
    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    var reseller = '';
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() === numero) {
        reseller = String(datos[i][COL.RESELLER] || '');
        break;
      }
    }
    if (!reseller) return { ok: false, error: 'Pedido no encontrado.' };

    var reactivados = 0;
    if (op === 'A') {
      // Stock actual FRESCO desde Carmen (sin cache: el ingreso puede ser de hace segundos)
      var stockMap = {};
      try {
        var cs = SpreadsheetApp.openById(CARMEN_SS_ID).getSheetByName('STOCK').getDataRange().getValues();
        for (var sc = 1; sc < cs.length; sc++) {
          var sCod = String(cs[sc][0] || '').trim().toUpperCase();
          if (sCod) stockMap[sCod] = parseInt(cs[sc][2]) || 0;
        }
      } catch(eSt) { Logger.log('WOS_procesarRespuestaIngreso stock: ' + eSt); }

      // Pendiente total por SKU (puede haber varias filas del mismo SKU)
      var pendPorSku = {};
      for (var p = 1; p < datos.length; p++) {
        if (String(datos[p][COL.NUMERO] || '').trim() !== numero) continue;
        if (String(datos[p][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
        var pSku  = String(datos[p][COL.SKU] || '').trim().toUpperCase();
        var pPend = (Number(datos[p][COL.CANT_SOL]) || 0) - (Number(datos[p][COL.CANT_DESP]) || 0) - (Number(datos[p][COL.CANT_CANCEL]) || 0);
        if (pSku && pPend > 0) pendPorSku[pSku] = (pendPorSku[pSku] || 0) + pPend;
      }

      var ahora = new Date(), skusOk = {};
      for (var r = 1; r < datos.length; r++) {
        if (String(datos[r][COL.NUMERO] || '').trim() !== numero) continue;
        if (String(datos[r][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
        var rSku = String(datos[r][COL.SKU] || '').trim().toUpperCase();
        if (!rSku || !(pendPorSku[rSku] > 0)) continue;
        if ((Number(stockMap[rSku]) || 0) < pendPorSku[rSku]) continue;   // cobertura parcial → queda en backorder
        var rEst = hoja.getRange(r + 1, COL.ESTADO + 1);
        rEst.clearDataValidations();
        rEst.setValue(EST.PREPARADO);
        hoja.getRange(r + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
        datos[r][COL.ESTADO] = EST.PREPARADO;   // mantener `datos` en memoria consistente (ver liberarConsolidar abajo)
        skusOk[rSku] = true;
        reactivados++;
      }
      if (reactivados) {
        SpreadsheetApp.flush();
        for (var sk in skusOk) {
          try { _wosCerrarReservas(numero, sk, 'Cumplida'); } catch(eCR) { Logger.log('respIngreso cerrarReservas: ' + eCR); }
        }
        // OT: si este pedido ya no tiene ninguna fila Backorder, liberar lo retenido por consolidación
        try { _wosLiberarConsolidarSiSinBackorder(hoja, datos, numero); } catch(eLib) { Logger.log('respIngreso liberarConsolidar: ' + eLib); }
      }
    }

    _wosLogAccion('Aviso de ingreso: ' + (op === 'A' ? 'despachar ahora' : 'esperar todo junto') +
      (reactivados ? ' \xb7 ' + reactivados + ' l\xednea(s) \x2192 Preparado' : ''), numero, reseller, String(operario || ''), '');
    try { _wosMarcarRespondidoIngreso(numero, op); } catch(eMk) { Logger.log('marcarRespondidoIngreso: ' + eMk); }
    return { ok: true, opcion: op, reseller: reseller, reactivados: reactivados };
  } catch(e) {
    Logger.log('WOS_procesarRespuestaIngreso ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Marca en NOTIF_INGRESOS_WOS que el pedido YA respondió (A o B) al aviso de ingreso parcial,
// para que el recordatorio/escalado (WOS_recordarIngresosPendientes, en WOS_GmailFlow.js) no
// lo siga persiguiendo. _WOS_NOTIF_ING_SHEET se define en WOS_GmailFlow.js (mismo proyecto).
function _wosMarcarRespondidoIngreso(numero, op) {
  var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_NOTIF_ING_SHEET);
  if (!hoja) return;
  var d = hoja.getDataRange().getValues();
  var val = 'respondido_' + op, changed = false;
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][2] || '').trim() !== numero) continue;
    var res = String(d[i][7] || '').trim();
    if (res !== 'pregunta_enviada' && res !== 'recordado') continue;
    hoja.getRange(i + 1, 8).setValue(val);
    changed = true;
  }
  if (changed) SpreadsheetApp.flush();
}

// Cola de avisos de ingreso (NOTIF_INGRESOS_WOS) para el panel de auditoría en WOS: qué se le
// avisó a cada reseller cuando llegó mercadería bloqueada, y qué resultado tuvo cada aviso.
function WOS_cargarColaIngresos() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_NOTIF_ING_SHEET);
    if (!hoja) return { ok: true, rows: [], resumen: {} };
    var d  = hoja.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var rows = [], resumen = {};
    for (var i = 1; i < d.length; i++) {
      var fecha    = d[i][0];
      var fechaMs  = (fecha instanceof Date) ? fecha.getTime() : 0;
      var fechaStr = (fecha instanceof Date) ? Utilities.formatDate(fecha, tz, 'dd/MM/yyyy HH:mm') : String(fecha || '');
      var resultado = String(d[i][7] || '').trim();
      var estado    = String(d[i][6] || '').trim();
      var key = resultado || estado || '?';
      resumen[key] = (resumen[key] || 0) + 1;
      rows.push({
        fecha: fechaStr, fechaMs: fechaMs, cas: String(d[i][1] || ''), pedido: String(d[i][2] || ''),
        reseller: String(d[i][3] || ''), sku: String(d[i][4] || ''), cantidad: Number(d[i][5]) || 0,
        estado: estado, resultado: resultado
      });
    }
    rows.sort(function(a, b) { return b.fechaMs - a.fechaMs; });
    if (rows.length > 300) rows = rows.slice(0, 300);
    return { ok: true, rows: rows, resumen: resumen };
  } catch(e) {
    Logger.log('WOS_cargarColaIngresos ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Confirmación de recepción por el reseller (link del mail de despacho): r=ok | prob.
// r=ok  → pasa las filas Entregado_Cerrado a Entregado_Confirmado + registra en WOS_QA.
// r=prob → registra el problema en WOS_QA (baja la precisión de resultado) + avisa a soporte.
function _doGetConfirmaEntrega(params) {
  var numero = String(params.num || params.n || '').trim();
  var r      = String(params.r || '').trim().toLowerCase();
  if (!numero || (r !== 'ok' && r !== 'prob')) {
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Enlace no v\xe1lido', 'El enlace de confirmaci\xf3n no es v\xe1lido. Comunicate con nuestro equipo.'));
  }
  try {
    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return HtmlService.createHtmlOutput(_rfHtml(false, 'Pedido no encontrado', 'No encontramos el pedido ' + numero + '. Comunicate con nuestro equipo.'));
    var datos = hoja.getDataRange().getValues();
    var reseller = '';
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() === numero) { reseller = String(datos[i][COL.RESELLER] || ''); break; }
    }
    if (!reseller) return HtmlService.createHtmlOutput(_rfHtml(false, 'Pedido no encontrado', 'No encontramos el pedido ' + numero + '.'));

    _wosRegistrarConfirmacion(numero, reseller, r, String(params.nota || ''));

    if (r === 'ok') {
      var ahora = new Date();
      for (var j = 1; j < datos.length; j++) {
        if (String(datos[j][COL.NUMERO] || '').trim() !== numero) continue;
        if (String(datos[j][COL.ESTADO] || '').trim() === EST.ENTREGADO) {
          var rE = hoja.getRange(j + 1, COL.ESTADO + 1);
          rE.clearDataValidations();
          rE.setValue(EST.ENTREGADO_CONF);
          hoja.getRange(j + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
        }
      }
      SpreadsheetApp.flush();
      _wosLogAccion('Reseller confirm\xf3 recepci\xf3n OK', numero, reseller, 'reseller_click', '');
      return HtmlService.createHtmlOutput(_rfHtml(true, '\xa1Gracias, ' + reseller + '!',
        'Registramos que recibiste el pedido <strong>' + numero + '</strong> <strong>en perfecto estado</strong>. \xa1Gracias por confirmar!'))
        .setTitle('Recepci\xf3n confirmada \xb7 ' + numero);
    }
    _wosLogAccion('Reseller report\xf3 PROBLEMA', numero, reseller, 'reseller_click', '');
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Gracias por avisar, ' + reseller,
      'Registramos que hubo un <strong>problema</strong> con el pedido <strong>' + numero + '</strong>. Nuestro equipo se va a comunicar con vos a la brevedad para resolverlo.'))
      .setTitle('Problema registrado \xb7 ' + numero);
  } catch(e) {
    Logger.log('_doGetConfirmaEntrega ERROR [' + numero + ']: ' + e);
    return HtmlService.createHtmlOutput(_rfHtml(false, 'Error interno', 'Ocurri\xf3 un error. Comunicate con nuestro equipo.'));
  }
}

function _rfHtml(ok, titulo, mensaje) {
  var titleColor = ok ? '#14713a' : '#922b21';
  var icon       = ok ? '&#x2705;' : '&#x274C;';
  return '<!DOCTYPE html><html><head>' +
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<style>body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:#f0f4f8;margin:0;padding:20px;min-height:100vh;display:flex;align-items:center;justify-content:center;box-sizing:border-box}' +
    '.card{background:#fff;border-radius:16px;padding:36px 28px;max-width:420px;width:100%;text-align:center;box-shadow:0 4px 24px rgba(0,0,0,.08)}' +
    '.ic{font-size:52px;margin-bottom:14px}h1{font-size:20px;color:' + titleColor + ';margin:0 0 12px;font-weight:700}' +
    'p{font-size:14px;color:#4a5568;line-height:1.65;margin:0}.brand{font-size:11px;color:#b0b8c1;margin-top:24px}</style>' +
    '</head><body><div class="card">' +
    '<div class="ic">' + icon + '</div>' +
    '<h1>' + titulo + '</h1>' +
    '<p>' + mensaje + '</p>' +
    '<div class="brand">BIDCOMAGRO &middot; Portal Resellers</div>' +
    '</div></body></html>';
}

function WOS_getManualUrl() {
  return ScriptApp.getService().getUrl() + '?page=manual';
}

// ── Identifica al operario logueado ──────────────────────────
function WOS_getUsuario() {
  try {
    var email = Session.getActiveUser().getEmail() || '';
    var emailL = email.toLowerCase();
    if (emailL) {
      var datos = SpreadsheetApp.openById(MASTER_SS_ID)
                    .getSheetByName('Usuarios_Internos').getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        if (String(datos[i][1] || '').trim().toLowerCase() === emailL) {
          return { email: email, nombre: String(datos[i][0] || email), tipo: String(datos[i][2] || ''), esAdmin: String(datos[i][2] || '').trim() === 'Admin' };
        }
      }
    }
    return { email: email, nombre: email || 'Desconocido', tipo: '', esAdmin: false };
  } catch(e) {
    Logger.log('WOS_getUsuario: ' + e);
    return { email: '', nombre: 'Desconocido', tipo: '' };
  }
}

// ── Log de acciones por operario ─────────────────────────────
// Escribe en hoja 'WOS_Log' del spreadsheet de NOTAS.
// Si la hoja no existe la crea con encabezados.
function _wosLogAccion(accion, numero, reseller, operario, detalle) {
  try {
    var ss   = SpreadsheetApp.openById(NOTAS_SS_ID);
    var hoja = ss.getSheetByName('WOS_Log');
    if (!hoja) {
      hoja = ss.insertSheet('WOS_Log');
      hoja.getRange(1, 1, 1, 6).setValues([['Fecha', 'Pedido', 'Reseller', 'Accion', 'Operario', 'Detalle']])
          .setFontWeight('bold').setBackground('#00a3e0').setFontColor('#ffffff');
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([new Date(), numero || '', reseller || '', accion || '', operario || '', detalle || '']);
  } catch(e) {
    Logger.log('_wosLogAccion: ' + e);
  }
}

// ══════════════════════════════════════════════════════════════
//  CALIDAD / PRECISIÓN (armado guiado)
//  Hoja WOS_QA: 1 fila por cierre de caja + confirmación del reseller.
// ══════════════════════════════════════════════════════════════
var WOS_QA_SHEET = 'WOS_QA';

function _wosGetHojaQA() {
  var ss   = SpreadsheetApp.openById(NOTAS_SS_ID);
  var hoja = ss.getSheetByName(WOS_QA_SHEET);
  if (!hoja) {
    hoja = ss.insertSheet(WOS_QA_SHEET);
    hoja.getRange(1, 1, 1, 15).setValues([[
      'Fecha', 'Pedido', 'Reseller', 'Operario', 'Items', 'Unidades',
      'CantOk', 'SnDupOk', 'EnvioOk', 'Overrides', 'Resultado', 'TiempoSeg',
      'Confirmacion', 'FechaConfirmacion', 'NotaProblema'
    ]]).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// Chequea N° de serie duplicados: internos (en la lista que se va a despachar) y
// externos (contra SERIALES ya despachados en OTROS pedidos). Devuelve la lista de duplicados.
function WOS_verificarSnDuplicados(numero, seriales) {
  try {
    numero = String(numero || '').trim();
    var lista = (Object.prototype.toString.call(seriales) === '[object Array]')
      ? seriales : String(seriales || '').split(',');
    var norm = [];
    for (var i = 0; i < lista.length; i++) {
      // quitar sufijo " xN" de bolsas → el código en sí es lo único
      var s = String(lista[i] || '').trim().replace(/\s+x\d+$/i, '').toUpperCase();
      if (s) norm.push(s);
    }
    // 1. Duplicados internos
    var seen = {}, internos = [];
    for (var j = 0; j < norm.length; j++) {
      if (seen[norm[j]]) { if (internos.indexOf(norm[j]) === -1) internos.push(norm[j]); }
      else seen[norm[j]] = true;
    }
    // 2. Duplicados contra otros pedidos ya despachados
    var externos = [];
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    for (var h = 0; h < hojas.length; h++) {
      var datos = hojas[h].getDataRange().getValues();
      for (var r = 1; r < datos.length; r++) {
        if (String(datos[r][COL.NUMERO] || '').trim() === numero) continue; // no chequear contra sí mismo
        var serStr = String(datos[r][COL.SERIALES] || '').trim();
        if (!serStr) continue;
        var partes = serStr.split(',');
        for (var m = 0; m < partes.length; m++) {
          var code = partes[m].trim().replace(/\s+x\d+$/i, '').toUpperCase();
          if (code && seen[code] && externos.indexOf(code) === -1) externos.push(code);
        }
      }
    }
    var dups = internos.slice();
    for (var e2 = 0; e2 < externos.length; e2++) if (dups.indexOf(externos[e2]) === -1) dups.push(externos[e2]);
    return { ok: true, hayDuplicados: dups.length > 0, duplicados: dups, internos: internos, externos: externos };
  } catch(e) {
    Logger.log('WOS_verificarSnDuplicados: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Registra el resultado de un cierre de caja guiado. qa: {reseller, operario, items, unidades,
// cantOk, snDupOk, envioOk, overrides, resultado, tiempoSeg}. Best-effort (no rompe el despacho).
function WOS_registrarQA(numero, qa) {
  try {
    qa = qa || {};
    var hoja = _wosGetHojaQA();
    var b = function(v) { return v === false ? 'NO' : 'S\xcd'; };
    var overridesStr = qa.overrides
      ? (typeof qa.overrides === 'string' ? qa.overrides : JSON.stringify(qa.overrides)) : '';
    var hayOverride = overridesStr && overridesStr !== '[]' && overridesStr !== '{}' && overridesStr !== '""';
    var resultado = qa.resultado ||
      ((qa.cantOk !== false && qa.snDupOk !== false && qa.envioOk !== false && !hayOverride) ? 'limpio' : 'flag');
    hoja.appendRow([
      new Date(), String(numero || ''), String(qa.reseller || ''), String(qa.operario || ''),
      Number(qa.items) || 0, Number(qa.unidades) || 0,
      b(qa.cantOk), b(qa.snDupOk), b(qa.envioOk),
      overridesStr, resultado, Number(qa.tiempoSeg) || 0,
      '', '', ''
    ]);
    return { ok: true, resultado: resultado };
  } catch(e) {
    Logger.log('WOS_registrarQA: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Actualiza la fila de WOS_QA del pedido con la confirmación del reseller (ok/problema).
// Si no hay fila (despacho manual viejo), crea una mínima para no perder el dato de resultado.
function _wosRegistrarConfirmacion(numero, reseller, r, nota) {
  try {
    var conf  = (r === 'ok') ? 'ok' : 'problema';
    var hoja  = _wosGetHojaQA();
    var datos = hoja.getDataRange().getValues();
    var target = -1, fallback = -1;
    for (var i = datos.length - 1; i >= 1; i--) {
      if (String(datos[i][1] || '').trim() !== numero) continue;
      if (fallback === -1) fallback = i;
      if (!String(datos[i][12] || '').trim()) { target = i; break; }  // col 13 (idx 12) = Confirmacion
    }
    if (target === -1) target = fallback;
    if (target >= 1) {
      hoja.getRange(target + 1, 13).setValue(conf);
      hoja.getRange(target + 1, 14).setValue(new Date());
      if (nota) hoja.getRange(target + 1, 15).setValue(nota);
    } else {
      hoja.appendRow([new Date(), numero, reseller || '', '', 0, 0, '', '', '', '', '', 0, conf, new Date(), nota || '']);
    }
  } catch(e) { Logger.log('_wosRegistrarConfirmacion: ' + e); }
}

// Métricas de precisión combinadas (proceso + resultado del reseller), global y por operador.
function WOS_getPrecisionMetrics(desdeISO, hastaISO) {
  try {
    var hoja  = _wosGetHojaQA();
    var datos = hoja.getDataRange().getValues();
    var desde = desdeISO ? new Date(desdeISO) : null;
    var hasta = hastaISO ? new Date(hastaISO) : null;
    var totalCierres = 0, limpios = 0, conf = 0, confOk = 0, preciso = 0;
    var porOp = {}, flags = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i][0];
      if (!(f instanceof Date)) continue;
      if (desde && f < desde) continue;
      if (hasta && f > hasta) continue;
      var pedido    = String(datos[i][1] || '');
      var op        = String(datos[i][3] || '') || '(sin operador)';
      var resultado = String(datos[i][10] || '').trim();   // col 11
      var confir    = String(datos[i][12] || '').trim();   // col 13
      if (!porOp[op]) porOp[op] = { cierres: 0, limpios: 0, conf: 0, confOk: 0, preciso: 0 };
      if (resultado) {   // fila con Resultado = cierre de caja guiado
        totalCierres++; porOp[op].cierres++;
        var esLimpio = (resultado === 'limpio');
        if (esLimpio) { limpios++; porOp[op].limpios++; }
        if (esLimpio && confir !== 'problema') { preciso++; porOp[op].preciso++; }
        if (!esLimpio || confir === 'problema') {
          flags.push({ pedido: pedido, operario: op, fechaMs: f.getTime(), resultado: resultado,
                       confirmacion: confir, overrides: String(datos[i][9] || ''), nota: String(datos[i][14] || '') });
        }
      }
      if (confir === 'ok' || confir === 'problema') {
        conf++; porOp[op].conf++;
        if (confir === 'ok') { confOk++; porOp[op].confOk++; }
      }
    }
    var pct = function(a, b) { return b > 0 ? Math.round((a / b) * 1000) / 10 : null; };
    var ranking = [];
    for (var o in porOp) if (porOp.hasOwnProperty(o)) {
      ranking.push({
        operario:     o,
        cierres:      porOp[o].cierres,
        procesoPct:   pct(porOp[o].limpios, porOp[o].cierres),
        resultadoPct: pct(porOp[o].confOk, porOp[o].conf),
        totalPct:     pct(porOp[o].preciso, porOp[o].cierres)
      });
    }
    ranking.sort(function(a, b) { return (b.totalPct || 0) - (a.totalPct || 0); });
    flags.sort(function(a, b) { return b.fechaMs - a.fechaMs; });
    var tz = Session.getScriptTimeZone();
    flags = flags.slice(0, 30).map(function(fl) {
      return { pedido: fl.pedido, operario: fl.operario, resultado: fl.resultado, confirmacion: fl.confirmacion,
               overrides: fl.overrides, nota: fl.nota,
               fecha: Utilities.formatDate(new Date(fl.fechaMs), tz, 'dd/MM HH:mm') };
    });
    return {
      ok: true,
      totalCierres: totalCierres,
      procesoPct:   pct(limpios, totalCierres),
      confRespondidos: conf,
      resultadoPct: pct(confOk, conf),
      totalPct:     pct(preciso, totalCierres),
      ranking: ranking,
      flags: flags
    };
  } catch(e) {
    Logger.log('WOS_getPrecisionMetrics: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Helper para probar desde el editor.
function WOS_previewPrecision() {
  var r = WOS_getPrecisionMetrics(null, null);
  Logger.log('Precisi\xf3n: ' + JSON.stringify(r, null, 2));
  return r;
}

// ── Setup: actualiza la data validation de columna J ──────────
// Ejecutar UNA VEZ desde el editor cuando se agregan estados nuevos.
function WOS_actualizarValidacion() {
  var estados = [
    'Pendiente_Revision', 'Confirmado', 'En_Espera_Reseller',
    'Cancelado', 'Preparado', 'Backorder', 'Preparado Parcial',
    'Reservado_Consolidar',
    'Entregado_Cerrado', 'Listo_Retiro', 'Entregado_Confirmado'
  ];
  var regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(estados, true)
    .setAllowInvalid(false)
    .build();

  var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
  var totalFilas = 0;
  for (var h = 0; h < hojas.length; h++) {
    var ultima = hojas[h].getLastRow();
    if (ultima < 2) continue;
    hojas[h].getRange(2, COL.ESTADO + 1, ultima - 1, 1).setDataValidation(regla);
    totalFilas += ultima - 1;
  }
  SpreadsheetApp.flush();
  Logger.log('WOS_actualizarValidacion OK: ' + totalFilas + ' filas · ' + estados.length + ' estados');
  return 'OK — ' + estados.join(', ');
}

// ── Email: busca el email del reseller en la hoja MASTER ──────
// ── Etiqueta de envío: datos del pedido + reseller ────────────
function WOS_getEtiquetaData(numero) {
  try {
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado' };
    var datos = hoja.getDataRange().getValues();
    var reseller = '', envio = '', obs = '';
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      reseller = reseller || String(datos[i][COL.RESELLER] || '').trim();
      envio    = envio    || String(datos[i][COL.ENVIO]    || '').trim();
      obs      = obs      || String(datos[i][COL.OBS]      || '').trim();
    }
    if (!reseller) return { ok: false, error: 'Pedido no encontrado' };

    // Buscar datos del reseller en MASTER
    var rsDatos = SpreadsheetApp.openById(MASTER_SS_ID)
      .getSheetByName('Resellers').getDataRange().getValues();
    var rLow = reseller.toLowerCase();
    var direccion = '', cp = '', localidad = '', provincia = '', telefono = '', cuit = '';
    for (var j = 1; j < rsDatos.length; j++) {
      if (String(rsDatos[j][0] || '').trim().toLowerCase() !== rLow) continue;
      cuit      = String(rsDatos[j][1] || '').trim();
      direccion = String(rsDatos[j][2] || '').trim();
      cp        = String(rsDatos[j][3] || '').trim();
      localidad = String(rsDatos[j][4] || '').trim();
      provincia = String(rsDatos[j][5] || '').trim();
      telefono  = String(rsDatos[j][6] || '').trim();
      break;
    }
    return {
      ok:        true,
      numero:    numero,
      reseller:  reseller,
      cuit:      cuit,
      direccion: direccion,
      cp:        cp,
      localidad: localidad,
      provincia: provincia,
      telefono:  telefono,
      envio:     envio,
      obs:       obs
    };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ── Email fallback: cancela sin threadId (nuevo email) ───────
function _enviarEmailEstado(numero, reseller, obs) {
  var email = _wosGetEmailReseller(reseller);
  if (!email) { Logger.log('_enviarEmailEstado: sin email para ' + reseller); return; }

  var asunto = 'Tu pedido ' + numero + ' fue cancelado — BIDCOMAGRO';
  var html = _wosPortalHead('Pedido cancelado — ' + numero) +
    "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + "</strong>:</p>" +
    "<p style='font-size:13px;color:#555;margin:0 0 14px'>Tu pedido <strong style='color:#e74c3c'>" + numero + "</strong> fue <strong>cancelado</strong>.</p>" +
    (obs ? "<div style='background:#fdecea;border:1px solid #f5a5a5;border-radius:8px;padding:12px 16px;margin-bottom:16px'><p style='margin:0;font-size:12px;color:#7f1919'><strong>Motivo:</strong> " + obs + "</p></div>" : '') +
    "<p style='font-size:13px;color:#555'>Si cre\xe9s que esto es un error, respond\xe9 este email y te ayudamos.</p>" +
    _wosPortalFoot('Pedido ' + numero + ' \xb7 ' + reseller + '.');

  try {
    GmailApp.sendEmail(email, asunto, '', {
      htmlBody: html,
      name:     'BIDCOMAGRO \xb7 Portal Resellers',
      replyTo:  _wosConfig().emailSoporte
    });
    Logger.log('WOS email [cancelado fallback] → ' + email + ' pedido ' + numero);
  } catch(e) {
    Logger.log('_enviarEmailEstado ERROR: ' + e);
  }
}

// ── Procesa las filas de una hoja de pedidos (Pedidos_resellers o Pedidos_OTs,
//    mismo layout COL) acumulando en mapaP/orden compartidos ──
function _procesarFilasPedidos(datos, stockMap, mapaP, orden) {
  for (var i = 1; i < datos.length; i++) {
    var r   = datos[i];
    var num = String(r[COL.NUMERO] || '').trim();
    if (!num) continue;

    if (!mapaP[num]) {
      var fechaRaw = r[COL.FECHA];
      var fechaStr = (fechaRaw instanceof Date)
        ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
        : String(fechaRaw || '');

      var feRaw   = r[COL.FECHA_ESTADO];
      // feMs: última acción del WOS (col S). Si nunca fue tocado, cae a fecha del pedido (col K).
      var feMs    = (feRaw instanceof Date) ? feRaw.getTime()
                  : (fechaRaw instanceof Date) ? fechaRaw.getTime() : 0;
      // fechaMs: siempre la fecha en que el reseller hizo el pedido (col K) — base del SLA.
      var fechaMs = (fechaRaw instanceof Date) ? fechaRaw.getTime() : 0;

      var fdRaw = r[COL.FECHA_DESPACHO];
      mapaP[num] = {
        numero:          num,
        reseller:        String(r[COL.RESELLER]         || ''),
        fecha:           fechaStr,
        fechaEstadoMs:   feMs,
        fechaMs:         fechaMs,
        envio:           String(r[COL.ENVIO]            || ''),
        pago:            String(r[COL.PAGO]             || ''),
        obs:             String(r[COL.OBS]              || ''),
        tracking:        String(r[COL.TRACKING]         || '').trim(),
        notaEntrega:     String(r[COL.NOTA_ENTREGA]     || '').trim(),
        neUrl:           String(r[COL.NE_URL]           || '').trim(),
        transportista:   String(r[COL.TRANSPORTISTA_DESP] || '').trim(),
        fechaDespacho:   (fdRaw instanceof Date) ? Utilities.formatDate(fdRaw, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
        esOT:            _esNumeroOT(num),
        pesoPrep:        Number(r[COL.PESO_PREP]) || 0,   // peso del paquete capturado al preparar
        items:           []
      };
      orden.push(num);
    }

    var cantSol    = Number(r[COL.CANT_SOL])    || 0;
    var cantDesp   = Number(r[COL.CANT_DESP])   || 0;
    var cantCancel = Number(r[COL.CANT_CANCEL]) || 0;
    var cantPendRaw = r[COL.CANT_PEND];
    // Primario: col G (fórmula =E−F−Z). Fallback (G vacía): E−F−Z, NO E−F, para no inflar el
    // pendiente con lo cancelado por el reseller (Opción B).
    var cantPend = (cantPendRaw !== '' && cantPendRaw !== null && cantPendRaw !== undefined)
                   ? Number(cantPendRaw) : (cantSol - cantDesp - cantCancel);
    if (isNaN(cantPend)) cantPend = cantSol - cantDesp - cantCancel;

    var skuKey = String(r[COL.SKU] || '').trim().toUpperCase();
    mapaP[num].items.push({
      row:         i + 1,
      sku:         String(r[COL.SKU]  || ''),
      desc:        String(r[COL.DESC] || ''),
      cantSol:     cantSol,
      cantDesp:    cantDesp,
      cantPend:    cantPend,
      precio:      Number(r[COL.PRECIO]) || 0,
      stockOri:    (r[COL.STOCK_ORI] !== '' && r[COL.STOCK_ORI] !== null && !isNaN(Number(r[COL.STOCK_ORI])))
                   ? Number(r[COL.STOCK_ORI]) : -1,
      stockActual:  (skuKey && stockMap[skuKey] !== undefined) ? stockMap[skuKey] : null,
      cantCancel:   cantCancel,
      seriales:     String(r[COL.SERIALES] || '').trim(),
      estado:       String(r[COL.ESTADO] || '')
    });
  }
}

// ── Carga todos los pedidos agrupados por número — fusiona Pedidos_resellers + Pedidos_OTs ──
function WOS_cargarPedidos() {
  try {
    var hojaRes = _getHojaPedidos();
    var hojaOT  = _getHojaPedidosOT();
    if (!hojaRes) return { ok: false, error: 'Hoja "' + HOJA_PEDIDOS + '" no encontrada.' };

    // Stock actual desde CARMEN (A=SKU, C=stock) — cache 5 min para evitar openById extra
    var stockMap = {};
    try {
      var _scCache = CacheService.getScriptCache();
      var _scCached = _scCache.get('wos_carmen_stock_v1');
      if (_scCached) {
        stockMap = JSON.parse(_scCached);
      } else {
        var carmenData = SpreadsheetApp.openById(CARMEN_SS_ID)
                           .getSheetByName('STOCK').getDataRange().getValues();
        for (var sc = 1; sc < carmenData.length; sc++) {
          var sCod = String(carmenData[sc][0] || '').trim().toUpperCase();
          if (sCod) stockMap[sCod] = parseInt(carmenData[sc][2]) || 0;
        }
        try { _scCache.put('wos_carmen_stock_v1', JSON.stringify(stockMap), 300); } catch(ec) {}
      }
    } catch(eSC) { Logger.log('WOS_cargarPedidos stockMap: ' + eSC); }

    var mapaP = {}, orden = [];
    _procesarFilasPedidos(hojaRes.getDataRange().getValues(), stockMap, mapaP, orden);
    if (hojaOT) _procesarFilasPedidos(hojaOT.getDataRange().getValues(), stockMap, mapaP, orden);

    // Cancelado NO entra en la prioridad: un pedido con ítems entregados + alguno cancelado
    // es un pedido ENTREGADO (parcialmente cancelado), no un pedido cancelado.
    var EST_PRIORIDAD = [EST.PENDIENTE, EST.CONFIRMADO, EST.EN_ESPERA, EST.PREPARADO, EST.PREP_PARCIAL, EST.BACKORDER, EST.RESERVADO_CONSOLIDAR];

    var result = [];
    for (var k = 0; k < orden.length; k++) {
      var ped = mapaP[orden[k]];

      var estSet = {};
      for (var m = 0; m < ped.items.length; m++) estSet[ped.items[m].estado] = true;
      ped.estado = ped.items.length ? ped.items[ped.items.length - 1].estado : '';
      var _prioMatch = false;
      for (var ep = 0; ep < EST_PRIORIDAD.length; ep++) {
        if (estSet[EST_PRIORIDAD[ep]]) { ped.estado = EST_PRIORIDAD[ep]; _prioMatch = true; break; }
      }
      // Sin ítems accionables pendientes → pedido cerrado: prioriza lo ENTREGADO sobre lo cancelado.
      // Cancelado solo queda como estado del pedido si TODOS los ítems están cancelados.
      if (!_prioMatch) {
        if      (estSet[EST.ENTREGADO])      ped.estado = EST.ENTREGADO;      // Entregado_Cerrado
        else if (estSet[EST.LISTO_RETIRO])   ped.estado = EST.LISTO_RETIRO;
        else if (estSet[EST.ENTREGADO_CONF]) ped.estado = EST.ENTREGADO_CONF; // Entregado_Confirmado
        else if (estSet[EST.CANCELADO])      ped.estado = EST.CANCELADO;
      }
      ped.esMixto   = Object.keys(estSet).length > 1;
      ped.estadoSet = estSet;

      ped.totalItems = ped.items.length;
      ped.totalSol   = 0;
      ped.totalDesp  = 0;
      ped.totalPend  = 0;
      ped.totalUSD   = 0; // valor total del pedido (cantSol × precio)
      ped.pendUSD    = 0; // valor aún pendiente de despacho (cantPend × precio)
      for (var j = 0; j < ped.items.length; j++) {
        ped.totalSol  += ped.items[j].cantSol;
        ped.totalDesp += ped.items[j].cantDesp;
        ped.totalPend += ped.items[j].cantPend;
        ped.totalUSD  += ped.items[j].precio * ped.items[j].cantSol;
        ped.pendUSD   += ped.items[j].precio * ped.items[j].cantPend;
      }
      result.push(ped);
    }

    result.reverse();
    return { ok: true, pedidos: result };
  } catch(e) {
    Logger.log('WOS_cargarPedidos: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Cancela un pedido con motivo, lo guarda en OBS y envía email
function WOS_cancelarPedido(numero, motivo, operario, reqToken) {
 return _wosLockIdempot(reqToken, function() {
  try {
    motivo   = String(motivo   || '').trim();
    operario = String(operario || '').trim();
    if (motivo) {
      var hoja  = _getHojaPorNumero(numero);
      if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
      var datos = hoja.getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
        var obsActual = String(datos[i][COL.OBS] || '').trim();
        var obsNueva  = motivo + (obsActual ? ' · ' + obsActual : '');
        hoja.getRange(i + 1, COL.OBS + 1).setValue(obsNueva);
      }
      SpreadsheetApp.flush();
    }
    return WOS_cambiarEstado(numero, EST.CANCELADO, operario);
  } catch(e) {
    Logger.log('WOS_cancelarPedido: ' + e);
    return { ok: false, error: e.toString() };
  }
 });
}

// ── Cambia el estado + graba timestamp + envía email si aplica
function WOS_cambiarEstado(numero, nuevoEstado, operario) {
  try {
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    var ahora = new Date();
    operario  = String(operario || '');

    var canReseller = '', canObs = '', canThreadId = '';

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      hoja.getRange(i + 1, COL.ESTADO      + 1).setValue(nuevoEstado);
      hoja.getRange(i + 1, COL.FECHA_ESTADO + 1).setValue(ahora);
      if (operario) hoja.getRange(i + 1, COL.OPERARIO + 1).setValue(operario);
      // Al CANCELAR, registrar en CANT_CANCEL (Z) lo que quedaba sin despachar (E−F) → CANT_PEND
      // (=E−F−Z) baja a 0. Sin esto la línea Cancelada seguía mostrando pendiente fantasma > 0 y
      // había que correr WOS_aplicarCanceladosSinCancel a mano. (Z=E−F absorbe cancelaciones parciales
      // previas; las unidades ya despachadas F quedan como entregadas, no se cancelan.)
      if (nuevoEstado === EST.CANCELADO) {
        var _cZ = Math.max(0, (Number(datos[i][COL.CANT_SOL]) || 0) - (Number(datos[i][COL.CANT_DESP]) || 0));
        hoja.getRange(i + 1, COL.CANT_CANCEL + 1).setValue(_cZ > 0 ? _cZ : '');
      }
      if (!canReseller) {
        canReseller = String(datos[i][COL.RESELLER]  || '');
        canObs      = String(datos[i][COL.OBS]       || '');
        canThreadId = String(datos[i][COL.THREAD_ID] || '').trim();
      }
    }
    SpreadsheetApp.flush();
    _wosLogAccion('Estado: ' + nuevoEstado, numero, canReseller, operario, '');

    if (nuevoEstado === EST.CANCELADO) {
      if (canThreadId) {
        var canHtml = _wosPortalHead('Pedido cancelado — ' + numero) +
          "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + canReseller + "</strong>:</p>" +
          "<p style='font-size:13px;color:#555;margin:0 0 14px'>Tu pedido " +
            "<strong style='color:#e74c3c'>" + numero + "</strong> fue <strong>cancelado</strong>." +
            (canObs ? " Motivo: <em>" + canObs + "</em>." : '') + "</p>" +
          "<p style='font-size:13px;color:#555'>Si tenés alguna consulta, respondé este email o escribinos a " +
            "<a href='mailto:" + _wosConfig().emailSoporte + "' style='color:#00a3e0'>" + _wosConfig().emailSoporte + "</a>.</p>" +
          _wosPortalFoot('Pedido ' + numero + ' · ' + canReseller + '.');
        var canPlain = 'Hola ' + canReseller + ',\n\nTu pedido ' + numero + ' fue cancelado.' +
          (canObs ? '\nMotivo: ' + canObs : '') + '\n\nConsultas: ' + _wosConfig().emailSoporte;
        try {
          // Reply al hilo apuntando a los destinatarios ORIGINALES (no a una conversación aparte),
          // garantizando que el reseller reciba la cancelación.
          var _canOk = _wosReplyHiloOriginal(canThreadId, canPlain, {
            htmlBody: canHtml,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  _wosConfig().emailSoporte
          }, [_wosGetEmailReseller(canReseller)]);
          if (!_canOk) _enviarEmailEstado(numero, canReseller, canObs);
        } catch(eReply) {
          Logger.log('WOS_cambiarEstado cancelacion reply error: ' + eReply);
          _enviarEmailEstado(numero, canReseller, canObs);
        }
      } else {
        _enviarEmailEstado(numero, canReseller, canObs);
      }
    }

    return { ok: true };
  } catch(e) {
    Logger.log('WOS_cambiarEstado: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── MAESTRO DE ARTÍCULOS (auto-construido) ────────────────────
// Registra qué SKU se prepara "por bolsa" (1 código SO por bolsa de N unidades) para
// que la próxima preparación de ese SKU venga pre-marcada. Se llena solo con el uso.

// Devuelve la hoja MAESTRO_ARTICULOS, creándola con encabezado si no existe.
function _wosGetHojaMaestro() {
  var ss = SpreadsheetApp.openById(NOTAS_SS_ID);
  var hoja = ss.getSheetByName(HOJA_MAESTRO);
  if (!hoja) {
    hoja = ss.insertSheet(HOJA_MAESTRO);
    hoja.appendRow(['SKU', 'Descripci\xf3n', 'Por bolsa', 'Bulto x defecto', '\xdaltima actualizaci\xf3n', 'Operador']);
    hoja.getRange(1, 1, 1, 6).setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function _invalidarMaestroCache() {
  try { CacheService.getScriptCache().remove('wos_maestro_v1'); } catch(e) {}
}

// Mapa { SKU(mayúsculas) → { porBolsa:bool, bulto:N } } para pre-marcar el modal de preparación.
function WOS_maestroArticulos() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('wos_maestro_v1');
    if (cached) { try { return { ok: true, map: JSON.parse(cached) }; } catch(e) {} }
    var hoja = _wosGetHojaMaestro();
    var data = hoja.getDataRange().getValues();
    var map = {};
    for (var i = 1; i < data.length; i++) {
      var sku = String(data[i][COL_MAESTRO.SKU] || '').trim();
      if (!sku) continue;
      var pb = data[i][COL_MAESTRO.POR_BOLSA];
      var s  = String(pb).trim().toUpperCase();
      var porBolsa = (pb === true || s === 'TRUE' || s === 'SI' || s === 'S\xcd' || s === '1');
      map[sku.toUpperCase()] = { porBolsa: porBolsa, bulto: parseInt(data[i][COL_MAESTRO.BULTO], 10) || 0 };
    }
    try { cache.put('wos_maestro_v1', JSON.stringify(map), 300); } catch(e) {}
    return { ok: true, map: map };
  } catch(e) {
    Logger.log('WOS_maestroArticulos: ' + e);
    return { ok: false, error: e.toString(), map: {} };
  }
}

// Upsert por SKU. list: [{ sku, desc, porBolsa, bulto }]. Refleja siempre la última elección
// del operario (marcar o desmarcar por bolsa). Nunca rompe la preparación (todo en try/catch).
function _wosUpsertMaestro(list, operario) {
  try {
    if (!list || !list.length) return;
    var hoja = _wosGetHojaMaestro();
    var data = hoja.getDataRange().getValues();
    var idx = {};
    for (var i = 1; i < data.length; i++) {
      var s = String(data[i][COL_MAESTRO.SKU] || '').trim().toUpperCase();
      if (s) idx[s] = i;
    }
    var now = new Date();
    for (var j = 0; j < list.length; j++) {
      var it  = list[j];
      var sku = String(it.sku || '').trim();
      if (!sku) continue;
      var su = sku.toUpperCase();
      var fila;
      if (idx[su] !== undefined) {
        fila = idx[su] + 1;
      } else {
        hoja.appendRow([sku, it.desc || '', '', '', '', '']);
        fila = hoja.getLastRow();
        idx[su] = fila - 1;
      }
      hoja.getRange(fila, COL_MAESTRO.SKU       + 1).setValue(sku);
      if (it.desc) hoja.getRange(fila, COL_MAESTRO.DESC + 1).setValue(it.desc);
      hoja.getRange(fila, COL_MAESTRO.POR_BOLSA + 1).setValue(it.porBolsa ? true : false);
      var bulto = parseInt(it.bulto, 10) || 0;
      if (bulto > 0) hoja.getRange(fila, COL_MAESTRO.BULTO + 1).setValue(bulto);
      hoja.getRange(fila, COL_MAESTRO.FECHA + 1).setValue(now);
      if (operario) hoja.getRange(fila, COL_MAESTRO.OPERADOR + 1).setValue(operario);
    }
    _invalidarMaestroCache();
  } catch(e) {
    Logger.log('_wosUpsertMaestro: ' + e);
  }
}

// ── Círculo cerrado con Stock Manager: resolver una bolsa por su código SO ─────
// SM registra en SO_ETIQUETAS (planilla MASTER) los códigos de bolsa con su cantidad de
// unidades (col CANTIDAD). Al escanear "SO-...-000041" en la preparación por bolsa, WOS
// trae esa cantidad para autocompletarla. Devuelve { ok, cantidad, sku } (cantidad 0 = no es bolsa).
function WOS_resolverBolsa(codigo) {
  try {
    codigo = String(codigo || '').trim().toUpperCase();
    if (!codigo) return { ok: true, cantidad: 0 };
    var map = _wosMapaBolsas();
    var hit = map[codigo];
    if (hit && hit.cantidad > 1) return { ok: true, cantidad: hit.cantidad, sku: hit.sku };
    return { ok: true, cantidad: 0 };
  } catch(e) {
    Logger.log('WOS_resolverBolsa: ' + e);
    return { ok: false, cantidad: 0, error: e.toString() };
  }
}

// Mapa { SO(mayúsculas) → {cantidad, sku} } SOLO de las bolsas (CANTIDAD>1) de SO_ETIQUETAS.
// Se cachea 2 min (el mapa es chico: excluye los SO de 1 unidad).
function _wosMapaBolsas() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get('wos_bolsas_v1');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  var map = {};
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('SO_ETIQUETAS');
    if (hoja) {
      var data = hoja.getDataRange().getValues();
      // Layout SM SO_ETIQUETAS: SO(0), SKU(1), DESC(2), FECHA(3), OPERADOR(4), CANTIDAD(5)
      for (var i = 1; i < data.length; i++) {
        var cant = parseInt(data[i][5], 10) || 0;
        if (cant > 1) {
          var so = String(data[i][0] || '').trim().toUpperCase();
          if (so) map[so] = { cantidad: cant, sku: String(data[i][1] || '').trim() };
        }
      }
    }
  } catch(e) { Logger.log('_wosMapaBolsas: ' + e); }
  try { var s = JSON.stringify(map); if (s.length < 90000) cache.put('wos_bolsas_v1', s, 120); } catch(e) {}
  return map;
}

// Marca el pedido como Preparado registrando el N° de serie / SO de CADA unidad o bolsa (OBLIGATORIO).
// Esto crea el manifiesto para poder auditar despachos (picker dice 10, reseller dice 4).
// Para consumibles (tornillos/gaskets) un mismo código representa toda una bolsa: en la lista de
// seriales queda como "SO-123 x50" y qtyPrep (unidades reales) lo manda el front (suma de bolsas).
// seriales: [{ row, qtyPrep, seriales:"SN1, SN2" | "SO-1 x50, SO-2 x20", ubicaciones, sku, desc, porBolsa, bulto }]
function WOS_prepararConSeriales(numero, seriales, operario, peso) {
  try {
    numero = String(numero || '').trim();
    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    var ahora = new Date();
    operario  = String(operario || '');
    var pesoNum = Number(peso) || 0;   // peso exacto del paquete (kg), capturado al preparar

    // getRange NO auto-expande columnas: asegurar que existan las cols AA/AB (UBIC_PREP/PESO_PREP)
    // antes de escribirlas. Insertar al final (después de Z) no desplaza A–Z ni afecta CANT_PEND (=E-F-Z).
    var _maxCol = hoja.getMaxColumns();
    if (_maxCol < COL.PESO_PREP + 1) hoja.insertColumnsAfter(_maxCol, (COL.PESO_PREP + 1) - _maxCol);

    // fila (1-indexada) → seriales / ubicaciones (bins) / cantidad preparada
    var serMap  = {};
    var ubicMap = {};
    var qtyMap  = {};   // fila → unidades preparadas (para estado Preparado vs Preparado Parcial)
    var maestroUpd = [];   // {sku, desc, porBolsa, bulto} para aprender qué SKU va por bolsa
    if (Object.prototype.toString.call(seriales) === '[object Array]') {
      for (var s = 0; s < seriales.length; s++) {
        var rw = parseInt(seriales[s].row, 10);
        if (rw > 0) {
          serMap[rw]  = String(seriales[s].seriales || '').trim();
          var bins = seriales[s].ubicaciones;
          ubicMap[rw] = (Object.prototype.toString.call(bins) === '[object Array]') ? bins : [];
          var qp = parseInt(seriales[s].qtyPrep, 10);
          if (!(qp > 0)) qp = serMap[rw] ? serMap[rw].split(',').length : 0;   // fallback: nº de SN
          qtyMap[rw] = qp;
          if (seriales[s].sku && seriales[s].porBolsa !== undefined) {
            maestroUpd.push({
              sku:      String(seriales[s].sku).trim(),
              desc:     String(seriales[s].desc || '').trim(),
              porBolsa: !!seriales[s].porBolsa,
              bulto:    parseInt(seriales[s].bulto, 10) || 0
            });
          }
        }
      }
    }

    var reseller = '', tocadas = 0;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      var fila = i + 1;
      // Solo tocar las filas enviadas en el payload; las demás (ej. sin stock en un backorder
      // parcial) quedan como están (Backorder). En el caso normal el payload trae todos los pendientes.
      if (qtyMap[fila] === undefined || qtyMap[fila] <= 0) continue;
      var estActual = String(datos[i][COL.ESTADO] || '');
      // No tocar filas ya cerradas (entregadas / canceladas)
      if (estActual === EST.ENTREGADO || estActual === EST.CANCELADO || estActual === EST.ENTREGADO_CONF) continue;
      // Estado según cantidad preparada vs pendiente de esa fila.
      // cantPend = solicitado - despachado - cancelado (igual que la col G del sheet y el front).
      var _cantSol    = Number(datos[i][COL.CANT_SOL])    || 0;
      var _cantDesp   = Number(datos[i][COL.CANT_DESP])   || 0;
      var _cantCancel = Number(datos[i][COL.CANT_CANCEL]) || 0;
      var _cantPend   = _cantSol - _cantDesp - _cantCancel;
      // Si la línea está En_Espera_Reseller (faltante esperando al reseller) o Reservado_Consolidar
      // (OT: disponible pero retenido a propósito hasta consolidar en 1 envío), preparar lo disponible
      // NO la saca de ese estado: se registran SN + peso pero el estado se mantiene para que la
      // respuesta A/B o la liberación por consolidación la sigan resolviendo. Resto: Preparado /
      // Preparado Parcial según lo preparado.
      var _estNuevo = (estActual === EST.EN_ESPERA || estActual === EST.RESERVADO_CONSOLIDAR) ? estActual
                    : ((qtyMap[fila] >= _cantPend) ? EST.PREPARADO : EST.PREP_PARCIAL);
      hoja.getRange(fila, COL.ESTADO       + 1).setValue(_estNuevo);
      hoja.getRange(fila, COL.FECHA_ESTADO + 1).setValue(ahora);
      if (operario) hoja.getRange(fila, COL.OPERARIO + 1).setValue(operario);
      if (serMap[fila] !== undefined && serMap[fila] !== '') {
        hoja.getRange(fila, COL.SERIALES + 1).setValue(serMap[fila]);
      }
      // Bins elegidos al preparar → col AA (JSON). El descuento del WMS se aplica al despachar.
      if (ubicMap[fila] !== undefined) {
        var _binsFila = ubicMap[fila] || [];
        hoja.getRange(fila, COL.UBIC_PREP + 1).setValue(_binsFila.length ? JSON.stringify(_binsFila) : '');
      }
      // Peso exacto del paquete → col AB. Se pide al preparar para que no se olvide al despachar.
      hoja.getRange(fila, COL.PESO_PREP + 1).setValue(pesoNum > 0 ? pesoNum : '');
      if (!reseller) reseller = String(datos[i][COL.RESELLER] || '');
      tocadas++;
    }
    if (!tocadas) return { ok: false, error: 'No se encontraron filas activas del pedido ' + numero + '.' };

    // Aprender qué SKU se prepara por bolsa (no rompe la preparación si falla).
    _wosUpsertMaestro(maestroUpd, operario);

    SpreadsheetApp.flush();
    _wosLogAccion('Preparado · N\xba de serie registrados', numero, reseller, operario, '');
    return { ok: true, filas: tocadas };
  } catch(e) {
    Logger.log('WOS_prepararConSeriales: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Reactiva solo los ítems en Backorder → Preparado, sin tocar los Entregado_Cerrado
// del primer despacho. Reemplaza el uso de WOS_cambiarEstado para la acción "reactivar".
function WOS_reactivarBackorder(numero, operario) {
  try {
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    _wosSetEstadoFiltrado(hoja, datos, numero, EST.BACKORDER, EST.PREPARADO);
    SpreadsheetApp.flush();
    // El backorder reactivado ya tiene su stock → cerrar las reservas del pedido (Cumplida)
    try { _wosCerrarReservas(numero, '', 'Cumplida'); } catch(eCR) { Logger.log('reactivarBackorder cerrarReservas: ' + eCR); }
    // OT: si este pedido ya no tiene ninguna fila Backorder, cualquier línea retenida por
    // consolidación (Reservado_Consolidar) se libera junto — lectura fresca, requerida.
    try {
      var datosFrescos = hoja.getDataRange().getValues();
      _wosLiberarConsolidarSiSinBackorder(hoja, datosFrescos, numero);
    } catch(eLib) { Logger.log('reactivarBackorder liberarConsolidar: ' + eLib); }
    var reseller = '';
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() === numero) {
        reseller = String(datos[i][COL.RESELLER] || '');
        break;
      }
    }
    _wosLogAccion('Backorder reactivado', numero, reseller, String(operario || ''), '');
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_reactivarBackorder: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Reemplazada por WOS_despacharCompleto (en WOS_GmailFlow.js)
function WOS_despacharPedido(numero) {
  Logger.log('WOS_despacharPedido: función deprecada, usar WOS_despacharCompleto para ' + numero);
  return { ok: false, error: 'Función deprecada. Usar WOS_despacharCompleto.' };
}

// ── Busca pedidos en Backorder que necesitan un SKU ───────────
// Devuelve lista de { numero, reseller, cantPend } ordenada por fecha ASC (FIFO)
function WOS_buscarBackorderPorSKU(sku) {
  try {
    var skuUp  = String(sku || '').trim().toUpperCase();
    if (!skuUp) return { ok: false, error: 'SKU vacío.' };

    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    var mapa = {}; // numero → { reseller, cantPend, fechaMs }
    for (var h = 0; h < hojas.length; h++) {
      var datos = hojas[h].getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        var num    = String(datos[i][COL.NUMERO]   || '').trim();
        var estado = String(datos[i][COL.ESTADO]   || '').trim();
        var rowSku = String(datos[i][COL.SKU]      || '').trim().toUpperCase();
        if (!num || estado !== EST.BACKORDER || rowSku !== skuUp) continue;

        var cantSol    = Number(datos[i][COL.CANT_SOL])    || 0;
        var cantDesp   = Number(datos[i][COL.CANT_DESP])   || 0;
        var cantCancel = Number(datos[i][COL.CANT_CANCEL]) || 0;
        // Pendiente real = E−F−Z: descontar lo cancelado por el reseller para no inflar la lista de compras.
        var cantPend = Math.max(0, cantSol - cantDesp - cantCancel);
        if (cantPend <= 0) continue;

        if (!mapa[num]) {
          var fRaw = datos[i][COL.FECHA];
          mapa[num] = {
            numero:   num,
            reseller: String(datos[i][COL.RESELLER] || ''),
            cantPend: 0,
            fechaMs:  (fRaw instanceof Date) ? fRaw.getTime() : 0
          };
        }
        mapa[num].cantPend += cantPend;
      }
    }

    var _resv = _wosReservasActivas();
    var lista = [];
    var keys  = Object.keys(mapa);
    for (var k = 0; k < keys.length; k++) {
      var _p = mapa[keys[k]];
      _p.reservado = (_resv.byPedidoSku[_p.numero] && _resv.byPedidoSku[_p.numero][skuUp]) ? _resv.byPedidoSku[_p.numero][skuUp] : 0;
      lista.push(_p);
    }
    // Reservados primero (unidades ya comprometidas a ese reseller), luego FIFO por fecha
    lista.sort(function(a, b) {
      var ar = a.reservado > 0 ? 0 : 1, br = b.reservado > 0 ? 0 : 1;
      if (ar !== br) return ar - br;
      return a.fechaMs - b.fechaMs;
    });

    return { ok: true, sku: skuUp, pedidos: lista };
  } catch(e) {
    Logger.log('WOS_buscarBackorderPorSKU ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Reactiva pedidos en Backorder cuando llega stock ─────────
// numeros: array de números de pedido a pasar a Preparado
// Notifica a cada reseller en su hilo original
function WOS_recibirMercaderia(sku, cantRecibida, numeros) {
  try {
    var skuUp  = String(sku          || '').trim().toUpperCase();
    var cant   = Number(cantRecibida || 0);
    if (!skuUp)       return { ok: false, error: 'SKU vacío.' };
    if (!numeros || !numeros.length) return { ok: false, error: 'Sin pedidos seleccionados.' };

    var reactivados = [];
    for (var n = 0; n < numeros.length; n++) {
      var numero = String(numeros[n] || '').trim();
      var hoja   = _getHojaPorNumero(numero);
      if (!hoja) continue;
      var datos  = hoja.getDataRange().getValues();
      var reseller = '', threadId = '';
      for (var i = 1; i < datos.length; i++) {
        if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
        if (!reseller) {
          reseller = String(datos[i][COL.RESELLER]  || '');
          threadId = String(datos[i][COL.THREAD_ID] || '').trim();
        }
      }
      if (!reseller) continue;

      // Reactivar SOLO las líneas en Backorder de ESTE SKU (es el stock que llegó). Nunca tocar
      // líneas ya Entregadas/Canceladas ni backorders de otros SKUs sin stock. Antes usaba
      // _wosSetEstado (sin filtro) → pisaba TODAS las filas del pedido a Preparado y revivía
      // entregados/cancelados (riesgo de doble despacho / descancelar).
      var _rmAhora = new Date(), _rmFlip = 0;
      for (var rm = 1; rm < datos.length; rm++) {
        if (String(datos[rm][COL.NUMERO] || '').trim() !== numero) continue;
        if (String(datos[rm][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
        if (String(datos[rm][COL.SKU]    || '').trim().toUpperCase() !== skuUp) continue;
        var _rmR = hoja.getRange(rm + 1, COL.ESTADO + 1);
        _rmR.clearDataValidations();
        _rmR.setValue(EST.PREPARADO);
        hoja.getRange(rm + 1, COL.FECHA_ESTADO + 1).setValue(_rmAhora);
        datos[rm][COL.ESTADO] = EST.PREPARADO;   // mantener `datos` en memoria consistente (ver liberarConsolidar abajo)
        _rmFlip++;
      }
      if (!_rmFlip) continue;   // este pedido no tenía backorder de ese SKU → no reactivar/notificar

      // La unidad reservada de este SKU llegó y se reactivó → cerrar la reserva (Cumplida)
      try { _wosCerrarReservas(numero, skuUp, 'Cumplida'); } catch(eCR) { Logger.log('recibirMercaderia cerrarReservas: ' + eCR); }
      // OT: si este pedido ya no tiene ninguna fila Backorder, liberar lo retenido por consolidación
      try { _wosLiberarConsolidarSiSinBackorder(hoja, datos, numero); } catch(eLib) { Logger.log('recibirMercaderia liberarConsolidar: ' + eLib); }

      if (threadId) {
        try {
          var html = _wosPortalHead('Stock disponible — Pedido ' + numero) +
            "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + reseller + "</strong>:</p>" +
            "<p style='font-size:13px;color:#555;line-height:1.6'>" +
              "El stock de <strong>" + skuUp + "</strong> que estabas esperando llegó. " +
              "Estamos preparando tu pedido <strong style='color:#00a3e0'>" + numero + "</strong> para el despacho " +
              "y te avisaremos en cuanto sea enviado." +
            "</p>" +
            _wosPortalFoot('Pedido ' + numero + ' · ' + reseller + '.');
          var plain = 'Hola ' + reseller + ',\n\nEl stock de ' + skuUp + ' llegó y estamos preparando tu pedido ' + numero + ' para el despacho. Te avisaremos cuando sea enviado.';
          // Reply al hilo apuntando a los destinatarios ORIGINALES (no a una conversación aparte).
          var _notifOk = _wosReplyHiloOriginal(threadId, plain, {
            htmlBody: html,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  _wosConfig().emailSoporte
          }, [_wosGetEmailReseller(reseller)]);
          if (!_notifOk) {
            var _emNotif = _wosGetEmailReseller(reseller);
            if (_emNotif) GmailApp.sendEmail(_emNotif, 'Stock disponible — Pedido ' + numero, plain, { htmlBody: html, name: 'BIDCOMAGRO · Portal Resellers', replyTo: _wosConfig().emailSoporte });
          }
        } catch(eN) { Logger.log('WOS_recibirMercaderia notif [' + numero + ']: ' + eN); }
      }
      reactivados.push(numero);
    }

    SpreadsheetApp.flush();
    Logger.log('WOS_recibirMercaderia: SKU=' + skuUp + ' cant=' + cant + ' → reactivados: ' + reactivados.join(', '));
    return { ok: true, reactivados: reactivados };
  } catch(e) {
    Logger.log('WOS_recibirMercaderia ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Consulta de stock para operarios ─────────────────────────
// Parsea "dd/MM/yyyy" o "dd/MM" (año actual) → Date; cualquier otra cosa → null.
function _wosEtaToDate(s) {
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

// Formatea una ETA (Date o texto) a "dd/MM/yyyy" limpio para mostrar. '0'/vacío → ''.
// Robustez: Sheets puede convertir el texto de la celda FECHA_ETA en una fecha real; sin esto
// se mostraría "Tue Aug 04 2026 00:00:00 GMT-0300 (…)".
function _wosEtaFmt(v) {
  var dt = _wosEtaToDate(v);
  if (dt) { try { return Utilities.formatDate(dt, Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch(e) {} }
  var s = String(v == null ? '' : v).trim();
  return (s === '0') ? '' : s;
}

// Retorna el mapa de unidades en camino { SKU: { total, ocs, batches, etaMin } }.
//   ocs     = { CAS: qty }               (compat con consumidores previos)
//   batches = [{ cas, air, eta, qty }]   ordenados por fecha asc (sin fecha al final)
//   etaMin  = fecha (string) del lote más próximo con ETA
// Versión ligera usada por Lista de Compras (evita cargar todo el stock).
function WOS_getEnCaminoMap() {
  try {
    var master      = SpreadsheetApp.openById(MASTER_SS_ID);
    var enCaminoMap = {};

    var dCAS = master.getSheetByName('COMPRAS_DJI');
    var dDET = master.getSheetByName('COMPRAS_DETALLE');
    if (!dCAS || !dDET) return { ok: true, map: {} };

    var casData = dCAS.getDataRange().getValues();
    var casActivos = {};
    for (var c = 1; c < casData.length; c++) {
      var casId  = String(casData[c][0] || '').trim().toUpperCase();
      var estado = String(casData[c][2] || '').trim();
      if (casId && estado !== 'En depósito' && estado.indexOf('Borrador') < 0) {
        casActivos[casId] = true;
      }
    }

    var detData = dDET.getDataRange().getValues();
    for (var d = 1; d < detData.length; d++) {
      var dCas = String(detData[d][0] || '').trim().toUpperCase();
      var dSku = String(detData[d][1] || '').trim().toUpperCase();
      var dPed = Number(detData[d][3]) || 0;
      var dRec = Number(detData[d][4]) || 0;
      if (!dSku || !casActivos[dCas]) continue;
      var pend = Math.max(0, dPed - dRec);
      if (pend > 0) {
        var dEta = _wosEtaFmt(detData[d][6]);            // FECHA_ETA (col G) → dd/MM/yyyy limpio
        var dAir = String(detData[d][7] || '').trim();   // N_AIR    (col H)
        if (!enCaminoMap[dSku]) enCaminoMap[dSku] = { total: 0, ocs: {}, _bmap: {} };
        enCaminoMap[dSku].total += pend;
        enCaminoMap[dSku].ocs[dCas] = (enCaminoMap[dSku].ocs[dCas] || 0) + pend;
        var _bKey = dCas + '|' + dEta + '|' + dAir;
        var _bm   = enCaminoMap[dSku]._bmap;
        if (!_bm[_bKey]) _bm[_bKey] = { cas: dCas, air: dAir, eta: dEta, qty: 0 };
        _bm[_bKey].qty += pend;
      }
    }

    // Reservas activas (unidades ya prometidas a un reseller que aceptó esperar)
    var _resv = _wosReservasActivas();

    // Consolidar lotes por SKU: ordenar por fecha (sin fecha al final) + etaMin,
    // y calcular disponibilidad restando las reservas activas por CAS.
    for (var _sk in enCaminoMap) {
      var _em  = enCaminoMap[_sk];
      var _tmp = [];
      for (var _bk in _em._bmap) { var _bb = _em._bmap[_bk]; _tmp.push({ b: _bb, dt: _wosEtaToDate(_bb.eta) }); }
      _tmp.sort(function(a, b) {   // fecha precalculada (no re-parsear en cada comparación)
        if (a.dt && b.dt) return a.dt - b.dt;
        if (a.dt) return -1;
        if (b.dt) return 1;
        return 0;
      });
      var _arr = [];
      for (var _ti = 0; _ti < _tmp.length; _ti++) _arr.push(_tmp[_ti].b);
      _em.batches = _arr;
      _em.etaMin  = (_arr.length && _arr[0].eta) ? _arr[0].eta : '';

      // Disponible = lotes menos reservas activas (repartidas por CAS, lote más próximo primero)
      var _rc = _resv.byCasSku[_sk] || {};
      var _rem = {}; for (var _c in _rc) _rem[_c] = _rc[_c];
      var _bDisp = [], _disp = 0, _etaMinDisp = '';
      for (var _bi3 = 0; _bi3 < _arr.length; _bi3++) {
        var _b3   = _arr[_bi3];
        var _take = Math.min(_rem[_b3.cas] || 0, _b3.qty);
        _rem[_b3.cas] = (_rem[_b3.cas] || 0) - _take;
        var _av = _b3.qty - _take;
        if (_av > 0) {
          _bDisp.push({ cas: _b3.cas, air: _b3.air, eta: _b3.eta, qty: _av });
          _disp += _av;
          if (!_etaMinDisp && _b3.eta) _etaMinDisp = _b3.eta;
        }
      }
      _em.batchesDisp = _bDisp;
      _em.disponible  = _disp;
      _em.reservado   = _em.total - _disp;
      _em.etaMinDisp  = _etaMinDisp;
      delete _em._bmap;
    }

    // Stock actual desde CARMEN — cache 5 min, clave compartida con WOS_cargarPedidos
    var stockMap = {};
    try {
      var cache = CacheService.getScriptCache();
      var cached = cache.get('wos_carmen_stock_v1');
      if (cached) {
        stockMap = JSON.parse(cached);
      } else {
        var carmenStock = SpreadsheetApp.openById(CARMEN_SS_ID)
                            .getSheetByName('STOCK').getDataRange().getValues();
        for (var sc = 1; sc < carmenStock.length; sc++) {
          var sCod = String(carmenStock[sc][0] || '').trim().toUpperCase();
          if (sCod) stockMap[sCod] = parseInt(carmenStock[sc][2]) || 0;
        }
        try { cache.put('wos_carmen_stock_v1', JSON.stringify(stockMap), 300); } catch(eCp) {}
      }
    } catch(eSC) { Logger.log('WOS_getEnCaminoMap stockMap: ' + eSC); }

    return { ok: true, map: enCaminoMap, stockMap: stockMap };
  } catch(e) {
    Logger.log('WOS_getEnCaminoMap ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ═══ RESERVAS DE UNIDADES EN CAMINO (compras DJI) ═══════════════════════
// Cuando un reseller acepta "Esperar" (Opción A) se reservan las unidades del/los
// lote(s) DJI más próximos, para no prometer las mismas unidades a dos resellers.
// Hoja RESERVAS_EN_CAMINO (MASTER): FECHA·PEDIDO·RESELLER·SKU·CAS·N_AIR·ETA·CANTIDAD·ESTADO
//   ESTADO: Activa (comprometida) · Cumplida (llegó y se reactivó) · Cancelada (Opción B / liberada)
var _WOS_RES_SHEET  = 'RESERVAS_EN_CAMINO';
var _WOS_RES        = { FECHA:0, PEDIDO:1, RESELLER:2, SKU:3, CAS:4, AIR:5, ETA:6, CANTIDAD:7, ESTADO:8 };
var _WOS_RES_HEADER = ['FECHA','PEDIDO','RESELLER','SKU','CAS','N_AIR','ETA','CANTIDAD','ESTADO'];

function _wosResSheet() {
  var ss = SpreadsheetApp.openById(MASTER_SS_ID);
  var h  = ss.getSheetByName(_WOS_RES_SHEET);
  if (!h) {
    h = ss.insertSheet(_WOS_RES_SHEET);
    h.appendRow(_WOS_RES_HEADER);
    h.setFrozenRows(1);
  }
  return h;
}

var _WOS_RES_CACHE = 'wos_reservas_activas_v1';
function _wosInvalidarReservasCache() {
  try { CacheService.getScriptCache().remove(_WOS_RES_CACHE); } catch(e) {}
}

// Lee reservas activas → { byCasSku:{SKU:{CAS:qty}}, byPedidoSku:{PEDIDO:{SKU:qty}}, total }
// Cacheada 60s (la invalidan reservar/cerrar). bypass=true fuerza lectura fresca (al asignar).
function _wosReservasActivas(bypass) {
  if (!bypass) {
    try {
      var c0 = CacheService.getScriptCache().get(_WOS_RES_CACHE);
      if (c0) return JSON.parse(c0);
    } catch(eC) {}
  }
  var out = { byCasSku: {}, byPedidoSku: {}, total: 0 };
  try {
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (h) {
      var d = h.getDataRange().getValues();
      var R = _WOS_RES;
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][R.ESTADO] || '').trim() !== 'Activa') continue;
        var sku = String(d[i][R.SKU]    || '').trim().toUpperCase();
        var cas = String(d[i][R.CAS]    || '').trim().toUpperCase();
        var ped = String(d[i][R.PEDIDO] || '').trim();
        var q   = Number(d[i][R.CANTIDAD]) || 0;
        if (!sku || q <= 0) continue;
        if (!out.byCasSku[sku])    out.byCasSku[sku]    = {};
        if (!out.byPedidoSku[ped]) out.byPedidoSku[ped] = {};
        out.byCasSku[sku][cas]    = (out.byCasSku[sku][cas]    || 0) + q;
        out.byPedidoSku[ped][sku] = (out.byPedidoSku[ped][sku] || 0) + q;
        out.total += q;
      }
    }
  } catch(e) { Logger.log('_wosReservasActivas: ' + e); }
  if (!bypass) {
    try {
      var pl = JSON.stringify(out);
      if (pl.length < 90000) CacheService.getScriptCache().put(_WOS_RES_CACHE, pl, 60);
    } catch(eP) {}
  }
  return out;
}

// Reserva, para un pedido que aceptó esperar, las unidades pendientes (E−F−Z de sus líneas
// en Backorder) contra los lotes DJI disponibles más próximos. Idempotente: no re-reserva lo ya
// reservado para ese pedido. Debe llamarse DESPUÉS de que las líneas queden en Backorder.
function _wosReservarEnCamino(numero, reseller) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(eL) { Logger.log('_wosReservarEnCamino lock: ' + eL); }
  try {
    numero = String(numero || '').trim();
    if (!numero) return { ok: false };
    _wosInvalidarReservasCache();                 // asignar siempre con datos frescos (evita doble reserva)
    var ec    = WOS_getEnCaminoMap();             // batchesDisp ya descuenta reservas activas
    var ecMap = (ec && ec.ok) ? ec.map : {};
    var resv  = _wosReservasActivas(true);
    var yaRes = resv.byPedidoSku[numero] || {};

    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false };
    var datos = hoja.getDataRange().getValues();
    var needBySku = {};
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      if (String(datos[i][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
      var sku = String(datos[i][COL.SKU] || '').trim().toUpperCase();
      if (!sku) continue;
      var pend = (Number(datos[i][COL.CANT_SOL]) || 0) - (Number(datos[i][COL.CANT_DESP]) || 0) - (Number(datos[i][COL.CANT_CANCEL]) || 0);
      if (pend > 0) needBySku[sku] = (needBySku[sku] || 0) + pend;
    }

    var nuevas = [], ahora = new Date(), totalRes = 0, etaProx = '', etaProxDt = null;
    for (var s in needBySku) {
      var need = needBySku[s] - (yaRes[s] || 0);   // descuenta lo ya reservado para este pedido
      if (need <= 0) continue;
      var em = ecMap[s];
      if (!em || !em.batchesDisp) continue;         // sin lotes disponibles → no se reserva (queda "a confirmar")
      for (var b = 0; b < em.batchesDisp.length && need > 0; b++) {
        var bt   = em.batchesDisp[b];
        var take = Math.min(need, bt.qty);
        if (take <= 0) continue;
        nuevas.push([ahora, numero, reseller || '', s, bt.cas, bt.air || '', bt.eta || '', take, 'Activa']);
        need -= take; totalRes += take;
        var _dt = _wosEtaToDate(bt.eta);
        if (_dt && (!etaProxDt || _dt < etaProxDt)) { etaProxDt = _dt; etaProx = bt.eta; }
      }
    }
    if (nuevas.length) {
      var h = _wosResSheet();
      h.getRange(h.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
      SpreadsheetApp.flush();
      _wosInvalidarReservasCache();
    }
    return { ok: true, reservas: nuevas.length, cantidad: totalRes, etaProx: etaProx };
  } catch(e) { Logger.log('_wosReservarEnCamino: ' + e); return { ok: false, error: e.toString() }; }
  finally { try { lock.releaseLock(); } catch(eR) {} }
}

// Cambia el estado de las reservas ACTIVAS de un pedido (opcionalmente filtrando por SKU).
// estadoNuevo: 'Cancelada' (Opción B / liberar) | 'Cumplida' (llegó y se reactivó).
function _wosCerrarReservas(numero, skuOpt, estadoNuevo) {
  try {
    numero = String(numero || '').trim();
    if (!numero) return { ok: false };
    var skuF = skuOpt ? String(skuOpt).trim().toUpperCase() : '';
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (!h) return { ok: true, cambiadas: 0 };
    var d = h.getDataRange().getValues();
    var R = _WOS_RES, chg = 0;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][R.ESTADO] || '').trim() !== 'Activa') continue;
      if (String(d[i][R.PEDIDO] || '').trim() !== numero) continue;
      if (skuF && String(d[i][R.SKU] || '').trim().toUpperCase() !== skuF) continue;
      h.getRange(i + 1, R.ESTADO + 1).setValue(estadoNuevo);
      chg++;
    }
    if (chg) { SpreadsheetApp.flush(); _wosInvalidarReservasCache(); }
    return { ok: true, cambiadas: chg };
  } catch(e) { Logger.log('_wosCerrarReservas: ' + e); return { ok: false, error: e.toString() }; }
}

// Estados de línea "cerrados" (ya no se le debe nada al reseller) — mismo criterio que el
// overlay Despacho parcial del front (DP_ESTADOS_CERRADOS). Cualquier otro estado con pend>0
// es deuda viva y justifica mantener una reserva de unidades en camino.
var _WOS_DP_CERRADOS = { 'Entregado_Cerrado':1, 'Entregado_Confirmado':1, 'Listo_Retiro':1, 'Cancelado':1 };

// Libera/cierra las reservas ACTIVAS cuyo pedido ya NO debe unidades de ese SKU (se despachó,
// se canceló, o el reseller nunca volvió). Evita que retengan unidades fantasma. Cuenta como
// deuda viva cualquier línea con pendiente>0 en estado no cerrado (no solo Backorder: el bloqueo
// global del overlay Despacho parcial reserva también para líneas En_Espera/Confirmado/Parcial).
// Idempotente y segura de correr seguido (la llama el detector cada 10 min).
// Cumplida si ese SKU tuvo despacho en el pedido; si no, Cancelada.
function WOS_reconciliarReservas() {
  try {
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (!h) return { ok: true, revisadas: 0, cerradas: 0 };
    var d = h.getDataRange().getValues();
    var R = _WOS_RES;

    // Necesidad viva por (pedido, sku): pendiente en Backorder + si hubo despacho de ese SKU
    var need  = {};
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    for (var hh = 0; hh < hojas.length; hh++) {
      var pd = hojas[hh].getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        var num = String(pd[i][COL.NUMERO] || '').trim();
        var sku = String(pd[i][COL.SKU]    || '').trim().toUpperCase();
        if (!num || !sku) continue;
        var est  = String(pd[i][COL.ESTADO] || '').trim();
        var pend = _WOS_DP_CERRADOS[est] ? 0
          : Math.max(0, (Number(pd[i][COL.CANT_SOL]) || 0) - (Number(pd[i][COL.CANT_DESP]) || 0) - (Number(pd[i][COL.CANT_CANCEL]) || 0));
        if (!need[num]) need[num] = {};
        if (!need[num][sku]) need[num][sku] = { pend: 0, desp: 0 };
        need[num][sku].pend += pend;
        need[num][sku].desp += Number(pd[i][COL.CANT_DESP]) || 0;
      }
    }

    var cerradas = 0, revisadas = 0;
    for (var r = 1; r < d.length; r++) {
      if (String(d[r][R.ESTADO] || '').trim() !== 'Activa') continue;
      revisadas++;
      var pnum = String(d[r][R.PEDIDO] || '').trim();
      var psku = String(d[r][R.SKU]    || '').trim().toUpperCase();
      var info = (need[pnum] && need[pnum][psku]) ? need[pnum][psku] : { pend: 0, desp: 0 };
      if (info.pend > 0) continue;   // sigue necesitando ese SKU en backorder → mantener la reserva
      h.getRange(r + 1, R.ESTADO + 1).setValue(info.desp > 0 ? 'Cumplida' : 'Cancelada');
      cerradas++;
    }
    if (cerradas) { SpreadsheetApp.flush(); _wosInvalidarReservasCache(); }
    if (cerradas) Logger.log('WOS_reconciliarReservas: ' + cerradas + ' cerrada(s) de ' + revisadas + ' activas');
    return { ok: true, revisadas: revisadas, cerradas: cerradas };
  } catch(e) { Logger.log('WOS_reconciliarReservas: ' + e); return { ok: false, error: e.toString() }; }
}

// Datos para el overlay "Despacho parcial": lotes DJI en camino disponibles por SKU (netos de
// reservas activas) + reservas activas con su ETA, para proyectar con qué lote y qué fecha se
// cumple cada línea pendiente y mostrar lo ya bloqueado 🔒.
function WOS_despachoParcialData() {
  try {
    var ec = WOS_getEnCaminoMap();
    if (!ec || !ec.ok) return { ok: false, error: (ec && ec.error) || 'en camino' };
    var bySku = {};
    for (var s in ec.map) {
      var em = ec.map[s];
      bySku[s] = { batchesDisp: em.batchesDisp || [], disponible: em.disponible || 0, reservado: em.reservado || 0 };
    }
    var reservas = [];
    var h = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(_WOS_RES_SHEET);
    if (h) {
      var d = h.getDataRange().getValues();
      var R = _WOS_RES;
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][R.ESTADO] || '').trim() !== 'Activa') continue;
        var q = Number(d[i][R.CANTIDAD]) || 0;
        if (q <= 0) continue;
        reservas.push({
          pedido: String(d[i][R.PEDIDO] || '').trim(),
          sku:    String(d[i][R.SKU]    || '').trim().toUpperCase(),
          eta:    _wosEtaFmt(d[i][R.ETA]),
          qty:    q,
          cas:    String(d[i][R.CAS] || '').trim(),
          air:    String(d[i][R.AIR] || '').trim()
        });
      }
    }
    return { ok: true, bySku: bySku, reservas: reservas };
  } catch(e) { Logger.log('WOS_despachoParcialData: ' + e); return { ok: false, error: e.toString() }; }
}

// Bloquea (reserva) las unidades DJI en camino para TODOS los pedidos con deuda, en orden FIFO
// global (pedido más viejo primero) — la misma asignación que muestra el overlay Despacho parcial.
// Solo reserva la parte NO cubierta por stock actual (la "falta"); lo cubierto por depósito no
// necesita reserva. Idempotente: descuenta lo ya reservado por (pedido, SKU) antes de asignar,
// así el botón se puede tocar las veces que haga falta sin duplicar.
function WOS_bloquearEnCaminoParciales() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch(eL) { Logger.log('WOS_bloquearEnCaminoParciales lock: ' + eL); }
  try {
    _wosInvalidarReservasCache();                 // asignar siempre con datos frescos
    var ec = WOS_getEnCaminoMap();                // batchesDisp ya descuenta reservas activas
    if (!ec || !ec.ok) return { ok: false, error: (ec && ec.error) || 'en camino' };
    var ecMap    = ec.map || {};
    var stockMap = ec.stockMap || {};
    var resv     = _wosReservasActivas(true);

    // Líneas con deuda de ambas hojas (pend>0, estado no cerrado), FIFO por fecha de pedido
    var lineas = [];
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    for (var hh = 0; hh < hojas.length; hh++) {
      var pd = hojas[hh].getDataRange().getValues();
      for (var i = 1; i < pd.length; i++) {
        var num = String(pd[i][COL.NUMERO] || '').trim();
        var sku = String(pd[i][COL.SKU]    || '').trim().toUpperCase();
        if (!num || !sku) continue;
        if (_WOS_DP_CERRADOS[String(pd[i][COL.ESTADO] || '').trim()]) continue;
        var pend = (Number(pd[i][COL.CANT_SOL]) || 0) - (Number(pd[i][COL.CANT_DESP]) || 0) - (Number(pd[i][COL.CANT_CANCEL]) || 0);
        if (pend <= 0) continue;
        var fR = pd[i][COL.FECHA];
        lineas.push({ num: num, res: String(pd[i][COL.RESELLER] || '').trim(), sku: sku,
                      pend: pend, f: (fR instanceof Date) ? fR.getTime() : 0 });
      }
    }
    lineas.sort(function(a, b) { return a.f - b.f; });

    // 1º el stock del depósito cubre lo que puede (sin reserva), 2º la falta toma lotes en camino
    var pool = {};
    var nuevas = [], ahora = new Date(), totalRes = 0, pedSet = {};
    for (var l = 0; l < lineas.length; l++) {
      var ln = lineas[l];
      if (pool[ln.sku] === undefined) pool[ln.sku] = Math.max(0, Number(stockMap[ln.sku]) || 0);
      var deStock = Math.min(ln.pend, pool[ln.sku]);
      pool[ln.sku] -= deStock;
      var falta = ln.pend - deStock;
      if (falta <= 0) continue;
      var yaPed = resv.byPedidoSku[ln.num];
      if (yaPed && yaPed[ln.sku] > 0) {          // ya bloqueado para este pedido → no duplicar
        var usa = Math.min(yaPed[ln.sku], falta);
        yaPed[ln.sku] -= usa;
        falta -= usa;
      }
      if (falta <= 0) continue;
      var em = ecMap[ln.sku];
      if (!em || !em.batchesDisp) continue;       // sin lotes en camino → queda "a confirmar"
      for (var b = 0; b < em.batchesDisp.length && falta > 0; b++) {
        var bt = em.batchesDisp[b];
        if (bt.qty <= 0) continue;
        var take = Math.min(falta, bt.qty);
        bt.qty -= take;                           // consumir el lote para las líneas siguientes
        nuevas.push([ahora, ln.num, ln.res, ln.sku, bt.cas, bt.air || '', bt.eta || '', take, 'Activa']);
        falta -= take; totalRes += take; pedSet[ln.num] = true;
      }
    }
    if (nuevas.length) {
      var h = _wosResSheet();
      h.getRange(h.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
      SpreadsheetApp.flush();
      _wosInvalidarReservasCache();
    }
    var nPed = 0; for (var k in pedSet) nPed++;
    return { ok: true, reservas: nuevas.length, cantidad: totalRes, pedidos: nPed };
  } catch(e) { Logger.log('WOS_bloquearEnCaminoParciales: ' + e); return { ok: false, error: e.toString() }; }
  finally { try { lock.releaseLock(); } catch(eR) {} }
}

// Carga las ubicaciones WMS de un conjunto de SKUs en una sola lectura.
// Devuelve { ok, map: { SKU: [{ubicacion, cantidad}] } } con locs ordenadas desc por cantidad.
function WOS_cargarUbicacionesPedido(skus) {
  try {
    var hojaUbic = SpreadsheetApp.openById(CARMEN_SS_ID).getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: true, map: {} };
    var d   = hojaUbic.getDataRange().getValues();
    var set = {};
    for (var s = 0; s < skus.length; s++) set[String(skus[s]).trim().toUpperCase()] = true;
    var map = {};
    for (var i = 1; i < d.length; i++) {
      var sku  = String(d[i][0] || '').trim().toUpperCase();
      var ubic = String(d[i][1] || '').trim();
      var cant = parseFloat(d[i][2]) || 0;
      if (!sku || !ubic || !set[sku]) continue;
      if (!map[sku]) map[sku] = [];
      map[sku].push({ ubicacion: ubic, cantidad: cant });
    }
    // ASC por cantidad: primero los bins con menos stock para vaciarlos antes
    for (var k in map) map[k].sort(function(a, b) { return a.cantidad - b.cantidad; });
    return { ok: true, map: map };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// Lista principal desde CARMEN (hoja STOCK: A=SKU, B=nombre, C=stock).
// Metadatos adicionales (min, ubicacion, modelos) desde STOCK_REPUESTOS en MASTER.
// q: filtro de búsqueda (SKU o descripción), vacío = todos.
// Fotos del catálogo unificado (hoja TODO, misma spreadsheet que usa Portal Reseller
// como LISTA_PRECIOS_SS_ID y Stock Manager para cargar las fotos faltantes — col B =
// código corto, col H = link de Drive). Cache 5 min, mismo criterio que el stockMap
// de Carmen en WOS_getEnCaminoMap — el catálogo de fotos no cambia a cada minuto.
function _wosFotoMapCatalogo() {
  var map = {};
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('wos_fotos_catalogo_v1');
    if (cached) return JSON.parse(cached);
    var hoja = SpreadsheetApp.openById(CATALOGO_REPUESTOS_ID).getSheetByName('TODO');
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var cod  = String(d[i][1] || '').trim().toUpperCase(); // col B = código corto
        var foto = String(d[i][7] || '').trim();               // col H = imagen
        if (cod && foto) map[cod] = foto;
      }
    }
    try { cache.put('wos_fotos_catalogo_v1', JSON.stringify(map), 300); } catch(eCp) {}
  } catch(e) { Logger.log('_wosFotoMapCatalogo: ' + e); }
  return map;
}

function WOS_cargarStock(q) {
  try {
    // Lista principal: CARMEN hoja STOCK (A=SKU, B=nombre, C=stock actual)
    var carmenSS    = SpreadsheetApp.openById(CARMEN_SS_ID);
    var hojaCarmen  = carmenSS.getSheetByName('STOCK');
    if (!hojaCarmen) return { ok: false, error: 'Hoja STOCK no encontrada en Carmen.' };

    // Abrir MASTER una sola vez para repMap + enCamino
    var master = SpreadsheetApp.openById(MASTER_SS_ID);

    // Ubicaciones WMS desde Carmen UBICACIONES tab: SKU → [{ubicacion, cantidad}]
    var ubicMap = {};
    try {
      var hojaUbic = carmenSS.getSheetByName(CARMEN_UBICACIONES_TAB);
      if (hojaUbic) {
        var dUbic = hojaUbic.getDataRange().getValues();
        for (var u = 1; u < dUbic.length; u++) {
          var uSku  = String(dUbic[u][0] || '').trim().toUpperCase();
          var uUbic = String(dUbic[u][1] || '').trim();
          var uCant = parseFloat(dUbic[u][2]) || 0;
          if (!uSku || !uUbic) continue;
          if (!ubicMap[uSku]) ubicMap[uSku] = [];
          ubicMap[uSku].push({ ubicacion: uUbic, cantidad: uCant });
        }
      }
    } catch(eUbic) { Logger.log('WOS_cargarStock ubicMap: ' + eUbic); }

    // Metadatos extra desde STOCK_REPUESTOS: min(D), categoria(E), ubicacion(F), modelos(G)
    var repMap = {};
    try {
      var repData = master.getSheetByName('STOCK_REPUESTOS').getDataRange().getValues();
      for (var r = 1; r < repData.length; r++) {
        var rCod = String(repData[r][0] || '').trim().toUpperCase();
        if (!rCod) continue;
        repMap[rCod] = {
          minimo:    parseInt(repData[r][3]) || 0,
          categoria: String(repData[r][4] || '').trim(),
          ubicacion: String(repData[r][5] || '').trim(),
          modelos:   String(repData[r][6] || '').trim()
        };
      }
    } catch(eR) { Logger.log('WOS_cargarStock repMap: ' + eR); }

    // Planificación: clasificación (col N) y mínimo (col O × 4) desde PLANILLA DE PLANIFICACION
    var planifMap = {};
    try {
      var cache = CacheService.getScriptCache();
      var cachedP = cache.get('wos_planif_map_v6');
      if (cachedP) {
        planifMap = JSON.parse(cachedP);
      } else {
        var ssPlanif = SpreadsheetApp.openById(PLANIF_SS_ID);
        var hojaPlanif = null;
        var hojas = ssPlanif.getSheets();
        for (var h = 0; h < hojas.length; h++) {
          if (hojas[h].getName().trim().toUpperCase() === 'PLANILLA DE PLANIFICACION') {
            hojaPlanif = hojas[h]; break;
          }
        }
        if (!hojaPlanif) throw new Error('Hoja planif no encontrada');
        var planifData = hojaPlanif.getDataRange().getValues();
        // Filas 1-2 son cabeceras (advertencia + nombres de col); datos desde índice 3
        for (var p = 3; p < planifData.length; p++) {
          var pCod   = String(planifData[p][0]  || '').trim().toUpperCase(); // col A: PN Corto
          var pClase = String(planifData[p][13] || '').trim().toUpperCase(); // col N: Clasificacion
          var pMin   = parseFloat(planifData[p][14]) || 0;                   // col O: Stock mínimo
          if (pCod) planifMap[pCod] = { clase: pClase, minimo: Math.round(pMin) };
        }
        if (Object.keys(planifMap).length > 0) {
          try { cache.put('wos_planif_map_v6', JSON.stringify(planifMap), 300); } catch(eCp) {}
        }
      }
    } catch(ePl) { Logger.log('WOS_cargarStock planifMap ERROR: ' + ePl); }

    // Unidades en camino — reutiliza WOS_getEnCaminoMap para evitar duplicar lógica
    var enCaminoMap = {};
    try {
      var ecRes = WOS_getEnCaminoMap();
      if (ecRes.ok) enCaminoMap = ecRes.map;
    } catch(eEC) { Logger.log('WOS_cargarStock enCamino: ' + eEC); }

    // Fotos del catálogo (hoja TODO) — para mostrar/ver la imagen de los ítems que ya tienen
    var fotoMap = _wosFotoMapCatalogo();

    var datos  = hojaCarmen.getDataRange().getValues();
    var filtro = String(q || '').trim().toLowerCase();
    var out    = [];

    for (var i = 1; i < datos.length; i++) {
      var cod  = String(datos[i][0] || '').trim();
      var desc = String(datos[i][1] || '').trim();
      if (!cod) continue;
      if (filtro && cod.toLowerCase().indexOf(filtro) < 0 && desc.toLowerCase().indexOf(filtro) < 0) continue;

      var codKey  = cod.toUpperCase();
      var actual  = parseInt(datos[i][2]) || 0;
      var meta    = repMap[codKey] || { minimo: 0, categoria: '', ubicacion: '', modelos: '' };
      var planif  = planifMap[codKey] || { clase: '', minimo: 0 };
      var minimo  = planif.minimo > 0 ? planif.minimo : meta.minimo;
      var estado  = actual <= 0 ? 'CRITICO' : (actual <= minimo ? 'BAJO' : 'OK');
      var ecData  = enCaminoMap[codKey];
      var ecTotal = ecData ? ecData.total : 0;
      var ecOcs   = [];
      if (ecData && ecData.batches) {
        for (var _bi = 0; _bi < ecData.batches.length; _bi++) {
          var _bt = ecData.batches[_bi];
          ecOcs.push(_bt.cas + ' (' + _bt.qty + 'u.)' + (_bt.eta ? ' · llega ~' + _bt.eta : ''));
        }
      } else if (ecData) {
        for (var ocId in ecData.ocs) ecOcs.push(ocId + ' (' + ecData.ocs[ocId] + 'u.)');
      }

      out.push({
        codigo:      cod,
        descripcion: desc,
        stockActual: actual,
        stockMinimo: minimo,
        clase:       planif.clase,
        estado:      estado,
        categoria:   meta.categoria,
        ubicacion:   (function() {
          var arr = ubicMap[codKey];
          if (arr && arr.length) {
            return arr.map(function(u) { return u.ubicacion + ' (' + u.cantidad + 'u.)'; }).join(' · ');
          }
          return meta.ubicacion; // fallback: ítem aún sin mapear en WMS
        })(),
        modelos:     String(datos[i][4] || '').trim(),
        enCamino:    ecTotal,
        enCaminoOcs: ecOcs,
        enCaminoETA: ecData ? (ecData.etaMin || '') : '',
        enCaminoReservado: ecData ? (ecData.reservado || 0) : 0,
        foto:        fotoMap[codKey] || ''
      });
    }

    // CRÍTICO → BAJO → OK; dentro de cada estado por código alfanuméricamente
    var orden = { 'CRITICO': 0, 'BAJO': 1, 'OK': 2 };
    out.sort(function(a, b) {
      var od = orden[a.estado] - orden[b.estado];
      return od !== 0 ? od : (a.codigo < b.codigo ? -1 : 1);
    });

    return { ok: true, items: out };
  } catch(e) {
    Logger.log('WOS_cargarStock ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Despacho en batch (múltiples pedidos, mismos bultos/tracking) ─
// batchItems: [{numero, despachos:[{row,cantDesp}]}]
// El costo se divide en partes iguales entre los pedidos del batch.
function WOS_despacharBatch(batchItems, transportista, bultos, costoEnvio, operario, reqToken) {
  try {
    if (!batchItems || !batchItems.length) return { ok: false, error: 'Sin pedidos en el batch.' };
    operario = String(operario || '');
    var costoPorPedido = batchItems.length > 1
      ? Math.round((costoEnvio || 0) / batchItems.length * 100) / 100
      : (Number(costoEnvio) || 0);
    var resultados = [];
    for (var i = 0; i < batchItems.length; i++) {
      var item = batchItems[i];
      try {
        // token de idempotencia por pedido (deriva del token del batch) → un reintento no duplica
        var itemToken = reqToken ? (reqToken + '::' + item.numero) : '';
        var res = WOS_despacharCompleto(item.numero, item.despachos, transportista, bultos, costoPorPedido, operario, itemToken);
        resultados.push({ numero: item.numero, ok: res.ok, error: res.error || '' });
      } catch(eI) {
        resultados.push({ numero: item.numero, ok: false, error: eI.toString() });
      }
    }
    var errores = [];
    for (var j = 0; j < resultados.length; j++) { if (!resultados[j].ok) errores.push(resultados[j]); }
    Logger.log('WOS_despacharBatch: ' + resultados.length + ' pedidos, ' + errores.length + ' errores');
    return { ok: true, resultados: resultados, errores: errores };
  } catch(e) {
    Logger.log('WOS_despacharBatch ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Resumen de envíos por reseller y mes ──────────────────────
// mesAnio: "YYYY-MM", e.g. "2025-06"
// Agrupa por NOTA_ENTREGA (cada evento de despacho = una fila).
function WOS_getResumenEnvios(reseller, mesAnio) {
  try {
    if (!reseller || !mesAnio) return { ok: false, error: 'Faltan parámetros.' };
    var partes = mesAnio.split('-');
    var anio   = parseInt(partes[0]) || 0;
    var mes    = parseInt(partes[1]) || 0;
    if (!anio || !mes) return { ok: false, error: 'Formato de mes inválido (YYYY-MM).' };

    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    var tz    = Session.getScriptTimeZone();

    var notasVistas = {};
    var envios      = [];

    for (var h = 0; h < hojas.length; h++) {
    var datos = hojas[h].getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var r      = datos[i];
      var rRes   = String(r[COL.RESELLER]     || '').trim();
      var rNota  = String(r[COL.NOTA_ENTREGA] || '').trim();
      var rFecRaw = r[COL.FECHA_DESPACHO];
      if (!rNota || !rFecRaw) continue;
      if (rRes.toLowerCase() !== reseller.toLowerCase()) continue;

      var fecDate = (rFecRaw instanceof Date) ? rFecRaw : new Date(rFecRaw);
      if (isNaN(fecDate.getTime())) continue;
      if (fecDate.getFullYear() !== anio || (fecDate.getMonth() + 1) !== mes) continue;
      if (notasVistas[rNota]) continue;
      notasVistas[rNota] = true;

      var fecStr = Utilities.formatDate(fecDate, tz, 'dd/MM/yyyy');
      envios.push({
        fecha:        fecStr,
        numero:       String(r[COL.NUMERO]          || '').trim(),
        notaEntrega:  rNota,
        transportista: String(r[COL.TRANSPORTISTA_DESP] || '').trim(),
        tracking:     String(r[COL.TRACKING]         || '').trim(),
        costoEnvio:   Number(r[COL.COSTO_ENVIO])    || 0,
        pesoEnvio:    Number(r[COL.PESO_ENVIO])      || 0
      });
    }
    } // fin loop hojas

    var totalCosto = 0;
    var totalPeso  = 0;
    for (var e = 0; e < envios.length; e++) {
      totalCosto += envios[e].costoEnvio;
      totalPeso  += envios[e].pesoEnvio;
    }

    var mesLabel = Utilities.formatDate(new Date(anio, mes - 1, 1), tz, 'MMMM yyyy');
    return {
      ok:         true,
      reseller:   reseller,
      mesAnio:    mesAnio,
      mesLabel:   mesLabel,
      envios:     envios,
      totalCosto: totalCosto,
      totalPeso:  Math.round(totalPeso * 100) / 100
    };
  } catch(e) {
    Logger.log('WOS_getResumenEnvios ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Envía el resumen de envíos por email al reseller ──────────
function WOS_enviarResumenEnvios(reseller, mesAnio, reqToken) {
 return _wosLockIdempot(reqToken, function() {
  try {
    var res = WOS_getResumenEnvios(reseller, mesAnio);
    if (!res.ok) return res;
    if (!res.envios.length) return { ok: false, error: 'No hay envíos para ese reseller y mes.' };

    var email = _wosGetEmailReseller(reseller);
    if (!email) return { ok: false, error: 'Email no encontrado para: ' + reseller };

    var tbodyRows = '';
    for (var i = 0; i < res.envios.length; i++) {
      var ev  = res.envios[i];
      var bg  = i % 2 === 0 ? '#ffffff' : '#f7f8fa';
      tbodyRows +=
        "<tr style='background:" + bg + "'>" +
        "<td style='padding:7px 10px;font-size:12px;color:#555'>"                                    + ev.fecha                                                        + "</td>" +
        "<td style='padding:7px 10px;font-size:12px;font-family:monospace;color:#00a3e0'>"            + ev.numero                                                       + "</td>" +
        "<td style='padding:7px 10px;font-size:12px;font-family:monospace;color:#1a1a2e'>"            + ev.notaEntrega                                                  + "</td>" +
        "<td style='padding:7px 10px;font-size:12px;color:#555'>"                                    + (ev.transportista || '—')                                        + "</td>" +
        "<td style='padding:7px 10px;font-size:12px;text-align:right;color:#555'>"                   + (ev.pesoEnvio  > 0 ? ev.pesoEnvio  + ' kg' : '—')               + "</td>" +
        "<td style='padding:7px 10px;font-size:13px;font-weight:700;text-align:right;color:#1a1a2e'>" + (ev.costoEnvio > 0 ? '$ ' + ev.costoEnvio.toFixed(2) : '—')    + "</td>" +
        "</tr>";
    }

    var tablaHtml =
      "<table width='100%' cellpadding='0' cellspacing='0' style='border-collapse:collapse;border:1px solid #dde3ea;border-radius:8px;overflow:hidden'>" +
      "<thead><tr style='background:#f0f5fa'>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888;letter-spacing:.05em'>Fecha</th>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888'>Pedido</th>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888'>Nota de Entrega</th>" +
        "<th style='padding:9px 10px;text-align:left;font-size:11px;font-weight:700;color:#888'>Transportista</th>" +
        "<th style='padding:9px 10px;text-align:right;font-size:11px;font-weight:700;color:#888'>Peso</th>" +
        "<th style='padding:9px 10px;text-align:right;font-size:11px;font-weight:700;color:#888'>Costo envío</th>" +
      "</tr></thead><tbody>" + tbodyRows + "</tbody></table>";

    var totalHtml =
      "<div style='text-align:right;margin-top:12px;padding:12px 16px;background:#e8f5fc;border-radius:8px'>" +
        "<span style='font-size:12px;color:#444'>Total envíos " + res.mesLabel + ": </span>" +
        "<strong style='font-size:16px;color:#00a3e0'>$ " + res.totalCosto.toFixed(2) + "</strong>" +
        (res.totalPeso > 0 ? "<span style='font-size:11px;color:#888;margin-left:14px'>Peso total: " + res.totalPeso + " kg</span>" : '') +
      "</div>";

    var htmlBody = _wosPortalHead('Resumen de Envíos — ' + res.mesLabel) +
      "<p style='font-size:14px;color:#666;margin:0 0 6px;line-height:1.5'>Hola <strong>" + reseller + "</strong>:</p>" +
      "<p style='font-size:13px;color:#555;margin:0 0 18px'>A continuación el detalle de los <strong>" + res.envios.length + " env" + (res.envios.length === 1 ? "ío" : "íos") + "</strong> realizados durante <strong>" + res.mesLabel + "</strong>.</p>" +
      tablaHtml + totalHtml +
      _wosPortalFoot('Resumen de Envíos · ' + reseller + ' · ' + res.mesLabel + '.');

    var plainBody =
      'Hola ' + reseller + ',\n\n' +
      'Resumen de envíos - ' + res.mesLabel + '\n' +
      '========================================\n';
    for (var j = 0; j < res.envios.length; j++) {
      var ev2 = res.envios[j];
      plainBody += ev2.fecha + ' | ' + ev2.numero + ' | ' + ev2.notaEntrega +
        ' | ' + (ev2.transportista || 'N/E') +
        (ev2.pesoEnvio  > 0 ? ' | ' + ev2.pesoEnvio  + ' kg'          : '') +
        (ev2.costoEnvio > 0 ? ' | $ ' + ev2.costoEnvio.toFixed(2) : '') + '\n';
    }
    plainBody += '========================================\n' +
      'Total: $ ' + res.totalCosto.toFixed(2) + (res.totalPeso > 0 ? ' · Peso total: ' + res.totalPeso + ' kg' : '');

    GmailApp.sendEmail(email, 'Resumen de envíos — ' + reseller + ' — ' + res.mesLabel, plainBody, {
      htmlBody: htmlBody,
      name:     'BIDCOMAGRO · Portal Resellers',
      replyTo:  _wosConfig().emailSoporte,
      cc:       _wosConfig().emailFact
    });

    Logger.log('WOS_enviarResumenEnvios OK: ' + reseller + ' ' + mesAnio + ' (' + res.envios.length + ' envíos)');
    return { ok: true, enviados: res.envios.length };
  } catch(e) {
    Logger.log('WOS_enviarResumenEnvios ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
 });
}

// Revierte un pedido despachado por error de vuelta a "Preparado"
function WOS_revertirAPreparado(numero, operario) {
  try {
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    var ahora = new Date();
    var reseller = '';
    var filas = [];

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      reseller = reseller || String(datos[i][COL.RESELLER] || '').trim();
      var estAct = String(datos[i][COL.ESTADO] || '').trim();
      // Solo revertir ítems que estaban Entregado_Cerrado o Backorder por este despacho
      if (estAct === EST.ENTREGADO || estAct === EST.BACKORDER || estAct === EST.LISTO_RETIRO) {
        filas.push(i + 1);
      }
    }

    if (!filas.length) return { ok: false, error: 'No hay ítems revertibles en este pedido (ya puede haber sido confirmado por el reseller).' };

    for (var f = 0; f < filas.length; f++) {
      var fila = filas[f];
      var r = hoja.getRange(fila, COL.ESTADO + 1);
      r.clearDataValidations();
      r.setValue(EST.PREPARADO);
      // CANT_DESP aislado
      hoja.getRange(fila, COL.CANT_DESP + 1).setValue('');
      // FECHA_DESPACHO (col 15), NOTA_ENTREGA (col 16), TRACKING (col 17) — bloque contiguo
      hoja.getRange(fila, COL.FECHA_DESPACHO + 1, 1, 3).setValues([['', '', '']]);
      // FECHA_ESTADO (col 19), TRANSPORTISTA_DESP (col 20) — bloque contiguo
      hoja.getRange(fila, COL.FECHA_ESTADO + 1, 1, 2).setValues([[ahora, '']]);
      // NE_URL aislado
      hoja.getRange(fila, COL.NE_URL + 1).setValue('');
    }
    SpreadsheetApp.flush();
    _wosLogAccion('Revertido a Preparado', numero, reseller, String(operario || ''), 'Revert manual por error de despacho');
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_revertirAPreparado ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Lee las últimas N entradas del log para un pedido específico
function WOS_getHistorial(numero) {
  try {
    var ss   = SpreadsheetApp.openById(NOTAS_SS_ID);
    var hoja = ss.getSheetByName('WOS_Log');
    if (!hoja) return { ok: true, entradas: [] };
    var datos = hoja.getDataRange().getValues();
    var tz    = Session.getScriptTimeZone();
    var entradas = [];
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][1] || '').trim() !== numero) continue;
      var fechaRaw = datos[i][0];
      entradas.push({
        fecha:    (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, tz, 'dd/MM/yy HH:mm') : String(fechaRaw || ''),
        accion:   String(datos[i][3] || ''),
        operario: String(datos[i][4] || '').replace(/@.*/, ''),
        detalle:  String(datos[i][5] || '')
      });
    }
    entradas.reverse(); // más reciente primero
    return { ok: true, entradas: entradas.slice(0, 20) };
  } catch(e) {
    return { ok: false, entradas: [] };
  }
}

// Actualiza el campo de observaciones de todas las filas del pedido
function WOS_actualizarObs(numero, obs) {
  try {
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      hoja.getRange(i + 1, COL.OBS + 1).setValue(obs || '');
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_actualizarObs: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Polling liviano — compara timestamps de FECHA_ESTADO contra ultimoMs del cliente
function WOS_checkCambios(ultimoMs) {
  try {
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    if (!hojas.length) return { ok: false };
    var nuevos = 0;
    for (var h = 0; h < hojas.length; h++) {
      var datos = hojas[h].getDataRange().getValues();
      for (var i = 1; i < datos.length; i++) {
        if (!String(datos[i][COL.NUMERO] || '').trim()) continue;
        var feRaw    = datos[i][COL.FECHA_ESTADO];
        var fechaRaw = datos[i][COL.FECHA];
        var ms = (feRaw    instanceof Date) ? feRaw.getTime()
               : (fechaRaw instanceof Date) ? fechaRaw.getTime() : 0;
        if (ms > ultimoMs) nuevos++;
      }
    }
    return { ok: true, cambiado: nuevos > 0, nuevos: nuevos };
  } catch(e) {
    return { ok: false };
  }
}

// ── REPORTE BACKORDER AUTOMÁTICO (Lun/Mie/Vie 10hs) ──────────────────────────

function WOS_reporteBackorder() {
  try {
    var masterSS = SpreadsheetApp.openById(MASTER_SS_ID);

    // 1. Destinatarios: Usuarios_Internos col B=email, col F=si
    var hojaU = masterSS.getSheetByName('Usuarios_Internos');
    if (!hojaU) { Logger.log('WOS_reporteBackorder: hoja Usuarios_Internos no encontrada'); return; }
    var usrs = hojaU.getDataRange().getValues();
    var destinatarios = [];
    for (var u = 1; u < usrs.length; u++) {
      var email  = String(usrs[u][1] || '').trim();
      var logInt = String(usrs[u][5] || '').trim().toLowerCase();
      if (email && logInt === 'si') destinatarios.push(email);
    }
    if (!destinatarios.length) { Logger.log('WOS_reporteBackorder: sin destinatarios logística internacional'); return; }

    // 2. Ítems en backorder del WOS, agrupados por SKU — fusiona ambas hojas
    var wosHojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    var backMap = {};
    for (var wh = 0; wh < wosHojas.length; wh++) {
    var wosData = wosHojas[wh].getDataRange().getValues();
    for (var i = 1; i < wosData.length; i++) {
      if (String(wosData[i][COL.ESTADO] || '').trim() !== EST.BACKORDER) continue;
      var sku  = String(wosData[i][COL.SKU]      || '').trim().toUpperCase();
      if (!sku) continue;
      var desc     = String(wosData[i][COL.DESC]     || '').trim();
      var numero   = String(wosData[i][COL.NUMERO]   || '').trim();
      var reseller = String(wosData[i][COL.RESELLER] || '').trim();
      // Necesidad real = E−F−Z: descontar lo cancelado por el reseller para no reportar de más a logística.
      var nec      = (Number(wosData[i][COL.CANT_SOL]  || 0) - Number(wosData[i][COL.CANT_DESP] || 0) - Number(wosData[i][COL.CANT_CANCEL] || 0));
      if (nec <= 0) continue;
      if (!backMap[sku]) backMap[sku] = { desc: desc, nec: 0, pedidos: [] };
      backMap[sku].nec += nec;
      backMap[sku].pedidos.push(numero + ' · ' + reseller + ' (' + nec + 'u)');
    }
    } // fin loop wosHojas

    // 3. Unidades en camino por SKU — misma lógica que WOS_getEnCaminoMap (excluye Borrador y En depósito)
    var casActivos = {};
    var hojaCAS = masterSS.getSheetByName('COMPRAS_DJI');
    if (hojaCAS) {
      var casData = hojaCAS.getDataRange().getValues();
      for (var c = 1; c < casData.length; c++) {
        var casId  = String(casData[c][0] || '').trim();
        var casEst = String(casData[c][2] || '').trim();
        if (casId && casEst !== 'En dep\xf3sito' && casEst.indexOf('Borrador') < 0) casActivos[casId] = true;
      }
    }
    var enCamino = {};
    var enCaminoEta = {}; // SKU → ETA más cercana (string) entre los lotes en camino
    var hojaDetalle = masterSS.getSheetByName('COMPRAS_DETALLE');
    if (hojaDetalle) {
      var detData = hojaDetalle.getDataRange().getValues();
      for (var d = 1; d < detData.length; d++) {
        if (!casActivos[String(detData[d][0] || '').trim()]) continue;
        var dSku  = String(detData[d][1] || '').trim().toUpperCase();
        var dPend = (Number(detData[d][3] || 0) - Number(detData[d][4] || 0));
        if (dSku && dPend > 0) {
          enCamino[dSku] = (enCamino[dSku] || 0) + dPend;
          var dEtaDt = _wosEtaToDate(detData[d][6]);
          if (dEtaDt) {
            var prevDt = enCaminoEta[dSku] ? _wosEtaToDate(enCaminoEta[dSku]) : null;
            if (!prevDt || dEtaDt < prevDt) enCaminoEta[dSku] = _wosEtaFmt(detData[d][6]);
          }
        }
      }
    }

    // 4. Clasificar: sin cobertura vs cubiertos
    var sinCubrir = [];
    var cubiertos = [];
    for (var sku in backMap) {
      var item    = backMap[sku];
      var camino  = enCamino[sku] || 0;
      var gap     = item.nec - camino;
      var entrada = { sku: sku, desc: item.desc, nec: item.nec, camino: camino, eta: enCaminoEta[sku] || '', gap: gap > 0 ? gap : 0, pedidos: item.pedidos };
      if (gap > 0) sinCubrir.push(entrada);
      else         cubiertos.push(entrada);
    }
    sinCubrir.sort(function(a, b) { return b.gap - a.gap; });
    cubiertos.sort(function(a, b) { return b.nec - a.nec; });

    // 5. Demanda perdida — ítems con CANT_CANCEL > 0 en los últimos 90 días, agrupados por SKU
    //    Fallback: filas con estado Cancelado sin CANT_CANCEL (pedidos anteriores al campo)
    var perdidoMap = {};
    var hace90 = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000);
    for (var p = 1; p < wosData.length; p++) {
      var pSku = String(wosData[p][COL.SKU] || '').trim().toUpperCase();
      if (!pSku) continue;
      var pFecha = wosData[p][COL.FECHA_ESTADO] instanceof Date ? wosData[p][COL.FECHA_ESTADO]
                 : (wosData[p][COL.FECHA] instanceof Date ? wosData[p][COL.FECHA] : null);
      if (!pFecha || pFecha < hace90) continue;
      var pCantCancel = Number(wosData[p][COL.CANT_CANCEL] || 0);
      var pEst        = String(wosData[p][COL.ESTADO] || '').trim();
      // Fuente primaria: CANT_CANCEL. Fallback: fila Cancelada sin CANT_CANCEL (datos viejos)
      var pCant = pCantCancel > 0 ? pCantCancel
                : (pEst === EST.CANCELADO ? Number(wosData[p][COL.CANT_SOL] || 0) : 0);
      if (pCant <= 0) continue;
      var pDesc     = String(wosData[p][COL.DESC]     || '').trim();
      var pReseller = String(wosData[p][COL.RESELLER] || '').trim();
      if (!perdidoMap[pSku]) perdidoMap[pSku] = { desc: pDesc, total: 0, resellers: {} };
      perdidoMap[pSku].total += pCant;
      perdidoMap[pSku].resellers[pReseller] = (perdidoMap[pSku].resellers[pReseller] || 0) + pCant;
    }
    var perdidos = [];
    for (var psk in perdidoMap) {
      var pm = perdidoMap[psk];
      var rList = [];
      for (var rn in pm.resellers) rList.push(rn + ' (' + pm.resellers[rn] + 'u)');
      perdidos.push({ sku: psk, desc: pm.desc, total: pm.total, resellers: rList });
    }
    perdidos.sort(function(a, b) { return b.total - a.total; });

    if (!sinCubrir.length && !cubiertos.length && !perdidos.length) {
      Logger.log('WOS_reporteBackorder: sin ítems en backorder ni demanda perdida, no se envía mail');
      return;
    }

    // 6. Enviar email
    var fechaStr = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', "EEEE dd/MM/yyyy 'a las' HH:mm");
    var html     = _wosBackorderEmailHTML(sinCubrir, cubiertos, perdidos, fechaStr);
    var asunto   = 'Backorder WOS — ' + sinCubrir.length + ' ítem' + (sinCubrir.length !== 1 ? 's' : '') + ' sin cobertura DJI';
    GmailApp.sendEmail(destinatarios[0], asunto, '', { htmlBody: html, name: 'WOS · BidcomAgro', cc: destinatarios.slice(1).join(',') });
    Logger.log('WOS_reporteBackorder enviado a: ' + destinatarios.join(', ') + ' | sin cobertura: ' + sinCubrir.length + ', cubiertos: ' + cubiertos.length);
  } catch(e) {
    Logger.log('WOS_reporteBackorder ERROR: ' + e);
  }
}

function _wosBackorderEmailHTML(sinCubrir, cubiertos, perdidos, fechaStr) {
  var rowsRojo = '';
  for (var i = 0; i < sinCubrir.length; i++) {
    var it = sinCubrir[i];
    rowsRojo +=
      '<tr style="border-bottom:1px solid #fecaca">' +
      '<td style="padding:8px 12px;font-family:monospace;font-weight:700;color:#1a56db;white-space:nowrap">' + it.sku + '</td>' +
      '<td style="padding:8px 12px;font-size:12px;color:#111">' + it.desc + '</td>' +
      '<td style="padding:8px 12px;text-align:center;font-weight:700;color:#7f1d1d">' + it.nec + '</td>' +
      '<td style="padding:8px 12px;text-align:center;color:#1e40af">' + it.camino + (it.eta ? '<br><span style="font-size:10px;color:#3b82f6;font-weight:600">llega ~' + it.eta + '</span>' : '') + '</td>' +
      '<td style="padding:8px 12px;text-align:center;font-weight:800;font-size:14px;color:#dc2626;background:#fef2f2">' + it.gap + '</td>' +
      '<td style="padding:8px 12px;font-size:11px;color:#555;line-height:1.6">' + it.pedidos.join('<br>') + '</td>' +
      '</tr>';
  }
  var rowsVerde = '';
  for (var j = 0; j < cubiertos.length; j++) {
    var cov = cubiertos[j];
    rowsVerde +=
      '<tr style="border-bottom:1px solid #d1fae5">' +
      '<td style="padding:6px 12px;font-family:monospace;font-size:11px;color:#374151">' + cov.sku + '</td>' +
      '<td style="padding:6px 12px;font-size:11px;color:#374151">' + cov.desc + '</td>' +
      '<td style="padding:6px 12px;text-align:center;font-size:12px">' + cov.nec + '</td>' +
      '<td style="padding:6px 12px;text-align:center;font-size:12px;color:#166534;font-weight:700">' + cov.camino + (cov.eta ? '<br><span style="font-size:10px;color:#3b82f6;font-weight:600">~' + cov.eta + '</span>' : '') + '</td>' +
      '<td style="padding:6px 12px;text-align:center" colspan="2"><span style="background:#dcfce7;color:#166534;padding:2px 10px;border-radius:99px;font-size:11px;font-weight:700">✓ Cubierto</span></td>' +
      '</tr>';
  }

  var thStyle   = 'padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;';
  var tableHdr  = function(color) {
    return '<thead><tr style="background:' + color + ';border-bottom:2px solid ' + (color === '#fef2f2' ? '#fecaca' : '#bbf7d0') + '">' +
      '<th style="' + thStyle + 'color:' + (color === '#fef2f2' ? '#991b1b' : '#166534') + ';text-align:left">SKU</th>' +
      '<th style="' + thStyle + 'color:' + (color === '#fef2f2' ? '#991b1b' : '#166534') + ';text-align:left">Descripción</th>' +
      '<th style="' + thStyle + 'color:' + (color === '#fef2f2' ? '#991b1b' : '#166534') + ';text-align:center">Necesario</th>' +
      '<th style="' + thStyle + 'color:' + (color === '#fef2f2' ? '#991b1b' : '#166534') + ';text-align:center">En camino</th>' +
      '<th style="' + thStyle + 'color:' + (color === '#fef2f2' ? '#991b1b' : '#166534') + ';text-align:center">Faltante</th>' +
      '<th style="' + thStyle + 'color:' + (color === '#fef2f2' ? '#991b1b' : '#166534') + ';text-align:left">Pedidos</th>' +
      '</tr></thead>';
  };

  var secRojo = sinCubrir.length
    ? '<h3 style="font-size:12px;font-weight:800;color:#b91c1c;margin:0 0 8px;text-transform:uppercase;letter-spacing:.5px">⚠ Sin cobertura — requieren acción</h3>' +
      '<div style="overflow-x:auto;margin-bottom:24px"><table style="width:100%;border-collapse:collapse;font-size:12px">' +
      tableHdr('#fef2f2') + '<tbody>' + rowsRojo + '</tbody></table></div>'
    : '';

  var secVerde = cubiertos.length
    ? '<h3 style="font-size:12px;font-weight:800;color:#166534;margin:0 0 8px;text-transform:uppercase;letter-spacing:.5px">✓ Cubiertos por compras en tránsito</h3>' +
      '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">' +
      tableHdr('#f0fdf4') + '<tbody>' + rowsVerde + '</tbody></table></div>'
    : '';

  var rowsPerd = '';
  for (var k = 0; k < perdidos.length; k++) {
    var pd = perdidos[k];
    rowsPerd +=
      '<tr style="border-bottom:1px solid #e5e7eb">' +
      '<td style="padding:7px 12px;font-family:monospace;font-size:11px;font-weight:700;color:#1a56db;white-space:nowrap">' + pd.sku + '</td>' +
      '<td style="padding:7px 12px;font-size:11px;color:#111">' + pd.desc + '</td>' +
      '<td style="padding:7px 12px;text-align:center;font-weight:700;color:#374151">' + pd.total + '</td>' +
      '<td style="padding:7px 12px;font-size:11px;color:#555;line-height:1.6">' + pd.resellers.join('<br>') + '</td>' +
      '</tr>';
  }
  var secPerdida = perdidos.length
    ? '<hr style="border:none;border-top:1px solid #e2e8f0;margin:24px 0">' +
      '<h3 style="font-size:12px;font-weight:800;color:#78350f;margin:0 0 4px;text-transform:uppercase;letter-spacing:.5px">📦 Demanda perdida — últimos 90 días</h3>' +
      '<p style="font-size:11px;color:#92400e;margin:0 0 10px">Ítems cancelados por resellers que podrías stockear a futuro. Ordenados por volumen total.</p>' +
      '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">' +
      '<thead><tr style="background:#fffbeb;border-bottom:2px solid #fde68a">' +
      '<th style="' + thStyle + 'color:#92400e;text-align:left">SKU</th>' +
      '<th style="' + thStyle + 'color:#92400e;text-align:left">Descripción</th>' +
      '<th style="' + thStyle + 'color:#92400e;text-align:center">Unidades canceladas</th>' +
      '<th style="' + thStyle + 'color:#92400e;text-align:left">Resellers</th>' +
      '</tr></thead><tbody>' + rowsPerd + '</tbody></table></div>'
    : '';

  var banner = sinCubrir.length
    ? '<div style="background:#fef2f2;border-left:4px solid #dc2626;padding:12px 16px;border-radius:0 6px 6px 0;margin-bottom:24px">' +
      '<strong style="color:#b91c1c;font-size:14px">⚠ ' + sinCubrir.length + ' ítem' + (sinCubrir.length !== 1 ? 's' : '') + ' en backorder sin cobertura DJI</strong>' +
      '<div style="font-size:12px;color:#7f1d1d;margin-top:4px">Estos repuestos no tienen unidades suficientes en compras activas (en tránsito). Se requiere gestionar una nueva compra ¡URGENTE!</div></div>'
    : '<div style="background:#f0fdf4;border-left:4px solid #16a34a;padding:12px 16px;border-radius:0 6px 6px 0;margin-bottom:24px">' +
      '<strong style="color:#166534;font-size:14px">✓ Todos los ítems en backorder están cubiertos por compras DJI en tránsito.</strong></div>';

  return '<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,Segoe UI,sans-serif">' +
    '<div style="max-width:720px;margin:24px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,.09)">' +
    '<div style="background:#1e3a8a;padding:20px 28px;display:flex;align-items:center;gap:14px">' +
    '<div style="background:#fff;color:#1e3a8a;font-weight:900;font-size:16px;padding:4px 10px;border-radius:6px;letter-spacing:-0.5px;flex-shrink:0">WOS</div>' +
    '<div><div style="color:#fff;font-size:15px;font-weight:700">Reporte de Backorder</div>' +
    '<div style="color:#93c5fd;font-size:12px">' + fechaStr + '</div></div>' +
    '</div>' +
    '<div style="padding:24px 28px">' + banner + secRojo + secVerde + secPerdida + '</div>' +
    '<div style="background:#f8fafc;border-top:1px solid #e2e8f0;padding:12px 28px;font-size:11px;color:#94a3b8;text-align:center">' +
    'WOS · BidcomAgro · Reporte automático — Lunes, Miércoles y Viernes a las 10 hs' +
    '</div></div></body></html>';
}

// ── Thread ID del hilo de backorder (guardado en WOS_CONFIG de MASTER) ──
function _wosGetBackorderThreadId() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('WOS_CONFIG');
    if (!hoja) return '';
    var data = hoja.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === 'BACKORDER_THREAD_ID') {
        return String(data[i][1] || '').trim();
      }
    }
  } catch(e) { Logger.log('_wosGetBackorderThreadId: ' + e); }
  return '';
}

function _wosSetBackorderThreadId(threadId) {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('WOS_CONFIG');
    if (!hoja) return;
    var data = hoja.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === 'BACKORDER_THREAD_ID') {
        hoja.getRange(i + 1, 2).setValue(threadId);
        _wosConfigCache = null;
        try { CacheService.getScriptCache().remove('wos_config_v1'); } catch(eC) {}
        return;
      }
    }
    hoja.appendRow(['BACKORDER_THREAD_ID', threadId]);
    _wosConfigCache = null;
    try { CacheService.getScriptCache().remove('wos_config_v1'); } catch(eC) {}
  } catch(e) { Logger.log('_wosSetBackorderThreadId: ' + e); }
}

// Correr UNA VEZ desde el editor para instalar los 3 triggers
function WOS_instalarTriggerBackorder() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) {
    if (triggers[t].getHandlerFunction() === 'WOS_reporteBackorder') {
      ScriptApp.deleteTrigger(triggers[t]);
    }
  }
  ScriptApp.newTrigger('WOS_reporteBackorder').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(10).create();
  ScriptApp.newTrigger('WOS_reporteBackorder').timeBased().onWeekDay(ScriptApp.WeekDay.WEDNESDAY).atHour(10).create();
  ScriptApp.newTrigger('WOS_reporteBackorder').timeBased().onWeekDay(ScriptApp.WeekDay.FRIDAY).atHour(10).create();
  Logger.log('Triggers instalados: Lunes, Miércoles y Viernes a las 10 hs');
}

// ─────────────────────────────────────────────────────────────
// MIGRACIÓN ONE-TIME: recrea Pedidos_OTs con el schema de 26 cols
// idéntico a Pedidos_resellers. Ejecutar UNA VEZ desde el editor.
// Borra todas las filas existentes (la hoja era de prueba/vacía).
// ─────────────────────────────────────────────────────────────
function WOS_migrarPedidosOTs() {
  var ss   = SpreadsheetApp.openById(NOTAS_SS_ID);
  var hoja = ss.getSheetByName(HOJA_PEDIDOS_OT);
  if (!hoja) {
    Logger.log('WOS_migrarPedidosOTs: hoja "' + HOJA_PEDIDOS_OT + '" no encontrada — creando...');
    hoja = ss.insertSheet(HOJA_PEDIDOS_OT);
  }

  // Borrar todo el contenido previo
  hoja.clearContents();
  hoja.clearFormats();

  // Encabezado idéntico al de Pedidos_resellers (26 columnas)
  var headers = [
    'NUMERO','RESELLER','SKU','DESC','CANT_SOL','CANT_DESP','CANT_PEND',
    'PRECIO','STOCK_ORI','ESTADO','FECHA','ENVIO','PAGO','OBS',
    'FECHA_DESPACHO','NOTA_ENTREGA','TRACKING','THREAD_ID','FECHA_ESTADO',
    'TRANSPORTISTA_DESP','COSTO_ENVIO','PESO_ENVIO','NE_URL','OPERARIO',
    'SERIALES','CANT_CANCEL'
  ];
  hoja.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Freeze + formato encabezado
  hoja.setFrozenRows(1);
  hoja.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  // Ancho de columnas útiles
  hoja.setColumnWidth(1,  110); // NUMERO
  hoja.setColumnWidth(2,  140); // RESELLER
  hoja.setColumnWidth(3,  130); // SKU
  hoja.setColumnWidth(4,  200); // DESC
  hoja.setColumnWidth(10, 140); // ESTADO

  SpreadsheetApp.flush();
  Logger.log('WOS_migrarPedidosOTs: OK — hoja "' + HOJA_PEDIDOS_OT + '" lista con 26 cols.');
  return 'OK';
}
