// @version 3.31
// ============================================================
//  WOS — Router HTTP (doGet) + utilidades de sesión/log
//  El resto de la lógica se reorganizó (2026-07-30, sin cambios
//  funcionales) en:
//    WOS_Calidad.js   — QA / precisión del armado guiado
//    WOS_Pedidos.js   — CRUD de pedidos, estado, maestro/bolsas, preparación
//    WOS_Reservas.js  — ETA de compras + reservas de unidades en camino
//    WOS_Stock.js     — consulta de stock, despacho parcial/batch, ubicaciones
//    WOS_Reportes.js  — resumen de envíos + reporte de backorder + trigger
// ============================================================

// Se bumpea a mano junto con "<!-- @version X.Y -->" (la de arriba de todo, línea 1) de
// Despacho_Index.html — el cliente la trae embebida al cargar la página y la vuelve a consultar
// cada tanto (WOS_obtenerVersionActual) para avisar si quedó una pestaña vieja abierta. Ver
// _wosChequearVersionNueva en Despacho_Index.html.
var WOS_VERSION = '3.71';

function WOS_obtenerVersionActual() { return WOS_VERSION; }

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
  var tmpl = HtmlService.createTemplateFromFile('Despacho_Index');
  // WOS_WEBAPP_URL (hardcodeada, ver Despacho_Env.js) en vez de ScriptApp.getService().getUrl():
  // ya está documentado en este proyecto que esa función no siempre resuelve bien fuera de una
  // request HTTP real; acá SÍ estamos en una request real, pero reusar la misma constante
  // estable evita depender de eso dos veces por motivos distintos.
  tmpl.DEPLOY_URL   = WOS_WEBAPP_URL;
  tmpl.WOS_VERSION  = WOS_VERSION;
  return tmpl.evaluate()
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
  return WOS_WEBAPP_URL + '?page=manual';
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
