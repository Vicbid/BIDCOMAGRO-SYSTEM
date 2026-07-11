// @version 3.5
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
    var hoja  = _getHojaPedidos();
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

    var msg = opcion === 'A'
      ? 'Registramos tu elecci\xf3n: <strong>Opci\xf3n A — Esperar el faltante en un segundo env\xedo.</strong><br>Despachamos lo disponible a la brevedad y los \xedtems faltantes llegan cuando haya stock.'
      : 'Registramos tu elecci\xf3n: <strong>Opci\xf3n B — Cancelar el faltante.</strong><br>Despachamos lo disponible y los \xedtems faltantes quedan cancelados.';
    return HtmlService.createHtmlOutput(_rfHtml(true, '\xa1Gracias, ' + reseller + '!', msg))
      .setTitle('Respuesta registrada \xb7 Pedido ' + numero);
  } catch(e) {
    Logger.log('_doGetRespFaltante ERROR [' + numero + ']: ' + e);
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
          return { email: email, nombre: String(datos[i][0] || email), tipo: String(datos[i][2] || '') };
        }
      }
    }
    return { email: email, nombre: email || 'Desconocido', tipo: '' };
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

// ── Setup: actualiza la data validation de columna J ──────────
// Ejecutar UNA VEZ desde el editor cuando se agregan estados nuevos.
function WOS_actualizarValidacion() {
  var estados = [
    'Pendiente_Revision', 'Confirmado', 'En_Espera_Reseller',
    'Cancelado', 'Preparado', 'Backorder', 'Preparado Parcial',
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
        items:           []
      };
      orden.push(num);
    }

    var cantSol  = Number(r[COL.CANT_SOL])  || 0;
    var cantDesp = Number(r[COL.CANT_DESP]) || 0;
    var cantPendRaw = r[COL.CANT_PEND];
    var cantPend = (cantPendRaw !== '' && cantPendRaw !== null && cantPendRaw !== undefined)
                   ? Number(cantPendRaw) : (cantSol - cantDesp);
    if (isNaN(cantPend)) cantPend = cantSol - cantDesp;

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
      cantCancel:   Number(r[COL.CANT_CANCEL] || 0),
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
    var EST_PRIORIDAD = [EST.PENDIENTE, EST.CONFIRMADO, EST.EN_ESPERA, EST.PREPARADO, EST.PREP_PARCIAL, EST.BACKORDER];

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
          GmailApp.getThreadById(canThreadId).replyAll(canPlain, {
            htmlBody: canHtml,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  _wosConfig().emailSoporte
          });
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

// Marca el pedido como Preparado registrando el N° de serie / SO de CADA unidad (OBLIGATORIO).
// Esto crea el manifiesto por unidad para poder auditar despachos (picker dice 10, reseller dice 4).
// seriales: [{ row: <fila 1-indexada>, seriales: "SN1, SN2, SO-...-003" }]
function WOS_prepararConSeriales(numero, seriales, operario) {
  try {
    numero = String(numero || '').trim();
    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    var ahora = new Date();
    operario  = String(operario || '');

    // fila (1-indexada) → seriales
    var serMap = {};
    if (Object.prototype.toString.call(seriales) === '[object Array]') {
      for (var s = 0; s < seriales.length; s++) {
        var rw = parseInt(seriales[s].row, 10);
        if (rw > 0) serMap[rw] = String(seriales[s].seriales || '').trim();
      }
    }

    var reseller = '', tocadas = 0;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      var estActual = String(datos[i][COL.ESTADO] || '');
      // No tocar filas ya cerradas (entregadas / canceladas) — solo lo que se está preparando
      if (estActual === EST.ENTREGADO || estActual === EST.CANCELADO || estActual === EST.ENTREGADO_CONF) continue;
      var fila = i + 1;
      hoja.getRange(fila, COL.ESTADO       + 1).setValue(EST.PREPARADO);
      hoja.getRange(fila, COL.FECHA_ESTADO + 1).setValue(ahora);
      if (operario) hoja.getRange(fila, COL.OPERARIO + 1).setValue(operario);
      if (serMap[fila] !== undefined && serMap[fila] !== '') {
        hoja.getRange(fila, COL.SERIALES + 1).setValue(serMap[fila]);
      }
      if (!reseller) reseller = String(datos[i][COL.RESELLER] || '');
      tocadas++;
    }
    if (!tocadas) return { ok: false, error: 'No se encontraron filas activas del pedido ' + numero + '.' };

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

        var cantSol  = Number(datos[i][COL.CANT_SOL])  || 0;
        var cantDesp = Number(datos[i][COL.CANT_DESP]) || 0;
        var cantPend = Math.max(0, cantSol - cantDesp);
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

    var lista = [];
    var keys  = Object.keys(mapa);
    for (var k = 0; k < keys.length; k++) lista.push(mapa[keys[k]]);
    lista.sort(function(a, b) { return a.fechaMs - b.fechaMs; }); // FIFO

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

      _wosSetEstado(hoja, datos, numero, EST.PREPARADO);

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
          GmailApp.getThreadById(threadId).replyAll(plain, {
            htmlBody: html,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  _wosConfig().emailSoporte
          });
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
// Retorna solo el mapa de unidades en camino { SKU: { total, ocs } }.
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
        if (!enCaminoMap[dSku]) enCaminoMap[dSku] = { total: 0, ocs: {} };
        enCaminoMap[dSku].total += pend;
        enCaminoMap[dSku].ocs[dCas] = (enCaminoMap[dSku].ocs[dCas] || 0) + pend;
      }
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
      if (ecData) {
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
        enCaminoOcs: ecOcs
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
      var nec      = (Number(wosData[i][COL.CANT_SOL]  || 0) - Number(wosData[i][COL.CANT_DESP] || 0));
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
    var hojaDetalle = masterSS.getSheetByName('COMPRAS_DETALLE');
    if (hojaDetalle) {
      var detData = hojaDetalle.getDataRange().getValues();
      for (var d = 1; d < detData.length; d++) {
        if (!casActivos[String(detData[d][0] || '').trim()]) continue;
        var dSku  = String(detData[d][1] || '').trim().toUpperCase();
        var dPend = (Number(detData[d][3] || 0) - Number(detData[d][4] || 0));
        if (dSku && dPend > 0) enCamino[dSku] = (enCamino[dSku] || 0) + dPend;
      }
    }

    // 4. Clasificar: sin cobertura vs cubiertos
    var sinCubrir = [];
    var cubiertos = [];
    for (var sku in backMap) {
      var item    = backMap[sku];
      var camino  = enCamino[sku] || 0;
      var gap     = item.nec - camino;
      var entrada = { sku: sku, desc: item.desc, nec: item.nec, camino: camino, gap: gap > 0 ? gap : 0, pedidos: item.pedidos };
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
      '<td style="padding:8px 12px;text-align:center;color:#1e40af">' + it.camino + '</td>' +
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
      '<td style="padding:6px 12px;text-align:center;font-size:12px;color:#166534;font-weight:700">' + cov.camino + '</td>' +
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
