// @version 1.3
// ============================================================
//  WOS — Pedidos: CRUD, estado, maestro de artículos/bolsas,
//  preparación con seriales, backorder por pedido.
//  Extraído de Despacho_Código.js 3.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// ── Setup: actualiza la data validation de columna J ──────────
// Ejecutar UNA VEZ desde el editor cuando se agregan estados nuevos.
function WOS_actualizarValidacion() {
  var u = WOS_getUsuario();
  if (!u.autorizado) return 'No autorizado.';
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
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// ── Email fallback: cancela sin threadId (nuevo email) ───────
function _enviarEmailEstado(numero, reseller, obs) {
  var email = _wosGetEmailReseller(reseller);
  var asunto = 'Tu pedido ' + numero + ' fue cancelado — BIDCOMAGRO';
  if (!email) {
    Logger.log('_enviarEmailEstado: sin email para ' + reseller);
    _wosRegistrarEmailLog(numero, '', 'Pedido cancelado', asunto, 'OMITIDO: sin email', '');
    return;
  }
  var html = _wosPortalHead('Pedido cancelado — ' + numero) +
    "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + _htmlEsc(reseller) + "</strong>:</p>" +
    "<p style='font-size:13px;color:#555;margin:0 0 14px'>Tu pedido <strong style='color:#e74c3c'>" + numero + "</strong> fue <strong>cancelado</strong>.</p>" +
    (obs ? "<div style='background:#fdecea;border:1px solid #f5a5a5;border-radius:8px;padding:12px 16px;margin-bottom:16px'><p style='margin:0;font-size:12px;color:#7f1919'><strong>Motivo:</strong> " + _htmlEsc(obs) + "</p></div>" : '') +
    "<p style='font-size:13px;color:#555'>Si cre\xe9s que esto es un error, respond\xe9 este email y te ayudamos.</p>" +
    _wosPortalFoot('Pedido ' + numero + ' \xb7 ' + reseller + '.');

  try {
    GmailApp.sendEmail(email, asunto, '', {
      htmlBody: html,
      name:     'BIDCOMAGRO \xb7 Portal Resellers',
      replyTo:  _wosConfig().emailSoporte
    });
    _wosRegistrarEmailLog(numero, email, 'Pedido cancelado', asunto, 'OK', '');
    Logger.log('WOS email [cancelado fallback] → ' + email + ' pedido ' + numero);
  } catch(e) {
    Logger.log('_enviarEmailEstado ERROR: ' + e);
    _wosRegistrarEmailLog(numero, email, 'Pedido cancelado', asunto, 'ERROR: ' + String(e).substring(0, 150), '');
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
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// ── Cancela un pedido con motivo, lo guarda en OBS y envía email
function WOS_cancelarPedido(numero, motivo, operario, reqToken) {
 var u = WOS_getUsuario();
 if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
 return _wosLockIdempot(reqToken, function() {
  try {
    motivo   = String(motivo   || '').trim();
    operario = String(operario || '').trim();
    // Motivo obligatorio (plan "sistema a prueba de errores") — antes se cancelaba igual sin
    // motivo, solo se omitía escribirlo en OBS. El front ya lo exige, esto lo hace autoritativo.
    if (!motivo) return { ok: false, error: 'El motivo es obligatorio.' };

    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();

    // Bloqueo duro (decidido con el usuario): si algún ítem de este pedido ya está entregado,
    // no se puede cancelar el pedido ENTERO desde acá — WOS_cambiarEstado no distingue filas y
    // pondría Cancelado hasta en lo ya entregado. Para tocar solo lo pendiente hay que usar
    // "Editar pedido (Admin)" (WOS_adminEditarPedido), que sí opera fila por fila. Lectura fresca
    // (recién leída arriba), no depende de lo que el cliente tenía cargado al abrir el modal.
    for (var iChk = 1; iChk < datos.length; iChk++) {
      if (String(datos[iChk][COL.NUMERO] || '').trim() !== numero) continue;
      var estChk = String(datos[iChk][COL.ESTADO] || '').trim();
      if (estChk === EST.ENTREGADO || estChk === EST.ENTREGADO_CONF || estChk === EST.LISTO_RETIRO) {
        return { ok: false, error: 'El pedido ' + numero + ' ya tiene ítems entregados — no se puede cancelar entero. Usá "Editar pedido (Admin)" para modificar solo lo que sigue pendiente.' };
      }
    }

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      var obsActual = String(datos[i][COL.OBS] || '').trim();
      var obsNueva  = _antiFormula(motivo) + (obsActual ? ' · ' + obsActual : '');
      hoja.getRange(i + 1, COL.OBS + 1).setValue(obsNueva);
    }
    SpreadsheetApp.flush();
    return WOS_cambiarEstado(numero, EST.CANCELADO, operario);
  } catch(e) {
    Logger.log('WOS_cancelarPedido: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
 });
}


// ── Cambia el estado + graba timestamp + envía email si aplica
function WOS_cambiarEstado(numero, nuevoEstado, operario) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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
          "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + _htmlEsc(canReseller) + "</strong>:</p>" +
          "<p style='font-size:13px;color:#555;margin:0 0 14px'>Tu pedido " +
            "<strong style='color:#e74c3c'>" + numero + "</strong> fue <strong>cancelado</strong>." +
            (canObs ? " Motivo: <em>" + _htmlEsc(canObs) + "</em>." : '') + "</p>" +
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
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
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.', map: {} };
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.', map: {} };
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
        hoja.appendRow([_antiFormula(sku), _antiFormula(it.desc || ''), '', '', '', '']);
        fila = hoja.getLastRow();
        idx[su] = fila - 1;
      }
      hoja.getRange(fila, COL_MAESTRO.SKU       + 1).setValue(_antiFormula(sku));
      if (it.desc) hoja.getRange(fila, COL_MAESTRO.DESC + 1).setValue(_antiFormula(it.desc));
      hoja.getRange(fila, COL_MAESTRO.POR_BOLSA + 1).setValue(it.porBolsa ? true : false);
      var bulto = parseInt(it.bulto, 10) || 0;
      if (bulto > 0) hoja.getRange(fila, COL_MAESTRO.BULTO + 1).setValue(bulto);
      hoja.getRange(fila, COL_MAESTRO.FECHA + 1).setValue(now);
      if (operario) hoja.getRange(fila, COL_MAESTRO.OPERADOR + 1).setValue(_antiFormula(operario));
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
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, cantidad: 0, error: 'No autorizado.' };
  try {
    codigo = String(codigo || '').trim().toUpperCase();
    if (!codigo) return { ok: true, cantidad: 0 };
    var map = _wosMapaBolsas();
    var hit = map[codigo];
    if (hit && hit.cantidad > 1) return { ok: true, cantidad: hit.cantidad, sku: hit.sku };
    return { ok: true, cantidad: 0 };
  } catch(e) {
    Logger.log('WOS_resolverBolsa: ' + e);
    return { ok: false, cantidad: 0, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
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
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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

    // BUG-PROOFING (plan "sistema a prueba de errores"): mismo patrón que WOS_despacharCompleto
    // (WOS_GmailFlow.js) — el cliente manda la "firma" (sku|desc|precio) que vio al ABRIR el
    // modal de Preparar por cada fila; si cambió desde entonces (ej. alguien la corrigió vía
    // Editar Admin mientras el operador estaba preparando), se aborta TODO antes de escribir nada.
    var _desactualizadosPrep = [];
    if (Object.prototype.toString.call(seriales) === '[object Array]') {
      for (var fp = 0; fp < seriales.length; fp++) {
        var _rowP = parseInt(seriales[fp].row, 10);
        if (!(_rowP > 0) || !seriales[fp].firma) continue;
        var _dRowP = datos[_rowP - 1];
        if (!_dRowP) continue;
        var _firmaActualP = String(_dRowP[COL.SKU] || '') + '|' + String(_dRowP[COL.DESC] || '') + '|' + (Number(_dRowP[COL.PRECIO]) || 0);
        if (seriales[fp].firma !== _firmaActualP) _desactualizadosPrep.push(_rowP);
      }
    }
    if (_desactualizadosPrep.length) {
      Logger.log('WOS_prepararConSeriales: ' + numero + ' desactualizado, filas ' + _desactualizadosPrep.join(','));
      return { ok: false, desactualizado: true, numero: numero,
        error: 'El pedido ' + numero + ' cambió (SKU/descripción/precio) desde que se abrió Preparar. No se guardó nada.' };
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
      if (operario) hoja.getRange(fila, COL.OPERARIO + 1).setValue(_antiFormula(operario));
      if (serMap[fila] !== undefined && serMap[fila] !== '') {
        hoja.getRange(fila, COL.SERIALES + 1).setValue(_antiFormula(serMap[fila]));
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Reactiva solo los ítems en Backorder → Preparado, sin tocar los Entregado_Cerrado
// del primer despacho. Reemplaza el uso de WOS_cambiarEstado para la acción "reactivar".
// Sin gate acá: además del botón del cliente, la llama internamente el flujo de cron
// _wosNotificarIngresoPedido (WOS_GmailFlow.js, vía WOS_notificarIngresos) cuando detecta que
// llegó todo el stock de un backorder — Session.getActiveUser() no se comporta igual bajo un
// trigger. Ver WOS_reactivarBackorderCliente para el uso gateado desde el cliente.
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Wrapper gateado para el botón del cliente ("Reactivar Backorder"). WOS_reactivarBackorder en
// sí queda sin gate — ver comentario arriba.
function WOS_reactivarBackorderCliente(numero, operario) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  return WOS_reactivarBackorder(numero, operario);
}


// Reemplazada por WOS_despacharCompleto (en WOS_GmailFlow.js)
function WOS_despacharPedido(numero) {
  Logger.log('WOS_despacharPedido: función deprecada, usar WOS_despacharCompleto para ' + numero);
  return { ok: false, error: 'Función deprecada. Usar WOS_despacharCompleto.' };
}


// ── Busca pedidos en Backorder que necesitan un SKU ───────────
// Devuelve lista de { numero, reseller, cantPend } ordenada por fecha ASC (FIFO)
function WOS_buscarBackorderPorSKU(sku) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// ── Reactiva pedidos en Backorder cuando llega stock ─────────
// numeros: array de números de pedido a pasar a Preparado
// Notifica a cada reseller en su hilo original
function WOS_recibirMercaderia(sku, cantRecibida, numeros) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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
            "<p style='font-size:14px;color:#666;margin:0 0 22px'>Hola <strong>" + _htmlEsc(reseller) + "</strong>:</p>" +
            "<p style='font-size:13px;color:#555;line-height:1.6'>" +
              "El stock de <strong>" + _htmlEsc(skuUp) + "</strong> que estabas esperando llegó. " +
              "Estamos preparando tu pedido <strong style='color:#00a3e0'>" + numero + "</strong> para el despacho " +
              "y te avisaremos en cuanto sea enviado." +
            "</p>" +
            _wosPortalFoot('Pedido ' + numero + ' · ' + reseller + '.');
          var plain = 'Hola ' + reseller + ',\n\nEl stock de ' + skuUp + ' llegó y estamos preparando tu pedido ' + numero + ' para el despacho. Te avisaremos cuando sea enviado.';
          // Reply al hilo apuntando a los destinatarios ORIGINALES (no a una conversación aparte).
          var _asuntoStock = 'Stock disponible — Pedido ' + numero;
          var _notifOk = _wosReplyHiloOriginal(threadId, plain, {
            htmlBody: html,
            name:     'BIDCOMAGRO · Portal Resellers',
            replyTo:  _wosConfig().emailSoporte
          }, [_wosGetEmailReseller(reseller)]);
          if (!_notifOk) {
            var _emNotif = _wosGetEmailReseller(reseller);
            if (_emNotif) {
              GmailApp.sendEmail(_emNotif, _asuntoStock, plain, { htmlBody: html, name: 'BIDCOMAGRO · Portal Resellers', replyTo: _wosConfig().emailSoporte });
              _wosRegistrarEmailLog(numero, _emNotif, 'Stock disponible', _asuntoStock, 'OK-FALLBACK', '');
            } else {
              _wosRegistrarEmailLog(numero, '', 'Stock disponible', _asuntoStock, 'OMITIDO: sin email', '');
            }
          } else {
            _wosRegistrarEmailLog(numero, _wosGetEmailReseller(reseller), 'Stock disponible', _asuntoStock, 'OK-THREAD', threadId);
          }
        } catch(eN) {
          Logger.log('WOS_recibirMercaderia notif [' + numero + ']: ' + eN);
          _wosRegistrarEmailLog(numero, '', 'Stock disponible', 'Stock disponible — Pedido ' + numero, 'ERROR: ' + String(eN).substring(0, 150), '');
        }
      }
      reactivados.push(numero);
    }

    SpreadsheetApp.flush();
    Logger.log('WOS_recibirMercaderia: SKU=' + skuUp + ' cant=' + cant + ' → reactivados: ' + reactivados.join(', '));
    return { ok: true, reactivados: reactivados };
  } catch(e) {
    Logger.log('WOS_recibirMercaderia ERROR: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Revierte un pedido despachado por error de vuelta a "Preparado"
function WOS_revertirAPreparado(numero, operario) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
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
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Lee las últimas N entradas del log para un pedido específico
function WOS_getHistorial(numero) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, entradas: [] };
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
function WOS_actualizarObs(numero, obs, operario) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    var hoja  = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado.' };
    var datos = hoja.getDataRange().getValues();
    var reseller = '', obsAnterior = '';
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      reseller    = reseller    || String(datos[i][COL.RESELLER] || '').trim();
      obsAnterior = obsAnterior || String(datos[i][COL.OBS]       || '').trim();
      hoja.getRange(i + 1, COL.OBS + 1).setValue(_antiFormula(obs || ''));
    }
    SpreadsheetApp.flush();
    // Bug reportado por el usuario (auditoría, "sistema a prueba de errores"): esta función
    // pisaba OBS sin dejar ningún rastro — quién lo cambió, ni el valor anterior.
    _wosLogAccion('Observación editada', numero, reseller, String(operario || ''),
      '"' + obsAnterior + '" → "' + String(obs || '') + '"');
    return { ok: true };
  } catch(e) {
    Logger.log('WOS_actualizarObs: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Polling liviano — compara timestamps de FECHA_ESTADO contra ultimoMs del cliente
function WOS_checkCambios(ultimoMs) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false };
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
