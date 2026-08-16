// @version 1.2
// ============================================================
//  WOS — Edición administrativa de pedidos: cambiar SKU, descripción,
//  precio o cantidad de un ítem, o cancelar el pendiente de una línea
//  puntual. Pedido del usuario tras corregir a mano en la hoja un
//  código que el reseller pidió mal: reemplaza esa edición manual
//  (sin aviso, sin registro) por una función con motivo OBLIGATORIO
//  y aviso por mail al reseller. Exclusivo de usuarios Admin
//  (Usuarios_Internos col C = 'Admin', ver WOS_getUsuario()).
// ============================================================


// Snapshot de un pedido para el modal de edición — mismo shape que usa
// el resto de WOS (ver _wosLeerPedido), gateado a Admin.
function WOS_obtenerPedidoParaEditar(numero) {
  try {
    var u = WOS_getUsuario();
    if (!u.esAdmin) return { ok: false, error: 'Esta acción es exclusiva de administradores.' };
    numero = String(numero || '').trim();
    var ped = _wosLeerPedido(numero);
    if (!ped.reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    return { ok: true, numero: numero, reseller: ped.reseller, items: ped.items };
  } catch(e) {
    Logger.log('WOS_obtenerPedidoParaEditar: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// cambios: [{ row, accion:'editar', sku, desc, precio, cantSol } | { row, accion:'eliminar' }]
// Reglas (validadas TODAS antes de escribir nada — atómico, sin aplicar la mitad de un lote):
//   - motivo obligatorio.
//   - SKU/descripción/precio solo se pueden tocar en líneas SIN despacho (CANT_DESP = 0) —
//     no se reescribe la identidad de algo que ya salió del depósito.
//   - la cantidad nueva nunca puede bajar de lo ya despachado+cancelado (CANT_SOL − CANT_PEND).
//   - "eliminar" cancela el pendiente de la línea (mismo mecanismo que WOS_cambiarEstado al
//     cancelar un pedido entero); si la línea ya tenía algo despachado, queda Entregado/Listo_Retiro
//     (nada pendiente); si no tenía nada despachado, queda Cancelado.
// Un solo mail al reseller con el detalle de TODOS los cambios de esta tanda + el motivo.
function WOS_adminEditarPedido(numero, cambios, motivo, operario) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); }
  catch(eLock) { return { ok: false, error: 'Otra operación está en curso sobre este pedido. Reintentá en unos segundos.' }; }
  try {
    var u = WOS_getUsuario();
    if (!u.esAdmin) return { ok: false, error: 'Esta acción es exclusiva de administradores.' };

    motivo = String(motivo || '').trim();
    if (!motivo) return { ok: false, error: 'El motivo es obligatorio.' };
    numero = String(numero || '').trim();
    if (!cambios || !cambios.length) return { ok: false, error: 'No hay cambios para aplicar.' };

    var hoja = _getHojaPorNumero(numero);
    if (!hoja) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    var datos = hoja.getDataRange().getValues();

    var porFila = {};      // row (1-indexed) → índice 0-based en `datos`
    var reseller = '', threadId = '', envio = '';
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][COL.NUMERO] || '').trim() !== numero) continue;
      porFila[i + 1] = i;
      if (!reseller) {
        reseller = String(datos[i][COL.RESELLER]  || '');
        threadId = String(datos[i][COL.THREAD_ID] || '').trim();
        envio    = String(datos[i][COL.ENVIO]     || '');
      }
    }
    if (!reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    var esRetiroPed = envio.toLowerCase().indexOf('retiro') >= 0;

    // 1) VALIDAR todo el lote antes de tocar la hoja.
    var aplicar = [];
    for (var c = 0; c < cambios.length; c++) {
      var cb  = cambios[c] || {};
      var fil = parseInt(cb.row, 10);
      if (!porFila.hasOwnProperty(fil)) return { ok: false, error: 'Fila inválida en el pedido (row ' + fil + ').' };
      var idx = porFila[fil];

      var curSku     = String(datos[idx][COL.SKU]  || '');
      var curDesc    = String(datos[idx][COL.DESC] || '');
      var curPrecio  = Number(datos[idx][COL.PRECIO])   || 0;
      var curCantSol = Number(datos[idx][COL.CANT_SOL]) || 0;
      var curCantDesp= Number(datos[idx][COL.CANT_DESP])|| 0;
      var curCantPend= Number(datos[idx][COL.CANT_PEND])|| 0;
      var curCantCanc= Number(datos[idx][COL.CANT_CANCEL]) || 0;

      // BUG-PROOFING (plan "sistema a prueba de errores"): mismo patrón que WOS_despacharCompleto/
      // WOS_prepararConSeriales — el cliente manda la firma (sku|desc|precio) que vio al abrir el
      // modal de Editar Admin; si cambió desde entonces (otro admin, u otra vía), se aborta el
      // LOTE ENTERO antes de tocar nada, para no pisar una edición ajena sin que nadie se entere.
      if (cb.firma) {
        var _firmaActualAE = curSku + '|' + curDesc + '|' + curPrecio;
        if (cb.firma !== _firmaActualAE) {
          return { ok: false, desactualizado: true, numero: numero,
            error: curSku + ' (fila ' + fil + '): cambió desde que se abrió Editar Admin. No se aplicó nada.' };
        }
      }

      if (cb.accion === 'eliminar') {
        if (curCantPend <= 0) return { ok: false, error: _htmlEsc(curSku) + ': no tiene cantidad pendiente, no hay nada para eliminar.' };
        aplicar.push({
          idx: idx, accion: 'eliminar',
          antes: { sku: curSku, desc: curDesc, cantSol: curCantSol, precio: curPrecio },
          cantCancelNueva: curCantCanc + curCantPend,
          quedaDespachado: curCantDesp > 0
        });
        continue;
      }

      if (cb.accion !== 'editar') return { ok: false, error: 'Acción desconocida para la fila ' + fil + '.' };

      var nuevoSku     = String(cb.sku  || '').trim();
      var nuevoDesc    = String(cb.desc || '').trim();
      var nuevoPrecio  = Number(cb.precio)  || 0;
      var nuevoCantSol = Number(cb.cantSol) || 0;
      if (!nuevoSku) return { ok: false, error: 'El SKU no puede quedar vacío (fila ' + fil + ').' };

      var tocaIdentidad = (nuevoSku !== curSku) || (nuevoDesc !== curDesc) || (nuevoPrecio !== curPrecio);
      if (tocaIdentidad && curCantDesp > 0) {
        return { ok: false, error: _htmlEsc(curSku) + ' (ya tiene ' + curCantDesp + ' u. despachadas): no se puede cambiar SKU, descripción ni precio de una línea con despacho — solo se puede ajustar la cantidad pendiente.' };
      }
      var piso = curCantSol - curCantPend;   // = ya despachado + ya cancelado
      if (nuevoCantSol < piso) {
        return { ok: false, error: _htmlEsc(curSku) + ': la cantidad no puede bajar de ' + piso + ' (ya despachado/cancelado de esta línea).' };
      }

      if (nuevoSku === curSku && nuevoDesc === curDesc && nuevoPrecio === curPrecio && nuevoCantSol === curCantSol) continue; // sin cambios reales

      aplicar.push({
        idx: idx, accion: 'editar',
        antes:   { sku: curSku,  desc: curDesc,  cantSol: curCantSol,  precio: curPrecio },
        despues: { sku: nuevoSku, desc: nuevoDesc, cantSol: nuevoCantSol, precio: nuevoPrecio }
      });
    }

    if (!aplicar.length) return { ok: false, error: 'No hay cambios reales para aplicar.' };

    // 2) APLICAR — recién acá se escribe la hoja.
    var ahora = new Date();
    operario  = String(operario || u.email || '');
    for (var a = 0; a < aplicar.length; a++) {
      var it   = aplicar[a];
      var fila = it.idx + 1;
      if (it.accion === 'eliminar') {
        hoja.getRange(fila, COL.CANT_CANCEL + 1).setValue(it.cantCancelNueva);
        var estFinal = it.quedaDespachado ? (esRetiroPed ? EST.LISTO_RETIRO : EST.ENTREGADO) : EST.CANCELADO;
        var rEst = hoja.getRange(fila, COL.ESTADO + 1);
        rEst.clearDataValidations();
        rEst.setValue(estFinal);
      } else {
        hoja.getRange(fila, COL.SKU     + 1).setValue(_antiFormula(it.despues.sku));
        hoja.getRange(fila, COL.DESC    + 1).setValue(_antiFormula(it.despues.desc));
        hoja.getRange(fila, COL.PRECIO  + 1).setValue(it.despues.precio);
        hoja.getRange(fila, COL.CANT_SOL+ 1).setValue(it.despues.cantSol);
      }
      hoja.getRange(fila, COL.FECHA_ESTADO + 1).setValue(ahora);
      if (operario) hoja.getRange(fila, COL.OPERARIO + 1).setValue(_antiFormula(operario));
    }
    SpreadsheetApp.flush();

    // 3) Log interno — una fila por línea tocada, con el motivo.
    for (var l = 0; l < aplicar.length; l++) {
      var itl = aplicar[l];
      var detalleLog = (itl.accion === 'eliminar')
        ? 'Eliminado: ' + itl.antes.sku + ' (' + itl.antes.cantSol + ' u.) · Motivo: ' + motivo
        : 'Editado: ' + itl.antes.sku + ' → ' + itl.despues.sku + ' · Motivo: ' + motivo;
      _wosLogAccion('Editado por Admin', numero, reseller, operario, detalleLog);
    }

    // 4) Mail al reseller — un solo mail con TODOS los cambios de esta tanda + motivo.
    var avisoEnviado = _wosAvisarEdicionAdmin(numero, reseller, threadId, aplicar, motivo, operario);

    return { ok: true, cambios: aplicar.length, avisoEnviado: avisoEnviado };
  } catch(e) {
    Logger.log('WOS_adminEditarPedido: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}


// Arma y manda el mail de "tu pedido fue corregido" — reply en el hilo original si existe,
// mail nuevo si no. Devuelve true/false según si se pudo avisar (nunca tira: best-effort,
// los cambios en la hoja ya están aplicados y no dependen de que el mail salga).
function _wosAvisarEdicionAdmin(numero, reseller, threadId, aplicar, motivo, operario) {
  try {
    var email = '';
    try { email = _wosGetEmailReseller(reseller); } catch(eEm) {}
    if (!threadId && !email) return false;

    function esc(s) { return String(s == null ? '' : s).replace(/[&<>]/g, function(m) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;' }[m]; }); }
    function filaCambio(campo, antes, despues) {
      return "<tr><td style='padding:5px 10px;color:#888;font-size:12px'>" + campo + "</td>" +
        "<td style='padding:5px 10px;font-size:12px;color:#999;text-decoration:line-through'>" + esc(antes) + "</td>" +
        "<td style='padding:5px 10px;font-size:12px;font-weight:700;color:#00a3e0'>" + esc(despues) + "</td></tr>";
    }

    var filasHtml = '', filasPlain = '';
    for (var i = 0; i < aplicar.length; i++) {
      var it = aplicar[i];
      if (it.accion === 'eliminar') {
        filasHtml +=
          "<div style='margin:0 0 10px;padding:10px 12px;background:#fff5f5;border-left:3px solid #e74c3c;border-radius:0 6px 6px 0'>" +
            "<strong style='font-size:12px;color:#c0392b'>Ítem eliminado</strong><br>" +
            "<span style='font-size:12px;color:#555'>" + esc(it.antes.sku) + " — " + esc(it.antes.desc) + " (" + it.antes.cantSol + " u.)</span>" +
          "</div>";
        filasPlain += '• Eliminado: ' + it.antes.sku + ' — ' + it.antes.desc + ' (' + it.antes.cantSol + ' u.)\n';
        continue;
      }
      var camposHtml = '', camposPlain = [];
      if (it.antes.sku !== it.despues.sku) { camposHtml += filaCambio('SKU', it.antes.sku, it.despues.sku); camposPlain.push('SKU: ' + it.antes.sku + ' → ' + it.despues.sku); }
      if (it.antes.desc !== it.despues.desc) { camposHtml += filaCambio('Descripción', it.antes.desc, it.despues.desc); camposPlain.push('Descripción: ' + it.antes.desc + ' → ' + it.despues.desc); }
      if (it.antes.precio !== it.despues.precio) { camposHtml += filaCambio('Precio', 'USD ' + it.antes.precio, 'USD ' + it.despues.precio); camposPlain.push('Precio: USD ' + it.antes.precio + ' → USD ' + it.despues.precio); }
      if (it.antes.cantSol !== it.despues.cantSol) { camposHtml += filaCambio('Cantidad', it.antes.cantSol + ' u.', it.despues.cantSol + ' u.'); camposPlain.push('Cantidad: ' + it.antes.cantSol + ' u. → ' + it.despues.cantSol + ' u.'); }
      filasHtml +=
        "<div style='margin:0 0 10px'>" +
          "<div style='font-size:12px;font-weight:700;color:#1a1f2e;margin-bottom:4px'>" + esc(it.despues.sku) + "</div>" +
          "<table style='width:100%;border-collapse:collapse;border:1px solid #e0e3e8;border-radius:6px;overflow:hidden'>" + camposHtml + "</table>" +
        "</div>";
      filasPlain += '• ' + it.despues.sku + ': ' + camposPlain.join(' · ') + '\n';
    }

    var motivoHtml =
      "<div style='margin:14px 0 0;padding:12px 14px;background:#fff8e1;border-left:3px solid #f39c12;border-radius:0 6px 6px 0'>" +
        "<strong style='font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:#7a5800'>Motivo</strong>" +
        "<p style='margin:4px 0 0;font-size:13px;color:#444;white-space:pre-wrap'>" + esc(motivo) + "</p>" +
      "</div>";

    var html = _wosPortalHead('Corrección en tu pedido — ' + numero) +
      "<p style='font-size:14px;color:#666;margin:0 0 18px'>Hola <strong>" + esc(reseller) + "</strong>:</p>" +
      "<p style='font-size:13px;color:#555;margin:0 0 16px'>Hicimos una corrección en tu pedido " +
        "<strong style='color:#00a3e0'>" + esc(numero) + "</strong>:</p>" +
      filasHtml + motivoHtml +
      "<p style='font-size:12px;color:#888;margin-top:16px'>Ante cualquier consulta, respondé este correo.</p>" +
      _wosPortalFoot('Pedido ' + numero + ' · ' + reseller + '.');

    var plain = 'Hola ' + reseller + ',\n\nHicimos una corrección en tu pedido ' + numero + ':\n\n' +
      filasPlain + '\nMotivo: ' + motivo;

    var asunto = 'Corrección en tu pedido — ' + numero;
    var opts = { htmlBody: html, name: 'BIDCOMAGRO · Portal Resellers', replyTo: _wosConfig().emailSoporte };
    var ok = false;
    try {
      ok = _wosReplyHiloOriginal(threadId, plain, opts, [email]);
      if (!ok && email) { GmailApp.sendEmail(email, asunto, plain, opts); ok = true; }
    } catch(eS) {
      try { if (email) { GmailApp.sendEmail(email, asunto, plain, opts); ok = true; } } catch(eS2) { Logger.log('_wosAvisarEdicionAdmin fallback: ' + eS2); }
    }
    _wosRegistrarEmailLog(numero, (email || ''), 'Edición Admin', asunto, ok ? 'OK' : 'ERROR', threadId);
    return ok;
  } catch(e) {
    Logger.log('_wosAvisarEdicionAdmin: ' + e);
    return false;
  }
}
