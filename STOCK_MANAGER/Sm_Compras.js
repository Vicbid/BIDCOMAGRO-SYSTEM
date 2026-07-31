// ── STOCK MANAGER — Compras ─────────────────────────────────────
// @version 1.6

// ============================================================
//  COMPRAS DJI
// ============================================================
var ESTADOS_CAS = [
  "Comprado","Pagado","Envío confirmado","Forwarder HK",
  "En vuelo","En aduana","En depósito"
];

function cargarCompras() {
  try {
    var d   = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var out = [];
    for (var i = 1; i < d.length; i++) {
      var f = d[i];
      var fechaVuelo  = f[8]  instanceof Date ? f[8]  : null;
      var fechaDepo   = f[10] instanceof Date ? f[10] : null;
      var diasTransito = (fechaVuelo && !fechaDepo)
        ? Math.floor((new Date() - fechaVuelo) / 86400000) : null;
      out.push({
        fila: i+1, cas: String(f[0]), fechaPedido: _fmtFecha(f[1]),
        estado: String(f[2]||"Comprado"), metodoPago: String(f[3]||""),
        fechas: {
          comprado:    _fmtFecha(f[4]),  pagado:       _fmtFecha(f[5]),
          confEnvio:   _fmtFecha(f[6]),  forwarderHK:  _fmtFecha(f[7]),
          vuelo:       _fmtFecha(f[8]),  aduana:       _fmtFecha(f[9]),
          deposito:    _fmtFecha(f[10])
        },
        diasEnTransito: diasTransito,
        operador: String(f[11]||""), observaciones: String(f[12]||""),
        ultimaActualizacion: f[13] instanceof Date ? _fmtFecha(f[13]) : (f[13] ? String(f[13]) : null),
        eta:    _fmtFecha(f[SCHEMA.COMPRAS_DJI.ETA]),
        etaISO: (f[SCHEMA.COMPRAS_DJI.ETA] instanceof Date)
                ? Utilities.formatDate(f[SCHEMA.COMPRAS_DJI.ETA], Session.getScriptTimeZone(), "yyyy-MM-dd") : ""
      });
    }
    out.sort(function(a,b){ return ESTADOS_CAS.indexOf(a.estado) - ESTADOS_CAS.indexOf(b.estado); });
    return out;
  } catch(e) { return []; }
}

function registrarCAS(cas, metodoPago, operador) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS);
    var d    = getSheetValues(hoja);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toUpperCase() === String(cas).trim().toUpperCase())
        return { ok: false, msg: "El CAS ya existe" };
    }
    hoja.appendRow([cas.trim().toUpperCase(), new Date(), "Comprado", metodoPago||"",
                    new Date(), "","","","","","", operador||"", ""]);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function actualizarEstadoCAS(cas, nuevoEstado, observaciones, operador, eta) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS);
    var d    = getSheetValues(hoja);
    var casB = String(cas).trim().toUpperCase();
    var colFecha = { "Comprado":5,"Pagado":6,"Envío confirmado":7,"Forwarder HK":8,"En vuelo":9,"En aduana":10,"En depósito":11 };
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toUpperCase() !== casB) continue;
      var estadoAnterior = String(d[i][SCHEMA.COMPRAS_DJI.ESTADO] || '');
      var ahora = new Date();

      hoja.getRange(i+1, 3).setValue(nuevoEstado);
      if (colFecha[nuevoEstado]) hoja.getRange(i+1, colFecha[nuevoEstado]).setValue(ahora);
      if (observaciones) hoja.getRange(i+1, 13).setValue(String(d[i][12]||"")+" | "+observaciones);
      hoja.getRange(i+1, 12).setValue(operador||"");
      hoja.getRange(i+1, 14).setValue(ahora); // Última actualización

      // ETA (llegada estimada) — solo si se pasó el parámetro. '' limpia; 'yyyy-MM-dd' setea la fecha.
      if (eta !== undefined && eta !== null) {
        var etaVal = "";
        var etaParts = String(eta).split("-");
        if (etaParts.length === 3) {
          etaVal = new Date(parseInt(etaParts[0],10), parseInt(etaParts[1],10)-1, parseInt(etaParts[2],10));
        }
        hoja.getRange(i+1, SCHEMA.COMPRAS_DJI.ETA + 1).setValue(etaVal);
      }

      // Registrar en historial
      _logHistorialCAS(casB, estadoAnterior, nuevoEstado, operador, observaciones);

      invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);
      if (nuevoEstado === "En depósito") {
        _alertarBackordersPendientes(cas);
      }
      return { ok: true };
    }
    return { ok: false, msg: "CAS no encontrado" };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function _logHistorialCAS(idCas, estadoAnterior, estadoNuevo, operador, observaciones) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.HISTORIAL_COMPRAS);
    if (!hoja) {
      hoja = getDb().insertSheet(SCHEMA.SHEETS.HISTORIAL_COMPRAS);
      hoja.appendRow(['Fecha','ID_CAS','Estado anterior','Estado nuevo','Operador','Observaciones']);
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([new Date(), idCas, estadoAnterior, estadoNuevo, operador||'', observaciones||'']);
    invalidateSheetValues(SCHEMA.SHEETS.HISTORIAL_COMPRAS);
  } catch(e) { Logger.log('_logHistorialCAS: ' + e); }
}

function recibirMercaderia(cas, items, operador, deposito) {
  // items = [{ codigo, descripcion, cantRecibida }]
  deposito = String(deposito || 'BA').trim().toUpperCase();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hojaStr = getSheet(SCHEMA.SHEETS.STOCK);
    var dStr    = getSheetValues(hojaStr);
    var stockIdx = {};
    for (var s = 1; s < dStr.length; s++) {
      stockIdx[String(dStr[s][0]).trim().toUpperCase()] = s+1;
    }

    // Índice de DB_REPUESTOS para enriquecer SKUs nuevos
    var dbRepIdx = {};
    var dDbRep = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    for (var dr = 1; dr < dDbRep.length; dr++) {
      var drSku = String(dDbRep[dr][1]||'').trim().toUpperCase();
      if (drSku) dbRepIdx[drSku] = dr;
    }

    var hoy = new Date();
    var strChanged = false;
    var nuevasFila = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var cod  = String(item.codigo).trim().toUpperCase();
      var cant = parseInt(item.cantRecibida)||0;
      if (cant <= 0) continue;
      var filaStr = stockIdx[cod];
      if (filaStr) {
        var actual = parseInt(dStr[filaStr-1][2])||0;
        var nuevo  = actual + cant;
        dStr[filaStr-1][2] = nuevo;
        dStr[filaStr-1][7] = hoy;
        // Rellenar campos vacíos desde DB_REPUESTOS si los tiene
        var drI = dbRepIdx[cod];
        if (drI !== undefined) {
          if (!dStr[filaStr-1][4] && dDbRep[drI][5]) dStr[filaStr-1][4] = String(dDbRep[drI][5]); // categoria
          if (!dStr[filaStr-1][6] && dDbRep[drI][3]) dStr[filaStr-1][6] = String(dDbRep[drI][3]); // modelos
        }
        strChanged = true;
        _registrarMovimiento("ENTRADA_COMPRA", cod, String(dStr[filaStr-1][1]),
          cant, nuevo, cas, operador||"", deposito);
      } else {
        // Código nuevo: enriquecer con datos de DB_REPUESTOS / CATALOGO_DJI si existen
        var drIdx = dbRepIdx[cod];
        var descFinal = (drIdx !== undefined && dDbRep[drIdx][2]) ? String(dDbRep[drIdx][2]) : (item.descripcion||'');
        var catFinal  = (drIdx !== undefined && _catValida(String(dDbRep[drIdx][5]||''))) ? String(dDbRep[drIdx][5]) : '';
        var modFinal  = (drIdx !== undefined && dDbRep[drIdx][3]) ? String(dDbRep[drIdx][3]) : '';
        var fobFinal  = (drIdx !== undefined && dDbRep[drIdx][6]) ? parseFloat(dDbRep[drIdx][6])||0 : 0;
        nuevasFila.push([cod, descFinal, cant, 0, catFinal, '', modFinal, hoy, '', false]);
        _registrarMovimiento("ENTRADA_COMPRA", cod, descFinal,
          cant, cant, cas, operador||"", deposito);
        // Si no existe en DB_REPUESTOS, crearlo con los datos disponibles
        if (drIdx === undefined) {
          var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
          if (hojaRep) {
            hojaRep.appendRow(['', cod, descFinal, modFinal, '', catFinal, fobFinal]);
            invalidateSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
          }
        }
      }
    }
    if (strChanged) hojaStr.getDataRange().setValues(dStr);
    for (var na = 0; na < nuevasFila.length; na++) {
      hojaStr.appendRow(nuevasFila[na]);
    }
    // Carmen se actualiza vía _escribirEnRecibidos más abajo — no tocar col C directamente
    // Cruzar con RESERVAS activas para este CAS
    var hojaRes  = getSheet(SCHEMA.SHEETS.RESERVAS);
    var dRes     = hojaRes ? getSheetValues(hojaRes) : [];
    var R        = SCHEMA.RESERVAS_STOCK;
    var destinos = [];
    var resChanged = false;
    for (var ri = 1; ri < dRes.length; ri++) {
      var resRow = dRes[ri];
      if (String(resRow[R.ESTADO]) !== 'Activa') continue;
      if (String(resRow[R.CAS_REF]).trim().toUpperCase() !== String(cas).trim().toUpperCase()) continue;
      var rSku = String(resRow[R.SKU]).trim().toUpperCase();
      for (var ii = 0; ii < items.length; ii++) {
        if (String(items[ii].codigo).trim().toUpperCase() === rSku) {
          dRes[ri][R.ESTADO] = 'Cumplida';
          resChanged = true;
          destinos.push({
            sku:        rSku,
            descripcion:String(resRow[R.DESCRIPCION]),
            cantidad:   parseInt(resRow[R.CANTIDAD]) || 0,
            origen:     String(resRow[R.ORIGEN]),
            referencia: String(resRow[R.ID_REFERENCIA])
          });
          break;
        }
      }
    }
    if (hojaRes && resChanged) {
      hojaRes.getDataRange().setValues(dRes);
      invalidateSheetValues(SCHEMA.SHEETS.RESERVAS);
    }

    // Actualizar COMPRAS_DETALLE: cantidades recibidas + estado por ítem
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (hojaCD) {
      var CD    = SCHEMA.COMPRAS_DETALLE;
      var casUp = String(cas).trim().toUpperCase();
      var dCD   = hojaCD.getDataRange().getValues();
      var cdChanged = false;
      for (var cdi = 1; cdi < dCD.length; cdi++) {
        if (String(dCD[cdi][CD.ID_CAS]).trim().toUpperCase() !== casUp) continue;
        var skuCD = String(dCD[cdi][CD.SKU]).trim().toUpperCase();
        for (var rii = 0; rii < items.length; rii++) {
          if (String(items[rii].codigo).trim().toUpperCase() !== skuCD) continue;
          var prevRecib  = parseInt(dCD[cdi][CD.CANTIDAD_RECIBIDA]) || 0;
          var newRecib   = prevRecib + (parseInt(items[rii].cantRecibida) || 0);
          var pedida     = parseInt(dCD[cdi][CD.CANTIDAD_PEDIDA]) || 0;
          var nuevoEst   = newRecib >= pedida ? 'Completo' : (newRecib > 0 ? 'Parcial' : 'Pendiente');
          dCD[cdi][CD.CANTIDAD_RECIBIDA] = newRecib;
          dCD[cdi][CD.ESTADO]            = nuevoEst;
          cdChanged = true;
          break;
        }
      }
      if (cdChanged) hojaCD.getDataRange().setValues(dCD);
      invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    }

    // Escribir cada ítem en la hoja Recibidos del spreadsheet Carmen
    _escribirEnRecibidos(cas, items, '');

    // Encolar avisos a resellers de WOS cuyas unidades en camino bloqueadas llegaron con este CAS
    _encolarNotifIngresoWOS(cas, items);

    // Marcar CAS como En depósito
    actualizarEstadoCAS(cas, "En depósito", "Recepción registrada", operador);
    return { ok: true, destinoMercaderia: destinos };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Cruza el CAS recibido con RESERVAS_EN_CAMINO (unidades bloqueadas por WOS a resellers que
// esperan) y encola un aviso por (pedido, SKU) en NOTIF_INGRESOS_WOS. El mail al reseller lo
// arma y envía WOS en su detector (corre cada 10 min): si con este ingreso el pedido queda
// completo lo prepara y avisa sin preguntar; si sigue faltando algo pregunta "¿despachamos ahora
// o esperás el resto?" (links A/B). SM solo encola — no toca reservas ni manda mails al reseller.
function _encolarNotifIngresoWOS(cas, items) {
  try {
    var casUp   = String(cas || '').trim().toUpperCase();
    var hojaRes = getSheet('RESERVAS_EN_CAMINO');
    if (!hojaRes) return;
    var dRes = hojaRes.getDataRange().getValues();
    // RESERVAS_EN_CAMINO (la mantiene WOS): FECHA·PEDIDO·RESELLER·SKU·CAS·N_AIR·ETA·CANTIDAD·ESTADO
    var recPorSku = {};
    for (var i = 0; i < items.length; i++) {
      var c = String(items[i].codigo || '').trim().toUpperCase();
      var q = parseInt(items[i].cantRecibida) || 0;
      if (c && q > 0) recPorSku[c] = (recPorSku[c] || 0) + q;
    }
    var filas = [], ahora = new Date();
    for (var r = 1; r < dRes.length; r++) {
      if (String(dRes[r][8] || '').trim() !== 'Activa') continue;
      if (String(dRes[r][4] || '').trim().toUpperCase() !== casUp) continue;
      var sku  = String(dRes[r][3] || '').trim().toUpperCase();
      var disp = recPorSku[sku] || 0;
      if (disp <= 0) continue;
      var take = Math.min(disp, parseInt(dRes[r][7]) || 0);
      if (take <= 0) continue;
      recPorSku[sku] = disp - take;   // recepción parcial: repartir lo recibido entre reservas en orden
      filas.push([ahora, casUp, String(dRes[r][1] || '').trim(), String(dRes[r][2] || ''), sku, take, 'Pendiente', '']);
    }
    if (!filas.length) return;
    var hojaQ = getSheet('NOTIF_INGRESOS_WOS');
    if (!hojaQ) {
      hojaQ = getSheet(SCHEMA.SHEETS.COMPRAS).getParent().insertSheet('NOTIF_INGRESOS_WOS');
      hojaQ.appendRow(['FECHA', 'CAS', 'PEDIDO', 'RESELLER', 'SKU', 'CANTIDAD', 'ESTADO', 'RESULTADO']);
      hojaQ.setFrozenRows(1);
    }
    hojaQ.getRange(hojaQ.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
  } catch(e) { Logger.log('_encolarNotifIngresoWOS: ' + e); }
}

// Cuando cambia la ETA de un CAS (Cruce externo → "Usar externo", ver sincronizarItemsCAS),
// avisa a los resellers que tengan una reserva ACTIVA de ese CAS/SKU en RESERVAS_EN_CAMINO de
// que se atrasa. Mismo patrón que _encolarNotifIngresoWOS: SM solo encola, el mail lo arma y
// envía WOS (WOS_notificarCambiosEta). `retrasos`: [{sku, desc, etaAnterior, etaNueva}].
// Devuelve la cantidad de avisos encolados.
function _encolarNotifEtaCambioWOS(cas, retrasos) {
  try {
    if (!retrasos || !retrasos.length) return 0;
    var casUp   = String(cas || '').trim().toUpperCase();
    var hojaRes = getSheet('RESERVAS_EN_CAMINO');
    if (!hojaRes) return 0;
    var dRes = hojaRes.getDataRange().getValues();
    // RESERVAS_EN_CAMINO (la mantiene WOS): FECHA·PEDIDO·RESELLER·SKU·CAS·N_AIR·ETA·CANTIDAD·ESTADO
    var etaPorSku = {};
    for (var t = 0; t < retrasos.length; t++) etaPorSku[retrasos[t].sku] = retrasos[t];

    var filas = [], ahora = new Date();
    for (var r = 1; r < dRes.length; r++) {
      if (String(dRes[r][8] || '').trim() !== 'Activa') continue;
      if (String(dRes[r][4] || '').trim().toUpperCase() !== casUp) continue;
      var sku = String(dRes[r][3] || '').trim().toUpperCase();
      var ret = etaPorSku[sku];
      if (!ret) continue;
      var cant = parseInt(dRes[r][7]) || 0;
      if (cant <= 0) continue;
      filas.push([ahora, casUp, String(dRes[r][1] || '').trim(), String(dRes[r][2] || ''), sku, cant, ret.etaAnterior, ret.etaNueva, 'Pendiente', '']);
    }
    if (!filas.length) return 0;
    var hojaQ = getSheet('NOTIF_ETA_CAMBIO_WOS');
    if (!hojaQ) {
      hojaQ = getSheet(SCHEMA.SHEETS.COMPRAS).getParent().insertSheet('NOTIF_ETA_CAMBIO_WOS');
      hojaQ.appendRow(['FECHA', 'CAS', 'PEDIDO', 'RESELLER', 'SKU', 'CANTIDAD', 'ETA_ANTERIOR', 'ETA_NUEVA', 'ESTADO', 'RESULTADO']);
      hojaQ.setFrozenRows(1);
    }
    hojaQ.getRange(hojaQ.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    return filas.length;
  } catch(e) { Logger.log('_encolarNotifEtaCambioWOS: ' + e); return 0; }
}

function _alertarBackordersPendientes(cas) {
  try {
    var casKey = String(cas || '').trim().toUpperCase();

    // SKUs del CAS recibido (de COMPRAS_DETALLE) — usados para filtrar qué backorders revisar
    var hojaCD   = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (!hojaCD) return;
    var dCD      = getSheetValues(hojaCD);
    var CD       = SCHEMA.COMPRAS_DETALLE;
    var skusCAS  = {}; // SKU → desc
    for (var ci = 1; ci < dCD.length; ci++) {
      if (String(dCD[ci][CD.ID_CAS] || '').trim().toUpperCase() !== casKey) continue;
      var sku = String(dCD[ci][CD.SKU] || '').trim().toUpperCase();
      if (sku) skusCAS[sku] = String(dCD[ci][CD.DESCRIPCION] || '');
    }
    if (!Object.keys(skusCAS).length) return;

    // Stock actual desde STOCK_REPUESTOS (ya actualizado por la recepción)
    var dStr     = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S        = SCHEMA.STOCK_REPUESTOS;
    var stockMap = {}; // SKU → stock actual
    for (var si = 1; si < dStr.length; si++) {
      var sCod = String(dStr[si][S.CODIGO] || '').trim().toUpperCase();
      if (sCod && skusCAS[sCod] !== undefined) stockMap[sCod] = parseInt(dStr[si][S.STOCK_ACTUAL]) || 0;
    }

    // Backorders en WOS Pedidos_resellers con estado 'Backorder' para esos SKUs
    var hojaWOS = SpreadsheetApp.openById('1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw')
                    .getSheetByName('Pedidos_resellers');
    if (!hojaWOS) return;
    var dWOS     = hojaWOS.getDataRange().getValues();
    // COL indices (Despacho_Env.js): 0=NUMERO,1=RESELLER,2=SKU,3=DESC,4=CANT_SOL,5=CANT_DESP,9=ESTADO,25=CANT_CANCEL
    var afectadas  = [];
    var procesados = {};
    for (var wi = 1; wi < dWOS.length; wi++) {
      if (String(dWOS[wi][9] || '').trim() !== 'Backorder') continue;
      var wSku = String(dWOS[wi][2] || '').trim().toUpperCase();
      if (!skusCAS[wSku]) continue;
      var numero = String(dWOS[wi][0] || '').trim();
      var key    = numero + '|' + wSku;
      if (procesados[key]) continue;
      procesados[key] = true;
      var pend = Math.max(0, (Number(dWOS[wi][4]) || 0) - (Number(dWOS[wi][5]) || 0) - (Number(dWOS[wi][25]) || 0));
      if (pend <= 0) continue;
      var stockActual = stockMap[wSku] !== undefined ? stockMap[wSku] : 0;
      afectadas.push({
        numero:   numero,
        reseller: String(dWOS[wi][1] || ''),
        sku:      wSku,
        desc:     skusCAS[wSku] || String(dWOS[wi][3] || ''),
        pend:     pend,
        stock:    stockActual,
        cubre:    stockActual >= pend
      });
    }
    if (!afectadas.length) return;

    var filas = afectadas.map(function(a) {
      var cobertura = a.cubre
        ? '<span style="color:#27ae60;font-weight:700">&#10003; Stock suficiente (' + a.stock + ' disp. / ' + a.pend + ' pend.)</span>'
        : '<span style="color:#e67e22;font-weight:700">Stock insuf.: ' + a.stock + '/' + a.pend + ' u.</span>';
      return '<tr>' +
        '<td style="padding:6px 10px;font-size:12px;font-weight:700;color:#00a3e0">' + a.numero + '</td>' +
        '<td style="padding:6px 10px;font-size:12px">' + a.reseller + '</td>' +
        '<td style="padding:6px 10px;font-size:12px">' + a.sku + '</td>' +
        '<td style="padding:6px 10px;font-size:12px">' + a.desc + '</td>' +
        '<td style="padding:6px 10px;font-size:12px;text-align:center">' + cobertura + '</td>' +
        '</tr>';
    }).join('');

    var totalCubren   = afectadas.filter(function(a) { return a.cubre; }).length;
    var totalParciales = afectadas.length - totalCubren;
    var asunto = '[WOS] Backorders desbloqueados — CAS ' + casKey +
      ' (' + (totalCubren ? totalCubren + ' total' : '') +
      (totalCubren && totalParciales ? ', ' : '') +
      (totalParciales ? totalParciales + ' parcial' : '') + ')';

    GmailApp.sendEmail(SM_CONFIG.EMAIL_SUPERVISOR, asunto, '', {
      htmlBody:
        '<div style="font-family:sans-serif;max-width:650px">' +
        '<div style="background:#00a3e0;padding:16px 20px;border-radius:8px 8px 0 0">' +
          '<span style="color:#fff;font-size:16px;font-weight:700">Backorders desbloqueados — ' + casKey + '</span>' +
        '</div>' +
        '<div style="background:#fff;border:1px solid #ddd;padding:18px 20px;border-radius:0 0 8px 8px">' +
          '<p style="font-size:13px;color:#444;margin:0 0 14px">La recepci\xf3n de <strong>' + casKey + '</strong> cubre (total o parcialmente) los siguientes backorders en WOS. Ingres\xe1 al WOS para despachar.</p>' +
          '<table style="width:100%;border-collapse:collapse;border:1px solid #eee">' +
            '<thead><tr style="background:#f5f5f5">' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">Pedido</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">Reseller</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">SKU</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">Descripci\xf3n</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:center">Cobertura</th>' +
            '</tr></thead>' +
            '<tbody>' + filas + '</tbody>' +
          '</table>' +
          '<p style="font-size:11px;color:#999;margin-top:14px">Cobertura calculada sobre las unidades recibidas en este CAS vs. unidades pendientes en cada pedido WOS.</p>' +
        '</div></div>',
      name: SM_CONFIG.NOMBRE_REMITENTE,
      replyTo: SM_CONFIG.EMAIL_SUPERVISOR
    });
  } catch(e) { Logger.log('_alertarBackordersPendientes: ' + e); }
}


// Escribe los ítems recibidos en la hoja "Recibidos" del spreadsheet Carmen.
// Formato: A=codigo, B=descripcion, C=cantRecibida, D=CAS, E=fecha dd/MM/yyyy, F=(vacío), G=observaciones del CAS (número de Air, etc.)
function _escribirEnRecibidos(cas, items, observaciones) {
  try {
    var ss   = _getCarmenSS();
    var hoja = ss.getSheetByName('Recibidos');
    if (!hoja) { Logger.log('_escribirEnRecibidos: hoja Recibidos no encontrada'); return; }

    var tz       = Session.getScriptTimeZone();
    var fechaStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    var casStr   = String(cas || '').trim().toUpperCase();

    // Leer observaciones del CAS desde COMPRAS_DJI (ahí se guarda el número de Air y otros seguimientos)
    var obs = String(observaciones || '');
    try {
      var dComp = getSheetValues(SCHEMA.SHEETS.COMPRAS);
      for (var c = 1; c < dComp.length; c++) {
        if (String(dComp[c][SCHEMA.COMPRAS_DJI.ID_CAS] || '').trim().toUpperCase() === casStr) {
          obs = String(dComp[c][SCHEMA.COMPRAS_DJI.OBSERVACIONES] || '');
          break;
        }
      }
    } catch(eObs) { Logger.log('_escribirEnRecibidos obs lookup: ' + eObs); }

    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);

    for (var i = 0; i < items.length; i++) {
      var it       = items[i];
      var cant     = parseInt(it.cantRecibida) || 0;
      if (cant <= 0) continue;
      var codKey  = String(it.codigo    || '').trim().toUpperCase();
      var ubicKey = String(it.ubicacion || '').trim().toUpperCase();

      // PN | Desc | Cant | Origen | Fecha | Comprobante | obs | Aparece Inv | Ubicación
      hoja.appendRow([codKey, String(it.descripcion || ''), cant, casStr, fechaStr, '', obs, '', ubicKey]);

      // Sumar a UBICACIONES si se indicó ubicación
      if (ubicKey && hojaUbic) {
        var dU    = hojaUbic.getDataRange().getValues();
        var found = false;
        for (var ui = 1; ui < dU.length; ui++) {
          if (String(dU[ui][0] || '').trim().toUpperCase() === codKey &&
              String(dU[ui][1] || '').trim().toUpperCase() === ubicKey) {
            hojaUbic.getRange(ui + 1, 3).setValue((parseFloat(dU[ui][2]) || 0) + cant);
            found = true;
            break;
          }
        }
        if (!found) hojaUbic.appendRow([codKey, ubicKey, cant]);
      }
    }
  } catch(e) {
    Logger.log('_escribirEnRecibidos: ' + e);
  }
}

// ============================================================
//  CRUCE COMPRAS EXTERNAS
// ============================================================
var _PEDIDOS_EXT_SS_ID = '15Y4tri7Egpa2Tjvq1kPXuuR7OUIsQVXgjPFUwthr7Sw';

// Normaliza el valor de "FECHA ESTIMADA EN PLANTA" del planner externo.
// Date → dd/MM/yyyy ; '0' o vacío → '' (sin ETA) ; texto → trim.
function _normEtaVal(v) {
  if (v instanceof Date) {
    try { return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
    catch(e) { return ''; }
  }
  var s = String(v == null ? '' : v).trim();
  if (!s || s === '0') return '';
  return s;
}

// Recién se avisa al reseller si el retraso supera esto — un corrimiento de 1-2 días del
// proveedor es normal y no vale la pena molestarlo; más de 1 semana sí es relevante.
var _SM_ETA_RETRASO_MIN_DIAS = 7;

// Parsea una ETA ya normalizada por _normEtaVal ("dd/MM/yyyy") a Date, para poder comparar
// si una fecha nueva es POSTERIOR a la anterior (retraso real, no solo un texto distinto).
function _smEtaToDate(s) {
  s = String(s || '').trim();
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return null;
  var dd = parseInt(m[1], 10), mm = parseInt(m[2], 10) - 1, yy = parseInt(m[3], 10);
  if (yy < 100) yy += 2000;
  var dt = new Date(yy, mm, dd);
  return isNaN(dt.getTime()) ? null : dt;
}

function cruzarComprasExternas() {
  try {
    var extSS    = SpreadsheetApp.openById(_PEDIDOS_EXT_SS_ID);
    var pedSheet = extSS.getSheetByName('Pedidos') || extSS.getSheetByName('Pedidos ');
    if (!pedSheet) {
      // Búsqueda tolerante a espacios
      var allSheets = extSS.getSheets();
      for (var si2 = 0; si2 < allSheets.length; si2++) {
        if (allSheets[si2].getName().trim() === 'Pedidos') { pedSheet = allSheets[si2]; break; }
      }
    }
    if (!pedSheet) return { ok: false, msg: 'Hoja "Pedidos" no encontrada en el sheet externo.' };

    var ext = pedSheet.getDataRange().getValues();

    // Encontrar fila de encabezados (primeras 10 filas)
    var hdrIdx = -1, cCas = -1, cAir = -1, cIng = -1, cEta = -1;
    for (var ri = 0; ri < Math.min(ext.length, 10) && hdrIdx < 0; ri++) {
      var foundCas = false, foundIng = false;
      for (var ci = 0; ci < ext[ri].length; ci++) {
        var v = String(ext[ri][ci] || '').trim().toLowerCase();
        if (v.indexOf('invoice') >= 0)                              { cCas = ci; foundCas = true; }
        if ((v.indexOf('n') === 0 && v.indexOf('air') >= 0) || v === 'air') cAir = ci;
        if (v.indexOf('estimada') >= 0 || v.indexOf('planta') >= 0) cEta = ci;
        if (v.indexOf('ingreso') >= 0 && v.indexOf('stock') >= 0)  { cIng = ci; foundIng = true; }
      }
      if (foundCas && foundIng) hdrIdx = ri;
    }
    if (hdrIdx < 0 || cCas < 0 || cIng < 0)
      return { ok: false, msg: 'No se encontraron columnas N\xb0 INVOICE / INGRESO A STOCK en el sheet externo.' };

    // Agrupar por CAS (N° INVOICE), parar al llegar a filas vacías
    var casMap = {};
    for (var di = hdrIdx + 1; di < ext.length; di++) {
      var casNum = String(ext[di][cCas] || '').trim().toUpperCase();
      if (!casNum) continue;
      var airNum = cAir >= 0 ? String(ext[di][cAir] || '').trim() : '';
      var ing    = String(ext[di][cIng] || '').trim().toUpperCase();
      if (!casMap[casNum]) casMap[casNum] = { cas: casNum, air: airNum, total: 0, si: 0, items: [] };
      casMap[casNum].total++;
      if (ing === 'SI') casMap[casNum].si++;
      // Cols D(3)=código, E(4)=descripción, F(5)=cantidad
      var cod  = String(ext[di][3] || '').trim();
      var desc = String(ext[di][4] || '').trim();
      var qty  = parseInt(ext[di][5], 10) || 0;
      var eta  = cEta >= 0 ? _normEtaVal(ext[di][cEta]) : '';
      if (cod && qty > 0) casMap[casNum].items.push({ codigo: cod, descripcion: desc, cantidad: qty, eta: eta, air: airNum });
    }

    // Estado actual de cada CAS en COMPRAS_DJI de SM
    var smData   = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var smCasMap = {};
    for (var si = 1; si < smData.length; si++) {
      var smCas = String(smData[si][SCHEMA.COMPRAS_DJI.ID_CAS] || '').trim().toUpperCase();
      if (smCas) smCasMap[smCas] = String(smData[si][SCHEMA.COMPRAS_DJI.ESTADO] || '');
    }

    // Items cargados en SM por CAS (COMPRAS_DETALLE)
    var smDetailByCas = {};
    try {
      var hojaDetalle = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
      if (hojaDetalle) {
        var dDet = getSheetValues(hojaDetalle);
        var CD   = SCHEMA.COMPRAS_DETALLE;
        for (var di = 1; di < dDet.length; di++) {
          var dCas = String(dDet[di][CD.ID_CAS] || '').trim().toUpperCase();
          var dSku = String(dDet[di][CD.SKU] || '').trim().toUpperCase();
          if (!dCas || !dSku) continue;
          if (!smDetailByCas[dCas]) smDetailByCas[dCas] = {};
          smDetailByCas[dCas][dSku] = {
            pedida:   parseInt(dDet[di][CD.CANTIDAD_PEDIDA])   || 0,
            recibida: parseInt(dDet[di][CD.CANTIDAD_RECIBIDA]) || 0,
            eta:      _normEtaVal(dDet[di][CD.FECHA_ETA])
          };
        }
      }
    } catch(eDet) { Logger.log('cruzarComprasExternas detalle: ' + eDet); }

    var nuevas = [], recibidas = [], diferencias = [];
    var keys = Object.keys(casMap);
    for (var ki = 0; ki < keys.length; ki++) {
      var e    = casMap[keys[ki]];
      var inSM = smCasMap.hasOwnProperty(e.cas);
      if (!inSM) {
        nuevas.push({ cas: e.cas, air: e.air, total: e.total, si: e.si, items: e.items });
        continue;
      }
      var estadoSM = smCasMap[e.cas];
      if (e.si > 0 && estadoSM !== 'En dep\xf3sito') {
        recibidas.push({ cas: e.cas, air: e.air, total: e.total, si: e.si, estadoSM: estadoSM });
      }
      // Comparar ítems para CAS que no están en depósito y tienen detalle en SM
      if (estadoSM === 'En dep\xf3sito') continue;
      var smDet = smDetailByCas[e.cas];
      if (!smDet || !e.items.length) continue;

      var extMap = {};
      for (var ei = 0; ei < e.items.length; ei++) {
        var eSku = String(e.items[ei].codigo || '').trim().toUpperCase();
        if (eSku) extMap[eSku] = { cantidad: e.items[ei].cantidad, desc: e.items[ei].descripcion, eta: e.items[ei].eta || '' };
      }

      var diffs = [];
      var allSkus = {};
      var ks = Object.keys(extMap); for (var ks0 = 0; ks0 < ks.length; ks0++) allSkus[ks[ks0]] = true;
      var ks2 = Object.keys(smDet);  for (var ks1 = 0; ks1 < ks2.length; ks1++) allSkus[ks2[ks1]] = true;

      var skuList = Object.keys(allSkus);
      for (var si2 = 0; si2 < skuList.length; si2++) {
        var sk      = skuList[si2];
        var extQty  = extMap[sk]  ? extMap[sk].cantidad      : null;
        var smQty   = smDet[sk]   ? smDet[sk].pedida         : null;
        var extDesc = extMap[sk]  ? extMap[sk].desc           : '';
        var extEta  = extMap[sk]  ? (extMap[sk].eta || '')    : '';
        var smEta   = smDet[sk]   ? (smDet[sk].eta || '')     : '';
        var qtyChg  = extQty !== smQty;
        // Cambio de ETA: solo cuando el SKU existe en ambos lados y hay al menos una fecha
        var etaChg  = !!(extMap[sk] && smDet[sk] && extEta !== smEta && (extEta || smEta));
        if (qtyChg || etaChg) {
          diffs.push({ sku: sk, desc: extDesc, ext: extQty, sm: smQty, etaExt: extEta, etaSm: smEta, qtyChg: !!qtyChg, etaChg: etaChg });
        }
      }
      if (diffs.length) diferencias.push({ cas: e.cas, estadoSM: estadoSM, air: e.air, diffs: diffs, extItems: e.items });
    }

    return { ok: true, nuevas: nuevas, recibidas: recibidas, diferencias: diferencias };
  } catch(e) {
    Logger.log('cruzarComprasExternas: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

// Reemplaza los ítems de COMPRAS_DETALLE para un CAS con los del sheet externo.
// Conserva CANTIDAD_RECIBIDA de las filas existentes que coincidan por SKU.
function sincronizarItemsCAS(cas) {
  try {
    var casKey = String(cas || '').trim().toUpperCase();
    if (!casKey) return { ok: false, error: 'CAS vacío' };

    // Leer items desde sheet externo
    var extSS    = SpreadsheetApp.openById(_PEDIDOS_EXT_SS_ID);
    var pedSheet = extSS.getSheetByName('Pedidos') || extSS.getSheetByName('Pedidos ');
    if (!pedSheet) {
      var allS = extSS.getSheets();
      for (var si = 0; si < allS.length; si++) {
        if (allS[si].getName().trim() === 'Pedidos') { pedSheet = allS[si]; break; }
      }
    }
    if (!pedSheet) return { ok: false, error: 'Hoja Pedidos no encontrada en sheet externo' };

    var ext    = pedSheet.getDataRange().getValues();
    var hdrIdx = -1, cCas = -1, cAir = -1, cEta = -1;
    for (var ri = 0; ri < Math.min(ext.length, 10) && hdrIdx < 0; ri++) {
      var hasInv = false, hasIng = false;
      for (var ci = 0; ci < ext[ri].length; ci++) {
        var v = String(ext[ri][ci] || '').trim().toLowerCase();
        if (v.indexOf('invoice') >= 0) { cCas = ci; hasInv = true; }
        if ((v.indexOf('n') === 0 && v.indexOf('air') >= 0) || v === 'air') cAir = ci;
        if (v.indexOf('estimada') >= 0 || v.indexOf('planta') >= 0) cEta = ci;
        if (v.indexOf('ingreso') >= 0 && v.indexOf('stock') >= 0) hasIng = true;
      }
      if (hasInv && hasIng) hdrIdx = ri;
    }
    if (hdrIdx < 0 || cCas < 0) return { ok: false, error: 'Columnas no encontradas en sheet externo' };

    var extItems = [];
    for (var di = hdrIdx + 1; di < ext.length; di++) {
      var extCas = String(ext[di][cCas] || '').trim().toUpperCase();
      if (extCas !== casKey) continue;
      var cod  = String(ext[di][3] || '').trim().toUpperCase();
      var desc = String(ext[di][4] || '').trim();
      var qty  = parseInt(ext[di][5], 10) || 0;
      var eta  = cEta >= 0 ? _normEtaVal(ext[di][cEta]) : '';
      var air  = cAir >= 0 ? String(ext[di][cAir] || '').trim() : '';
      if (cod && qty > 0) extItems.push({ sku: cod, desc: desc, cantidad: qty, eta: eta, air: air });
    }
    if (!extItems.length) return { ok: false, error: 'No se encontraron ítems para ' + casKey + ' en el sheet externo' };

    // Leer COMPRAS_DETALLE actual para conservar CANTIDAD_RECIBIDA
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (!hojaCD) return { ok: false, error: 'Hoja COMPRAS_DETALLE no encontrada' };
    var CD      = SCHEMA.COMPRAS_DETALLE;
    var WIDTH   = 8; // ID_CAS, SKU, DESC, PEDIDA, RECIBIDA, ESTADO, FECHA_ETA, N_AIR
    var HEADER  = ['ID_CAS','SKU','DESCRIPCION','CANTIDAD_PEDIDA','CANTIDAD_RECIBIDA','ESTADO','FECHA_ETA','N_AIR'];
    var dCD     = hojaCD.getDataRange().getValues();
    // Normaliza cualquier fila a WIDTH columnas (rellena filas viejas de 6 col con '')
    function _padRow(r) { r = r.slice(0, WIDTH); while (r.length < WIDTH) r.push(''); return r; }
    var recibMap = {}, etaAnteriorMap = {};
    var rowsKeep = [HEADER]; // encabezado canónico de 8 columnas
    for (var ri2 = 1; ri2 < dCD.length; ri2++) {
      var rCas = String(dCD[ri2][CD.ID_CAS] || '').trim().toUpperCase();
      if (rCas === casKey) {
        // guardar recibido y ETA previa por SKU, descartar la fila (se reemplaza)
        var rSku = String(dCD[ri2][CD.SKU] || '').trim().toUpperCase();
        recibMap[rSku]      = parseInt(dCD[ri2][CD.CANTIDAD_RECIBIDA]) || 0;
        etaAnteriorMap[rSku] = _normEtaVal(dCD[ri2][CD.FECHA_ETA]);
      } else {
        rowsKeep.push(_padRow(dCD[ri2])); // otras CAS: conservar, normalizadas a 8 col
      }
    }

    // Agregar nuevas filas con datos del sheet externo (incluye ETA + N° AIR por línea).
    // De paso, detectar RETRASOS reales (ETA nueva posterior a la anterior) para avisar a
    // los resellers que tengan una reserva activa de ese CAS/SKU — ver _encolarNotifEtaCambioWOS.
    var retrasos = [];
    for (var xi = 0; xi < extItems.length; xi++) {
      var xIt  = extItems[xi];
      var xRec = recibMap[xIt.sku] || 0;
      var xEst = xRec >= xIt.cantidad ? 'Recibido' : (xRec > 0 ? 'Parcial' : 'Pendiente');
      rowsKeep.push([casKey, xIt.sku, xIt.desc, xIt.cantidad, xRec, xEst, xIt.eta || '', xIt.air || '']);

      var etaAnt = etaAnteriorMap[xIt.sku] || '';
      if (etaAnt && xIt.eta && etaAnt !== xIt.eta) {
        var dAnt = _smEtaToDate(etaAnt), dNva = _smEtaToDate(xIt.eta);
        if (dAnt && dNva) {
          var diasRetraso = (dNva.getTime() - dAnt.getTime()) / 86400000;
          // Solo avisar si el retraso supera _SM_ETA_RETRASO_MIN_DIAS (corrimientos chicos del
          // proveedor, 1-2 días, son normales y no ameritan un mail al reseller).
          if (diasRetraso > _SM_ETA_RETRASO_MIN_DIAS) {
            retrasos.push({ sku: xIt.sku, desc: xIt.desc, etaAnterior: etaAnt, etaNueva: xIt.eta });
          }
        }
      }
    }

    hojaCD.clearContents();
    hojaCD.getRange(1, 1, rowsKeep.length, WIDTH).setValues(rowsKeep);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    SpreadsheetApp.flush();

    var notificaciones = 0;
    if (retrasos.length) {
      try { notificaciones = _encolarNotifEtaCambioWOS(casKey, retrasos) || 0; }
      catch(eNot) { Logger.log('sincronizarItemsCAS notif ETA: ' + eNot); }
    }
    return { ok: true, cas: casKey, items: extItems.length, retrasosSku: retrasos.length, notificaciones: notificaciones };
  } catch(e) {
    Logger.log('sincronizarItemsCAS: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ============================================================
//  DETALLE DE ITEMS POR CAS
// ============================================================
function obtenerItemsPorCAS(idCas) {
  try {
    var casB = String(idCas).trim().toUpperCase();
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var R    = SCHEMA.RESERVAS_STOCK;
    var tz   = Session.getScriptTimeZone();

    var dMov     = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var recibidos = [];
    for (var i = 1; i < dMov.length; i++) {
      var f = dMov[i];
      if (String(f[M.TIPO] || '').trim() !== 'ENTRADA_COMPRA') continue;
      if (String(f[M.REFERENCIA] || '').trim().toUpperCase() !== casB) continue;
      recibidos.push({
        sku:         String(f[M.CODIGO]      || ''),
        descripcion: String(f[M.DESCRIPCION] || ''),
        cantidad:    Math.abs(parseInt(f[M.CANTIDAD]) || 0),
        fecha:       f[M.FECHA] instanceof Date
                     ? Utilities.formatDate(f[M.FECHA], tz, 'dd/MM/yyyy')
                     : String(f[M.FECHA] || '')
      });
    }

    var reservados = [];
    var hojaRes = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (hojaRes) {
      var dRes = getSheetValues(hojaRes);
      for (var j = 1; j < dRes.length; j++) {
        var r = dRes[j];
        if (String(r[R.CAS_REF] || '').trim().toUpperCase() !== casB) continue;
        if (String(r[R.ESTADO]  || '') !== 'Activa') continue;
        reservados.push({
          sku:        String(r[R.SKU]          || ''),
          descripcion:String(r[R.DESCRIPCION]  || ''),
          cantidad:   parseInt(r[R.CANTIDAD])  || 0,
          origen:     String(r[R.ORIGEN]       || ''),
          referencia: String(r[R.ID_REFERENCIA]|| '')
        });
      }
    }

    return { cas: idCas, recibidos: recibidos, reservados: reservados };
  } catch(e) {
    return { cas: idCas, recibidos: [], reservados: [], error: e.toString() };
  }
}

// ============================================================
//  COMPRAS DETALLE — manifiesto de ítems por CAS
// ============================================================

function crearHojaComprasDetalle() {
  try {
    var db = getDb();
    var existing = db.getSheetByName(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (existing) return { ok: false, msg: 'La hoja ya existe' };
    var hoja = db.insertSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    hoja.appendRow(['ID_CAS','SKU','DESCRIPCION','CANTIDAD_PEDIDA','CANTIDAD_RECIBIDA','ESTADO','FECHA_ETA','N_AIR']);
    hoja.setFrozenRows(1);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function vincularItemsACAS(idCas, listaItems) {
  // listaItems: [{sku, descripcion, cantPedida}]
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (!hoja) return { ok: false, msg: 'Hoja COMPRAS_DETALLE no existe. Ejecutá crearHojaComprasDetalle() primero.' };
    var casKey = String(idCas).trim().toUpperCase();
    var CD = SCHEMA.COMPRAS_DETALLE;

    // Full replace: delete existing rows for this CAS (iterate backwards)
    var d = hoja.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--) {
      if (String(d[i][CD.ID_CAS]).trim().toUpperCase() === casKey) {
        hoja.deleteRow(i + 1);
      }
    }

    if (listaItems.length) {
      var newRows = [];
      for (var k = 0; k < listaItems.length; k++) {
        var item = listaItems[k];
        newRows.push([
          casKey,
          String(item.sku || '').trim().toUpperCase(),
          String(item.descripcion || ''),
          parseInt(item.cantPedida) || 0,
          0,
          'Pendiente'
        ]);
      }
      var startRow = hoja.getLastRow() + 1;
      hoja.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
      SpreadsheetApp.flush();
    }

    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    return { ok: true, items: listaItems.length };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function obtenerDetalleCAS(idCas) {
  try {
    var casKey = String(idCas).trim().toUpperCase();
    var CD = SCHEMA.COMPRAS_DETALLE;
    var M  = SCHEMA.MOVIMIENTOS_STOCK;
    var R  = SCHEMA.RESERVAS_STOCK;
    var tz = Session.getScriptTimeZone();

    // 1. Manifest
    var manifest = [];
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (hojaCD) {
      var dCD = getSheetValues(hojaCD);
      for (var ci = 1; ci < dCD.length; ci++) {
        if (String(dCD[ci][CD.ID_CAS]).trim().toUpperCase() !== casKey) continue;
        manifest.push({
          sku:          String(dCD[ci][CD.SKU]               || ''),
          descripcion:  String(dCD[ci][CD.DESCRIPCION]       || ''),
          cantPedida:   parseInt(dCD[ci][CD.CANTIDAD_PEDIDA])   || 0,
          cantRecibida: parseInt(dCD[ci][CD.CANTIDAD_RECIBIDA]) || 0,
          estado:       String(dCD[ci][CD.ESTADO]            || 'Pendiente')
        });
      }
    }

    // 2. Received (from MOVIMIENTOS ENTRADA_COMPRA)
    var recibidos = [];
    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    for (var mi = 1; mi < dMov.length; mi++) {
      if (String(dMov[mi][M.TIPO]       || '').trim() !== 'ENTRADA_COMPRA') continue;
      if (String(dMov[mi][M.REFERENCIA] || '').trim().toUpperCase() !== casKey) continue;
      recibidos.push({
        sku:         String(dMov[mi][M.CODIGO]      || ''),
        descripcion: String(dMov[mi][M.DESCRIPCION] || ''),
        cantidad:    Math.abs(parseInt(dMov[mi][M.CANTIDAD]) || 0),
        fecha:       dMov[mi][M.FECHA] instanceof Date
                     ? Utilities.formatDate(dMov[mi][M.FECHA], tz, 'dd/MM/yyyy')
                     : String(dMov[mi][M.FECHA] || '')
      });
    }

    // 3. Active reservations
    var reservados = [];
    var hojaRes = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (hojaRes) {
      var dRes = getSheetValues(hojaRes);
      for (var ri = 1; ri < dRes.length; ri++) {
        if (String(dRes[ri][R.CAS_REF] || '').trim().toUpperCase() !== casKey) continue;
        if (String(dRes[ri][R.ESTADO]  || '') !== 'Activa') continue;
        reservados.push({
          sku:        String(dRes[ri][R.SKU]          || ''),
          descripcion:String(dRes[ri][R.DESCRIPCION]  || ''),
          cantidad:   parseInt(dRes[ri][R.CANTIDAD])  || 0,
          origen:     String(dRes[ri][R.ORIGEN]       || ''),
          referencia: String(dRes[ri][R.ID_REFERENCIA]|| '')
        });
      }
    }

    // 4. Coverage: para cada SKU del manifiesto, ¿cuántas unidades quedan PENDIENTES de
    //    entregar a resellers? Fuente = Pedidos_resellers (WOS) — la demanda real del canal.
    //    Pendiente por línea = max(0, CANT_SOL − CANT_DESP − CANT_CANCEL): descuenta lo ya
    //    despachado y lo cancelado (una línea cerrada o anulada aporta 0). Misma fórmula que el
    //    mail de "Backorders desbloqueados". Antes se leía SOLICITUDES_DESPACHO (solo repuestos de
    //    OT + ventas directas), que NO incluye los pedidos de resellers → cobertura equivocada.
    var cobertura = [];
    if (manifest.length) {
      var pendMap = {};
      try {
        var hojaPR = SpreadsheetApp.openById(WOS_NOTAS_SS_ID).getSheetByName('Pedidos_resellers');
        if (hojaPR) {
          var dPR = hojaPR.getDataRange().getValues();
          // COL (Despacho_Env.js): 2=SKU, 4=CANT_SOL(E), 5=CANT_DESP(F), 9=ESTADO(J), 25=CANT_CANCEL(Z)
          for (var pr = 1; pr < dPR.length; pr++) {
            var skuPR = String(dPR[pr][2] || '').trim().toUpperCase();
            if (!skuPR) continue;
            var solPR  = Number(dPR[pr][4])  || 0;
            var despPR = Number(dPR[pr][5])  || 0;
            var cancPR = Number(dPR[pr][25]) || 0;
            if (cancPR <= 0 && String(dPR[pr][9] || '').trim().toLowerCase() === 'cancelado') cancPR = solPR; // dato viejo: cancelado sin CANT_CANCEL
            var pendPR = Math.max(0, solPR - despPR - cancPR);
            if (pendPR > 0) pendMap[skuPR] = (pendMap[skuPR] || 0) + pendPR;
          }
        }
      } catch(ePR) { Logger.log('obtenerDetalleCAS cobertura Pedidos_resellers: ' + ePR); }
      for (var ci2 = 0; ci2 < manifest.length; ci2++) {
        var skuUp  = manifest[ci2].sku.trim().toUpperCase();
        var pending = pendMap[skuUp] || 0;
        var enCamino = manifest[ci2].cantPedida - manifest[ci2].cantRecibida;
        cobertura.push({ sku: manifest[ci2].sku, pedido: pending, enCamino: enCamino, cubre: enCamino >= pending });
      }
    }

    return { cas: idCas, manifest: manifest, recibidos: recibidos, reservados: reservados, cobertura: cobertura };
  } catch(e) {
    return { cas: idCas, manifest: [], recibidos: [], reservados: [], cobertura: [], error: e.toString() };
  }
}

// ============================================================
//  ANÁLISIS DE VELOCIDAD DE CONSUMO
// ============================================================
function analizarVelocidadConsumo(leadDias) {
  try {
    if (!leadDias) leadDias = 45;
    var hoy    = new Date();
    var corte  = new Date(hoy.getTime() - 30 * 86400000);
    var dMov   = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var dStock = getSheetValues(SCHEMA.SHEETS.STOCK);
    var M = SCHEMA.MOVIMIENTOS_STOCK;
    var S = SCHEMA.STOCK_REPUESTOS;

    var consumo30 = {};
    for (var m = 1; m < dMov.length; m++) {
      var fm = dMov[m];
      if (!(fm[M.FECHA] instanceof Date) || fm[M.FECHA] < corte) continue;
      var tipo = String(fm[M.TIPO] || '');
      if (tipo !== 'SALIDA_DESPACHO' && tipo !== 'EGRESO') continue;
      var cod = String(fm[M.CODIGO] || '').trim().toUpperCase();
      if (!cod) continue;
      consumo30[cod] = (consumo30[cod] || 0) + Math.abs(parseInt(fm[M.CANTIDAD]) || 0);
    }

    var out = [];
    for (var s = 1; s < dStock.length; s++) {
      if (!dStock[s][S.CODIGO]) continue;
      var codS    = String(dStock[s][S.CODIGO]).trim().toUpperCase();
      var actS    = parseInt(dStock[s][S.STOCK_ACTUAL]) || 0;
      var unids   = consumo30[codS] || 0;
      var burnDay = unids / 30;
      var diasR   = burnDay > 0 ? Math.round(actS / burnDay) : null;
      var flagged = diasR !== null && diasR < leadDias && actS > 0;
      out.push({
        codigo:        codS,
        descripcion:   String(dStock[s][S.DESCRIPCION] || ''),
        stockActual:   actS,
        unidades30d:   unids,
        burnRatePerDay:Math.round(burnDay * 100) / 100,
        diasRestantes: diasR,
        sugeridoCompra:flagged
      });
    }
    out.sort(function(a, b) {
      if (a.diasRestantes === null && b.diasRestantes === null) return 0;
      if (a.diasRestantes === null) return 1;
      if (b.diasRestantes === null) return -1;
      return a.diasRestantes - b.diasRestantes;
    });
    return out;
  } catch(e) { return []; }
}

// ============================================================
//  LOOP AUTOMÁTICO — BORRADOR DE COMPRA POR BACKORDER
// ============================================================
function crearBorradorCompraAutomatico(operador) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!operador) {
      try { operador = Session.getActiveUser().getEmail(); } catch(eu) { operador = 'Sistema'; }
    }

    var alertas = obtenerAlertasStockCritico();
    // Solo quiebres con OTs bloqueadas
    var candidatos = [];
    for (var a = 0; a < alertas.length; a++) {
      if (alertas[a].estado === 'QUIEBRE' && alertas[a].bloqueadas > 0) {
        candidatos.push(alertas[a]);
      }
    }
    if (!candidatos.length) return { ok: false, msg: 'No hay SKUs en quiebre con OTs bloqueadas.' };

    var tz  = Session.getScriptTimeZone();
    var hoy = new Date();
    var cas = 'BORRADOR-AUTO-' + Utilities.formatDate(hoy, tz, 'yyyyMMdd-HHmm');

    var hojaCompras = getSheet(SCHEMA.SHEETS.COMPRAS);
    var dComp = getSheetValues(hojaCompras);
    // Verificar duplicado del mismo día (evita doble click)
    var prefixHoy = 'BORRADOR-AUTO-' + Utilities.formatDate(hoy, tz, 'yyyyMMdd');
    for (var ci = 1; ci < dComp.length; ci++) {
      var casExist = String(dComp[ci][0] || '').trim().toUpperCase();
      if (casExist.indexOf(prefixHoy.toUpperCase()) === 0) {
        return { ok: false, msg: 'Ya existe un borrador automático para hoy: ' + casExist + '. Eliminalo o editalo desde la sección Compras.' };
      }
    }

    // Crear fila en COMPRAS_DJI con estado "Borrador"
    var obsItems = candidatos.map(function(c) { return c.codigo + ' (' + c.bloqueadas + ' OTs)'; }).join(', ');
    hojaCompras.appendRow([
      cas.toUpperCase(),
      hoy,
      'Borrador',
      '',
      '', '', '', '', '', '', '',
      operador || 'Sistema',
      'Auto-generado. SKUs bloqueados: ' + obsItems
    ]);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);

    // Escribir ítems en COMPRAS_DETALLE
    var listaItems = candidatos.map(function(c) {
      return {
        sku:        c.codigo,
        descripcion:c.descripcion,
        cantPedida: c.bloqueadas  // 1 unidad por OT bloqueada como mínimo
      };
    });
    var resVinc = vincularItemsACAS(cas.toUpperCase(), listaItems);
    if (!resVinc.ok) {
      Logger.log('crearBorradorCompraAutomatico: vincularItemsACAS falló: ' + resVinc.msg);
    }

    // Notificar al supervisor
    _notificarBorradorAutoCompra(cas.toUpperCase(), candidatos, operador);

    return { ok: true, cas: cas.toUpperCase(), items: candidatos.length };
  } catch(e) {
    Logger.log('crearBorradorCompraAutomatico: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

function _notificarBorradorAutoCompra(cas, candidatos, operador) {
  try {
    var filas = candidatos.map(function(c) {
      return "<tr>" +
        "<td style='padding:6px 10px;font-size:12px;font-weight:700;color:#e53935'>" + c.codigo + "</td>" +
        "<td style='padding:6px 10px;font-size:12px'>" + c.descripcion + "</td>" +
        "<td style='padding:6px 10px;font-size:12px;text-align:center'>" + c.bloqueadas + "</td>" +
      "</tr>";
    }).join('');
    GmailApp.sendEmail(SM_CONFIG.EMAIL_SUPERVISOR,
      '[Stock Manager] Borrador de compra generado — ' + candidatos.length + ' SKU(s) bloqueados',
      '', {
        htmlBody:
          "<div style='font-family:sans-serif;max-width:640px'>" +
          "<div style='background:#e53935;padding:16px 20px;border-radius:8px 8px 0 0'>" +
            "<span style='color:#fff;font-size:16px;font-weight:700'>Borrador de compra auto-generado</span>" +
          "</div>" +
          "<div style='background:#fff;border:1px solid #ddd;padding:18px 20px;border-radius:0 0 8px 8px'>" +
            "<p style='font-size:13px;color:#444;margin:0 0 6px'>CAS: <strong>" + cas + "</strong></p>" +
            "<p style='font-size:13px;color:#444;margin:0 0 14px'>Operador: " + (operador || 'Sistema') + "</p>" +
            "<p style='font-size:13px;color:#444;margin:0 0 10px'>Los siguientes SKUs están en <strong>quiebre total</strong> y tienen OTs esperando esos repuestos:</p>" +
            "<table style='width:100%;border-collapse:collapse;border:1px solid #eee'>" +
              "<thead><tr style='background:#f5f5f5'>" +
                "<th style='padding:6px 10px;font-size:11px;text-align:left'>SKU</th>" +
                "<th style='padding:6px 10px;font-size:11px;text-align:left'>Descripción</th>" +
                "<th style='padding:6px 10px;font-size:11px;text-align:center'>OTs bloqueadas</th>" +
              "</tr></thead>" +
              "<tbody>" + filas + "</tbody>" +
            "</table>" +
            "<p style='font-size:12px;color:#666;margin-top:14px'>Ingresá al <strong>Stock Manager → Compras</strong> para revisar y confirmar el pedido.</p>" +
          "</div></div>",
        name: SM_CONFIG.NOMBRE_REMITENTE,
        replyTo: SM_CONFIG.EMAIL_SUPERVISOR
      });
  } catch(e) { Logger.log('_notificarBorradorAutoCompra: ' + e); }
}

// ============================================================
//  SINCRONIZAR CATÁLOGO DJI → DB_REPUESTOS
// ============================================================
function sincronizarCatalogoDJI(operador) {
  try {
    var hojaCat = getSheet(SCHEMA.SHEETS.CATALOGO_DJI);
    if (!hojaCat) return { ok: false, msg: 'Hoja CATALOGO_DJI no encontrada. Creala primero con asegurarHojas().' };
    var dCat = getSheetValues(hojaCat);
    if (dCat.length < 2) return { ok: false, msg: 'El catálogo está vacío — pegá la lista DJI debajo de los encabezados.' };

    // Detectar columnas por nombre de header (case-insensitive, acepta variantes)
    var headers = dCat[0].map(function(h) { return String(h||'').trim().toUpperCase().replace(/[\s\-_]+/g, ''); });
    function findCol() {
      var aliases = Array.prototype.slice.call(arguments);
      for (var a = 0; a < aliases.length; a++) {
        var idx = headers.indexOf(aliases[a].toUpperCase().replace(/[\s\-_]+/g, ''));
        if (idx !== -1) return idx;
      }
      return -1;
    }
    var COL = {
      CODIGO:      findCol('SIMPLIFIEDPARTNUMBER','SIMPLIFIED PART NUMBER','CODIGO','SKU','PARTNUMBER','PART NUMBER','CODIGOREPUESTO'),
      DESCRIPCION: findCol('ENGLISHNAME','ENGLISH NAME','DESCRIPCION','DESCRIPTION','NOMBRE','NAME','DESC'),
      MODELOS:     findCol('APPLICABLEMODELS','APPLICABLE MODELS','MODELOS','MODELS','COMPATIBLE','MODELOCOMPATIBLE'),
      PRECIO_FOB:  findCol('SOUTHAMERICA','SOUTH AMERICA','PRECIOFOB','PRECIO FOB','PRECIO','PRICE','FOB','PRECIOUSD'),
      SKU_DJI:     findCol('MATERIALNUMBER','MATERIAL NUMBER','SKUDJI','SKU DJI','CODIGODJI','PARTNODJI'),
      CATEGORIA:   findCol('CATEGORIA','CATEGORY','CAT'),
      PESO:        findCol('PESO','WEIGHT','PESOG','WEIGHTG')
    };
    if (COL.CODIGO === -1) return { ok: false, msg: 'No se encontró columna "Simplified part number" en la primera fila del catálogo.' };

    // Construir mapa del catálogo
    var catalog = {};
    for (var i = 1; i < dCat.length; i++) {
      var row = dCat[i];
      var cod = COL.CODIGO !== -1 ? String(row[COL.CODIGO]||'').trim().toUpperCase() : '';
      if (!cod) continue;
      catalog[cod] = {
        descripcion: COL.DESCRIPCION !== -1 ? String(row[COL.DESCRIPCION]||'').trim() : '',
        categoria:   COL.CATEGORIA   !== -1 ? String(row[COL.CATEGORIA]  ||'').trim() : '',
        modelos:     COL.MODELOS     !== -1 ? String(row[COL.MODELOS]    ||'').trim() : '',
        precioFOB:   COL.PRECIO_FOB  !== -1 ? (parseFloat(String(row[COL.PRECIO_FOB]||'0').replace(/[^0-9.,\-]/g,'').replace(',','.')) || 0) : 0,
        skuDJI:      COL.SKU_DJI     !== -1 ? String(row[COL.SKU_DJI]   ||'').trim() : ''
      };
    }

    var totalCatalog = Object.keys(catalog).length;
    if (totalCatalog === 0) return { ok: false, msg: 'No se encontraron filas con código válido en el catálogo.' };

    // Upsert en DB_REPUESTOS
    // Schema DB_REPUESTOS: [0]ID [1]SKU [2]DESC [3]MODELOS [4]? [5]CATEGORIA [6]PRECIO_FOB
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (!hojaRep) return { ok: false, msg: 'Hoja DB_REPUESTOS no encontrada.' };
    var dRep = hojaRep.getDataRange().getValues();
    var repIdx = {};
    for (var r = 1; r < dRep.length; r++) {
      var skuR = String(dRep[r][1]||'').trim().toUpperCase();
      if (skuR) repIdx[skuR] = r;
    }

    var nuevos = 0, actualizados = 0, appendRows = [];
    var codsOrdenados = Object.keys(catalog);
    for (var k = 0; k < codsOrdenados.length; k++) {
      var cod2 = codsOrdenados[k];
      var c    = catalog[cod2];
      if (repIdx[cod2] !== undefined) {
        var ri = repIdx[cod2];
        var changed = false;
        // Siempre sobreescribir desde el catálogo (corrige datos previos incorrectos)
        if (c.descripcion && dRep[ri][2] !== c.descripcion) { dRep[ri][2] = c.descripcion; changed = true; }
        if (c.modelos     && dRep[ri][3] !== c.modelos)     { dRep[ri][3] = c.modelos;     changed = true; }
        if (c.precioFOB   && dRep[ri][6] !== c.precioFOB)   { dRep[ri][6] = c.precioFOB;   changed = true; }
        // Categoría: solo si es texto válido (no número) — previene que valores numéricos del catálogo DJI contaminen este campo
        if (c.categoria && _catValida(c.categoria) && dRep[ri][5] !== c.categoria) { dRep[ri][5] = c.categoria; changed = true; }
        // Limpiar categoría si actualmente tiene un número inválido
        if (dRep[ri][5] && !_catValida(String(dRep[ri][5]))) { dRep[ri][5] = ''; changed = true; }
        if (changed) actualizados++;
      } else {
        appendRows.push(['', cod2, c.descripcion, c.modelos, '', _catValida(c.categoria) ? c.categoria : '', c.precioFOB]);
        nuevos++;
      }
    }
    if (actualizados > 0) hojaRep.getDataRange().setValues(dRep);
    for (var na = 0; na < appendRows.length; na++) hojaRep.appendRow(appendRows[na]);
    if (nuevos > 0 || actualizados > 0) invalidateSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);

    // Actualizar STOCK_REPUESTOS: sobreescribir modelos y corregir categorías numéricas
    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    var dStock    = hojaStock.getDataRange().getValues();
    var stockChanged = false;
    for (var s = 1; s < dStock.length; s++) {
      var skuS = String(dStock[s][0]||'').trim().toUpperCase();
      if (!skuS) continue;
      // Limpiar categorías numéricas aunque el SKU no esté en el catálogo
      if (dStock[s][4] && !_catValida(String(dStock[s][4]))) { dStock[s][4] = ''; stockChanged = true; }
      if (!catalog[skuS]) continue;
      var cat = catalog[skuS];
      if (cat.modelos && dStock[s][6] !== cat.modelos)                         { dStock[s][6] = cat.modelos;   stockChanged = true; }
      if (cat.categoria && _catValida(cat.categoria) && !dStock[s][4])         { dStock[s][4] = cat.categoria; stockChanged = true; }
    }
    if (stockChanged) {
      hojaStock.getDataRange().setValues(dStock);
      invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    }

    return { ok: true, total: totalCatalog, nuevos: nuevos, actualizados: actualizados };
  } catch(e) {
    Logger.log('sincronizarCatalogoDJI: ' + e);
    return { ok: false, msg: e.toString() };
  }
}
