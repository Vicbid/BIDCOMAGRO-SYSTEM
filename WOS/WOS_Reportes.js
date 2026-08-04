// @version 1.2
// ============================================================
//  WOS — Reportes: resumen de envíos por reseller/mes + reporte
//  de backorder (demanda perdida) por mail + su trigger.
//  Extraído de Despacho_Código.js 3.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


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

    var _asuntoResumen = 'Resumen de envíos — ' + reseller + ' — ' + res.mesLabel;
    GmailApp.sendEmail(email, _asuntoResumen, plainBody, {
      htmlBody: htmlBody,
      name:     'BIDCOMAGRO · Portal Resellers',
      replyTo:  _wosConfig().emailSoporte,
      cc:       _wosConfig().emailFact
    });
    _wosRegistrarEmailLog(reseller + ' · ' + res.mesLabel, email, 'Resumen de envíos', _asuntoResumen, 'OK', '');

    Logger.log('WOS_enviarResumenEnvios OK: ' + reseller + ' ' + mesAnio + ' (' + res.envios.length + ' envíos)');
    return { ok: true, enviados: res.envios.length };
  } catch(e) {
    Logger.log('WOS_enviarResumenEnvios ERROR: ' + e);
    _wosRegistrarEmailLog(reseller + ' · ' + mesAnio, (typeof email !== 'undefined' ? email : ''), 'Resumen de envíos', 'Resumen de envíos — ' + reseller, 'ERROR: ' + String(e).substring(0, 150), '');
    return { ok: false, error: e.toString() };
  }
 });
}


// Lista de resellers con email válido cargado en la hoja Resellers del MASTER — misma fuente
// que _wosGetEmailReseller (WOS_GmailFlow.js), leída una sola vez acá para poder iterar TODOS
// los resellers (antes solo existía "elegir 1 reseller" en el modal de Resumen de Envíos).
function _wosListaResellersConEmail() {
  try {
    var d = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Resellers').getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      var nombre = String(d[i][COL_RS.NOMBRE] || '').trim();
      var email  = String(d[i][COL_RS.EMAIL]  || '').trim();
      if (nombre && email && email.indexOf('@') !== -1) out.push({ nombre: nombre, email: email });
    }
    return out;
  } catch(e) { Logger.log('_wosListaResellersConEmail: ' + e); return []; }
}

// ── Resumen de envíos EN LOTE — manda el mail de todos los resellers de una, en vez de
// 1 por 1. Mismo patrón que WOS_despacharBatch (WOS_Stock.js): loop con try/catch por ítem
// (un reseller con error no aborta el resto) + token de idempotencia derivado por ítem.
// Reusa WOS_enviarResumenEnvios tal cual (mismo mail, mismo log) — acá solo se itera.
function WOS_enviarResumenEnviosLote(mesAnio, reqToken) {
  try {
    if (!mesAnio) return { ok: false, error: 'Falta el mes.' };
    var resellers = _wosListaResellersConEmail();
    if (!resellers.length) return { ok: false, error: 'No se encontraron resellers con email cargado.' };

    var resultados = [];
    for (var i = 0; i < resellers.length; i++) {
      var nombre = resellers[i].nombre;
      try {
        var itemToken = reqToken ? (reqToken + '::' + nombre) : '';
        var res = WOS_enviarResumenEnvios(nombre, mesAnio, itemToken);
        if (res && res.ok) {
          resultados.push({ reseller: nombre, ok: true, enviados: res.enviados });
        } else {
          // "No hay envíos para ese reseller y mes" no es un error real, es la mayoría de los
          // casos (no todos los resellers despacharon ese mes) — se cuenta aparte.
          var sinEnvios = res && /No hay env/i.test(res.error || '');
          resultados.push({ reseller: nombre, ok: false, skip: !!sinEnvios, error: (res && res.error) || 'Error desconocido' });
        }
      } catch(eI) {
        resultados.push({ reseller: nombre, ok: false, skip: false, error: eI.toString() });
      }
    }
    var enviados = 0, saltados = 0, errores = 0;
    for (var j = 0; j < resultados.length; j++) {
      if (resultados[j].ok) enviados++;
      else if (resultados[j].skip) saltados++;
      else errores++;
    }
    Logger.log('WOS_enviarResumenEnviosLote: ' + mesAnio + ' → ' + enviados + ' enviados, ' + saltados + ' sin envíos, ' + errores + ' con error');
    return { ok: true, resultados: resultados, enviados: enviados, saltados: saltados, errores: errores };
  } catch(e) {
    Logger.log('WOS_enviarResumenEnviosLote ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Excel descargable con el resumen de envíos de TODOS los resellers de un mes ──
// Mismo patrón que CC_exportarXLS (PORTAL_RESELLER/RS_CuentaCorriente.js): crea una
// spreadsheet temporal, la exporta como blob XLSX, la borra y devuelve el base64 para que
// el cliente arme la descarga con un data: URI (sin re-auth: WOS ya usa DriveApp/SpreadsheetApp
// para los PDFs de nota de entrega, mismo scope ya concedido).
function WOS_exportarResumenEnviosXLS(mesAnio) {
  var ssTemp = null;
  try {
    if (!mesAnio) return { ok: false, error: 'Falta el mes.' };
    var resellers = _wosListaResellersConEmail();
    var filas = [];
    var totalCosto = 0, totalPeso = 0;
    var mesLabel = mesAnio;

    for (var i = 0; i < resellers.length; i++) {
      var res = WOS_getResumenEnvios(resellers[i].nombre, mesAnio);
      if (!res || !res.ok || !res.envios.length) continue;
      mesLabel = res.mesLabel || mesLabel;
      for (var j = 0; j < res.envios.length; j++) {
        var ev = res.envios[j];
        filas.push([
          resellers[i].nombre, ev.fecha, ev.numero, ev.notaEntrega,
          ev.transportista || '', ev.tracking || '', ev.pesoEnvio || 0, ev.costoEnvio || 0
        ]);
        totalCosto += Number(ev.costoEnvio) || 0;
        totalPeso  += Number(ev.pesoEnvio)  || 0;
      }
    }
    if (!filas.length) return { ok: false, error: 'No hay envíos registrados para ese mes.' };

    ssTemp = SpreadsheetApp.create('WOS_resumen_envios_' + new Date().getTime());
    var hoja = ssTemp.getActiveSheet();
    hoja.setName('Resumen Envíos');
    hoja.appendRow(['Reseller', 'Fecha', 'Pedido', 'Nota de Entrega', 'Transportista', 'Tracking', 'Peso (kg)', 'Costo Envío']);
    for (var k = 0; k < filas.length; k++) hoja.appendRow(filas[k]);
    hoja.appendRow([]);
    hoja.appendRow(['', '', '', '', '', 'TOTAL', Math.round(totalPeso * 100) / 100, totalCosto]);
    hoja.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#ffffff');
    hoja.autoResizeColumns(1, 8);

    var file   = DriveApp.getFileById(ssTemp.getId());
    var blob   = file.getAs(MimeType.MICROSOFT_EXCEL);
    var base64 = Utilities.base64Encode(blob.getBytes());
    file.setTrashed(true);
    ssTemp = null;

    return { ok: true, base64: base64, nombre: 'ResumenEnvios_' + String(mesLabel).replace(/\s+/g, '_') + '.xlsx' };
  } catch(e) {
    Logger.log('WOS_exportarResumenEnviosXLS ERROR: ' + e);
    if (ssTemp) { try { DriveApp.getFileById(ssTemp.getId()).setTrashed(true); } catch(eT) {} }
    return { ok: false, error: e.toString() };
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
    _wosRegistrarEmailLog('REPORTE_BACKORDER', destinatarios.join(', '), 'Reporte backorder', asunto, 'OK', '');
    Logger.log('WOS_reporteBackorder enviado a: ' + destinatarios.join(', ') + ' | sin cobertura: ' + sinCubrir.length + ', cubiertos: ' + cubiertos.length);
  } catch(e) {
    Logger.log('WOS_reporteBackorder ERROR: ' + e);
    _wosRegistrarEmailLog('REPORTE_BACKORDER', '', 'Reporte backorder', 'Backorder WOS', 'ERROR: ' + String(e).substring(0, 150), '');
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
