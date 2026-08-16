// @version 1.3
// ============================================================
//  HUB PRO — Solicitud de factura: descuento del reseller, XLS de
//  detalle, mail a administración/facturación.
//  Extraído de HUB_Código.js 2.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// ============================================================
//  FACTURACIÓN — FLUJO COMPLETO
//  1. getMailAdministracion()        → busca email en Usuarios_Internos
//  2. _obtenerDescuentoReseller()    → lee descuento por reseller (fallback 40%)
//  3. _generarXLSFactura()           → crea XLSX temporal en Drive y lo devuelve como Blob
//  4. _construirEmailFacturacion()   → arma el HTML del cuerpo
//  5. solicitarFactura(data)         → punto de entrada desde el frontend
// ============================================================

function getMailAdministracion() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.USUARIOS);
    if (!hoja) return null;
    var datos = getSheetValues(hoja);
    for (var i = 1; i < datos.length; i++) {
      var nombre = String(datos[i][0]||'').trim().toLowerCase();
      var email  = String(datos[i][1]||'').trim();
      if (nombre === 'administracion' || nombre === 'administración') {
        return (email && email.indexOf('@') !== -1) ? email : null;
      }
    }
    return null;
  } catch(e) {
    Logger.log('getMailAdministracion: ' + e);
    return null;
  }
}


function _obtenerDescuentoReseller(nombreReseller) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    if (!datos || datos.length < 2) return 0.40;
    var header = datos[0].map(function(c){ return String(c||'').toLowerCase().trim(); });
    var colDesc = -1;
    for (var j = 0; j < header.length; j++) {
      if (header[j].indexOf('descuento') !== -1 || header[j].indexOf('discount') !== -1) {
        colDesc = j; break;
      }
    }
    if (colDesc < 0) return 0.40;
    var nb = String(nombreReseller||'').trim().toLowerCase();
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][0]||'').trim().toLowerCase() === nb) {
        var val = parseFloat(String(datos[i][colDesc]||'').replace(',','.'));
        if (!isNaN(val) && val > 0 && val <= 1) return val;
      }
    }
    return 0.40;
  } catch(e) { return 0.40; }
}


function _generarXLSFactura(ot, items, totalGeneral, data) {
  var nombre = 'FACTURA_' + ot + '_' + new Date().getTime();
  var tmpSS = SpreadsheetApp.create(nombre);
  var ws = tmpSS.getSheets()[0];
  ws.setName('Detalle');

  var headers = ['CÓDIGO', 'DESCRIPCIÓN', 'CANTIDAD', 'PVP UNITARIO (USD)', 'TOTAL ITEM (USD)'];
  ws.getRange(1, 1, 1, 5).setValues([headers])
    .setFontWeight('bold')
    .setBackground('#00a3e0')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  if (items.length > 0) {
    var rows = items.map(function(item) {
      return [
        _antiFormula(item.codigo),
        _antiFormula(item.descripcion),
        item.cantidad,
        item.pvp > 0 ? item.pvp : '',
        item.total > 0 ? item.total : ''
      ];
    });
    var rango = ws.getRange(2, 1, rows.length, 5);
    rango.setValues(rows);
    ws.getRange(2, 4, rows.length, 2).setNumberFormat('[$$-es-AR]#,##0.00');
  }

  var filaTotal = items.length + 3;
  ws.getRange(filaTotal, 4).setValue('TOTAL GENERAL').setFontWeight('bold').setHorizontalAlignment('right');
  ws.getRange(filaTotal, 5).setValue(totalGeneral > 0 ? totalGeneral : 0)
    .setFontWeight('bold').setNumberFormat('[$$-es-AR]#,##0.00').setBackground('#e8f4f9');

  var filaInfo = filaTotal + 2;
  ws.getRange(filaInfo,     1).setValue('OT: ' + ot);
  ws.getRange(filaInfo + 1, 1).setValue('Reseller: ' + (data.reseller || ''));
  ws.getRange(filaInfo + 2, 1).setValue('Descuento aplicado: ' + Math.round(_obtenerDescuentoReseller(data.reseller) * 100) + '% (Precio Reseller)');
  ws.getRange(filaInfo + 3, 1).setValue('Generado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'));
  ws.getRange(filaInfo, 1, 4, 1).setFontColor('#888888').setFontSize(10);

  ws.autoResizeColumns(1, 5);
  SpreadsheetApp.flush();

  var blob = tmpSS.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  blob.setName('Factura_' + ot + '.xlsx');
  DriveApp.getFileById(tmpSS.getId()).setTrashed(true);
  return blob;
}


function _construirEmailFacturacion(data, items, totalGeneral, infoCliente, descuento) {
  var tablaHTML =
    "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:16px'>" +
    "<thead><tr style='background:#00a3e0'>" +
    "<th style='padding:8px 10px;color:#fff;text-align:left;border:1px solid #0088bb'>CÓDIGO</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:left;border:1px solid #0088bb'>DESCRIPCIÓN</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:center;border:1px solid #0088bb'>CANTIDAD</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:right;border:1px solid #0088bb'>PVP UNITARIO</th>" +
    "<th style='padding:8px 10px;color:#fff;text-align:right;border:1px solid #0088bb'>TOTAL ITEM</th>" +
    "</tr></thead><tbody>";

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var bg = i % 2 === 0 ? '#ffffff' : '#f7f9fc';
    tablaHTML +=
      "<tr style='background:" + bg + "'>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;font-weight:700;color:#00a3e0'>" + _htmlEsc(item.codigo) + "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0'>" + _htmlEsc(item.descripcion) + "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:center'>" + item.cantidad + "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>" +
        (item.pvp > 0 ? "USD " + _fmtNum(item.pvp) + " <span style='color:#888;font-size:10px'>(-" + Math.round(descuento*100) + "% desc.)</span>" : "—") +
      "</td>" +
      "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right;font-weight:700'>" +
        (item.total > 0 ? "USD " + _fmtNum(item.total) : "—") +
      "</td></tr>";
  }

  if (totalGeneral > 0) {
    tablaHTML +=
      "<tr style='background:#e8f4f9;font-weight:700'>" +
      "<td colspan='4' style='padding:8px 10px;border:1px solid #e0e0e0;text-align:right'>TOTAL GENERAL</td>" +
      "<td style='padding:8px 10px;border:1px solid #e0e0e0;text-align:right;color:#00a3e0;font-size:14px'>" +
        "USD " + _fmtNum(totalGeneral) +
      "</td></tr>";
  }
  tablaHTML += "</tbody></table>";

  var resellerInfo =
    filaDetalle('Reseller / Empresa', '<strong>' + _htmlEsc(data.reseller) + '</strong>') +
    (infoCliente ? (
      filaDetalle('CUIT', _htmlEsc(infoCliente.cuit) || '—') +
      filaDetalle('Dirección', _htmlEsc((infoCliente.direccion||'—') + (infoCliente.localidad ? ', ' + infoCliente.localidad : ''))) +
      filaDetalle('Código Postal', _htmlEsc(infoCliente.cp) || '—') +
      filaDetalle('Teléfono', _htmlEsc(infoCliente.telefono) || '—')
    ) : '');

  var cuerpo =
    bloqueCard('📋 Detalle de la Orden',
      filaDetalle('OT', '<strong>' + data.ot + '</strong>') +
      filaDetalle('Equipo / Modelo', data.equipo ? _htmlEsc(data.equipo) : '—') +
      filaDetalle('N° de Serie', data.sn ? _htmlEsc(data.sn) : '—') +
      (data.cas ? filaDetalle('Caso DJI (CAS/FWR)', _htmlEsc(data.cas)) : '') +
      filaDetalle('Garantía', 'OOW — Fuera de garantía'),
      '#00a3e0') +
    "<div style='margin-bottom:16px'>" +
      "<p style='font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase;letter-spacing:.04em'>Repuestos a facturar (" + Math.round(descuento*100) + "% dto. reseller)</p>" +
      tablaHTML +
    "</div>" +
    "<div style='background:#f5f9fc;border:1px solid #ddeef7;border-radius:8px;padding:4px 16px;margin-bottom:16px'>" +
      "<p style='font-size:11px;font-weight:700;color:#888;text-transform:uppercase;margin:10px 0 4px'>Datos del reseller</p>" +
      resellerInfo +
    "</div>" +
    bloqueCard('📎 Adjunto XLS', 'Se incluye el detalle en formato Excel (.xlsx). Generá la factura y enviala al reseller.', '#27ae60');

  return construirEmailHTML(
    'Solicitud de Facturación — OT ' + data.ot,
    'Estimado equipo de Administración,<br>Se solicita la emisión de factura por repuestos despachados en la siguiente orden:',
    cuerpo,
    'Procesá este pedido a la brevedad y emití la factura correspondiente al reseller.'
  );
}


function solicitarFactura(data) {
  var _u = identificarUsuario();
  if (!_u) return { ok: false, msg: 'No autorizado.' };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    // Reportado por el usuario: Reseller / Reseller Propio no tienen nada real que facturarle a
    // Administración por esta vía (el reseller ya paga los repuestos por su cuenta, vía Portal/
    // WOS) — mismo criterio que el mail automático de facturación (enviarNotificaciones bloque 4).
    // Server-side porque el gate del botón en el cliente es solo conveniencia de UI, no seguridad.
    if (String(data && data.circuito || '') !== 'Taller') {
      return { ok: false, msg: 'La solicitud de facturación a Administración es exclusiva del circuito Taller Central.' };
    }

    var emailAdmin = getMailAdministracion();
    if (!emailAdmin) return { ok: false, msg: 'No se encontró "Administracion" en Usuarios_Internos (Columna A) con email en Columna B.' };

    var infoCliente = obtenerInfoCliente(data.reseller);
    var descuento   = _obtenerDescuentoReseller(data.reseller);

    // Leer precios desde DB_REPUESTOS col F (índice 5)
    var precioMap = {};
    var hojaDB = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaDB) {
      var dRep = getSheetValues(hojaDB);
      for (var pr = 1; pr < dRep.length; pr++) {
        var codR = String(dRep[pr][1]||'').trim().toUpperCase();
        if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5]||'0').replace(',','.')) || 0;
      }
    }

    // Parsear repuestos — se factura lo despachado (E:)
    var items = [];
    var totalGeneral = 0;
    if (data.repuestos && data.repuestos !== 'Sin consumo de repuestos') {
      var ls = data.repuestos.split(' ; ');
      for (var i = 0; i < ls.length; i++) {
        var p = ls[i].split(' | ');
        if (p.length < 3) continue;
        var cod  = p[0].trim();
        var desc = p[1].trim();
        var cant = parseInt(p[2].split('E:')[1]) || 0;
        if (cant <= 0) continue;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var pvp     = pvpBase > 0 ? pvpBase * (1 - descuento) : 0;
        var total   = pvp * cant;
        totalGeneral += total;
        items.push({ codigo: cod, descripcion: desc, cantidad: cant, pvp: pvp, total: total });
      }
    }
    if (!items.length) return { ok: false, msg: 'No hay repuestos despachados (Env. > 0) para facturar en esta OT.' };

    // Generar XLSX como Blob
    var xlsBlob = _generarXLSFactura(data.ot, items, totalGeneral, data);

    // Enviar email con adjunto
    var htmlEmail = _construirEmailFacturacion(data, items, totalGeneral, infoCliente, descuento);
    var asunto    = 'SOLICITUD DE FACTURACIÓN - OT: ' + data.ot + ' - ' + data.reseller;

    GmailApp.sendEmail(emailAdmin, asunto, '', {
      htmlBody:    htmlEmail,
      name:        CONFIG.NOMBRE_REMITENTE,
      replyTo:     CONFIG.EMAIL_SUPERVISOR,
      attachments: [xlsBlob]
    });
    registrarEmailLog(data.ot, emailAdmin, 'Administración', asunto, 'OK');

    // Actualizar estado de la OT en la hoja
    var fila    = parseInt(data.fila);
    var hojaOT  = getSheet(SCHEMA.SHEETS.OT);
    var estadoAnt = String(data.estado || '');
    var _ahoraFact = new Date();
    hojaOT.getRange(fila, SCHEMA.OT.ESTADO      + 1).setValue('ESPERANDO FACTURA');
    hojaOT.getRange(fila, SCHEMA.OT.FECHA_ESTADO + 1).setValue(_ahoraFact);
    // Historial + sello de concurrencia (ver actualizarOrden, HUB_OTs.js) — esta OT cambió
    // server-side, así que el cliente necesita el timestamp nuevo para no chocar con un
    // CONFLICT falso en el próximo guardado.
    var _celHistFact = hojaOT.getRange(fila, SCHEMA.OT.HISTORIAL_ESTADOS + 1);
    var _histFact = [];
    try { var _hrFact = _celHistFact.getValue(); if (_hrFact) _histFact = JSON.parse(_hrFact); } catch(e2) {}
    _histFact.push({ f: _ahoraFact.getTime(), ant: estadoAnt, nvo: 'ESPERANDO FACTURA', tec: '' });
    _celHistFact.setValue(JSON.stringify(_histFact));
    hojaOT.getRange(fila, SCHEMA.OT.ULTIMA_MODIFICACION + 1).setValue(_ahoraFact);

    var operador = Session.getActiveUser().getEmail();
    registrarLog(data.ot, operador, operador, 'FACTURACIÓN', estadoAnt, 'ESPERANDO FACTURA',
                 'Solicitud de factura enviada · mail: ' + emailAdmin);

    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.OT);

    return { ok: true, mailEnviado: emailAdmin };

  } catch(e) {
    Logger.log('solicitarFactura: ' + e);
    return { ok: false, msg: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}
