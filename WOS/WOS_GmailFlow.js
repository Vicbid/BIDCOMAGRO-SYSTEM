// ============================================================
// @version 2.22
//  WOS — Gestión de hilos Gmail · V-1.0 (Hitos 2–5)
//
//  Hito 1 vive en PORTAL_RESELLER/RS_Pedidos.js.
//  Este módulo requiere que COL.THREAD_ID (col R) ya esté
//  poblado por el Portal al confirmar el pedido.
//
//  TRIGGER a instalar: WOS_instalarTriggerDetector()
//  (ejecutar UNA VEZ desde el editor de Apps Script)
// ============================================================

// ── Helpers internos ─────────────────────────────────────────

// ── Template de email idéntico al Portal Reseller ─────────────
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

// Genera PDF de Nota de Entrega, lo guarda en Drive y devuelve { blob, url, nombreNota }
// notaNumStr: "01", "02", etc. (usar _wosNextNotaNum para obtenerlo)
// transp y costoEnvio son opcionales. bultos = [{tracking, peso}, ...]
function _wosGenerarPDF(numero, notaNumStr, reseller, items, fecha, transp, bultos, costoEnvio) {
  var tempSs = null;
  try {
    var meta     = _wosGetResellerMeta(reseller);
    var fechaStr = fecha || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    transp       = String(transp || '').trim();
    bultos       = bultos || [];
    costoEnvio   = Number(costoEnvio) || 0;
    var pesoTotal = 0;
    for (var bi = 0; bi < bultos.length; bi++) pesoTotal += Number(bultos[bi].peso) || 0;
    pesoTotal = Math.round(pesoTotal * 1000) / 1000;

    tempSs = SpreadsheetApp.create('TEMP_NE_' + numero);
    var sheet = tempSs.getActiveSheet();
    sheet.setName('Nota de Entrega');
    var ri = 1;

    // ── 1. CABECERA ───────────────────────────────────────────
    sheet.getRange(ri, 1, 3, 4).merge().setValue('BIDCOMAGRO')
      .setFontSize(20).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#00a3e0').setVerticalAlignment('middle').setHorizontalAlignment('left');
    sheet.getRange(ri,     5, 1, 3).merge().setValue('NOTA DE ENTREGA')
      .setFontSize(11).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('bottom');
    sheet.getRange(ri + 1, 5, 1, 3).merge().setValue('Nº ' + numero + '-' + notaNumStr)
      .setFontSize(10).setFontColor('#cce8f4')
      .setBackground('#007ab3').setHorizontalAlignment('right');
    sheet.getRange(ri + 2, 5, 1, 3).merge().setValue('Fecha: ' + fechaStr)
      .setFontSize(9).setFontColor('#cce8f4')
      .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('top');
    sheet.setRowHeight(ri, 24); sheet.setRowHeight(ri+1, 18); sheet.setRowHeight(ri+2, 20);
    ri += 3;

    // ── 2. DATOS DEL RESELLER ─────────────────────────────────
    sheet.setRowHeight(ri, 6);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    sheet.getRange(ri, 1, 1, 7).merge().setValue('DATOS DEL RESELLER DESTINATARIO')
      .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
      .setBackground('#2d3436').setHorizontalAlignment('left').setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22);
    ri++;

    var resellerRows = [
      ['Nombre / Razón Social', meta.nombre    || '—'],
      ['Dirección Comercial',   meta.direccion || '—'],
      ['Teléfono de Contacto',  meta.telefono  || '—'],
      ['Transportista',         transp         || '—'],
      ['Costo de Envío (ARS)',        costoEnvio > 0 ? '$ ' + _formatMoneda(costoEnvio) : '—']
    ];
    for (var bri = 0; bri < bultos.length; bri++) {
      var bt     = bultos[bri];
      var bLabel = bultos.length > 1 ? 'Bulto ' + (bri + 1) : 'Nº de Seguimiento';
      var bVal   = (bt.tracking || '—') + (bt.peso > 0 ? '  ·  ' + bt.peso + ' kg' : '');
      resellerRows.push([bLabel, bVal]);
    }
    if (pesoTotal > 0 && bultos.length > 1) resellerRows.push(['Peso Total', pesoTotal + ' kg']);

    for (var rri = 0; rri < resellerRows.length; rri++) {
      var rBg = rri % 2 === 0 ? '#f7f8fa' : '#ffffff';
      sheet.getRange(ri, 1, 1, 2).merge().setValue(resellerRows[rri][0])
        .setFontSize(9).setFontWeight('bold').setFontColor('#5e6778').setBackground(rBg);
      sheet.getRange(ri, 3, 1, 5).merge().setValue(resellerRows[rri][1])
        .setFontSize(9).setFontWeight(rri === 0 ? 'bold' : 'normal')
        .setFontColor('#1a1a2e').setBackground(rBg);
      sheet.setRowHeight(ri, 20);
      ri++;
    }

    // ── 3. TABLA DE ÍTEMS ─────────────────────────────────────
    sheet.setRowHeight(ri, 8);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#ffffff');
    ri++;

    sheet.getRange(ri, 1, 1, 7).setValues([['SKU', 'Descripción', 'Cant.', 'PVP USD', 'Precio Res. USD', 'Subtotal USD', '']])
      .setBackground('#00a3e0').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(9).setVerticalAlignment('middle');
    sheet.setRowHeight(ri, 22);
    ri++;

    // Determinar si renderizar agrupado por caja
    var hasCajaAssign = bultos.length > 1;
    if (hasCajaAssign) {
      var anyAssigned = false;
      for (var ci = 0; ci < items.length; ci++) {
        if (items[ci].cajaIdx !== undefined) { anyAssigned = true; break; }
      }
      hasCajaAssign = anyAssigned;
    }

    var total = 0;
    if (hasCajaAssign) {
      // Ordenar ítems por cajaIdx
      var sortedItems = items.slice().sort(function(a, b) {
        return (a.cajaIdx || 0) - (b.cajaIdx || 0);
      });
      var currentCajaIdx = -1;
      var altRow = 0;
      for (var i = 0; i < sortedItems.length; i++) {
        var it      = sortedItems[i];
        var cajaIdx = it.cajaIdx !== undefined ? it.cajaIdx : 0;
        // Fila separadora de caja
        if (cajaIdx !== currentCajaIdx) {
          currentCajaIdx = cajaIdx;
          altRow = 0;
          var blt      = bultos[cajaIdx] || {};
          var cajaLbl  = 'Caja ' + (cajaIdx + 1);
          if (blt.tracking) cajaLbl += '  ·  Tracking: ' + blt.tracking;
          if (blt.peso > 0) cajaLbl += '  ·  ' + blt.peso + ' kg';
          sheet.getRange(ri, 1, 1, 7).merge().setValue(cajaLbl)
            .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
            .setBackground('#005082').setHorizontalAlignment('left').setVerticalAlignment('middle');
          sheet.setRowHeight(ri, 20);
          ri++;
        }
        var cant  = Number(it.cantDesp) || 0;
        var prec  = Number(it.precio)   || 0;
        var pvp   = prec > 0 ? Math.round(prec / 0.60 * 100) / 100 : 0;
        var sub   = cant * prec;
        total    += sub;
        var rowBg = (altRow % 2 === 0) ? '#ffffff' : '#f5f7fa';
        sheet.getRange(ri, 1, 1, 7).setValues([[
          it.sku  || '—',
          it.desc || '—',
          cant,
          pvp  > 0 ? 'USD ' + _formatMoneda(pvp)  : '—',
          prec > 0 ? 'USD ' + _formatMoneda(prec) : '—',
          sub  > 0 ? 'USD ' + _formatMoneda(sub)  : '—',
          ''
        ]]).setFontSize(9).setBackground(rowBg).setVerticalAlignment('middle');
        sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#00a3e0');
        sheet.getRange(ri, 3).setHorizontalAlignment('center');
        sheet.getRange(ri, 4, 1, 3).setHorizontalAlignment('right');
        sheet.setRowHeight(ri, 20);
        ri++;
        altRow++;
      }
    } else {
      for (var i = 0; i < items.length; i++) {
        var it   = items[i];
        var cant = Number(it.cantDesp) || 0;
        var prec = Number(it.precio)   || 0;
        var pvp  = prec > 0 ? Math.round(prec / 0.60 * 100) / 100 : 0;
        var sub  = cant * prec;
        total   += sub;
        var rowBg = (i % 2 === 0) ? '#ffffff' : '#f5f7fa';
        sheet.getRange(ri, 1, 1, 7).setValues([[
          it.sku  || '—',
          it.desc || '—',
          cant,
          pvp  > 0 ? 'USD ' + _formatMoneda(pvp)  : '—',
          prec > 0 ? 'USD ' + _formatMoneda(prec) : '—',
          sub  > 0 ? 'USD ' + _formatMoneda(sub)  : '—',
          ''
        ]]).setFontSize(9).setBackground(rowBg).setVerticalAlignment('middle');
        sheet.getRange(ri, 1).setFontWeight('bold').setFontColor('#00a3e0');
        sheet.getRange(ri, 3).setHorizontalAlignment('center');
        sheet.getRange(ri, 4, 1, 3).setHorizontalAlignment('right');
        sheet.setRowHeight(ri, 20);
        ri++;
      }
    }

    // ── 4. FILA TOTAL ─────────────────────────────────────────
    if (total > 0) {
      sheet.getRange(ri, 1, 1, 5).merge()
        .setValue('TOTAL (precio reseller · no incluye impuestos)')
        .setFontSize(10).setFontWeight('bold').setFontColor('#5e6778')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 6).setValue('USD ' + _formatMoneda(total))
        .setFontSize(11).setFontWeight('bold').setFontColor('#00a3e0')
        .setBackground('#e8f5fc').setHorizontalAlignment('right').setVerticalAlignment('middle');
      sheet.getRange(ri, 7).setValue('').setBackground('#e8f5fc');
      sheet.setRowHeight(ri, 24);
      ri++;
    }

    // ── 5. PIE ────────────────────────────────────────────────
    sheet.setRowHeight(ri, 4);
    sheet.getRange(ri, 1, 1, 7).merge().setBackground('#00a3e0');
    ri++;
    sheet.getRange(ri, 1, 1, 7).merge()
      .setValue('Documento generado automáticamente — WOS · BIDCOMAGRO · ' + meta.nombre)
      .setFontSize(8).setFontStyle('italic').setFontColor('#9ba5b4')
      .setHorizontalAlignment('center').setBackground('#ffffff');
    sheet.setRowHeight(ri, 16);

    // Anchos de columna
    sheet.setColumnWidth(1, 110); // SKU
    sheet.setColumnWidth(2, 210); // Descripción
    sheet.setColumnWidth(3, 48);  // Cant.
    sheet.setColumnWidth(4, 78);  // PVP USD
    sheet.setColumnWidth(5, 82);  // Precio Res.
    sheet.setColumnWidth(6, 82);  // Subtotal
    sheet.setColumnWidth(7, 1);   // spacer

    // ── HOJA ANEXO: NÚMEROS DE SERIE ─────────────────────────────
    var itemsConSerial = [];
    for (var sci = 0; sci < items.length; sci++) {
      if (items[sci].seriales) itemsConSerial.push(items[sci]);
    }
    if (itemsConSerial.length > 0) {
      // Nota al pie de la NE principal
      sheet.getRange(ri, 1, 1, 7).merge()
        .setValue('Ver Anexo — N\xfameros de Serie (p\xe1gina siguiente)')
        .setFontSize(8).setFontStyle('italic').setFontColor('#00a3e0')
        .setHorizontalAlignment('center').setBackground('#f0f8ff');
      sheet.setRowHeight(ri, 14);
      ri++;

      var sSheet = tempSs.insertSheet('N\xba de Serie');
      var sri = 1;

      // Cabecera
      sSheet.getRange(sri, 1, 3, 4).merge().setValue('BIDCOMAGRO')
        .setFontSize(16).setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#00a3e0').setVerticalAlignment('middle');
      sSheet.getRange(sri,     5, 1, 3).merge().setValue('ANEXO — N\xdaMEROS DE SERIE')
        .setFontSize(10).setFontWeight('bold').setFontColor('#ffffff')
        .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('bottom');
      sSheet.getRange(sri + 1, 5, 1, 3).merge().setValue('NE ' + numero + '-' + notaNumStr)
        .setFontSize(9).setFontColor('#cce8f4')
        .setBackground('#007ab3').setHorizontalAlignment('right');
      sSheet.getRange(sri + 2, 5, 1, 3).merge().setValue(meta.nombre || '')
        .setFontSize(9).setFontColor('#cce8f4')
        .setBackground('#007ab3').setHorizontalAlignment('right').setVerticalAlignment('top');
      sSheet.setRowHeight(sri, 24); sSheet.setRowHeight(sri+1, 18); sSheet.setRowHeight(sri+2, 18);
      sri += 3;

      sSheet.setRowHeight(sri, 8);
      sSheet.getRange(sri, 1, 1, 7).merge().setBackground('#ffffff');
      sri++;

      for (var sii = 0; sii < itemsConSerial.length; sii++) {
        var sIt    = itemsConSerial[sii];
        var sNums  = String(sIt.seriales).split(',');

        // Sub-cabecera del ítem
        sSheet.getRange(sri, 1, 1, 7).merge()
          .setValue(sIt.sku + '   \xb7   ' + sIt.desc + '   (' + sIt.cantDesp + ' u.)')
          .setFontSize(9).setFontWeight('bold').setFontColor('#ffffff')
          .setBackground('#2d3436');
        sSheet.setRowHeight(sri, 22);
        sri++;

        // Seriales numerados
        for (var sni = 0; sni < sNums.length; sni++) {
          var snVal = String(sNums[sni]).trim();
          if (!snVal) continue;
          // "SO-123 x50" = bolsa de 50 unidades con un único código → mostrarlo claro para auditar.
          var _bag = snVal.match(/^(.*?)\s+x(\d+)$/i);
          if (_bag) snVal = _bag[1] + '   \xb7   bolsa de ' + _bag[2] + ' u.';
          var snBg = sni % 2 === 0 ? '#f7f8fa' : '#ffffff';
          sSheet.getRange(sri, 1, 1, 2).merge()
            .setValue(sni + 1)
            .setFontSize(9).setFontColor('#5e6778').setBackground(snBg)
            .setHorizontalAlignment('center').setFontWeight('bold');
          sSheet.getRange(sri, 3, 1, 5).merge()
            .setValue(snVal)
            .setFontSize(10).setFontColor('#1a1a2e').setBackground(snBg)
            .setFontFamily('Courier New').setFontWeight('bold');
          sSheet.setRowHeight(sri, 22);
          sri++;
        }

        if (sii < itemsConSerial.length - 1) {
          sSheet.setRowHeight(sri, 8);
          sSheet.getRange(sri, 1, 1, 7).merge().setBackground('#ffffff');
          sri++;
        }
      }

      // Pie
      sSheet.setRowHeight(sri, 4);
      sSheet.getRange(sri, 1, 1, 7).merge().setBackground('#00a3e0');
      sri++;
      sSheet.getRange(sri, 1, 1, 7).merge()
        .setValue('Documento generado autom\xe1ticamente — WOS \xb7 BIDCOMAGRO \xb7 ' + meta.nombre)
        .setFontSize(8).setFontStyle('italic').setFontColor('#9ba5b4')
        .setHorizontalAlignment('center').setBackground('#ffffff');
      sSheet.setRowHeight(sri, 16);

      // Anchos
      sSheet.setColumnWidth(1, 40);
      sSheet.setColumnWidth(2, 10);
      sSheet.setColumnWidth(3, 350);
      sSheet.setColumnWidth(4, 1);
      sSheet.setColumnWidth(5, 1);
      sSheet.setColumnWidth(6, 1);
      sSheet.setColumnWidth(7, 1);
    }

    SpreadsheetApp.flush();

    var nombreNota = 'NE_' + numero + '-' + notaNumStr + '.pdf';
    var ssFile     = DriveApp.getFileById(tempSs.getId());
    var pdfBlob    = ssFile.getAs('application/pdf');   // captura el PDF con el formato actual
    pdfBlob.setName(nombreNota);

    var pdfUrl   = '';
    var sheetUrl = '';   // versión spreadsheet de la MISMA NE (Google Sheet, link)
    try {
      var folder  = DriveApp.getFolderById(_wosConfig().pdfFolderId);
      var pdfFile = folder.createFile(pdfBlob);
      pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      pdfUrl = pdfFile.getUrl();

      // Conservar la hoja temporal como Google Sheet (misma NE, versión planilla).
      // Se renombra, se mueve a la carpeta de NE y se comparte con link — no se manda a la papelera.
      ssFile.setName('NE_' + numero + '-' + notaNumStr);
      try { folder.addFile(ssFile); DriveApp.getRootFolder().removeFile(ssFile); } catch(eMov) {}
      ssFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheetUrl = ssFile.getUrl();
    } catch(eDrive) {
      Logger.log('_wosGenerarPDF Drive save: ' + eDrive);
      // Si falló el guardado, no dejar la hoja huérfana en Drive
      try { ssFile.setTrashed(true); } catch(eT2) {}
    }

    return { blob: pdfBlob, url: pdfUrl, nombreNota: nombreNota, sheetUrl: sheetUrl };
  } catch(e) {
    Logger.log('_wosGenerarPDF: ' + e);
    if (tempSs) { try { DriveApp.getFileById(tempSs.getId()).setTrashed(true); } catch(eT) {} }
    return null;
  }
}


// ─────────────────────────────────────────────────────────────
//  HITO 4b — Despachar Completo (reemplaza WOS_despacharPedido)
//  Registra cantidades, guarda tracking, responde en el hilo
//  original y notifica a facturación en un solo paso.
//
//  despachos:    [{row (1-indexed), cantDesp}]
//  transportista: string requerido ("Correo Argentino", "Credifin", "Via Cargo")
//  bultos:        [{tracking, peso}] — un objeto por bulto físico
//  costoEnvio:    costo total del envío (opcional)
//  reqToken:      token de idempotencia (opcional) — si el mismo token ya se
//                 procesó, se devuelve el resultado previo sin re-ejecutar
//                 (protege contra doble-click / reintento tras respuesta perdida)
// ─────────────────────────────────────────────────────────────
// Combina códigos de seguimiento sin pisar ni duplicar. Acepta valores ya
// combinados ("T1 | T2") en cualquiera de los dos lados. Devuelve "T1 | T2 | ...".
function _wosMergeTracking(oldVal, newVal) {
  var out = [];
  function _add(v) {
    var parts = String(v || '').split('|');
    for (var i = 0; i < parts.length; i++) {
      var t = parts[i].trim();
      if (t && out.indexOf(t) === -1) out.push(t);
    }
  }
  _add(oldVal);
  _add(newVal);
  return out.join(' | ');
}

// ─────────────────────────────────────────────────────────────
// RECUPERACIÓN de códigos de seguimiento pisados: los parsea de los mails de
// despacho del hilo del pedido (cada despacho mandó "Código de seguimiento: ...").
//   numero  : N° de pedido, ej "PR-0035".
//   guardar : true → fusiona los códigos recuperados en las filas ya despachadas
//             (col Q); false/omitido → solo devuelve el reporte (no toca la planilla).
function WOS_recuperarTracking(numero, guardar) {
  try {
    numero = String(numero || '').trim();
    if (!numero) return { ok: false, error: 'Falta el número de pedido' };
    var ped = _wosLeerPedido(numero);
    if (!ped.hoja) return { ok: false, error: 'Pedido no encontrado: ' + numero };

    // Reunir los mensajes del hilo (por threadId guardado; respaldo: búsqueda por N°)
    var mensajes = [];
    try {
      if (ped.threadId) {
        var th = GmailApp.getThreadById(ped.threadId);
        if (th) mensajes = th.getMessages();
      }
    } catch(eT) { Logger.log('WOS_recuperarTracking thread: ' + eT); }
    if (!mensajes.length) {
      try {
        var hilos = GmailApp.search('subject:"' + numero + '"', 0, 10);
        for (var h = 0; h < hilos.length; h++) {
          var ms = hilos[h].getMessages();
          for (var mm = 0; mm < ms.length; mm++) mensajes.push(ms[mm]);
        }
      } catch(eS) { Logger.log('WOS_recuperarTracking search: ' + eS); }
    }

    // Parsear "Código de seguimiento: ..." de cada mensaje (uno por despacho)
    var re = /C[o\xf3]digo de seguimiento:\s*([^\n\r]+)/i;
    var detalle = [], todos = [];
    for (var i = 0; i < mensajes.length; i++) {
      var body = '';
      try { body = mensajes[i].getPlainBody() || ''; } catch(eB) { continue; }
      var mtch = body.match(re);
      if (!mtch) continue;
      var partes = mtch[1].split(/\s*[|,]\s*/);
      var limpios = [];
      for (var c = 0; c < partes.length; c++) {
        var t = String(partes[c] || '').trim();
        if (!t) continue;
        if (limpios.indexOf(t) === -1) limpios.push(t);
        if (todos.indexOf(t)   === -1) todos.push(t);
      }
      if (!limpios.length) continue;
      var fecha = '';
      try { fecha = Utilities.formatDate(mensajes[i].getDate(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); } catch(eD) {}
      detalle.push({ fecha: fecha, codigos: limpios });
    }

    var res = { ok: true, numero: numero, codigos: todos, detalle: detalle, mensajesRevisados: mensajes.length };

    // Escritura opcional: fusiona TODOS los códigos recuperados en las filas ya despachadas
    if (guardar && todos.length) {
      var escritas = 0;
      var mergeStr = todos.join(' | ');
      for (var r = 1; r < ped.datos.length; r++) {
        if (String(ped.datos[r][COL.NUMERO] || '').trim() !== numero) continue;
        var yaDesp = ped.datos[r][COL.FECHA_DESPACHO] || String(ped.datos[r][COL.TRACKING] || '').trim();
        if (!yaDesp) continue; // solo filas efectivamente despachadas
        var mergedRow = _wosMergeTracking(ped.datos[r][COL.TRACKING], mergeStr);
        ped.hoja.getRange(r + 1, COL.TRACKING + 1).setValue(mergedRow);
        escritas++;
      }
      res.filasActualizadas = escritas;
    }

    Logger.log('WOS_recuperarTracking ' + numero + ': ' + todos.length + ' código(s) → ' + JSON.stringify(todos));
    return res;
  } catch(e) {
    Logger.log('WOS_recuperarTracking: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// Wrapper para correr DESDE EL EDITOR (no acepta parámetros ahí):
// editá NUMERO y GUARDAR, ejecutá, y mirá el resultado en el Log (Ver → Registros).
//   1ª pasada con GUARDAR=false para revisar los códigos encontrados.
//   2ª pasada con GUARDAR=true para escribirlos en la planilla.
function WOS_recuperarTrackingTest() {
  var NUMERO  = 'PR-0000';   // ← poné acá el N° del pedido
  var GUARDAR = false;       // ← poné true cuando quieras escribirlos
  var r = WOS_recuperarTracking(NUMERO, GUARDAR);
  Logger.log(JSON.stringify(r, null, 2));
  return r;
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

// Diagnóstico de acceso a Carmen para el descuento de stock del despacho.
// Ejecutar desde el editor de Apps Script de WOS (Run ▶) y ver el resultado.
// Dice si abre la planilla, lista las pestañas (para ver si "Entregados" existe / cómo se
// llama) y prueba escribir+borrar una fila de test en "Entregados".
function WOS_diagnosticoCarmen() {
  var out = { carmenId: CARMEN_SS_ID, abrio: false, ok: false };
  try {
    var ss = SpreadsheetApp.openById(CARMEN_SS_ID);
    out.abrio = true;
    out.nombrePlanilla = ss.getName();
    var tabs = ss.getSheets().map(function(s) { return s.getName(); });
    out.pestanas = tabs;
    out.tieneEntregados  = tabs.indexOf('Entregados') !== -1;
    out.tieneUbicaciones = tabs.indexOf(CARMEN_UBICACIONES_TAB) !== -1;
    var hoja = ss.getSheetByName('Entregados');
    if (hoja) {
      hoja.appendRow(['__TEST_WOS__', 'diagnostico (borrar)', 0, 'TEST', '', '', new Date(), '']);
      SpreadsheetApp.flush();
      hoja.deleteRow(hoja.getLastRow());
      SpreadsheetApp.flush();
      out.escrituraOK = true;
    }
    out.ok = out.tieneEntregados && out.escrituraOK === true;
  } catch(e) {
    out.error = e.toString();
  }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}

// ── RECUPERACIÓN de líneas de "Entregados" que no se escribieron ──────────────
// (p. ej. cuando el sheet de Carmen se quedó sin filas y el appendRow falló).
// Reconcilia, por (pedido, SKU): esperado = CANT_DESP (lo despachado, acumulado en la hoja
// de pedidos) vs actual = suma de líneas en "Entregados" con ese pedido como Origen.
// La diferencia positiva = lo que falta descontar. IDEMPOTENTE: solo escribe el gap.
// desdeISO (opcional): 'YYYY-MM-DD' (desde medianoche) o 'YYYY-MM-DDTHH:mm:ss' (hora exacta,
// p.ej. '2026-07-15T12:00:00'). Considera solo filas con FECHA_DESPACHO desde ese instante.
function _wosCalcularEntregadosFaltantes(desdeISO) {
  var desde = null;
  if (desdeISO) { var _s = String(desdeISO); desde = new Date(_s.indexOf('T') >= 0 ? _s : _s + 'T00:00:00'); }
  var tz = Session.getScriptTimeZone();

  // 1. Esperado por (numero, SKU): sumar CANT_DESP de las filas despachadas de ambas hojas.
  var esperado = {};
  var hojas = [_getHojaPedidos(), _getHojaPedidosOT()];
  for (var h = 0; h < hojas.length; h++) {
    var hoja = hojas[h];
    if (!hoja) continue;
    var d = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var cd = Number(d[i][COL.CANT_DESP]) || 0;
      if (cd <= 0) continue;
      var fd = d[i][COL.FECHA_DESPACHO];
      if (desde) { if (!(fd instanceof Date) || fd < desde) continue; }
      var numero = String(d[i][COL.NUMERO] || '').trim();
      var sku    = String(d[i][COL.SKU] || '').trim().toUpperCase();
      if (!numero || !sku) continue;
      var key = numero + '||' + sku;
      if (!esperado[key]) esperado[key] = { numero: numero, sku: sku, desc: String(d[i][COL.DESC] || ''), cant: 0, fecha: '' };
      esperado[key].cant += cd;
      if (fd instanceof Date) esperado[key].fecha = Utilities.formatDate(fd, tz, 'dd/MM/yyyy');
    }
  }

  // 2. Actual en "Entregados" por (numero, SKU): sumar cantidad donde Origen==numero y SKU==sku.
  var carmenSS = SpreadsheetApp.openById(CARMEN_SS_ID);
  var hojaEnt  = carmenSS.getSheetByName('Entregados');
  if (!hojaEnt) return { ok: false, error: 'No existe la pesta\xf1a "Entregados" en Carmen.' };
  var actual = {};
  var dE = hojaEnt.getDataRange().getValues();
  // Entregados: 0=SKU, 1=desc, 2=cant, 3=Origen(pedido), 6=fecha, 7=ubic
  for (var e = 1; e < dE.length; e++) {
    var sku2 = String(dE[e][0] || '').trim().toUpperCase();
    var orig = String(dE[e][3] || '').trim();
    if (!sku2 || !orig) continue;
    // El Origen de un despacho puede venir como "PR-00028" o "PR-00028 (recupero)".
    var origNum = orig.replace(/\s*\(recupero\)\s*$/i, '').trim();
    var k2 = origNum + '||' + sku2;
    actual[k2] = (actual[k2] || 0) + (Number(dE[e][2]) || 0);
  }

  // 3. Gaps
  var faltantes = [], totalUnidades = 0;
  for (var key in esperado) {
    var esp = esperado[key];
    var act = actual[key] || 0;
    var falta = esp.cant - act;
    if (falta > 0) {
      faltantes.push({ numero: esp.numero, sku: esp.sku, desc: esp.desc, esperado: esp.cant, actual: act, faltante: falta, fecha: esp.fecha });
      totalUnidades += falta;
    }
  }
  faltantes.sort(function(a, b) { return b.faltante - a.faltante; });
  return { ok: true, faltantes: faltantes, totalUnidades: totalUnidades, hojaEnt: hojaEnt };
}

// PREVIEW (solo lectura): corré esto primero desde el editor de WOS y revisá la lista.
function WOS_previewEntregadosFaltantes(desdeISO) {
  try {
    var r = _wosCalcularEntregadosFaltantes(desdeISO);
    if (!r.ok) return r;
    Logger.log('PREVIEW Entregados faltantes: ' + r.faltantes.length + ' l\xedneas, ' + r.totalUnidades + ' unidades.');
    for (var i = 0; i < r.faltantes.length; i++) {
      var f = r.faltantes[i];
      Logger.log('  ' + f.numero + ' | ' + f.sku + ' | ' + (f.desc || '') +
        ' | despachado=' + f.esperado + ' en Entregados=' + f.actual + ' \x2192 FALTA ' + f.faltante +
        (f.fecha ? ' | ' + f.fecha : ''));
    }
    return { ok: true, lineas: r.faltantes.length, totalUnidades: r.totalUnidades, faltantes: r.faltantes };
  } catch(e) { Logger.log('WOS_previewEntregadosFaltantes: ' + e); return { ok: false, error: e.toString() }; }
}

// APLICAR: escribe en "Entregados" solo las líneas faltantes (marcadas "<pedido> (recupero)").
function WOS_aplicarEntregadosFaltantes(desdeISO) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch(eL) { return { ok: false, error: 'Otra operaci\xf3n de despacho en curso. Reintent\xe1 en unos segundos.' }; }
  try {
    var r = _wosCalcularEntregadosFaltantes(desdeISO);
    if (!r.ok) return r;
    if (!r.faltantes.length) return { ok: true, escritas: 0, totalUnidades: 0, mensaje: 'No hay l\xedneas faltantes: todo cuadra.' };
    var hojaEnt = r.hojaEnt;
    var tz = Session.getScriptTimeZone();
    var fecha = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    for (var i = 0; i < r.faltantes.length; i++) {
      var f = r.faltantes[i];
      hojaEnt.appendRow([f.sku, f.desc, f.faltante, f.numero + ' (recupero)', '', '', fecha, '']);
    }
    SpreadsheetApp.flush();
    Logger.log('RECUPERADAS ' + r.faltantes.length + ' l\xedneas de Entregados (' + r.totalUnidades + ' unidades).');
    return { ok: true, escritas: r.faltantes.length, totalUnidades: r.totalUnidades, detalle: r.faltantes };
  } catch(e) {
    Logger.log('WOS_aplicarEntregadosFaltantes: ' + e);
    return { ok: false, error: e.toString() };
  } finally { try { lock.releaseLock(); } catch(eF) {} }
}

// ── Wrappers sin argumentos para correr desde el botón "Ejecutar" del editor ──
// El botón Run no permite pasar 'desdeISO'. Estas fijan el corte del incidente.
// INCIDENTE 15/07: el sheet "Entregados" se quedó sin filas ~12:00 de ayer y NADA de lo
// despachado DESPUÉS quedó registrado. Corte exacto = 2026-07-15 12:00 (hora Buenos Aires).
// Como después del corte "Entregados" está vacío, para esas filas actual=0 y "faltante" =
// cantidad DESPACHADA completa: NO resta nada, escribe tal cual lo que entregaste y no se
// anotó. Deja afuera lo anterior (junio y el 06/07), que es otro tema.
var _WOS_CORTE = '2026-07-15T12:00:00';
function WOS_previewFaltantes_Desde15Jul12h() { return WOS_previewEntregadosFaltantes(_WOS_CORTE); }
function WOS_aplicarFaltantes_Desde15Jul12h() { return WOS_aplicarEntregadosFaltantes(_WOS_CORTE); }

// ── Diagnóstico/arreglo: líneas 'Cancelado' sin CANT_CANCEL (col Z) ───────────
// Antes del fix de Opción B, una línea cancelada quedaba con estado 'Cancelado' pero
// CANT_CANCEL vacío → CANT_PEND (=E−F−Z) seguía > 0 y bloqueaba la preparación del pedido
// ("No se puede preparar — sin stock suficiente"). Escanea Pedidos_resellers y Pedidos_OTs.
//   · WOS_previewCanceladosSinCancel(): SOLO lista las filas afectadas (no toca nada).
//   · WOS_aplicarCanceladosSinCancel(): setea Z = pendiente (CANT_SOL − CANT_DESP) → CANT_PEND = 0.
// Correr desde el editor de Apps Script (botón Ejecutar). Ver el resultado en el log / return.
function WOS_previewCanceladosSinCancel() { return _wosScanCanceladosSinCancel(false); }
function WOS_aplicarCanceladosSinCancel() { return _wosScanCanceladosSinCancel(true); }

function _wosScanCanceladosSinCancel(aplicar) {
  var out = { ok: true, aplicar: !!aplicar, afectados: [], total: 0, porHoja: {} };
  try {
    var hojas = [
      { nombre: 'Pedidos_resellers', hoja: _getHojaPedidos() },
      { nombre: 'Pedidos_OTs',       hoja: _getHojaPedidosOT() }
    ];
    for (var h = 0; h < hojas.length; h++) {
      var hoja = hojas[h].hoja;
      if (!hoja) { out.porHoja[hojas[h].nombre] = 'hoja no encontrada'; continue; }
      var datos = hoja.getDataRange().getValues();
      var cnt = 0;
      for (var i = 1; i < datos.length; i++) {
        if (String(datos[i][COL.ESTADO] || '').trim() !== EST.CANCELADO) continue;
        var sol  = Number(datos[i][COL.CANT_SOL])    || 0;
        var desp = Number(datos[i][COL.CANT_DESP])   || 0;
        var canc = Number(datos[i][COL.CANT_CANCEL]) || 0;
        var pend = sol - desp - canc;   // lo que la fórmula col G sigue mostrando como pendiente
        if (canc > 0 || pend <= 0) continue;   // ya tiene Z, o ya está en 0 → nada que arreglar
        var fila = i + 1;
        var nuevoZ = Math.max(0, sol - desp);
        out.afectados.push({
          hoja:     hojas[h].nombre,
          fila:     fila,
          numero:   String(datos[i][COL.NUMERO]   || ''),
          reseller: String(datos[i][COL.RESELLER] || ''),
          sku:      String(datos[i][COL.SKU]      || ''),
          cantSol:  sol,
          cantDesp: desp,
          pendiente: pend,
          nuevoCancel: nuevoZ
        });
        cnt++;
        if (aplicar) hoja.getRange(fila, COL.CANT_CANCEL + 1).setValue(nuevoZ);
      }
      out.porHoja[hojas[h].nombre] = cnt;
    }
    if (aplicar) SpreadsheetApp.flush();
    out.total = out.afectados.length;
    Logger.log('WOS ' + (aplicar ? 'APLICAR' : 'PREVIEW') + ' cancelados sin CANT_CANCEL: ' +
      out.total + ' fila(s) → ' + JSON.stringify(out.porHoja));
    for (var a = 0; a < out.afectados.length; a++) {
      var r = out.afectados[a];
      Logger.log('  [' + r.hoja + ' fila ' + r.fila + '] ' + r.numero + ' · ' + r.reseller +
        ' · ' + r.sku + ' · sol=' + r.cantSol + ' desp=' + r.cantDesp + ' → Z=' + r.nuevoCancel);
    }
    return out;
  } catch(e) {
    Logger.log('_wosScanCanceladosSinCancel: ' + e);
    return { ok: false, error: e.toString(), afectados: out.afectados, total: out.afectados.length };
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

    // Actualizar HUB PRO OT: E: en REPUESTOS (col Q) + estado si completo
    if (_esNumeroOT(numero)) {
      try {
        var otNum = numero.replace(/^OT-/, '');
        // SKU → total despachado (previo + este despacho)
        var skuDespMap = {};
        for (var di = 1; di < ped.datos.length; di++) {
          if (String(ped.datos[di][COL.NUMERO] || '').trim() !== numero) continue;
          var dSku = String(ped.datos[di][COL.SKU] || '').trim().toUpperCase();
          if (!dSku) continue;
          var totalDi = (Number(ped.datos[di][COL.CANT_DESP]) || 0) +
                        ((despMap[di + 1] !== undefined) ? (Number(despMap[di + 1]) || 0) : 0);
          skuDespMap[dSku] = (skuDespMap[dSku] || 0) + totalDi;
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
            // Estado: solo pasa a "Repuestos enviados" si el despacho fue completo
            if (!hayBackorder) hubHoja.getRange(hi + 1, 5).setValue('Repuestos enviados');
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
    var ped = _wosLeerPedido(numero);
    if (!ped.reseller) return { ok: false, error: 'Pedido no encontrado: ' + numero };
    // Sin threadId: el bloque try/catch de más abajo ya envía email nuevo como fallback.

    // Cliente externo (super-RTV): puede no estar en Resellers. El aviso va por replyAll al hilo
    // (threadId ya garantizado arriba); el email directo es solo fallback (abajo). No se aborta.
    var email = _wosGetEmailReseller(ped.reseller);

    // ── Tabla de faltantes ────────────────────────────────────
    var tablaFalt =
      "<table style='width:100%;border-collapse:collapse;font-size:12px;margin:16px 0'>" +
      "<thead><tr style='background:#fdecea'>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>SKU</th>" +
        "<th style='padding:7px 10px;text-align:left;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Solicitado</th>" +
        "<th style='padding:7px 10px;text-align:center;border-bottom:2px solid #f5a5a5;font-size:10px;font-weight:700;text-transform:uppercase;color:#7f1919'>Disponible</th>" +
      "</tr></thead><tbody>";
    for (var i = 0; i < faltantes.length; i++) {
      var f = faltantes[i];
      tablaFalt +=
        "<tr style='border-bottom:1px solid #f0f2f5'>" +
          "<td style='padding:8px 10px;font-family:monospace;font-weight:700;color:#e74c3c'>" + (f.sku  || '') + "</td>" +
          "<td style='padding:8px 10px;color:#1a202c'>"                                      + (f.desc || '') + "</td>" +
          "<td style='padding:8px 10px;text-align:center;font-weight:700'>"                  + (f.cantSol  || 0) + "</td>" +
          "<td style='padding:8px 10px;text-align:center;font-weight:700;color:" +
            (f.cantDisp > 0 ? '#1a9e4a' : '#e74c3c') + "'>"                                 + (f.cantDisp || 0) + "</td>" +
        "</tr>";
    }
    tablaFalt += "</tbody></table>";

    // ¿Hay algo disponible para despachar?
    var hayDisponible = false;
    for (var fd = 0; fd < faltantes.length; fd++) {
      if ((faltantes[fd].cantDisp || 0) > 0) { hayDisponible = true; break; }
    }

    // ── Cuerpo HTML ───────────────────────────────────────────
    var introDisp = hayDisponible
      ? "<p style='font-size:13px;color:#155724;background:#e8f7ee;border:1px solid #8fd4a8;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
          "&#128666; <strong>Vamos a despachar lo disponible</strong> a la brevedad.<br>" +
          "<span style='font-size:12px;color:#4a5568'>Solo necesitamos saber qué preferís hacer con los ítems faltantes:</span>" +
        "</p>"
      : "<p style='font-size:13px;color:#7c3c00;background:#fff3e0;border:1px solid #ffba7b;border-radius:8px;padding:12px 16px;margin:16px 0 4px'>" +
          "&#9888; <strong>No tenemos stock disponible</strong> para ninguno de los ítems faltantes.<br>" +
          "<span style='font-size:12px;color:#4a5568'>Por favor indicanos cómo querés proceder:</span>" +
        "</p>";

    var _wosUrl = '';
    try { _wosUrl = ScriptApp.getService().getUrl(); } catch(eUrl) { Logger.log('WOS URL error: ' + eUrl); }
    var _itemsStr = '';
    for (var fi = 0; fi < faltantes.length; fi++) {
      if (_itemsStr) _itemsStr += ',';
      _itemsStr += String(faltantes[fi].sku || '').toUpperCase().replace(/[,: ]/g, '') + ':' + (faltantes[fi].cantDisp || 0);
    }
    var _urlA = _wosUrl ? _wosUrl + '?page=resp_faltante&num=' + encodeURIComponent(numero) + '&op=A' : '';
    var _urlB = _wosUrl ? _wosUrl + '?page=resp_faltante&num=' + encodeURIComponent(numero) + '&op=B&items=' + encodeURIComponent(_itemsStr) : '';

    var _botonesHtml = _urlA
      ? "<div style='text-align:center;margin:22px 0 8px'>" +
          "<a href='" + _urlA + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#3730a3;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>Opci\xf3n A — Esperar el faltante</a>" +
          "<a href='" + _urlB + "' style='display:inline-block;margin:6px;padding:13px 22px;background:#92400e;color:#fff;border-radius:8px;font-size:14px;font-weight:700;text-decoration:none'>Opci\xf3n B — Cancelar el faltante</a>" +
        "</div>" +
        "<p style='font-size:11px;color:#888;text-align:center;margin:4px 0 0'>Si los botones no funcionan, respond\xe9 este correo con <strong>\"Opci\xf3n A\"</strong> o <strong>\"Opci\xf3n B\"</strong>.</p>"
      : "<p style='font-size:12px;color:#555;background:#f5f8fc;border-radius:6px;padding:10px 14px;line-height:1.6'>" +
          "Respond\xe9 este correo indicando tu elecci\xf3n: <strong>\"Opci\xf3n A\"</strong> o <strong>\"Opci\xf3n B\"</strong>." +
        "</p>";

    var opcionesHtml = introDisp +
      "<p style='font-size:13px;color:#1a1f2e;font-weight:700;margin:18px 0 10px'>\xbfQu\xe9 hacemos con el faltante?</p>" +
      "<div style='border:1px solid #c7d2fe;border-radius:8px;padding:14px 18px;margin-bottom:8px;background:#eef2ff'>" +
        "<p style='margin:0 0 4px;font-size:13px;color:#3730a3;font-weight:700'>OPCI\xd3N A — Esperar el faltante (segundo env\xedo)</p>" +
        "<p style='margin:0;font-size:12px;color:#4a5568'>Los \xedtems faltantes quedan pendientes. Cuando ingresen al stock los despachamos en un segundo env\xedo.</p>" +
      "</div>" +
      "<div style='border:1px solid #ffba7b;border-radius:8px;padding:14px 18px;margin-bottom:16px;background:#fff3e0'>" +
        "<p style='margin:0 0 4px;font-size:13px;color:#7c3c00;font-weight:700'>OPCI\xd3N B — Cancelar el faltante</p>" +
        "<p style='margin:0;font-size:12px;color:#4a5568'>Cancelamos definitivamente los \xedtems que no tenemos en stock. Solo se factura lo que se despacha ahora.</p>" +
      "</div>" +
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

    // ── Plain text (incluye WOSDATA para que Hito 3 lo parsee) ──
    var wosDataJson = JSON.stringify({ numero: numero, faltantes: faltantes });
    var plainBody =
      'Hola ' + ped.reseller + ',\n\n' +
      'Al procesar tu pedido ' + numero + ' detectamos un faltante de stock.\n' +
      (hayDisponible ? 'Vamos a despachar lo disponible. ' : '') +
      'Por favor indicanos qué hacer con los ítems faltantes:\n\n' +
      'OPCI\xd3N A: Esperar el faltante y recibirlo en un segundo env\xedo cuando el stock est\xe9 disponible.\n' +
      'OPCI\xd3N B: Cancelar definitivamente los \xedtems faltantes.\n\n' +
      'Respond\xe9 este correo con tu elecci\xf3n (escrib\xed "Opci\xf3n A" u "Opci\xf3n B").\n\n' +
      '===WOSDATA===\n' + wosDataJson + '\n===ENDWOSDATA===';

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

    // ── Cambiar estado por ítem: faltantes → En_Espera_Reseller, disponibles → Preparado ──
    var faltSet = {};
    for (var ff = 0; ff < faltantes.length; ff++) {
      faltSet[String(faltantes[ff].sku || '').trim().toUpperCase()] = true;
    }
    _wosSetEstadoPorSku(ped.hoja, ped.datos, numero, faltSet, EST.EN_ESPERA, EST.PREPARADO);
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
//  UTILIDAD — Recuperar Thread IDs perdidos
//  Ejecutar UNA VEZ desde el editor de Apps Script.
//  Para cada pedido sin Thread ID en col R, busca en Gmail
//  el hilo original por número de pedido y lo restaura.
//  Ver resultado en Ver → Registros de ejecución.
// ─────────────────────────────────────────────────────────────
function WOS_recuperarThreadIds() {
  var hojas = [_wosHoja(), _getHojaPedidosOT()].filter(Boolean);

  // Agrupar filas por pedido, solo los que no tienen threadId; registra la hoja de origen
  var sinThread = {}; // numero → { hoja, filas: [] }
  for (var h = 0; h < hojas.length; h++) {
    var hoja  = hojas[h];
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var num = String(datos[i][COL.NUMERO]    || '').trim();
      var tid = String(datos[i][COL.THREAD_ID] || '').trim();
      if (!num || tid) continue;
      if (!sinThread[num]) sinThread[num] = { hoja: hoja, filas: [] };
      sinThread[num].filas.push(i + 1); // filas 1-indexed
    }
  }

  var numeros = Object.keys(sinThread);
  Logger.log('WOS_recuperarThreadIds: ' + numeros.length + ' pedidos sin Thread ID → ' + numeros.join(', '));
  if (!numeros.length) return { ok: true, recuperados: 0, noEncontrados: [] };

  var recuperados = 0;
  var noEncontrados = [];

  for (var n = 0; n < numeros.length; n++) {
    var numero = numeros[n];
    var hoja   = sinThread[numero].hoja;
    var rows   = sinThread[numero].filas;

    // Buscar en Gmail: el Portal usa asunto "[PEDIDO] PR-XXXXX — ..."
    // Un solo número es suficiente porque los IDs son únicos
    var threads = [];
    try {
      threads = GmailApp.search('subject:"' + numero + '"', 0, 10);
    } catch(eS) {
      Logger.log('WOS_recuperarThreadIds: error buscando ' + numero + ' → ' + eS);
      noEncontrados.push(numero);
      continue;
    }

    if (!threads || threads.length === 0) {
      Logger.log('WOS_recuperarThreadIds: ' + numero + ' → sin resultados en Gmail');
      noEncontrados.push(numero);
      continue;
    }

    // Si hay más de un hilo (raro), usar el más antiguo (el del pedido original)
    var threadId = threads[threads.length - 1].getId();

    for (var r = 0; r < rows.length; r++) {
      hoja.getRange(rows[r], COL.THREAD_ID + 1).setValue(threadId);
    }
    recuperados++;
    Logger.log('WOS_recuperarThreadIds: ' + numero + ' (' + rows.length + ' fila/s) → ' + threadId);
  }

  SpreadsheetApp.flush();
  Logger.log('WOS_recuperarThreadIds RESULTADO: recuperados=' + recuperados +
             ' | no encontrados=' + (noEncontrados.length ? noEncontrados.join(', ') : 'ninguno'));
  return { ok: true, recuperados: recuperados, noEncontrados: noEncontrados };
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

    SpreadsheetApp.flush();

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
