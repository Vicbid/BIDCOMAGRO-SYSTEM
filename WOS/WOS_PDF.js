// @version 1.0
// ============================================================
//  WOS — Generación del PDF de nota de entrega.
//  Extraído de WOS_GmailFlow.js 2.30 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


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
