// ── STOCK MANAGER — Stock ─────────────────────────────────────
// @version 1.1

//  CATÁLOGO REPUESTOS DJI (para Borrador de Pedido)
function cargarCatalogoBorrador() {
  try {
    var stkMap = _getCarmenStockMap(); // { SKU_UPPER: stockActual }
    var ss   = SpreadsheetApp.openById(CATALOGO_REPUESTOS_ID);
    var hoja = ss.getSheets()[0];
    var d    = hoja.getDataRange().getValues();
    var out  = [];
    // Cols: 0=Código Largo, 1=Código Corto, 2=Descripción, 3=Modelo, 4=Cant/eq, 5=Precio PVP
    for (var i = 1; i < d.length; i++) {
      var codL = String(d[i][0] || '').trim();
      var codC = String(d[i][1] || '').trim();
      var desc = String(d[i][2] || '').trim();
      if (!codL || !desc) continue;
      // Buscar stock: primero por Código Largo, luego por Código Corto
      var stk = stkMap[codL.toUpperCase()];
      if (stk === undefined) stk = stkMap[codC.toUpperCase()];
      out.push({
        codigo:      codL,
        codigoCorto: codC,
        descripcion: desc,
        modelo:      String(d[i][3] || '').trim(),
        stockActual: stk !== undefined ? stk : null
      });
    }
    return { ok: true, items: out };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

//  STOCK
function cargarStock(filtro) {
  try {
    // Carmen es la fuente primaria de ítems y stock actual
    var hojaCarmen = _getCarmenSS().getSheetByName('STOCK');
    var dCarmen    = hojaCarmen ? hojaCarmen.getDataRange().getValues() : [];
    // STOCK_REPUESTOS provee metadata (mínimo, categoría, ubicación, etc.)
    var dMaster    = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S          = SCHEMA.STOCK_REPUESTOS;
    var masterMap  = {};
    for (var mi = 1; mi < dMaster.length; mi++) {
      var mk = String(dMaster[mi][S.CODIGO] || '').trim().toUpperCase();
      if (mk) masterMap[mk] = dMaster[mi];
    }
    var out = [];
    var q   = filtro ? filtro.toLowerCase().trim() : "";

    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var hoy  = new Date();
    var lastCountMap = {};
    for (var m = 1; m < dMov.length; m++) {
      var mTipo = String(dMov[m][M.TIPO] || '');
      if (mTipo.indexOf('AJUSTE') === -1 && mTipo.indexOf('CONTEO') === -1) continue;
      var mCod = String(dMov[m][M.CODIGO] || '').trim().toUpperCase();
      if (!mCod) continue;
      var mFecha = dMov[m][M.FECHA];
      if (!(mFecha instanceof Date)) continue;
      if (!lastCountMap[mCod] || mFecha > lastCountMap[mCod].fecha) {
        lastCountMap[mCod] = { fecha: mFecha, dias: Math.floor((hoy - mFecha) / 86400000) };
      }
    }

    var dUbic = [];
    var ubicMultiMap = _getCarmenUbicMap(); // tab UBICACIONES en Carmen

    // TABLA_POSICIONES: bins WMS con BIN_ID y TIPO_ALMACEN
    var dPos = [];
    try { dPos = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES); } catch(ep) {}
    var TP = SCHEMA.TABLA_POSICIONES;
    var binMap = {};    // SKU → [{ binId, cantidad, tipoAlmacen }]
    var binTotalMap = {}; // SKU → suma de cantidades en bins
    for (var p = 1; p < dPos.length; p++) {
      var pSku = String(dPos[p][TP.SKU] || '').trim().toUpperCase();
      if (!pSku) continue;
      var pCant = parseInt(dPos[p][TP.CANTIDAD]) || 0;
      if (!binMap[pSku]) { binMap[pSku] = []; binTotalMap[pSku] = 0; }
      binMap[pSku].push({
        binId:       String(dPos[p][TP.BIN_ID]       || ''),
        cantidad:    pCant,
        tipoAlmacen: String(dPos[p][TP.TIPO_ALMACEN] || '')
      });
      binTotalMap[pSku] += pCant;
    }

    var diasMap = calcularDiasDeStockPorSku();

    // Mapa SKU → unidades reservadas (RESERVAS_STOCK activas)
    var reservadoMap = {};
    try {
      var hojaRes = getSheet(SCHEMA.SHEETS.RESERVAS);
      if (hojaRes) {
        var dResS = getSheetValues(hojaRes);
        var RS    = SCHEMA.RESERVAS_STOCK;
        for (var rsi = 1; rsi < dResS.length; rsi++) {
          if (String(dResS[rsi][RS.ESTADO] || '') !== 'Activa') continue;
          var rsSku  = String(dResS[rsi][RS.SKU]      || '').trim().toUpperCase();
          var rsCant = parseInt(dResS[rsi][RS.CANTIDAD]) || 0;
          if (rsSku && rsCant > 0) reservadoMap[rsSku] = (reservadoMap[rsSku] || 0) + rsCant;
        }
      }
    } catch(eR) { Logger.log('cargarStock reservado: ' + eR); }

    // Mapa SKU → unidades en camino desde compras DJI activas (excluye Borrador y En depósito)
    var enCaminoMap    = {};
    var enCaminoCasMap = {};
    try {
      var dCAS = getSheetValues(SCHEMA.SHEETS.COMPRAS);
      var casEstadoMap = {};
      var casEtaMap    = {};
      var estadosExcluidos = { 'En depósito': true };
      for (var ci = 1; ci < dCAS.length; ci++) {
        var casId  = String(dCAS[ci][0] || '').trim().toUpperCase();
        var casEst = String(dCAS[ci][2] || '').trim();
        if (casId) {
          casEstadoMap[casId] = casEst;
          casEtaMap[casId]    = _fmtFecha(dCAS[ci][SCHEMA.COMPRAS_DJI.ETA]);
        }
      }
      var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
      if (hojaCD) {
        var dCD = getSheetValues(hojaCD);
        var CD  = SCHEMA.COMPRAS_DETALLE;
        for (var ci2 = 1; ci2 < dCD.length; ci2++) {
          var cdCas  = String(dCD[ci2][CD.ID_CAS] || '').trim().toUpperCase();
          var cdSku  = String(dCD[ci2][CD.SKU]    || '').trim().toUpperCase();
          var cdPed  = parseInt(dCD[ci2][CD.CANTIDAD_PEDIDA])   || 0;
          var cdRec  = parseInt(dCD[ci2][CD.CANTIDAD_RECIBIDA]) || 0;
          var cdEst  = casEstadoMap[cdCas] || '';
          if (!cdSku || !cdEst || estadosExcluidos[cdEst] || cdEst.indexOf('Borrador') !== -1) continue;
          var pendiente = Math.max(0, cdPed - cdRec);
          if (pendiente > 0) {
            enCaminoMap[cdSku] = (enCaminoMap[cdSku] || 0) + pendiente;
            if (!enCaminoCasMap[cdSku]) enCaminoCasMap[cdSku] = [];
            enCaminoCasMap[cdSku].push({ cas: cdCas, estado: cdEst, cant: pendiente, eta: casEtaMap[cdCas] || '' });
          }
        }
      }
    } catch(eC) { Logger.log('cargarStock enCamino: ' + eC); }

    // Cols Carmen STOCK: 0=PN, 1=Descripción, 2=Stock Actual, 3=Modelo, 4=Serie
    for (var i = 1; i < dCarmen.length; i++) {
      var cod  = String(dCarmen[i][0] || '').trim();
      var desc = String(dCarmen[i][1] || '').trim();
      if (!cod || !desc) continue;
      if (q && cod.toLowerCase().indexOf(q) === -1 && desc.toLowerCase().indexOf(q) === -1) continue;
      var codKey = cod.toUpperCase();
      // Stock: Carmen col 2 es la fuente; fallback bins WMS
      var act = (binTotalMap[codKey] !== undefined) ? binTotalMap[codKey] : (parseInt(dCarmen[i][2]) || 0);
      // Metadata desde STOCK_REPUESTOS (si existe)
      var mf  = masterMap[codKey];
      var min = mf ? (parseInt(mf[S.STOCK_MINIMO]) || 0) : 0;
      var cat = mf ? String(mf[S.CATEGORIA]  || '') : '';
      var ubi = mf ? String(mf[S.UBICACION]  || '') : '';
      var mod = mf ? String(mf[S.MODELOS]    || '') : String(dCarmen[i][4] || '');
      var uEnt = mf ? _fmtFecha(mf[S.ULTIMA_ENTRADA]) : '—';
      var uSal = mf ? _fmtFecha(mf[S.ULTIMA_SALIDA])  : '—';
      var reqSN = mf ? mf[S.REQUIERE_SN] === true : false;
      var estado = act <= 0 ? 'CRÍTICO' : (act <= min ? 'BAJO' : 'OK');
      var cEntry   = lastCountMap[codKey];
      var diasStock = diasMap.hasOwnProperty(codKey) ? diasMap[codKey] : undefined;
      out.push({
        fila: i+1, codigo: cod, descripcion: desc,
        stockActual: act, stockMinimo: min,
        categoria: cat, ubicacion: ubi, modelos: mod,
        ultimaEntrada: uEnt, ultimaSalida: uSal,
        requireSN: reqSN,
        estado: estado,
        fechaConteo: cEntry ? _fmtFecha(cEntry.fecha) : '—',
        diasDesdeConteo: cEntry ? cEntry.dias : 9999,
        ubicaciones: ubicMultiMap[codKey] || [],
        bins: binMap[codKey] || [],
        diasStock: (diasStock === null || diasStock === undefined) ? null : diasStock,
        enCamino:    enCaminoMap[codKey]    || 0,
        enCaminoCas: enCaminoCasMap[codKey] || [],
        reservado:   reservadoMap[codKey]   || 0
      });
    }
    return out;
  } catch(e) { return []; }
}

function ajustarInventario(codigo, cantNueva, motivo, operador, ubicacion) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.STOCK);
    var d    = getSheetValues(hoja);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toUpperCase() !== String(codigo).trim().toUpperCase()) continue;
      var anterior = parseInt(d[i][2])||0;
      var nueva    = parseInt(cantNueva)||0;
      var diff     = nueva - anterior;
      // Stock Actual en Carmen es fórmula — no escribir ahí; solo registrar el movimiento
      _registrarMovimiento("AJUSTE_INVENTARIO", codigo, String(d[i][1]), diff, nueva,
        "Ajuste: "+motivo, operador||"Sistema");
      // Origen LEGIBLE en la línea de Carmen (Recibidos/Entregados): motivo + operador
      var _origenAjuste = (motivo ? 'Ajuste: ' + motivo : 'Ajuste') + (operador ? ' \xb7 ' + operador : '');
      _registrarMovimientoCarmen(codigo, String(d[i][1]), ubicacion, diff, _origenAjuste);
      return { ok: true, anterior: anterior, nueva: nueva };
    }
    return { ok: false, msg: "Código no encontrado" };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function editarConfigStock(fila, minimo, categoria, ubicacion, requireSN) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.STOCK);
    var f = parseInt(fila);
    var catStr = String(categoria||'').trim().toUpperCase();
    if (catStr && !_catValida(catStr)) return { ok: false, msg: 'Categoría inválida: debe ser texto (A/B/C/D), no un número.' };
    hoja.getRange(f, 4).setValue(parseInt(minimo)||0);
    hoja.getRange(f, 5).setValue(catStr);
    hoja.getRange(f, 6).setValue(String(ubicacion));
    if (requireSN !== undefined) {
      hoja.getRange(f, SCHEMA.STOCK_REPUESTOS.REQUIERE_SN + 1).setValue(requireSN === true || requireSN === 'true');
    }
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function importarStockInicial(filas) {
  // filas = array de { codigo, descripcion, stockActual, modelos }
  try {
    var hoja = getSheet(SCHEMA.SHEETS.STOCK);
    var d    = getSheetValues(hoja);
    var codigosExistentes = {};
    for (var e = 1; e < d.length; e++) codigosExistentes[String(d[e][0]).trim().toUpperCase()] = e+1;
    var nuevos = 0, actualizados = 0;
    for (var i = 0; i < filas.length; i++) {
      var f   = filas[i];
      var cod = String(f.codigo||"").trim().toUpperCase();
      if (!cod) continue;
      if (codigosExistentes[cod]) {
        hoja.getRange(codigosExistentes[cod], 3).setValue(parseInt(f.stockActual)||0);
        actualizados++;
      } else {
        hoja.appendRow([cod, f.descripcion||"", parseInt(f.stockActual)||0, 0, "", "", f.modelos||"", "", "", false]);
        nuevos++;
      }
    }
    return { ok: true, nuevos: nuevos, actualizados: actualizados };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function obtenerMovimientos(codigo, limite) {
  try {
    var d    = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var out  = [], max = limite || 50;
    var codB = codigo ? String(codigo).trim().toUpperCase() : null;
    for (var i = d.length - 1; i >= 1 && out.length < max; i--) {
      if (codB && String(d[i][SCHEMA.MOVIMIENTOS_STOCK.CODIGO]).trim().toUpperCase() !== codB) continue;
      out.push({
        fila:           i + 1,
        fecha:          _fmtFecha(d[i][SCHEMA.MOVIMIENTOS_STOCK.FECHA]),
        tipo:           String(d[i][SCHEMA.MOVIMIENTOS_STOCK.TIPO]         || ""),
        codigo:         String(d[i][SCHEMA.MOVIMIENTOS_STOCK.CODIGO]       || ""),
        descripcion:    String(d[i][SCHEMA.MOVIMIENTOS_STOCK.DESCRIPCION]  || ""),
        cantidad:       parseInt(d[i][SCHEMA.MOVIMIENTOS_STOCK.CANTIDAD])  || 0,
        stockResultante:parseInt(d[i][SCHEMA.MOVIMIENTOS_STOCK.STOCK_RESULTANTE]) || 0,
        referencia:     String(d[i][SCHEMA.MOVIMIENTOS_STOCK.REFERENCIA]   || "—"),
        operador:       String(d[i][SCHEMA.MOVIMIENTOS_STOCK.OPERADOR]     || ""),
        observaciones:  String(d[i][SCHEMA.MOVIMIENTOS_STOCK.OBSERVACIONES]|| "")
      });
    }
    return out;
  } catch(e) { Logger.log("obtenerMovimientos: " + e); return []; }
}

// Alias para el tab de auditoría del frontend
function obtenerUltimosLogs(limite) {
  return obtenerMovimientos(null, limite || 200);
}

// Genera un vale PDF firmable a partir de una fila de MOVIMIENTOS_STOCK.
// fila: número de fila 1-based en la hoja (fila 1 = encabezado, fila 2 = primer registro).
function generarValeMovimiento(fila) {
  var S    = SCHEMA.MOVIMIENTOS_STOCK;
  var d    = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
  var row  = d[fila - 1];
  if (!row || fila < 2) {
    return { success: false, error: 'Registro no encontrado (fila ' + fila + ').' };
  }

  var fecha       = row[S.FECHA];
  var tipo        = String(row[S.TIPO]           || '');
  var codigo      = String(row[S.CODIGO]         || '');
  var descripcion = String(row[S.DESCRIPCION]    || '');
  var cantidad    = row[S.CANTIDAD];
  var stockResult = row[S.STOCK_RESULTANTE];
  var referencia  = String(row[S.REFERENCIA]     || '—');
  var operador    = String(row[S.OPERADOR]        || '—');
  var observ      = String(row[S.OBSERVACIONES]  || '');
  var fechaStr    = _fmtFecha(fecha instanceof Date ? fecha : new Date(fecha));
  var ahora       = _fmtFecha(new Date());

  try {
    var docName = 'Vale_' + codigo + '_' + fechaStr.replace(/\//g, '-');
    var doc     = DocumentApp.create(docName);
    var body    = doc.getBody();
    body.clear();

    // Membrete
    body.appendParagraph('BIDCOMAGRO')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('Servicio Técnico Oficial DJI Agras · Stock Manager')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setFontSize(9);
    body.appendHorizontalRule();

    body.appendParagraph('VALE DE MOVIMIENTO DE STOCK')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');

    // Tabla de datos
    var tableData = [
      ['FECHA DEL MOVIMIENTO', fechaStr],
      ['TIPO',                  tipo],
      ['CÓDIGO (SKU)',           codigo],
      ['DESCRIPCIÓN',           descripcion],
      ['CANTIDAD',              String(cantidad)],
      ['STOCK RESULTANTE',      String(stockResult)],
      ['REFERENCIA / OT',       referencia],
      ['RESPONSABLE',           operador]
    ];
    if (observ) tableData.push(['OBSERVACIONES', observ]);

    var table = body.appendTable(tableData);
    for (var r = 0; r < table.getNumRows(); r++) {
      table.getCell(r, 0)
           .setFontSize(9)
           .setBold(true)
           .setBackgroundColor('#f0f2f5');
      table.getCell(r, 1).setFontSize(11);
    }

    body.appendParagraph('');
    body.appendParagraph(
      '_______________________          _______________________          _______________________'
    ).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('  Responsable de almacén                  Autorizado por                          Recibido / Firma')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setFontSize(9);
    body.appendParagraph('');
    body.appendParagraph('Generado el ' + ahora + '  ·  BIDCOMAGRO  ·  Sistema de Gestión de Stock')
        .setFontSize(8)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();

    var pdfBlob = DriveApp.getFileById(doc.getId())
                    .getAs(MimeType.PDF)
                    .setName(docName + '.pdf');
    var pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Eliminar el Doc temporal
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return { success: true, url: pdfFile.getUrl(), nombre: docName };

  } catch(e) {
    Logger.log('generarValeMovimiento [fila=' + fila + ']: ' + e);
    return { success: false, error: e.message };
  }
}

function obtenerAlertasStockCritico() {
  var S         = SCHEMA.STOCK_REPUESTOS;
  var data      = getSheetValues(SCHEMA.SHEETS.STOCK);
  var carmenMap = _getCarmenStockMap();

  // Cruzar con OTs en "Espera de repuestos" para contar OTs bloqueadas por SKU
  var otsEspera = {};
  try {
    var dOT = getSheetValues(SCHEMA.SHEETS.OT);
    for (var o = 1; o < dOT.length; o++) {
      var estado = String(dOT[o][SCHEMA.OT.ESTADO]||"");
      if (estado !== "Espera de repuestos") continue;
      var repStr = String(dOT[o][SCHEMA.OT.REPUESTOS]||"").trim();
      if (!repStr || repStr === "Sin consumo de repuestos") continue;
      var partes = repStr.split(" ; ");
      for (var rp = 0; rp < partes.length; rp++) {
        var skuRaw = String(partes[rp].split(" | ")[0]).trim().toUpperCase();
        if (skuRaw) otsEspera[skuRaw] = (otsEspera[skuRaw]||0) + 1;
      }
    }
  } catch(e2) { Logger.log("obtenerAlertasStockCritico OTs: " + e2); }

  var alertas = [];
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var sku    = String(row[S.CODIGO]||"").trim().toUpperCase();
    var actual = (carmenMap[sku] !== undefined) ? carmenMap[sku] : Number(row[S.STOCK_ACTUAL]);
    var minimo = Number(row[S.STOCK_MINIMO]);
    if (!sku) continue;
    if (actual <= minimo) {
      var bloqueadas = otsEspera[sku] || 0;
      alertas.push({
        sku:       sku,
        nombre:    row[S.DESCRIPCION],
        actual:    actual,
        minimo:    minimo,
        categoria: String(row[S.CATEGORIA]||""),
        estado:    actual === 0 ? 'QUIEBRE' : 'BAJO',
        bloqueadas: bloqueadas   // OTs que esperan este SKU
      });
    }
  }
  // Ordenar: quiebres primero, luego por OTs bloqueadas desc
  alertas.sort(function(a, b) {
    if (a.estado !== b.estado) return a.estado === 'QUIEBRE' ? -1 : 1;
    return b.bloqueadas - a.bloqueadas;
  });
  return alertas;
}

// ============================================================
//  CREAR REPUESTO MANUAL (para piezas DJI no compradas aún)
// ============================================================
function crearRepuesto(data) {
  try {
    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    var hojaRep   = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    var dStock    = getSheetValues(hojaStock);
    var cod       = String(data.codigo||"").trim().toUpperCase();
    if (!cod) return { ok: false, msg: "Código requerido" };

    // Verificar si ya existe en STOCK_REPUESTOS
    for (var i = 1; i < dStock.length; i++) {
      if (String(dStock[i][0]).trim().toUpperCase() === cod)
        return { ok: false, msg: "El código ya existe en el stock" };
    }

    // Agregar a STOCK_REPUESTOS con stock 0
    hojaStock.appendRow([
      cod,
      String(data.descripcion||""),
      0,  // stock inicial 0
      parseInt(data.minimo)||0,
      String(data.categoria||"D").toUpperCase(),
      String(data.ubicacion||""),
      String(data.modelos||""),
      "", ""
    ]);

    // También agregar a DB_REPUESTOS si no existe
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      var existeEnDB = false;
      for (var j = 1; j < dRep.length; j++) {
        if (String(dRep[j][1]||"").trim().toUpperCase() === cod) { existeEnDB=true; break; }
      }
      if (!existeEnDB) {
        hojaRep.appendRow([
          "", cod,
          String(data.descripcion||""),
          String(data.modelos||""),
          "", "", parseFloat(data.precioFOB)||0
        ]);
      }
    }

    return { ok: true };
  } catch(e) {
    Logger.log("crearRepuesto: "+e);
    return { ok: false, msg: e.toString() };
  }
}

function generarHTMLPendientes() {
  try {
    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
    var hojaSolicitudes = ss.getSheetByName(SCHEMA.SHEETS.SOLICITUDES);

    if (!hojaSolicitudes) return "Error: No se encontró la hoja " + SCHEMA.SHEETS.SOLICITUDES;

    var datos = getSheetValues(hojaSolicitudes);
    var pendientes = datos.filter(function(fila, index) {
      return index > 0 && String(fila[8]).trim().toLowerCase() === "pendiente";
    });

    if (pendientes.length === 0) return "VACIO";

    var html = '<html><head><style>' +
               'body { font-family: Arial, sans-serif; padding: 20px; }' +
               'table { width: 100%; border-collapse: collapse; margin-top: 15px; }' +
               'th, td { border: 1px solid #000; padding: 6px; font-size: 11px; text-align: left; }' +
               'th { background-color: #eee; }' +
               '.urgente { font-weight: bold; text-decoration: underline; }' +
               '@media print { .no-print { display: none; } }' +
               '</style></head><body>' +
               '<h2>PICKING LIST - BIDCOMAGRO</h2>' +
               '<p>Fecha: ' + new Date().toLocaleString() + '</p>' +
               '<table><thead><tr>' +
               '<th>REF / OT</th><th>RESELLER</th><th>CÓDIGO</th><th>DESCRIPCIÓN</th><th>CANT.</th><th>URG.</th>' +
               '</tr></thead><tbody>';

    pendientes.forEach(function(p) {
      html += '<tr>' +
              '<td>' + p[2] + '</td>' +
              '<td>' + p[3] + '</td>' +
              '<td><strong>' + p[4] + '</strong></td>' +
              '<td>' + p[5] + '</td>' +
              '<td style="text-align:center">' + p[6] + '</td>' +
              '<td>' + p[9] + '</td>' +
              '</tr>';
    });

    html += '</tbody></table></body></html>';
    return html;

  } catch (e) {
    return "Error: " + e.toString();
  }
}

// ============================================================
//  GENERADOR DE PEDIDO DJI — algoritmo inteligente
// ============================================================
function calcularPedidoDJI() {
  try {
    var dStr      = getSheetValues(SCHEMA.SHEETS.STOCK);
    var carmenMap = _getCarmenStockMap();
    var dMov    = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var dRep    = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var dDet    = getSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    var dComp   = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var dReserv = getSheetValues(SCHEMA.SHEETS.RESERVAS);
    var hace90  = new Date(new Date().getTime() - 90 * 86400000);

    // Lead time estimado en días según estado — valores históricos si hay datos, si no, fijos
    var _LEAD_FIJO = {
      'Pendiente':        120,
      'Comprado':          90,
      'Pagado':            75,
      'Envío confirmado':  60,
      'Forwarder HK':      45,
      'En vuelo':          20,
      'En aduana':         10,
      'En depósito':        3
    };
    var _leadHist = {};
    try { _leadHist = calcularLeadTimeHistorico(); } catch(elh) {}
    var LEAD_TIME_DIAS = {};
    var _estados = Object.keys(_LEAD_FIJO);
    for (var le = 0; le < _estados.length; le++) {
      var _est = _estados[le];
      LEAD_TIME_DIAS[_est] = (_leadHist[_est] && _leadHist[_est] > 0) ? _leadHist[_est] : _LEAD_FIJO[_est];
    }

    // Estado y lead time estimado por CAS ID
    var ESTADOS_ACTIVOS = { 'Pendiente':1,'Comprado':1,'Pagado':1,'Envío confirmado':1,'Forwarder HK':1,'En vuelo':1,'En aduana':1,'En depósito':1 };
    var casEstadoMap = {};
    for (var c = 1; c < dComp.length; c++) {
      var casId = String(dComp[c][SCHEMA.COMPRAS_DJI.ID_CAS]||'').trim().toUpperCase();
      var estC  = String(dComp[c][SCHEMA.COMPRAS_DJI.ESTADO]||'');
      if (casId && ESTADOS_ACTIVOS[estC]) {
        casEstadoMap[casId] = { estado: estC, leadTime: LEAD_TIME_DIAS[estC] || 60 };
      }
    }

    // En camino por SKU + lead time mínimo esperado (la compra más avanzada gana)
    var enCaminoMap     = {};
    var enCaminoLeadMap = {};
    for (var d = 1; d < dDet.length; d++) {
      var casRef  = String(dDet[d][SCHEMA.COMPRAS_DETALLE.ID_CAS]||'').trim().toUpperCase();
      var casMeta = casEstadoMap[casRef];
      if (!casMeta) continue;
      var skuD   = String(dDet[d][SCHEMA.COMPRAS_DETALLE.SKU]||'').trim().toUpperCase();
      var pedida = parseInt(dDet[d][SCHEMA.COMPRAS_DETALLE.CANTIDAD_PEDIDA])||0;
      var recib  = parseInt(dDet[d][SCHEMA.COMPRAS_DETALLE.CANTIDAD_RECIBIDA])||0;
      var pend   = Math.max(0, pedida - recib);
      if (!skuD || pend <= 0) continue;
      enCaminoMap[skuD] = (enCaminoMap[skuD]||0) + pend;
      if (enCaminoLeadMap[skuD] === undefined || casMeta.leadTime < enCaminoLeadMap[skuD]) {
        enCaminoLeadMap[skuD] = casMeta.leadTime;
      }
    }

    // Reservado por SKU: reservas vigentes
    var reservadoMap = {};
    var S_RES = SCHEMA.RESERVAS_STOCK;
    for (var rv = 1; rv < dReserv.length; rv++) {
      var estR = String(dReserv[rv][S_RES.ESTADO]||'').trim().toUpperCase();
      if (estR === 'DESPACHADA' || estR === 'CANCELADA') continue;
      var skuR  = String(dReserv[rv][S_RES.SKU]||'').trim().toUpperCase();
      var cantR = parseInt(dReserv[rv][S_RES.CANTIDAD])||0;
      if (skuR && cantR > 0) reservadoMap[skuR] = (reservadoMap[skuR]||0) + cantR;
    }

    // Modelos activos configurados
    var modelosActivos = getModelosActivos();

    // DB_REPUESTOS: mapa de metadata + set de SKUs que aplican a modelos activos
    var dbRepMap        = {};
    var skusModeloActivo = {};
    for (var r = 1; r < dRep.length; r++) {
      var codR     = String(dRep[r][1]||'').trim().toUpperCase();
      if (!codR) continue;
      var modelosR = String(dRep[r][3]||'').trim();
      var fobR     = parseFloat(String(dRep[r][6]||'0').replace(/[^0-9.,\-]/g,'').replace(',','.')) || 0;
      dbRepMap[codR] = { desc: String(dRep[r][2]||'').trim(), modelos: modelosR, fob: fobR, cat: String(dRep[r][4]||'').trim() };
      if (modelosActivos.length > 0 && modelosR) {
        var mUp = modelosR.toUpperCase();
        for (var ma = 0; ma < modelosActivos.length; ma++) {
          if (mUp.indexOf(String(modelosActivos[ma]).toUpperCase()) !== -1) {
            skusModeloActivo[codR] = true;
            break;
          }
        }
      }
    }

    // Salidas últimos 90 días por código
    var salidas90 = {};
    for (var m = 1; m < dMov.length; m++) {
      var tipo = String(dMov[m][1]||'');
      if (tipo !== 'SALIDA_DESPACHO' && tipo !== 'EGRESO' && tipo.indexOf('SALIDA') === -1) continue;
      var fecha = dMov[m][0];
      if (!(fecha instanceof Date) || fecha < hace90) continue;
      var cod = String(dMov[m][2]||'').trim().toUpperCase();
      salidas90[cod] = (salidas90[cod]||0) + Math.abs(parseInt(dMov[m][4])||0);
    }

    var CRIT_ORD = { 'critico':0, 'urgente':1, 'bajo':2, 'modelo':3, 'restock':4, 'ok':5 };
    var out = [];
    var skusEnStock = {};

    for (var i = 1; i < dStr.length; i++) {
      var f        = dStr[i];
      var cod2     = String(f[0]||'').trim().toUpperCase();
      if (!cod2) continue;
      skusEnStock[cod2] = true;

      var stockAct = (carmenMap[cod2] !== undefined) ? carmenMap[cod2] : (parseInt(f[2])||0);
      var stockMin = parseInt(f[3])||0;
      var sal90    = salidas90[cod2]||0;
      var enCam    = enCaminoMap[cod2]||0;
      var reserv   = reservadoMap[cod2]||0;
      var isModelo = skusModeloActivo[cod2] || false;

      // Burn rate diario y stock físico disponible (sin reservas)
      var burnRateDia = sal90 / 90;
      var stockFisico = Math.max(0, stockAct - reserv);

      // Días de stock restantes con el stock físico actual
      var diasStock = burnRateDia > 0 ? stockFisico / burnRateDia : Infinity;

      // Lead time de la compra más próxima a llegar
      var leadTime = enCaminoLeadMap[cod2] !== undefined ? enCaminoLeadMap[cod2] : 999;

      // Target 2 meses (sal90 = 3 meses → × 2/3 = 2 meses)
      var target2M    = Math.ceil(sal90 * 2 / 3);
      var stockFuturo = stockFisico + enCam;

      // URGENTE: el stock físico se agota antes de que llegue la compra en camino
      var seAgota = burnRateDia > 0 && enCam > 0 && diasStock < leadTime;
      // Unidades necesarias para cubrir el período hasta que llegue la compra
      var coberturaInmediata = seAgota ? Math.max(0, Math.ceil(burnRateDia * leadTime) - stockFisico) : 0;

      // Gap de largo plazo: lo que falta para tener 2 meses DESPUÉS de recibir lo que viene
      var gapLargoPlazo = Math.max(0, target2M - stockFuturo);

      // Si es modelo activo, asegurar al menos 1 unidad disponible en todo momento
      var modeloGap = (isModelo && stockFuturo < 1) ? 1 : 0;

      var sugerido, criticidad, motivo;

      if (sal90 > 0 || isModelo) {
        sugerido = Math.max(gapLargoPlazo, coberturaInmediata, modeloGap);

        if (stockFisico <= 0 && enCam === 0) {
          criticidad = 'critico';
          motivo     = 'Sin stock';
        } else if (seAgota) {
          criticidad = 'urgente';
          motivo     = 'Se agota en ' + Math.ceil(diasStock) + 'd — llega en ~' + leadTime + 'd';
        } else if (stockAct <= stockMin && stockMin > 0) {
          criticidad = 'bajo';
          motivo     = 'Bajo mínimo';
        } else if (sugerido > 0) {
          criticidad = 'restock';
          motivo     = 'Restock 2M';
        } else {
          criticidad = 'ok';
          motivo     = 'Cubierto';
        }
      } else {
        if (stockMin > 0 && stockAct < stockMin) {
          sugerido   = stockMin - stockAct;
          criticidad = 'bajo';
          motivo     = 'Bajo mínimo';
        } else {
          sugerido   = 0;
          criticidad = 'ok';
          motivo     = '';
        }
      }

      var dbInfo = dbRepMap[cod2] || {};
      var fob = dbInfo.fob || 0;
      out.push({
        codigo:      cod2,
        descripcion: String(f[1]||''),
        stockActual: stockAct,
        stockMinimo: stockMin,
        enCamino:    enCam,
        reservado:   reserv,
        salidas90:   sal90,
        target2M:    target2M,
        sugerido:    sugerido,
        precioFOB:   fob,
        totalFOB:    fob * sugerido,
        categoria:   String(f[4]||''),
        criticidad:  criticidad,
        motivo:      motivo,
        incluir:     sugerido > 0
      });
    }

    // Modelos activos: agregar repuestos del catálogo que nunca tuvieron stock
    if (modelosActivos.length > 0) {
      for (var db = 1; db < dRep.length; db++) {
        var dbCod = String(dRep[db][1]||'').trim().toUpperCase();
        if (!dbCod || skusEnStock[dbCod] || !skusModeloActivo[dbCod]) continue;

        var enCamDB    = enCaminoMap[dbCod] || 0;
        var sugeridoDB = Math.max(0, 1 - enCamDB);
        var fobDB      = parseFloat(String(dRep[db][6]||'0').replace(/[^0-9.,\-]/g,'').replace(',','.')) || 0;
        out.push({
          codigo:      dbCod,
          descripcion: String(dRep[db][2]||'').trim(),
          stockActual: 0,
          stockMinimo: 0,
          enCamino:    enCamDB,
          reservado:   0,
          salidas90:   0,
          target2M:    1,
          sugerido:    sugeridoDB,
          precioFOB:   fobDB,
          totalFOB:    fobDB * sugeridoDB,
          categoria:   String(dRep[db][4]||'').trim(),
          criticidad:  enCamDB > 0 ? 'ok' : 'modelo',
          motivo:      enCamDB > 0 ? 'Modelo activo · en camino' : 'Modelo activo · sin stock',
          incluir:     sugeridoDB > 0
        });
      }
    }

    out.sort(function(a,b){
      var diff = (CRIT_ORD[a.criticidad]||5) - (CRIT_ORD[b.criticidad]||5);
      if (diff !== 0) return diff;
      return (a.categoria||'Z').localeCompare(b.categoria||'Z');
    });
    return out;
  } catch(e) { Logger.log('calcularPedidoDJI: '+e); return []; }
}

// Previsión de demanda: calcula días de stock restante por SKU basado en
// rotación de los últimos 90 días. Retorna mapa { SKU: diasRestantes }
// diasRestantes = null si no hubo salidas (sin rotación = stock sin consumo)
function calcularDiasDeStockPorSku() {
  try {
    var dMov      = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var dStr      = getSheetValues(SCHEMA.SHEETS.STOCK);
    var carmenMap = _getCarmenStockMap();
    var hace90 = new Date(new Date().getTime() - 90 * 86400000);
    var S = SCHEMA.STOCK_REPUESTOS;

    var salidas90 = {};
    for (var m = 1; m < dMov.length; m++) {
      if (String(dMov[m][1]||"").indexOf("SALIDA") === -1) continue;
      var fecha = dMov[m][0];
      if (!(fecha instanceof Date) || fecha < hace90) continue;
      var cod = String(dMov[m][2]||"").trim().toUpperCase();
      salidas90[cod] = (salidas90[cod]||0) + Math.abs(parseInt(dMov[m][4])||0);
    }

    var resultado = {};
    for (var i = 1; i < dStr.length; i++) {
      var sku = String(dStr[i][S.CODIGO]||"").trim().toUpperCase();
      if (!sku) continue;
      var stockActual = (carmenMap[sku] !== undefined) ? carmenMap[sku] : (parseInt(dStr[i][S.STOCK_ACTUAL])||0);
      var sal = salidas90[sku] || 0;
      if (sal === 0) {
        resultado[sku] = null; // sin rotación, no aplica semáforo
      } else {
        var rotDiaria = sal / 90;
        resultado[sku] = stockActual > 0 ? Math.round(stockActual / rotDiaria) : 0;
      }
    }
    return resultado;
  } catch(e) { Logger.log("calcularDiasDeStockPorSku: " + e); return {}; }
}

function registrarPedidoEnviado(items, cas, metodoPago, operador) {
  try {
    var res = registrarCAS(cas, metodoPago, operador);
    if (!res.ok) return res;

    var rows = [];
    var fecha = new Date();
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (!item.incluir || !item.ajustado) continue;
      // Preparamos la fila para MOVIMIENTOS_STOCK
      rows.push([fecha, "PEDIDO_DJI", item.codigo, item.descripcion, item.ajustado, 0, cas, operador, ""]);
    }

    if (rows.length > 0) {
      var hMov = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
      hMov.getRange(hMov.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ============================================================
//  TRANSFERENCIA INTERNA WMS
//  Mueve stock entre bins dentro de TABLA_POSICIONES.
//  No altera el saldo total en STOCK_REPUESTOS.
// ============================================================
function transferirStock(sku, origen, destino, cant) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var cantN  = parseInt(cant) || 0;
    if (cantN <= 0) throw new Error('La cantidad debe ser mayor a 0');
    var codKey = String(sku).trim().toUpperCase();
    var origenK = String(origen).trim();
    var destinoK = String(destino).trim();
    if (!origenK || !destinoK) throw new Error('Origen y destino son requeridos');
    if (origenK === destinoK) throw new Error('Origen y destino no pueden ser iguales');

    var hoja    = getSheet(SCHEMA.SHEETS.TABLA_POSICIONES);
    if (!hoja) throw new Error('Hoja TABLA_POSICIONES no existe. Ejecutá asegurarHojas() primero.');
    var TP      = SCHEMA.TABLA_POSICIONES;
    var datos   = hoja.getDataRange().getValues();

    var origenFila  = -1, origenCant = 0;
    var destinoFila = -1;

    for (var i = 1; i < datos.length; i++) {
      var fSku = String(datos[i][TP.SKU]    || '').trim().toUpperCase();
      var fBin = String(datos[i][TP.BIN_ID] || '').trim();
      if (fSku !== codKey) continue;
      if (fBin === origenK)  { origenFila  = i + 1; origenCant  = parseInt(datos[i][TP.CANTIDAD]) || 0; }
      if (fBin === destinoK) { destinoFila = i + 1; }
    }

    if (origenFila < 0) throw new Error('SKU "' + sku + '" no encontrado en bin "' + origenK + '"');
    if (origenCant < cantN) throw new Error('Stock insuficiente en ' + origenK + '. Disponible: ' + origenCant + ', solicitado: ' + cantN);

    // Decrementar origen
    hoja.getRange(origenFila, TP.CANTIDAD + 1).setValue(origenCant - cantN);

    // Incrementar destino o crear fila nueva
    if (destinoFila > 0) {
      var destCant = parseInt(datos[destinoFila - 1][TP.CANTIDAD]) || 0;
      hoja.getRange(destinoFila, TP.CANTIDAD + 1).setValue(destCant + cantN);
    } else {
      hoja.appendRow([codKey, destinoK, cantN, '']);
    }

    invalidateSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);

    _registrarMovimiento(
      'TRANSFERENCIA_INTERNA', codKey, '', 0,
      (origenCant - cantN),
      'De: ' + origenK + ' → A: ' + destinoK + ' · ' + cantN + ' u.',
      Session.getActiveUser().getEmail()
    );

    return { ok: true, sku: codKey, origen: origenK, destino: destinoK, cantidad: cantN };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

// ============================================================
//  CARGAR BINS DE UN SKU (para el modal WMS)
// ============================================================
function cargarBinsSKU(sku) {
  try {
    var codKey = String(sku).trim().toUpperCase();
    var dPos   = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);
    var TP     = SCHEMA.TABLA_POSICIONES;
    var bins   = [];
    for (var i = 1; i < dPos.length; i++) {
      if (String(dPos[i][TP.SKU] || '').trim().toUpperCase() !== codKey) continue;
      bins.push({
        binId:       String(dPos[i][TP.BIN_ID]       || ''),
        cantidad:    parseInt(dPos[i][TP.CANTIDAD])   || 0,
        tipoAlmacen: String(dPos[i][TP.TIPO_ALMACEN] || '')
      });
    }
    bins.sort(function(a, b) { return _binSortKey(a.binId) < _binSortKey(b.binId) ? -1 : 1; });
    return { ok: true, bins: bins };
  } catch(e) { return { ok: false, msg: e.toString(), bins: [] }; }
}

/**
 * Calcula el saldo exacto de un SKU leyendo la historia completa del ledger.
 */
function obtenerSaldoRealEnMemoria(sku) {
  var skuKey = String(sku || '').trim().toUpperCase();
  if (!skuKey) {
    Logger.log('[LEDGER] obtenerSaldoRealEnMemoria: SKU vacío');
    return { ok: false, saldo: 0, movimientos: 0, error: 'WMS_SKU_VACIO' };
  }
  var M = SCHEMA.MOVIMIENTOS_STOCK;
  try {
    var dMov      = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var saldo     = 0;
    var countMov  = 0;
    for (var i = 1; i < dMov.length; i++) {
      if (String(dMov[i][M.CODIGO] || '').trim().toUpperCase() !== skuKey) continue;
      // CANTIDAD ya viene con signo: positivo = entrada, negativo = salida
      var delta = parseFloat(dMov[i][M.CANTIDAD]) || 0;
      saldo    += delta;
      countMov++;
    }
    return { ok: true, saldo: Math.round(saldo), movimientos: countMov };
  } catch(e) {
    var snap = '[LEDGER][obtenerSaldoRealEnMemoria] sku=' + skuKey + ' err=' + e.message;
    Logger.log(snap);
    console.error(snap);
    return { ok: false, saldo: 0, movimientos: 0, error: e.message };
  }
}

/**
 * Punto de entrada ÚNICO para toda mutación de inventario.
 * Patrón: Lock → Verify (cache) → Append (ledger) → Sync (cache) → Release
 */
function registrarEventoLedgerSeguro(params) {
  var sku         = String(params.sku         || '').trim().toUpperCase();
  var delta       = parseFloat(params.delta)  || 0;
  var tipo        = String(params.tipo        || 'EGRESO').trim();
  var referencia  = String(params.referencia  || '');
  var operador    = String(params.operador    || '');
  var descripcion = String(params.descripcion || '');

  // Validaciones previas al lock (fail-fast sin costo de I/O)
  if (!sku)     return { ok: false, error: 'WMS_SKU_VACIO' };
  if (delta === 0) return { ok: false, error: 'WMS_DELTA_CERO' };

  var lock     = LockService.getScriptLock();
  var snapshot = {
    ts:         new Date().toISOString(),
    sku:        sku,
    delta:      delta,
    tipo:       tipo,
    referencia: referencia,
    operador:   operador
  };

  try {
    if (!lock.tryLock(15000)) {
      snapshot.error = 'WMS_LOCK_TIMEOUT';
      console.error('[LEDGER][LOCK_TIMEOUT] ' + JSON.stringify(snapshot));
      Logger.log('[LEDGER][LOCK_TIMEOUT] ' + JSON.stringify(snapshot));
      return { ok: false, error: 'WMS_LOCK_TIMEOUT', snapshot: snapshot };
    }

    // 1. Leer vista materializada — O(1) lookup, autoritativa bajo lock
    var hojaStr = getSheet(SCHEMA.SHEETS.STOCK);
    var dStr    = getSheetValues(hojaStr);
    var S       = SCHEMA.STOCK_REPUESTOS;
    var filaIdx = -1;
    for (var i = 1; i < dStr.length; i++) {
      if (String(dStr[i][S.CODIGO] || '').trim().toUpperCase() === sku) {
        filaIdx = i;
        break;
      }
    }
    if (filaIdx === -1) {
      snapshot.error = 'WMS_SKU_NOT_FOUND';
      console.error('[LEDGER][SKU_NOT_FOUND] ' + JSON.stringify(snapshot));
      Logger.log('[LEDGER][SKU_NOT_FOUND] ' + JSON.stringify(snapshot));
      return { ok: false, error: 'WMS_SKU_NOT_FOUND', snapshot: snapshot };
    }

    var saldoCache = parseInt(dStr[filaIdx][S.STOCK_ACTUAL]) || 0;
    if (!descripcion) descripcion = String(dStr[filaIdx][S.DESCRIPCION] || '');
    snapshot.saldoCache = saldoCache;

    // 2. Gatekeeper para egresos — usa la vista materializada (correcta bajo lock)
    if (delta < 0) {
      var cantSolicitada = Math.abs(delta);
      if (saldoCache < cantSolicitada) {
        snapshot.error         = 'WMS_INSUFFICIENT_FUNDS';
        snapshot.cantSolicitada = cantSolicitada;
        console.error('[LEDGER][INSUFFICIENT_FUNDS] ' + JSON.stringify(snapshot));
        Logger.log('[LEDGER][INSUFFICIENT_FUNDS] ' + JSON.stringify(snapshot));
        return { ok: false, error: 'WMS_INSUFFICIENT_FUNDS', snapshot: snapshot };
      }
    }

    // 3. Calcular saldo resultante
    var saldoNuevo = saldoCache + delta;
    var ahora      = new Date();

    // 4. Append al ledger — operación inmutable, fuente de verdad
    var hojaLedger = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
    if (!hojaLedger) {
      snapshot.error = 'WMS_LEDGER_WRITE_FAIL';
      console.error('[LEDGER][NO_LEDGER_SHEET] ' + JSON.stringify(snapshot));
      Logger.log('[LEDGER][NO_LEDGER_SHEET] ' + JSON.stringify(snapshot));
      return { ok: false, error: 'WMS_LEDGER_WRITE_FAIL', snapshot: snapshot };
    }
    hojaLedger.appendRow([
      ahora,
      tipo,
      sku,
      descripcion,
      delta,
      saldoNuevo,
      referencia,
      operador,
      ''
    ]);

    // 5. Sync write-through: actualizar vista materializada (la misma celda, no rebuild)
    hojaStr.getRange(filaIdx + 1, S.STOCK_ACTUAL + 1).setValue(saldoNuevo);
    if (delta > 0) hojaStr.getRange(filaIdx + 1, S.ULTIMA_ENTRADA + 1).setValue(ahora);
    if (delta < 0) hojaStr.getRange(filaIdx + 1, S.ULTIMA_SALIDA  + 1).setValue(ahora);

    // Espejo a Carmen: TODO ingreso escribe una línea en "Recibidos" y TODO egreso en
    // "Entregados" (mismo formato que compras y que WOS al despachar). El Stock Actual de
    // Carmen es fórmula (Recibidos/Entregados) → nunca se toca con setValue.
    // Origen LEGIBLE = referencia + operador. Ubicación tomada del propio ítem (STOCK_REPUESTOS).
    var _ubicItem      = String(dStr[filaIdx][S.UBICACION] || '').trim();
    var _origenLegible = (referencia || tipo || 'Movimiento') + (operador ? ' \xb7 ' + operador : '');
    _registrarMovimientoCarmen(sku, descripcion, _ubicItem, delta, _origenLegible);

    // 6. Invalidar caché de getSheetValues para que la próxima lectura sea fresca
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    SpreadsheetApp.flush();

    return { ok: true, saldoResultante: saldoNuevo };

  } catch(e) {
    snapshot.error   = e.message;
    snapshot.stack   = e.stack || '(no stack)';
    console.error('[LEDGER][EXCEPTION] ' + JSON.stringify(snapshot));
    Logger.log('[LEDGER][EXCEPTION] ' + JSON.stringify(snapshot));
    return { ok: false, error: e.message, snapshot: snapshot };
  } finally {
    // El lock SIEMPRE se libera — incluso si appendRow falla a mitad
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Reconcilia la vista materializada (STOCK_REPUESTOS) contra el ledger.
 * EJECUTAR: manualmente desde el editor, o vía trigger nocturno.
 */
function reconciliarCacheDesdeMovimientos() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Lock largo solo permitido en operación de mantenimiento
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var S    = SCHEMA.STOCK_REPUESTOS;
    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);

    // Acumular suma de deltas por SKU desde el ledger completo
    var saldosLedger = {};
    for (var i = 1; i < dMov.length; i++) {
      var cod = String(dMov[i][M.CODIGO] || '').trim().toUpperCase();
      if (!cod) continue;
      saldosLedger[cod] = (saldosLedger[cod] || 0) + (parseFloat(dMov[i][M.CANTIDAD]) || 0);
    }

    // Comparar contra la vista materializada y corregir drift
    var hojaStr  = getSheet(SCHEMA.SHEETS.STOCK);
    var dStr     = hojaStr.getDataRange().getValues();
    var discrepancias = [];
    var reconciliados = 0;

    for (var r = 1; r < dStr.length; r++) {
      var sku = String(dStr[r][S.CODIGO] || '').trim().toUpperCase();
      if (!sku || saldosLedger[sku] === undefined) continue;
      var saldoLedger = Math.max(0, Math.round(saldosLedger[sku]));
      var saldoCache  = parseInt(dStr[r][S.STOCK_ACTUAL]) || 0;
      if (saldoCache !== saldoLedger) {
        discrepancias.push({ sku: sku, cache: saldoCache, ledger: saldoLedger, drift: saldoLedger - saldoCache });
        dStr[r][S.STOCK_ACTUAL] = saldoLedger;
        reconciliados++;
      }
    }

    if (reconciliados > 0) {
      hojaStr.getDataRange().setValues(dStr);
      SpreadsheetApp.flush();
      invalidateSheetValues(SCHEMA.SHEETS.STOCK);
      Logger.log('[RECONCILIAR] ' + reconciliados + ' discrepancias corregidas: ' + JSON.stringify(discrepancias));
    } else {
      Logger.log('[RECONCILIAR] Caché consistente con el ledger. Sin drift detectado.');
    }

    return {
      ok:             true,
      reconciliados:  reconciliados,
      totalSkus:      Object.keys(saldosLedger).length,
      discrepancias:  discrepancias
    };
  } catch(e) {
    var snap = '[RECONCILIAR][EXCEPTION] ' + e.message;
    console.error(snap);
    Logger.log(snap);
    return { ok: false, error: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ============================================================
//  CLASIFICACIÓN ABC — Reclasificación mensual por rotación
// ============================================================

function recalcularClasificacionABC() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var hoy   = new Date();
    var corte = new Date(hoy.getTime() - 90 * 86400000);
    var M     = SCHEMA.MOVIMIENTOS_STOCK;
    var S     = SCHEMA.STOCK_REPUESTOS;

    // 1. Sumar unidades de salida por SKU en los últimos 90 días
    var dMov  = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var consumo = {};
    for (var m = 1; m < dMov.length; m++) {
      var f = dMov[m];
      if (!(f[M.FECHA] instanceof Date) || f[M.FECHA] < corte) continue;
      var tipo = String(f[M.TIPO] || '');
      if (tipo !== 'SALIDA_DESPACHO' && tipo !== 'EGRESO') continue;
      var cod = String(f[M.CODIGO] || '').trim().toUpperCase();
      if (!cod) continue;
      consumo[cod] = (consumo[cod] || 0) + Math.abs(parseInt(f[M.CANTIDAD]) || 0);
    }

    // 2. Leer stock sin caché para garantizar datos frescos
    var hStock = getSheet(SCHEMA.SHEETS.STOCK);
    var dStock = hStock.getDataRange().getValues();
    var nRows  = dStock.length - 1;
    if (nRows <= 0) return { ok: false, msg: 'Sin datos en stock' };

    // 3. Ordenar SKUs por consumo descendente para el corte Pareto
    var skuList = [];
    for (var i = 1; i < dStock.length; i++) {
      var sku = String(dStock[i][S.CODIGO] || '').trim().toUpperCase();
      if (!sku) continue;
      skuList.push({ sku: sku, cons: consumo[sku] || 0 });
    }
    skuList.sort(function(a, b) { return b.cons - a.cons; });

    var totalConsumo = 0;
    for (var k = 0; k < skuList.length; k++) totalConsumo += skuList[k].cons;

    // 4. Asignar categorías acumulando el porcentaje del total
    var mapaCategoria = {};
    var acum = 0;
    for (var j = 0; j < skuList.length; j++) {
      var item = skuList[j];
      if (item.cons === 0 || totalConsumo === 0) { mapaCategoria[item.sku] = 'D'; continue; }
      acum += item.cons;
      var pct = acum / totalConsumo;
      mapaCategoria[item.sku] = pct <= 0.80 ? 'A' : pct <= 0.95 ? 'B' : 'C';
    }

    // 5. Construir columna de categorías y escribir en bulk (solo col 5)
    var colData = [];
    var conteo  = { A: 0, B: 0, C: 0, D: 0 };
    for (var r = 1; r < dStock.length; r++) {
      var skuR = String(dStock[r][S.CODIGO] || '').trim().toUpperCase();
      var cat  = (skuR && mapaCategoria[skuR]) ? mapaCategoria[skuR] : 'D';
      colData.push([cat]);
      conteo[cat] = (conteo[cat] || 0) + 1;
    }

    hStock.getRange(2, S.CATEGORIA + 1, colData.length, 1).setValues(colData);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);

    _logClasificacion(hoy, conteo, colData.length);

    return {
      ok:            true,
      actualizado:   colData.length,
      porCategoria:  conteo,
      fechaEjecucion: Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    };
  } catch(e) {
    Logger.log('recalcularClasificacionABC: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// Persiste el resultado de la reclasificación en la hoja LOGS.
function _logClasificacion(fecha, conteo, total) {
  try {
    var hLogs = getSheet('LOGS');
    if (!hLogs) {
      hLogs = getSS().insertSheet('LOGS');
      hLogs.appendRow(['Fecha', 'Evento', 'Detalle']);
    }
    hLogs.appendRow([
      fecha, 'RECLASIFICACION_ABC',
      'Total:' + total + ' A:' + (conteo.A||0) + ' B:' + (conteo.B||0) + ' C:' + (conteo.C||0) + ' D:' + (conteo.D||0)
    ]);
  } catch(e) { Logger.log('_logClasificacion: ' + e); }
}

// Retorna la última ejecución de la clasificación ABC desde LOGS, o null.
function obtenerUltimaClasificacion() {
  try {
    var hLogs = getSheet('LOGS');
    if (!hLogs) return null;
    var d = hLogs.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1]) !== 'RECLASIFICACION_ABC') continue;
      var det = String(d[i][2]);
      var out = {
        fecha: d[i][0] instanceof Date
          ? Utilities.formatDate(d[i][0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
          : String(d[i][0]),
        porCategoria: {}
      };
      var parts = det.split(' ');
      for (var p = 0; p < parts.length; p++) {
        var kv = parts[p].split(':');
        if (kv.length === 2 && 'ABCD'.indexOf(kv[0]) !== -1 && kv[0].length === 1) {
          out.porCategoria[kv[0]] = parseInt(kv[1]) || 0;
        }
      }
      return out;
    }
    return null;
  } catch(e) { return null; }
}

// Instala un trigger que ejecuta recalcularClasificacionABC() el día 1 de cada mes a las 02:00.
function configurarTriggerMensualABC() {
  var existentes = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existentes.length; i++) {
    if (existentes[i].getHandlerFunction() === 'recalcularClasificacionABC') {
      ScriptApp.deleteTrigger(existentes[i]);
    }
  }
  ScriptApp.newTrigger('recalcularClasificacionABC')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  Logger.log('Trigger mensual ABC configurado: día 1 de cada mes a las 02:00');
  return { ok: true };
}

// ============================================================
//  LEAD TIME HISTÓRICO REAL
// ============================================================
function calcularLeadTimeHistorico() {
  return getCachedData('leadtime_hist_v2', function() {
    var dComp = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var C = SCHEMA.COMPRAS_DJI;

    // estado → array de días que faltaban desde ese estado hasta depósito
    var acum = {
      'Comprado':        [],
      'Pagado':          [],
      'Envío confirmado':[],
      'Forwarder HK':    [],
      'En vuelo':        [],
      'En aduana':       []
    };

    var stateCols = [
      ['Comprado',         C.FECHA_COMPRADO],
      ['Pagado',           C.FECHA_PAGADO],
      ['Envío confirmado', C.FECHA_ENVIO],
      ['Forwarder HK',     C.FECHA_FORWARDER],
      ['En vuelo',         C.FECHA_VUELO],
      ['En aduana',        C.FECHA_ADUANA]
    ];

    for (var i = 1; i < dComp.length; i++) {
      var fDep = dComp[i][C.FECHA_DEPOSITO];
      if (!(fDep instanceof Date)) continue;
      for (var s = 0; s < stateCols.length; s++) {
        var fEst = dComp[i][stateCols[s][1]];
        if (!(fEst instanceof Date)) continue;
        var dias = Math.round((fDep - fEst) / 86400000);
        if (dias > 0 && dias < 400) acum[stateCols[s][0]].push(dias);
      }
    }

    var result = { _n: dComp.length - 1 };
    var estados = Object.keys(acum);
    for (var e = 0; e < estados.length; e++) {
      var arr = acum[estados[e]];
      result[estados[e]] = arr.length > 0
        ? Math.round(arr.reduce(function(a,b){return a+b;},0) / arr.length)
        : null;
    }
    return result;
  }, 3600);
}

// ============================================================
//  SLOTTING AUTOMÁTICO
// ============================================================
function calcularSlottingOptimo() {
  try {
    var dPos = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);
    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    var TP   = SCHEMA.TABLA_POSICIONES;
    var S    = SCHEMA.STOCK_REPUESTOS;

    // Mapa ABC y descripción desde STOCK_REPUESTOS
    var abcMap  = {};
    var descMap = {};
    for (var i = 1; i < dStr.length; i++) {
      var sku = String(dStr[i][S.CODIGO]||'').trim().toUpperCase();
      if (!sku) continue;
      abcMap[sku]  = String(dStr[i][S.CATEGORIA]||'D').trim().toUpperCase() || 'D';
      descMap[sku] = String(dStr[i][S.DESCRIPCION]||'');
    }

    // Construir lista de asignaciones bin → SKU
    var bins = [];
    var maxEstante = 0;
    for (var p = 1; p < dPos.length; p++) {
      var pSku  = String(dPos[p][TP.SKU]||'').trim().toUpperCase();
      var binId = String(dPos[p][TP.BIN_ID]||'').trim();
      if (!pSku || !binId) continue;
      var partes  = binId.split('-');
      var estante = parseInt(partes[0]) || 999;
      if (estante > maxEstante) maxEstante = estante;
      bins.push({
        sku:         pSku,
        descripcion: descMap[pSku] || pSku,
        binId:       binId,
        estante:     estante,
        cantidad:    parseInt(dPos[p][TP.CANTIDAD])||0,
        tipoAlmacen: String(dPos[p][TP.TIPO_ALMACEN]||''),
        categoria:   abcMap[pSku] || 'D'
      });
    }

    if (bins.length === 0) return { ok: true, sugerencias: [], msg: 'Sin datos en TABLA_POSICIONES.' };

    // Umbral: 1/3 del rango de estantes = zona prime
    var thresh = Math.max(2, Math.ceil(maxEstante / 3));

    // Encontrar items A/B mal ubicados y C/D ocupando zonas prime
    var sugerencias = [];
    var skusProcesados = {};

    for (var x = 0; x < bins.length; x++) {
      var item = bins[x];
      if (skusProcesados[item.sku]) continue;
      skusProcesados[item.sku] = true;

      var cat       = item.categoria;
      var isAltaRot = (cat === 'A' || cat === 'B');
      var isFar     = item.estante > thresh * 2;
      var isPrime   = item.estante <= thresh;

      if (isAltaRot && isFar) {
        // A/B en zona lejana — buscar C/D en zona prime para intercambiar
        var swap = null;
        for (var y = 0; y < bins.length; y++) {
          if (!skusProcesados[bins[y].sku] && bins[y].estante <= thresh &&
              (bins[y].categoria === 'C' || bins[y].categoria === 'D')) {
            swap = bins[y];
            break;
          }
        }
        sugerencias.push({
          tipo:        'REUBICAR',
          sku:         item.sku,
          descripcion: item.descripcion,
          binActual:   item.binId,
          categoria:   cat,
          problema:    'Alta rotación (' + cat + ') en estante ' + item.estante + ' — lejos del despacho',
          swapSku:     swap ? swap.sku : null,
          swapDesc:    swap ? swap.descripcion : null,
          swapBin:     swap ? swap.binId : null,
          binSugerido: swap ? swap.binId : 'Zona estante 1–' + thresh
        });
      } else if (!isAltaRot && isPrime) {
        // C/D en zona prime — verificar si algún A/B está lejos
        var hayABLejos = false;
        for (var z = 0; z < bins.length; z++) {
          if ((bins[z].categoria === 'A' || bins[z].categoria === 'B') && bins[z].estante > thresh * 2) {
            hayABLejos = true;
            break;
          }
        }
        if (hayABLejos) {
          sugerencias.push({
            tipo:        'LIBERAR',
            sku:         item.sku,
            descripcion: item.descripcion,
            binActual:   item.binId,
            categoria:   cat,
            problema:    'Baja rotación (' + cat + ') ocupa estante prime ' + item.estante,
            swapSku:     null,
            swapDesc:    null,
            swapBin:     null,
            binSugerido: 'Zona estante ' + (thresh * 2 + 1) + '+'
          });
        }
      }
    }

    return { ok: true, sugerencias: sugerencias, totalBins: bins.length, maxEstante: maxEstante, thresh: thresh };
  } catch(e) {
    Logger.log('calcularSlottingOptimo: ' + e);
    return { ok: false, msg: e.toString(), sugerencias: [] };
  }
}

// ============================================================
//  REABASTECIMIENTO AUTOMÁTICO
// ============================================================

function revisarReabastecimientoAuto() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('auto_restock_habilitado') !== '1') return;
  try {
    var pedido = calcularPedidoDJI();
    var criticos = [];
    for (var p = 0; p < pedido.length; p++) {
      if ((pedido[p].criticidad === 'critico' || pedido[p].criticidad === 'urgente') && pedido[p].sugerido > 0) {
        criticos.push(pedido[p]);
      }
    }
    if (!criticos.length) { Logger.log('revisarReabastecimientoAuto: no hay items criticos/urgentes.'); return; }

    var hoy = new Date();
    var tz  = Session.getScriptTimeZone();
    var prefix = 'AUTO-' + Utilities.formatDate(hoy, tz, 'yyyyMMdd');

    // Verificar si ya existe un borrador automático hoy (para no duplicar)
    var dComp = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    for (var ci = 1; ci < dComp.length; ci++) {
      if (String(dComp[ci][0]||'').toUpperCase().indexOf(prefix) === 0) {
        Logger.log('revisarReabastecimientoAuto: ya existe ' + dComp[ci][0] + ' — omitiendo.');
        return;
      }
    }

    var cas = prefix + '-' + Utilities.formatDate(hoy, tz, 'HHmm');
    var hojaComp = getSheet(SCHEMA.SHEETS.COMPRAS);
    hojaComp.appendRow([
      cas, hoy, 'Borrador', '', '', '', '', '', '', '', '',
      'Sistema · Auto-reabastecimiento',
      'Auto-generado: ' + criticos.length + ' SKU(s) en quiebre/urgente.',
      hoy
    ]);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);

    var listaItems = [];
    for (var li = 0; li < criticos.length; li++) {
      listaItems.push({ sku: criticos[li].codigo, descripcion: criticos[li].descripcion, cantPedida: criticos[li].sugerido });
    }
    vincularItemsACAS(cas, listaItems);
    _notificarAutoRestock(cas, criticos);
    Logger.log('revisarReabastecimientoAuto: creado ' + cas + ' · ' + criticos.length + ' ítems.');
  } catch(e) {
    Logger.log('revisarReabastecimientoAuto error: ' + e);
  }
}

function _notificarAutoRestock(cas, items) {
  try {
    var filas = items.map(function(c) {
      var chipColor = c.criticidad === 'critico' ? '#c0392b' : '#b94600';
      return "<tr style='border-bottom:1px solid #eee'>" +
        "<td style='padding:7px 10px;font-weight:700;color:" + chipColor + ";font-size:12px'>" + c.codigo + "</td>" +
        "<td style='padding:7px 10px;font-size:12px'>" + c.descripcion + "</td>" +
        "<td style='padding:7px 10px;text-align:center;font-size:12px'>" + c.sugerido + "</td>" +
        "<td style='padding:7px 10px;text-align:center'><span style='font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;" +
          (c.criticidad === 'critico' ? "background:#fde8e8;color:#c0392b" : "background:#ffe5d0;color:#b94600") +
        "'>" + c.criticidad.toUpperCase() + "</span></td></tr>";
    }).join('');
    GmailApp.sendEmail(SM_CONFIG.EMAIL_SUPERVISOR,
      '[Stock Manager] Reabastecimiento automático — ' + items.length + ' SKU(s) críticos',
      '', { htmlBody:
        "<div style='font-family:sans-serif;max-width:640px'>" +
        "<div style='background:#c0392b;padding:16px 20px;border-radius:8px 8px 0 0'>" +
          "<span style='color:#fff;font-size:15px;font-weight:700'>Reabastecimiento automático generado</span></div>" +
        "<div style='background:#fff;border:1px solid #ddd;padding:18px 20px'>" +
          "<p style='font-size:13px;color:#444;margin:0 0 6px'>CAS: <strong>" + cas + "</strong></p>" +
          "<p style='font-size:13px;color:#444;margin:0 0 14px'>El sistema detectó items críticos/urgentes y generó un borrador de compra automáticamente.</p>" +
          "<table style='width:100%;border-collapse:collapse;border:1px solid #eee'>" +
            "<thead><tr style='background:#f5f5f5'><th style='padding:7px 10px;font-size:11px;text-align:left'>SKU</th>" +
            "<th style='padding:7px 10px;font-size:11px;text-align:left'>Descripción</th>" +
            "<th style='padding:7px 10px;font-size:11px'>Sugerido</th>" +
            "<th style='padding:7px 10px;font-size:11px'>Criticidad</th></tr></thead>" +
            "<tbody>" + filas + "</tbody></table>" +
        "</div>" +
        "<div style='background:#f9f9f9;border:1px solid #ddd;border-top:none;padding:12px 20px;border-radius:0 0 8px 8px'>" +
          "<p style='font-size:11px;color:#aaa;margin:0'>Generado automáticamente por Stock Manager · BIDCOMAGRO.<br>Revisá y confirmá el borrador en la sección Compras DJI.</p>" +
        "</div></div>"
      , name: SM_CONFIG.NOMBRE_REMITENTE });
  } catch(e) { Logger.log('_notificarAutoRestock: ' + e); }
}

function configurarAutoReabastecimiento(habilitar) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (habilitar) {
      props.setProperty('auto_restock_habilitado', '1');
      // Crear trigger diario si no existe
      var triggers = ScriptApp.getProjectTriggers();
      for (var t = 0; t < triggers.length; t++) {
        if (triggers[t].getHandlerFunction() === 'revisarReabastecimientoAuto') {
          return { ok: true, msg: 'Ya estaba habilitado. Trigger activo.' };
        }
      }
      ScriptApp.newTrigger('revisarReabastecimientoAuto')
        .timeBased().everyDays(1).atHour(8).create();
      return { ok: true, msg: 'Auto-reabastecimiento habilitado. Corre todos los días a las 8am.' };
    } else {
      props.setProperty('auto_restock_habilitado', '0');
      var triggers2 = ScriptApp.getProjectTriggers();
      for (var t2 = 0; t2 < triggers2.length; t2++) {
        if (triggers2[t2].getHandlerFunction() === 'revisarReabastecimientoAuto') {
          ScriptApp.deleteTrigger(triggers2[t2]);
        }
      }
      return { ok: true, msg: 'Auto-reabastecimiento deshabilitado.' };
    }
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function getConfigAutoReabastecimiento() {
  try {
    var props = PropertiesService.getScriptProperties();
    var habilitado = props.getProperty('auto_restock_habilitado') === '1';
    var triggerActivo = false;
    var triggers = ScriptApp.getProjectTriggers();
    for (var t = 0; t < triggers.length; t++) {
      if (triggers[t].getHandlerFunction() === 'revisarReabastecimientoAuto') { triggerActivo = true; break; }
    }
    return { ok: true, habilitado: habilitado, triggerActivo: triggerActivo };
  } catch(e) { return { ok: false, habilitado: false, triggerActivo: false }; }
}

// ============================================================
//  RESERVAS DE STOCK EN TRÁNSITO
// ============================================================

function crearHojaReservas() {
  try {
    var db = getDb();
    var existing = db.getSheetByName(SCHEMA.SHEETS.RESERVAS);
    if (existing) return { ok: false, msg: 'La hoja ya existe' };
    var hoja = db.insertSheet(SCHEMA.SHEETS.RESERVAS);
    hoja.appendRow(['ID','Fecha','SKU','Descripción','Cantidad','Origen','ID_Referencia','Estado','CAS_Ref','Operador','Observaciones']);
    hoja.setFrozenRows(1);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function crearReserva(sku, descripcion, cantidad, origen, idReferencia, casRef, operador) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (!hoja) return { ok: false, msg: 'Hoja RESERVAS_STOCK no encontrada. Ejecutar crearHojaReservas() primero.' };
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var id = 'RES-' + new Date().getTime();
      var R  = SCHEMA.RESERVAS_STOCK;
      var row = new Array(11);
      row[R.ID]            = id;
      row[R.FECHA]         = new Date();
      row[R.SKU]           = String(sku).trim().toUpperCase();
      row[R.DESCRIPCION]   = String(descripcion || '');
      row[R.CANTIDAD]      = parseInt(cantidad) || 1;
      row[R.ORIGEN]        = String(origen || 'VENTAS');
      row[R.ID_REFERENCIA] = String(idReferencia || '');
      row[R.ESTADO]        = 'Activa';
      row[R.CAS_REF]       = String(casRef || '');
      row[R.OPERADOR]      = String(operador || '');
      row[R.OBSERVACIONES] = '';
      hoja.appendRow(row);
      invalidateSheetValues(SCHEMA.SHEETS.RESERVAS);
      return { ok: true, id: id };
    } finally {
      if (lock.hasLock()) lock.releaseLock();
    }
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function obtenerReservasActivas() {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.RESERVAS);
    if (!datos || datos.length < 2) return [];
    var R   = SCHEMA.RESERVAS_STOCK;
    var tz  = Session.getScriptTimeZone();
    var out = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (String(f[R.ESTADO]) !== 'Activa') continue;
      out.push({
        id:          String(f[R.ID]),
        fecha:       f[R.FECHA] instanceof Date ? Utilities.formatDate(f[R.FECHA], tz, 'dd/MM/yyyy') : String(f[R.FECHA]),
        sku:         String(f[R.SKU]),
        descripcion: String(f[R.DESCRIPCION]),
        cantidad:    parseInt(f[R.CANTIDAD]) || 0,
        origen:      String(f[R.ORIGEN]),
        referencia:  String(f[R.ID_REFERENCIA]),
        casRef:      String(f[R.CAS_REF]),
        operador:    String(f[R.OPERADOR])
      });
    }
    return out;
  } catch(e) { return []; }
}

// ============================================================
//  MULTI-DEPÓSITO
// ============================================================

function getStockPorDeposito(sku) {
  try {
    var skuKey = String(sku||'').trim().toUpperCase();
    var dMov   = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M      = SCHEMA.MOVIMIENTOS_STOCK;
    var saldos = {};

    for (var i = 1; i < dMov.length; i++) {
      if (String(dMov[i][M.CODIGO]||'').trim().toUpperCase() !== skuKey) continue;
      var dep  = String(dMov[i][M.DEPOSITO]||'BA').trim() || 'BA';
      var cant = parseFloat(dMov[i][M.CANTIDAD]) || 0;
      saldos[dep] = (saldos[dep]||0) + cant;
    }

    // Eliminar depósitos con saldo 0 o negativo
    var resultado = {};
    var deps = Object.keys(saldos);
    for (var d = 0; d < deps.length; d++) {
      if (saldos[deps[d]] > 0) resultado[deps[d]] = Math.round(saldos[deps[d]]);
    }
    return { ok: true, depositos: resultado };
  } catch(e) { return { ok: false, msg: e.toString(), depositos: {} }; }
}

function listarDepositos() {
  try {
    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var deps = {};
    for (var i = 1; i < dMov.length; i++) {
      var dep = String(dMov[i][M.DEPOSITO]||'BA').trim() || 'BA';
      deps[dep] = true;
    }
    var lista = Object.keys(deps).sort();
    if (lista.length === 0) lista = ['BA'];
    return lista;
  } catch(e) { return ['BA']; }
}

function transferirEntreDepositos(sku, cantidad, depositoOrigen, depositoDestino, operador) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var cod = String(sku||'').trim().toUpperCase();
    var cant = parseInt(cantidad)||0;
    var org  = String(depositoOrigen||'BA').trim().toUpperCase();
    var dst  = String(depositoDestino||'BA').trim().toUpperCase();

    if (!cod || cant <= 0) return { ok: false, msg: 'SKU y cantidad requeridos.' };
    if (org === dst) return { ok: false, msg: 'Origen y destino deben ser diferentes.' };

    var saldos = getStockPorDeposito(cod);
    var saldoOrg = (saldos.depositos && saldos.depositos[org]) || 0;
    if (saldoOrg < cant) return { ok: false, msg: 'Stock insuficiente en ' + org + '. Disponible: ' + saldoOrg + ' u.' };

    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S    = SCHEMA.STOCK_REPUESTOS;
    var desc = '';
    for (var i = 1; i < dStr.length; i++) {
      if (String(dStr[i][S.CODIGO]||'').trim().toUpperCase() === cod) { desc = String(dStr[i][S.DESCRIPCION]||''); break; }
    }

    var saldoAct = parseInt(dStr[1] && dStr[1][S.STOCK_ACTUAL]||0);
    var ref = 'TRANSF:' + org + '→' + dst;
    _registrarMovimiento('TRANSFERENCIA_SALIDA',  cod, desc,  0, saldoAct, ref, operador||'', org);
    _registrarMovimiento('TRANSFERENCIA_ENTRADA', cod, desc,  0, saldoAct, ref, operador||'', dst);

    return { ok: true, msg: 'Transferido ' + cant + ' u. de ' + org + ' a ' + dst + '.' };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

// ── TEST — ejecutar desde el editor GAS, eliminar después de verificar ──────
function TEST_generarValeMovimiento() {
  // 1. Insertar fila de prueba en MOVIMIENTOS_STOCK
  var hojaMov = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
  hojaMov.appendRow([
    new Date(), 'EGRESO', 'TEST-SKU-001', 'Repuesto de prueba PDF',
    -1, 99, 'OT #TEST-001', Session.getActiveUser().getEmail(), 'Fila de prueba — eliminar'
  ]);
  SpreadsheetApp.flush();
  invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);

  // 2. Obtener la fila recién insertada
  var filaTest = hojaMov.getLastRow();
  Logger.log('Fila de prueba insertada: ' + filaTest);

  // 3. Generar el vale PDF
  var resultado = generarValeMovimiento(filaTest);
  Logger.log(JSON.stringify(resultado));

  if (resultado.success) {
    Logger.log('PDF generado correctamente: ' + resultado.url);
  } else {
    Logger.log('Error: ' + resultado.error);
  }

  // 4. Limpiar la fila de prueba
  hojaMov.deleteRow(filaTest);
  invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
  Logger.log('Fila de prueba eliminada.');
}
