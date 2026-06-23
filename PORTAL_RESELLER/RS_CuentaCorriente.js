// ============================================================
// @version 1.0
//  PORTAL RESELLER — Módulo Cuenta Corriente (Solo Lectura)
//  Fuente: CC.VtaCte (spreadsheet externo, ID en servidor)
//  Seguridad: nombre de sesión revalidado contra hoja Resellers
//             antes de cualquier acceso a datos de CC.
//  Caché: UserCache 5 min — datos históricos fijos, sin mutaciones.
// ============================================================

var _CC_SS_ID        = '1KVaNIQVJPPTA-bAfDjReQanC5a24KOuHRNQQoq2R4sg';
var _CC_HOJA         = 'CC.VtaCte';
var _CC_HOJA_COB     = 'Cobranzas';
var _CC_CACHE_TTL    = 300; // 5 minutos

// Índices de columnas — hoja Cobranzas
var _CC_COB = {
  FECHA_PAGO:  0,   // A
  ID_VENTA:    1,   // B
  MEDIO_PAGO:  4,   // E
  MONEDA:      5,   // F
  CUIT:        6,   // G
  EMISOR:      7,   // H
  BANCO:       8,   // I
  REFERENCIA:  9,   // J
  FECHA_COBRO: 11,  // L
  MONTO_ARS:   12,  // M
  TC:          13,  // N
  MONTO_USD:   14,  // O
  COMPROBANTE: 18   // S
};

// Índices de columnas — espejo exacto del esquema de CC.VtaCte
var _CC = {
  FECHA:        0,
  ID_VENTA:     1,
  RESELLER:     2,
  RAZON_SOCIAL: 3,
  VENTAS_USD:   4,
  COBR_USD:     5,
  SALDO_USD:    6,
  VENTAS_ARS:   7,
  COBR_ARS:     8,
  SALDO_ARS:    9,
  ID_ENTREGA:   10,
  OPERACION:    11,
  CONDICION:    12,
  METODO_PAGO:  13,
  SKU:          14,
  RTV_ZONAL:    15,
  COMENTARIO:   16
};

// ── Seguridad: revalida que `nombre` sea un reseller legítimo ──
// Devuelve true solo si existe una fila en Resellers con ese nombre exacto.
// Impide que un cliente manipule el param para ver datos de otro reseller.
function _CC_resellerValido(nombre) {
  try {
    var d       = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var nombreB = String(nombre || '').trim().toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === nombreB) return true;
    }
    return false;
  } catch(e) {
    Logger.log('_CC_resellerValido: ' + e);
    return false;
  }
}

// ── Lee y cachea los movimientos del reseller desde CC.VtaCte ──
// El cache es por nombre de reseller. TTL corto por si hay ajustes manuales.
function _CC_leerDatos(nombre) {
  var cache    = CacheService.getUserCache();
  var cacheKey = 'cc_v1_' + nombre.toLowerCase().replace(/[^a-z0-9]/g, '_');

  var cached = cache.get(cacheKey);
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }

  var hoja = SpreadsheetApp.openById(_CC_SS_ID).getSheetByName(_CC_HOJA);
  if (!hoja) { Logger.log('_CC_leerDatos: hoja no encontrada'); return []; }

  var datos    = hoja.getDataRange().getValues();
  var nombreB  = nombre.trim().toLowerCase();
  var tz       = Session.getScriptTimeZone();
  var filas    = [];

  for (var i = 1; i < datos.length; i++) {
    var r = datos[i];

    // Saltar filas vacías o de otro reseller
    if (!r[_CC.ID_VENTA]) continue;
    if (String(r[_CC.RESELLER] || '').trim().toLowerCase() !== nombreB) continue;

    // Normalizar fecha — puede llegar como Date o string DD/MM/YYYY
    var fechaRaw = r[_CC.FECHA];
    var fechaStr = '';
    if (fechaRaw instanceof Date) {
      fechaStr = Utilities.formatDate(fechaRaw, tz, 'dd/MM/yyyy');
    } else {
      fechaStr = String(fechaRaw || '').trim();
    }

    filas.push({
      fecha:       fechaStr,
      idVenta:     String(r[_CC.ID_VENTA]     || ''),
      razonSocial: String(r[_CC.RAZON_SOCIAL] || ''),
      ventasUSD:   Number(r[_CC.VENTAS_USD])  || 0,
      cobrUSD:     Number(r[_CC.COBR_USD])    || 0,
      saldoUSD:    Number(r[_CC.SALDO_USD])   || 0,
      ventasARS:   Number(r[_CC.VENTAS_ARS])  || 0,
      cobrARS:     Number(r[_CC.COBR_ARS])    || 0,
      saldoARS:    Number(r[_CC.SALDO_ARS])   || 0,
      idEntrega:   String(r[_CC.ID_ENTREGA]   || ''),
      operacion:   String(r[_CC.OPERACION]    || ''),
      condicion:   String(r[_CC.CONDICION]    || ''),
      metodoPago:  String(r[_CC.METODO_PAGO]  || ''),
      sku:         String(r[_CC.SKU]          || ''),
      rtvZonal:    String(r[_CC.RTV_ZONAL]    || ''),
      comentario:  String(r[_CC.COMENTARIO]   || '')
    });
  }

  try {
    var payload = JSON.stringify(filas);
    if (payload.length < 90000) cache.put(cacheKey, payload, _CC_CACHE_TTL);
  } catch(e) { Logger.log('_CC_leerDatos cache.put: ' + e); }

  return filas;
}

// ── Lee cobranzas de la hoja "Cobranzas" para un set de ID Venta ──
// idSet: objeto { idVenta: true } — solo IDs que pertenecen al reseller.
// Retorna mapa { idVenta: [ cobranza, ... ] } para lookup O(1) en el detalle.
// Cachea junto con los movimientos usando la misma clave de reseller.
function _CC_leerCobranzas(idSet) {
  try {
    var hoja  = SpreadsheetApp.openById(_CC_SS_ID).getSheetByName(_CC_HOJA_COB);
    if (!hoja) return {};
    var datos = hoja.getDataRange().getValues();
    var tz    = Session.getScriptTimeZone();
    var mapa  = {};

    for (var i = 1; i < datos.length; i++) {
      var r   = datos[i];
      var idV = String(r[_CC_COB.ID_VENTA] || '').trim();
      if (!idV || !idSet[idV]) continue; // seguridad: solo IDs del reseller

      // Normalizar fechas
      var fPago  = r[_CC_COB.FECHA_PAGO]  instanceof Date
        ? Utilities.formatDate(r[_CC_COB.FECHA_PAGO],  tz, 'dd/MM/yyyy') : String(r[_CC_COB.FECHA_PAGO]  || '');
      var fCobro = r[_CC_COB.FECHA_COBRO] instanceof Date
        ? Utilities.formatDate(r[_CC_COB.FECHA_COBRO], tz, 'dd/MM/yyyy') : String(r[_CC_COB.FECHA_COBRO] || '');

      var cob = {
        fechaPago:   fPago,
        medioPago:   String(r[_CC_COB.MEDIO_PAGO]  || ''),
        moneda:      String(r[_CC_COB.MONEDA]       || ''),
        cuit:        String(r[_CC_COB.CUIT]         || ''),
        emisor:      String(r[_CC_COB.EMISOR]       || ''),
        banco:       String(r[_CC_COB.BANCO]        || ''),
        referencia:  String(r[_CC_COB.REFERENCIA]   || ''),
        fechaCobro:  fCobro,
        montoARS:    Number(r[_CC_COB.MONTO_ARS])   || 0,
        tc:          Number(r[_CC_COB.TC])           || 0,
        montoUSD:    Number(r[_CC_COB.MONTO_USD])    || 0,
        comprobante: String(r[_CC_COB.COMPROBANTE]  || '')
      };

      if (!mapa[idV]) mapa[idV] = [];
      mapa[idV].push(cob);
    }
    return mapa;
  } catch(e) {
    Logger.log('_CC_leerCobranzas: ' + e);
    return {};
  }
}

// ── Aplica filtros de fecha y tipo de operación (server-side) ──
function _CC_aplicarFiltros(filas, filtros) {
  if (!filtros) return filas;

  var desde = _CC_parseFecha(filtros.desde || '');
  var hasta  = _CC_parseFecha(filtros.hasta  || '');
  var opFil  = String(filtros.operacion || '').trim().toLowerCase();

  var resultado = [];
  for (var i = 0; i < filas.length; i++) {
    var f  = filas[i];
    var ok = true;

    if (desde) {
      var fd = _CC_parseFecha(f.fecha);
      if (!fd || fd < desde) ok = false;
    }
    if (ok && hasta) {
      var fh = _CC_parseFecha(f.fecha);
      if (!fh || fh > hasta) ok = false;
    }
    if (ok && opFil && f.operacion.toLowerCase().indexOf(opFil) === -1) ok = false;

    if (ok) resultado.push(f);
  }
  return resultado;
}

// ── KPIs calculados sobre el conjunto filtrado ─────────────────
// Opción A: saldo = suma algebraica (ventas - cobranzas)
function _CC_calcularKpis(filas) {
  var totalVentasUSD = 0, totalCobrUSD = 0;
  var totalVentasARS = 0, totalCobrARS = 0;

  for (var i = 0; i < filas.length; i++) {
    totalVentasUSD += filas[i].ventasUSD;
    totalCobrUSD   += filas[i].cobrUSD;
    totalVentasARS += filas[i].ventasARS;
    totalCobrARS   += filas[i].cobrARS;
  }

  return {
    saldoUSD:      Math.round((totalVentasUSD - totalCobrUSD) * 100) / 100,
    saldoARS:      Math.round((totalVentasARS - totalCobrARS) * 100) / 100,
    totalVentasUSD:Math.round(totalVentasUSD * 100) / 100,
    totalCobrUSD:  Math.round(totalCobrUSD   * 100) / 100,
    totalVentasARS:Math.round(totalVentasARS * 100) / 100,
    totalCobrARS:  totalCobrARS,
    cantidad:      filas.length
  };
}

// ── Punto de entrada público ───────────────────────────────────
// Llamado desde el cliente con google.script.run
// params.nombre  → nombre de empresa (revalidado en servidor)
// params.filtros → { desde, hasta, operacion } (opcionales)
function CC_obtenerMovimientos(params) {
  try {
    var nombre  = String((params && params.nombre) || '').trim();
    var filtros = (params && params.filtros) ? params.filtros : null;

    if (!nombre) return { ok: false, error: 'Sesión no identificada.' };

    // Revalidar que el nombre pertenece a un reseller legítimo
    if (!_CC_resellerValido(nombre)) {
      Logger.log('CC_obtenerMovimientos: acceso denegado para "' + nombre + '"');
      return { ok: false, error: 'Acceso no autorizado.' };
    }

    var filas     = _CC_leerDatos(nombre);
    var filtradas = _CC_aplicarFiltros(filas, filtros);
    var kpis      = _CC_calcularKpis(filtradas);

    // Construir set de ID Venta propios del reseller (todos, no solo los filtrados)
    // para que el detalle pueda ver cobranzas aunque no estén en el período filtrado
    var idSet = {};
    for (var iv = 0; iv < filas.length; iv++) {
      if (filas[iv].idVenta) idSet[filas[iv].idVenta] = true;
    }
    var cobranzas = _CC_leerCobranzas(idSet);

    // Más recientes primero en el listado
    var ordenadas = filtradas.slice().reverse();

    return { ok: true, movimientos: ordenadas, kpis: kpis, cobranzas: cobranzas };
  } catch(e) {
    Logger.log('CC_obtenerMovimientos: ' + e);
    return { ok: false, error: 'Error interno. Intentá más tarde.' };
  }
}

// ── Exportar como XLSX ────────────────────────────────────────
function CC_exportarXLS(params) {
  try {
    var nombre = String((params && params.nombre) || '').trim();
    if (!nombre || !_CC_resellerValido(nombre)) return { ok: false, error: 'No autorizado.' };

    var filas = _CC_aplicarFiltros(_CC_leerDatos(nombre), params.filtros || null);

    var ssTemp = SpreadsheetApp.create('CC_export_' + new Date().getTime());
    var hoja   = ssTemp.getActiveSheet();
    hoja.setName('Cuenta Corriente');
    hoja.appendRow([
      'Fecha','ID Venta','Operación','Condición','Método Pago',
      'Ventas USD','Cobranzas USD','Saldo USD',
      'Ventas ARS','Cobranzas ARS','Saldo ARS',
      'ID Entrega','SKU','RTV/Zonal','Comentario'
    ]);
    for (var i = 0; i < filas.length; i++) {
      var f = filas[i];
      hoja.appendRow([
        f.fecha, f.idVenta, f.operacion, f.condicion, f.metodoPago,
        f.ventasUSD, f.cobrUSD, f.saldoUSD,
        f.ventasARS, f.cobrARS, f.saldoARS,
        f.idEntrega, f.sku, f.rtvZonal, f.comentario
      ]);
    }
    // Autoformat encabezado
    hoja.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#ffffff');

    var file   = DriveApp.getFileById(ssTemp.getId());
    var blob   = file.getAs(MimeType.MICROSOFT_EXCEL);
    var base64 = Utilities.base64Encode(blob.getBytes());
    file.setTrashed(true);

    return { ok: true, base64: base64, nombre: 'CuentaCorriente_' + nombre + '.xlsx' };
  } catch(e) {
    Logger.log('CC_exportarXLS: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Exportar como PDF ─────────────────────────────────────────
function CC_exportarPDF(params) {
  try {
    var nombre = String((params && params.nombre) || '').trim();
    if (!nombre || !_CC_resellerValido(nombre)) return { ok: false, error: 'No autorizado.' };

    var filas = _CC_aplicarFiltros(_CC_leerDatos(nombre), params.filtros || null);
    var tz    = Session.getScriptTimeZone();
    var ahora = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');

    var docName = 'CC_' + nombre + '_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
    var doc     = DocumentApp.create(docName);
    var body    = doc.getBody();
    body.clear();

    body.appendParagraph('BIDCOMAGRO — Cuenta Corriente')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(nombre + '  ·  Generado el ' + ahora)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setFontSize(9);
    body.appendHorizontalRule();
    body.appendParagraph('');

    var headers = ['Fecha','ID Venta','Operación','Condición','Ventas USD','Cobr. USD','Saldo USD','Ventas ARS','Cobr. ARS','Saldo ARS'];
    var tableData = [headers];
    for (var i = 0; i < filas.length; i++) {
      var f = filas[i];
      tableData.push([
        f.fecha, f.idVenta, f.operacion, f.condicion,
        f.ventasUSD ? 'U$S '+f.ventasUSD : '—',
        f.cobrUSD   ? 'U$S '+f.cobrUSD   : '—',
        'U$S '+f.saldoUSD,
        f.ventasARS ? '$'+f.ventasARS : '—',
        f.cobrARS   ? '$'+f.cobrARS   : '—',
        '$'+f.saldoARS
      ]);
    }
    var table = body.appendTable(tableData);
    var hdr   = table.getRow(0);
    for (var c = 0; c < hdr.getNumCells(); c++) {
      hdr.getCell(c).setFontSize(8).setBold(true).setBackgroundColor('#00a3e0');
    }
    for (var r = 1; r < table.getNumRows(); r++) {
      for (var c2 = 0; c2 < table.getRow(r).getNumCells(); c2++) {
        table.getRow(r).getCell(c2).setFontSize(8);
      }
    }

    body.appendParagraph('');
    body.appendParagraph('Total de registros: ' + filas.length)
        .setFontSize(9).setItalic(true);

    doc.saveAndClose();
    var pdfBlob = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
    var base64  = Utilities.base64Encode(pdfBlob.getBytes());
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return { ok: true, base64: base64, nombre: docName + '.pdf' };
  } catch(e) {
    Logger.log('CC_exportarPDF: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Utilidad: parsea string DD/MM/YYYY a Date ──────────────────
function _CC_parseFecha(str) {
  if (!str) return null;
  var p = String(str).trim().split('/');
  if (p.length !== 3) return null;
  var d = parseInt(p[0], 10), m = parseInt(p[1], 10) - 1, y = parseInt(p[2], 10);
  if (isNaN(d) || isNaN(m) || isNaN(y)) return null;
  return new Date(y, m, d);
}
