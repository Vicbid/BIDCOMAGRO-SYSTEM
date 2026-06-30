// ── STOCK MANAGER — Ventas ─────────────────────────────────────

// ============================================================
//  VENTAS DIRECTAS
// ============================================================
function parsearCarritoReseller(datosFilas) {
  try {
    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    var stockMap = {};
    for (var s = 1; s < dStr.length; s++) {
      var cod = String(dStr[s][0]).trim().toUpperCase();
      stockMap[cod] = parseInt(dStr[s][2]) || 0;
    }

    var items = [];
    if (!datosFilas || datosFilas.length === 0) throw new Error("El archivo está vacÍ­o o no tiene el formato correcto.");

    for (var i = 0; i < datosFilas.length; i++) {
      var f = datosFilas[i];
      var cod = String(f.codigo || "").trim().toUpperCase();
      if (!cod || cod === "UNDEFINED") continue; // Ignorar filas vacÍ­as

      var stock = stockMap[cod] !== undefined ? stockMap[cod] : -1;
      items.push({
        codigo: cod,
        descripcion: String(f.descripcion || "Sin descripción"),
        cantidad: parseInt(f.cantidad) || 0,
        precioUSD: parseFloat(f.precioUSD) || 0,
        totalUSD: (parseFloat(f.precioUSD) || 0) * (parseInt(f.cantidad) || 0),
        stockDisponible: stock,
        sinStock: stock >= 0 && stock < (parseInt(f.cantidad) || 0)
      });
    }

    if (items.length === 0) throw new Error("No se detectaron códigos de repuestos válidos en el archivo.");
    return { ok: true, items: items };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  }
}

function confirmarVentaDirecta(reseller, items, nOrdenEntrega, operador) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    var ss = getSS();
    var hojaVentas = getSheet(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var hojaSolic  = getSheet(SCHEMA.SHEETS.SOLICITUDES);
    var hoy = new Date();

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var cod = String(item.codigo).trim().toUpperCase();
      var cant = parseInt(item.cantidad) || 0;

      // 1. REGISTRO CONTABLE (Historial de ventas)
      hojaVentas.appendRow([
        hoy,
        nOrdenEntrega,
        "",
        reseller,
        cod,
        item.descripcion,
        cant,
        item.precioUSD,
        "Venta directa"
      ]);

      // 2. REGISTRO OPERATIVO (Para que aparezca en la pestaña Despachos)
      var idSolicitud = "VD-" + nOrdenEntrega + "-" + i;
      hojaSolic.appendRow([
        idSolicitud,
        hoy,
        nOrdenEntrega,
        reseller,
        cod,
        item.descripcion,
        cant,
        0,
        "Pendiente",
        "NORMAL",
        "", "", ""
      ]);

      // IMPORTANTE: El stock NO se toca en esta función.
    }

    return { ok: true, nOrden: nOrdenEntrega };

  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function cargarVentas(limite) {
  try {
    var d   = getSheetValues(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var max = limite || 200;
    var out = [], porOrden = {};
    for (var i = d.length-1; i >= 1 && out.length < max; i--) {
      var f = d[i];
      var nOrd = String(f[1]||"");
      if (!porOrden[nOrd]) {
        porOrden[nOrd] = {
          fecha: _fmtFecha(f[0]), nOrden: nOrd,
          nFactura: String(f[2]||""), reseller: String(f[3]||""),
          items: [], totalUSD: 0
        };
        out.push(porOrden[nOrd]);
      }
      var precio = parseFloat(f[7])||0, cant = parseInt(f[6])||0;
      porOrden[nOrd].items.push({
        codigo: String(f[4]), descripcion: String(f[5]),
        cantidad: cant, precioUSD: precio, totalUSD: precio*cant
      });
      porOrden[nOrd].totalUSD += precio*cant;
    }
    return out;
  } catch(e) { return []; }
}

function completarFactura(nOrdenEntrega, nFactura) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var d    = hoja.getDataRange().getValues();
    var nOrd = String(nOrdenEntrega).trim();
    var changed = false;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1]||"").trim() === nOrd) {
        d[i][2] = nFactura;
        changed  = true;
      }
    }
    if (changed) hoja.getDataRange().setValues(d);
    return { ok: changed };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function generarNumeroOrden() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var maxNum = 0;
    for (var i = 1; i < d.length; i++) {
      var n = String(d[i][1]||"").replace(/[^0-9]/g,"");
      if (n) maxNum = Math.max(maxNum, parseInt(n)||0);
    }
    var num = String(maxNum + 1);
    while (num.length < 4) num = "0" + num;
    return "OE-" + num;
  } catch(e) { return "OE-0001"; }
}

function obtenerResellers() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var lista = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0]) lista.push(String(d[i][0]).trim());
    }
    return lista.sort();
  } catch(e) { return []; }
}
