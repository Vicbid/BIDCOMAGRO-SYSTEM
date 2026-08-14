// ── STOCK MANAGER — WMS ─────────────────────────────────────
// @version 1.2

// ── WMS — GESTIÓN DE UBICACIONES EN CARMEN ──────────────────
function cargarUbicacionesItem(sku) {
  try {
    var codKey   = String(sku).trim().toUpperCase();
    var hojaUbic = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: true, ubicaciones: [] };
    var d = hojaUbic.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() !== codKey) continue;
      out.push({
        ubicacion: String(d[i][1] || '').trim(),
        cantidad:  parseFloat(d[i][2]) || 0
      });
    }
    return { ok: true, ubicaciones: out };
  } catch(e) { return { ok: false, error: e.message }; }
}

function guardarUbicacionInicial(sku, ubicacion, cantidadInicial) {
  try {
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
    var codKey  = String(sku).trim().toUpperCase();
    var ubicKey = String(ubicacion).trim().toUpperCase();
    var cantIni = Math.max(0, parseInt(cantidadInicial) || 0);
    var d = hojaUbic.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() === codKey &&
          String(d[i][1] || '').trim().toUpperCase() === ubicKey) {
        hojaUbic.getRange(i + 1, 3).setValue(cantIni);
        return { ok: true };
      }
    }
    // Escribir fila nueva con col B forzada a texto para evitar auto-conversión a fecha (ej: "1-1" → 1 ene)
    var lr = hojaUbic.getLastRow() + 1;
    hojaUbic.getRange(lr, 1).setValue(codKey);
    hojaUbic.getRange(lr, 2).setNumberFormat('@').setValue(ubicKey);
    hojaUbic.getRange(lr, 3).setValue(cantIni);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}

// Retorna todos los ítems registrados en una ubicación específica
function cargarUbicacionesSector(ubicacion) {
  try {
    var ubicKey  = String(ubicacion).trim().toUpperCase();
    var hojaUbic = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: true, items: [] };
    var d   = hojaUbic.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim().toUpperCase() !== ubicKey) continue;
      out.push({
        sku:      String(d[i][0] || '').trim(),
        cantidad: parseFloat(d[i][2]) || 0
      });
    }
    return { ok: true, items: out };
  } catch(e) { return { ok: false, error: e.message }; }
}

function eliminarUbicacion(sku, ubicacion) {
  try {
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe' };
    var codKey  = String(sku).trim().toUpperCase();
    var ubicKey = String(ubicacion).trim().toUpperCase();
    var d = hojaUbic.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() === codKey &&
          String(d[i][1] || '').trim().toUpperCase() === ubicKey) {
        hojaUbic.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Ubicación no encontrada' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// Repara la columna B (ubicacion) de Carmen UBICACIONES: fuerza formato texto en todas las filas
// para corregir valores que Sheets auto-convirtió a fechas (ej: "1-1" → 1-ene-2001).
function SM_repararFormatoUbicaciones() {
  try {
    var hoja = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hoja) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { ok: true, reparadas: 0 };
    var rangeB = hoja.getRange(2, 2, lastRow - 1, 1);
    var vals   = rangeB.getValues();
    // Forzar formato texto en toda la columna B de datos
    rangeB.setNumberFormat('@');
    // Detectar y corregir celdas que se guardaron como Date (número serial)
    var corregidas = 0;
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i][0];
      if (v instanceof Date) {
        // Era una fecha → no podemos recuperar el texto original, marcar para revisión manual
        hoja.getRange(i + 2, 2).setValue('⚠ REVISAR');
        corregidas++;
      }
    }
    Logger.log('SM_repararFormatoUbicaciones: ' + corregidas + ' celdas con fecha detectadas');
    return { ok: true, reparadas: corregidas };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// Mueve `cant` unidades de un SKU de una ubicación a otra en UBICACIONES de Carmen.
// Lockeada (igual que SM_unificarUbicaciones): lee toda la hoja una vez y después borra/escribe
// por número de fila calculado de esa lectura — sin lock, dos llamadas casi simultáneas (ahora
// puede llegar tanto desde "Mover" del WMS como desde "Transferencia Interna", Sm_Stock.js)
// podían calcular filas antes de que la otra terminara de escribir y pisar/borrar la fila
// equivocada, perdiendo en silencio la ubicación de un SKU que no tenía nada que ver con esa
// operación puntual.
function SM_moverStock(sku, origen, destino, cant) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
    var codKey = String(sku).trim().toUpperCase();
    var orgKey = String(origen).trim().toUpperCase();
    var dstKey = String(destino).trim().toUpperCase();
    cant       = Math.max(1, parseInt(cant) || 0);
    if (orgKey === dstKey) return { ok: false, error: 'Origen y destino son el mismo bin' };
    var d = hojaUbic.getDataRange().getValues();
    var orgRow = -1, dstRow = -1, orgCant = 0, dstCant = 0;
    for (var i = 1; i < d.length; i++) {
      var rowSku  = String(d[i][0] || '').trim().toUpperCase();
      var rowUbic = String(d[i][1] || '').trim().toUpperCase();
      if (rowSku !== codKey) continue;
      if (rowUbic === orgKey) { orgRow = i + 1; orgCant = parseFloat(d[i][2]) || 0; }
      if (rowUbic === dstKey) { dstRow = i + 1; dstCant = parseFloat(d[i][2]) || 0; }
    }
    if (orgRow < 0) return { ok: false, error: 'El SKU no está en el bin de origen' };
    if (cant > orgCant) return { ok: false, error: 'Cantidad a mover (' + cant + ') supera el stock en origen (' + orgCant + ')' };
    var newOrgCant = orgCant - cant;
    if (newOrgCant === 0) {
      hojaUbic.deleteRow(orgRow);
      // Si dstRow estaba después del orgRow, su índice bajó uno
      if (dstRow > orgRow) dstRow--;
    } else {
      hojaUbic.getRange(orgRow, 3).setValue(newOrgCant);
    }
    if (dstRow > 0) {
      hojaUbic.getRange(dstRow, 3).setValue(dstCant + cant);
    } else {
      var lrMv = hojaUbic.getLastRow() + 1;
      hojaUbic.getRange(lrMv, 1).setValue(codKey);
      hojaUbic.getRange(lrMv, 2).setNumberFormat('@').setValue(dstKey);
      hojaUbic.getRange(lrMv, 3).setValue(dstCant + cant);
    }
    Logger.log('SM_moverStock: ' + codKey + ' ' + cant + 'u. ' + orgKey + ' → ' + dstKey);
    return { ok: true, msg: cant + ' u. movidas de ' + orgKey + ' a ' + dstKey };
  } catch(e) {
    Logger.log('SM_moverStock: ' + e);
    return { ok: false, error: e.message };
  } finally {
    try { lock.releaseLock(); } catch(eL) {}
  }
}

// Reescribe de una la distribución de un SKU entre ubicaciones de Carmen — pantalla "Unificar
// ubicaciones" (SM_Index.html, tab WMS/Ubicaciones): el operador ve los códigos repartidos en
// más de un bin y arma a mano la distribución final (puede vaciar bines, dejar solo uno, agregar
// uno nuevo, etc.). distribucionFinal = [{ubicacion, cantidad}]. Valida que la suma coincida con
// el total actualmente registrado para ese SKU antes de tocar nada (protege contra datos
// desactualizados en el cliente — ej. otra persona movió stock mientras el operador editaba).
function SM_unificarUbicaciones(sku, distribucionFinal) {
  try {
    var codKey = String(sku).trim().toUpperCase();
    var destino = {};
    var totalDestino = 0;
    for (var i = 0; i < (distribucionFinal || []).length; i++) {
      var u = String(distribucionFinal[i].ubicacion || '').trim().toUpperCase();
      var c = Math.max(0, parseInt(distribucionFinal[i].cantidad, 10) || 0);
      if (!u || c <= 0) continue;
      destino[u] = (destino[u] || 0) + c;
      totalDestino += c;
    }
    if (!Object.keys(destino).length) return { ok: false, error: 'La distribución final no puede quedar vacía.' };

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss   = _getCarmenSS();
      var hoja = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
      if (!hoja) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
      var d = hoja.getDataRange().getValues();

      var filas = []; // { row, ubicacion, cantidad } — filas actuales de este SKU
      var totalActual = 0;
      for (var r = 1; r < d.length; r++) {
        if (String(d[r][0] || '').trim().toUpperCase() !== codKey) continue;
        var ub = String(d[r][1] || '').trim().toUpperCase();
        var ca = parseFloat(d[r][2]) || 0;
        filas.push({ row: r + 1, ubicacion: ub, cantidad: ca });
        totalActual += ca;
      }
      if (!filas.length) return { ok: false, error: 'Este SKU no tiene ubicaciones registradas.' };
      if (totalDestino !== totalActual) {
        return { ok: false, error: 'La suma de la nueva distribución (' + totalDestino + ') no coincide con el stock actual en ubicaciones (' + totalActual + '). Actualizá la página e intentá de nuevo.' };
      }

      // De abajo hacia arriba para poder borrar filas sin romper los índices de las que faltan.
      var escritas = {};
      for (var f = filas.length - 1; f >= 0; f--) {
        var fila  = filas[f];
        var nueva = destino.hasOwnProperty(fila.ubicacion) ? destino[fila.ubicacion] : 0;
        if (nueva > 0) {
          hoja.getRange(fila.row, 3).setValue(nueva);
          escritas[fila.ubicacion] = true;
        } else {
          hoja.deleteRow(fila.row);
        }
      }
      // Ubicaciones nuevas que no existían antes en este SKU.
      var ubics = Object.keys(destino);
      for (var k = 0; k < ubics.length; k++) {
        var uK = ubics[k];
        if (escritas[uK]) continue;
        var lr = hoja.getLastRow() + 1;
        hoja.getRange(lr, 1).setValue(codKey);
        hoja.getRange(lr, 2).setNumberFormat('@').setValue(uK);
        hoja.getRange(lr, 3).setValue(destino[uK]);
      }

      Logger.log('SM_unificarUbicaciones: ' + codKey + ' → ' + JSON.stringify(destino));
      return { ok: true };
    } finally {
      try { lock.releaseLock(); } catch(eL) {}
    }
  } catch(e) {
    Logger.log('SM_unificarUbicaciones: ' + e);
    return { ok: false, error: e.message };
  }
}

// Mueve todos los SKUs de una caja completa de un bin a otro.
// origenBin4 = BIN_ID exacto (ej: "1-A-2-3"), destinoBin4 = BIN_ID destino (ej: "2-B-1-1")
function SM_moverCaja(origenBin4, destinoBin4) {
  try {
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
    var orgKey = String(origenBin4).trim().toUpperCase();
    var dstKey = String(destinoBin4).trim().toUpperCase();
    if (orgKey === dstKey) return { ok: false, error: 'Origen y destino son iguales' };
    var d = hojaUbic.getDataRange().getValues();
    // Recolectar filas de origen + mapear destinos existentes
    var orgRows   = [];
    var dstLookup = {};
    for (var i = 1; i < d.length; i++) {
      var rowSku  = String(d[i][0] || '').trim().toUpperCase();
      var rowUbic = String(d[i][1] || '').trim().toUpperCase();
      var rowCant = parseFloat(d[i][2]) || 0;
      if (rowUbic === orgKey)  orgRows.push({ rowIdx: i + 1, sku: rowSku, cant: rowCant });
      if (rowUbic === dstKey)  dstLookup[rowSku] = { rowIdx: i + 1, cant: rowCant };
    }
    if (!orgRows.length) return { ok: false, error: 'No se encontraron ítems en ' + orgKey };
    // Actualizar o crear filas de destino
    for (var k = 0; k < orgRows.length; k++) {
      var og = orgRows[k];
      if (dstLookup[og.sku]) {
        hojaUbic.getRange(dstLookup[og.sku].rowIdx, 3).setValue(dstLookup[og.sku].cant + og.cant);
        dstLookup[og.sku].cant += og.cant;
      } else {
        var lr = hojaUbic.getLastRow() + 1;
        hojaUbic.getRange(lr, 1).setValue(og.sku);
        hojaUbic.getRange(lr, 2).setNumberFormat('@').setValue(dstKey);
        hojaUbic.getRange(lr, 3).setValue(og.cant);
        dstLookup[og.sku] = { rowIdx: lr, cant: og.cant };
      }
    }
    // Eliminar filas de origen en orden inverso para preservar índices
    orgRows.sort(function(a, b) { return b.rowIdx - a.rowIdx; });
    for (var r = 0; r < orgRows.length; r++) hojaUbic.deleteRow(orgRows[r].rowIdx);
    Logger.log('SM_moverCaja: ' + orgKey + ' → ' + dstKey + ' (' + orgRows.length + ' SKUs)');
    return { ok: true, movidas: orgRows.length };
  } catch(e) {
    Logger.log('SM_moverCaja: ' + e);
    return { ok: false, error: e.message };
  }
}

// Escribe el conteo completo de una ubicación: sobrescribe la cantidad de cada SKU provisto
// (lo borra si lo pusieron en 0). NO toca ningún otro SKU del bin que no venga en `items` —
// antes sí lo hacía ("lo que no vino en la lista, se borra"), y eso perdía en silencio
// cualquier ubicación que otro operador hubiera mapeado a ese mismo bin (Recibir compra, Mover
// stock, etc.) mientras esta pantalla de Contar seguía abierta con una lista vieja: al guardar,
// ese SKU nuevo no estaba en `items` y se eliminaba aunque el operador nunca lo vio ni lo tocó.
// Reporte del operador: "hay ubicaciones que estoy seguro que mapeé pero se perdieron" — ver
// [[carmen_fuente_stock]]. Las bajas explícitas de un ítem puntual ya las maneja
// `eliminarUbicacion` en tiempo real (botón de tacho de cada fila), así que acá alcanza con
// nunca borrar lo que el operador no tiene cargado en pantalla. `items` = [{sku, cantidad}]
function guardarConteoUbicacion(ubicacion, items) {
  try {
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
    var ubicKey = String(ubicacion).trim().toUpperCase();
    // Construir mapa de cantidades a aplicar (solo los SKUs que el operador tiene en pantalla)
    var nuevoMapa = {};
    for (var k = 0; k < items.length; k++) {
      var s = String(items[k].sku || '').trim().toUpperCase();
      var c = Math.max(0, parseInt(items[k].cantidad) || 0);
      if (s) nuevoMapa[s] = c;
    }
    var d = hojaUbic.getDataRange().getValues();
    // Procesar filas existentes (en reversa para poder eliminar) — solo las que el operador
    // tiene en `items`; cualquier otra fila del bin se deja intacta.
    var procesadas = {};
    for (var i = d.length - 1; i >= 1; i--) {
      var rowSku  = String(d[i][0] || '').trim().toUpperCase();
      var rowUbic = String(d[i][1] || '').trim().toUpperCase();
      if (rowUbic !== ubicKey) continue;
      if (!nuevoMapa.hasOwnProperty(rowSku)) continue;
      if (nuevoMapa[rowSku] === 0) {
        hojaUbic.deleteRow(i + 1);
      } else {
        hojaUbic.getRange(i + 1, 3).setValue(nuevoMapa[rowSku]);
      }
      procesadas[rowSku] = true;
    }
    // Agregar SKUs nuevos (no existían antes)
    var keys = Object.keys(nuevoMapa);
    for (var m = 0; m < keys.length; m++) {
      if (!procesadas[keys[m]] && nuevoMapa[keys[m]] > 0) {
        var lr2 = hojaUbic.getLastRow() + 1;
        hojaUbic.getRange(lr2, 1).setValue(keys[m]);
        hojaUbic.getRange(lr2, 2).setNumberFormat('@').setValue(ubicKey);
        hojaUbic.getRange(lr2, 3).setValue(nuevoMapa[keys[m]]);
      }
    }
    return { ok: true };
  } catch(e) {
    Logger.log('guardarConteoUbicacion: ' + e);
    return { ok: false, error: e.message };
  }
}

// Devuelve ítems con stock en Carmen que no tienen ninguna ubicación en UBICACIONES.
// Ordenados desc por stock. Limita a 200 ítems para no saturar.
function cargarSinMapear() {
  try {
    var ss       = _getCarmenSS();
    var hojaStk  = ss.getSheetByName('STOCK');
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaStk) return { ok: false, error: 'Hoja STOCK no encontrada en Carmen' };

    // Set de SKUs ya mapeados
    var mapeados = {};
    if (hojaUbic) {
      var dU = hojaUbic.getDataRange().getValues();
      for (var u = 1; u < dU.length; u++) {
        var sk = String(dU[u][0] || '').trim().toUpperCase();
        if (sk) mapeados[sk] = true;
      }
    }

    var dS  = hojaStk.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < dS.length; i++) {
      var cod   = String(dS[i][0] || '').trim().toUpperCase();
      var stock = parseFloat(dS[i][2]) || 0;
      if (!cod || stock <= 0 || mapeados[cod]) continue;
      out.push({ sku: cod, descripcion: String(dS[i][1] || ''), stock: stock });
    }
    out.sort(function(a, b) { return b.stock - a.stock; });
    return { ok: true, items: out.slice(0, 200), total: out.length };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ============================================================
//  LAYOUT DE ALMACÉN — mapa de estantes, paños y alturas
//  BIN_ID = ESTANTE-PAÑO-ALTURA (ej: "1-A-2")
// ============================================================

function SM_cargarLayout() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    if (!hoja) return { ok: true, estantes: [] };
    var d  = hoja.getDataRange().getValues();
    var LA = SCHEMA.LAYOUT_ALMACEN;
    var map = {};
    for (var i = 1; i < d.length; i++) {
      var rawEst  = d[i][LA.ESTANTE];
      var rawPano = d[i][LA.PANO];
      // Si Sheets convirtió algún valor a Date, la fila está corrupta — saltear
      if (rawEst instanceof Date || rawPano instanceof Date) continue;
      var est     = String(rawEst  || '').trim();
      var ordEst  = Number(d[i][LA.ORDEN_ESTANTE]) || 0;
      var pano    = String(rawPano || '').trim();
      var ordPano = Number(d[i][LA.ORDEN_PANO]) || 0;
      var alt     = Math.max(1, Number(d[i][LA.NUM_ALTURAS]) || 1);
      if (!est || !pano) continue;
      if (!map[est]) map[est] = { estante: est, orden: ordEst, panos: [] };
      map[est].panos.push({ pano: pano, orden: ordPano, alturas: alt });
    }
    var estantes = [];
    var keys = Object.keys(map);
    for (var k = 0; k < keys.length; k++) estantes.push(map[keys[k]]);
    estantes.sort(function(a, b) {
      var na = parseFloat(a.estante), nb = parseFloat(b.estante);
      if (!isNaN(na) && !isNaN(nb)) return na - nb;
      return a.orden - b.orden || String(a.estante).localeCompare(String(b.estante));
    });
    for (var e = 0; e < estantes.length; e++) {
      estantes[e].panos.sort(function(a, b) { return a.orden - b.orden || String(a.pano).localeCompare(String(b.pano)); });
    }
    return { ok: true, estantes: estantes };
  } catch(e) {
    Logger.log('SM_cargarLayout: ' + e);
    return { ok: false, error: e.message };
  }
}

function SM_guardarLayout(estantes) {
  try {
    var ss   = getSS();
    var hoja = ss.getSheetByName(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    if (!hoja) hoja = ss.insertSheet(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    hoja.clearContents();
    // Forzar cols A y C como texto antes de escribir para que Sheets no
    // convierta valores numéricos/alfanuméricos a fechas automáticamente
    hoja.getRange('A:A').setNumberFormat('@');
    hoja.getRange('C:C').setNumberFormat('@');
    hoja.getRange(1, 1).setValue('ESTANTE');
    hoja.getRange(1, 2).setValue('ORDEN_ESTANTE');
    hoja.getRange(1, 3).setValue('PAÑO');
    hoja.getRange(1, 4).setValue('ORDEN_PAÑO');
    hoja.getRange(1, 5).setValue('NUM_ALTURAS');
    hoja.getRange(1, 1, 1, 5).setFontWeight('bold');
    var lr = 2; // fila 1 = header
    for (var i = 0; i < estantes.length; i++) {
      var e = estantes[i];
      for (var j = 0; j < e.panos.length; j++) {
        var p = e.panos[j];
        hoja.getRange(lr, 1).setNumberFormat('@').setValue(String(e.estante || ''));
        hoja.getRange(lr, 2).setValue(i + 1);
        hoja.getRange(lr, 3).setNumberFormat('@').setValue(String(p.pano || ''));
        hoja.getRange(lr, 4).setValue(j + 1);
        hoja.getRange(lr, 5).setValue(Math.max(1, Number(p.alturas) || 1));
        lr++;
      }
    }
    invalidateSheetValues(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    Logger.log('SM_guardarLayout: ' + estantes.length + ' estantes guardados');
    return { ok: true };
  } catch(e) {
    Logger.log('SM_guardarLayout: ' + e);
    return { ok: false, error: e.message };
  }
}
