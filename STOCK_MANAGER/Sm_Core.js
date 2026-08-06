// ── STOCK MANAGER — Core ─────────────────────────────────────
// ============================================================
//  STOCK MANAGER BIDCOMAGRO v4.8 — SM_Core.gs
//  Variables globales, Carmen helpers, utilidades base
// ============================================================

var MASTER_SHEET_ID = "1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc";

var SM_CONFIG = {
  EMAIL_SUPERVISOR:  "soporteagrasdji@bidcom.com.ar",
  NOMBRE_REMITENTE:  "BIDCOMAGRO · Stock Manager",
  DIAS_STOCK_RECOM:  90,
  ALERTA_CATEGORIAS: ["A","B"]
};


// SINGLETON
var _SS = null;
function getSS() {
  if (!_SS) _SS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  return _SS;
}
function getSheet(nombre) { return getSS().getSheetByName(nombre); }

// ── CARMEN — spreadsheet externo, fuente de verdad para stock actual ──
var CARMEN_SS_ID         = '1-BH5m-LXFYhBZxqpSFVhIz5jwzFgJmLWH8Qvkh4PSCI';
var _carmenSS            = null;
var _carmenStockMapCache = null;

function _getCarmenSS() {
  if (!_carmenSS) _carmenSS = SpreadsheetApp.openById(CARMEN_SS_ID);
  return _carmenSS;
}

// Retorna { SKU → stockActual } desde hoja "STOCK" de Carmen (col A=codigo, col C=stock)
function _getCarmenStockMap() {
  if (_carmenStockMapCache) return _carmenStockMapCache;
  try {
    var hoja = _getCarmenSS().getSheetByName('STOCK');
    if (!hoja) return {};
    var d = hoja.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var c = String(d[i][0] || '').trim().toUpperCase();
      if (c) m[c] = parseInt(d[i][2]) || 0;
    }
    _carmenStockMapCache = m;
    return m;
  } catch(e) { Logger.log('_getCarmenStockMap: ' + e); return {}; }
}

// Retorna { SKU → [{ubicacion, cantidad}] } desde tab UBICACIONES de Carmen
function _getCarmenUbicMap() {
  var m = {};
  try {
    var hoja = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hoja) return m;
    var d = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var sku  = String(d[i][0] || '').trim().toUpperCase();
      var ubic = String(d[i][1] || '').trim();
      var cant = parseInt(d[i][2]) || 0;
      if (!sku || !ubic) continue;
      if (!m[sku]) m[sku] = [];
      m[sku].push({ ubicacion: ubic, cantidad: cant });
    }
  } catch(e) { Logger.log('_getCarmenUbicMap: ' + e); }
  return m;
}

// Registra un movimiento en Carmen (Entregados/Recibidos) y actualiza tally en UBICACIONES
// diff > 0 → entrada (Recibidos); diff < 0 → salida (Entregados)
// UBICACIONES col C = número directo (tally); STOCK col C de Carmen = fórmula de Carmen, nunca se toca
function _registrarMovimientoCarmen(sku, desc, ubicacion, diff, referencia) {
  try {
    var ss      = _getCarmenSS();
    var codKey  = String(sku       || '').trim().toUpperCase();
    var ubicKey = String(ubicacion || '').trim();
    var refStr  = String(referencia || 'Ajuste').trim();
    var descStr = String(desc      || '').trim();
    var cant    = Math.abs(diff);
    var fecha   = new Date();

    if (diff < 0) {
      var hojaEnt = ss.getSheetByName(CARMEN_ENTREGADOS_TAB);
      if (hojaEnt) hojaEnt.appendRow([codKey, descStr, cant, refStr, '', '', fecha, ubicKey]);
    } else if (diff > 0) {
      var hojaRec = ss.getSheetByName(CARMEN_RECIBIDOS_TAB);
      if (hojaRec) hojaRec.appendRow([codKey, descStr, cant, refStr, fecha, '', '', '', ubicKey]);
    }

    // Actualizar tally en UBICACIONES (col C = número directo, sin fórmula)
    if (ubicKey && diff !== 0) {
      var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
      if (hojaUbic) {
        var dU = hojaUbic.getDataRange().getValues();
        for (var i = 1; i < dU.length; i++) {
          if (String(dU[i][0] || '').trim().toUpperCase() === codKey &&
              String(dU[i][1] || '').trim().toUpperCase() === ubicKey) {
            var nueva = Math.max(0, (parseFloat(dU[i][2]) || 0) + diff);
            hojaUbic.getRange(i + 1, 3).setValue(nueva);
            return;
          }
        }
        // Fila nueva
        hojaUbic.appendRow([codKey, ubicKey, Math.max(0, diff)]);
      }
    }
  } catch(e) { Logger.log('_registrarMovimientoCarmen: ' + e); }
}

// Actualiza col C (stock actual) de un SKU en la hoja "STOCK" de Carmen
function _actualizarCarmenStock(sku, nuevoSaldo) {
  try {
    var hoja    = _getCarmenSS().getSheetByName('STOCK');
    if (!hoja) return;
    var codBusc = String(sku).trim().toUpperCase();
    var d       = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() === codBusc) {
        hoja.getRange(i + 1, 3).setValue(nuevoSaldo);
        _carmenStockMapCache = null;
        return;
      }
    }
    Logger.log('_actualizarCarmenStock: SKU no encontrado en Carmen: ' + sku);
  } catch(e) { Logger.log('_actualizarCarmenStock: ' + e); }
}

// Restaura la fórmula de Carmen STOCK col C en todas las filas que tienen un número plano.
// Ejecutar UNA VEZ desde el editor de Apps Script después de actualizar el código.
// La fórmula asume: col A = SKU, col F = stock base, pivots 'Entreg dinámica' y 'Recib dinámica'.
function restaurarFormulasCarmenStock() {
  try {
    var hoja = _getCarmenSS().getSheetByName('STOCK');
    if (!hoja) return { ok: false, error: 'Hoja STOCK no encontrada en Carmen' };
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { ok: true, restauradas: 0 };
    var formulas  = hoja.getRange(2, 3, lastRow - 1, 1).getFormulas();
    var newForms  = [];
    var restauradas = 0;
    for (var i = 0; i < formulas.length; i++) {
      var row = i + 2;
      if (!formulas[i][0]) {
        // Celda sin fórmula (número plano) — restaurar
        newForms.push(['=F' + row + '-(SI.ERROR(BUSCARV(A' + row + ',\'Entreg dinámica\'!A:B,2,0),0))+(SI.ERROR(BUSCARV(A' + row + ',\'Recib dinámica\'!A:B,2,0),0))']);
        restauradas++;
      } else {
        newForms.push([formulas[i][0]]); // ya tiene fórmula, no tocar
      }
    }
    hoja.getRange(2, 3, lastRow - 1, 1).setFormulas(newForms);
    SpreadsheetApp.flush();
    Logger.log('restaurarFormulasCarmenStock: ' + restauradas + ' filas restauradas de ' + (lastRow - 1) + ' totales');
    return { ok: true, restauradas: restauradas, total: lastRow - 1 };
  } catch(e) {
    Logger.log('restaurarFormulasCarmenStock ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── ASEGURAR HOJAS ───────────────────────────────────────────
function asegurarHojas() {
  var ss = getSS();
  var hojas = {
    "STOCK_REPUESTOS":    ["Código","Descripción","Stock actual","Stock mínimo","Categoría","Ubicación","Modelos compatibles","Última entrada","Última salida"],
    "MOVIMIENTOS_STOCK":  ["Fecha","Tipo","Código","Descripción","Cantidad","Stock resultante","Referencia","Operador","Observaciones","Depósito"],
    "COMPRAS_DJI":        ["CAS","Fecha pedido","Estado","Método pago","Fecha comprado","Fecha pagado","Fecha conf. envío","Fecha forwarder HK","Fecha vuelo","Fecha aduana","Fecha depósito","Operador","Observaciones","Última actualización"],
    "HISTORIAL_COMPRAS":  ["Fecha","ID_CAS","Estado anterior","Estado nuevo","Operador","Observaciones"],
    "CATALOGO_DJI":       ["Material Number","Simplified part number","代理商系统名称优化","English name","Applicable models","unit","pcs","South America"],
    "VENTAS_DIRECTAS":    ["Fecha","N° Orden entrega","N° Factura","Reseller","Código","Descripción","Cantidad","Precio USD","Observaciones"],
    "SOLICITUDES_DESPACHO": ["ID","Fecha","OT","Reseller","Código","Descripción","Cant. solicitada","Cant. despachada","Estado","Urgencia","Fecha despacho","Operador","Observaciones"],
    "STOCK_UBICACIONES":    ["SKU","Ubicación","Cantidad"],
    "TABLA_POSICIONES":     ["SKU","BIN_ID","CANTIDAD","TIPO_ALMACEN"],
    "LAYOUT_ALMACEN":       ["ESTANTE","ORDEN_ESTANTE","PAÑO","ORDEN_PAÑO","NUM_ALTURAS"]
  };
  var keys = Object.keys(hojas);
  for (var i = 0; i < keys.length; i++) {
    var nombre = keys[i];
    if (!ss.getSheetByName(nombre)) {
      var h = ss.insertSheet(nombre);
      h.appendRow(hojas[nombre]);
      h.getRange(1, 1, 1, hojas[nombre].length).setFontWeight("bold");
    }
  }
}

// Se bumpea a mano junto con "<!-- @version X.Y -->" de SM_Index.html — el cliente la trae
// embebida al cargar la página y la vuelve a consultar cada tanto (SM_obtenerVersionActual)
// para avisar si quedó una pestaña vieja abierta. Ver _smChequearVersionNueva en SM_Index.html.
var SM_VERSION = '5.43';

function SM_obtenerVersionActual() { return SM_VERSION; }

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet() {
  asegurarHojas();
  var tmpl = HtmlService.createTemplateFromFile('SM_Index');
  tmpl.DEPLOY_URL = ScriptApp.getService().getUrl();
  tmpl.SM_VERSION = SM_VERSION;
  return tmpl.evaluate()
    .setTitle('Stock Manager - BIDCOMAGRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
//  CARGA INICIAL — limpieza de categorías numéricas
// ============================================================
// Limpia categorías numéricas en STOCK_REPUESTOS y DB_REPUESTOS.
// Se llama automáticamente al cargar el dashboard la primera vez después de un deploy.
function _limpiarCatsNumericas() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('cats_limpias_v2')) return;  // ya corrió
    var changed = false;
    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    if (hojaStock) {
      var dS = hojaStock.getDataRange().getValues();
      for (var i = 1; i < dS.length; i++) {
        if (dS[i][4] && !_catValida(String(dS[i][4]))) { dS[i][4] = ''; changed = true; }
      }
      if (changed) { hojaStock.getDataRange().setValues(dS); invalidateSheetValues(SCHEMA.SHEETS.STOCK); }
    }
    var repChanged = false;
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaRep) {
      var dR = hojaRep.getDataRange().getValues();
      for (var j = 1; j < dR.length; j++) {
        if (dR[j][5] && !_catValida(String(dR[j][5]))) { dR[j][5] = ''; repChanged = true; }
      }
      if (repChanged) { hojaRep.getDataRange().setValues(dR); invalidateSheetValues(SCHEMA.SHEETS.DB_REPUESTOS); }
    }
    props.setProperty('cats_limpias_v2', '1');
  } catch(e) { Logger.log('_limpiarCatsNumericas: ' + e); }
}

function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail().trim().toLowerCase();
    var d = getSheetValues(SCHEMA.SHEETS.USUARIOS);
    // Usuarios_Internos layout: col 0=nombre, col 1=email, col 2=rol, col 3=esTecnico
    for (var i = 1; i < d.length; i++) {
      var rowEmail = String(d[i][1] || '').trim().toLowerCase();
      if (rowEmail === email) {
        var rol = String(d[i][2] || '').trim().toLowerCase();
        return (rol === 'admin') ? 'ADMIN' : 'OPERADOR';
      }
    }
    return 'OPERADOR';
  } catch(e) {
    return 'OPERADOR';
  }
}

// Lista de operadores para los selectores del front. Toma los nombres de la hoja
// Usuarios_Internos (col 0) SOLO de los que son técnicos (col D / índice 3 == "si"),
// únicos y ordenados. Reemplaza las listas hardcodeadas.
function getOperadores() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.USUARIOS);  // Usuarios_Internos: 0=nombre,1=email,2=rol,3=esTecnico
    var vistos = {}, out = [];
    for (var i = 1; i < d.length; i++) {
      var esTecnico = String(d[i][3] || '').trim().toLowerCase();  // col D
      if (esTecnico !== 'si' && esTecnico !== 'sí') continue;      // solo técnicos
      var nombre = String(d[i][0] || '').trim();
      if (!nombre) continue;
      var key = nombre.toLowerCase();
      if (vistos[key]) continue;
      vistos[key] = true;
      out.push(nombre);
    }
    out.sort(function(a, b) { return a.localeCompare(b, 'es'); });
    return { ok: true, operadores: out };
  } catch(e) {
    Logger.log('getOperadores: ' + e);
    return { ok: false, operadores: [], error: e.toString() };
  }
}

// ============================================================
//  HELPERS INTERNOS
// ============================================================

// Devuelve true si el valor es una categoría de texto válida (no un número puro).
// Previene que columnas numéricas del Excel DJI (unit, pcs) contaminen el campo categoría.
function _catValida(v) {
  var s = String(v || '').trim();
  return s.length > 0 && isNaN(Number(s));
}

function _fmtFecha(v) {
  if (!v || !(v instanceof Date)) return "";
  try { return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy"); }
  catch(e) { return ""; }
}

function _registrarMovimiento(tipo, codigo, desc, cantidad, stockResult, referencia, operador, deposito) {
  var hojaLedger = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
  if (!hojaLedger) {
    var err = '[_registrarMovimiento] Hoja MOVIMIENTOS_STOCK no encontrada. tipo=' + tipo + ' sku=' + codigo;
    Logger.log(err);
    console.error(err);
    throw new Error('WMS_LEDGER_WRITE_FAIL: ' + err);
  }
  hojaLedger.appendRow([
    new Date(), tipo, codigo, desc, cantidad, stockResult,
    referencia || '', operador || '', '', String(deposito || 'BA')
  ]);
}

// ============================================================
//  REGISTRAR MOVIMIENTO SEGURO
//  Acepta firma posicional o por objeto:
//    registrarMovimientoSeguro(sku, cantidad, tipo, motivo, operador)
//    registrarMovimientoSeguro({ sku, cantidad, tipo, motivo, usuario, nuevaUbicacion })
//  tipo: 'INGRESO' | 'EGRESO'
//  nuevaUbicacion (opcional): actualiza la ubicación física del ítem junto con el movimiento.
//  Valida stock antes de egresar, actualiza timestamp,
//  registra en MOVIMIENTOS_STOCK y hace flush atómico.
// ============================================================
function registrarMovimientoSeguro(sku, cantidad, tipo, motivo, operador) {
  // Soporte para firma por objeto: { sku, cantidad, tipo, motivo, usuario, nuevaUbicacion }
  var nuevaUbicacion = null;
  if (sku && typeof sku === 'object') {
    var d = sku;
    nuevaUbicacion = d.nuevaUbicacion || null;
    cantidad       = d.cantidad;
    tipo           = d.tipo;
    motivo         = d.motivo         || d.referencia || '';
    operador       = d.usuario        || d.operador   || '';
    sku            = d.sku;
  }

  var codBusc = String(sku || '').trim().toUpperCase();
  var cant    = parseInt(cantidad) || 0;
  var tipoStr = String(tipo || 'EGRESO').trim();

  // Delegar al ledger engine para la lógica de lock+verify+append+sync
  var delta  = (tipoStr === 'EGRESO') ? -cant : cant;
  var result = registrarEventoLedgerSeguro({
    sku:        codBusc,
    delta:      delta,
    tipo:       tipoStr,
    referencia: motivo   || '',
    operador:   operador || ''
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  // Actualizar ubicación si fue solicitada (no la maneja el ledger engine genérico)
  if (nuevaUbicacion) {
    try {
      var hojaStr = getSheet(SCHEMA.SHEETS.STOCK);
      var dStr    = getSheetValues(hojaStr);
      var S       = SCHEMA.STOCK_REPUESTOS;
      for (var i = 1; i < dStr.length; i++) {
        if (String(dStr[i][S.CODIGO] || '').trim().toUpperCase() !== codBusc) continue;
        hojaStr.getRange(i + 1, S.UBICACION + 1).setValue(String(nuevaUbicacion).trim());
        invalidateSheetValues(SCHEMA.SHEETS.STOCK);
        break;
      }
    } catch(eUbic) {
      Logger.log('[registrarMovimientoSeguro] ubicacion update failed: ' + eUbic.message);
    }
  }

  return {
    success:    true,
    stockNuevo: result.saldoResultante,
    ubicacion:  nuevaUbicacion || ''
  };
}

// ============================================================
//  MODELOS ACTIVOS — configuración persistente en PropertiesService
// ============================================================
function getModelosActivos() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('sm_modelos_activos');
    if (!raw) return [];
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch(e) {
    Logger.log('getModelosActivos: ' + e);
    return [];
  }
}

function setModelosActivos(modelos) {
  try {
    var lista = Array.isArray(modelos) ? modelos : [];
    // Normalizar: trim, sin duplicados, sin vacíos
    var mapa = {};
    var limpia = [];
    for (var i = 0; i < lista.length; i++) {
      var m = String(lista[i]).trim();
      if (m && !mapa[m.toUpperCase()]) {
        mapa[m.toUpperCase()] = true;
        limpia.push(m);
      }
    }
    PropertiesService.getScriptProperties().setProperty('sm_modelos_activos', JSON.stringify(limpia));
    return { ok: true, modelos: limpia };
  } catch(e) {
    Logger.log('setModelosActivos: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

/**
 * Extrae los nombres de modelos únicos desde DB_REPUESTOS (col 3 = MODELOS).
 * Separa por coma, barra, punto y coma y retorna lista ordenada sin duplicados.
 * Usada para popular el datalist de autocomplete en el generador de pedido.
 */
function getModelosDisponibles() {
  try {
    var dRep = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var mapa = {};
    for (var i = 1; i < dRep.length; i++) {
      var raw = String(dRep[i][3] || '').trim();
      if (!raw) continue;
      var partes = raw.split(/[,\/;|]+/);
      for (var p = 0; p < partes.length; p++) {
        var m = partes[p].trim();
        if (m && m.length > 1) mapa[m] = true;
      }
    }
    return Object.keys(mapa).sort();
  } catch(e) {
    Logger.log('getModelosDisponibles: ' + e);
    return [];
  }
}

// Parsea BIN_ID "ESTANTE-NIVEL" para ordenamiento lógico de ruta
function _binSortKey(binId) {
  var parts = String(binId || '').trim().split('-');
  var pasillo = (parts[0] || '').toUpperCase();
  var estante = parts[1] ? (isNaN(parts[1]) ? parts[1] : ('00' + parts[1]).slice(-3)) : '';
  var nivel   = parts[2] ? (isNaN(parts[2]) ? parts[2] : ('00' + parts[2]).slice(-3)) : '';
  return pasillo + '|' + estante + '|' + nivel;
}
