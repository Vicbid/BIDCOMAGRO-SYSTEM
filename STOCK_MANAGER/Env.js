var MASTER_SHEET_ID        = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc';
var WOS_NOTAS_SS_ID        = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw'; // WOS — contiene Pedidos_resellers y Pedidos_OTs
var CATALOGO_REPUESTOS_ID  = '1DWjX4JxHskP1uHa7YXTPpbgh2MD35hs43SpUvhP9Vn0';
var CARMEN_UBICACIONES_TAB = 'UBICACIONES'; // tab resumen WMS (fórmulas SUMIFS, solo lectura)
var CARMEN_ENTREGADOS_TAB  = 'Entregados';  // log de salidas en Carmen
var CARMEN_RECIBIDOS_TAB   = 'Recibidos';   // log de entradas en Carmen

var SCHEMA = {
  SHEETS: {
    OT: 'Ordenes de trabajo',
    STOCK: 'STOCK_REPUESTOS',
    RESELLERS: 'Resellers',
    SOLICITUDES: 'SOLICITUDES_DESPACHO',
    VENTAS_DIRECTAS: 'VENTAS_DIRECTAS',
    COMPRAS: 'COMPRAS_DJI',
    MOVIMIENTOS: 'MOVIMIENTOS_STOCK',
    DB_REPUESTOS: 'DB_REPUESTOS',
    LOGS: 'LOGS',
    EMAIL_LOGS: 'EMAIL_LOGS',
    EQUIPOS: 'EQUIPOS',
    USUARIOS: 'Usuarios_Internos',
    DEUDA_RESELLERS: 'DEUDA_RESELLERS',
    PRECIOS_MANO_OBRA: 'Precios_mano_obra',
    RESERVAS: 'RESERVAS_STOCK',
    COMPRAS_DETALLE: 'COMPRAS_DETALLE',
    USUARIOS_CONFIG: 'USUARIOS_CONFIG',
    STOCK_UBICACIONES: 'STOCK_UBICACIONES',
    TABLA_POSICIONES:  'TABLA_POSICIONES',
    LAYOUT_ALMACEN:    'LAYOUT_ALMACEN',
    HISTORIAL_COMPRAS: 'HISTORIAL_COMPRAS',
    CATALOGO_DJI:      'CATALOGO_DJI'
  },
  OT: {
    FECHA_INGRESO: 0,
    FECHA_CIERRE: 1,
    OT: 2,
    GARANTIA: 3,
    ESTADO: 4,
    EQUIPO: 5,
    SN: 6,
    RESELLER: 7,
    TECNICO: 9,
    MENSAJES: 11,
    TRABAJO: 12,
    FACTURA: 13,
    CAS: 14,
    REPUESTOS: 16,
    PRIORIDAD: 17,
    CIRCUITO: 18,
    MANO_OBRA: 22,
    NOTAS_INTERNAS: 23
  },
  STOCK_REPUESTOS: {
    CODIGO: 0,
    DESCRIPCION: 1,
    STOCK_ACTUAL: 2,
    STOCK_MINIMO: 3,
    CATEGORIA: 4,
    UBICACION: 5,
    MODELOS: 6,
    ULTIMA_ENTRADA: 7,
    ULTIMA_SALIDA: 8,
    REQUIERE_SN: 9
  },
  SOLICITUDES_DESPACHO: {
    ID: 0,
    FECHA: 1,
    OT: 2,
    RESELLER: 3,
    CODIGO: 4,
    DESCRIPCION: 5,
    CANT_SOLICITADA: 6,
    CANT_DESPACHADA: 7,
    ESTADO: 8,
    URGENCIA: 9,
    FECHA_DESPACHO: 10,
    OPERADOR: 11,
    OBSERVACIONES: 12
  },
  RESELLERS: {
    NOMBRE: 0,
    CUIT: 1,
    DIRECCION: 2,
    CP: 3,
    LOCALIDAD: 4,
    TELEFONO: 6,
    EMAIL: 9,
    PIN: 10
  },
  MOVIMIENTOS_STOCK: {
    FECHA:            0,
    TIPO:             1,
    CODIGO:           2,
    DESCRIPCION:      3,
    CANTIDAD:         4,
    STOCK_RESULTANTE: 5,
    REFERENCIA:       6,
    OPERADOR:         7,
    OBSERVACIONES:    8,
    DEPOSITO:         9
  },
  RESERVAS_STOCK: {
    ID:           0,
    FECHA:        1,
    SKU:          2,
    DESCRIPCION:  3,
    CANTIDAD:     4,
    ORIGEN:       5,
    ID_REFERENCIA:6,
    ESTADO:       7,
    CAS_REF:      8,
    OPERADOR:     9,
    OBSERVACIONES:10
  },
  COMPRAS_DJI: {
    ID_CAS:               0,
    FECHA_PEDIDO:         1,
    ESTADO:               2,
    METODO_PAGO:          3,
    FECHA_COMPRADO:       4,
    FECHA_PAGADO:         5,
    FECHA_ENVIO:          6,
    FECHA_FORWARDER:      7,
    FECHA_VUELO:          8,
    FECHA_ADUANA:         9,
    FECHA_DEPOSITO:       10,
    OPERADOR:             11,
    OBSERVACIONES:        12,
    ULTIMA_ACTUALIZACION: 13
  },
  HISTORIAL_COMPRAS: {
    FECHA:           0,
    ID_CAS:          1,
    ESTADO_ANTERIOR: 2,
    ESTADO_NUEVO:    3,
    OPERADOR:        4,
    OBSERVACIONES:   5
  },
  COMPRAS_DETALLE: {
    ID_CAS:            0,
    SKU:               1,
    DESCRIPCION:       2,
    CANTIDAD_PEDIDA:   3,
    CANTIDAD_RECIBIDA: 4,
    ESTADO:            5
  },
  USUARIOS_CONFIG: {
    EMAIL: 0,
    ROL:   1
  },
  STOCK_UBICACIONES: {
    SKU:      0,
    UBICACION: 1,
    CANTIDAD:  2
  },
  TABLA_POSICIONES: {
    SKU:          0,
    BIN_ID:       1,
    CANTIDAD:     2,
    TIPO_ALMACEN: 3
  },
  LAYOUT_ALMACEN: {
    ESTANTE:       0,
    ORDEN_ESTANTE: 1,
    PANO:          2,
    ORDEN_PANO:    3,
    NUM_ALTURAS:   4
  }
};



var _DB = null;
var CACHE_TTL = 60;

function getDb() {
  if (!_DB) {
    try {
      _DB = SpreadsheetApp.openById(MASTER_SHEET_ID);
    } catch(e) {
      _DB = SpreadsheetApp.getActiveSpreadsheet();
    }
  }
  return _DB;
}

function getSheet(nombre) {
  return getDb().getSheetByName(nombre);
}

function _cacheKey(nombre) {
  return 'sheetData:' + nombre;
}

function _serializeCell(value) {
  if (value instanceof Date) {
    return { __type: 'date', v: value.getTime() };
  }
  return value === undefined ? null : value;
}

function _deserializeCell(value) {
  if (value && typeof value === 'object' && value.__type === 'date') {
    return new Date(value.v);
  }
  return value;
}

function _serializeValues(values) {
  var out = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var outRow = [];
    for (var j = 0; j < row.length; j++) {
      outRow.push(_serializeCell(row[j]));
    }
    out.push(outRow);
  }
  return JSON.stringify(out);
}

function _deserializeValues(serialized) {
  var raw = JSON.parse(serialized);
  var out = [];
  for (var i = 0; i < raw.length; i++) {
    var row = raw[i];
    var outRow = [];
    for (var j = 0; j < row.length; j++) {
      outRow.push(_deserializeCell(row[j]));
    }
    out.push(outRow);
  }
  return out;
}

function getSheetValues(nombre, force) {
  var cache = CacheService.getScriptCache();
  var keyName = (nombre && typeof nombre.getName === 'function') ? nombre.getName() : String(nombre);
  var key = _cacheKey(keyName);
  if (!force) {
    var cached = cache.get(key);
    if (cached) {
      try {
        return _deserializeValues(cached);
      } catch(e) {}
    }
  }
  var sheet = (typeof nombre === 'string' || nombre instanceof String) ? getSheet(nombre) : nombre;
  var values = sheet ? sheet.getDataRange().getValues() : [];
  try {
    var payload = _serializeValues(values);
    if (payload.length < 90000) {
      cache.put(key, payload, CACHE_TTL);
    }
  } catch(e) {}
  return values;
}

function invalidateSheetValues(nombre) {
  var keyName = (nombre && typeof nombre.getName === 'function') ? nombre.getName() : String(nombre);
  CacheService.getScriptCache().remove(_cacheKey(keyName));
}

/**
 * Wrapper optimizado para CacheService con callback
 * @param {string} key Clave única de caché
 * @param {function} callback Función que devuelve datos si no hay caché
 * @param {number} ttl TTL en segundos (default: 600s = 10 min)
 * @returns {*} Datos cacheados o frescos
 */
function getCachedData(key, callback, ttl) {
  if (!ttl) ttl = 600;
  var cache = CacheService.getScriptCache();
  var cached = cache.get(key);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch(e) {}
  }
  var data = callback();
  try {
    cache.put(key, JSON.stringify(data), ttl);
  } catch(e) {}
  return data;
}
