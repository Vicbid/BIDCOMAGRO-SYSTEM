var MASTER_SHEET_ID = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc';

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
    RESERVAS: 'RESERVAS_STOCK'
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
    CLIENTE: 10,
    TECNICO: 9,
    MENSAJES: 11,
    TRABAJO: 12,
    FACTURA: 13,
    FECHA_ACTIVACION: 13,
    CAS: 14,
    REPUESTOS: 16,
    PRIORIDAD: 17,
    CIRCUITO: 18,
    FECHA_ESTADO: 20,
    MANO_OBRA: 22,
    NOTAS_INTERNAS: 23,
    HISTORIAL_ESTADOS: 24,
    ULTIMA_MODIFICACION: 25
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
    PROVINCIA: 5,
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
    OBSERVACIONES:    8
  },
  DB_REPUESTOS: {
    CODIGO: 1,
    DESCRIPCION: 2,
    MODELOS: 3
  },
  EQUIPOS: {
    NOMBRE: 0,
    TIPO:   1,
    PREFIJO:2,
    MESES:  3
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
