// @version 1.12
var MASTER_SHEET_ID = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc';
var STOCK_SHEET_ID  = '1-BH5m-LXFYhBZxqpSFVhIz5jwzFgJmLWH8Qvkh4PSCI';
var LISTA_PRECIOS_SS_ID = '1DWjX4JxHskP1uHa7YXTPpbgh2MD35hs43SpUvhP9Vn0';

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
    PEDIDOS_REPUESTOS: 'PEDIDOS_REPUESTOS',
    ITEMS_SIN_CATALOGO: 'ITEMS_SIN_CATALOGO',
    LISTA_REPUESTOS:    'Lista_Repuestos',
    TABLA_POSICIONES:   'TABLA_POSICIONES',
    STOCK_INVENTARIO:   'STOCK',
    LOG_DEMANDA_PERDIDA:'LOG_DEMANDA_PERDIDA',
    PEDIDOS_CAMPANA:    'Pedidos_Campaña',
    COTIZACIONES:       'COTIZACIONES',
    KITS:               'KITS',
    NOVEDADES:          'NOVEDADES',
    EVENTOS:               'EVENTOS',
    INSCRIPCIONES_EVENTOS: 'INSCRIPCIONES_EVENTOS',
    VIDEOS_MODELOS:        'VIDEOS_ARMADO_DESARMADO',
    COMPRAS_DETALLE:       'COMPRAS_DETALLE',
    RESERVAS_EN_CAMINO:    'RESERVAS_EN_CAMINO',
    CONFIG_PROSPECTOS:     'CONFIG_PROSPECTOS'
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
    CLIENTE: 10,
    MENSAJES: 11,
    TRABAJO: 12,
    FECHA_ACTIVACION: 13,
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
    ULTIMA_SALIDA: 8
  },
  DB_REPUESTOS: {
    CODIGO: 1,
    DESCRIPCION: 2,
    MODELOS: 3,
    DESCRIPCION_ES: 7,   // Col H — descripción en español
    REEMPLAZADO_POR: 8   // Col I — SKU del repuesto reemplazante (vacío si vigente)
  },
  EQUIPOS: {
    NOMBRE: 0,
    TIPO: 1,
    PREFIJO: 2,
    MESES: 3
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
    OBSERVACIONES: 12,
    BIN_ID:        13,
    FORMA_PAGO:    14,
    ENVIO:         15
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
    PIN: 10,
    AFTERSALES: 11,
    EMAIL_RTV: 13,
    GRUPO: 14,      // Col O — nombre del grupo (igual para todas las sucursales del mismo dueño)
    PIN_GRUPO: 15,  // Col P — PIN del grupo (mismo valor en todas las filas del grupo)
    ACTIVO: 16,     // Col Q — estado de la cuenta: vacío = activo; "NO"/"BAJA"/"INACTIVO"/"FALSE"/"0" = desactivado
    DESCUENTO: 17   // Col R — % de descuento sobre precio de lista (vacío = 40%, histórico). 0 = sin descuento.
                    // Se localiza por encabezado "Descuento"/"Dto" si existe; este índice es el fallback.
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
  PEDIDOS_REPUESTOS: {
    ID:               0,
    FECHA:            1,
    RESELLER:         2,
    EMAIL:            3,
    CANT_ITEMS:       4,
    CANT_SIN_CATALOGO:5,
    ESTADO:           6,
    OBSERVACIONES:    7,
    ITEMS_JSON:       8,
    PDF_URL:          9,
    TOTAL_USD:        10,
    FORMA_PAGO:       11,
    ENVIO:            12,
    RTV_EMAIL:        13,  // Solo pedidos de prospecto (RS_Prospectos.js): email del RTV que lo cargó
    DTO_PROPUESTO_PCT:14   // Solo pedidos de prospecto: % de descuento que propuso el RTV — no definitivo, lo confirma/ajusta el autorizador
  },
  ITEMS_SIN_CATALOGO: {
    FECHA:      0,
    RESELLER:   1,
    SKU:        2,
    CANTIDAD:   3,
    PEDIDO_REF: 4
  },
  LISTA_REPUESTOS: {
    CODIGO:  0,
    PRECIO:  4
  },
  VENTAS_DIRECTAS: {
    FECHA:        0,
    TIPO:         1,
    REFERENCIA:   2,
    RESELLER:     3,
    EMAIL:        4,
    CODIGO:       5,
    DESCRIPCION:  6,
    CANTIDAD:     7,
    PRECIO_USD:   8,
    SUBTOTAL_USD: 9,
    ESTADO:       10
  },
  TABLA_POSICIONES: {
    BIN_ID:       0,
    CODIGO:       1,
    DESCRIPCION:  2,
    STOCK_EN_BIN: 3,
    ZONA:         4
  },
  STOCK_INVENTARIO: {
    CODIGO:       0,
    DESCRIPCION:  1,
    STOCK_ACTUAL: 2
  },
  PEDIDOS_CAMPANA: {
    TIMESTAMP:   0,  // Col A — fecha/hora de envío
    CAMPANA:     1,  // Col B — ID de campaña, e.g. "ago-2026"
    RESELLER:    2,  // Col C — nombre del reseller
    CODIGO:      3,  // Col D — SKU del repuesto
    DESCRIPCION: 4,  // Col E — descripción
    CANTIDAD:    5   // Col F — cantidad estimada
  }
};

var _DB       = null;
var _STOCK_DB = null;
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

// Quita tildes y convierte a minúsculas para comparaciones de búsqueda insensibles a acentos.
function _normText(s) {
  return String(s || '').toLowerCase()
    .replace(/[áàäâã]/g, 'a').replace(/[éèëê]/g, 'e')
    .replace(/[íìïî]/g, 'i').replace(/[óòöôõ]/g, 'o')
    .replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n');
}

function getStockDb() {
  if (!_STOCK_DB) {
    _STOCK_DB = SpreadsheetApp.openById(STOCK_SHEET_ID);
  }
  return _STOCK_DB;
}

function getStockSheetValues(nombre) {
  var cache = CacheService.getScriptCache();
  var key = _cacheKey('stk:' + nombre);
  var cached = cache.get(key);
  if (cached) {
    try { return _deserializeValues(cached); } catch(e) {}
  }
  var sheet;
  try {
    sheet = getStockDb().getSheetByName(nombre);
  } catch(e) {
    return [];
  }
  var values = sheet ? sheet.getDataRange().getValues() : [];
  try {
    var payload = _serializeValues(values);
    if (payload.length < 90000) cache.put(key, payload, CACHE_TTL);
  } catch(e) {}
  return values;
}
