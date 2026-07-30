// @version 1.2
var NOTAS_SS_ID            = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
var CARMEN_SS_ID           = '1-BH5m-LXFYhBZxqpSFVhIz5jwzFgJmLWH8Qvkh4PSCI';
var MASTER_SS_ID           = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc';
var PLANIF_SS_ID           = '1GxqNYGq9Uf4hyo2fobuyY5BF9h9Uox1N8wUv2BhUuy0';
var CARMEN_UBICACIONES_TAB = 'UBICACIONES'; // tab WMS en Carmen
var HOJA_PEDIDOS    = 'Pedidos_resellers';
var HOJA_PEDIDOS_OT = 'Pedidos_OTs'; // mismo schema (COL) que Pedidos_resellers — repuestos de reparaciones (HUB PRO)
// Maestro de artículos auto-construido: cuando un operario marca un SKU como "por bolsa"
// en la preparación, ese SKU queda registrado acá y la próxima vez viene pre-marcado.
var HOJA_MAESTRO    = 'MAESTRO_ARTICULOS'; // tab en NOTAS_SS_ID
var COL_MAESTRO = {
  SKU:       0,   // A
  DESC:      1,   // B: última descripción vista
  POR_BOLSA: 2,   // C: TRUE = se prepara por bolsa (1 código SO por bolsa de N unidades)
  BULTO:     3,   // D: tamaño de bolsa por defecto (ej. 50) — solo sugerencia
  FECHA:     4,   // E: última actualización
  OPERADOR:  5    // F: quién lo marcó por última vez
};
// Defaults — sobreescritos por WOS_CONFIG en MASTER si existe
var EMAIL_SOPORTE     = 'soporteagrasdji@bidcom.com.ar';
var EMAIL_FACTURACION = 'Cecilia.f@bidcom.com.ar,lucia.c@bidcom.com.ar';
var WOS_PDF_FOLDER_ID = '1yVefFM-vZ-Skmg2a_V9fsA-3XTAKx3Pz';

var _wosConfigCache = null;

function _wosConfig() {
  if (_wosConfigCache) return _wosConfigCache;
  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get('wos_config_v1');
    var map    = cached ? JSON.parse(cached) : null;
    if (!map) {
      map = {};
      var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('WOS_CONFIG');
      if (hoja) {
        var data = hoja.getDataRange().getValues();
        for (var i = 0; i < data.length; i++) {
          if (data[i][0]) map[String(data[i][0]).trim()] = String(data[i][1] || '').trim();
        }
      }
      try { cache.put('wos_config_v1', JSON.stringify(map), 300); } catch(eC) {}
    }
    _wosConfigCache = {
      emailSoporte: map['EMAIL_SOPORTE']     || EMAIL_SOPORTE,
      emailFact:    map['EMAIL_FACTURACION'] || EMAIL_FACTURACION,
      pdfFolderId:  map['WOS_PDF_FOLDER_ID'] || WOS_PDF_FOLDER_ID
    };
    return _wosConfigCache;
  } catch(e) {
    return { emailSoporte: EMAIL_SOPORTE, emailFact: EMAIL_FACTURACION, pdfFolderId: WOS_PDF_FOLDER_ID };
  }
}

var COL = {
  NUMERO:         0,   // A
  RESELLER:       1,   // B
  SKU:            2,   // C
  DESC:           3,   // D
  CANT_SOL:       4,   // E
  CANT_DESP:      5,   // F
  CANT_PEND:      6,   // G (fórmula =E-F-Z en sheet)
  PRECIO:         7,   // H
  STOCK_ORI:      8,   // I
  ESTADO:         9,   // J
  FECHA:          10,  // K
  ENVIO:          11,  // L
  PAGO:           12,  // M
  OBS:            13,  // N
  FECHA_DESPACHO: 14,  // O: fecha del despacho (cuando se envía a facturar)
  NOTA_ENTREGA:   15,  // P: código de nota de entrega, e.g. NE_PR-00028-01
  TRACKING:       16,  // Q: Número de seguimiento del envío
  THREAD_ID:      17,  // R: Gmail Thread ID del hilo ancla (Portal → col 18 en getRange)
  FECHA_ESTADO:       18,  // S: timestamp del último cambio de estado (WOS)
  TRANSPORTISTA_DESP: 19,  // T: transportista real usado al despachar
  COSTO_ENVIO:        20,  // U: costo del envío (ARS)
  PESO_ENVIO:         21,  // V: peso del envío en kg
  NE_URL:             22,  // W: link Drive de la Nota de Entrega
  OPERARIO:           23,  // X: email del operario que realizó la última acción
  SERIALES:           24,  // Y: números de serie despachados (separados por coma)
  CANT_CANCEL:        25,  // Z: unidades canceladas por reseller (Opción B) — fórmula CANT_PEND actualizar a =E-F-Z
  UBIC_PREP:          26,  // AA: bins elegidos al preparar, JSON [{bin,cant}] por ítem — descuento WMS se aplica al despachar
  PESO_PREP:          27,  // AB: peso exacto del paquete (kg) capturado al preparar — pre-llena el peso del bulto al despachar
};

var COL_RS = { NOMBRE: 0, EMAIL: 9 };

var EST = {
  PENDIENTE:      'Pendiente_Revision',
  CONFIRMADO:     'Confirmado',
  EN_ESPERA:      'En_Espera_Reseller',
  CANCELADO:      'Cancelado',
  PREPARADO:      'Preparado',
  BACKORDER:      'Backorder',
  PREP_PARCIAL:   'Preparado Parcial',
  ENTREGADO:      'Entregado_Cerrado',
  LISTO_RETIRO:   'Listo_Retiro',
  ENTREGADO_CONF: 'Entregado_Confirmado',
  // Solo pedidos de OT (ver _esNumeroOT): ítem YA disponible pero retenido a propósito
  // porque el contacto de la OT eligió "esperar y consolidar en 1 solo envío". A
  // diferencia de Backorder, NO significa "falta comprar/reponer" — no debe entrar en
  // WOS_reporteBackorder ni en el backorderPred del Command Center de HUB_PRO.
  RESERVADO_CONSOLIDAR: 'Reservado_Consolidar'
};

function _getHojaPedidos() {
  return SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName(HOJA_PEDIDOS);
}

function _getHojaPedidosOT() {
  return SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName(HOJA_PEDIDOS_OT);
}

// Resuelve a qué hoja pertenece un número de pedido — 'OT-' = repuestos de reparación, resto = Pedidos_resellers.
// Ambas hojas comparten el mismo layout de columnas (COL).
function _esNumeroOT(numero) {
  return String(numero || '').trim().toUpperCase().indexOf('OT-') === 0;
}
function _getHojaPorNumero(numero) {
  return _esNumeroOT(numero) ? _getHojaPedidosOT() : _getHojaPedidos();
}
