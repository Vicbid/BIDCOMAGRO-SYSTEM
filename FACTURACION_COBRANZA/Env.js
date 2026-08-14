// @version 1.2
// ============================================================
//  FACTURACION_COBRANZA — Configuración
//  Herramienta chica para 2 personas de administración (facturación y cobranza) que hoy
//  dependen de que les avisen por mail/WhatsApp qué pedido de repuestos de WOS hay que
//  facturar o cobrar. Este proyecto lee (SOLO LECTURA) los pedidos de WOS y de HUB_PRO,
//  y escribe únicamente en su propia spreadsheet — nunca toca las hojas de esos proyectos.
// ============================================================

// Pedidos de repuestos de WOS — SOLO LECTURA. Mismos IDs/nombres que WOS/Despacho_Env.js.
var NOTAS_SS_ID     = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
var HOJA_PEDIDOS    = 'Pedidos_resellers';
var HOJA_PEDIDOS_OT = 'Pedidos_OTs'; // mismo layout de columnas (COL) que Pedidos_resellers

// Órdenes de trabajo de HUB_PRO — SOLO LECTURA. Solo para resolver circuito (excluir
// "Taller", que consume el repuesto puertas adentro y nunca se factura a nadie) y el
// nombre del cliente de los pedidos ligados a una OT. Mismo ID que HUB_PRO/Env.js.
var MASTER_SHEET_ID = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc';
var HOJA_OT         = 'Ordenes de trabajo';
var COL_OT = { OT: 2, RESELLER: 7, CLIENTE: 10, CIRCUITO: 18 };
var CIRCUITO_TALLER = 'Taller';

// Columnas de Pedidos_resellers / Pedidos_OTs (0-based) — mismo layout que WOS/Despacho_Env.js COL.
var COL = {
  NUMERO:         0,   // A
  RESELLER:       1,   // B
  SKU:            2,   // C
  DESC:           3,   // D
  CANT_SOL:       4,   // E
  CANT_DESP:      5,   // F
  PRECIO:         7,   // H
  ESTADO:         9,   // J
  FECHA_DESPACHO: 14,  // O: fecha del despacho (cuando se envía a facturar)
  NOTA_ENTREGA:   15   // P
};

// Spreadsheet PROPIA de este proyecto. Acá y SOLO acá escribe este proyecto: las marcas de
// Facturado/Cobrado y la config de roles.
var FACT_SS_ID     = '1Mbq0LUoK7C-P2udSEGnmbBu7s-PwUOQvkxcVq4XKdbk';
var FACT_TAB       = 'PEDIDOS_FACTURACION';
var FACT_CONFIG_TAB = '_CONFIG';

// Columnas de la hoja PEDIDOS_FACTURACION (0-based) — 1 fila por NUMERO de pedido, no por línea.
// MONTO = base calculada (precio × cantidad despachada, sin impuesto — el impuesto varía por
// reseller y hoy no está sistematizado en ningún lado, así que NO se intenta calcular acá).
// MONTO_FACTURADO / IDVENTA los carga la propia facturación al marcar (puede ajustar el monto
// real facturado, con impuesto ya aplicado, y anotar el ID de venta de Masterchief para
// trazabilidad). MONTO_COBRADO / MONEDA_COBRADO los carga cobranza (puede diferir del
// facturado: distinto tipo de cambio, pago en pesos, etc. — se registra lo que realmente entró).
var COL_FACT = {
  NUMERO:           0,
  TIPO:             1,  // 'Reseller' | 'OT'
  CLIENTE:          2,
  MONTO:            3,  // base USD, sin impuesto (informativo/punto de partida)
  FECHA_DESPACHO:   4,
  FACTURADO:        5,
  FACTURADO_FECHA:  6,
  FACTURADO_POR:    7,
  IDVENTA:          8,  // ID de venta / factura en Masterchief (lo carga facturación)
  MONTO_FACTURADO:  9,  // USD, lo que realmente se facturó (con impuesto ya aplicado)
  COBRADO:          10,
  COBRADO_FECHA:    11,
  COBRADO_POR:      12,
  MONTO_COBRADO:    13, // lo que realmente entró
  MONEDA_COBRADO:   14  // 'USD' | 'ARS'
};

var FACT_HEADERS = [
  'Número', 'Tipo', 'Cliente', 'Monto base (USD)', 'Fecha despacho',
  'Facturado', 'Facturado fecha', 'Facturado por', 'ID Venta (Masterchief)', 'Monto facturado (USD)',
  'Cobrado', 'Cobrado fecha', 'Cobrado por', 'Monto cobrado', 'Moneda cobrado'
];

// _CONFIG: qué emails ven la pestaña de facturación y cuáles la de cobranza (separados por
// coma). Un email en las 2 listas ve las 2 pestañas. Vacío = nadie ve esa pestaña todavía
// (hay que cargar los mails acá antes de repartir la URL).
var FACT_CONFIG_DEFAULTS = [
  ['FACTURACION_EMAILS', ''],
  ['COBRANZA_EMAILS',    '']
];
