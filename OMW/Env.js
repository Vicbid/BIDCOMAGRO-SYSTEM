// @version 1.4
var OMW_SS_ID       = '1phTp8v59Gji90ZMZB9x-0vvLhYDyA2Bx-1yoY9WEQSs';
var PORTAL_SS_ID    = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc'; // Portal Reseller master sheet
var COTIZADOR_SS_ID = '1WiE7y1rQdg0bkY2_BkNr3jlwGsJs3aDv3Nd7V2bGEn0'; // Cotizador BidcomAgro
var TAB_CONDICIONES = 'Condiciones ExpoAgro 2026';

// Columnas tab Condiciones (0-based, col A = índice 0)
var CC = {
  CAT:     1,  // B  Categoría
  SKU:     2,  // C
  METODO:  3,  // D  Metodo Pago
  PVP:     4,  // E
  IVA:     5,  // F  % (decimal en GAS, ej. 0.105)
  COM_LL:  6,  // G  Comision Llena %
  DESC_CO: 7,  // H  DescComp %
  VOUCHER: 8,  // I
  CFT:     9,  // J  CFT Cliente % (negativo = descuento contado)
  DESC_RE: 10, // K  Descuento Reseller %
  PCT_COM: 11, // L  %Comision efectiva
  COMISION:12, // M  Comision USD
  SUB_CL:  13, // N  Subtotal Cliente USD
  IVA_CL:  14, // O  IVA Cliente USD
  TOTAL:   15  // P  Total Cliente USD
};

// Columnas PEDIDOS tab (0-based, expandido a 18 columnas)
var CP = {
  ID:         0,  // A  AG-DRON-XXXXX
  NUMERO:     1,  // B  AG-DRON-001
  RESELLER:   2,  // C
  CLIENTE:    3,  // D
  IMPORTE:    4,  // E  número USD
  PAGO:       5,  // F
  STAGE:      6,  // G  02_PENDIENTE_RTV | 03_APROBADO_ADMIN | 04_ACREDITADO | 99_RECHAZADO_RTV
  SLA_INICIO: 7,  // H  Date — se usa para calcular horas transcurridas
  ALERTA:     8,  // I
  OBS_RTV:    9,  // J
  REF_BANCO:  10, // K
  FECHA_CREA: 11, // L
  FECHA_MOD:  12, // M
  ID_COTIZ:   13, // N  Referencia a cotización original (COT-2026-XXXXX)
  SUBTOTAL:   14, // O  Congelado al formalizar
  DESCUENTO:  15, // P  Descuento aplicado (congelado)
  FECHA_FORMAL:16 // Q  Timestamp del traspaso atómico
};

// Stages definidos como constantes para visibilidad y mantenibilidad
var OMW_STAGES = {
  NUEVA_COTIZACION:      '01_NUEVA_COTIZACION',       // Solo en COTIZACIONES (no en PEDIDOS)
  CONVERTIDO_EN_PEDIDO:  'CONVERTIDO_EN_PEDIDO',      // Estado en COTIZACIONES cuando se formaliza
  PENDIENTE_RTV:         '02_PENDIENTE_RTV',          // Primer stage en PEDIDOS
  APROBADO_ADMIN:        '03_APROBADO_ADMIN',         // Tras aprobación RTV
  ACREDITADO:            '04_ACREDITADO',             // Tras acreditación Admin
  DESPACHO:              '05_EN_DESPACHO',            // En logística
  COMPLETADO:            '06_COMPLETADO',             // Fin normal
  RECHAZADO_RTV:         '99_RECHAZADO_RTV'           // Rechazado por RTV (re-editable por Reseller)
};

// Columnas PRODUCTOS tab (0-based)
var CPR = { ID_PED: 0, CANT: 1, NOMBRE: 2, SKU: 3 };

// Columnas MENSAJES tab (0-based)
var CM = { ID_PED: 0, ROL: 1, TEXTO: 2, TS: 3, EMAIL: 4 };

// Columnas USUARIOS tab (0-based)
var CU = { EMAIL: 0, NOMBRE: 1, ROL: 2, RESELLER: 3 };

// Nombres de tabs
var TAB_PEDIDOS       = 'PEDIDOS';
var TAB_PRODUCTOS     = 'PRODUCTOS';
var TAB_MENSAJES      = 'MENSAJES';
var TAB_USUARIOS      = 'USUARIOS';
var TAB_COTIZACIONES  = 'COTIZACIONES';

// Columnas COTIZACIONES tab (0-based, expandido para estados de formalización)
var CCOT = {
  ID:         0,  // A  COT-2026-XXXXX
  RESELLER:   1,  // B  empresa del reseller
  NOMBRE_CLI: 2,  // C
  EMPRESA_CLI:3,  // D
  CUIT:       4,  // E
  EMAIL_CLI:  5,  // F
  TEL:        6,  // G
  SKU:        7,  // H
  CAT:        8,  // I  categoria
  METODO:     9,  // J  metodo de pago
  PVP:        10, // K  precio venta publico USD
  SUBTOTAL:   11, // L  subtotal cliente USD
  IVA_CL:     12, // M  IVA cliente USD
  TOTAL:      13, // N  total cliente USD
  ESTADO:     14, // O  NUEVA | SEGUIMIENTO | CERRADA | CONVERTIDO_EN_PEDIDO
  FECHA_CREA: 15, // P
  NOTAS:      16  // Q
};
