// @version 1.16
// ============================================================
//  COMANDAS · CARGA MASTERCHIEF — Configuración
// ============================================================

// Sheet de VENTAS — SOLO LECTURA. El programa NUNCA escribe acá.
var CP_SS_ID        = '1KVaNIQVJPPTA-bAfDjReQanC5a24KOuHRNQQoq2R4sg';
var CP_TAB          = 'Ventas';
var CP_MASTER_TAB   = 'Comandas Master';  // misma planilla; estado de despacho + guía OCA (solo lectura)
var CP_DESPACHADO   = 'DESPACHADO';        // valor de col F que marca la comanda como despachada
var CP_RESELLERS_TAB = 'Resellers';        // misma planilla; nombre + mail del reseller (solo lectura)

// Sheet donde SÍ se escribe (separado de Ventas): registro del momento CARGAR
// y de las comandas ya cargadas en Masterchief por el operador.
var CP_LOG_SS_ID     = '1mOOeUDPORa9d1csQJ1fCL4ON592SvjMr1y3VKmWBN44';
var CP_LOG_TAB       = 'CARGAR_LOG';   // momento en que se marcó CARGAR
var CP_ENVIOS_TAB    = 'ENVIOS';       // 1 fila por ENVÍO (una venta puede tener varios): comanda + productos + guía
var CP_CONFIG_TAB    = '_CONFIG';      // parámetros del mail (destinatarios, asunto, URL tracking OCA)
var CP_RTV_TAB       = 'RTV';          // nombre + mail de cada RTV (para el mail final del Paso 5)
var CP_PEND_TAB      = 'PENDIENTES_ENTREGA'; // lo que falta enviar: 1 producto por línea
var CP_AUDIT_TAB     = 'AUDITORIA';    // trazabilidad: quién + cuándo + qué acción (crear/borrar envío, mails)
var CP_TIEMPOS_TAB   = 'TIEMPOS';      // 1 fila por envío: cuánto tardó cada etapa (CARGAR→comanda→autorización→despacho)

// Fecha de arranque del reporte de tiempos: un envío con autorización ANTERIOR a esta fecha se
// excluye ENTERO del reporte (ninguna etapa, ni siquiera "cargar comanda"). El usuario pidió esto
// porque los nombres de estado en Comandas Master cambiaron con el tiempo y eso ensuciaba los
// datos históricos de autorización/despacho — se arranca a medir limpio desde acá en adelante.
// Para reiniciar la medición de nuevo en el futuro, simplemente actualizá esta fecha.
var CP_TIEMPOS_DESDE = new Date(2026, 7, 6); // 06/08/2026 (mes 0-indexado: 7 = agosto)

// Sheet de KITS (receta/BOM): traduce el KIT (col C) a sus componentes (col A) a cargar
//   A = SKU componente | B = Descripción componente | C = KIT | D = Descripción Kit
//   La cantidad de cada componente = cantidad de filas repetidas para ese KIT.
var CP_KITS_SS_ID = '1WzE6yMz1MPVanrX5prwTOG5OvabSbvbDB0YBjOIq-oA';
var CP_KITS_TAB   = 'Actual';

// Carpeta de Drive con los PDF de las comandas ("Comandas Master C."). El nombre de cada PDF
// EMPIEZA con el N° de comanda (ej. "15861196 - FLO AGRO.pdf"). Se busca por ese número.
// (El ID anterior '1TotuLh...' apuntaba a una carpeta INEXISTENTE → nunca encontraba los PDF.)
var CP_PDF_FOLDER_ID = '11otuLnpmXa94ycEfyea1V7bletFNha4Y';

// Carpeta de Drive con los "documentos definidos" (manual, certificados, instructivos, etc.)
// que se adjuntan al reseller SOLO en el PRIMER envío de cada venta. Se adjuntan TODOS los
// archivos que estén en esta carpeta. Para cambiar qué se manda, agregás/sacás archivos de la
// carpeta (sin tocar código). VACÍO = no se adjunta ningún documento (solo el PDF de la comanda).
var CP_DOCS_FOLDER_ID = '';   // ← PEGAR ACÁ el ID de la carpeta de Drive con los documentos

// Valor en ID_Entrega que dispara el aviso "cargar en Masterchief"
var CP_FLAG    = 'CARGAR';

// Columnas de la hoja "Ventas" (0-based, col A = 0)
// Fecha | ID Negocio | ID_Venta | ID_Entrega | Operacion | Tipo | Condicion |
// MetodoPago | SKU | Cantidad | Descripcion | Alic IVA | Reseller | Razon Social |
// %Com | %Fact | P. Unitario | PV USD | IVA USD | Total USD | TC.Op | PV ARS |
// IVA ARS | Total ARS | Comentarios | RTV | Carta de Compromiso | Aprobacion Comercial
var VC = {
  FECHA:            0,  // A
  ID_NEGOCIO:       1,  // B
  ID_VENTA:         2,  // C
  ID_ENTREGA:       3,  // D
  OPERACION:        4,  // E
  TIPO:             5,  // F
  CONDICION:        6,  // G
  METODO_PAGO:      7,  // H
  SKU:              8,  // I
  CANTIDAD:         9,  // J
  DESCRIPCION:     10,  // K
  ALIC_IVA:        11,  // L
  RESELLER:        12,  // M
  RAZON_SOCIAL:    13,  // N
  PCT_COM:         14,  // O
  PCT_FACT:        15,  // P
  P_UNITARIO:      16,  // Q
  PV_USD:          17,  // R
  IVA_USD:         18,  // S
  TOTAL_USD:       19,  // T
  TC_OP:           20,  // U
  PV_ARS:          21,  // V
  IVA_ARS:         22,  // W
  TOTAL_ARS:       23,  // X
  COMENTARIOS:     24,  // Y
  RTV:             25,  // Z
  CARTA_COMPROMISO:26,  // AA
  APROB_COMERCIAL: 27   // AB
};
