// @version 1.0
// SharedLogger.js — ES5 estricto
// Copiar este archivo a cada módulo (HUB_PRO, STOCK_MANAGER, PORTAL_RESELLER).
// Depende de: getSheet() definida en Env.js del módulo correspondiente.

var _LOGGER_NOMBRE_MODULO = "BIDCOMAGRO"; // sobreescribir en cada módulo si se prefiere

function registrarLogAuditoria(modulo, accion, detalle) {
  var usuario    = "";
  try { usuario = Session.getActiveUser().getEmail() || ""; } catch(e) {}
  if (!usuario) usuario = "Sistema[" + modulo + "]";

  var ahora     = new Date();
  var timestamp = ahora.toISOString();

  // Normalizar detalle: acepta string u objeto
  var detalleStr;
  try {
    detalleStr = (typeof detalle === "object" && detalle !== null)
      ? JSON.stringify(detalle)
      : String(detalle || "");
  } catch(e) {
    detalleStr = "[no-serializable]";
  }

  // 1. Hoja LOGS — consumo por dashboards internos
  try {
    var hojaLog = getSheet("LOGS");
    if (hojaLog) {
      hojaLog.appendRow([ahora, modulo, usuario, accion, detalleStr]);
    }
  } catch(eLog) {
    Logger.log("ERROR_LOG_SHEET | " + modulo + " | " + accion + " | " + eLog.toString());
  }

  // 2. GCP Cloud Logging — capa inmutable (requiere proyecto GCP estándar vinculado)
  // console.log genera un log inalterable en Cloud Logging con retención configurable.
  console.log(JSON.stringify({
    timestamp: timestamp,
    modulo:    modulo,
    operador:  usuario,
    accion:    accion,
    detalle:   detalleStr
  }));
}
