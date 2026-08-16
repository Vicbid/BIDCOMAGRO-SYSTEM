// @version 1.4

/* ════════════════════════════════════════════
   ENTRY POINT
════════════════════════════════════════════ */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('OMW_Index')
    .setTitle('OMW · Pipeline Equipos DJI AGRAS & Enterprise')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ════════════════════════════════════════════
   HELPERS INTERNOS
════════════════════════════════════════════ */
function _ss()       { return SpreadsheetApp.openById(OMW_SS_ID); }
function _portal()   { return SpreadsheetApp.openById(PORTAL_SS_ID); }
function _hoja(tab)  { var h = _ss().getSheetByName(tab); if (!h) throw new Error('Tab "'+tab+'" no encontrada.'); return h; }

function _fmtTs(v) {
  var d = v instanceof Date ? v : new Date(v);
  return Utilities.formatDate(d, 'America/Argentina/Buenos_Aires', 'dd/MM HH:mm');
}

function _horasDesde(v) {
  if (!v) return 0;
  var d = v instanceof Date ? v : new Date(v);
  return Math.floor((new Date() - d) / 3600000);
}

// Lee el último ID correlativo de PEDIDOS de forma segura.
// Formato esperado: 'AG-DRON-XXXXX' donde XXXXX es numérico.
// Retorna el siguiente número correlativo.
function _nextIdCorrelativo() {
  var d = _hoja(TAB_PEDIDOS).getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < d.length; i++) {
    var idStr = String(d[i][CP.ID] || '').trim();
    if (!idStr || idStr.indexOf('AG-DRON-') === -1) continue;
    var numPart = idStr.substring('AG-DRON-'.length);
    var v = parseInt(numPart);
    if (!isNaN(v) && v > max) max = v;
  }
  return max + 1;
}

// Valida que el callerEmail tenga uno de los roles permitidos en USUARIOS.
// Lanza error si no está autorizado. Retorna el rol.
function _validarCaller(callerEmail, rolesPermitidos) {
  var mailB = String(callerEmail || '').trim().toLowerCase();
  if (!mailB) throw new Error('Sesión inválida. Reiniciá la aplicación.');
  var d = _hoja(TAB_USUARIOS).getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][CU.EMAIL] || '').trim().toLowerCase() === mailB) {
      var rol = String(d[i][CU.ROL] || '').toLowerCase();
      if (rolesPermitidos.indexOf(rol) === -1) throw new Error('Tu rol no permite esta acción.');
      return rol;
    }
  }
  throw new Error('Usuario no autorizado.');
}

// Retorna EXCLUSIVAMENTE pedidos de la tabla PEDIDOS (no borradores de COTIZACIONES).
// Filtra solo registros cuyo stage sea >= '02_PENDIENTE_RTV' para evitar que borradores
// o estados pre-pipeline aparezcan en el Kanban de RTVs/Admin/Depósito.
// Agrupa por stage para el tablero: rtv, admin, depo, done, rechazado.
function _pedidosAll() {
  var dPed  = _hoja(TAB_PEDIDOS).getDataRange().getValues();
  var dProd = _hoja(TAB_PRODUCTOS).getDataRange().getValues();

  var prodMap = {};
  for (var j = 1; j < dProd.length; j++) {
    var pid = String(dProd[j][CPR.ID_PED]);
    if (!prodMap[pid]) prodMap[pid] = [];
    prodMap[pid].push({ q: parseInt(dProd[j][CPR.CANT]) || 1, n: String(dProd[j][CPR.NOMBRE] || '') });
  }

  var result = { rtv: [], admin: [], depo: [], done: [], rechazado: [] };
  for (var i = 1; i < dPed.length; i++) {
    var r = dPed[i];
    if (!r[CP.ID]) continue;
    
    var stageRaw = String(r[CP.STAGE] || '').trim();
    
    // FILTRO CRÍTICO: Solo mostrar pedidos en pipeline activo (>= '02_PENDIENTE_RTV')
    // Rechazados (99) TAMBIÉN se incluyen para que Resellers puedan corregir
    if (stageRaw.indexOf('01_') === 0) continue; // Excluir cotizaciones nuevas
    
    var idStr = String(r[CP.ID]);
    var horas = _horasDesde(r[CP.SLA_INICIO]);
    var imp   = parseFloat(r[CP.IMPORTE]) || 0;
    
    // Mapeo bidireccional: nuevo formato ('03_APROBADO_ADMIN') + viejo formato ('admin')
    var stageKey = 'rtv'; // default cubre '02_PENDIENTE_RTV' y viejo 'rtv'
    var sL = stageRaw.toLowerCase();
    if (stageRaw === OMW_STAGES.APROBADO_ADMIN || stageRaw === '03_APROBADO_ADMIN' || sL === 'admin') stageKey = 'admin';
    else if (stageRaw === OMW_STAGES.ACREDITADO || stageRaw === '04_ACREDITADO') stageKey = 'depo';
    else if (stageRaw === OMW_STAGES.DESPACHO || stageRaw === '05_EN_DESPACHO' || sL === 'depo') stageKey = 'depo';
    else if (stageRaw === OMW_STAGES.COMPLETADO || stageRaw === '06_COMPLETADO' || sL === 'done') stageKey = 'done';
    else if (stageRaw === OMW_STAGES.RECHAZADO_RTV || stageRaw === '99_RECHAZADO_RTV' || sL === 'rechazado') stageKey = 'rechazado';
    
    var ped = {
      id:       idStr,
      num:      String(r[CP.NUMERO]   || ''),
      reseller: String(r[CP.RESELLER] || ''),
      cliente:  String(r[CP.CLIENTE]  || ''),
      importe:  'USD ' + imp.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }),
      pago:     String(r[CP.PAGO]     || ''),
      stage:    stageRaw.toLowerCase(),
      horas:    horas,
      sla:      horas >= 24,
      alerta:   String(r[CP.ALERTA]   || ''),
      prods:    prodMap[idStr] || [],
      idCotiz:  String(r[CP.ID_COTIZ] || ''),
      obsRTV:   String(r[CP.OBS_RTV]  || ''),
      subtotal: parseFloat(r[CP.SUBTOTAL] || 0) || 0,
      descuento: parseFloat(r[CP.DESCUENTO] || 0) || 0,
      stage:    stageKey  // Enviar el stageKey normalizado ('rtv', 'admin', 'depo', 'done', 'rechazado')
    };
    
    if (result[stageKey]) result[stageKey].push(ped);
  }
  return result;
}

function _findRow(pedidoId) {
  var h = _hoja(TAB_PEDIDOS);
  var d = h.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][CP.ID]) === String(pedidoId)) return { h: h, row: i + 1 };
  }
  throw new Error('Pedido ' + pedidoId + ' no encontrado');
}

function _setCell(pedidoId, col, val) {
  var r = _findRow(pedidoId);
  r.h.getRange(r.row, col + 1).setValue(val);
  r.h.getRange(r.row, CP.FECHA_MOD + 1).setValue(new Date());
}

/* ════════════════════════════════════════════
   LOGIN — dos flujos separados
════════════════════════════════════════════ */

// Bidcom interno: valida email contra tab USUARIOS.
// Sin contraseña — el acceso a la URL ya es restringido internamente.
function OMW_loginInterno(email) {
  var mailB = String(email || '').trim().toLowerCase();
  if (!mailB) return { ok: false, error: 'Ingresá tu email' };
  try {
    var d = _hoja(TAB_USUARIOS).getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][CU.EMAIL] || '').trim().toLowerCase() !== mailB) continue;
      var rol = String(d[i][CU.ROL] || '').toLowerCase();
      if (!rol) return { ok: false, error: 'Usuario sin rol asignado. Contactá a Administración.' };
      return {
        ok:       true,
        email:    mailB,
        nombre:   String(d[i][CU.NOMBRE] || mailB),
        rol:      rol,
        reseller: ''
      };
    }
    return { ok: false, error: 'Email no encontrado en usuarios internos. Contactá a Administración.' };
  } catch(e) {
    return { ok: false, error: 'Error al verificar usuario: ' + e.message };
  }
}

// Reseller: valida nombre (Col A) + PIN (Col K) contra Portal Reseller master sheet.
// Mismo mecanismo que Portal Reseller (validarAccesoInicial).
function OMW_loginReseller(nombre, pin) {
  var nombreB = String(nombre || '').trim().toLowerCase();
  var pinB    = String(pin    || '').trim();
  if (!nombreB || !pinB) return { ok: false, error: 'Completá empresa y contraseña' };
  try {
    var h = _portal().getSheetByName('Resellers');
    if (!h) return { ok: false, error: 'Error al conectar con Portal Reseller' };
    var d = h.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var rNombre = String(d[i][0]  || '').trim().toLowerCase(); // Col A: NOMBRE
      var rPin    = String(d[i][10] || '').trim();               // Col K: PIN
      if (!rNombre || !rPin) continue;
      if (rNombre !== nombreB) continue;
      if (rPin !== pinB) return { ok: false, error: 'Contraseña incorrecta' };
      return {
        ok:       true,
        email:    String(d[i][9] || '').trim(), // Col J: EMAIL
        nombre:   String(d[i][0] || '').trim(), // Col A: NOMBRE
        rol:      'reseller',
        reseller: String(d[i][0] || '').trim()
      };
    }
    return { ok: false, error: 'Empresa no encontrada. Verificá la selección.' };
  } catch(e) {
    Logger.log('OMW_loginReseller: ' + e);
    return { ok: false, error: 'Error de conexión con Portal Reseller. Intentá de nuevo.' };
  }
}

// Retorna lista de nombres de resellers para el dropdown del login.
function OMW_getResellersPortal() {
  try {
    var h = _portal().getSheetByName('Resellers');
    if (!h) return [];
    var d = h.getDataRange().getValues();
    var list = [];
    for (var i = 1; i < d.length; i++) {
      var n = String(d[i][0] || '').trim();
      if (n) list.push(n);
    }
    list.sort();
    return list;
  } catch(e) {
    Logger.log('OMW_getResellersPortal: ' + e);
    return [];
  }
}

/* ════════════════════════════════════════════
   API PÚBLICA — datos
════════════════════════════════════════════ */

// Retorna todos los pedidos agrupados por stage.
// La identidad del usuario ya fue verificada en el login client-side.
function OMW_getPedidos() {
  return _pedidosAll();
}

function OMW_getMensajes(pedidoId) {
  var d = _hoja(TAB_MENSAJES).getDataRange().getValues();
  var msgs = [];
  for (var i = 1; i < d.length; i++) {
    if (String(d[i][CM.ID_PED]) === String(pedidoId)) {
      msgs.push({
        rol:   String(d[i][CM.ROL]   || ''),
        txt:   String(d[i][CM.TEXTO] || ''),
        ts:    _fmtTs(d[i][CM.TS]),
        email: String(d[i][CM.EMAIL] || '')
      });
    }
  }
  return msgs;
}

function OMW_sendMensaje(pedidoId, texto, callerEmail, callerRol) {
  var mailB = String(callerEmail || '').trim();
  var rolB  = String(callerRol  || '').trim().toLowerCase();
  if (!mailB || !rolB) throw new Error('Sesión inválida. Reiniciá la aplicación.');
  if (!String(texto || '').trim()) throw new Error('Mensaje vacío');
  _hoja(TAB_MENSAJES).appendRow([
    String(pedidoId), rolB, String(texto).trim(), new Date(), mailB
  ]);
  return OMW_getMensajes(pedidoId);
}

/* ════════════════════════════════════════════
   API PÚBLICA — acciones (validan caller)
════════════════════════════════════════════ */
function OMW_autorizarRTV(pedidoId, observaciones, callerEmail) {
  _validarCaller(callerEmail, ['rtv']);
  _setCell(pedidoId, CP.STAGE, OMW_STAGES.APROBADO_ADMIN);
  if (observaciones) _setCell(pedidoId, CP.OBS_RTV, String(observaciones));
  return { exito: true, mensaje: 'Pedido aprobado y enviado a Administración' };
}

function OMW_rechazarRTV(pedidoId, motivo, callerEmail) {
  _validarCaller(callerEmail, ['rtv']);
  _setCell(pedidoId, CP.STAGE, OMW_STAGES.RECHAZADO_RTV);
  if (motivo) _setCell(pedidoId, CP.OBS_RTV, String(motivo));
  return { exito: true, mensaje: 'Pedido rechazado y disponible para corrección por Reseller' };
}

function OMW_acreditarAdmin(pedidoId, refBancaria, callerEmail) {
  _validarCaller(callerEmail, ['admin', 'administracion']);
  _setCell(pedidoId, CP.STAGE, OMW_STAGES.ACREDITADO);
  if (refBancaria) _setCell(pedidoId, CP.REF_BANCO, String(refBancaria));
  return { exito: true, mensaje: 'Acreditación registrada, pedido enviado a logística' };
}

function OMW_generarComanda(pedidoId, callerEmail) {
  _validarCaller(callerEmail, ['logistica', 'admin']);
  _setCell(pedidoId, CP.STAGE, OMW_STAGES.COMPLETADO);
  return { exito: true, mensaje: 'Comanda generada, pedido completado' };
}

/* ════════════════════════════════════════════
   FORMALIZACIÓN ATÓMICA: Cotización → Pedido
════════════════════════════════════════════ */

// DIRECTIVA 2: Logística de Traspaso Atómico y Seguro
// Transforma una cotización de borrador a pedido en el pipeline activo.
// Utiliza LockService para evitar colisiones de escritura entre múltiples resellers.
// Esta es la ÚNICA forma de migrar una cotización a pedido; NO hay creación manual desde PEDIDOS.
function OMW_formalizarPedido(idCotizacion, resellerNombre, callerEmail) {
  var lock = LockService.getScriptLock();
  var lockAdquirido = false;
  
  try {
    // A) CONTROL DE CONCURRENCIA: Adquirir lock con timeout de 15 segundos
    lockAdquirido = lock.tryLock(15000);
    if (!lockAdquirido) {
      return {
        exito: false,
        mensaje: 'Sistema ocupado. Intenta nuevamente en 10 segundos.',
        codigoError: 'LOCK_TIMEOUT'
      };
    }
    
    // B) VERIFICACIÓN DE ESTADO: Confirmar que la cotización existe y NO está ya procesada
    var hCot = _hojaCotiz();
    var dCot = hCot.getDataRange().getValues();
    var cotizRow = -1;
    var cotizData = null;
    
    for (var i = 1; i < dCot.length; i++) {
      if (String(dCot[i][CCOT.ID]) === String(idCotizacion)) {
        cotizRow = i;
        cotizData = dCot[i];
        break;
      }
    }
    
    if (cotizRow === -1) {
      return {
        exito: false,
        mensaje: 'Cotización no encontrada: ' + idCotizacion,
        codigoError: 'COTIZ_NOT_FOUND'
      };
    }
    
    // Verificar que el estado NO sea ya 'CONVERTIDO_EN_PEDIDO'
    var estadoActual = String(cotizData[CCOT.ESTADO] || '');
    if (estadoActual === OMW_STAGES.CONVERTIDO_EN_PEDIDO) {
      return {
        exito: false,
        mensaje: 'Esta cotización ya fue procesada previamente y convertida en pedido.',
        codigoError: 'ALREADY_CONVERTED'
      };
    }
    
    // Verificar propiedad: el reseller que formaliza debe coincidir
    var resellerCotiz = String(cotizData[CCOT.RESELLER] || '').trim().toLowerCase();
    var resellerReq   = String(resellerNombre || '').trim().toLowerCase();
    if (resellerCotiz !== resellerReq) {
      return {
        exito: false,
        mensaje: 'No tienes permisos para formalizar esta cotización.',
        codigoError: 'UNAUTHORIZED'
      };
    }
    
    // C) EXTRACCIÓN DE DATOS: Congelar el snapshot de precios de la cotización
    var seq = _nextIdCorrelativo();
    var idPedido = 'AG-DRON-' + String(seq).padStart(5, '0');
    
    var reseller    = String(cotizData[CCOT.RESELLER]    || '');
    var nombreCli   = String(cotizData[CCOT.NOMBRE_CLI]  || '');
    var empresaCli  = String(cotizData[CCOT.EMPRESA_CLI] || '');
    var emailCli    = String(cotizData[CCOT.EMAIL_CLI]   || '');
    var sku         = String(cotizData[CCOT.SKU]         || '');
    var categoria   = String(cotizData[CCOT.CAT]         || '');
    var metodoPago  = String(cotizData[CCOT.METODO]      || '');
    var pvp         = parseFloat(cotizData[CCOT.PVP])     || 0;
    var subtotalCL  = parseFloat(cotizData[CCOT.SUBTOTAL]) || 0;
    var ivaCL       = parseFloat(cotizData[CCOT.IVA_CL])  || 0;
    var totalCL     = parseFloat(cotizData[CCOT.TOTAL])   || 0;
    
    // El sistema NO recalcula precios; respeta el valor pactado en la cotización
    var importe = totalCL;
    
    // D) INSERCIÓN LIMPIA CON CORRELATIVO SEGURO: Agregar fila a PEDIDOS con 18 columnas
    var hPed = _hoja(TAB_PEDIDOS);
    var ahora = new Date();
    
    hPed.appendRow([
      idPedido,                       // A: ID (AG-DRON-XXXXX)
      idPedido,                       // B: NUMERO (idéntico a ID)
      reseller,                       // C: RESELLER
      nombreCli,                      // D: CLIENTE (nombre contacto)
      importe,                        // E: IMPORTE (total congelado)
      metodoPago,                     // F: PAGO
      OMW_STAGES.PENDIENTE_RTV,       // G: STAGE (inicial '02_PENDIENTE_RTV')
      ahora,                          // H: SLA_INICIO (cuando entra al pipeline)
      '',                             // I: ALERTA (vacío inicialmente)
      '',                             // J: OBS_RTV (vacío inicialmente)
      '',                             // K: REF_BANCO (vacío inicialmente)
      ahora,                          // L: FECHA_CREA
      ahora,                          // M: FECHA_MOD
      idCotizacion,                   // N: ID_COTIZ (trazabilidad → COT-2026-XXXXX)
      subtotalCL,                     // O: SUBTOTAL (congelado)
      pvp - subtotalCL,               // P: DESCUENTO (si lo hay, congelado)
      ahora                           // Q: FECHA_FORMAL (timestamp del traspaso atómico)
    ]);
    
    // D.5) INSERTAR PRODUCTO EN TABLA PRODUCTOS para que RTV vea qué se cotiza
    var hProd = _hoja(TAB_PRODUCTOS);
    var nombreProducto = sku + ' · ' + categoria;  // Ej: "DJI Agras T50 · Aeronave Premium"
    hProd.appendRow([
      idPedido,           // A: ID_PED (referencia al pedido)
      1,                  // B: CANTIDAD (siempre 1 equipamiento)
      nombreProducto,     // C: NOMBRE (descriptivo para RTV)
      sku                 // D: SKU (código interno)
    ]);
    
    // E) ACTUALIZAR ESTADO DE COTIZACIÓN a 'CONVERTIDO_EN_PEDIDO'
    // Utilizar setValues para actualizar la fila específica
    hCot.getRange(cotizRow + 1, CCOT.ESTADO + 1).setValue(OMW_STAGES.CONVERTIDO_EN_PEDIDO);
    
    // F) FLUSH: Forzar propagación de cambios al servidor
    SpreadsheetApp.flush();
    
    return {
      exito: true,
      idPedido: idPedido,
      mensaje: 'Cotización formalizada exitosamente. Pedido ' + idPedido + ' creado en pipeline de RTV.',
      codigoError: null
    };
    
  } catch(e) {
    Logger.log('OMW_formalizarPedido ERROR: ' + e.toString());
    return {
      exito: false,
      mensaje: 'Error al procesar formalización: ' + e.message,
      codigoError: 'SYSTEM_ERROR'
    };
  } finally {
    // Siempre liberar el lock al finalizar
    if (lockAdquirido) {
      lock.releaseLock();
    }
  }
}

/* ════════════════════════════════════════════
   DIRECTIVA 4: Flujo de Retorno y Re-envío Sólido
════════════════════════════════════════════ */

// Función CRÍTICA: Permite que un Reseller corrija y re-envíe un pedido rechazado.
// NO genera una fila nueva; actualiza la fila existente en la tabla PEDIDOS.
// Limpiar observaciones previas del RTV para evitar confusión.
function OMW_reenviarPedidoRechazado(pedidoId, datosCorregidos, resellerNombre, callerEmail) {
  var lock = LockService.getScriptLock();
  var lockAdquirido = false;
  
  try {
    // A) CONTROL DE CONCURRENCIA: Adquirir lock
    lockAdquirido = lock.tryLock(15000);
    if (!lockAdquirido) {
      return {
        exito: false,
        mensaje: 'Sistema ocupado. Intenta nuevamente en 10 segundos.',
        codigoError: 'LOCK_TIMEOUT'
      };
    }
    
    // B) ENCONTRAR PEDIDO EXISTENTE
    var hPed = _hoja(TAB_PEDIDOS);
    var dPed = hPed.getDataRange().getValues();
    var pedidoRow = -1;
    var pedidoData = null;
    
    for (var i = 1; i < dPed.length; i++) {
      if (String(dPed[i][CP.ID]) === String(pedidoId)) {
        pedidoRow = i;
        pedidoData = dPed[i];
        break;
      }
    }
    
    if (pedidoRow === -1) {
      return {
        exito: false,
        mensaje: 'Pedido no encontrado: ' + pedidoId,
        codigoError: 'PEDIDO_NOT_FOUND'
      };
    }
    
    // C) VERIFICAR QUE EL PEDIDO ESTÉ EN ESTADO RECHAZADO
    var stageActual = String(pedidoData[CP.STAGE] || '');
    if (stageActual !== OMW_STAGES.RECHAZADO_RTV) {
      return {
        exito: false,
        mensaje: 'Este pedido no está en estado rechazado. Stage actual: ' + stageActual,
        codigoError: 'INVALID_STAGE'
      };
    }
    
    // D) VERIFICAR PROPIEDAD: El reseller que re-envía debe ser el dueño del pedido
    var resellerPedido = String(pedidoData[CP.RESELLER] || '').trim().toLowerCase();
    var resellerReq    = String(resellerNombre || '').trim().toLowerCase();
    if (resellerPedido !== resellerReq) {
      return {
        exito: false,
        mensaje: 'No tienes permisos para re-enviar este pedido.',
        codigoError: 'UNAUTHORIZED'
      };
    }
    
    // E) ACTUALIZAR VALORES CORREGIDOS EN LA MISMA FILA (NO crear nueva)
    // Pueden ser: importe, metodoPago, datos del cliente, notas, etc.
    var rowNum = pedidoRow + 1;
    
    if (datosCorregidos.importe !== undefined && datosCorregidos.importe !== null) {
      hPed.getRange(rowNum, CP.IMPORTE + 1).setValue(parseFloat(datosCorregidos.importe));
    }
    
    if (datosCorregidos.pago !== undefined && datosCorregidos.pago !== null) {
      hPed.getRange(rowNum, CP.PAGO + 1).setValue(String(datosCorregidos.pago));
    }
    
    if (datosCorregidos.cliente !== undefined && datosCorregidos.cliente !== null) {
      hPed.getRange(rowNum, CP.CLIENTE + 1).setValue(String(datosCorregidos.cliente));
    }
    
    if (datosCorregidos.alerta !== undefined && datosCorregidos.alerta !== null) {
      hPed.getRange(rowNum, CP.ALERTA + 1).setValue(String(datosCorregidos.alerta));
    }
    
    // F) LIMPIAR OBSERVACIONES PREVIAS DEL RTV (Col J)
    // Esto evita que observaciones de rechazo antiguas confundan al equipo
    hPed.getRange(rowNum, CP.OBS_RTV + 1).setValue('');
    
    // G) CAMBIAR STAGE DE VUELTA A '02_PENDIENTE_RTV' (re-entrando al pipeline)
    hPed.getRange(rowNum, CP.STAGE + 1).setValue(OMW_STAGES.PENDIENTE_RTV);
    
    // H) ACTUALIZAR FECHA DE MODIFICACIÓN
    hPed.getRange(rowNum, CP.FECHA_MOD + 1).setValue(new Date());
    
    // I) FLUSH: Forzar propagación
    SpreadsheetApp.flush();
    
    return {
      exito: true,
      idPedido: pedidoId,
      mensaje: 'Pedido corregido y re-enviado a RTV exitosamente.',
      codigoError: null
    };
    
  } catch(e) {
    Logger.log('OMW_reenviarPedidoRechazado ERROR: ' + e.toString());
    return {
      exito: false,
      mensaje: 'Error al re-enviar pedido: ' + e.message,
      codigoError: 'SYSTEM_ERROR'
    };
  } finally {
    if (lockAdquirido) {
      lock.releaseLock();
    }
  }
}

function OMW_crearPedido(datos, callerEmail) {
  _validarCaller(callerEmail, ['rtv', 'admin']);
  if (!datos.reseller || !datos.cliente || !datos.importe) throw new Error('Faltan datos obligatorios');
  var seq    = _nextIdCorrelativo();
  var id     = 'AG-DRON-' + String(seq).padStart(5, '0');
  var numero = id;  // Mismo que ID
  _hoja(TAB_PEDIDOS).appendRow([
    id, numero, String(datos.reseller), String(datos.cliente),
    parseFloat(datos.importe) || 0, String(datos.pago || ''),
    OMW_STAGES.PENDIENTE_RTV, new Date(), String(datos.alerta || ''),
    '', '', new Date(), new Date(),
    String(datos.idCotiz || ''),  // ID_COTIZ para trazabilidad
    0, 0, new Date()  // SUBTOTAL, DESCUENTO, FECHA_FORMAL
  ]);
  var prods = datos.prods || [];
  var hProd = _hoja(TAB_PRODUCTOS);
  for (var i = 0; i < prods.length; i++) {
    if (!String(prods[i].nombre || '').trim()) continue;
    hProd.appendRow([id, parseInt(prods[i].cant) || 1, String(prods[i].nombre), String(prods[i].sku || '')]);
  }
  return { exito: true, id: id, numero: numero, mensaje: 'Pedido creado exitosamente' };
}

/* ════════════════════════════════════════════
   COTIZADOR — lee "Condiciones ExpoAgro 2026"
════════════════════════════════════════════ */

function _cotizSS() { return SpreadsheetApp.openById(COTIZADOR_SS_ID); }

// Devuelve { rows, start } donde start = índice de primera fila de datos
function _condData() {
  var h = _cotizSS().getSheetByName(TAB_CONDICIONES);
  if (!h) throw new Error('Tab "' + TAB_CONDICIONES + '" no encontrada en Cotizador');
  var rows = h.getDataRange().getValues();
  var start = 1;
  for (var i = 0; i < Math.min(rows.length, 20); i++) {
    var c = String(rows[i][CC.CAT] || '').toLowerCase().replace(/[áa]/g,'a').trim();
    if (c === 'categoria') { start = i + 1; break; }
  }
  return { rows: rows, start: start };
}

function _parseNum(v) {
  if (typeof v === 'number') return v;
  return parseFloat(String(v || '').replace(/[^0-9.,-]/g, '').replace(',', '.')) || 0;
}

// Formatea decimal de GAS (0.105) o número entero (52000) como "10.5%" o "USD 52,000"
function _fmtPct(v) {
  if (typeof v === 'number' && Math.abs(v) <= 2) return (v * 100).toFixed(2).replace(/\.?0+$/, '') + '%';
  return String(v || '').trim();
}

function _fmtUSD(v) {
  var n = _parseNum(v);
  return 'USD ' + n.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

function _isNoDisp(v) {
  return String(v || '').toLowerCase().indexOf('disp') !== -1 && String(v || '').toLowerCase().indexOf('no') !== -1;
}

// Retorna array de { sku, cat, metodos[] } — un ítem por SKU único con sus métodos disponibles.
// Usado para poblar los selectores del cotizador en el frontend.
function OMW_getCatalogo() {
  try {
    var data = _condData();
    var rows = data.rows, s = data.start;
    var map = {};
    for (var i = s; i < rows.length; i++) {
      var sku    = String(rows[i][CC.SKU]    || '').trim();
      var cat    = String(rows[i][CC.CAT]    || '').trim();
      var metodo = String(rows[i][CC.METODO] || '').trim();
      if (!sku || !cat || !metodo) continue;
      if (!map[sku]) map[sku] = { cat: cat, metodos: [] };
      var totalVal = rows[i][CC.TOTAL];
      if (!_isNoDisp(totalVal)) map[sku].metodos.push(metodo);
    }
    return Object.keys(map).sort().map(function(sku) {
      return { sku: sku, cat: map[sku].cat, metodos: map[sku].metodos };
    });
  } catch(e) {
    Logger.log('OMW_getCatalogo error: ' + e);
    return [];
  }
}

// Retorna TODA la tabla de condiciones en un solo round-trip.
// El cliente cachea esto al abrir el cotizador por primera vez y hace búsquedas locales.
// Incluye precios pre-formateados para evitar cálculos client-side.
function OMW_getTablaCondiciones() {
  try {
    var data = _condData();
    var rows = data.rows, s = data.start;
    var result = [];
    for (var i = s; i < rows.length; i++) {
      var sku    = String(rows[i][CC.SKU]    || '').trim();
      var cat    = String(rows[i][CC.CAT]    || '').trim();
      var metodo = String(rows[i][CC.METODO] || '').trim();
      if (!sku || !cat || !metodo) continue;
      if (_isNoDisp(rows[i][CC.TOTAL])) continue;
      result.push({
        sku:         sku,
        cat:         cat,
        metodo:      metodo,
        pvp:         _parseNum(rows[i][CC.PVP]),
        pvpStr:      _fmtUSD(rows[i][CC.PVP]),
        ivaStr:      _fmtPct(rows[i][CC.IVA]),
        comLlenaStr: _fmtPct(rows[i][CC.COM_LL]),
        descCompStr: _fmtPct(rows[i][CC.DESC_CO]),
        voucherStr:  _fmtPct(rows[i][CC.VOUCHER]),
        cftStr:      _fmtPct(rows[i][CC.CFT]),
        descResStr:  _fmtPct(rows[i][CC.DESC_RE]),
        pctComStr:   _fmtPct(rows[i][CC.PCT_COM]),
        comision:    _parseNum(rows[i][CC.COMISION]),
        comisionStr: _fmtUSD(rows[i][CC.COMISION]),
        subtotalCL:  _parseNum(rows[i][CC.SUB_CL]),
        subtotalStr: _fmtUSD(rows[i][CC.SUB_CL]),
        ivaCL:       _parseNum(rows[i][CC.IVA_CL]),
        ivaCLStr:    _fmtUSD(rows[i][CC.IVA_CL]),
        totalCL:     _parseNum(rows[i][CC.TOTAL]),
        totalStr:    _fmtUSD(rows[i][CC.TOTAL])
      });
    }
    return result;
  } catch(e) {
    Logger.log('OMW_getTablaCondiciones error: ' + e);
    return [];
  }
}

// Retorna el breakdown completo de precios para un SKU + método de pago específico.
function OMW_getCotizacion(sku, metodoPago) {
  try {
    var data = _condData();
    var rows = data.rows, s = data.start;
    var skuN   = String(sku        || '').trim().toUpperCase();
    var metN   = String(metodoPago || '').trim().toLowerCase();
    for (var i = s; i < rows.length; i++) {
      if (String(rows[i][CC.SKU]    || '').trim().toUpperCase() !== skuN) continue;
      if (String(rows[i][CC.METODO] || '').trim().toLowerCase() !== metN) continue;
      var totalRaw = rows[i][CC.TOTAL];
      if (_isNoDisp(totalRaw)) return { ok: false, error: 'Condición no disponible para esta combinación' };
      return {
        ok:           true,
        categoria:    String(rows[i][CC.CAT]     || ''),
        sku:          String(rows[i][CC.SKU]     || ''),
        metodoPago:   String(rows[i][CC.METODO]  || ''),
        pvp:          _parseNum(rows[i][CC.PVP]),
        pvpStr:       _fmtUSD(rows[i][CC.PVP]),
        ivaStr:       _fmtPct(rows[i][CC.IVA]),
        comLlenaStr:  _fmtPct(rows[i][CC.COM_LL]),
        descCompStr:  _fmtPct(rows[i][CC.DESC_CO]),
        voucherStr:   _fmtPct(rows[i][CC.VOUCHER]),
        cftStr:       _fmtPct(rows[i][CC.CFT]),
        descResStr:   _fmtPct(rows[i][CC.DESC_RE]),
        pctComStr:    _fmtPct(rows[i][CC.PCT_COM]),
        comision:     _parseNum(rows[i][CC.COMISION]),
        comisionStr:  _fmtUSD(rows[i][CC.COMISION]),
        subtotalCL:   _parseNum(rows[i][CC.SUB_CL]),
        subtotalStr:  _fmtUSD(rows[i][CC.SUB_CL]),
        ivaCL:        _parseNum(rows[i][CC.IVA_CL]),
        ivaCLStr:     _fmtUSD(rows[i][CC.IVA_CL]),
        totalCL:      _parseNum(rows[i][CC.TOTAL]),
        totalStr:     _fmtUSD(rows[i][CC.TOTAL])
      };
    }
    return { ok: false, error: 'Combinación no encontrada en el catálogo' };
  } catch(e) {
    Logger.log('OMW_getCotizacion error: ' + e);
    return { ok: false, error: 'Error al leer cotizador: ' + e.message };
  }
}

/* ════════════════════════════════════════════
   COTIZACIONES — guarda / consulta cotizaciones de resellers
════════════════════════════════════════════ */

function _hojaCotiz() {
  var h = _ss().getSheetByName(TAB_COTIZACIONES);
  if (!h) {
    h = _ss().insertSheet(TAB_COTIZACIONES);
    var hdrs = ['ID','Reseller','Nombre_Cliente','Empresa_Cliente','CUIT','Email_Cliente',
                'Telefono','SKU','Categoria','Metodo_Pago','PVP_USD','Subtotal_CL_USD',
                'IVA_CL_USD','Total_CL_USD','Estado','Fecha_Creacion','Notas'];
    h.getRange(1, 1, 1, hdrs.length).setValues([hdrs])
     .setFontWeight('bold').setBackground('#1A2438').setFontColor('#E2E8F0');
    h.setFrozenRows(1);
    h.setColumnWidths(1, hdrs.length, 140);
  }
  return h;
}

function _nextCotizId() {
  var d = _hojaCotiz().getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < d.length; i++) {
    var idStr = String(d[i][CCOT.ID] || '').trim();
    if (!idStr || idStr.indexOf('COT-2026-') === -1) continue;
    var numPart = idStr.substring('COT-2026-'.length);
    var v = parseInt(numPart);
    if (!isNaN(v) && v > max) max = v;
  }
  var newSeq = max + 1;
  return 'COT-2026-' + String(newSeq).padStart(5, '0');
}

// Guarda una cotización con datos del cliente. Retorna { exito, id, mensaje }.
// Genera ID único en formato 'COT-2026-XXXXX'. El estado inicial es 'NUEVA'.
// La cotización se guarda EXCLUSIVAMENTE en la pestaña 'COTIZACIONES' y NO aparecerá en el Kanban de pedidos.
function OMW_guardarCotizacion(datos) {
  try {
    if (!datos || !datos.reseller || !datos.sku) throw new Error('Datos incompletos');
    var id = _nextCotizId();
    _hojaCotiz().appendRow([
      id,  // ID en formato COT-2026-XXXXX
      String(datos.reseller    || ''),
      String(datos.nombreCli   || ''),
      String(datos.empresaCli  || ''),
      String(datos.cuit        || ''),
      String(datos.emailCli    || ''),
      String(datos.tel         || ''),
      String(datos.sku         || ''),
      String(datos.categoria   || ''),
      String(datos.metodoPago  || ''),
      datos.pvp        || 0,
      datos.subtotalCL || 0,
      datos.ivaCL      || 0,
      datos.totalCL    || 0,
      OMW_STAGES.NUEVA_COTIZACION,  // Estado inicial: '01_NUEVA_COTIZACION'
      new Date(),
      String(datos.notas || '')
    ]);
    return { exito: true, id: id, mensaje: 'Cotización guardada exitosamente' };
  } catch(e) {
    Logger.log('OMW_guardarCotizacion: ' + e);
    return { exito: false, mensaje: e.message };
  }
}

// Retorna cotizaciones de un reseller específico (por nombre exacto), ordenadas newest-first.
// Las cotizaciones son borradores READ-ONLY hasta ser formalizadas a pedidos.
function OMW_getCotizaciones(resellerNombre) {
  try {
    var h = _hojaCotiz();
    var d = h.getDataRange().getValues();
    var nameB = String(resellerNombre || '').trim().toLowerCase();
    var list = [];
    for (var i = 1; i < d.length; i++) {
      if (!d[i][CCOT.ID]) continue;
      var rowRes = String(d[i][CCOT.RESELLER] || '').trim().toLowerCase();
      if (nameB && rowRes !== nameB) continue;
      var tot = parseFloat(d[i][CCOT.TOTAL]) || 0;
      var estado = String(d[i][CCOT.ESTADO] || OMW_STAGES.NUEVA_COTIZACION);
      list.push({
        id:         String(d[i][CCOT.ID]),
        reseller:   String(d[i][CCOT.RESELLER]    || ''),
        nombreCli:  String(d[i][CCOT.NOMBRE_CLI]  || ''),
        empresaCli: String(d[i][CCOT.EMPRESA_CLI] || ''),
        cuit:       String(d[i][CCOT.CUIT]        || ''),
        emailCli:   String(d[i][CCOT.EMAIL_CLI]   || ''),
        tel:        String(d[i][CCOT.TEL]         || ''),
        sku:        String(d[i][CCOT.SKU]         || ''),
        categoria:  String(d[i][CCOT.CAT]         || ''),
        metodoPago: String(d[i][CCOT.METODO]      || ''),
        pvp:        parseFloat(d[i][CCOT.PVP])      || 0,
        subtotalCL: parseFloat(d[i][CCOT.SUBTOTAL]) || 0,
        ivaCL:      parseFloat(d[i][CCOT.IVA_CL])  || 0,
        totalCL:    tot,
        totalStr:   'USD ' + tot.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 }),
        estado:     estado,
        fecha:      _fmtTs(d[i][CCOT.FECHA_CREA]),
        notas:      String(d[i][CCOT.NOTAS]      || ''),
        puedeFormalizar: (estado !== OMW_STAGES.CONVERTIDO_EN_PEDIDO)
      });
    }
    list.reverse(); // newest first
    return list;
  } catch(e) {
    Logger.log('OMW_getCotizaciones: ' + e);
    return [];
  }
}

// Actualiza estado de una cotización. Solo el reseller dueño puede hacerlo.
// Devuelve {exito, mensaje} según formato Enterprise.
function OMW_actualizarEstadoCotiz(cotizId, nuevoEstado, resellerNombre) {
  try {
    var estadosValidos = ['Nueva','Seguimiento','Cerrada'];
    if (estadosValidos.indexOf(nuevoEstado) === -1) throw new Error('Estado inválido');
    var h = _hojaCotiz();
    var d = h.getDataRange().getValues();
    var nameB = String(resellerNombre || '').trim().toLowerCase();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][CCOT.ID]) !== String(cotizId)) continue;
      var rowRes = String(d[i][CCOT.RESELLER] || '').trim().toLowerCase();
      if (rowRes !== nameB) return { exito: false, mensaje: 'No autorizado' };
      h.getRange(i + 1, CCOT.ESTADO + 1).setValue(nuevoEstado);
      return { exito: true, mensaje: 'Estado actualizado' };
    }
    return { exito: false, mensaje: 'Cotización no encontrada' };
  } catch(e) {
    Logger.log('OMW_actualizarEstadoCotiz: ' + e);
    return { exito: false, mensaje: e.message };
  }
}

/* ════════════════════════════════════════════
   SETUP — ejecutar UNA VEZ desde el editor GAS
   Crea las tabs que falten en el Sheet existente.
════════════════════════════════════════════ */
function setupOMW() {
  var ss  = _ss();
  var log = [];

  function _ensureTab(name, headers, colW) {
    var h = ss.getSheetByName(name);
    if (!h) {
      h = ss.insertSheet(name);
      h.getRange(1, 1, 1, headers.length).setValues([headers])
        .setFontWeight('bold').setBackground('#1A2438').setFontColor('#E2E8F0');
      h.setFrozenRows(1);
      if (colW) h.setColumnWidths(1, headers.length, colW);
      log.push('✅ Tab creada: ' + name);
    } else {
      log.push('⚠️  Tab ya existe: ' + name);
    }
    return h;
  }

  // PEDIDOS: Schema expandido a 18 columnas (Enterprise)
  _ensureTab(TAB_PEDIDOS,
    ['ID','Numero','Reseller','Cliente','Importe_USD','Forma_Pago','Stage',
     'SLA_inicio','Alerta','Obs_RTV','Ref_Bancaria','Fecha_creacion','Fecha_modificacion',
     'ID_Cotizacion','Subtotal_Congelado','Descuento_Congelado','Fecha_Formal'], 160);

  _ensureTab(TAB_PRODUCTOS, ['ID_pedido','Cantidad','Nombre','SKU']);
  _ensureTab(TAB_MENSAJES,  ['ID_pedido','Rol','Texto','Timestamp','Email']);

  var hUsr = _ensureTab(TAB_USUARIOS, ['Email','Nombre','Rol','Reseller_asociado']);
  if (hUsr.getLastRow() === 1) {
    hUsr.appendRow(['victor@bidcom.com.ar', 'Victor', 'admin', '']);
    log.push('👤 Usuario inicial agregado');
  }

  // COTIZACIONES: Schema con estado para formalización
  _ensureTab(TAB_COTIZACIONES,
    ['ID','Reseller','Nombre_Cliente','Empresa_Cliente','CUIT','Email_Cliente',
     'Telefono','SKU','Categoria','Metodo_Pago','PVP_USD','Subtotal_CL_USD',
     'IVA_CL_USD','Total_CL_USD','Estado','Fecha_Creacion','Notas'], 140);

  log.forEach(function(l){ Logger.log(l); });
  return log.join('\n');
}
