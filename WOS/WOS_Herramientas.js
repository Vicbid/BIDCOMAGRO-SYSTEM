// @version 1.1
// ============================================================
//  WOS — Herramientas de recuperación/diagnóstico GENÉRICAS (no atadas
//  a un incidente puntual, idempotentes, pensadas para correr desde el
//  editor de Apps Script si el problema vuelve a pasar):
//    WOS_recuperarTracking/Test — reconstruye códigos de seguimiento pisados
//    WOS_diagnosticoCarmen      — verifica acceso/escritura a Carmen
//    WOS_recuperarThreadIds     — restaura Thread ID perdido buscando en Gmail
//  Extraído de WOS_GmailFlow.js 2.30 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================



// ─────────────────────────────────────────────────────────────
//  HITO 4b — Despachar Completo (reemplaza WOS_despacharPedido)
//  Registra cantidades, guarda tracking, responde en el hilo
//  original y notifica a facturación en un solo paso.
//
//  despachos:    [{row (1-indexed), cantDesp}]
//  transportista: string requerido ("Correo Argentino", "Credifin", "Via Cargo")
//  bultos:        [{tracking, peso}] — un objeto por bulto físico
//  costoEnvio:    costo total del envío (opcional)
//  reqToken:      token de idempotencia (opcional) — si el mismo token ya se
//                 procesó, se devuelve el resultado previo sin re-ejecutar
//                 (protege contra doble-click / reintento tras respuesta perdida)
// ─────────────────────────────────────────────────────────────
// Combina códigos de seguimiento sin pisar ni duplicar. Acepta valores ya
// combinados ("T1 | T2") en cualquiera de los dos lados. Devuelve "T1 | T2 | ...".
function _wosMergeTracking(oldVal, newVal) {
  var out = [];
  function _add(v) {
    var parts = String(v || '').split('|');
    for (var i = 0; i < parts.length; i++) {
      var t = parts[i].trim();
      if (t && out.indexOf(t) === -1) out.push(t);
    }
  }
  _add(oldVal);
  _add(newVal);
  return out.join(' | ');
}


// ─────────────────────────────────────────────────────────────
// RECUPERACIÓN de códigos de seguimiento pisados: los parsea de los mails de
// despacho del hilo del pedido (cada despacho mandó "Código de seguimiento: ...").
//   numero  : N° de pedido, ej "PR-0035".
//   guardar : true → fusiona los códigos recuperados en las filas ya despachadas
//             (col Q); false/omitido → solo devuelve el reporte (no toca la planilla).
function WOS_recuperarTracking(numero, guardar) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    numero = String(numero || '').trim();
    if (!numero) return { ok: false, error: 'Falta el número de pedido' };
    var ped = _wosLeerPedido(numero);
    if (!ped.hoja) return { ok: false, error: 'Pedido no encontrado: ' + numero };

    // Reunir los mensajes del hilo (por threadId guardado; respaldo: búsqueda por N°)
    var mensajes = [];
    try {
      if (ped.threadId) {
        var th = GmailApp.getThreadById(ped.threadId);
        if (th) mensajes = th.getMessages();
      }
    } catch(eT) { Logger.log('WOS_recuperarTracking thread: ' + eT); }
    if (!mensajes.length) {
      try {
        var hilos = GmailApp.search('subject:"' + numero + '"', 0, 10);
        for (var h = 0; h < hilos.length; h++) {
          var ms = hilos[h].getMessages();
          for (var mm = 0; mm < ms.length; mm++) mensajes.push(ms[mm]);
        }
      } catch(eS) { Logger.log('WOS_recuperarTracking search: ' + eS); }
    }

    // Parsear "Código de seguimiento: ..." de cada mensaje (uno por despacho)
    var re = /C[o\xf3]digo de seguimiento:\s*([^\n\r]+)/i;
    var detalle = [], todos = [];
    for (var i = 0; i < mensajes.length; i++) {
      var body = '';
      try { body = mensajes[i].getPlainBody() || ''; } catch(eB) { continue; }
      var mtch = body.match(re);
      if (!mtch) continue;
      var partes = mtch[1].split(/\s*[|,]\s*/);
      var limpios = [];
      for (var c = 0; c < partes.length; c++) {
        var t = String(partes[c] || '').trim();
        if (!t) continue;
        if (limpios.indexOf(t) === -1) limpios.push(t);
        if (todos.indexOf(t)   === -1) todos.push(t);
      }
      if (!limpios.length) continue;
      var fecha = '';
      try { fecha = Utilities.formatDate(mensajes[i].getDate(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); } catch(eD) {}
      detalle.push({ fecha: fecha, codigos: limpios });
    }

    var res = { ok: true, numero: numero, codigos: todos, detalle: detalle, mensajesRevisados: mensajes.length };

    // Escritura opcional: fusiona TODOS los códigos recuperados en las filas ya despachadas
    if (guardar && todos.length) {
      var escritas = 0;
      var mergeStr = todos.join(' | ');
      for (var r = 1; r < ped.datos.length; r++) {
        if (String(ped.datos[r][COL.NUMERO] || '').trim() !== numero) continue;
        var yaDesp = ped.datos[r][COL.FECHA_DESPACHO] || String(ped.datos[r][COL.TRACKING] || '').trim();
        if (!yaDesp) continue; // solo filas efectivamente despachadas
        var mergedRow = _antiFormula(_wosMergeTracking(ped.datos[r][COL.TRACKING], mergeStr));
        ped.hoja.getRange(r + 1, COL.TRACKING + 1).setValue(mergedRow);
        escritas++;
      }
      res.filasActualizadas = escritas;
    }

    Logger.log('WOS_recuperarTracking ' + numero + ': ' + todos.length + ' código(s) → ' + JSON.stringify(todos));
    return res;
  } catch(e) {
    Logger.log('WOS_recuperarTracking: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Wrapper para correr DESDE EL EDITOR (no acepta parámetros ahí):
// editá NUMERO y GUARDAR, ejecutá, y mirá el resultado en el Log (Ver → Registros).
//   1ª pasada con GUARDAR=false para revisar los códigos encontrados.
//   2ª pasada con GUARDAR=true para escribirlos en la planilla.
function WOS_recuperarTrackingTest() {
  var NUMERO  = 'PR-0000';   // ← poné acá el N° del pedido
  var GUARDAR = false;       // ← poné true cuando quieras escribirlos
  var r = WOS_recuperarTracking(NUMERO, GUARDAR);
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}


// Diagnóstico de acceso a Carmen para el descuento de stock del despacho.
// Ejecutar desde el editor de Apps Script de WOS (Run ▶) y ver el resultado.
// Dice si abre la planilla, lista las pestañas (para ver si "Entregados" existe / cómo se
// llama) y prueba escribir+borrar una fila de test en "Entregados".
function WOS_diagnosticoCarmen() {
  var out = { carmenId: CARMEN_SS_ID, abrio: false, ok: false };
  var u = WOS_getUsuario();
  if (!u.autorizado) { out.error = 'No autorizado.'; return out; }
  try {
    var ss = SpreadsheetApp.openById(CARMEN_SS_ID);
    out.abrio = true;
    out.nombrePlanilla = ss.getName();
    var tabs = ss.getSheets().map(function(s) { return s.getName(); });
    out.pestanas = tabs;
    out.tieneEntregados  = tabs.indexOf('Entregados') !== -1;
    out.tieneUbicaciones = tabs.indexOf(CARMEN_UBICACIONES_TAB) !== -1;
    var hoja = ss.getSheetByName('Entregados');
    if (hoja) {
      hoja.appendRow(['__TEST_WOS__', 'diagnostico (borrar)', 0, 'TEST', '', '', new Date(), '']);
      SpreadsheetApp.flush();
      hoja.deleteRow(hoja.getLastRow());
      SpreadsheetApp.flush();
      out.escrituraOK = true;
    }
    out.ok = out.tieneEntregados && out.escrituraOK === true;
  } catch(e) {
    out.error = e.toString();
  }
  Logger.log(JSON.stringify(out, null, 2));
  return out;
}


// ─────────────────────────────────────────────────────────────
//  UTILIDAD — Recuperar Thread IDs perdidos
//  Ejecutar UNA VEZ desde el editor de Apps Script.
//  Para cada pedido sin Thread ID en col R, busca en Gmail
//  el hilo original por número de pedido y lo restaura.
//  Ver resultado en Ver → Registros de ejecución.
// ─────────────────────────────────────────────────────────────
function WOS_recuperarThreadIds() {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.', recuperados: 0, noEncontrados: [] };
  var hojas = [_wosHoja(), _getHojaPedidosOT()].filter(Boolean);

  // Agrupar filas por pedido, solo los que no tienen threadId; registra la hoja de origen
  var sinThread = {}; // numero → { hoja, filas: [] }
  for (var h = 0; h < hojas.length; h++) {
    var hoja  = hojas[h];
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      var num = String(datos[i][COL.NUMERO]    || '').trim();
      var tid = String(datos[i][COL.THREAD_ID] || '').trim();
      if (!num || tid) continue;
      if (!sinThread[num]) sinThread[num] = { hoja: hoja, filas: [] };
      sinThread[num].filas.push(i + 1); // filas 1-indexed
    }
  }

  var numeros = Object.keys(sinThread);
  Logger.log('WOS_recuperarThreadIds: ' + numeros.length + ' pedidos sin Thread ID → ' + numeros.join(', '));
  if (!numeros.length) return { ok: true, recuperados: 0, noEncontrados: [] };

  var recuperados = 0;
  var noEncontrados = [];

  for (var n = 0; n < numeros.length; n++) {
    var numero = numeros[n];
    var hoja   = sinThread[numero].hoja;
    var rows   = sinThread[numero].filas;

    // Buscar en Gmail: el Portal usa asunto "[PEDIDO] PR-XXXXX — ..."
    // Un solo número es suficiente porque los IDs son únicos
    var threads = [];
    try {
      threads = GmailApp.search('subject:"' + numero + '"', 0, 10);
    } catch(eS) {
      Logger.log('WOS_recuperarThreadIds: error buscando ' + numero + ' → ' + eS);
      noEncontrados.push(numero);
      continue;
    }

    if (!threads || threads.length === 0) {
      Logger.log('WOS_recuperarThreadIds: ' + numero + ' → sin resultados en Gmail');
      noEncontrados.push(numero);
      continue;
    }

    // Si hay más de un hilo (raro), usar el más antiguo (el del pedido original)
    var threadId = threads[threads.length - 1].getId();

    for (var r = 0; r < rows.length; r++) {
      hoja.getRange(rows[r], COL.THREAD_ID + 1).setValue(threadId);
    }
    recuperados++;
    Logger.log('WOS_recuperarThreadIds: ' + numero + ' (' + rows.length + ' fila/s) → ' + threadId);
  }

  SpreadsheetApp.flush();
  Logger.log('WOS_recuperarThreadIds RESULTADO: recuperados=' + recuperados +
             ' | no encontrados=' + (noEncontrados.length ? noEncontrados.join(', ') : 'ninguno'));
  return { ok: true, recuperados: recuperados, noEncontrados: noEncontrados };
}
