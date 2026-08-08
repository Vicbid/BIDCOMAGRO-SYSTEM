// @version 1.8
var MASTER_SS_ID = '1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc';
var NOTAS_SS_ID  = '1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw';
var CARMEN_SS_ID = '1-BH5m-LXFYhBZxqpSFVhIz5jwzFgJmLWH8Qvkh4PSCI'; // stock en vivo (tab 'STOCK'), para LAUNCH_recuperarPedidos

// ── ComandasPedidos · _CONFIG (se edita desde el Launcher) ──
// Planilla de LOG de ComandasPedidos donde vive la pestaña _CONFIG (clave/valor).
// El scope 'spreadsheets' del Launcher ya permite leerla/escribirla por ID (sin re-auth).
var CP_LOG_SS_ID  = '1mOOeUDPORa9d1csQJ1fCL4ON592SvjMr1y3VKmWBN44';
var CP_CONFIG_TAB = '_CONFIG';

// Estado de la cuenta del reseller: col Q de la hoja Resellers (índice 16, columna 1-based 17).
// Vacío = activo. NO / BAJA / INACTIVO / FALSE / 0 / false = desactivado. Debe coincidir con RS_Auth.js _resellerInactivo.
var RES_ACTIVO_COL = 17; // 1-based para getRange/setValue
function _launchResellerInactivo(fila) {
  var raw = fila[16];
  if (raw === false) return true;
  var v = String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
  if (!v) return false;
  return (v === 'NO' || v === 'BAJA' || v === 'INACTIVO' || v === 'INACTIVA'
          || v === 'FALSE' || v === '0' || v === 'DESACTIVADO' || v === 'DESACTIVADA');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Launcher_Index')
    .setTitle('BIDCOMAGRO · Sistema Central')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function obtenerKPIs() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var ssN  = SpreadsheetApp.openById(NOTAS_SS_ID);
    var tz   = Session.getScriptTimeZone();
    var hoy  = new Date();
    var primerDiaMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
    var TERMINAL = ['Finalizado', 'Entregado', 'CANCELADO', 'Rechazado DJI', 'Sin respuesta · Cerrado'];

    // ── HUB PRO ──────────────────────────────────────────────────
    var hub = { abiertas: 0, urgentes: 0, sinTecnico: 0, finMes: 0 };
    var dOT = ss.getSheetByName('Ordenes de trabajo');
    if (dOT) {
      var rOT = dOT.getDataRange().getValues();
      for (var i = 1; i < rOT.length; i++) {
        var est  = String(rOT[i][4]  || '').trim();
        var prio = String(rOT[i][17] || '').trim();
        var tec  = String(rOT[i][9]  || '').trim();
        var fCierre = rOT[i][1];
        if (TERMINAL.indexOf(est) === -1) {
          hub.abiertas++;
          if (prio === 'URGENTE') hub.urgentes++;
          if (!tec || tec === 'Sin asignar') hub.sinTecnico++;
        }
        if (est === 'Finalizado' && fCierre instanceof Date && fCierre >= primerDiaMes) hub.finMes++;
      }
    }

    // ── STOCK MANAGER -───────────────────────────────────────────
    var stock = { total: 0, critico: 0, bajo: 0, despachos: 0 };
    var dSt = ss.getSheetByName('STOCK_REPUESTOS');
    if (dSt) {
      var rSt = dSt.getDataRange().getValues();
      for (var s = 1; s < rSt.length; s++) {
        if (!String(rSt[s][0] || '').trim()) continue;
        stock.total++;
        var act = parseInt(rSt[s][2]) || 0;
        var min = parseInt(rSt[s][3]) || 0;
        if (act <= 0 || (min > 0 && act < min * 0.5)) stock.critico++;
        else if (min > 0 && act < min)                  stock.bajo++;
      }
    }
    var dSol = ss.getSheetByName('SOLICITUDES_DESPACHO');
    if (dSol) {
      var rSol = dSol.getDataRange().getValues();
      for (var d = 1; d < rSol.length; d++) {
        if (String(rSol[d][8] || '') === 'Pendiente') stock.despachos++;
      }
    }

    // ── PORTAL RESELLER ──────────────────────────────────────────
    var portal = { resellers: 0, otMes: 0 };
    var dRes = ss.getSheetByName('Resellers');
    if (dRes) portal.resellers = Math.max(0, dRes.getLastRow() - 1);
    if (dOT) {
      var rOT2 = dOT.getDataRange().getValues();
      for (var p = 1; p < rOT2.length; p++) {
        var fIng = rOT2[p][0];
        var res  = String(rOT2[p][7] || '').trim();
        if (fIng instanceof Date && fIng >= primerDiaMes && res) portal.otMes++;
      }
    }

    // ── WOS ──────────────────────────────────────────────────────
    var wos = { activos: 0, aDespachar: 0, enEspera: 0, despMes: 0 };
    var WOS_TERMINAL   = ['Cancelado', 'Entregado_Cerrado', 'Entregado_Confirmado'];
    var WOS_ADESPACHAR = ['Preparado', 'Preparado Parcial', 'Listo_Retiro'];
    var dPed = ssN.getSheetByName('Pedidos_resellers');
    if (dPed) {
      var rPed = dPed.getDataRange().getValues();
      for (var w = 1; w < rPed.length; w++) {
        if (!String(rPed[w][0] || '').trim()) continue;
        var wEst  = String(rPed[w][9]  || '').trim();
        var wFDesp = rPed[w][14];
        if (WOS_TERMINAL.indexOf(wEst) === -1) {
          wos.activos++;
          if (WOS_ADESPACHAR.indexOf(wEst) !== -1) wos.aDespachar++;
          if (wEst === 'En_Espera_Reseller')        wos.enEspera++;
        }
        if ((wEst === 'Entregado_Cerrado' || wEst === 'Entregado_Confirmado') &&
            wFDesp instanceof Date && wFDesp >= primerDiaMes) {
          wos.despMes++;
        }
      }
    }

    return {
      ok: true,
      hub:    hub,
      stock:  stock,
      portal: portal,
      wos:    wos,
      urls:   obtenerUrls(),
      ts:     Utilities.formatDate(hoy, tz, 'dd/MM/yyyy HH:mm')
    };
  } catch(e) {
    Logger.log('obtenerKPIs: ' + e);
    return { ok: false, error: e.message };
  }
}

function obtenerUrls() {
  var p = PropertiesService.getScriptProperties();
  return {
    hub:    p.getProperty('URL_HUB')    || '',
    stock:  p.getProperty('URL_STOCK')  || '',
    portal: p.getProperty('URL_PORTAL') || '',
    wos:    p.getProperty('URL_WOS')    || ''
  };
}

function guardarUrls(hubUrl, stockUrl, portalUrl, wosUrl) {
  try {
    var p = PropertiesService.getScriptProperties();
    p.setProperty('URL_HUB',    String(hubUrl    || ''));
    p.setProperty('URL_STOCK',  String(stockUrl  || ''));
    p.setProperty('URL_PORTAL', String(portalUrl || ''));
    p.setProperty('URL_WOS',    String(wosUrl    || ''));
    return { ok: true };
  } catch(e) {
    Logger.log('guardarUrls: ' + e);
    return { ok: false, error: e.message };
  }
}

// ─────────────────────────────────────────────────────────────
//  HERRAMIENTAS DE MANTENIMIENTO
// ─────────────────────────────────────────────────────────────

function ejecutarHerramienta(id) {
  try {
    if (id === 'wos-recuperar-threads') return _herr_wosRecuperarThreadIds();
    if (id === 'wos-cuota-gmail')       return _herr_cuotaGmail();
    if (id === 'sys-health')            return _herr_sysHealth();
    return { ok: false, error: 'Herramienta desconocida: ' + id };
  } catch(e) {
    Logger.log('ejecutarHerramienta [' + id + ']: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── WOS: Recuperar Thread IDs perdidos ───────────────────────
function _herr_wosRecuperarThreadIds() {
  var hoja  = SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName('Pedidos_resellers');
  if (!hoja) return { ok: false, error: 'Hoja Pedidos_resellers no encontrada.' };
  var datos = hoja.getDataRange().getValues();

  var sinThread = {};
  for (var i = 1; i < datos.length; i++) {
    var num = String(datos[i][0]  || '').trim(); // col A: NUMERO
    var tid = String(datos[i][17] || '').trim(); // col R: THREAD_ID
    if (!num || tid) continue;
    if (!sinThread[num]) sinThread[num] = [];
    sinThread[num].push(i + 1);
  }

  var numeros = Object.keys(sinThread);
  if (!numeros.length) {
    return { ok: true, recuperados: 0, noEncontrados: [],
             msg: 'Todos los pedidos ya tienen Thread ID. Nada que recuperar.' };
  }

  var recuperados = 0, noEncontrados = [];
  for (var n = 0; n < numeros.length; n++) {
    var numero = numeros[n];
    var rows   = sinThread[numero];
    var threads = [];
    try {
      threads = GmailApp.search('subject:"' + numero + '"', 0, 10);
    } catch(eS) {
      noEncontrados.push(numero);
      continue;
    }
    if (!threads || !threads.length) {
      noEncontrados.push(numero);
      continue;
    }
    var threadId = threads[threads.length - 1].getId();
    for (var r = 0; r < rows.length; r++) {
      hoja.getRange(rows[r], 18).setValue(threadId); // col R
    }
    recuperados++;
  }

  SpreadsheetApp.flush();
  var msg = 'Recuperados: ' + recuperados + ' de ' + numeros.length + ' pedidos sin Thread ID.';
  if (noEncontrados.length) msg += ' No encontrados en Gmail: ' + noEncontrados.join(', ') + '.';
  return { ok: true, recuperados: recuperados, total: numeros.length,
           noEncontrados: noEncontrados, msg: msg };
}

// ── Recuperar pedidos que no quedaron registrados en Pedidos_resellers ──────
// Reportado por el usuario: RS_Pedidos (Portal Reseller) dejó de escribir filas en
// Pedidos_resellers en algún momento — el pedido se creaba igual (mail + PDF al reseller,
// sin que note nada) pero la fila que WOS necesita para despachar nunca se escribía. La
// única copia completa que sobrevive es PEDIDOS_REPUESTOS (col "Items JSON"), que SIEMPRE
// se escribe (paso E de confirmarPedidoPortal, posterior al intento fallido). Esta función
// reconstruye, a pedido del usuario, las filas faltantes a partir de ese respaldo.
//
// spec acepta, mezclados y separados por coma: números sueltos ("PR-0100" o "100", se
// autocompleta a 4 dígitos) y rangos con la palabra "a" ("PR-0100 a PR-0110"). Ej:
// "PR-0050, PR-0100 a PR-0110, 200"

// Parsea el texto libre del usuario a una lista de números "PR-####" sin duplicados.
function _launchParsePedidoSpec(spec) {
  var out = [], seen = {};
  function normalizar(s) {
    s = String(s || '').trim().toUpperCase();
    if (!s) return '';
    var m = s.match(/^(?:PR-)?(\d+)$/);
    if (!m) return '';
    var n = m[1];
    while (n.length < 4) n = '0' + n;
    return 'PR-' + n;
  }
  var tokens = String(spec || '').split(',');
  for (var i = 0; i < tokens.length; i++) {
    var tok = tokens[i].trim();
    if (!tok) continue;
    var rangeM = tok.match(/^(.+?)\s+a\s+(.+)$/i);
    if (rangeM) {
      var desde = normalizar(rangeM[1]), hasta = normalizar(rangeM[2]);
      var dNum  = desde ? parseInt(desde.replace('PR-', ''), 10) : NaN;
      var hNum  = hasta ? parseInt(hasta.replace('PR-', ''), 10) : NaN;
      // Tope de 500 para que un typo en el rango no genere miles de filas por accidente.
      if (!isNaN(dNum) && !isNaN(hNum) && hNum >= dNum && (hNum - dNum) < 500) {
        for (var n2 = dNum; n2 <= hNum; n2++) {
          var num2 = String(n2); while (num2.length < 4) num2 = '0' + num2;
          var full2 = 'PR-' + num2;
          if (!seen[full2]) { seen[full2] = true; out.push(full2); }
        }
      }
      continue;
    }
    var single = normalizar(tok);
    if (single && !seen[single]) { seen[single] = true; out.push(single); }
  }
  return out;
}

function LAUNCH_recuperarPedidos(spec) {
  try {
    var numeros = _launchParsePedidoSpec(spec);
    if (!numeros.length) return { ok: false, error: 'No pude interpretar ningún número de pedido en "' + spec + '". Formato: PR-0100, PR-0102, PR-0105 a PR-0110' };

    var hojaPed = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('PEDIDOS_REPUESTOS');
    if (!hojaPed) return { ok: false, error: 'No se encontró la hoja PEDIDOS_REPUESTOS.' };
    var dPed = hojaPed.getDataRange().getValues();
    // col: 0 ID · 1 Fecha · 2 Reseller · 3 Email · 4 CantItems · 5 CantSinCatalogo ·
    //      6 Estado · 7 Observaciones · 8 Items JSON · 9 PDF URL · 10 Total USD · 11 FormaPago · 12 Envio
    var porNumero = {};
    for (var i = 1; i < dPed.length; i++) {
      var id = String(dPed[i][0] || '').trim();
      if (id && !porNumero[id]) porNumero[id] = dPed[i];
    }

    var hojaNota = SpreadsheetApp.openById(NOTAS_SS_ID).getSheetByName('Pedidos_resellers');
    if (!hojaNota) return { ok: false, error: 'No se encontró la hoja Pedidos_resellers.' };
    var dNota = hojaNota.getDataRange().getValues();

    // Pedidos que YA tienen al menos 1 fila ahí → nunca duplicar, aunque el usuario los
    // vuelva a pasar por error (rango que se superpone con otro ya recuperado, etc.)
    var yaExiste = {};
    for (var j = 1; j < dNota.length; j++) {
      var numJ = String(dNota[j][0] || '').trim();
      if (numJ) yaExiste[numJ] = true;
    }

    // Stock actual (Carmen) para decidir Confirmado/Pendiente_Revision igual que el flujo
    // normal — best-effort: si falla, todo lo recuperado queda en Pendiente_Revision (más
    // conservador, alguien lo revisa a mano antes de prepararlo).
    var stockMap = {};
    try {
      var dStock = SpreadsheetApp.openById(CARMEN_SS_ID).getSheetByName('STOCK').getDataRange().getValues();
      for (var s = 1; s < dStock.length; s++) {
        var cod = String(dStock[s][0] || '').trim().toUpperCase();
        if (cod) stockMap[cod] = Number(dStock[s][2]) || 0;
      }
    } catch (eStk) { Logger.log('LAUNCH_recuperarPedidos stockMap: ' + eStk); }

    var tz = Session.getScriptTimeZone();
    var resultado = { recuperados: [], filas: 0, yaExistian: [], noEncontrados: [], sinItems: [] };

    for (var k = 0; k < numeros.length; k++) {
      var numero = numeros[k];
      if (yaExiste[numero]) { resultado.yaExistian.push(numero); continue; }
      var fila = porNumero[numero];
      if (!fila) { resultado.noEncontrados.push(numero); continue; }

      var items = null;
      try { items = JSON.parse(String(fila[8] || '')); } catch (eJ) { items = null; }
      if (!items || !items.length) { resultado.sinItems.push(numero); continue; }

      var reseller  = String(fila[2]  || '');
      var obs       = String(fila[7]  || '');
      var formaPago = String(fila[11] || '');
      var envio     = String(fila[12] || '');
      var fechaOrig = (fila[1] instanceof Date) ? fila[1] : new Date();
      var notaObs   = '[Recuperado ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm') + ' desde PEDIDOS_REPUESTOS]' + (obs ? ' ' + obs : '');

      var agregadas = 0;
      for (var it = 0; it < items.length; it++) {
        var item  = items[it] || {};
        var skuUp = String(item.sku || '').trim().toUpperCase();
        var cant  = Number(item.cantidad) || 0;
        if (!skuUp && !item.descripcion) continue;   // línea vacía/corrupta en el JSON, se saltea
        var prec     = Number(item.precio) || 0;
        var stkH     = stockMap[skuUp] !== undefined ? stockMap[skuUp] : '';
        var estadoNota = (stkH !== '' && stkH >= cant) ? 'Confirmado' : 'Pendiente_Revision';
        var newRow = hojaNota.getLastRow() + 1;
        hojaNota.appendRow([
          numero, reseller, item.sku || '', item.descripcion || '', cant, 0,
          '=E' + newRow + '-F' + newRow + '-Z' + newRow,
          prec, stkH, estadoNota, fechaOrig, envio, formaPago, notaObs
        ]);
        agregadas++;
      }
      if (agregadas) { resultado.recuperados.push(numero); resultado.filas += agregadas; yaExiste[numero] = true; }
      else resultado.sinItems.push(numero);
    }

    if (resultado.filas) SpreadsheetApp.flush();

    var msg = 'Recuperados: ' + resultado.recuperados.length + ' pedido(s), ' + resultado.filas + ' fila(s) nueva(s).';
    if (resultado.yaExistian.length)    msg += '\nYa existían (sin tocar): ' + resultado.yaExistian.join(', ') + '.';
    if (resultado.noEncontrados.length) msg += '\nNo encontrados en PEDIDOS_REPUESTOS: ' + resultado.noEncontrados.join(', ') + '.';
    if (resultado.sinItems.length)      msg += '\nSin ítems recuperables: ' + resultado.sinItems.join(', ') + '.';
    msg += '\nOjo: quedan sin Thread ID — para que WOS conteste en el hilo original de cada uno, correr después "Recuperar Thread IDs perdidos".';

    return { ok: true, total: numeros.length, resultado: resultado, msg: msg };
  } catch (e) {
    Logger.log('LAUNCH_recuperarPedidos: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Cuota diaria de Gmail restante ───────────────────────────
function _herr_cuotaGmail() {
  var restante = MailApp.getRemainingDailyQuota();
  var msg = 'Cuota Gmail disponible hoy: ' + restante + ' emails restantes.';
  return { ok: true, msg: msg, valor: restante };
}

// ── Health check general del sistema ─────────────────────────
function _herr_sysHealth() {
  var lines = [];
  var ok = true;
  try {
    var ss = SpreadsheetApp.openById(MASTER_SS_ID);
    var hojas = ['Ordenes de trabajo', 'STOCK_REPUESTOS', 'Resellers', 'SOLICITUDES_DESPACHO'];
    for (var h = 0; h < hojas.length; h++) {
      var hoja = ss.getSheetByName(hojas[h]);
      if (hoja) lines.push('✓ MASTER → ' + hojas[h] + ' (' + (hoja.getLastRow()-1) + ' filas)');
      else      { lines.push('✗ MASTER → ' + hojas[h] + ' NO ENCONTRADA'); ok = false; }
    }
  } catch(e) { lines.push('✗ MASTER spreadsheet: ' + e); ok = false; }
  try {
    var ssN = SpreadsheetApp.openById(NOTAS_SS_ID);
    var hojasN = ['Pedidos_resellers', 'WOS_Log'];
    for (var hn = 0; hn < hojasN.length; hn++) {
      var hojaN = ssN.getSheetByName(hojasN[hn]);
      if (hojaN) lines.push('✓ NOTAS → ' + hojasN[hn] + ' (' + (hojaN.getLastRow()-1) + ' filas)');
      else       { lines.push('✗ NOTAS → ' + hojasN[hn] + ' NO ENCONTRADA'); ok = false; }
    }
  } catch(e) { lines.push('✗ NOTAS spreadsheet: ' + e); ok = false; }
  try {
    var quota = MailApp.getRemainingDailyQuota();
    lines.push('✓ Gmail: ' + quota + ' emails/día disponibles');
  } catch(e) { lines.push('✗ Gmail: ' + e); ok = false; }
  return { ok: ok, msg: lines.join('\n') };
}

// ── Avisos Portal Reseller ────────────────────────────────────
function LAUNCH_getAviso() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('PORTAL_CONFIG');
    if (!hoja) return { activo: false, titulo: '', cuerpo: '' };
    var data = hoja.getDataRange().getValues();
    var map = {};
    for (var i = 0; i < data.length; i++) {
      map[String(data[i][0] || '').trim()] = String(data[i][1] || '').trim();
    }
    return {
      activo: String(map['AVISO_ACTIVO'] || '').toUpperCase() === 'TRUE',
      titulo: map['AVISO_TITULO'] || '',
      cuerpo: map['AVISO_CUERPO'] || ''
    };
  } catch(e) { return { activo: false, titulo: '', cuerpo: '' }; }
}

function LAUNCH_getWosConfig() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('WOS_CONFIG');
    var def  = { emailSoporte: 'soporteagrasdji@bidcom.com.ar', emailFact: 'Cecilia.f@bidcom.com.ar,lucia.c@bidcom.com.ar', pdfFolderId: '1yVefFM-vZ-Skmg2a_V9fsA-3XTAKx3Pz' };
    if (!hoja) return def;
    var data = hoja.getDataRange().getValues();
    var map  = {};
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) map[String(data[i][0]).trim()] = String(data[i][1] || '').trim();
    }
    return {
      emailSoporte: map['EMAIL_SOPORTE']     || def.emailSoporte,
      emailFact:    map['EMAIL_FACTURACION'] || def.emailFact,
      pdfFolderId:  map['WOS_PDF_FOLDER_ID'] || def.pdfFolderId
    };
  } catch(e) { return { emailSoporte: '', emailFact: '', pdfFolderId: '' }; }
}

function LAUNCH_setWosConfig(data) {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName('WOS_CONFIG') || ss.insertSheet('WOS_CONFIG');
    hoja.clearContents();
    hoja.getRange(1, 1, 3, 2).setValues([
      ['EMAIL_SOPORTE',     String(data.emailSoporte || '').trim()],
      ['EMAIL_FACTURACION', String(data.emailFact    || '').trim()],
      ['WOS_PDF_FOLDER_ID', String(data.pdfFolderId  || '').trim()]
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

// ── ComandasPedidos · _CONFIG ────────────────────────────────
// Esquema de la config editable de ComandasPedidos desde el Launcher.
// DEBE reflejar CP_CONFIG_DEFAULTS de ComandasPedidos/CP_Código.js (claves y sentido).
// tipo: 'sino' (SI/NO), 'num' (numérico), 'email'/'text' (texto).
var CP_CFG_SCHEMA = [
  { grupo:'Automático',   clave:'AUTO_MAIL_DESPACHO',     label:'Envío automático al reseller + RTV', tipo:'sino', ayuda:'SI = al detectar la guía manda solo el mail final al reseller y RTV. Requiere el trigger CP_autoMailEnvios instalado en ComandasPedidos.' },
  { grupo:'Automático',   clave:'EMAIL_PRUEBA',           label:'Modo prueba (redirigir TODO a este mail)', tipo:'email', ayuda:'Si tiene un mail, TODOS los correos se mandan SOLO ahí (no a los reales). VACIAR para volver a producción.' },
  { grupo:'Destinatarios',clave:'MAIL_APROBACION',        label:'Mail de aprobación (Sole)', tipo:'text', ayuda:'Destinatario del pedido de autorización en Masterchief.' },
  { grupo:'Destinatarios',clave:'MAIL_DESTINATARIOS',     label:'Destinatarios fijos del despacho', tipo:'text', ayuda:'Se agregan siempre al mail de despacho. Varios separados por coma.' },
  { grupo:'Destinatarios',clave:'MAIL_CC',                label:'CC (con copia)', tipo:'text', ayuda:'Varios separados por coma.' },
  { grupo:'Destinatarios',clave:'MAIL_BCC',               label:'CCO (copia oculta)', tipo:'text', ayuda:'Copia oculta del mail al reseller + aprobación. Varios con coma.' },
  { grupo:'Textos',       clave:'MAIL_ASUNTO',            label:'Asunto · mail al reseller', tipo:'text', ayuda:'Placeholders disponibles: {IDVENTA} {COMANDA} {CLIENTE}.' },
  { grupo:'Textos',       clave:'ASUNTO_APROBACION',      label:'Asunto · mail de aprobación', tipo:'text', ayuda:'Placeholders disponibles: {COMANDA} {IDVENTA}.' },
  { grupo:'Textos',       clave:'MAIL_REMITENTE_NOMBRE',  label:'Nombre del remitente', tipo:'text', ayuda:'Aparece como el "De" del correo.' },
  { grupo:'Textos',       clave:'OCA_TRACKING_URL',       label:'URL de tracking OCA', tipo:'text', ayuda:'Placeholder {GUIA}. Se usa para linkear la guía en el mail.' },
  { grupo:'Tiempos',      clave:'RECORDATORIO_HORAS',     label:'Recordatorio a Sole (horas)', tipo:'num', ayuda:'Cada cuántas horas recordar una comanda sin despachar.' },
  { grupo:'Tiempos',      clave:'SLA_WARN_HORAS',         label:'SLA · pasa a amarillo (horas)', tipo:'num', ayuda:'Antigüedad a partir de la cual el semáforo se pone amarillo.' },
  { grupo:'Tiempos',      clave:'SLA_DANGER_HORAS',       label:'SLA · pasa a rojo (horas)', tipo:'num', ayuda:'Antigüedad a partir de la cual el semáforo se pone rojo.' },
  { grupo:'Permisos',     clave:'OPERADORES_AUTORIZADOS', label:'Operadores autorizados', tipo:'text', ayuda:'Mails (coma) que pueden crear/borrar envíos y mandar mail. VACÍO = todos.' },
  { grupo:'Permisos',     clave:'MAIL_MAX_POR_10MIN',     label:'Tope de mails por 10 min', tipo:'num', ayuda:'Límite anti-abuso de correos enviados por ventana de 10 minutos.' }
];

// Lee la pestaña _CONFIG de ComandasPedidos → { ok, schema, valores:{CLAVE:valor}, existeHoja }.
function LAUNCH_getComandasConfig() {
  try {
    var ss   = SpreadsheetApp.openById(CP_LOG_SS_ID);
    var hoja = ss.getSheetByName(CP_CONFIG_TAB);
    var valores = {};
    if (hoja) {
      var data = hoja.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var k = String(data[i][0] || '').trim();
        if (k) valores[k.toUpperCase()] = String(data[i][1] == null ? '' : data[i][1]);
      }
    }
    return { ok: true, schema: CP_CFG_SCHEMA, valores: valores, existeHoja: !!hoja };
  } catch(e) {
    Logger.log('LAUNCH_getComandasConfig: ' + e);
    return { ok: false, error: e.toString(), schema: CP_CFG_SCHEMA, valores: {} };
  }
}

// Escribe SOLO las claves recibidas (por clave, sin borrar la hoja) → preserva cualquier otra
// fila/clave que ComandasPedidos maneje y que el Launcher no conozca. Crea la hoja si no existe.
function LAUNCH_setComandasConfig(data) {
  try {
    data = data || {};
    var ss   = SpreadsheetApp.openById(CP_LOG_SS_ID);
    var hoja = ss.getSheetByName(CP_CONFIG_TAB);
    if (!hoja) {
      hoja = ss.insertSheet(CP_CONFIG_TAB);
      hoja.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]);
      hoja.setFrozenRows(1);
      hoja.setColumnWidth(1, 210); hoja.setColumnWidth(2, 460);
    }
    // índice fila (1-based) por clave existente, para actualizar sin pisar claves ajenas
    var last = hoja.getLastRow();
    var filaDe = {};
    if (last >= 2) {
      var col = hoja.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < col.length; i++) {
        var k = String(col[i][0] || '').trim().toUpperCase();
        if (k) filaDe[k] = i + 2;
      }
    }
    var nuevas = [];
    Object.keys(data).forEach(function(clave) {
      var K = String(clave || '').trim().toUpperCase();
      if (!K) return;
      var val = data[clave] == null ? '' : String(data[clave]);
      if (filaDe[K]) hoja.getRange(filaDe[K], 2).setValue(val);
      else nuevas.push([K, val]);
    });
    if (nuevas.length) hoja.getRange(hoja.getLastRow() + 1, 1, nuevas.length, 2).setValues(nuevas);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) {
    Logger.log('LAUNCH_setComandasConfig: ' + e);
    return { ok: false, error: e.toString() };
  }
}

function LAUNCH_getCampana() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('PORTAL_CAMPANA');
    if (!hoja) return { activa: '', label: '', vence: '' };
    var data = hoja.getDataRange().getValues();
    var map  = {};
    for (var i = 0; i < data.length; i++) {
      map[String(data[i][0] || '').trim()] = String(data[i][1] || '').trim();
    }
    return {
      activa: map['CAMPANA_ACTIVA'] || '',
      label:  map['CAMPANA_LABEL']  || '',
      vence:  map['CAMPANA_VENCE']  || ''
    };
  } catch(e) { return { activa: '', label: '', vence: '' }; }
}

function LAUNCH_setCampana(data) {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName('PORTAL_CAMPANA') || ss.insertSheet('PORTAL_CAMPANA');
    hoja.clearContents();
    var activaId = data.activa ? String(data.label || '').trim() : '';
    hoja.getRange(1, 1, 3, 2).setValues([
      ['CAMPANA_ACTIVA', activaId],
      ['CAMPANA_LABEL',  String(data.label || '').trim()],
      ['CAMPANA_VENCE',  String(data.vence || '').trim()]
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function LAUNCH_getCampanaStats() {
  try {
    var ss    = SpreadsheetApp.openById(MASTER_SS_ID);
    var hCfg  = ss.getSheetByName('PORTAL_CAMPANA');
    var campanaActiva = '';
    if (hCfg) {
      var cfgData = hCfg.getDataRange().getValues();
      for (var c = 0; c < cfgData.length; c++) {
        if (String(cfgData[c][0]).trim() === 'CAMPANA_ACTIVA') {
          campanaActiva = String(cfgData[c][1] || '').trim();
          break;
        }
      }
    }
    var hoja = ss.getSheetByName('Pedidos_Campaña');
    if (!hoja || !campanaActiva) return { ok: true, total: 0, resellers: [] };
    var data = hoja.getDataRange().getValues();
    var map  = {};
    var tz   = Session.getScriptTimeZone();
    for (var i = 0; i < data.length; i++) {
      var rowCamp = String(data[i][1] || '').trim();
      var rowRes  = String(data[i][2] || '').trim();
      if (rowCamp !== campanaActiva || !rowRes) continue;
      if (!map[rowRes]) map[rowRes] = { items: 0, fecha: '' };
      map[rowRes].items++;
      var ts = data[i][0];
      if (ts instanceof Date) {
        try { map[rowRes].fecha = Utilities.formatDate(ts, tz, 'dd/MM HH:mm'); } catch(ef) {}
      }
    }
    var resellers = [];
    for (var r in map) {
      resellers.push({ nombre: r, items: map[r].items, fecha: map[r].fecha });
    }
    resellers.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });
    return { ok: true, total: resellers.length, resellers: resellers };
  } catch(e) { return { ok: false, total: 0, resellers: [] }; }
}

function LAUNCH_enviarRecordatorioCampana() {
  try {
    var ss = SpreadsheetApp.openById(MASTER_SS_ID);

    // Configuración de campaña
    var hCfg = ss.getSheetByName('PORTAL_CAMPANA');
    if (!hCfg) return { ok: false, error: 'Campaña no configurada.' };
    var cfgData = hCfg.getDataRange().getValues();
    var cfg = {};
    for (var c = 0; c < cfgData.length; c++) {
      cfg[String(cfgData[c][0]).trim()] = String(cfgData[c][1] || '').trim();
    }
    var campanaActiva = cfg['CAMPANA_ACTIVA'];
    var campanaLabel  = cfg['CAMPANA_LABEL'] || campanaActiva;
    var campanaVence  = cfg['CAMPANA_VENCE'] || '';
    if (!campanaActiva) return { ok: false, error: 'No hay campaña activa.' };

    // Quiénes ya enviaron estimado
    var submitted = {};
    var hPed = ss.getSheetByName('Pedidos_Campaña');
    if (hPed && hPed.getLastRow() > 0) {
      var pedData = hPed.getDataRange().getValues();
      for (var p = 0; p < pedData.length; p++) {
        if (String(pedData[p][1] || '').trim() === campanaActiva) {
          submitted[String(pedData[p][2] || '').trim().toLowerCase()] = true;
        }
      }
    }

    // Resellers con email
    var hRes = ss.getSheetByName('Resellers');
    if (!hRes) return { ok: false, error: 'Hoja Resellers no encontrada.' };
    var resData = hRes.getDataRange().getValues();

    var portalUrl    = PropertiesService.getScriptProperties().getProperty('URL_PORTAL') || '';
    var emailsEnviados = {};
    var enviados = 0, sinEmail = 0, yaCompletaron = 0;

    for (var r = 1; r < resData.length; r++) {
      var nombre = String(resData[r][0] || '').trim();
      var email  = String(resData[r][9] || '').trim().toLowerCase();
      if (!nombre) continue;
      if (submitted[nombre.toLowerCase()]) { yaCompletaron++; continue; }
      if (!email) { sinEmail++; continue; }
      if (emailsEnviados[email]) continue; // evitar duplicados (grupos)

      var subject = 'Recordatorio: ' + campanaLabel +
                    (campanaVence ? ' · Fecha límite: ' + campanaVence : '');
      var body =
        'Hola ' + nombre + ',\n\n' +
        'Te recordamos que todavía no cargaste tu estimado para la ' + campanaLabel + '.\n\n' +
        (campanaVence ? 'Fecha límite: ' + campanaVence + '\n\n' : '') +
        'Podés ingresar al Portal Reseller y dirigirte a la sección "Campaña" para completar tu pedido estimado.\n' +
        (portalUrl ? portalUrl + '\n\n' : '\n') +
        'Recordá que el estimado no representa un compromiso de compra y podés modificarlo hasta la fecha límite.\n\n' +
        'Saludos,\nBIDCOMAGRO';

      try {
        MailApp.sendEmail(email, subject, body);
        emailsEnviados[email] = true;
        enviados++;
      } catch(em) { Logger.log('Recordatorio email error [' + email + ']: ' + em); }
    }

    return { ok: true, enviados: enviados, sinEmail: sinEmail, yaCompletaron: yaCompletaron };
  } catch(e) {
    Logger.log('LAUNCH_enviarRecordatorioCampana: ' + e);
    return { ok: false, error: e.toString() };
  }
}

function LAUNCH_getVistasAviso() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('PORTAL_AVISO_VISTAS');
    if (!hoja) return { ok: true, vistas: [] };
    var data = hoja.getDataRange().getValues();
    var tz = Session.getScriptTimeZone();
    var vistas = [];
    for (var i = 0; i < data.length; i++) {
      if (!data[i][0]) continue;
      vistas.push({
        nombre: String(data[i][0]),
        fecha: data[i][1] instanceof Date ? Utilities.formatDate(data[i][1], tz, 'dd/MM HH:mm') : ''
      });
    }
    return { ok: true, vistas: vistas };
  } catch(e) { return { ok: false, vistas: [] }; }
}

function LAUNCH_limpiarVistasAviso() {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('PORTAL_AVISO_VISTAS');
    if (hoja) hoja.clearContents();
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function LAUNCH_setAviso(data) {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName('PORTAL_CONFIG') || ss.insertSheet('PORTAL_CONFIG');
    hoja.clearContents();
    hoja.getRange(1, 1, 3, 2).setValues([
      ['AVISO_ACTIVO', data.activo ? 'TRUE' : 'FALSE'],
      ['AVISO_TITULO', String(data.titulo || '').trim()],
      ['AVISO_CUERPO', String(data.cuerpo || '').trim()]
    ]);
    return { ok: true };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

// ── Curso / Eventos · Portal Reseller ────────────────────────
// Gestión de eventos desde el Launcher. Las hojas EVENTOS e INSCRIPCIONES_EVENTOS viven en el
// MASTER (mismas que usa RS_Eventos.js del Portal); el scope 'spreadsheets' ya permite leerlas/escribirlas.
//   EVENTOS:  A=ID B=Título C=Fecha D=Lugar E=Descripción F=Fecha límite G=Activo (vacío/SI visible)
//   INSCRIPCIONES_EVENTOS: A=Fecha B=EventoID C=Evento D=Reseller E=Email reseller F=Asiste(SI/NO) G=Nombre H=Email I=Comentario
var _LAUNCH_EVENTOS_TAB = 'EVENTOS';
var _LAUNCH_INSCRIP_TAB = 'INSCRIPCIONES_EVENTOS';

// eventoId → { totalPersonas, resellers: { nombre: { asiste, asistentes:[{nombre,email}] } } }
function _launchAgruparInscripciones(ss) {
  var out  = {};
  var hoja = ss.getSheetByName(_LAUNCH_INSCRIP_TAB);
  if (!hoja) return out;
  var d = hoja.getDataRange().getValues();
  for (var i = 1; i < d.length; i++) {
    var eid = String(d[i][1] || '').trim();
    var rs  = String(d[i][3] || '').trim();
    if (!eid || !rs) continue;
    if (!out[eid]) out[eid] = { totalPersonas: 0, resellers: {} };
    if (!out[eid].resellers[rs]) out[eid].resellers[rs] = { asiste: false, asistentes: [] };
    if (String(d[i][5] || '').toUpperCase() === 'SI') {
      out[eid].resellers[rs].asiste = true;
      var nom = String(d[i][6] || '').trim();
      if (nom) { out[eid].resellers[rs].asistentes.push({ nombre: nom, email: String(d[i][7] || '') }); out[eid].totalPersonas++; }
    }
  }
  return out;
}

function _launchEvActivo(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  if (s === '') return true;
  return !(s === 'no' || s === 'false' || s === '0' || s === 'oculto' || s === 'inactivo');
}

function LAUNCH_getEventos() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_EVENTOS_TAB);
    var tz   = Session.getScriptTimeZone();
    var agr  = _launchAgruparInscripciones(ss);
    var out  = [];
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var id     = String(d[i][0] || '').trim();
        var titulo = String(d[i][1] || '').trim();
        if (!id && !titulo) continue;
        var fecha  = (d[i][2] instanceof Date) ? Utilities.formatDate(d[i][2], tz, 'dd/MM/yyyy') : String(d[i][2] || '');
        var limite = (d[i][5] instanceof Date) ? Utilities.formatDate(d[i][5], tz, 'dd/MM/yyyy') : String(d[i][5] || '');
        var a = agr[id] || { totalPersonas: 0, resellers: {} };
        var van = 0, rk = Object.keys(a.resellers);
        for (var r = 0; r < rk.length; r++) if (a.resellers[rk[r]].asiste) van++;
        out.push({
          id: id, titulo: titulo, fecha: fecha, lugar: String(d[i][3] || ''),
          descripcion: String(d[i][4] || ''), limite: limite, activo: _launchEvActivo(d[i][6]),
          personas: a.totalPersonas, resellers: van
        });
      }
    }
    return { ok: true, eventos: out };
  } catch(e) { Logger.log('LAUNCH_getEventos: ' + e); return { ok: false, error: e.toString(), eventos: [] }; }
}

function LAUNCH_saveEvento(data) {
  try {
    data = data || {};
    var titulo = String(data.titulo || '').trim();
    if (!titulo) return { ok: false, error: 'Falta el título del evento.' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_EVENTOS_TAB);
    if (!hoja) {
      hoja = ss.insertSheet(_LAUNCH_EVENTOS_TAB);
      hoja.appendRow(['ID', 'Título', 'Fecha', 'Lugar', 'Descripción', 'Fecha límite inscripción', 'Activo']);
      hoja.setFrozenRows(1);
      hoja.getRange(1, 1, 1, 7).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    }
    var id = String(data.id || '').trim();
    if (!id) id = 'CURSO-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
    var fila = [
      id, titulo, String(data.fecha || '').trim(), String(data.lugar || '').trim(),
      String(data.descripcion || '').trim(), String(data.limite || '').trim(),
      data.activo ? 'SI' : 'NO'
    ];
    var d = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim() === id) { filaIdx = i + 1; break; }
    }
    if (filaIdx > 0) hoja.getRange(filaIdx, 1, 1, 7).setValues([fila]);
    else             hoja.appendRow(fila);
    SpreadsheetApp.flush();
    return { ok: true, id: id };
  } catch(e) { Logger.log('LAUNCH_saveEvento: ' + e); return { ok: false, error: e.toString() }; }
}

function LAUNCH_getEventoStats(eventoId) {
  try {
    eventoId = String(eventoId || '').trim();
    var ss  = SpreadsheetApp.openById(MASTER_SS_ID);
    var agr = _launchAgruparInscripciones(ss);
    var a   = agr[eventoId] || { totalPersonas: 0, resellers: {} };
    var rk  = Object.keys(a.resellers).sort();
    var lista = [], van = 0;
    for (var i = 0; i < rk.length; i++) {
      var g = a.resellers[rk[i]];
      if (g.asiste) van++;
      lista.push({ reseller: rk[i], asiste: g.asiste, personas: g.asistentes.length, asistentes: g.asistentes });
    }
    return { ok: true, totalPersonas: a.totalPersonas, totalResellers: van, lista: lista };
  } catch(e) { return { ok: false, error: e.toString(), totalPersonas: 0, totalResellers: 0, lista: [] }; }
}

// ── Videos de armado y desarmado · Portal Reseller ────────────
// Gestión de links (2 por modelo: armado y desarmado, que DJI comparte directo) desde el
// Launcher. La hoja vive en el MASTER (misma que usa RS_VideosModelos.js del Portal); el scope
// 'spreadsheets' ya permite leerla/escribirla — sin re-auth. Mismo patrón que EVENTOS, pero acá
// el propio nombre del modelo hace de identificador (no hay un ID separado).
//   VIDEOS_ARMADO_DESARMADO: A=Modelo B=Link armado C=Link desarmado D=Activo (vacío/SI visible)
var _LAUNCH_VIDEOS_TAB = 'VIDEOS_ARMADO_DESARMADO';

function LAUNCH_getVideosModelos() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_VIDEOS_TAB);
    var out  = [];
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var modelo = String(d[i][0] || '').trim();
        if (!modelo) continue;
        out.push({
          modelo: modelo, armado: String(d[i][1] || '').trim(), desarmado: String(d[i][2] || '').trim(),
          activo: _launchEvActivo(d[i][3])
        });
      }
    }
    out.sort(function(a, b) { return a.modelo.localeCompare(b.modelo); });
    return { ok: true, modelos: out };
  } catch(e) { Logger.log('LAUNCH_getVideosModelos: ' + e); return { ok: false, error: e.toString(), modelos: [] }; }
}

// data = { modelo, armado, desarmado, activo, modeloOriginal }. modeloOriginal (opcional) permite
// renombrar un modelo sin duplicar la fila — se busca por ese nombre, se guarda con el nuevo.
function LAUNCH_saveVideoModelo(data) {
  try {
    data = data || {};
    var modelo = String(data.modelo || '').trim();
    if (!modelo) return { ok: false, error: 'Falta el nombre del modelo.' };
    var armado    = String(data.armado    || '').trim();
    var desarmado = String(data.desarmado || '').trim();
    if (!armado && !desarmado) return { ok: false, error: 'Cargá al menos un link (armado o desarmado).' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_VIDEOS_TAB);
    if (!hoja) {
      hoja = ss.insertSheet(_LAUNCH_VIDEOS_TAB);
      hoja.appendRow(['Modelo', 'Link armado', 'Link desarmado', 'Activo']);
      hoja.setFrozenRows(1);
      hoja.getRange(1, 1, 1, 4).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    }
    var fila = [modelo, armado, desarmado, data.activo === false ? 'NO' : 'SI'];
    var buscar = String(data.modeloOriginal || modelo).trim().toLowerCase();
    var d = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() === buscar) { filaIdx = i + 1; break; }
    }
    if (filaIdx > 0) hoja.getRange(filaIdx, 1, 1, 4).setValues([fila]);
    else             hoja.appendRow(fila);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_saveVideoModelo: ' + e); return { ok: false, error: e.toString() }; }
}


// ── Config "Venta a prospectos (RTV)" — hoja Clave/Valor en el master, leída directo ──
// por PORTAL_RESELLER (RS_Prospectos.js). EMAIL_AUTORIZADOR: quién recibe el mail para
// autorizar cantidades. DESCUENTO_PCT: % de descuento sobre lista para estos pedidos
// (0 = PVP). Self-provisioning igual que el resto de configs de este Launcher.
var _LAUNCH_CONFIG_PROSPECTOS_TAB = 'CONFIG_PROSPECTOS';
function LAUNCH_getConfigProspectos() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_CONFIG_PROSPECTOS_TAB);
    var out  = { emailAutorizador: '', descuentoPct: 0 };
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var clave = String(d[i][0] || '').trim().toUpperCase();
        var valor = d[i][1];
        if (clave === 'EMAIL_AUTORIZADOR') out.emailAutorizador = String(valor || '').trim();
        else if (clave === 'DESCUENTO_PCT') out.descuentoPct = Number(valor) || 0;
      }
    }
    return { ok: true, emailAutorizador: out.emailAutorizador, descuentoPct: out.descuentoPct };
  } catch(e) { Logger.log('LAUNCH_getConfigProspectos: ' + e); return { ok: false, error: e.toString(), emailAutorizador: '', descuentoPct: 0 }; }
}

function LAUNCH_saveConfigProspectos(emailAutorizador, descuentoPct) {
  try {
    var email = String(emailAutorizador || '').trim();
    var pct   = Number(descuentoPct);
    if (!email)     return { ok: false, error: 'Falta el email del autorizador.' };
    if (isNaN(pct) || pct < 0 || pct > 100) return { ok: false, error: 'Descuento inválido (0-100).' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_CONFIG_PROSPECTOS_TAB);
    if (!hoja) {
      hoja = ss.insertSheet(_LAUNCH_CONFIG_PROSPECTOS_TAB);
      hoja.appendRow(['Clave', 'Valor']);
      hoja.setFrozenRows(1);
      hoja.getRange(1, 1, 1, 2).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
      hoja.appendRow(['EMAIL_AUTORIZADOR', '']);
      hoja.appendRow(['DESCUENTO_PCT', 0]);
    }
    var d = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var clave = String(d[i][0] || '').trim().toUpperCase();
      if (clave === 'EMAIL_AUTORIZADOR') hoja.getRange(i + 1, 2).setValue(email);
      else if (clave === 'DESCUENTO_PCT') hoja.getRange(i + 1, 2).setValue(pct);
    }
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_saveConfigProspectos: ' + e); return { ok: false, error: e.toString() }; }
}


function LAUNCH_getResellersLista() {
  try {
    var hRes = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Resellers');
    if (!hRes) return [];
    var data  = hRes.getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < data.length; i++) {
      var nombre = String(data[i][0] || '').trim();
      if (nombre) lista.push(nombre);
    }
    return lista;
  } catch(e) { return []; }
}

// Lista de resellers con su estado de activación (para el modal de gestión).
function LAUNCH_getResellersConEstado() {
  try {
    var hRes = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Resellers');
    if (!hRes) return [];
    var data  = hRes.getDataRange().getValues();
    var lista = [];
    for (var i = 1; i < data.length; i++) {
      var nombre = String(data[i][0] || '').trim();
      if (!nombre) continue;
      lista.push({ nombre: nombre, activo: !_launchResellerInactivo(data[i]) });
    }
    return lista;
  } catch(e) { Logger.log('LAUNCH_getResellersConEstado: ' + e); return []; }
}

// Da de baja (activo=false) o reactiva (activo=true) un reseller escribiendo col Q.
// Reactivar = celda vacía; baja = "NO". Conserva email/PIN. Afecta TODAS las filas del mismo nombre.
function LAUNCH_setResellerActivo(nombre, activo) {
  try {
    var nom = String(nombre || '').trim();
    if (!nom) return { ok: false, error: 'Nombre requerido.' };
    var hRes = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Resellers');
    if (!hRes) return { ok: false, error: 'Hoja Resellers no encontrada.' };
    var data   = hRes.getDataRange().getValues();
    var nomLow = nom.toLowerCase();
    var valor  = activo ? '' : 'NO';
    var filas  = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() !== nomLow) continue;
      hRes.getRange(i + 1, RES_ACTIVO_COL).setValue(valor);
      filas++;
    }
    if (!filas) return { ok: false, error: 'Reseller no encontrado.' };
    SpreadsheetApp.flush();
    return { ok: true, activo: !!activo, filas: filas };
  } catch(e) {
    Logger.log('LAUNCH_setResellerActivo: ' + e);
    return { ok: false, error: e.toString() };
  }
}

function _pinEmailHtml(empresa, pin, esReset) {
  var intro = esReset
    ? 'Se generó un nuevo PIN para tu cuenta en el Portal Reseller de BIDCOMAGRO.'
    : 'A continuación encontrás tus credenciales de acceso al Portal Reseller de BIDCOMAGRO.';
  var portalUrl = PropertiesService.getScriptProperties().getProperty('URL_PORTAL') || '';
  return '<div style="font-family:Arial,sans-serif;max-width:480px;margin:auto">' +
    '<div style="background:#0077b6;padding:24px 28px;border-radius:10px 10px 0 0">' +
      '<h2 style="color:#fff;margin:0;font-size:18px">BIDCOMAGRO · Portal Reseller</h2>' +
    '</div>' +
    '<div style="background:#f5f8fa;padding:24px 28px;border-radius:0 0 10px 10px;border:1px solid #dde3ea">' +
      '<p style="margin:0 0 16px;color:#222;font-size:14px">' + intro + '</p>' +
      '<table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:18px">' +
        '<tr><td style="padding:8px 12px;background:#fff;border:1px solid #dde3ea;color:#555;width:40%">Empresa</td>' +
            '<td style="padding:8px 12px;background:#fff;border:1px solid #dde3ea;color:#222;font-weight:700">' + empresa + '</td></tr>' +
        '<tr><td style="padding:8px 12px;background:#fff;border:1px solid #dde3ea;color:#555">PIN</td>' +
            '<td style="padding:8px 12px;background:#fff;border:1px solid #dde3ea;color:#0077b6;font-size:22px;font-weight:700;letter-spacing:6px">' + pin + '</td></tr>' +
      '</table>' +
      (portalUrl ? '<p style="margin:0 0 18px;font-size:13px"><a href="' + portalUrl + '" style="color:#0077b6">' + portalUrl + '</a></p>' : '') +
      '<p style="margin:0;font-size:12px;color:#888">Si no solicitaste este mensaje, podés ignorarlo.</p>' +
    '</div>' +
  '</div>';
}

function LAUNCH_recordarPin(nombre) {
  try {
    if (!nombre) return { ok: false, error: 'Nombre requerido.' };
    var hRes = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Resellers');
    if (!hRes) return { ok: false, error: 'Hoja Resellers no encontrada.' };
    var data   = hRes.getDataRange().getValues();
    var nomLow = String(nombre).trim().toLowerCase();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() !== nomLow) continue;
      var empresa = String(data[i][0]  || '').trim();
      var email   = String(data[i][9]  || '').trim();
      var pin     = String(data[i][10] || '').trim();
      if (!email) return { ok: false, error: 'Sin email registrado.' };
      if (!pin)   return { ok: false, error: 'Sin PIN registrado.' };
      MailApp.sendEmail({
        to:       email,
        subject:  'Tus credenciales de acceso — Portal Reseller BIDCOMAGRO',
        htmlBody: _pinEmailHtml(empresa, pin, false)
      });
      return { ok: true, email: email };
    }
    return { ok: false, error: 'Reseller no encontrado.' };
  } catch(e) {
    Logger.log('LAUNCH_recordarPin: ' + e);
    return { ok: false, error: e.toString() };
  }
}

function LAUNCH_resetPin(nombre) {
  try {
    if (!nombre) return { ok: false, error: 'Nombre requerido.' };
    var hRes = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Resellers');
    if (!hRes) return { ok: false, error: 'Hoja Resellers no encontrada.' };
    var data   = hRes.getDataRange().getValues();
    var nomLow = String(nombre).trim().toLowerCase();
    var usados = {};
    for (var j = 1; j < data.length; j++) {
      var p = String(data[j][10] || '').trim();
      if (p) usados[p] = true;
    }
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim().toLowerCase() !== nomLow) continue;
      var empresa = String(data[i][0] || '').trim();
      var email   = String(data[i][9] || '').trim();
      if (!email) return { ok: false, error: 'Sin email registrado.' };
      var pin, intentos = 0;
      do {
        pin = String(Math.floor(1000 + Math.random() * 9000));
        intentos++;
      } while (usados[pin] && intentos < 100);
      hRes.getRange(i + 1, 11).setValue(pin);
      MailApp.sendEmail({
        to:       email,
        subject:  'Tu nuevo PIN de acceso — Portal Reseller BIDCOMAGRO',
        htmlBody: _pinEmailHtml(empresa, pin, true)
      });
      return { ok: true, email: email };
    }
    return { ok: false, error: 'Reseller no encontrado.' };
  } catch(e) {
    Logger.log('LAUNCH_resetPin: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── Migración: normaliza el separador de repuestos en 'Ordenes de trabajo' ─────────
// Versiones viejas del Portal unían los repuestos de la col Q con '\n', pero el HUB los lee con
// split(' ; ') → tomaba solo el primero. Esto recorre la columna y reescribe las celdas que tengan
// salto de línea al separador canónico ' ; '. aplicar=false → solo cuenta/preview (no escribe nada).
function LAUNCH_migrarRepuestosSeparador(aplicar) {
  try {
    var hoja = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName('Ordenes de trabajo');
    if (!hoja) return { ok: false, error: "No se encontró la hoja 'Ordenes de trabajo'." };
    var COL_REP = 17, COL_OT = 3;  // Q = REPUESTOS, C = N° de OT
    var last = hoja.getLastRow();
    if (last < 2) return { ok: true, revisadas: 0, corregidas: 0, ejemplos: [], aplicado: false };
    var rng  = hoja.getRange(2, COL_REP, last - 1, 1);
    var vals = rng.getValues();
    var ots  = hoja.getRange(2, COL_OT, last - 1, 1).getValues();
    var corregidas = 0, ejemplos = [];
    for (var i = 0; i < vals.length; i++) {
      var v = String(vals[i][0] == null ? '' : vals[i][0]);
      if (v.indexOf('\n') === -1 && v.indexOf('\r') === -1) continue;   // sin salto de línea → nada que migrar
      var partes = v.split(/ ; |\r?\n/), limpio = [];
      for (var p = 0; p < partes.length; p++) { var t = partes[p].trim(); if (t) limpio.push(t); }
      var nuevo = limpio.join(' ; ');
      if (nuevo !== v) {
        corregidas++;
        if (ejemplos.length < 6) ejemplos.push({ ot: String(ots[i][0] || ''), antes: v, despues: nuevo });
        if (aplicar) vals[i][0] = nuevo;
      }
    }
    if (aplicar && corregidas) rng.setValues(vals);
    return { ok: true, revisadas: vals.length, corregidas: corregidas, ejemplos: ejemplos, aplicado: !!aplicar };
  } catch(e) {
    Logger.log('LAUNCH_migrarRepuestosSeparador: ' + e);
    return { ok: false, error: e.toString() };
  }
}
