// @version 1.16
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

function _launchVideosHoja(ss) {
  var hoja = ss.getSheetByName(_LAUNCH_VIDEOS_TAB);
  if (!hoja) {
    hoja = ss.insertSheet(_LAUNCH_VIDEOS_TAB);
    hoja.appendRow(['Modelo', 'Componente', 'Link armado', 'Link desarmado', 'Activo', 'Es dron', 'Drones asociados']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 7).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    return hoja;
  }
  // Migración: hojas creadas antes de agrupar por dron tenían solo 5 columnas. Se agregan
  // 'Es dron'/'Drones asociados' vacías para todas las filas existentes — a propósito no
  // se adivina cuáles son drones acá: ni esa lista ni la de qué accesorio sirve para qué
  // dron se hardcodean en ningún lado del código, porque el catálogo de modelos cambia
  // con el tiempo. Se marca a mano desde el Launcher (checkbox "Es un dron" por fila).
  if (hoja.getLastColumn() < 7) {
    hoja.getRange(1, 6, 1, 2).setValues([['Es dron', 'Drones asociados']])
      .setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
  }
  return hoja;
}

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
          modelo: modelo, componente: String(d[i][1] || '').trim(),
          armado: String(d[i][2] || '').trim(), desarmado: String(d[i][3] || '').trim(),
          activo: _launchEvActivo(d[i][4]),
          esDron: String(d[i][5] || '').trim().toUpperCase() === 'SI',
          dronesAsociados: String(d[i][6] || '').split(';').map(function(s) { return s.trim(); }).filter(function(s) { return s; })
        });
      }
    }
    out.sort(function(a, b) {
      var c = a.modelo.localeCompare(b.modelo);
      return c !== 0 ? c : a.componente.localeCompare(b.componente);
    });
    return { ok: true, modelos: out };
  } catch(e) { Logger.log('LAUNCH_getVideosModelos: ' + e); return { ok: false, error: e.toString(), modelos: [] }; }
}

// data = { modelo, componente, armado, desarmado, activo, esDron, dronesAsociados, modeloOriginal, componenteOriginal }.
// La clave real es (modelo, componente) — un modelo con varias piezas (ej. T100: Aeronave /
// Sistema de siembra / Sistema de elevación) tiene una fila por componente. *Original (opcional)
// permite renombrar sin duplicar la fila — se busca por esos valores, se guarda con los nuevos.
// esDron: true si la fila es una pieza propia de un dron. dronesAsociados: array de nombres
// de dron para los que sirve este accesorio (solo aplica si esDron=false) — un mismo cargador
// o generador puede servir para varios modelos, se guarda separado por ';'.
function LAUNCH_saveVideoModelo(data) {
  try {
    data = data || {};
    var modelo = String(data.modelo || '').trim();
    if (!modelo) return { ok: false, error: 'Falta el nombre del modelo.' };
    var componente = String(data.componente || '').trim();
    var armado     = String(data.armado     || '').trim();
    var desarmado  = String(data.desarmado  || '').trim();
    if (!armado && !desarmado) return { ok: false, error: 'Cargá al menos un link (armado o desarmado).' };
    var esDron = !!data.esDron;
    var dronesAsociados = (!esDron && Array.isArray(data.dronesAsociados)) ? data.dronesAsociados : [];
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchVideosHoja(ss);
    var fila = [modelo, componente, armado, desarmado, data.activo === false ? 'NO' : 'SI', esDron ? 'SI' : 'NO', dronesAsociados.join(';')];
    var buscarModelo     = String(data.modeloOriginal     || modelo).trim().toLowerCase();
    var buscarComponente = String(data.componenteOriginal != null ? data.componenteOriginal : componente).trim().toLowerCase();
    var d = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() !== buscarModelo) continue;
      if (String(d[i][1] || '').trim().toLowerCase() !== buscarComponente) continue;
      filaIdx = i + 1; break;
    }
    if (filaIdx > 0) hoja.getRange(filaIdx, 1, 1, 7).setValues([fila]);
    else             hoja.appendRow(fila);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_saveVideoModelo: ' + e); return { ok: false, error: e.toString() }; }
}

// ── Carga inicial del catálogo DJI en inglés (curado a mano con el usuario, agosto 2026) ──
// Un solo uso: agrega las filas que falten por (modelo, componente); si una ya existe,
// la deja como está (no pisa ediciones manuales posteriores). Botón "Cargar catálogo DJI
// (EN)" en el modal de Videos del Launcher.
function LAUNCH_seedVideosModelosDji() {
  var CDN = 'https://cdn.djivideos.com/watch/';
  var filas = [
    // [Modelo, Componente, ID armado, ID desarmado]
    ['T100',  'Aeronave',              'c13a7902-30c0-4d82-9a0e-3cf486917b8d', 'fa3b6935-c21c-4707-9522-bb80c327befe'],
    ['T100',  'Sistema de siembra',    '38c9a4d4-bb33-4386-b393-3fa03599832e', 'c60596d5-7cb2-4dab-8829-2b3dc0b731ab'],
    ['T100',  'Sistema de elevación',  '7ae824b6-d382-40b4-ad3e-cf4d1b1e5062', '4b514ba0-47c0-424b-97ea-815ac957941f'],
    ['T70',   'Aeronave',              '46990dd5-390a-426e-b37c-9e4474136028', '1a332314-c2a3-42b0-a899-9b2a460ef3a6'],
    ['T70',   'Sistema de siembra',    '83a3201d-c652-4c11-805e-8c03c2c07360', '397c1479-3437-4cf0-9d37-81f2b44a4880'],
    ['T25P',  'Aeronave',              '5b42c81f-6825-4e13-bec6-49ee6acf5ffb', '69411446-87a6-42a5-953d-7d1c045031b7'],
    ['T25P',  'Sistema de siembra',    'df1084b9-d096-4b90-a904-cbe1a9639224', 'f7872915-52ed-4b81-b3d0-b85c39a1f0a1'],
    ['T25',   '',                      'fc2efca0-94ed-44b6-9a67-fffc1c81bcfd', 'b863b1ff-e81a-4ed3-bad6-8a99a3e5e185'],
    ['T50',   'Aeronave',              'faa7d926-e0ae-4921-a7ce-9b28fb060253', '97a80444-33bd-498f-bce9-a63d55eda3d4'],
    ['T50',   'Sistema de siembra',    'eb6a15fc-4b33-4691-83eb-bb9801a61b3a', 'cf05d57e-604c-4af8-8002-08de049c810b'],
    ['T55',   'Aeronave',              'e36409c0-45b3-4236-a7cf-eb3503523e7b', 'd2a3e922-ccc3-4e29-92fb-9d75a8981a00'],
    ['T55',   'Sistema de elevación (DL100)', '3016551e-e20d-43b2-9b0c-dde317345eb9', '0e57ae19-6088-4597-b8fb-8e20ca193e64'],
    ['T70S',  'Sistema de elevación (DL100)', '3016551e-e20d-43b2-9b0c-dde317345eb9', '0e57ae19-6088-4597-b8fb-8e20ca193e64'],
    ['T70S',  'Sistema de siembra (DS125L)',  'ff47da19-94be-412a-8e9b-268fafc4b5cc', '379fb0ac-20ff-40f5-8b91-17fcdd38f0d8'],
    ['T100S', 'Sistema de elevación (DL100)', '3016551e-e20d-43b2-9b0c-dde317345eb9', '0e57ae19-6088-4597-b8fb-8e20ca193e64'],
    ['T100S', 'Sistema de siembra (DS125L)',  'ff47da19-94be-412a-8e9b-268fafc4b5cc', '379fb0ac-20ff-40f5-8b91-17fcdd38f0d8'],
    ['DJI RC Plus 2',                 '', '3a832d5f-825e-481d-80b0-c70a13c77333', '83e10176-3227-410d-9b11-3c6f019137c9'],
    ['DJI RC Plus',                   '', 'f6c913fe-c69f-4c01-826c-a125ff2720e7', 'dc5ec0e8-02fb-4121-a7b5-2c9493c202c7'],
    ['D-RTK 3',                       '', '85c47220-8e46-4a8e-b751-e12335536b3b', '6ad80901-2042-4d49-83f2-72ffed1452b4'],
    ['DB2160 Intelligent Flight Battery', '', '244e8ab4-ac35-41bc-a099-a5509cce7436', 'f5c4a00f-21bc-478a-9d4b-eb6a6ff74ccc'],
    ['DB1580 Intelligent Flight Battery', '', '2ffa0446-da58-42a4-8765-78e39a04898d', 'a14b2ad4-3d4a-487d-aa0d-aadc6b2447b3'],
    ['DB1050 Intelligent Flight Battery', '', '019d74c6-5366-4482-b06d-1cf1ddc2d46b', 'b08ef542-065a-46f7-9295-7e48b1f5bbee'],
    ['C12000 Smart Charger',          '', 'ef7daa72-6c9d-4b5c-9195-9caba83155a3', '34d562ba-13fa-42b3-b095-43948780c77c'],
    ['T40 Intelligent Charger',       '', '18d838f1-da3b-4c3c-a6f8-19b93d8631f7', 'df12f7b8-1140-4cfc-8d2b-029054eeba85'],
    ['T20P Intelligent Charger',      '', 'e7d912f6-5ca6-4711-99e7-03f905d19019', '5a18e473-6cc0-4fb0-b541-8515df8cd077'],
    ['C7000 Intelligent Charger',     '', '6a97a823-eec3-4259-978f-b040645f35c6', '6d2ae79d-1ac4-4c09-af77-c7087f79cafb'],
    ['D14000 Generator',              '', '7d913f56-cf7d-4307-ba81-408b4a6c2d0a', '17373bd0-39a6-4e2e-97fc-b5b2f3530e1a'],
    ['D6000i Generator',              '', 'f2df2299-27d9-4f86-a354-27c72a8e08ce', '5be7960d-a768-4f08-91f8-98240f1b6b4d'],
    ['D12000iEP Generator', 'Parte 1', '408f9140-3490-4394-a222-3c9d307847b2', '800a0faf-e337-435c-b27d-55256c4aec05'],
    ['D12000iEP Generator', 'Parte 2', 'd3cd3e81-d274-4c37-93b5-b731ef803eaa', '30af96d4-e89a-4370-bead-7c037c43d79f'],
    ['D8000iE Generator',             '', 'a71e3d49-6cfe-403f-b361-d4ecf16ad9df', 'fbc62646-d8a3-4717-84a8-724b162ce1bb'],
    ['DJI O4 Relay',                  '', '498e2e62-4ea0-4702-90ae-a0ef84a335f7', '4328485a-c500-4c9b-a01d-33d70c984bcd'],
    ['DN6720 Mist Sprinkler Combo',   '', '0cac4cad-c039-4e62-bb2b-b89ae0f17ac5', '923af150-ab2c-431e-aa90-e9820d4caebb']
  ];
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchVideosHoja(ss);
    var d = hoja.getDataRange().getValues();
    var existentes = {};
    for (var i = 1; i < d.length; i++) {
      var k = String(d[i][0] || '').trim().toLowerCase() + '||' + String(d[i][1] || '').trim().toLowerCase();
      existentes[k] = true;
    }
    var agregados = 0, saltados = 0;
    for (var f = 0; f < filas.length; f++) {
      var modelo = filas[f][0], componente = filas[f][1];
      var key = modelo.trim().toLowerCase() + '||' + componente.trim().toLowerCase();
      if (existentes[key]) { saltados++; continue; }
      // 'Es dron' queda en NO por default — se marca a mano desde el Launcher (ver
      // _launchVideosHoja), no se adivina por nombre de modelo.
      hoja.appendRow([modelo, componente, CDN + filas[f][2], CDN + filas[f][3], 'SI', 'NO', '']);
      existentes[key] = true;
      agregados++;
    }
    SpreadsheetApp.flush();
    return { ok: true, agregados: agregados, saltados: saltados };
  } catch(e) { Logger.log('LAUNCH_seedVideosModelosDji: ' + e); return { ok: false, error: e.toString() }; }
}


// ── Material de Cursos — Portal Reseller ──────────────────────────────────
// Pedido del usuario: "tengo información valiosa de cursos que hemos hecho, ¿dónde podría
// colocarla?" — se decidió con el usuario un modal buscable propio (mismo patrón que Videos
// armado/desarmado más arriba: hoja self-provisioning en el MASTER, gestión desde el Launcher,
// lectura read-only desde el Portal). A diferencia de Videos, acá no hace falta agrupar/anidar
// nada — cada fila es un curso/material independiente, así que el modelo de datos es plano.
//   CURSOS_MATERIAL: A=Título B=Categoría C=Descripción D=Link E=Fecha F=Activo (vacío/SI visible)
var _LAUNCH_CURSOS_TAB = 'CURSOS_MATERIAL';

function _launchCursosHoja(ss) {
  var hoja = ss.getSheetByName(_LAUNCH_CURSOS_TAB);
  if (!hoja) {
    hoja = ss.insertSheet(_LAUNCH_CURSOS_TAB);
    hoja.appendRow(['Título', 'Categoría', 'Descripción', 'Link', 'Fecha', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 6).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
  }
  return hoja;
}

function LAUNCH_getCursosMaterial() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = ss.getSheetByName(_LAUNCH_CURSOS_TAB);
    var out  = [];
    if (hoja) {
      var d = hoja.getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        var titulo = String(d[i][0] || '').trim();
        if (!titulo) continue;
        var fechaRaw = d[i][4];
        out.push({
          titulo: titulo, categoria: String(d[i][1] || '').trim(),
          descripcion: String(d[i][2] || '').trim(), link: String(d[i][3] || '').trim(),
          fecha: (fechaRaw instanceof Date) ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fechaRaw || '').trim(),
          activo: _launchEvActivo(d[i][5])
        });
      }
    }
    // Más reciente primero (sin fecha, al final) — es la lectura natural de "cursos que ya dimos".
    out.sort(function(a, b) { return (b.fecha || '').localeCompare(a.fecha || ''); });
    return { ok: true, cursos: out };
  } catch(e) { Logger.log('LAUNCH_getCursosMaterial: ' + e); return { ok: false, error: e.toString(), cursos: [] }; }
}

// data = { titulo, categoria, descripcion, link, fecha, activo, tituloOriginal }.
// La clave real es el título — tituloOriginal (opcional) permite renombrar sin duplicar la fila.
function LAUNCH_saveCursoMaterial(data) {
  try {
    data = data || {};
    var titulo = String(data.titulo || '').trim();
    if (!titulo) return { ok: false, error: 'Falta el título del curso.' };
    var link = String(data.link || '').trim();
    if (!link) return { ok: false, error: 'Cargá el link del material.' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchCursosHoja(ss);
    var fila = [
      titulo, String(data.categoria || '').trim(), String(data.descripcion || '').trim(),
      link, String(data.fecha || '').trim(), data.activo === false ? 'NO' : 'SI'
    ];
    var buscar = String(data.tituloOriginal || titulo).trim().toLowerCase();
    var d = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() !== buscar) continue;
      filaIdx = i + 1; break;
    }
    if (filaIdx > 0) hoja.getRange(filaIdx, 1, 1, 6).setValues([fila]);
    else             hoja.appendRow(fila);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_saveCursoMaterial: ' + e); return { ok: false, error: e.toString() }; }
}


// ── Plantillas Cotizador — hoja compartida con PORTAL_RESELLER (RS_Cotizador.js) ──
// Mismo nombre de pestaña y MISMO esquema de columnas que
// _asegurarHojaPlantillasCotizador() de RS_Cotizador.js — no reordenar ni agregar columnas
// acá sin actualizar ese archivo también, los dos proyectos leen/escriben la misma hoja.
// PLANTILLAS_COTIZADOR: A=Plantilla B=Reseller (vacío=general, acá SIEMPRE vacío — las
// privadas de un reseller se arman desde el Portal, no desde acá) C=Tipo (REPUESTO|MANO_OBRA)
// D=SKU_Codigo E=Descripción F=Cantidad G=Fecha. Cada plantilla = varias filas.
var _LAUNCH_PLANTILLAS_TAB = 'PLANTILLAS_COTIZADOR';

function _launchPlantillasHoja(ss) {
  var hoja = ss.getSheetByName(_LAUNCH_PLANTILLAS_TAB);
  if (!hoja) {
    hoja = ss.insertSheet(_LAUNCH_PLANTILLAS_TAB);
    hoja.appendRow(['Plantilla', 'Reseller', 'Tipo', 'SKU_Codigo', 'Descripción', 'Cantidad', 'Fecha']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 7).setBackground('#3a9e3a').setFontColor('#fff').setFontWeight('bold');
  }
  return hoja;
}

// Solo las GENERALES (Reseller vacío) — las privadas de cada reseller no se administran acá.
function LAUNCH_getPlantillasCotizador() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchPlantillasHoja(ss);
    var d    = hoja.getDataRange().getValues();
    var mapa = {};
    for (var i = 1; i < d.length; i++) {
      var nombre = String(d[i][0] || '').trim();
      if (!nombre) continue;
      if (String(d[i][1] || '').trim()) continue; // con Reseller cargado → privada, no es de acá
      if (!mapa[nombre]) mapa[nombre] = { nombre: nombre, items: [], mo: [] };
      var tipo  = String(d[i][2] || '').trim().toUpperCase();
      var linea = { sku: String(d[i][3] || '').trim(), descripcion: String(d[i][4] || '').trim(), cantidad: Number(d[i][5]) || 1 };
      if (tipo === 'MANO_OBRA') mapa[nombre].mo.push(linea); else mapa[nombre].items.push(linea);
    }
    var out = [];
    for (var k in mapa) out.push(mapa[k]);
    out.sort(function(a, b) { return a.nombre.localeCompare(b.nombre); });
    return { ok: true, plantillas: out };
  } catch(e) { Logger.log('LAUNCH_getPlantillasCotizador: ' + e); return { ok: false, error: e.toString(), plantillas: [] }; }
}

// data = { nombre, nombreOriginal (para renombrar sin duplicar), items: [{sku,cantidad,descripcion}],
//          mo: [{sku,descripcion}] }. Reemplaza TODAS las filas de la plantilla (borra + reescribe)
// en vez de tratar de mergear línea por línea — más simple y sin riesgo de filas huérfanas.
function LAUNCH_guardarPlantillaCotizador(data) {
  try {
    data = data || {};
    var nombre = String(data.nombre || '').trim();
    if (!nombre) return { ok: false, error: 'Falta el nombre de la plantilla.' };
    var items = data.items || [];
    var mo    = data.mo || [];
    if (!items.length && !mo.length) return { ok: false, error: 'Agregá al menos un repuesto o una mano de obra.' };

    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchPlantillasHoja(ss);
    var buscar = String(data.nombreOriginal || nombre).trim().toLowerCase();
    var d = hoja.getDataRange().getValues();
    var filasBorrar = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim()) continue; // solo generales
      if (String(d[i][0] || '').trim().toLowerCase() === buscar) filasBorrar.push(i + 1);
    }
    filasBorrar.sort(function(a, b) { return b - a; });
    for (var f = 0; f < filasBorrar.length; f++) hoja.deleteRow(filasBorrar[f]);

    var ahora = new Date();
    var filas = [];
    for (var j = 0; j < items.length; j++) {
      var it = items[j];
      var sku = String(it.sku || '').trim();
      var descIt = String(it.descripcion || '').trim();
      if (!sku && !descIt) continue;
      var cant = Math.floor(Number(it.cantidad)) || 1;
      filas.push([nombre, '', 'REPUESTO', sku, descIt, cant, ahora]);
    }
    for (var m = 0; m < mo.length; m++) {
      var mm = mo[m];
      var cod = String(mm.sku || '').trim();
      var dsc = String(mm.descripcion || '').trim();
      if (!cod && !dsc) continue;
      filas.push([nombre, '', 'MANO_OBRA', cod, dsc, 1, ahora]);
    }
    if (!filas.length) return { ok: false, error: 'Agregá al menos un repuesto o una mano de obra.' };
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_guardarPlantillaCotizador: ' + e); return { ok: false, error: e.toString() }; }
}

function LAUNCH_eliminarPlantillaCotizador(nombre) {
  try {
    nombre = String(nombre || '').trim();
    if (!nombre) return { ok: false, error: 'Falta el nombre de la plantilla.' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchPlantillasHoja(ss);
    var d = hoja.getDataRange().getValues();
    var nombreLower = nombre.toLowerCase();
    var filasBorrar = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim()) continue; // solo generales
      if (String(d[i][0] || '').trim().toLowerCase() === nombreLower) filasBorrar.push(i + 1);
    }
    if (!filasBorrar.length) return { ok: false, error: 'No se encontró esa plantilla.' };
    filasBorrar.sort(function(a, b) { return b - a; });
    for (var f = 0; f < filasBorrar.length; f++) hoja.deleteRow(filasBorrar[f]);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_eliminarPlantillaCotizador: ' + e); return { ok: false, error: e.toString() }; }
}


// ── Repuestos recomendados DJI / Mantenimiento sugerido DJI (Portal Reseller) ──
// Pedido del usuario: 2 listas más en "Documentación técnica", cada una "1 link por
// modelo" (sin la complejidad de accesorios compartidos entre modelos que sí tiene
// Videos armado/desarmado — decisión confirmada). Mismo modelo de datos plano que
// CURSOS_MATERIAL, pero con Modelo como clave en vez de Título. Hojas independientes
// (no comparten pestaña) porque conceptualmente son documentos distintos — mismo
// criterio que Videos/Cursos siendo pestañas separadas.
//   REPUESTOS_RECOMENDADOS_DJI / MANTENIMIENTO_DJI: A=Modelo B=Link C=Activo (vacío/SI visible)
var _LAUNCH_REPUESTOS_REC_TAB = 'REPUESTOS_RECOMENDADOS_DJI';
var _LAUNCH_MANTENIMIENTO_TAB = 'MANTENIMIENTO_DJI';

function _launchDocModeloHoja(ss, nombreTab) {
  var hoja = ss.getSheetByName(nombreTab);
  if (!hoja) {
    hoja = ss.insertSheet(nombreTab);
    hoja.appendRow(['Modelo', 'Link', 'Activo']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 3).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(2, 360);
  }
  return hoja;
}
function _launchDocModeloGet(nombreTab) {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchDocModeloHoja(ss, nombreTab);
    var d    = hoja.getDataRange().getValues();
    var out  = [];
    for (var i = 1; i < d.length; i++) {
      var modelo = String(d[i][0] || '').trim();
      if (!modelo) continue;
      out.push({ modelo: modelo, link: String(d[i][1] || '').trim(), activo: _launchEvActivo(d[i][2]) });
    }
    out.sort(function(a, b) { return a.modelo.localeCompare(b.modelo); });
    return { ok: true, items: out };
  } catch(e) { Logger.log('_launchDocModeloGet(' + nombreTab + '): ' + e); return { ok: false, error: e.toString(), items: [] }; }
}
// data = { modelo, link, activo, modeloOriginal } — mismo criterio de upsert por clave que
// LAUNCH_saveCursoMaterial (acá la clave es el Modelo, renombrable vía modeloOriginal).
function _launchDocModeloSave(nombreTab, data) {
  try {
    data = data || {};
    var modelo = String(data.modelo || '').trim();
    if (!modelo) return { ok: false, error: 'Falta el modelo.' };
    var link = String(data.link || '').trim();
    if (!link) return { ok: false, error: 'Cargá el link.' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchDocModeloHoja(ss, nombreTab);
    var fila = [modelo, link, data.activo === false ? 'NO' : 'SI'];
    var buscar = String(data.modeloOriginal || modelo).trim().toLowerCase();
    var d = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() !== buscar) continue;
      filaIdx = i + 1; break;
    }
    if (filaIdx > 0) hoja.getRange(filaIdx, 1, 1, 3).setValues([fila]);
    else             hoja.appendRow(fila);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('_launchDocModeloSave(' + nombreTab + '): ' + e); return { ok: false, error: e.toString() }; }
}
function _launchDocModeloEliminar(nombreTab, modelo) {
  try {
    modelo = String(modelo || '').trim();
    if (!modelo) return { ok: false, error: 'Falta el modelo.' };
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchDocModeloHoja(ss, nombreTab);
    var d = hoja.getDataRange().getValues();
    var modeloLower = modelo.toLowerCase();
    var fila = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toLowerCase() === modeloLower) { fila = i + 1; break; }
    }
    if (fila < 0) return { ok: false, error: 'No se encontró ese modelo.' };
    hoja.deleteRow(fila);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('_launchDocModeloEliminar(' + nombreTab + '): ' + e); return { ok: false, error: e.toString() }; }
}

function LAUNCH_getRepuestosRecomendados()        { return _launchDocModeloGet(_LAUNCH_REPUESTOS_REC_TAB); }
function LAUNCH_saveRepuestoRecomendado(data)      { return _launchDocModeloSave(_LAUNCH_REPUESTOS_REC_TAB, data); }
function LAUNCH_eliminarRepuestoRecomendado(modelo){ return _launchDocModeloEliminar(_LAUNCH_REPUESTOS_REC_TAB, modelo); }

function LAUNCH_getMantenimientoDji()        { return _launchDocModeloGet(_LAUNCH_MANTENIMIENTO_TAB); }
function LAUNCH_saveMantenimientoDji(data)   { return _launchDocModeloSave(_LAUNCH_MANTENIMIENTO_TAB, data); }
function LAUNCH_eliminarMantenimientoDji(modelo) { return _launchDocModeloEliminar(_LAUNCH_MANTENIMIENTO_TAB, modelo); }


// ── "Pedir algo a BidcomAgro" — solicitudes libres que mandan los resellers desde "Contacto y
// soporte" del Portal (RS_Solicitudes.js). No es un catálogo para editar (nombre+link) como
// el resto de este archivo — es una cola de pedidos para marcar resueltos, así que no tiene
// un "guardar" genérico tipo _launchDocModeloSave: cada card de la lista responde directo.
var _LAUNCH_SOLICITUDES_TAB = 'SOLICITUDES_RESELLER';

function _launchSolicitudesHoja(ss) {
  var hoja = ss.getSheetByName(_LAUNCH_SOLICITUDES_TAB);
  if (!hoja) {
    hoja = ss.insertSheet(_LAUNCH_SOLICITUDES_TAB);
    hoja.appendRow(['ID', 'Fecha', 'Reseller', 'Asunto', 'Detalle', 'Estado', 'Respuesta', 'FechaRespuesta']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 8).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
    hoja.setColumnWidth(4, 220);
    hoja.setColumnWidth(5, 320);
    hoja.setColumnWidth(7, 320);
  }
  return hoja;
}

// Todas las solicitudes de todos los resellers — Pendientes primero, y dentro de cada grupo
// (Pendiente/Resuelto) las más recientes arriba.
function LAUNCH_getSolicitudes() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchSolicitudesHoja(ss);
    var d    = hoja.getDataRange().getValues();
    var out  = [];
    for (var i = 1; i < d.length; i++) {
      var id = String(d[i][0] || '').trim();
      if (!id) continue;
      var fechaCell = d[i][1];
      out.push({
        id:         id,
        fecha:      fechaCell instanceof Date ? Utilities.formatDate(fechaCell, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(fechaCell || ''),
        fechaOrden: fechaCell instanceof Date ? fechaCell.getTime() : 0,
        reseller:   String(d[i][2] || '').trim(),
        asunto:     String(d[i][3] || '').trim(),
        detalle:    String(d[i][4] || '').trim(),
        estado:     String(d[i][5] || 'Pendiente').trim(),
        respuesta:  String(d[i][6] || '').trim()
      });
    }
    out.sort(function(a, b) {
      var aPend = a.estado !== 'Resuelto', bPend = b.estado !== 'Resuelto';
      if (aPend !== bPend) return aPend ? -1 : 1;
      return b.fechaOrden - a.fechaOrden;
    });
    return { ok: true, items: out };
  } catch(e) { Logger.log('LAUNCH_getSolicitudes: ' + e); return { ok: false, error: e.toString(), items: [] }; }
}

// data = { id, estado, respuesta } — estado: 'Resuelto' o 'Pendiente' (reabrir una ya resuelta).
function LAUNCH_responderSolicitud(data) {
  try {
    data = data || {};
    var id = String(data.id || '').trim();
    if (!id) return { ok: false, error: 'Falta el ID de la solicitud.' };
    var estado = data.estado === 'Pendiente' ? 'Pendiente' : 'Resuelto';
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchSolicitudesHoja(ss);
    var d = hoja.getDataRange().getValues();
    var filaIdx = -1;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim() === id) { filaIdx = i + 1; break; }
    }
    if (filaIdx < 0) return { ok: false, error: 'No se encontró esa solicitud.' };
    var respuesta = String(data.respuesta || '').trim();
    hoja.getRange(filaIdx, 6, 1, 3).setValues([[estado, respuesta, estado === 'Resuelto' ? new Date() : '']]);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(e) { Logger.log('LAUNCH_responderSolicitud: ' + e); return { ok: false, error: e.toString() }; }
}

// ── Encuestas de Satisfacción Postventa (Portal Reseller) — solo lectura, sin acción de
// responder: es el reporte de lo que ya contestaron los resellers, no una cola para resolver.
var _LAUNCH_ENCUESTAS_TAB = 'ENCUESTA_POSTVENTA';
function _launchEncuestasHoja(ss) {
  var hoja = ss.getSheetByName(_LAUNCH_ENCUESTAS_TAB);
  if (!hoja) {
    hoja = ss.insertSheet(_LAUNCH_ENCUESTAS_TAB);
    hoja.appendRow(['ID', 'Fecha', 'Reseller', 'PuntPortal', 'PuntGarantia', 'PuntRepuesto']);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, 6).setBackground('#00a3e0').setFontColor('#fff').setFontWeight('bold');
  }
  return hoja;
}

// Todas las respuestas + los 3 promedios generales, ya calculados server-side.
function LAUNCH_getEncuestas() {
  try {
    var ss   = SpreadsheetApp.openById(MASTER_SS_ID);
    var hoja = _launchEncuestasHoja(ss);
    var d    = hoja.getDataRange().getValues();
    var out  = [];
    var sumaPortal = 0, sumaGarantia = 0, sumaRepuesto = 0;
    for (var i = 1; i < d.length; i++) {
      var id = String(d[i][0] || '').trim();
      if (!id) continue;
      var fechaCell   = d[i][1];
      var pPortal     = Number(d[i][3]) || 0;
      var pGarantia   = Number(d[i][4]) || 0;
      var pRepuesto   = Number(d[i][5]) || 0;
      sumaPortal += pPortal; sumaGarantia += pGarantia; sumaRepuesto += pRepuesto;
      out.push({
        id:           id,
        fecha:        fechaCell instanceof Date ? Utilities.formatDate(fechaCell, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(fechaCell || ''),
        fechaOrden:   fechaCell instanceof Date ? fechaCell.getTime() : 0,
        reseller:     String(d[i][2] || '').trim(),
        puntPortal:   pPortal,
        puntGarantia: pGarantia,
        puntRepuesto: pRepuesto
      });
    }
    out.sort(function(a, b) { return b.fechaOrden - a.fechaOrden; });
    var n = out.length;
    return {
      ok: true, items: out, total: n,
      promPortal:   n ? Math.round((sumaPortal   / n) * 10) / 10 : 0,
      promGarantia: n ? Math.round((sumaGarantia / n) * 10) / 10 : 0,
      promRepuesto: n ? Math.round((sumaRepuesto / n) * 10) / 10 : 0
    };
  } catch(e) {
    Logger.log('LAUNCH_getEncuestas: ' + e);
    return { ok: false, error: e.toString(), items: [], total: 0, promPortal: 0, promGarantia: 0, promRepuesto: 0 };
  }
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
