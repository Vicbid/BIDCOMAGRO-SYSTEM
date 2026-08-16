// @version 1.1
// ============================================================
//  WOS — Calidad / Precisión (armado guiado)
//  Extraído de Despacho_Código.js 3.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// ══════════════════════════════════════════════════════════════
//  CALIDAD / PRECISIÓN (armado guiado)
//  Hoja WOS_QA: 1 fila por cierre de caja + confirmación del reseller.
// ══════════════════════════════════════════════════════════════
var WOS_QA_SHEET = 'WOS_QA';

function _wosGetHojaQA() {
  var ss   = SpreadsheetApp.openById(NOTAS_SS_ID);
  var hoja = ss.getSheetByName(WOS_QA_SHEET);
  if (!hoja) {
    hoja = ss.insertSheet(WOS_QA_SHEET);
    hoja.getRange(1, 1, 1, 15).setValues([[
      'Fecha', 'Pedido', 'Reseller', 'Operario', 'Items', 'Unidades',
      'CantOk', 'SnDupOk', 'EnvioOk', 'Overrides', 'Resultado', 'TiempoSeg',
      'Confirmacion', 'FechaConfirmacion', 'NotaProblema'
    ]]).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}


// Chequea N° de serie duplicados: internos (en la lista que se va a despachar) y
// externos (contra SERIALES ya despachados en OTROS pedidos). Devuelve la lista de duplicados.
function WOS_verificarSnDuplicados(numero, seriales) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    numero = String(numero || '').trim();
    var lista = (Object.prototype.toString.call(seriales) === '[object Array]')
      ? seriales : String(seriales || '').split(',');
    var norm = [];
    for (var i = 0; i < lista.length; i++) {
      // quitar sufijo " xN" de bolsas → el código en sí es lo único
      var s = String(lista[i] || '').trim().replace(/\s+x\d+$/i, '').toUpperCase();
      if (s) norm.push(s);
    }
    // 1. Duplicados internos
    var seen = {}, internos = [];
    for (var j = 0; j < norm.length; j++) {
      if (seen[norm[j]]) { if (internos.indexOf(norm[j]) === -1) internos.push(norm[j]); }
      else seen[norm[j]] = true;
    }
    // 2. Duplicados contra otros pedidos ya despachados
    var externos = [];
    var hojas = [_getHojaPedidos(), _getHojaPedidosOT()].filter(Boolean);
    for (var h = 0; h < hojas.length; h++) {
      var datos = hojas[h].getDataRange().getValues();
      for (var r = 1; r < datos.length; r++) {
        if (String(datos[r][COL.NUMERO] || '').trim() === numero) continue; // no chequear contra sí mismo
        var serStr = String(datos[r][COL.SERIALES] || '').trim();
        if (!serStr) continue;
        var partes = serStr.split(',');
        for (var m = 0; m < partes.length; m++) {
          var code = partes[m].trim().replace(/\s+x\d+$/i, '').toUpperCase();
          if (code && seen[code] && externos.indexOf(code) === -1) externos.push(code);
        }
      }
    }
    var dups = internos.slice();
    for (var e2 = 0; e2 < externos.length; e2++) if (dups.indexOf(externos[e2]) === -1) dups.push(externos[e2]);
    return { ok: true, hayDuplicados: dups.length > 0, duplicados: dups, internos: internos, externos: externos };
  } catch(e) {
    Logger.log('WOS_verificarSnDuplicados: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Registra el resultado de un cierre de caja guiado. qa: {reseller, operario, items, unidades,
// cantOk, snDupOk, envioOk, overrides, resultado, tiempoSeg}. Best-effort (no rompe el despacho).
function WOS_registrarQA(numero, qa) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    qa = qa || {};
    var hoja = _wosGetHojaQA();
    var b = function(v) { return v === false ? 'NO' : 'S\xcd'; };
    var overridesStr = qa.overrides
      ? (typeof qa.overrides === 'string' ? qa.overrides : JSON.stringify(qa.overrides)) : '';
    var hayOverride = overridesStr && overridesStr !== '[]' && overridesStr !== '{}' && overridesStr !== '""';
    var resultado = qa.resultado ||
      ((qa.cantOk !== false && qa.snDupOk !== false && qa.envioOk !== false && !hayOverride) ? 'limpio' : 'flag');
    hoja.appendRow([
      new Date(), String(numero || ''), _antiFormula(qa.reseller || ''), _antiFormula(qa.operario || ''),
      Number(qa.items) || 0, Number(qa.unidades) || 0,
      b(qa.cantOk), b(qa.snDupOk), b(qa.envioOk),
      overridesStr, resultado, Number(qa.tiempoSeg) || 0,
      '', '', ''
    ]);
    return { ok: true, resultado: resultado };
  } catch(e) {
    Logger.log('WOS_registrarQA: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Actualiza la fila de WOS_QA del pedido con la confirmación del reseller (ok/problema).
// Si no hay fila (despacho manual viejo), crea una mínima para no perder el dato de resultado.
function _wosRegistrarConfirmacion(numero, reseller, r, nota) {
  try {
    var conf  = (r === 'ok') ? 'ok' : 'problema';
    var hoja  = _wosGetHojaQA();
    var datos = hoja.getDataRange().getValues();
    var target = -1, fallback = -1;
    for (var i = datos.length - 1; i >= 1; i--) {
      if (String(datos[i][1] || '').trim() !== numero) continue;
      if (fallback === -1) fallback = i;
      if (!String(datos[i][12] || '').trim()) { target = i; break; }  // col 13 (idx 12) = Confirmacion
    }
    if (target === -1) target = fallback;
    if (target >= 1) {
      hoja.getRange(target + 1, 13).setValue(conf);
      hoja.getRange(target + 1, 14).setValue(new Date());
      if (nota) hoja.getRange(target + 1, 15).setValue(_antiFormula(nota));
    } else {
      hoja.appendRow([new Date(), numero, _antiFormula(reseller || ''), '', 0, 0, '', '', '', '', '', 0, conf, new Date(), _antiFormula(nota || '')]);
    }
  } catch(e) { Logger.log('_wosRegistrarConfirmacion: ' + e); }
}


// Métricas de precisión combinadas (proceso + resultado del reseller), global y por operador.
function WOS_getPrecisionMetrics(desdeISO, hastaISO) {
  var u = WOS_getUsuario();
  if (!u.autorizado) return { ok: false, error: 'No autorizado.' };
  try {
    var hoja  = _wosGetHojaQA();
    var datos = hoja.getDataRange().getValues();
    var desde = desdeISO ? new Date(desdeISO) : null;
    var hasta = hastaISO ? new Date(hastaISO) : null;
    var totalCierres = 0, limpios = 0, conf = 0, confOk = 0, preciso = 0;
    var porOp = {}, flags = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i][0];
      if (!(f instanceof Date)) continue;
      if (desde && f < desde) continue;
      if (hasta && f > hasta) continue;
      var pedido    = String(datos[i][1] || '');
      var op        = String(datos[i][3] || '') || '(sin operador)';
      var resultado = String(datos[i][10] || '').trim();   // col 11
      var confir    = String(datos[i][12] || '').trim();   // col 13
      if (!porOp[op]) porOp[op] = { cierres: 0, limpios: 0, conf: 0, confOk: 0, preciso: 0 };
      if (resultado) {   // fila con Resultado = cierre de caja guiado
        totalCierres++; porOp[op].cierres++;
        var esLimpio = (resultado === 'limpio');
        if (esLimpio) { limpios++; porOp[op].limpios++; }
        if (esLimpio && confir !== 'problema') { preciso++; porOp[op].preciso++; }
        if (!esLimpio || confir === 'problema') {
          flags.push({ pedido: pedido, operario: op, fechaMs: f.getTime(), resultado: resultado,
                       confirmacion: confir, overrides: String(datos[i][9] || ''), nota: String(datos[i][14] || '') });
        }
      }
      if (confir === 'ok' || confir === 'problema') {
        conf++; porOp[op].conf++;
        if (confir === 'ok') { confOk++; porOp[op].confOk++; }
      }
    }
    var pct = function(a, b) { return b > 0 ? Math.round((a / b) * 1000) / 10 : null; };
    var ranking = [];
    for (var o in porOp) if (porOp.hasOwnProperty(o)) {
      ranking.push({
        operario:     o,
        cierres:      porOp[o].cierres,
        procesoPct:   pct(porOp[o].limpios, porOp[o].cierres),
        resultadoPct: pct(porOp[o].confOk, porOp[o].conf),
        totalPct:     pct(porOp[o].preciso, porOp[o].cierres)
      });
    }
    ranking.sort(function(a, b) { return (b.totalPct || 0) - (a.totalPct || 0); });
    flags.sort(function(a, b) { return b.fechaMs - a.fechaMs; });
    var tz = Session.getScriptTimeZone();
    flags = flags.slice(0, 30).map(function(fl) {
      return { pedido: fl.pedido, operario: fl.operario, resultado: fl.resultado, confirmacion: fl.confirmacion,
               overrides: fl.overrides, nota: fl.nota,
               fecha: Utilities.formatDate(new Date(fl.fechaMs), tz, 'dd/MM HH:mm') };
    });
    return {
      ok: true,
      totalCierres: totalCierres,
      procesoPct:   pct(limpios, totalCierres),
      confRespondidos: conf,
      resultadoPct: pct(confOk, conf),
      totalPct:     pct(preciso, totalCierres),
      ranking: ranking,
      flags: flags
    };
  } catch(e) {
    Logger.log('WOS_getPrecisionMetrics: ' + e);
    return { ok: false, error: 'No se pudo procesar la solicitud. Intentá de nuevo.' };
  }
}


// Helper para probar desde el editor.
function WOS_previewPrecision() {
  var r = WOS_getPrecisionMetrics(null, null);
  Logger.log('Precisi\xf3n: ' + JSON.stringify(r, null, 2));
  return r;
}
