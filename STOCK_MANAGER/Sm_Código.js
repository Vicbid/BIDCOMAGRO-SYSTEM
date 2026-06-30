// ============================================================
//  STOCK MANAGER BIDCOMAGRO v4.1 — SM_Codigo.gs
//  Proyecto: Stock Manager
//
//  Comparte el mismo Google Sheet que HUB PRO y Portal.
//  Lee y escribe en hojas: STOCK_REPUESTOS, MOVIMIENTOS_STOCK,
//  COMPRAS_DJI, VENTAS_DIRECTAS, SOLICITUDES_DESPACHO,
//  DB_REPUESTOS, Resellers, Ordenes de trabajo
// ============================================================

var MASTER_SHEET_ID = "1YeQl4vTQ5pTFahZ8Z9Jab7rP42xFD4_hEvpW_JDXjRc";

var SM_CONFIG = {
  EMAIL_SUPERVISOR:  "soporteagrasdji@bidcom.com.ar",
  NOMBRE_REMITENTE:  "BIDCOMAGRO · Stock Manager",
  DIAS_STOCK_RECOM:  90,
  ALERTA_CATEGORIAS: ["A","B"]
};


// SINGLETON
var _SS = null;
function getSS() {
  if (!_SS) _SS = SpreadsheetApp.openById(MASTER_SHEET_ID);
  return _SS;
}
function getSheet(nombre) { return getSS().getSheetByName(nombre); }

// ── CARMEN — spreadsheet externo, fuente de verdad para stock actual ──
var CARMEN_SS_ID         = '1-BH5m-LXFYhBZxqpSFVhIz5jwzFgJmLWH8Qvkh4PSCI';
var _carmenSS            = null;
var _carmenStockMapCache = null;

function _getCarmenSS() {
  if (!_carmenSS) _carmenSS = SpreadsheetApp.openById(CARMEN_SS_ID);
  return _carmenSS;
}

// Retorna { SKU → stockActual } desde hoja "STOCK" de Carmen (col A=codigo, col C=stock)
function _getCarmenStockMap() {
  if (_carmenStockMapCache) return _carmenStockMapCache;
  try {
    var hoja = _getCarmenSS().getSheetByName('STOCK');
    if (!hoja) return {};
    var d = hoja.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var c = String(d[i][0] || '').trim().toUpperCase();
      if (c) m[c] = parseInt(d[i][2]) || 0;
    }
    _carmenStockMapCache = m;
    return m;
  } catch(e) { Logger.log('_getCarmenStockMap: ' + e); return {}; }
}

// Retorna { SKU → [{ubicacion, cantidad}] } desde tab UBICACIONES de Carmen
function _getCarmenUbicMap() {
  var m = {};
  try {
    var hoja = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hoja) return m;
    var d = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      var sku  = String(d[i][0] || '').trim().toUpperCase();
      var ubic = String(d[i][1] || '').trim();
      var cant = parseInt(d[i][2]) || 0;
      if (!sku || !ubic) continue;
      if (!m[sku]) m[sku] = [];
      m[sku].push({ ubicacion: ubic, cantidad: cant });
    }
  } catch(e) { Logger.log('_getCarmenUbicMap: ' + e); }
  return m;
}

// Registra un movimiento en Carmen (Entregados/Recibidos) y actualiza tally en UBICACIONES
// diff > 0 → entrada (Recibidos); diff < 0 → salida (Entregados)
// UBICACIONES col C = número directo (tally); STOCK col C de Carmen = fórmula de Carmen, nunca se toca
function _registrarMovimientoCarmen(sku, desc, ubicacion, diff, referencia) {
  try {
    var ss      = _getCarmenSS();
    var codKey  = String(sku       || '').trim().toUpperCase();
    var ubicKey = String(ubicacion || '').trim();
    var refStr  = String(referencia || 'Ajuste').trim();
    var descStr = String(desc      || '').trim();
    var cant    = Math.abs(diff);
    var fecha   = new Date();

    if (diff < 0) {
      var hojaEnt = ss.getSheetByName(CARMEN_ENTREGADOS_TAB);
      if (hojaEnt) hojaEnt.appendRow([codKey, descStr, cant, refStr, '', '', fecha, ubicKey]);
    } else if (diff > 0) {
      var hojaRec = ss.getSheetByName(CARMEN_RECIBIDOS_TAB);
      if (hojaRec) hojaRec.appendRow([codKey, descStr, cant, refStr, fecha, '', '', '', ubicKey]);
    }

    // Actualizar tally en UBICACIONES (col C = número directo, sin fórmula)
    if (ubicKey && diff !== 0) {
      var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
      if (hojaUbic) {
        var dU = hojaUbic.getDataRange().getValues();
        for (var i = 1; i < dU.length; i++) {
          if (String(dU[i][0] || '').trim().toUpperCase() === codKey &&
              String(dU[i][1] || '').trim().toUpperCase() === ubicKey) {
            var nueva = Math.max(0, (parseFloat(dU[i][2]) || 0) + diff);
            hojaUbic.getRange(i + 1, 3).setValue(nueva);
            return;
          }
        }
        // Fila nueva
        hojaUbic.appendRow([codKey, ubicKey, Math.max(0, diff)]);
      }
    }
  } catch(e) { Logger.log('_registrarMovimientoCarmen: ' + e); }
}

// Actualiza col C (stock actual) de un SKU en la hoja "STOCK" de Carmen
function _actualizarCarmenStock(sku, nuevoSaldo) {
  try {
    var hoja    = _getCarmenSS().getSheetByName('STOCK');
    if (!hoja) return;
    var codBusc = String(sku).trim().toUpperCase();
    var d       = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() === codBusc) {
        hoja.getRange(i + 1, 3).setValue(nuevoSaldo);
        _carmenStockMapCache = null;
        return;
      }
    }
    Logger.log('_actualizarCarmenStock: SKU no encontrado en Carmen: ' + sku);
  } catch(e) { Logger.log('_actualizarCarmenStock: ' + e); }
}

// Restaura la fórmula de Carmen STOCK col C en todas las filas que tienen un número plano.
// Ejecutar UNA VEZ desde el editor de Apps Script después de actualizar el código.
// La fórmula asume: col A = SKU, col F = stock base, pivots 'Entreg dinámica' y 'Recib dinámica'.
function restaurarFormulasCarmenStock() {
  try {
    var hoja = _getCarmenSS().getSheetByName('STOCK');
    if (!hoja) return { ok: false, error: 'Hoja STOCK no encontrada en Carmen' };
    var lastRow = hoja.getLastRow();
    if (lastRow < 2) return { ok: true, restauradas: 0 };
    var formulas  = hoja.getRange(2, 3, lastRow - 1, 1).getFormulas();
    var newForms  = [];
    var restauradas = 0;
    for (var i = 0; i < formulas.length; i++) {
      var row = i + 2;
      if (!formulas[i][0]) {
        // Celda sin fórmula (número plano) — restaurar
        newForms.push(['=F' + row + '-(SI.ERROR(BUSCARV(A' + row + ',\'Entreg dinámica\'!A:B,2,0),0))+(SI.ERROR(BUSCARV(A' + row + ',\'Recib dinámica\'!A:B,2,0),0))']);
        restauradas++;
      } else {
        newForms.push([formulas[i][0]]); // ya tiene fórmula, no tocar
      }
    }
    hoja.getRange(2, 3, lastRow - 1, 1).setFormulas(newForms);
    SpreadsheetApp.flush();
    Logger.log('restaurarFormulasCarmenStock: ' + restauradas + ' filas restauradas de ' + (lastRow - 1) + ' totales');
    return { ok: true, restauradas: restauradas, total: lastRow - 1 };
  } catch(e) {
    Logger.log('restaurarFormulasCarmenStock ERROR: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ── ASEGURAR HOJAS ───────────────────────────────────────────
function asegurarHojas() {
  var ss = getSS();
  var hojas = {
    "STOCK_REPUESTOS":    ["Código","Descripción","Stock actual","Stock mínimo","Categoría","Ubicación","Modelos compatibles","Última entrada","Última salida"],
    "MOVIMIENTOS_STOCK":  ["Fecha","Tipo","Código","Descripción","Cantidad","Stock resultante","Referencia","Operador","Observaciones","Depósito"],
    "COMPRAS_DJI":        ["CAS","Fecha pedido","Estado","Método pago","Fecha comprado","Fecha pagado","Fecha conf. envío","Fecha forwarder HK","Fecha vuelo","Fecha aduana","Fecha depósito","Operador","Observaciones","Última actualización"],
    "HISTORIAL_COMPRAS":  ["Fecha","ID_CAS","Estado anterior","Estado nuevo","Operador","Observaciones"],
    "CATALOGO_DJI":       ["Material Number","Simplified part number","代理商系统名称优化","English name","Applicable models","unit","pcs","South America"],
    "VENTAS_DIRECTAS":    ["Fecha","N° Orden entrega","N° Factura","Reseller","Código","Descripción","Cantidad","Precio USD","Observaciones"],
    "SOLICITUDES_DESPACHO": ["ID","Fecha","OT","Reseller","Código","Descripción","Cant. solicitada","Cant. despachada","Estado","Urgencia","Fecha despacho","Operador","Observaciones"],
    "STOCK_UBICACIONES":    ["SKU","Ubicación","Cantidad"],
    "TABLA_POSICIONES":     ["SKU","BIN_ID","CANTIDAD","TIPO_ALMACEN"],
    "LAYOUT_ALMACEN":       ["PASILLO","ORDEN_PASILLO","ESTANTE","ORDEN_ESTANTE","NUM_NIVELES"]
  };
  var keys = Object.keys(hojas);
  for (var i = 0; i < keys.length; i++) {
    var nombre = keys[i];
    if (!ss.getSheetByName(nombre)) {
      var h = ss.insertSheet(nombre);
      h.appendRow(hojas[nombre]);
      h.getRange(1, 1, 1, hojas[nombre].length).setFontWeight("bold");
    }
  }
}

// ── ENTRY POINT ──────────────────────────────────────────────
function doGet() {
  asegurarHojas();
  return HtmlService.createHtmlOutputFromFile('SM_Index')
    .setTitle('Stock Manager - BIDCOMAGRO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
//  CARGA INICIAL — dashboard y catálogo
// ============================================================
// Limpia categorías numéricas en STOCK_REPUESTOS y DB_REPUESTOS.
// Se llama automáticamente al cargar el dashboard la primera vez después de un deploy.
function _limpiarCatsNumericas() {
  try {
    var props = PropertiesService.getScriptProperties();
    if (props.getProperty('cats_limpias_v2')) return;  // ya corrió
    var changed = false;
    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    if (hojaStock) {
      var dS = hojaStock.getDataRange().getValues();
      for (var i = 1; i < dS.length; i++) {
        if (dS[i][4] && !_catValida(String(dS[i][4]))) { dS[i][4] = ''; changed = true; }
      }
      if (changed) { hojaStock.getDataRange().setValues(dS); invalidateSheetValues(SCHEMA.SHEETS.STOCK); }
    }
    var repChanged = false;
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaRep) {
      var dR = hojaRep.getDataRange().getValues();
      for (var j = 1; j < dR.length; j++) {
        if (dR[j][5] && !_catValida(String(dR[j][5]))) { dR[j][5] = ''; repChanged = true; }
      }
      if (repChanged) { hojaRep.getDataRange().setValues(dR); invalidateSheetValues(SCHEMA.SHEETS.DB_REPUESTOS); }
    }
    props.setProperty('cats_limpias_v2', '1');
  } catch(e) { Logger.log('_limpiarCatsNumericas: ' + e); }
}

function cargarDashboard() {
  try {
    _limpiarCatsNumericas();
    var hoy = new Date();
    var ss = getSS();
    
    // Leemos todo el libro de una vez para minimizar llamadas al servidor
    var hojas = [SCHEMA.SHEETS.STOCK, SCHEMA.SHEETS.COMPRAS, SCHEMA.SHEETS.DB_REPUESTOS, SCHEMA.SHEETS.MOVIMIENTOS];
    var db = {};
    hojas.forEach(function(h) {
      var sheet = ss.getSheetByName(h);
      db[h] = sheet ? getSheetValues(sheet) : [];
    });

    // Carmen — fuente de verdad para stock actual
    var carmenMap = _getCarmenStockMap();

    // Mapa de precios FOB desde DB_REPUESTOS (col 1 = código, col 6 = precio FOB)
    var precioFOB = {};
    var dRep = db[SCHEMA.SHEETS.DB_REPUESTOS];
    for (var r = 1; r < dRep.length; r++) {
      var codR = String(dRep[r][1] || '').trim().toUpperCase();
      if (codR) precioFOB[codR] = parseFloat(String(dRep[r][6] || '0').replace(',', '.')) || 0;
    }

    // KPIs de stock (Procesado en memoria)
    var bajoMinimo = { A:0, B:0, C:0, D:0, total:0, quiebre:0 };
    var valorDeposito = 0;
    var alertasCriticas = [];
    var dStock = db[SCHEMA.SHEETS.STOCK];
    var S = SCHEMA.STOCK_REPUESTOS;

    // Mapa de OTs bloqueadas en "Espera de repuestos" por SKU
    var otsEsperaPorSku = {};
    try {
      var dOT = getSheetValues(SCHEMA.SHEETS.OT);
      for (var oi = 1; oi < dOT.length; oi++) {
        if (String(dOT[oi][SCHEMA.OT.ESTADO]||"") !== "Espera de repuestos") continue;
        var repStr = String(dOT[oi][SCHEMA.OT.REPUESTOS]||"").trim();
        if (!repStr || repStr === "Sin consumo de repuestos") continue;
        var ps = repStr.split(" ; ");
        for (var pi = 0; pi < ps.length; pi++) {
          var skuOT = String(ps[pi].split(" | ")[0]).trim().toUpperCase();
          if (skuOT) otsEsperaPorSku[skuOT] = (otsEsperaPorSku[skuOT]||0) + 1;
        }
      }
    } catch(e3) {}

    for (var i = 1; i < dStock.length; i++) {
      if (!dStock[i][S.CODIGO]) continue;
      var skuD = String(dStock[i][S.CODIGO]).trim().toUpperCase();
      var actual = (carmenMap[skuD] !== undefined) ? carmenMap[skuD] : (parseInt(dStock[i][S.STOCK_ACTUAL]) || 0);
      var minimo = parseInt(dStock[i][S.STOCK_MINIMO]) || 0;
      var cat = String(dStock[i][S.CATEGORIA] || '').trim().toUpperCase();

      var fob = precioFOB[skuD] || 0;
      valorDeposito += fob * actual;

      if (actual <= minimo) {
        bajoMinimo.total++;
        if (actual === 0) bajoMinimo.quiebre++;
        if (bajoMinimo[cat] !== undefined) bajoMinimo[cat]++;
        if (cat === 'A' || cat === 'B') {
          alertasCriticas.push({
            codigo:      skuD,
            descripcion: String(dStock[i][S.DESCRIPCION]),
            stockActual: actual,
            stockMinimo: minimo,
            categoria:   cat,
            ubicacion:   String(dStock[i][S.UBICACION] || ''),
            estado:      (actual === 0) ? 'QUIEBRE' : 'BAJO',
            bloqueadas:  otsEsperaPorSku[skuD] || 0
          });
        }
      }
    }
    // Ordenar alertas: quiebres primero, luego por OTs bloqueadas
    alertasCriticas.sort(function(a, b) {
      if (a.estado !== b.estado) return a.estado === 'QUIEBRE' ? -1 : 1;
      return b.bloqueadas - a.bloqueadas;
    });

    // Pedidos pendientes de despacho — fusiona Pedidos_resellers + Pedidos_OTs del WOS
    // COL: 0=NUMERO, 1=RESELLER, 6=CANT_PEND, 9=ESTADO, 10=FECHA
    var solicPendientes = [];
    try {
      var wosSS = SpreadsheetApp.openById(WOS_NOTAS_SS_ID);
      var wosHojas = [wosSS.getSheetByName('Pedidos_resellers'), wosSS.getSheetByName('Pedidos_OTs')].filter(Boolean);
      var ESTADOS_CERRADOS_WOS = ['Entregado_Cerrado', 'Cancelado', 'Entregado_Confirmado'];
      var pedMap = {};
      for (var wh = 0; wh < wosHojas.length; wh++) {
        var wosD = wosHojas[wh].getDataRange().getValues();
        for (var wp = 1; wp < wosD.length; wp++) {
          var wNum  = String(wosD[wp][0] || '').trim();
          var wEst  = String(wosD[wp][9] || '').trim();
          if (!wNum || ESTADOS_CERRADOS_WOS.indexOf(wEst) !== -1) continue;
          var wPend = Number(wosD[wp][6]) || 0;
          if (wPend <= 0) continue;
          var wFec  = wosD[wp][10];
          if (!pedMap[wNum]) {
            pedMap[wNum] = {
              numero:   wNum,
              reseller: String(wosD[wp][1] || ''),
              estado:   wEst,
              cantPend: 0,
              fecha:    (wFec instanceof Date) ? wFec : null
            };
          }
          pedMap[wNum].cantPend += wPend;
        }
      }
      var wKeys = Object.keys(pedMap);
      for (var wk = 0; wk < wKeys.length; wk++) {
        var ped = pedMap[wKeys[wk]];
        solicPendientes.push({
          numero:     ped.numero,
          reseller:   ped.reseller,
          estado:     ped.estado,
          cantPend:   ped.cantPend,
          diasEspera: ped.fecha ? Math.floor((hoy - ped.fecha) / 86400000) : 0
        });
      }
      solicPendientes.sort(function(a, b) { return b.diasEspera - a.diasEspera; });
    } catch(eWP) { Logger.log('cargarDashboard WOS pedidos: ' + eWP); }

    // Métricas logísticas desde COMPRAS_DJI (ya cargado en db, sin round-trip extra)
    var dCom = db[SCHEMA.SHEETS.COMPRAS];
    var sumLead = 0, cntLead = 0, sumAduana = 0, cntAduana = 0;
    for (var c = 1; c < dCom.length; c++) {
      var fc = dCom[c];
      var fPedido   = fc[1]  instanceof Date ? fc[1]  : null;
      var fAduanaC  = fc[9]  instanceof Date ? fc[9]  : null;
      var fDepositoC = fc[10] instanceof Date ? fc[10] : null;
      if (fPedido   && fDepositoC) { sumLead   += Math.floor((fDepositoC - fPedido)   / 86400000); cntLead++; }
      if (fAduanaC  && fDepositoC) { sumAduana += Math.floor((fDepositoC - fAduanaC)  / 86400000); cntAduana++; }
    }
    var totalSKU = dStock.length > 1 ? dStock.length - 1 : 0;

    // Pronóstico de quiebre por burn rate (últimos 30 días de MOVIMIENTOS)
    var LEAD_DIAS   = cntLead ? Math.round(sumLead / cntLead) : 45;
    var dMov        = db[SCHEMA.SHEETS.MOVIMIENTOS] || [];
    var M           = SCHEMA.MOVIMIENTOS_STOCK;
    var corte30     = new Date(hoy.getTime() - 30 * 86400000);
    var consumo30   = {};
    var lastMov     = {};
    for (var mv = 1; mv < dMov.length; mv++) {
      var fm = dMov[mv];
      if (!(fm[M.FECHA] instanceof Date)) continue;
      var codMov = String(fm[M.CODIGO] || '').trim().toUpperCase();
      if (!codMov) continue;
      if (!lastMov[codMov] || fm[M.FECHA] > lastMov[codMov]) lastMov[codMov] = fm[M.FECHA];
      if (fm[M.FECHA] < corte30) continue;
      var tipoMov = String(fm[M.TIPO] || '');
      if (tipoMov !== 'SALIDA_DESPACHO' && tipoMov !== 'EGRESO') continue;
      consumo30[codMov] = (consumo30[codMov] || 0) + Math.abs(parseInt(fm[M.CANTIDAD]) || 0);
    }
    var stockMapPron = {};
    for (var sp2 = 1; sp2 < dStock.length; sp2++) {
      if (!dStock[sp2][S.CODIGO]) continue;
      var codPr = String(dStock[sp2][S.CODIGO]).trim().toUpperCase();
      stockMapPron[codPr] = {
        descripcion: String(dStock[sp2][S.DESCRIPCION] || ''),
        stockActual: (carmenMap[codPr] !== undefined) ? carmenMap[codPr] : (parseInt(dStock[sp2][S.STOCK_ACTUAL]) || 0)
      };
    }
    var pronosticoQuiebre = [];
    var codsMov = Object.keys(consumo30);
    for (var pk = 0; pk < codsMov.length; pk++) {
      var codP    = codsMov[pk];
      var unids   = consumo30[codP];
      if (unids <= 0) continue;
      var info    = stockMapPron[codP];
      if (!info || info.stockActual <= 0) continue;
      var burnDay = unids / 30;
      var diasR   = Math.round(info.stockActual / burnDay);
      if (diasR >= LEAD_DIAS) continue;
      pronosticoQuiebre.push({ codigo: codP, descripcion: info.descripcion, stockActual: info.stockActual, unidades30d: unids, diasRestantes: diasR });
    }
    pronosticoQuiebre.sort(function(a, b) { return a.diasRestantes - b.diasRestantes; });

    // Top valor del inventario (precio FOB × stock)
    var topValor = [];
    for (var tv = 1; tv < dStock.length; tv++) {
      if (!dStock[tv][S.CODIGO]) continue;
      var codT  = String(dStock[tv][S.CODIGO]).trim().toUpperCase();
      var actT  = (carmenMap[codT] !== undefined) ? carmenMap[codT] : (parseInt(dStock[tv][S.STOCK_ACTUAL]) || 0);
      var fobT  = precioFOB[codT] || 0;
      if (fobT <= 0 || actT <= 0) continue;
      topValor.push({ codigo: codT, descripcion: String(dStock[tv][S.DESCRIPCION] || ''), stockActual: actT, precioFOB: fobT, valorTotal: Math.round(fobT * actT * 100) / 100 });
    }
    topValor.sort(function(a, b) { return b.valorTotal - a.valorTotal; });
    topValor = topValor.slice(0, 8);

    // Stock inmovilizado (sin movimientos en 90 días)
    var corte90      = new Date(hoy.getTime() - 90 * 86400000);
    var inmovilizado = [];
    for (var im = 1; im < dStock.length; im++) {
      if (!dStock[im][S.CODIGO]) continue;
      var codIM  = String(dStock[im][S.CODIGO]).trim().toUpperCase();
      var actIM  = (carmenMap[codIM] !== undefined) ? carmenMap[codIM] : (parseInt(dStock[im][S.STOCK_ACTUAL]) || 0);
      if (actIM <= 0) continue;
      var lastD  = lastMov[codIM];
      if (lastD && lastD >= corte90) continue;
      inmovilizado.push({
        codigo:     codIM,
        descripcion:String(dStock[im][S.DESCRIPCION] || ''),
        stockActual:actIM,
        diasSinMov: lastD ? Math.floor((hoy - lastD) / 86400000) : null
      });
    }
    inmovilizado.sort(function(a, b) {
      if (a.diasSinMov === null) return -1;
      if (b.diasSinMov === null) return 1;
      return b.diasSinMov - a.diasSinMov;
    });

    // Rotación mensual = salidas 30d / stock total actual
    var totalSalidas30 = 0;
    var codsMov2 = Object.keys(consumo30);
    for (var rm = 0; rm < codsMov2.length; rm++) totalSalidas30 += consumo30[codsMov2[rm]];
    var totalStockAct = 0;
    for (var rs = 1; rs < dStock.length; rs++) {
      var codRs = String(dStock[rs][S.CODIGO] || '').trim().toUpperCase();
      totalStockAct += (carmenMap[codRs] !== undefined) ? carmenMap[codRs] : (parseInt(dStock[rs][S.STOCK_ACTUAL]) || 0);
    }
    var rotacionMensual = (totalStockAct > 0) ? Math.round((totalSalidas30 / totalStockAct) * 100) / 100 : null;

    // CAS en tránsito — solo los que están físicamente en movimiento
    var CAS_ACTIVOS = ['En vuelo', 'En aduana'];
    var casTransito = [];
    for (var ct = 1; ct < dCom.length; ct++) {
      var fct    = dCom[ct];
      var casEst = String(fct[2] || '').trim();
      if (CAS_ACTIVOS.indexOf(casEst) === -1) continue;
      casTransito.push({
        cas:        String(fct[0] || ''),
        fechaPedido:_fmtFecha(fct[1]),
        estado:     casEst,
        metodoPago: String(fct[3] || '')
      });
    }

    // Cruzar pronóstico de quiebre con unidades en camino (COMPRAS_DETALLE)
    try {
      var casEstMap_d = {};
      for (var cip = 1; cip < dCom.length; cip++) {
        var casIdP  = String(dCom[cip][0] || '').trim().toUpperCase();
        var casEstP = String(dCom[cip][2] || '').trim();
        if (casIdP && casEstP !== 'En depósito' && casEstP.indexOf('Borrador') === -1)
          casEstMap_d[casIdP] = casEstP;
      }
      var enCaminoProno = {};
      var hojaCDP = getSS().getSheetByName(SCHEMA.SHEETS.COMPRAS_DETALLE);
      if (hojaCDP) {
        var dCDP = getSheetValues(hojaCDP);
        var CDP  = SCHEMA.COMPRAS_DETALLE;
        for (var cdp = 1; cdp < dCDP.length; cdp++) {
          var casP = String(dCDP[cdp][CDP.ID_CAS] || '').trim().toUpperCase();
          var skuP = String(dCDP[cdp][CDP.SKU]    || '').trim().toUpperCase();
          var pedP = parseInt(dCDP[cdp][CDP.CANTIDAD_PEDIDA])   || 0;
          var recP = parseInt(dCDP[cdp][CDP.CANTIDAD_RECIBIDA]) || 0;
          if (!skuP || !casEstMap_d[casP]) continue;
          var pendP = Math.max(0, pedP - recP);
          if (pendP > 0) {
            if (!enCaminoProno[skuP]) { enCaminoProno[skuP] = { total: 0, cas: casP, estado: casEstMap_d[casP] }; }
            enCaminoProno[skuP].total += pendP;
          }
        }
      }
      for (var prk = 0; prk < pronosticoQuiebre.length; prk++) {
        pronosticoQuiebre[prk].enCaminoInfo = enCaminoProno[pronosticoQuiebre[prk].codigo] || null;
      }
    } catch(ecp) { Logger.log('enCaminoProno: ' + ecp); }

    return {
      bajoMinimo:       bajoMinimo,
      valorDeposito:    Math.round(valorDeposito * 100) / 100,
      solicPendientes:  solicPendientes,
      casTransito:      casTransito,
      alertasCriticas:  alertasCriticas,
      pronosticoQuiebre:pronosticoQuiebre,
      topValor:         topValor,
      inmovilizado:     inmovilizado.slice(0, 12),
      rotacionMensual:  rotacionMensual,
      leadDias:         LEAD_DIAS,
      metricas: {
        leadTime:       cntLead   ? Math.round(sumLead   / cntLead)   : null,
        aduana:         cntAduana ? Math.round(sumAduana / cntAduana) : null,
        indiceQuiebre:  totalSKU  ? Math.round((bajoMinimo.quiebre / totalSKU) * 100) : 0
      }
    };
  } catch(e) {
    Logger.log("Error en Dashboard: " + e);
    return { bajoMinimo:{total:0,A:0,B:0,C:0,D:0}, solicPendientes:[], casTransito:[], alertasCriticas:[] };
  }
}


//  CATÁLOGO REPUESTOS DJI (para Borrador de Pedido)
function cargarCatalogoBorrador() {
  try {
    var stkMap = _getCarmenStockMap(); // { SKU_UPPER: stockActual }
    var ss   = SpreadsheetApp.openById(CATALOGO_REPUESTOS_ID);
    var hoja = ss.getSheets()[0];
    var d    = hoja.getDataRange().getValues();
    var out  = [];
    // Cols: 0=Código Largo, 1=Código Corto, 2=Descripción, 3=Modelo, 4=Cant/eq, 5=Precio PVP
    for (var i = 1; i < d.length; i++) {
      var codL = String(d[i][0] || '').trim();
      var codC = String(d[i][1] || '').trim();
      var desc = String(d[i][2] || '').trim();
      if (!codL || !desc) continue;
      // Buscar stock: primero por Código Largo, luego por Código Corto
      var stk = stkMap[codL.toUpperCase()];
      if (stk === undefined) stk = stkMap[codC.toUpperCase()];
      out.push({
        codigo:      codL,
        codigoCorto: codC,
        descripcion: desc,
        modelo:      String(d[i][3] || '').trim(),
        stockActual: stk !== undefined ? stk : null
      });
    }
    return { ok: true, items: out };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

//  STOCK
function cargarStock(filtro) {
  try {
    // Carmen es la fuente primaria de ítems y stock actual
    var hojaCarmen = _getCarmenSS().getSheetByName('STOCK');
    var dCarmen    = hojaCarmen ? hojaCarmen.getDataRange().getValues() : [];
    // STOCK_REPUESTOS provee metadata (mínimo, categoría, ubicación, etc.)
    var dMaster    = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S          = SCHEMA.STOCK_REPUESTOS;
    var masterMap  = {};
    for (var mi = 1; mi < dMaster.length; mi++) {
      var mk = String(dMaster[mi][S.CODIGO] || '').trim().toUpperCase();
      if (mk) masterMap[mk] = dMaster[mi];
    }
    var out = [];
    var q   = filtro ? filtro.toLowerCase().trim() : "";

    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var hoy  = new Date();
    var lastCountMap = {};
    for (var m = 1; m < dMov.length; m++) {
      var mTipo = String(dMov[m][M.TIPO] || '');
      if (mTipo.indexOf('AJUSTE') === -1 && mTipo.indexOf('CONTEO') === -1) continue;
      var mCod = String(dMov[m][M.CODIGO] || '').trim().toUpperCase();
      if (!mCod) continue;
      var mFecha = dMov[m][M.FECHA];
      if (!(mFecha instanceof Date)) continue;
      if (!lastCountMap[mCod] || mFecha > lastCountMap[mCod].fecha) {
        lastCountMap[mCod] = { fecha: mFecha, dias: Math.floor((hoy - mFecha) / 86400000) };
      }
    }

    var dUbic = [];
    var ubicMultiMap = _getCarmenUbicMap(); // tab UBICACIONES en Carmen

    // TABLA_POSICIONES: bins WMS con BIN_ID y TIPO_ALMACEN
    var dPos = [];
    try { dPos = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES); } catch(ep) {}
    var TP = SCHEMA.TABLA_POSICIONES;
    var binMap = {};    // SKU → [{ binId, cantidad, tipoAlmacen }]
    var binTotalMap = {}; // SKU → suma de cantidades en bins
    for (var p = 1; p < dPos.length; p++) {
      var pSku = String(dPos[p][TP.SKU] || '').trim().toUpperCase();
      if (!pSku) continue;
      var pCant = parseInt(dPos[p][TP.CANTIDAD]) || 0;
      if (!binMap[pSku]) { binMap[pSku] = []; binTotalMap[pSku] = 0; }
      binMap[pSku].push({
        binId:       String(dPos[p][TP.BIN_ID]       || ''),
        cantidad:    pCant,
        tipoAlmacen: String(dPos[p][TP.TIPO_ALMACEN] || '')
      });
      binTotalMap[pSku] += pCant;
    }

    var diasMap = calcularDiasDeStockPorSku();

    // Mapa SKU → unidades reservadas (RESERVAS_STOCK activas)
    var reservadoMap = {};
    try {
      var hojaRes = getSheet(SCHEMA.SHEETS.RESERVAS);
      if (hojaRes) {
        var dResS = getSheetValues(hojaRes);
        var RS    = SCHEMA.RESERVAS_STOCK;
        for (var rsi = 1; rsi < dResS.length; rsi++) {
          if (String(dResS[rsi][RS.ESTADO] || '') !== 'Activa') continue;
          var rsSku  = String(dResS[rsi][RS.SKU]      || '').trim().toUpperCase();
          var rsCant = parseInt(dResS[rsi][RS.CANTIDAD]) || 0;
          if (rsSku && rsCant > 0) reservadoMap[rsSku] = (reservadoMap[rsSku] || 0) + rsCant;
        }
      }
    } catch(eR) { Logger.log('cargarStock reservado: ' + eR); }

    // Mapa SKU → unidades en camino desde compras DJI activas (excluye Borrador y En depósito)
    var enCaminoMap    = {};
    var enCaminoCasMap = {};
    try {
      var dCAS = getSheetValues(SCHEMA.SHEETS.COMPRAS);
      var casEstadoMap = {};
      var estadosExcluidos = { 'En depósito': true };
      for (var ci = 1; ci < dCAS.length; ci++) {
        var casId  = String(dCAS[ci][0] || '').trim().toUpperCase();
        var casEst = String(dCAS[ci][2] || '').trim();
        if (casId) casEstadoMap[casId] = casEst;
      }
      var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
      if (hojaCD) {
        var dCD = getSheetValues(hojaCD);
        var CD  = SCHEMA.COMPRAS_DETALLE;
        for (var ci2 = 1; ci2 < dCD.length; ci2++) {
          var cdCas  = String(dCD[ci2][CD.ID_CAS] || '').trim().toUpperCase();
          var cdSku  = String(dCD[ci2][CD.SKU]    || '').trim().toUpperCase();
          var cdPed  = parseInt(dCD[ci2][CD.CANTIDAD_PEDIDA])   || 0;
          var cdRec  = parseInt(dCD[ci2][CD.CANTIDAD_RECIBIDA]) || 0;
          var cdEst  = casEstadoMap[cdCas] || '';
          if (!cdSku || !cdEst || estadosExcluidos[cdEst] || cdEst.indexOf('Borrador') !== -1) continue;
          var pendiente = Math.max(0, cdPed - cdRec);
          if (pendiente > 0) {
            enCaminoMap[cdSku] = (enCaminoMap[cdSku] || 0) + pendiente;
            if (!enCaminoCasMap[cdSku]) enCaminoCasMap[cdSku] = [];
            enCaminoCasMap[cdSku].push({ cas: cdCas, estado: cdEst, cant: pendiente });
          }
        }
      }
    } catch(eC) { Logger.log('cargarStock enCamino: ' + eC); }

    // Cols Carmen STOCK: 0=PN, 1=Descripción, 2=Stock Actual, 3=Modelo, 4=Serie
    for (var i = 1; i < dCarmen.length; i++) {
      var cod  = String(dCarmen[i][0] || '').trim();
      var desc = String(dCarmen[i][1] || '').trim();
      if (!cod || !desc) continue;
      if (q && cod.toLowerCase().indexOf(q) === -1 && desc.toLowerCase().indexOf(q) === -1) continue;
      var codKey = cod.toUpperCase();
      // Stock: Carmen col 2 es la fuente; fallback bins WMS
      var act = (binTotalMap[codKey] !== undefined) ? binTotalMap[codKey] : (parseInt(dCarmen[i][2]) || 0);
      // Metadata desde STOCK_REPUESTOS (si existe)
      var mf  = masterMap[codKey];
      var min = mf ? (parseInt(mf[S.STOCK_MINIMO]) || 0) : 0;
      var cat = mf ? String(mf[S.CATEGORIA]  || '') : '';
      var ubi = mf ? String(mf[S.UBICACION]  || '') : '';
      var mod = mf ? String(mf[S.MODELOS]    || '') : String(dCarmen[i][4] || '');
      var uEnt = mf ? _fmtFecha(mf[S.ULTIMA_ENTRADA]) : '—';
      var uSal = mf ? _fmtFecha(mf[S.ULTIMA_SALIDA])  : '—';
      var reqSN = mf ? mf[S.REQUIERE_SN] === true : false;
      var estado = act <= 0 ? 'CRÍTICO' : (act <= min ? 'BAJO' : 'OK');
      var cEntry   = lastCountMap[codKey];
      var diasStock = diasMap.hasOwnProperty(codKey) ? diasMap[codKey] : undefined;
      out.push({
        fila: i+1, codigo: cod, descripcion: desc,
        stockActual: act, stockMinimo: min,
        categoria: cat, ubicacion: ubi, modelos: mod,
        ultimaEntrada: uEnt, ultimaSalida: uSal,
        requireSN: reqSN,
        estado: estado,
        fechaConteo: cEntry ? _fmtFecha(cEntry.fecha) : '—',
        diasDesdeConteo: cEntry ? cEntry.dias : 9999,
        ubicaciones: ubicMultiMap[codKey] || [],
        bins: binMap[codKey] || [],
        diasStock: (diasStock === null || diasStock === undefined) ? null : diasStock,
        enCamino:    enCaminoMap[codKey]    || 0,
        enCaminoCas: enCaminoCasMap[codKey] || [],
        reservado:   reservadoMap[codKey]   || 0
      });
    }
    return out;
  } catch(e) { return []; }
}

function getUserRole() {
  try {
    var email = Session.getActiveUser().getEmail().trim().toLowerCase();
    var d = getSheetValues(SCHEMA.SHEETS.USUARIOS);
    // Usuarios_Internos layout: col 0=nombre, col 1=email, col 2=rol, col 3=esTecnico
    for (var i = 1; i < d.length; i++) {
      var rowEmail = String(d[i][1] || '').trim().toLowerCase();
      if (rowEmail === email) {
        var rol = String(d[i][2] || '').trim().toLowerCase();
        return (rol === 'admin') ? 'ADMIN' : 'OPERADOR';
      }
    }
    return 'OPERADOR';
  } catch(e) {
    return 'OPERADOR';
  }
}

function ajustarInventario(codigo, cantNueva, motivo, operador, ubicacion) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.STOCK);
    var d    = getSheetValues(hoja);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toUpperCase() !== String(codigo).trim().toUpperCase()) continue;
      var anterior = parseInt(d[i][2])||0;
      var nueva    = parseInt(cantNueva)||0;
      var diff     = nueva - anterior;
      // Stock Actual en Carmen es fórmula — no escribir ahí; solo registrar el movimiento
      _registrarMovimiento("AJUSTE_INVENTARIO", codigo, String(d[i][1]), diff, nueva,
        "Ajuste: "+motivo, operador||"Sistema");
      _registrarMovimientoCarmen(codigo, String(d[i][1]), ubicacion, diff, motivo || 'Ajuste');
      return { ok: true, anterior: anterior, nueva: nueva };
    }
    return { ok: false, msg: "Código no encontrado" };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

// ── WMS — GESTIÓN DE UBICACIONES EN CARMEN ──────────────────
function cargarUbicacionesItem(sku) {
  try {
    var codKey   = String(sku).trim().toUpperCase();
    var hojaUbic = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: true, ubicaciones: [] };
    var d = hojaUbic.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() !== codKey) continue;
      out.push({
        ubicacion: String(d[i][1] || '').trim(),
        cantidad:  parseFloat(d[i][2]) || 0
      });
    }
    return { ok: true, ubicaciones: out };
  } catch(e) { return { ok: false, error: e.message }; }
}

function guardarUbicacionInicial(sku, ubicacion, cantidadInicial) {
  try {
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe en Carmen' };
    var codKey  = String(sku).trim().toUpperCase();
    var ubicKey = String(ubicacion).trim().toUpperCase();
    var cantIni = Math.max(0, parseInt(cantidadInicial) || 0);
    var d = hojaUbic.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() === codKey &&
          String(d[i][1] || '').trim().toUpperCase() === ubicKey) {
        hojaUbic.getRange(i + 1, 3).setValue(cantIni); // col C directo, sin fórmula
        return { ok: true };
      }
    }
    hojaUbic.appendRow([codKey, ubicKey, cantIni]); // PN | Ubicación | Cantidad (número)
    return { ok: true };
  } catch(e) { return { ok: false, error: e.message }; }
}

// Retorna todos los ítems registrados en una ubicación específica
function cargarUbicacionesSector(ubicacion) {
  try {
    var ubicKey  = String(ubicacion).trim().toUpperCase();
    var hojaUbic = _getCarmenSS().getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: true, items: [] };
    var d   = hojaUbic.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1] || '').trim().toUpperCase() !== ubicKey) continue;
      out.push({
        sku:      String(d[i][0] || '').trim(),
        cantidad: parseFloat(d[i][2]) || 0
      });
    }
    return { ok: true, items: out };
  } catch(e) { return { ok: false, error: e.message }; }
}

function eliminarUbicacion(sku, ubicacion) {
  try {
    var ss       = _getCarmenSS();
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaUbic) return { ok: false, error: 'Tab UBICACIONES no existe' };
    var codKey  = String(sku).trim().toUpperCase();
    var ubicKey = String(ubicacion).trim().toUpperCase();
    var d = hojaUbic.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0] || '').trim().toUpperCase() === codKey &&
          String(d[i][1] || '').trim().toUpperCase() === ubicKey) {
        hojaUbic.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Ubicación no encontrada' };
  } catch(e) { return { ok: false, error: e.message }; }
}

// Devuelve ítems con stock en Carmen que no tienen ninguna ubicación en UBICACIONES.
// Ordenados desc por stock. Limita a 200 ítems para no saturar.
function cargarSinMapear() {
  try {
    var ss       = _getCarmenSS();
    var hojaStk  = ss.getSheetByName('STOCK');
    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);
    if (!hojaStk) return { ok: false, error: 'Hoja STOCK no encontrada en Carmen' };

    // Set de SKUs ya mapeados
    var mapeados = {};
    if (hojaUbic) {
      var dU = hojaUbic.getDataRange().getValues();
      for (var u = 1; u < dU.length; u++) {
        var sk = String(dU[u][0] || '').trim().toUpperCase();
        if (sk) mapeados[sk] = true;
      }
    }

    var dS  = hojaStk.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < dS.length; i++) {
      var cod   = String(dS[i][0] || '').trim().toUpperCase();
      var stock = parseFloat(dS[i][2]) || 0;
      if (!cod || stock <= 0 || mapeados[cod]) continue;
      out.push({ sku: cod, descripcion: String(dS[i][1] || ''), stock: stock });
    }
    out.sort(function(a, b) { return b.stock - a.stock; });
    return { ok: true, items: out.slice(0, 200), total: out.length };
  } catch(e) { return { ok: false, error: e.message }; }
}

// ============================================================
//  TRANSFERENCIA INTERNA WMS
//  Mueve stock entre bins dentro de TABLA_POSICIONES.
//  No altera el saldo total en STOCK_REPUESTOS.
// ============================================================
function transferirStock(sku, origen, destino, cant) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var cantN  = parseInt(cant) || 0;
    if (cantN <= 0) throw new Error('La cantidad debe ser mayor a 0');
    var codKey = String(sku).trim().toUpperCase();
    var origenK = String(origen).trim();
    var destinoK = String(destino).trim();
    if (!origenK || !destinoK) throw new Error('Origen y destino son requeridos');
    if (origenK === destinoK) throw new Error('Origen y destino no pueden ser iguales');

    var hoja    = getSheet(SCHEMA.SHEETS.TABLA_POSICIONES);
    if (!hoja) throw new Error('Hoja TABLA_POSICIONES no existe. Ejecutá asegurarHojas() primero.');
    var TP      = SCHEMA.TABLA_POSICIONES;
    var datos   = hoja.getDataRange().getValues();

    var origenFila  = -1, origenCant = 0;
    var destinoFila = -1;

    for (var i = 1; i < datos.length; i++) {
      var fSku = String(datos[i][TP.SKU]    || '').trim().toUpperCase();
      var fBin = String(datos[i][TP.BIN_ID] || '').trim();
      if (fSku !== codKey) continue;
      if (fBin === origenK)  { origenFila  = i + 1; origenCant  = parseInt(datos[i][TP.CANTIDAD]) || 0; }
      if (fBin === destinoK) { destinoFila = i + 1; }
    }

    if (origenFila < 0) throw new Error('SKU "' + sku + '" no encontrado en bin "' + origenK + '"');
    if (origenCant < cantN) throw new Error('Stock insuficiente en ' + origenK + '. Disponible: ' + origenCant + ', solicitado: ' + cantN);

    // Decrementar origen
    hoja.getRange(origenFila, TP.CANTIDAD + 1).setValue(origenCant - cantN);

    // Incrementar destino o crear fila nueva
    if (destinoFila > 0) {
      var destCant = parseInt(datos[destinoFila - 1][TP.CANTIDAD]) || 0;
      hoja.getRange(destinoFila, TP.CANTIDAD + 1).setValue(destCant + cantN);
    } else {
      hoja.appendRow([codKey, destinoK, cantN, '']);
    }

    invalidateSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);

    _registrarMovimiento(
      'TRANSFERENCIA_INTERNA', codKey, '', 0,
      (origenCant - cantN),
      'De: ' + origenK + ' → A: ' + destinoK + ' · ' + cantN + ' u.',
      Session.getActiveUser().getEmail()
    );

    return { ok: true, sku: codKey, origen: origenK, destino: destinoK, cantidad: cantN };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

// ============================================================
//  CARGAR BINS DE UN SKU (para el modal WMS)
// ============================================================
function cargarBinsSKU(sku) {
  try {
    var codKey = String(sku).trim().toUpperCase();
    var dPos   = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);
    var TP     = SCHEMA.TABLA_POSICIONES;
    var bins   = [];
    for (var i = 1; i < dPos.length; i++) {
      if (String(dPos[i][TP.SKU] || '').trim().toUpperCase() !== codKey) continue;
      bins.push({
        binId:       String(dPos[i][TP.BIN_ID]       || ''),
        cantidad:    parseInt(dPos[i][TP.CANTIDAD])   || 0,
        tipoAlmacen: String(dPos[i][TP.TIPO_ALMACEN] || '')
      });
    }
    bins.sort(function(a, b) { return _binSortKey(a.binId) < _binSortKey(b.binId) ? -1 : 1; });
    return { ok: true, bins: bins };
  } catch(e) { return { ok: false, msg: e.toString(), bins: [] }; }
}

// Parsea BIN_ID "PASILLO-ESTANTE-NIVEL" para ordenamiento lógico de ruta
function _binSortKey(binId) {
  var parts = String(binId || '').trim().split('-');
  var pasillo = (parts[0] || '').toUpperCase();
  var estante = parts[1] ? (isNaN(parts[1]) ? parts[1] : ('00' + parts[1]).slice(-3)) : '';
  var nivel   = parts[2] ? (isNaN(parts[2]) ? parts[2] : ('00' + parts[2]).slice(-3)) : '';
  return pasillo + '|' + estante + '|' + nivel;
}

// ============================================================
//  REGISTRAR MOVIMIENTO SEGURO
//  Acepta firma posicional o por objeto:
//    registrarMovimientoSeguro(sku, cantidad, tipo, motivo, operador)
//    registrarMovimientoSeguro({ sku, cantidad, tipo, motivo, usuario, nuevaUbicacion })
//  tipo: 'INGRESO' | 'EGRESO'
//  nuevaUbicacion (opcional): actualiza la ubicación física del ítem junto con el movimiento.
//  Valida stock antes de egresar, actualiza timestamp,
//  registra en MOVIMIENTOS_STOCK y hace flush atómico.
// ============================================================
function registrarMovimientoSeguro(sku, cantidad, tipo, motivo, operador) {
  // Soporte para firma por objeto: { sku, cantidad, tipo, motivo, usuario, nuevaUbicacion }
  var nuevaUbicacion = null;
  if (sku && typeof sku === 'object') {
    var d = sku;
    nuevaUbicacion = d.nuevaUbicacion || null;
    cantidad       = d.cantidad;
    tipo           = d.tipo;
    motivo         = d.motivo         || d.referencia || '';
    operador       = d.usuario        || d.operador   || '';
    sku            = d.sku;
  }

  var codBusc = String(sku || '').trim().toUpperCase();
  var cant    = parseInt(cantidad) || 0;
  var tipoStr = String(tipo || 'EGRESO').trim();

  // Delegar al ledger engine para la lógica de lock+verify+append+sync
  var delta  = (tipoStr === 'EGRESO') ? -cant : cant;
  var result = registrarEventoLedgerSeguro({
    sku:        codBusc,
    delta:      delta,
    tipo:       tipoStr,
    referencia: motivo   || '',
    operador:   operador || ''
  });

  if (!result.ok) {
    return { success: false, error: result.error };
  }

  // Actualizar ubicación si fue solicitada (no la maneja el ledger engine genérico)
  if (nuevaUbicacion) {
    try {
      var hojaStr = getSheet(SCHEMA.SHEETS.STOCK);
      var dStr    = getSheetValues(hojaStr);
      var S       = SCHEMA.STOCK_REPUESTOS;
      for (var i = 1; i < dStr.length; i++) {
        if (String(dStr[i][S.CODIGO] || '').trim().toUpperCase() !== codBusc) continue;
        hojaStr.getRange(i + 1, S.UBICACION + 1).setValue(String(nuevaUbicacion).trim());
        invalidateSheetValues(SCHEMA.SHEETS.STOCK);
        break;
      }
    } catch(eUbic) {
      Logger.log('[registrarMovimientoSeguro] ubicacion update failed: ' + eUbic.message);
    }
  }

  return {
    success:    true,
    stockNuevo: result.saldoResultante,
    ubicacion:  nuevaUbicacion || ''
  };
}

function editarConfigStock(fila, minimo, categoria, ubicacion, requireSN) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.STOCK);
    var f = parseInt(fila);
    var catStr = String(categoria||'').trim().toUpperCase();
    if (catStr && !_catValida(catStr)) return { ok: false, msg: 'Categoría inválida: debe ser texto (A/B/C/D), no un número.' };
    hoja.getRange(f, 4).setValue(parseInt(minimo)||0);
    hoja.getRange(f, 5).setValue(catStr);
    hoja.getRange(f, 6).setValue(String(ubicacion));
    if (requireSN !== undefined) {
      hoja.getRange(f, SCHEMA.STOCK_REPUESTOS.REQUIERE_SN + 1).setValue(requireSN === true || requireSN === 'true');
    }
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function importarStockInicial(filas) {
  // filas = array de { codigo, descripcion, stockActual, modelos }
  try {
    var hoja = getSheet(SCHEMA.SHEETS.STOCK);
    var d    = getSheetValues(hoja);
    var codigosExistentes = {};
    for (var e = 1; e < d.length; e++) codigosExistentes[String(d[e][0]).trim().toUpperCase()] = e+1;
    var nuevos = 0, actualizados = 0;
    for (var i = 0; i < filas.length; i++) {
      var f   = filas[i];
      var cod = String(f.codigo||"").trim().toUpperCase();
      if (!cod) continue;
      if (codigosExistentes[cod]) {
        hoja.getRange(codigosExistentes[cod], 3).setValue(parseInt(f.stockActual)||0);
        actualizados++;
      } else {
        hoja.appendRow([cod, f.descripcion||"", parseInt(f.stockActual)||0, 0, "", "", f.modelos||"", "", "", false]);
        nuevos++;
      }
    }
    return { ok: true, nuevos: nuevos, actualizados: actualizados };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function obtenerMovimientos(codigo, limite) {
  try {
    var d    = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var out  = [], max = limite || 50;
    var codB = codigo ? String(codigo).trim().toUpperCase() : null;
    for (var i = d.length - 1; i >= 1 && out.length < max; i--) {
      if (codB && String(d[i][SCHEMA.MOVIMIENTOS_STOCK.CODIGO]).trim().toUpperCase() !== codB) continue;
      out.push({
        fila:           i + 1,
        fecha:          _fmtFecha(d[i][SCHEMA.MOVIMIENTOS_STOCK.FECHA]),
        tipo:           String(d[i][SCHEMA.MOVIMIENTOS_STOCK.TIPO]         || ""),
        codigo:         String(d[i][SCHEMA.MOVIMIENTOS_STOCK.CODIGO]       || ""),
        descripcion:    String(d[i][SCHEMA.MOVIMIENTOS_STOCK.DESCRIPCION]  || ""),
        cantidad:       parseInt(d[i][SCHEMA.MOVIMIENTOS_STOCK.CANTIDAD])  || 0,
        stockResultante:parseInt(d[i][SCHEMA.MOVIMIENTOS_STOCK.STOCK_RESULTANTE]) || 0,
        referencia:     String(d[i][SCHEMA.MOVIMIENTOS_STOCK.REFERENCIA]   || "—"),
        operador:       String(d[i][SCHEMA.MOVIMIENTOS_STOCK.OPERADOR]     || ""),
        observaciones:  String(d[i][SCHEMA.MOVIMIENTOS_STOCK.OBSERVACIONES]|| "")
      });
    }
    return out;
  } catch(e) { Logger.log("obtenerMovimientos: " + e); return []; }
}

// Alias para el tab de auditoría del frontend
function obtenerUltimosLogs(limite) {
  return obtenerMovimientos(null, limite || 200);
}

// Genera un vale PDF firmable a partir de una fila de MOVIMIENTOS_STOCK.
// fila: número de fila 1-based en la hoja (fila 1 = encabezado, fila 2 = primer registro).
function generarValeMovimiento(fila) {
  var S    = SCHEMA.MOVIMIENTOS_STOCK;
  var d    = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
  var row  = d[fila - 1];
  if (!row || fila < 2) {
    return { success: false, error: 'Registro no encontrado (fila ' + fila + ').' };
  }

  var fecha       = row[S.FECHA];
  var tipo        = String(row[S.TIPO]           || '');
  var codigo      = String(row[S.CODIGO]         || '');
  var descripcion = String(row[S.DESCRIPCION]    || '');
  var cantidad    = row[S.CANTIDAD];
  var stockResult = row[S.STOCK_RESULTANTE];
  var referencia  = String(row[S.REFERENCIA]     || '—');
  var operador    = String(row[S.OPERADOR]        || '—');
  var observ      = String(row[S.OBSERVACIONES]  || '');
  var fechaStr    = _fmtFecha(fecha instanceof Date ? fecha : new Date(fecha));
  var ahora       = _fmtFecha(new Date());

  try {
    var docName = 'Vale_' + codigo + '_' + fechaStr.replace(/\//g, '-');
    var doc     = DocumentApp.create(docName);
    var body    = doc.getBody();
    body.clear();

    // Membrete
    body.appendParagraph('BIDCOMAGRO')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('Servicio Técnico Oficial DJI Agras · Stock Manager')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setFontSize(9);
    body.appendHorizontalRule();

    body.appendParagraph('VALE DE MOVIMIENTO DE STOCK')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('');

    // Tabla de datos
    var tableData = [
      ['FECHA DEL MOVIMIENTO', fechaStr],
      ['TIPO',                  tipo],
      ['CÓDIGO (SKU)',           codigo],
      ['DESCRIPCIÓN',           descripcion],
      ['CANTIDAD',              String(cantidad)],
      ['STOCK RESULTANTE',      String(stockResult)],
      ['REFERENCIA / OT',       referencia],
      ['RESPONSABLE',           operador]
    ];
    if (observ) tableData.push(['OBSERVACIONES', observ]);

    var table = body.appendTable(tableData);
    for (var r = 0; r < table.getNumRows(); r++) {
      table.getCell(r, 0)
           .setFontSize(9)
           .setBold(true)
           .setBackgroundColor('#f0f2f5');
      table.getCell(r, 1).setFontSize(11);
    }

    body.appendParagraph('');
    body.appendParagraph(
      '_______________________          _______________________          _______________________'
    ).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph('  Responsable de almacén                  Autorizado por                          Recibido / Firma')
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
        .setFontSize(9);
    body.appendParagraph('');
    body.appendParagraph('Generado el ' + ahora + '  ·  BIDCOMAGRO  ·  Sistema de Gestión de Stock')
        .setFontSize(8)
        .setAlignment(DocumentApp.HorizontalAlignment.CENTER);

    doc.saveAndClose();

    var pdfBlob = DriveApp.getFileById(doc.getId())
                    .getAs(MimeType.PDF)
                    .setName(docName + '.pdf');
    var pdfFile = DriveApp.createFile(pdfBlob);
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Eliminar el Doc temporal
    DriveApp.getFileById(doc.getId()).setTrashed(true);

    return { success: true, url: pdfFile.getUrl(), nombre: docName };

  } catch(e) {
    Logger.log('generarValeMovimiento [fila=' + fila + ']: ' + e);
    return { success: false, error: e.message };
  }
}

function obtenerAlertasStockCritico() {
  var S         = SCHEMA.STOCK_REPUESTOS;
  var data      = getSheetValues(SCHEMA.SHEETS.STOCK);
  var carmenMap = _getCarmenStockMap();

  // Cruzar con OTs en "Espera de repuestos" para contar OTs bloqueadas por SKU
  var otsEspera = {};
  try {
    var dOT = getSheetValues(SCHEMA.SHEETS.OT);
    for (var o = 1; o < dOT.length; o++) {
      var estado = String(dOT[o][SCHEMA.OT.ESTADO]||"");
      if (estado !== "Espera de repuestos") continue;
      var repStr = String(dOT[o][SCHEMA.OT.REPUESTOS]||"").trim();
      if (!repStr || repStr === "Sin consumo de repuestos") continue;
      var partes = repStr.split(" ; ");
      for (var rp = 0; rp < partes.length; rp++) {
        var skuRaw = String(partes[rp].split(" | ")[0]).trim().toUpperCase();
        if (skuRaw) otsEspera[skuRaw] = (otsEspera[skuRaw]||0) + 1;
      }
    }
  } catch(e2) { Logger.log("obtenerAlertasStockCritico OTs: " + e2); }

  var alertas = [];
  for (var i = 1; i < data.length; i++) {
    var row    = data[i];
    var sku    = String(row[S.CODIGO]||"").trim().toUpperCase();
    var actual = (carmenMap[sku] !== undefined) ? carmenMap[sku] : Number(row[S.STOCK_ACTUAL]);
    var minimo = Number(row[S.STOCK_MINIMO]);
    if (!sku) continue;
    if (actual <= minimo) {
      var bloqueadas = otsEspera[sku] || 0;
      alertas.push({
        sku:       sku,
        nombre:    row[S.DESCRIPCION],
        actual:    actual,
        minimo:    minimo,
        categoria: String(row[S.CATEGORIA]||""),
        estado:    actual === 0 ? 'QUIEBRE' : 'BAJO',
        bloqueadas: bloqueadas   // OTs que esperan este SKU
      });
    }
  }
  // Ordenar: quiebres primero, luego por OTs bloqueadas desc
  alertas.sort(function(a, b) {
    if (a.estado !== b.estado) return a.estado === 'QUIEBRE' ? -1 : 1;
    return b.bloqueadas - a.bloqueadas;
  });
  return alertas;
}



// ============================================================
//  COMPRAS DJI
// ============================================================
var ESTADOS_CAS = [
  "Comprado","Pagado","Envío confirmado","Forwarder HK",
  "En vuelo","En aduana","En depósito"
];

function cargarCompras() {
  try {
    var d   = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var out = [];
    for (var i = 1; i < d.length; i++) {
      var f = d[i];
      var fechaVuelo  = f[8]  instanceof Date ? f[8]  : null;
      var fechaDepo   = f[10] instanceof Date ? f[10] : null;
      var diasTransito = (fechaVuelo && !fechaDepo)
        ? Math.floor((new Date() - fechaVuelo) / 86400000) : null;
      out.push({
        fila: i+1, cas: String(f[0]), fechaPedido: _fmtFecha(f[1]),
        estado: String(f[2]||"Comprado"), metodoPago: String(f[3]||""),
        fechas: {
          comprado:    _fmtFecha(f[4]),  pagado:       _fmtFecha(f[5]),
          confEnvio:   _fmtFecha(f[6]),  forwarderHK:  _fmtFecha(f[7]),
          vuelo:       _fmtFecha(f[8]),  aduana:       _fmtFecha(f[9]),
          deposito:    _fmtFecha(f[10])
        },
        diasEnTransito: diasTransito,
        operador: String(f[11]||""), observaciones: String(f[12]||""),
        ultimaActualizacion: f[13] instanceof Date ? _fmtFecha(f[13]) : (f[13] ? String(f[13]) : null)
      });
    }
    out.sort(function(a,b){ return ESTADOS_CAS.indexOf(a.estado) - ESTADOS_CAS.indexOf(b.estado); });
    return out;
  } catch(e) { return []; }
}

function registrarCAS(cas, metodoPago, operador) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS);
    var d    = getSheetValues(hoja);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toUpperCase() === String(cas).trim().toUpperCase())
        return { ok: false, msg: "El CAS ya existe" };
    }
    hoja.appendRow([cas.trim().toUpperCase(), new Date(), "Comprado", metodoPago||"",
                    new Date(), "","","","","","", operador||"", ""]);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function actualizarEstadoCAS(cas, nuevoEstado, observaciones, operador) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS);
    var d    = getSheetValues(hoja);
    var casB = String(cas).trim().toUpperCase();
    var colFecha = { "Comprado":5,"Pagado":6,"Envío confirmado":7,"Forwarder HK":8,"En vuelo":9,"En aduana":10,"En depósito":11 };
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim().toUpperCase() !== casB) continue;
      var estadoAnterior = String(d[i][SCHEMA.COMPRAS_DJI.ESTADO] || '');
      var ahora = new Date();

      hoja.getRange(i+1, 3).setValue(nuevoEstado);
      if (colFecha[nuevoEstado]) hoja.getRange(i+1, colFecha[nuevoEstado]).setValue(ahora);
      if (observaciones) hoja.getRange(i+1, 13).setValue(String(d[i][12]||"")+" | "+observaciones);
      hoja.getRange(i+1, 12).setValue(operador||"");
      hoja.getRange(i+1, 14).setValue(ahora); // Última actualización

      // Registrar en historial
      _logHistorialCAS(casB, estadoAnterior, nuevoEstado, operador, observaciones);

      invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);
      if (nuevoEstado === "En depósito") {
        _alertarBackordersPendientes(cas);
      }
      return { ok: true };
    }
    return { ok: false, msg: "CAS no encontrado" };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function _logHistorialCAS(idCas, estadoAnterior, estadoNuevo, operador, observaciones) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.HISTORIAL_COMPRAS);
    if (!hoja) {
      hoja = getDb().insertSheet(SCHEMA.SHEETS.HISTORIAL_COMPRAS);
      hoja.appendRow(['Fecha','ID_CAS','Estado anterior','Estado nuevo','Operador','Observaciones']);
      hoja.setFrozenRows(1);
    }
    hoja.appendRow([new Date(), idCas, estadoAnterior, estadoNuevo, operador||'', observaciones||'']);
    invalidateSheetValues(SCHEMA.SHEETS.HISTORIAL_COMPRAS);
  } catch(e) { Logger.log('_logHistorialCAS: ' + e); }
}

function recibirMercaderia(cas, items, operador, deposito) {
  // items = [{ codigo, descripcion, cantRecibida }]
  deposito = String(deposito || 'BA').trim().toUpperCase();
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hojaStr = getSheet(SCHEMA.SHEETS.STOCK);
    var dStr    = getSheetValues(hojaStr);
    var stockIdx = {};
    for (var s = 1; s < dStr.length; s++) {
      stockIdx[String(dStr[s][0]).trim().toUpperCase()] = s+1;
    }

    // Índice de DB_REPUESTOS para enriquecer SKUs nuevos
    var dbRepIdx = {};
    var dDbRep = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    for (var dr = 1; dr < dDbRep.length; dr++) {
      var drSku = String(dDbRep[dr][1]||'').trim().toUpperCase();
      if (drSku) dbRepIdx[drSku] = dr;
    }

    var hoy = new Date();
    var strChanged = false;
    var nuevasFila = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var cod  = String(item.codigo).trim().toUpperCase();
      var cant = parseInt(item.cantRecibida)||0;
      if (cant <= 0) continue;
      var filaStr = stockIdx[cod];
      if (filaStr) {
        var actual = parseInt(dStr[filaStr-1][2])||0;
        var nuevo  = actual + cant;
        dStr[filaStr-1][2] = nuevo;
        dStr[filaStr-1][7] = hoy;
        // Rellenar campos vacíos desde DB_REPUESTOS si los tiene
        var drI = dbRepIdx[cod];
        if (drI !== undefined) {
          if (!dStr[filaStr-1][4] && dDbRep[drI][5]) dStr[filaStr-1][4] = String(dDbRep[drI][5]); // categoria
          if (!dStr[filaStr-1][6] && dDbRep[drI][3]) dStr[filaStr-1][6] = String(dDbRep[drI][3]); // modelos
        }
        strChanged = true;
        _registrarMovimiento("ENTRADA_COMPRA", cod, String(dStr[filaStr-1][1]),
          cant, nuevo, cas, operador||"", deposito);
      } else {
        // Código nuevo: enriquecer con datos de DB_REPUESTOS / CATALOGO_DJI si existen
        var drIdx = dbRepIdx[cod];
        var descFinal = (drIdx !== undefined && dDbRep[drIdx][2]) ? String(dDbRep[drIdx][2]) : (item.descripcion||'');
        var catFinal  = (drIdx !== undefined && _catValida(String(dDbRep[drIdx][5]||''))) ? String(dDbRep[drIdx][5]) : '';
        var modFinal  = (drIdx !== undefined && dDbRep[drIdx][3]) ? String(dDbRep[drIdx][3]) : '';
        var fobFinal  = (drIdx !== undefined && dDbRep[drIdx][6]) ? parseFloat(dDbRep[drIdx][6])||0 : 0;
        nuevasFila.push([cod, descFinal, cant, 0, catFinal, '', modFinal, hoy, '', false]);
        _registrarMovimiento("ENTRADA_COMPRA", cod, descFinal,
          cant, cant, cas, operador||"", deposito);
        // Si no existe en DB_REPUESTOS, crearlo con los datos disponibles
        if (drIdx === undefined) {
          var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
          if (hojaRep) {
            hojaRep.appendRow(['', cod, descFinal, modFinal, '', catFinal, fobFinal]);
            invalidateSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
          }
        }
      }
    }
    if (strChanged) hojaStr.getDataRange().setValues(dStr);
    for (var na = 0; na < nuevasFila.length; na++) {
      hojaStr.appendRow(nuevasFila[na]);
    }
    // Carmen se actualiza vía _escribirEnRecibidos más abajo — no tocar col C directamente
    // Cruzar con RESERVAS activas para este CAS
    var hojaRes  = getSheet(SCHEMA.SHEETS.RESERVAS);
    var dRes     = hojaRes ? getSheetValues(hojaRes) : [];
    var R        = SCHEMA.RESERVAS_STOCK;
    var destinos = [];
    var resChanged = false;
    for (var ri = 1; ri < dRes.length; ri++) {
      var resRow = dRes[ri];
      if (String(resRow[R.ESTADO]) !== 'Activa') continue;
      if (String(resRow[R.CAS_REF]).trim().toUpperCase() !== String(cas).trim().toUpperCase()) continue;
      var rSku = String(resRow[R.SKU]).trim().toUpperCase();
      for (var ii = 0; ii < items.length; ii++) {
        if (String(items[ii].codigo).trim().toUpperCase() === rSku) {
          dRes[ri][R.ESTADO] = 'Cumplida';
          resChanged = true;
          destinos.push({
            sku:        rSku,
            descripcion:String(resRow[R.DESCRIPCION]),
            cantidad:   parseInt(resRow[R.CANTIDAD]) || 0,
            origen:     String(resRow[R.ORIGEN]),
            referencia: String(resRow[R.ID_REFERENCIA])
          });
          break;
        }
      }
    }
    if (hojaRes && resChanged) {
      hojaRes.getDataRange().setValues(dRes);
      invalidateSheetValues(SCHEMA.SHEETS.RESERVAS);
    }

    // Actualizar COMPRAS_DETALLE: cantidades recibidas + estado por ítem
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (hojaCD) {
      var CD    = SCHEMA.COMPRAS_DETALLE;
      var casUp = String(cas).trim().toUpperCase();
      var dCD   = hojaCD.getDataRange().getValues();
      var cdChanged = false;
      for (var cdi = 1; cdi < dCD.length; cdi++) {
        if (String(dCD[cdi][CD.ID_CAS]).trim().toUpperCase() !== casUp) continue;
        var skuCD = String(dCD[cdi][CD.SKU]).trim().toUpperCase();
        for (var rii = 0; rii < items.length; rii++) {
          if (String(items[rii].codigo).trim().toUpperCase() !== skuCD) continue;
          var prevRecib  = parseInt(dCD[cdi][CD.CANTIDAD_RECIBIDA]) || 0;
          var newRecib   = prevRecib + (parseInt(items[rii].cantRecibida) || 0);
          var pedida     = parseInt(dCD[cdi][CD.CANTIDAD_PEDIDA]) || 0;
          var nuevoEst   = newRecib >= pedida ? 'Completo' : (newRecib > 0 ? 'Parcial' : 'Pendiente');
          dCD[cdi][CD.CANTIDAD_RECIBIDA] = newRecib;
          dCD[cdi][CD.ESTADO]            = nuevoEst;
          cdChanged = true;
          break;
        }
      }
      if (cdChanged) hojaCD.getDataRange().setValues(dCD);
      invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    }

    // Escribir cada ítem en la hoja Recibidos del spreadsheet Carmen
    _escribirEnRecibidos(cas, items, '');

    // Marcar CAS como En depósito
    actualizarEstadoCAS(cas, "En depósito", "Recepción registrada", operador);
    return { ok: true, destinoMercaderia: destinos };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function _alertarBackordersPendientes(cas) {
  try {
    var casKey = String(cas || '').trim().toUpperCase();

    // SKUs del CAS recibido (de COMPRAS_DETALLE) — usados para filtrar qué backorders revisar
    var hojaCD   = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (!hojaCD) return;
    var dCD      = getSheetValues(hojaCD);
    var CD       = SCHEMA.COMPRAS_DETALLE;
    var skusCAS  = {}; // SKU → desc
    for (var ci = 1; ci < dCD.length; ci++) {
      if (String(dCD[ci][CD.ID_CAS] || '').trim().toUpperCase() !== casKey) continue;
      var sku = String(dCD[ci][CD.SKU] || '').trim().toUpperCase();
      if (sku) skusCAS[sku] = String(dCD[ci][CD.DESCRIPCION] || '');
    }
    if (!Object.keys(skusCAS).length) return;

    // Stock actual desde STOCK_REPUESTOS (ya actualizado por la recepción)
    var dStr     = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S        = SCHEMA.STOCK_REPUESTOS;
    var stockMap = {}; // SKU → stock actual
    for (var si = 1; si < dStr.length; si++) {
      var sCod = String(dStr[si][S.CODIGO] || '').trim().toUpperCase();
      if (sCod && skusCAS[sCod] !== undefined) stockMap[sCod] = parseInt(dStr[si][S.STOCK_ACTUAL]) || 0;
    }

    // Backorders en WOS Pedidos_resellers con estado 'Backorder' para esos SKUs
    var hojaWOS = SpreadsheetApp.openById('1IjCHG0BZ4ZiISca10d9GYU2gDQvwDgWibDaStjb1giw')
                    .getSheetByName('Pedidos_resellers');
    if (!hojaWOS) return;
    var dWOS     = hojaWOS.getDataRange().getValues();
    // COL indices (Despacho_Env.js): 0=NUMERO,1=RESELLER,2=SKU,3=DESC,4=CANT_SOL,5=CANT_DESP,9=ESTADO,25=CANT_CANCEL
    var afectadas  = [];
    var procesados = {};
    for (var wi = 1; wi < dWOS.length; wi++) {
      if (String(dWOS[wi][9] || '').trim() !== 'Backorder') continue;
      var wSku = String(dWOS[wi][2] || '').trim().toUpperCase();
      if (!skusCAS[wSku]) continue;
      var numero = String(dWOS[wi][0] || '').trim();
      var key    = numero + '|' + wSku;
      if (procesados[key]) continue;
      procesados[key] = true;
      var pend = Math.max(0, (Number(dWOS[wi][4]) || 0) - (Number(dWOS[wi][5]) || 0) - (Number(dWOS[wi][25]) || 0));
      if (pend <= 0) continue;
      var stockActual = stockMap[wSku] !== undefined ? stockMap[wSku] : 0;
      afectadas.push({
        numero:   numero,
        reseller: String(dWOS[wi][1] || ''),
        sku:      wSku,
        desc:     skusCAS[wSku] || String(dWOS[wi][3] || ''),
        pend:     pend,
        stock:    stockActual,
        cubre:    stockActual >= pend
      });
    }
    if (!afectadas.length) return;

    var filas = afectadas.map(function(a) {
      var cobertura = a.cubre
        ? '<span style="color:#27ae60;font-weight:700">&#10003; Stock suficiente (' + a.stock + ' disp. / ' + a.pend + ' pend.)</span>'
        : '<span style="color:#e67e22;font-weight:700">Stock insuf.: ' + a.stock + '/' + a.pend + ' u.</span>';
      return '<tr>' +
        '<td style="padding:6px 10px;font-size:12px;font-weight:700;color:#00a3e0">' + a.numero + '</td>' +
        '<td style="padding:6px 10px;font-size:12px">' + a.reseller + '</td>' +
        '<td style="padding:6px 10px;font-size:12px">' + a.sku + '</td>' +
        '<td style="padding:6px 10px;font-size:12px">' + a.desc + '</td>' +
        '<td style="padding:6px 10px;font-size:12px;text-align:center">' + cobertura + '</td>' +
        '</tr>';
    }).join('');

    var totalCubren   = afectadas.filter(function(a) { return a.cubre; }).length;
    var totalParciales = afectadas.length - totalCubren;
    var asunto = '[WOS] Backorders desbloqueados — CAS ' + casKey +
      ' (' + (totalCubren ? totalCubren + ' total' : '') +
      (totalCubren && totalParciales ? ', ' : '') +
      (totalParciales ? totalParciales + ' parcial' : '') + ')';

    GmailApp.sendEmail(SM_CONFIG.EMAIL_SUPERVISOR, asunto, '', {
      htmlBody:
        '<div style="font-family:sans-serif;max-width:650px">' +
        '<div style="background:#00a3e0;padding:16px 20px;border-radius:8px 8px 0 0">' +
          '<span style="color:#fff;font-size:16px;font-weight:700">Backorders desbloqueados — ' + casKey + '</span>' +
        '</div>' +
        '<div style="background:#fff;border:1px solid #ddd;padding:18px 20px;border-radius:0 0 8px 8px">' +
          '<p style="font-size:13px;color:#444;margin:0 0 14px">La recepci\xf3n de <strong>' + casKey + '</strong> cubre (total o parcialmente) los siguientes backorders en WOS. Ingres\xe1 al WOS para despachar.</p>' +
          '<table style="width:100%;border-collapse:collapse;border:1px solid #eee">' +
            '<thead><tr style="background:#f5f5f5">' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">Pedido</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">Reseller</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">SKU</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:left">Descripci\xf3n</th>' +
              '<th style="padding:6px 10px;font-size:11px;text-align:center">Cobertura</th>' +
            '</tr></thead>' +
            '<tbody>' + filas + '</tbody>' +
          '</table>' +
          '<p style="font-size:11px;color:#999;margin-top:14px">Cobertura calculada sobre las unidades recibidas en este CAS vs. unidades pendientes en cada pedido WOS.</p>' +
        '</div></div>',
      name: SM_CONFIG.NOMBRE_REMITENTE,
      replyTo: SM_CONFIG.EMAIL_SUPERVISOR
    });
  } catch(e) { Logger.log('_alertarBackordersPendientes: ' + e); }
}


// Escribe los ítems recibidos en la hoja "Recibidos" del spreadsheet Carmen.
// Formato: A=codigo, B=descripcion, C=cantRecibida, D=CAS, E=fecha dd/MM/yyyy, F=(vacío), G=observaciones del CAS (número de Air, etc.)
function _escribirEnRecibidos(cas, items, observaciones) {
  try {
    var ss   = _getCarmenSS();
    var hoja = ss.getSheetByName('Recibidos');
    if (!hoja) { Logger.log('_escribirEnRecibidos: hoja Recibidos no encontrada'); return; }

    var tz       = Session.getScriptTimeZone();
    var fechaStr = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    var casStr   = String(cas || '').trim().toUpperCase();

    // Leer observaciones del CAS desde COMPRAS_DJI (ahí se guarda el número de Air y otros seguimientos)
    var obs = String(observaciones || '');
    try {
      var dComp = getSheetValues(SCHEMA.SHEETS.COMPRAS);
      for (var c = 1; c < dComp.length; c++) {
        if (String(dComp[c][SCHEMA.COMPRAS_DJI.ID_CAS] || '').trim().toUpperCase() === casStr) {
          obs = String(dComp[c][SCHEMA.COMPRAS_DJI.OBSERVACIONES] || '');
          break;
        }
      }
    } catch(eObs) { Logger.log('_escribirEnRecibidos obs lookup: ' + eObs); }

    var hojaUbic = ss.getSheetByName(CARMEN_UBICACIONES_TAB);

    for (var i = 0; i < items.length; i++) {
      var it       = items[i];
      var cant     = parseInt(it.cantRecibida) || 0;
      if (cant <= 0) continue;
      var codKey  = String(it.codigo    || '').trim().toUpperCase();
      var ubicKey = String(it.ubicacion || '').trim().toUpperCase();

      // PN | Desc | Cant | Origen | Fecha | Comprobante | obs | Aparece Inv | Ubicación
      hoja.appendRow([codKey, String(it.descripcion || ''), cant, casStr, fechaStr, '', obs, '', ubicKey]);

      // Sumar a UBICACIONES si se indicó ubicación
      if (ubicKey && hojaUbic) {
        var dU    = hojaUbic.getDataRange().getValues();
        var found = false;
        for (var ui = 1; ui < dU.length; ui++) {
          if (String(dU[ui][0] || '').trim().toUpperCase() === codKey &&
              String(dU[ui][1] || '').trim().toUpperCase() === ubicKey) {
            hojaUbic.getRange(ui + 1, 3).setValue((parseFloat(dU[ui][2]) || 0) + cant);
            found = true;
            break;
          }
        }
        if (!found) hojaUbic.appendRow([codKey, ubicKey, cant]);
      }
    }
  } catch(e) {
    Logger.log('_escribirEnRecibidos: ' + e);
  }
}

// ============================================================
//  MODELOS ACTIVOS — configuración persistente en PropertiesService
// ============================================================
function getModelosActivos() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty('sm_modelos_activos');
    if (!raw) return [];
    var parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch(e) {
    Logger.log('getModelosActivos: ' + e);
    return [];
  }
}

function setModelosActivos(modelos) {
  try {
    var lista = Array.isArray(modelos) ? modelos : [];
    // Normalizar: trim, sin duplicados, sin vacíos
    var mapa = {};
    var limpia = [];
    for (var i = 0; i < lista.length; i++) {
      var m = String(lista[i]).trim();
      if (m && !mapa[m.toUpperCase()]) {
        mapa[m.toUpperCase()] = true;
        limpia.push(m);
      }
    }
    PropertiesService.getScriptProperties().setProperty('sm_modelos_activos', JSON.stringify(limpia));
    return { ok: true, modelos: limpia };
  } catch(e) {
    Logger.log('setModelosActivos: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

// ============================================================
//  CRUCE COMPRAS EXTERNAS
// ============================================================
var _PEDIDOS_EXT_SS_ID = '15Y4tri7Egpa2Tjvq1kPXuuR7OUIsQVXgjPFUwthr7Sw';

function cruzarComprasExternas() {
  try {
    var extSS    = SpreadsheetApp.openById(_PEDIDOS_EXT_SS_ID);
    var pedSheet = extSS.getSheetByName('Pedidos') || extSS.getSheetByName('Pedidos ');
    if (!pedSheet) {
      // Búsqueda tolerante a espacios
      var allSheets = extSS.getSheets();
      for (var si2 = 0; si2 < allSheets.length; si2++) {
        if (allSheets[si2].getName().trim() === 'Pedidos') { pedSheet = allSheets[si2]; break; }
      }
    }
    if (!pedSheet) return { ok: false, msg: 'Hoja "Pedidos" no encontrada en el sheet externo.' };

    var ext = pedSheet.getDataRange().getValues();

    // Encontrar fila de encabezados (primeras 10 filas)
    var hdrIdx = -1, cCas = -1, cAir = -1, cIng = -1;
    for (var ri = 0; ri < Math.min(ext.length, 10) && hdrIdx < 0; ri++) {
      var foundCas = false, foundIng = false;
      for (var ci = 0; ci < ext[ri].length; ci++) {
        var v = String(ext[ri][ci] || '').trim().toLowerCase();
        if (v.indexOf('invoice') >= 0)                              { cCas = ci; foundCas = true; }
        if ((v.indexOf('n') === 0 && v.indexOf('air') >= 0) || v === 'air') cAir = ci;
        if (v.indexOf('ingreso') >= 0 && v.indexOf('stock') >= 0)  { cIng = ci; foundIng = true; }
      }
      if (foundCas && foundIng) hdrIdx = ri;
    }
    if (hdrIdx < 0 || cCas < 0 || cIng < 0)
      return { ok: false, msg: 'No se encontraron columnas N\xb0 INVOICE / INGRESO A STOCK en el sheet externo.' };

    // Agrupar por CAS (N° INVOICE), parar al llegar a filas vacías
    var casMap = {};
    for (var di = hdrIdx + 1; di < ext.length; di++) {
      var casNum = String(ext[di][cCas] || '').trim().toUpperCase();
      if (!casNum) continue;
      var airNum = cAir >= 0 ? String(ext[di][cAir] || '').trim() : '';
      var ing    = String(ext[di][cIng] || '').trim().toUpperCase();
      if (!casMap[casNum]) casMap[casNum] = { cas: casNum, air: airNum, total: 0, si: 0, items: [] };
      casMap[casNum].total++;
      if (ing === 'SI') casMap[casNum].si++;
      // Cols D(3)=código, E(4)=descripción, F(5)=cantidad
      var cod  = String(ext[di][3] || '').trim();
      var desc = String(ext[di][4] || '').trim();
      var qty  = parseInt(ext[di][5], 10) || 0;
      if (cod && qty > 0) casMap[casNum].items.push({ codigo: cod, descripcion: desc, cantidad: qty });
    }

    // Estado actual de cada CAS en COMPRAS_DJI de SM
    var smData   = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var smCasMap = {};
    for (var si = 1; si < smData.length; si++) {
      var smCas = String(smData[si][SCHEMA.COMPRAS_DJI.ID_CAS] || '').trim().toUpperCase();
      if (smCas) smCasMap[smCas] = String(smData[si][SCHEMA.COMPRAS_DJI.ESTADO] || '');
    }

    // Items cargados en SM por CAS (COMPRAS_DETALLE)
    var smDetailByCas = {};
    try {
      var hojaDetalle = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
      if (hojaDetalle) {
        var dDet = getSheetValues(hojaDetalle);
        var CD   = SCHEMA.COMPRAS_DETALLE;
        for (var di = 1; di < dDet.length; di++) {
          var dCas = String(dDet[di][CD.ID_CAS] || '').trim().toUpperCase();
          var dSku = String(dDet[di][CD.SKU] || '').trim().toUpperCase();
          if (!dCas || !dSku) continue;
          if (!smDetailByCas[dCas]) smDetailByCas[dCas] = {};
          smDetailByCas[dCas][dSku] = {
            pedida:   parseInt(dDet[di][CD.CANTIDAD_PEDIDA])   || 0,
            recibida: parseInt(dDet[di][CD.CANTIDAD_RECIBIDA]) || 0
          };
        }
      }
    } catch(eDet) { Logger.log('cruzarComprasExternas detalle: ' + eDet); }

    var nuevas = [], recibidas = [], diferencias = [];
    var keys = Object.keys(casMap);
    for (var ki = 0; ki < keys.length; ki++) {
      var e    = casMap[keys[ki]];
      var inSM = smCasMap.hasOwnProperty(e.cas);
      if (!inSM) {
        nuevas.push({ cas: e.cas, air: e.air, total: e.total, si: e.si, items: e.items });
        continue;
      }
      var estadoSM = smCasMap[e.cas];
      if (e.si > 0 && estadoSM !== 'En dep\xf3sito') {
        recibidas.push({ cas: e.cas, air: e.air, total: e.total, si: e.si, estadoSM: estadoSM });
      }
      // Comparar ítems para CAS que no están en depósito y tienen detalle en SM
      if (estadoSM === 'En dep\xf3sito') continue;
      var smDet = smDetailByCas[e.cas];
      if (!smDet || !e.items.length) continue;

      var extMap = {};
      for (var ei = 0; ei < e.items.length; ei++) {
        var eSku = String(e.items[ei].codigo || '').trim().toUpperCase();
        if (eSku) extMap[eSku] = { cantidad: e.items[ei].cantidad, desc: e.items[ei].descripcion };
      }

      var diffs = [];
      var allSkus = {};
      var ks = Object.keys(extMap); for (var ks0 = 0; ks0 < ks.length; ks0++) allSkus[ks[ks0]] = true;
      var ks2 = Object.keys(smDet);  for (var ks1 = 0; ks1 < ks2.length; ks1++) allSkus[ks2[ks1]] = true;

      var skuList = Object.keys(allSkus);
      for (var si2 = 0; si2 < skuList.length; si2++) {
        var sk      = skuList[si2];
        var extQty  = extMap[sk]  ? extMap[sk].cantidad      : null;
        var smQty   = smDet[sk]   ? smDet[sk].pedida         : null;
        var extDesc = extMap[sk]  ? extMap[sk].desc           : '';
        if (extQty !== smQty) diffs.push({ sku: sk, desc: extDesc, ext: extQty, sm: smQty });
      }
      if (diffs.length) diferencias.push({ cas: e.cas, estadoSM: estadoSM, air: e.air, diffs: diffs, extItems: e.items });
    }

    return { ok: true, nuevas: nuevas, recibidas: recibidas, diferencias: diferencias };
  } catch(e) {
    Logger.log('cruzarComprasExternas: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

// Reemplaza los ítems de COMPRAS_DETALLE para un CAS con los del sheet externo.
// Conserva CANTIDAD_RECIBIDA de las filas existentes que coincidan por SKU.
function sincronizarItemsCAS(cas) {
  try {
    var casKey = String(cas || '').trim().toUpperCase();
    if (!casKey) return { ok: false, error: 'CAS vacío' };

    // Leer items desde sheet externo
    var extSS    = SpreadsheetApp.openById(_PEDIDOS_EXT_SS_ID);
    var pedSheet = extSS.getSheetByName('Pedidos') || extSS.getSheetByName('Pedidos ');
    if (!pedSheet) {
      var allS = extSS.getSheets();
      for (var si = 0; si < allS.length; si++) {
        if (allS[si].getName().trim() === 'Pedidos') { pedSheet = allS[si]; break; }
      }
    }
    if (!pedSheet) return { ok: false, error: 'Hoja Pedidos no encontrada en sheet externo' };

    var ext    = pedSheet.getDataRange().getValues();
    var hdrIdx = -1, cCas = -1;
    for (var ri = 0; ri < Math.min(ext.length, 10) && hdrIdx < 0; ri++) {
      var hasInv = false, hasIng = false;
      for (var ci = 0; ci < ext[ri].length; ci++) {
        var v = String(ext[ri][ci] || '').trim().toLowerCase();
        if (v.indexOf('invoice') >= 0) { cCas = ci; hasInv = true; }
        if (v.indexOf('ingreso') >= 0 && v.indexOf('stock') >= 0) hasIng = true;
      }
      if (hasInv && hasIng) hdrIdx = ri;
    }
    if (hdrIdx < 0 || cCas < 0) return { ok: false, error: 'Columnas no encontradas en sheet externo' };

    var extItems = [];
    for (var di = hdrIdx + 1; di < ext.length; di++) {
      var extCas = String(ext[di][cCas] || '').trim().toUpperCase();
      if (extCas !== casKey) continue;
      var cod  = String(ext[di][3] || '').trim().toUpperCase();
      var desc = String(ext[di][4] || '').trim();
      var qty  = parseInt(ext[di][5], 10) || 0;
      if (cod && qty > 0) extItems.push({ sku: cod, desc: desc, cantidad: qty });
    }
    if (!extItems.length) return { ok: false, error: 'No se encontraron ítems para ' + casKey + ' en el sheet externo' };

    // Leer COMPRAS_DETALLE actual para conservar CANTIDAD_RECIBIDA
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (!hojaCD) return { ok: false, error: 'Hoja COMPRAS_DETALLE no encontrada' };
    var CD      = SCHEMA.COMPRAS_DETALLE;
    var dCD     = hojaCD.getDataRange().getValues();
    var recibMap = {};
    var rowsKeep = [dCD[0]]; // encabezado
    for (var ri2 = 1; ri2 < dCD.length; ri2++) {
      var rCas = String(dCD[ri2][CD.ID_CAS] || '').trim().toUpperCase();
      if (rCas === casKey) {
        // guardar recibido por SKU, descartar la fila (se reemplaza)
        var rSku = String(dCD[ri2][CD.SKU] || '').trim().toUpperCase();
        recibMap[rSku] = parseInt(dCD[ri2][CD.CANTIDAD_RECIBIDA]) || 0;
      } else {
        rowsKeep.push(dCD[ri2]);
      }
    }

    // Agregar nuevas filas con datos del sheet externo
    for (var xi = 0; xi < extItems.length; xi++) {
      var xIt  = extItems[xi];
      var xRec = recibMap[xIt.sku] || 0;
      var xEst = xRec >= xIt.cantidad ? 'Recibido' : (xRec > 0 ? 'Parcial' : 'Pendiente');
      rowsKeep.push([casKey, xIt.sku, xIt.desc, xIt.cantidad, xRec, xEst]);
    }

    hojaCD.clearContents();
    hojaCD.getRange(1, 1, rowsKeep.length, rowsKeep[0].length).setValues(rowsKeep);
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    SpreadsheetApp.flush();
    return { ok: true, cas: casKey, items: extItems.length };
  } catch(e) {
    Logger.log('sincronizarItemsCAS: ' + e);
    return { ok: false, error: e.toString() };
  }
}

// ============================================================
//  GENERADOR DE PEDIDO DJI — algoritmo inteligente
// ============================================================
function calcularPedidoDJI() {
  try {
    var dStr      = getSheetValues(SCHEMA.SHEETS.STOCK);
    var carmenMap = _getCarmenStockMap();
    var dMov    = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var dRep    = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var dDet    = getSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    var dComp   = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var dReserv = getSheetValues(SCHEMA.SHEETS.RESERVAS);
    var hace90  = new Date(new Date().getTime() - 90 * 86400000);

    // Lead time estimado en días según estado — valores históricos si hay datos, si no, fijos
    var _LEAD_FIJO = {
      'Pendiente':        120,
      'Comprado':          90,
      'Pagado':            75,
      'Envío confirmado':  60,
      'Forwarder HK':      45,
      'En vuelo':          20,
      'En aduana':         10,
      'En depósito':        3
    };
    var _leadHist = {};
    try { _leadHist = calcularLeadTimeHistorico(); } catch(elh) {}
    var LEAD_TIME_DIAS = {};
    var _estados = Object.keys(_LEAD_FIJO);
    for (var le = 0; le < _estados.length; le++) {
      var _est = _estados[le];
      LEAD_TIME_DIAS[_est] = (_leadHist[_est] && _leadHist[_est] > 0) ? _leadHist[_est] : _LEAD_FIJO[_est];
    }

    // Estado y lead time estimado por CAS ID
    var ESTADOS_ACTIVOS = { 'Pendiente':1,'Comprado':1,'Pagado':1,'Envío confirmado':1,'Forwarder HK':1,'En vuelo':1,'En aduana':1,'En depósito':1 };
    var casEstadoMap = {};
    for (var c = 1; c < dComp.length; c++) {
      var casId = String(dComp[c][SCHEMA.COMPRAS_DJI.ID_CAS]||'').trim().toUpperCase();
      var estC  = String(dComp[c][SCHEMA.COMPRAS_DJI.ESTADO]||'');
      if (casId && ESTADOS_ACTIVOS[estC]) {
        casEstadoMap[casId] = { estado: estC, leadTime: LEAD_TIME_DIAS[estC] || 60 };
      }
    }

    // En camino por SKU + lead time mínimo esperado (la compra más avanzada gana)
    var enCaminoMap     = {};
    var enCaminoLeadMap = {};
    for (var d = 1; d < dDet.length; d++) {
      var casRef  = String(dDet[d][SCHEMA.COMPRAS_DETALLE.ID_CAS]||'').trim().toUpperCase();
      var casMeta = casEstadoMap[casRef];
      if (!casMeta) continue;
      var skuD   = String(dDet[d][SCHEMA.COMPRAS_DETALLE.SKU]||'').trim().toUpperCase();
      var pedida = parseInt(dDet[d][SCHEMA.COMPRAS_DETALLE.CANTIDAD_PEDIDA])||0;
      var recib  = parseInt(dDet[d][SCHEMA.COMPRAS_DETALLE.CANTIDAD_RECIBIDA])||0;
      var pend   = Math.max(0, pedida - recib);
      if (!skuD || pend <= 0) continue;
      enCaminoMap[skuD] = (enCaminoMap[skuD]||0) + pend;
      if (enCaminoLeadMap[skuD] === undefined || casMeta.leadTime < enCaminoLeadMap[skuD]) {
        enCaminoLeadMap[skuD] = casMeta.leadTime;
      }
    }

    // Reservado por SKU: reservas vigentes
    var reservadoMap = {};
    var S_RES = SCHEMA.RESERVAS_STOCK;
    for (var rv = 1; rv < dReserv.length; rv++) {
      var estR = String(dReserv[rv][S_RES.ESTADO]||'').trim().toUpperCase();
      if (estR === 'DESPACHADA' || estR === 'CANCELADA') continue;
      var skuR  = String(dReserv[rv][S_RES.SKU]||'').trim().toUpperCase();
      var cantR = parseInt(dReserv[rv][S_RES.CANTIDAD])||0;
      if (skuR && cantR > 0) reservadoMap[skuR] = (reservadoMap[skuR]||0) + cantR;
    }

    // Modelos activos configurados
    var modelosActivos = getModelosActivos();

    // DB_REPUESTOS: mapa de metadata + set de SKUs que aplican a modelos activos
    var dbRepMap        = {};
    var skusModeloActivo = {};
    for (var r = 1; r < dRep.length; r++) {
      var codR     = String(dRep[r][1]||'').trim().toUpperCase();
      if (!codR) continue;
      var modelosR = String(dRep[r][3]||'').trim();
      var fobR     = parseFloat(String(dRep[r][6]||'0').replace(/[^0-9.,\-]/g,'').replace(',','.')) || 0;
      dbRepMap[codR] = { desc: String(dRep[r][2]||'').trim(), modelos: modelosR, fob: fobR, cat: String(dRep[r][4]||'').trim() };
      if (modelosActivos.length > 0 && modelosR) {
        var mUp = modelosR.toUpperCase();
        for (var ma = 0; ma < modelosActivos.length; ma++) {
          if (mUp.indexOf(String(modelosActivos[ma]).toUpperCase()) !== -1) {
            skusModeloActivo[codR] = true;
            break;
          }
        }
      }
    }

    // Salidas últimos 90 días por código
    var salidas90 = {};
    for (var m = 1; m < dMov.length; m++) {
      var tipo = String(dMov[m][1]||'');
      if (tipo !== 'SALIDA_DESPACHO' && tipo !== 'EGRESO' && tipo.indexOf('SALIDA') === -1) continue;
      var fecha = dMov[m][0];
      if (!(fecha instanceof Date) || fecha < hace90) continue;
      var cod = String(dMov[m][2]||'').trim().toUpperCase();
      salidas90[cod] = (salidas90[cod]||0) + Math.abs(parseInt(dMov[m][4])||0);
    }

    var CRIT_ORD = { 'critico':0, 'urgente':1, 'bajo':2, 'modelo':3, 'restock':4, 'ok':5 };
    var out = [];
    var skusEnStock = {};

    for (var i = 1; i < dStr.length; i++) {
      var f        = dStr[i];
      var cod2     = String(f[0]||'').trim().toUpperCase();
      if (!cod2) continue;
      skusEnStock[cod2] = true;

      var stockAct = (carmenMap[cod2] !== undefined) ? carmenMap[cod2] : (parseInt(f[2])||0);
      var stockMin = parseInt(f[3])||0;
      var sal90    = salidas90[cod2]||0;
      var enCam    = enCaminoMap[cod2]||0;
      var reserv   = reservadoMap[cod2]||0;
      var isModelo = skusModeloActivo[cod2] || false;

      // Burn rate diario y stock físico disponible (sin reservas)
      var burnRateDia = sal90 / 90;
      var stockFisico = Math.max(0, stockAct - reserv);

      // Días de stock restantes con el stock físico actual
      var diasStock = burnRateDia > 0 ? stockFisico / burnRateDia : Infinity;

      // Lead time de la compra más próxima a llegar
      var leadTime = enCaminoLeadMap[cod2] !== undefined ? enCaminoLeadMap[cod2] : 999;

      // Target 2 meses (sal90 = 3 meses → × 2/3 = 2 meses)
      var target2M    = Math.ceil(sal90 * 2 / 3);
      var stockFuturo = stockFisico + enCam;

      // URGENTE: el stock físico se agota antes de que llegue la compra en camino
      var seAgota = burnRateDia > 0 && enCam > 0 && diasStock < leadTime;
      // Unidades necesarias para cubrir el período hasta que llegue la compra
      var coberturaInmediata = seAgota ? Math.max(0, Math.ceil(burnRateDia * leadTime) - stockFisico) : 0;

      // Gap de largo plazo: lo que falta para tener 2 meses DESPUÉS de recibir lo que viene
      var gapLargoPlazo = Math.max(0, target2M - stockFuturo);

      // Si es modelo activo, asegurar al menos 1 unidad disponible en todo momento
      var modeloGap = (isModelo && stockFuturo < 1) ? 1 : 0;

      var sugerido, criticidad, motivo;

      if (sal90 > 0 || isModelo) {
        sugerido = Math.max(gapLargoPlazo, coberturaInmediata, modeloGap);

        if (stockFisico <= 0 && enCam === 0) {
          criticidad = 'critico';
          motivo     = 'Sin stock';
        } else if (seAgota) {
          criticidad = 'urgente';
          motivo     = 'Se agota en ' + Math.ceil(diasStock) + 'd — llega en ~' + leadTime + 'd';
        } else if (stockAct <= stockMin && stockMin > 0) {
          criticidad = 'bajo';
          motivo     = 'Bajo mínimo';
        } else if (sugerido > 0) {
          criticidad = 'restock';
          motivo     = 'Restock 2M';
        } else {
          criticidad = 'ok';
          motivo     = 'Cubierto';
        }
      } else {
        if (stockMin > 0 && stockAct < stockMin) {
          sugerido   = stockMin - stockAct;
          criticidad = 'bajo';
          motivo     = 'Bajo mínimo';
        } else {
          sugerido   = 0;
          criticidad = 'ok';
          motivo     = '';
        }
      }

      var dbInfo = dbRepMap[cod2] || {};
      var fob = dbInfo.fob || 0;
      out.push({
        codigo:      cod2,
        descripcion: String(f[1]||''),
        stockActual: stockAct,
        stockMinimo: stockMin,
        enCamino:    enCam,
        reservado:   reserv,
        salidas90:   sal90,
        target2M:    target2M,
        sugerido:    sugerido,
        precioFOB:   fob,
        totalFOB:    fob * sugerido,
        categoria:   String(f[4]||''),
        criticidad:  criticidad,
        motivo:      motivo,
        incluir:     sugerido > 0
      });
    }

    // Modelos activos: agregar repuestos del catálogo que nunca tuvieron stock
    if (modelosActivos.length > 0) {
      for (var db = 1; db < dRep.length; db++) {
        var dbCod = String(dRep[db][1]||'').trim().toUpperCase();
        if (!dbCod || skusEnStock[dbCod] || !skusModeloActivo[dbCod]) continue;

        var enCamDB    = enCaminoMap[dbCod] || 0;
        var sugeridoDB = Math.max(0, 1 - enCamDB);
        var fobDB      = parseFloat(String(dRep[db][6]||'0').replace(/[^0-9.,\-]/g,'').replace(',','.')) || 0;
        out.push({
          codigo:      dbCod,
          descripcion: String(dRep[db][2]||'').trim(),
          stockActual: 0,
          stockMinimo: 0,
          enCamino:    enCamDB,
          reservado:   0,
          salidas90:   0,
          target2M:    1,
          sugerido:    sugeridoDB,
          precioFOB:   fobDB,
          totalFOB:    fobDB * sugeridoDB,
          categoria:   String(dRep[db][4]||'').trim(),
          criticidad:  enCamDB > 0 ? 'ok' : 'modelo',
          motivo:      enCamDB > 0 ? 'Modelo activo · en camino' : 'Modelo activo · sin stock',
          incluir:     sugeridoDB > 0
        });
      }
    }

    out.sort(function(a,b){
      var diff = (CRIT_ORD[a.criticidad]||5) - (CRIT_ORD[b.criticidad]||5);
      if (diff !== 0) return diff;
      return (a.categoria||'Z').localeCompare(b.categoria||'Z');
    });
    return out;
  } catch(e) { Logger.log('calcularPedidoDJI: '+e); return []; }
}

// Previsión de demanda: calcula días de stock restante por SKU basado en
// rotación de los últimos 90 días. Retorna mapa { SKU: diasRestantes }
// diasRestantes = null si no hubo salidas (sin rotación = stock sin consumo)
function calcularDiasDeStockPorSku() {
  try {
    var dMov      = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var dStr      = getSheetValues(SCHEMA.SHEETS.STOCK);
    var carmenMap = _getCarmenStockMap();
    var hace90 = new Date(new Date().getTime() - 90 * 86400000);
    var S = SCHEMA.STOCK_REPUESTOS;

    var salidas90 = {};
    for (var m = 1; m < dMov.length; m++) {
      if (String(dMov[m][1]||"").indexOf("SALIDA") === -1) continue;
      var fecha = dMov[m][0];
      if (!(fecha instanceof Date) || fecha < hace90) continue;
      var cod = String(dMov[m][2]||"").trim().toUpperCase();
      salidas90[cod] = (salidas90[cod]||0) + Math.abs(parseInt(dMov[m][4])||0);
    }

    var resultado = {};
    for (var i = 1; i < dStr.length; i++) {
      var sku = String(dStr[i][S.CODIGO]||"").trim().toUpperCase();
      if (!sku) continue;
      var stockActual = (carmenMap[sku] !== undefined) ? carmenMap[sku] : (parseInt(dStr[i][S.STOCK_ACTUAL])||0);
      var sal = salidas90[sku] || 0;
      if (sal === 0) {
        resultado[sku] = null; // sin rotación, no aplica semáforo
      } else {
        var rotDiaria = sal / 90;
        resultado[sku] = stockActual > 0 ? Math.round(stockActual / rotDiaria) : 0;
      }
    }
    return resultado;
  } catch(e) { Logger.log("calcularDiasDeStockPorSku: " + e); return {}; }
}

function registrarPedidoEnviado(items, cas, metodoPago, operador) {
  try {
    var res = registrarCAS(cas, metodoPago, operador);
    if (!res.ok) return res;

    var rows = [];
    var fecha = new Date();
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      if (!item.incluir || !item.ajustado) continue;
      // Preparamos la fila para MOVIMIENTOS_STOCK
      rows.push([fecha, "PEDIDO_DJI", item.codigo, item.descripcion, item.ajustado, 0, cas, operador, ""]);
    }

    if (rows.length > 0) {
      var hMov = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
      hMov.getRange(hMov.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}


// ============================================================
//  VENTAS DIRECTAS
// ============================================================
function parsearCarritoReseller(datosFilas) {
  try {
    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    var stockMap = {};
    for (var s = 1; s < dStr.length; s++) {
      var cod = String(dStr[s][0]).trim().toUpperCase();
      stockMap[cod] = parseInt(dStr[s][2]) || 0;
    }
    
    var items = [];
    if (!datosFilas || datosFilas.length === 0) throw new Error("El archivo está vacÍ­o o no tiene el formato correcto.");

    for (var i = 0; i < datosFilas.length; i++) {
      var f = datosFilas[i];
      var cod = String(f.codigo || "").trim().toUpperCase();
      if (!cod || cod === "UNDEFINED") continue; // Ignorar filas vacÍ­as

      var stock = stockMap[cod] !== undefined ? stockMap[cod] : -1;
      items.push({
        codigo: cod,
        descripcion: String(f.descripcion || "Sin descripción"),
        cantidad: parseInt(f.cantidad) || 0,
        precioUSD: parseFloat(f.precioUSD) || 0,
        totalUSD: (parseFloat(f.precioUSD) || 0) * (parseInt(f.cantidad) || 0),
        stockDisponible: stock,
        sinStock: stock >= 0 && stock < (parseInt(f.cantidad) || 0)
      });
    }
    
    if (items.length === 0) throw new Error("No se detectaron códigos de repuestos válidos en el archivo.");
    return { ok: true, items: items };
  } catch(e) { 
    return { ok: false, msg: e.toString() }; 
  }
}

function confirmarVentaDirecta(reseller, items, nOrdenEntrega, operador) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    var ss = getSS();
    var hojaVentas = getSheet(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var hojaSolic  = getSheet(SCHEMA.SHEETS.SOLICITUDES);
    var hoy = new Date();

    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var cod = String(item.codigo).trim().toUpperCase();
      var cant = parseInt(item.cantidad) || 0;

      // 1. REGISTRO CONTABLE (Historial de ventas)
      hojaVentas.appendRow([
        hoy, 
        nOrdenEntrega, 
        "", 
        reseller, 
        cod,
        item.descripcion, 
        cant, 
        item.precioUSD, 
        "Venta directa"
      ]);

      // 2. REGISTRO OPERATIVO (Para que aparezca en la pestaa Despachos)
      var idSolicitud = "VD-" + nOrdenEntrega + "-" + i;
      hojaSolic.appendRow([
        idSolicitud,
        hoy,
        nOrdenEntrega, 
        reseller,
        cod,
        item.descripcion,
        cant,
        0,             
        "Pendiente",   
        "NORMAL",      
        "", "", ""     
      ]);
      
      // IMPORTANTE: AquÍ­ borramos cualquier lÍ­nea que diga:
      // hojaStr.getRange(...).setValue(...) o _registrarMovimiento(...)
      // El stock NO se toca en esta función.
    }

    return { ok: true, nOrden: nOrdenEntrega };

  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function cargarVentas(limite) {
  try {
    var d   = getSheetValues(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var max = limite || 200;
    var out = [], porOrden = {};
    for (var i = d.length-1; i >= 1 && out.length < max; i--) {
      var f = d[i];
      var nOrd = String(f[1]||"");
      if (!porOrden[nOrd]) {
        porOrden[nOrd] = {
          fecha: _fmtFecha(f[0]), nOrden: nOrd,
          nFactura: String(f[2]||""), reseller: String(f[3]||""),
          items: [], totalUSD: 0
        };
        out.push(porOrden[nOrd]);
      }
      var precio = parseFloat(f[7])||0, cant = parseInt(f[6])||0;
      porOrden[nOrd].items.push({
        codigo: String(f[4]), descripcion: String(f[5]),
        cantidad: cant, precioUSD: precio, totalUSD: precio*cant
      });
      porOrden[nOrd].totalUSD += precio*cant;
    }
    return out;
  } catch(e) { return []; }
}

function completarFactura(nOrdenEntrega, nFactura) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var d    = hoja.getDataRange().getValues();
    var nOrd = String(nOrdenEntrega).trim();
    var changed = false;
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][1]||"").trim() === nOrd) {
        d[i][2] = nFactura;
        changed  = true;
      }
    }
    if (changed) hoja.getDataRange().setValues(d);
    return { ok: changed };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function generarNumeroOrden() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.VENTAS_DIRECTAS);
    var maxNum = 0;
    for (var i = 1; i < d.length; i++) {
      var n = String(d[i][1]||"").replace(/[^0-9]/g,"");
      if (n) maxNum = Math.max(maxNum, parseInt(n)||0);
    }
    var num = String(maxNum + 1);
    while (num.length < 4) num = "0" + num;
    return "OE-" + num;
  } catch(e) { return "OE-0001"; }
}

function obtenerResellers() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var lista = [];
    for (var i = 1; i < d.length; i++) {
      if (d[i][0]) lista.push(String(d[i][0]).trim());
    }
    return lista.sort();
  } catch(e) { return []; }
}


// ============================================================
//  SINCRONIZAR CATÁLOGO DJI → DB_REPUESTOS
//  Lee la hoja CATALOGO_DJI (columnas detectadas por header),
//  hace upsert en DB_REPUESTOS y rellena campos vacíos en STOCK.
// ============================================================
function sincronizarCatalogoDJI(operador) {
  try {
    var hojaCat = getSheet(SCHEMA.SHEETS.CATALOGO_DJI);
    if (!hojaCat) return { ok: false, msg: 'Hoja CATALOGO_DJI no encontrada. Creala primero con asegurarHojas().' };
    var dCat = getSheetValues(hojaCat);
    if (dCat.length < 2) return { ok: false, msg: 'El catálogo está vacío — pegá la lista DJI debajo de los encabezados.' };

    // Detectar columnas por nombre de header (case-insensitive, acepta variantes)
    var headers = dCat[0].map(function(h) { return String(h||'').trim().toUpperCase().replace(/[\s\-_]+/g, ''); });
    function findCol() {
      var aliases = Array.prototype.slice.call(arguments);
      for (var a = 0; a < aliases.length; a++) {
        var idx = headers.indexOf(aliases[a].toUpperCase().replace(/[\s\-_]+/g, ''));
        if (idx !== -1) return idx;
      }
      return -1;
    }
    var COL = {
      CODIGO:      findCol('SIMPLIFIEDPARTNUMBER','SIMPLIFIED PART NUMBER','CODIGO','SKU','PARTNUMBER','PART NUMBER','CODIGOREPUESTO'),
      DESCRIPCION: findCol('ENGLISHNAME','ENGLISH NAME','DESCRIPCION','DESCRIPTION','NOMBRE','NAME','DESC'),
      MODELOS:     findCol('APPLICABLEMODELS','APPLICABLE MODELS','MODELOS','MODELS','COMPATIBLE','MODELOCOMPATIBLE'),
      PRECIO_FOB:  findCol('SOUTHAMERICA','SOUTH AMERICA','PRECIOFOB','PRECIO FOB','PRECIO','PRICE','FOB','PRECIOUSD'),
      SKU_DJI:     findCol('MATERIALNUMBER','MATERIAL NUMBER','SKUDJI','SKU DJI','CODIGODJI','PARTNODJI'),
      CATEGORIA:   findCol('CATEGORIA','CATEGORY','CAT'),
      PESO:        findCol('PESO','WEIGHT','PESOG','WEIGHTG')
    };
    if (COL.CODIGO === -1) return { ok: false, msg: 'No se encontró columna "Simplified part number" en la primera fila del catálogo.' };

    // Construir mapa del catálogo
    var catalog = {};
    for (var i = 1; i < dCat.length; i++) {
      var row = dCat[i];
      var cod = COL.CODIGO !== -1 ? String(row[COL.CODIGO]||'').trim().toUpperCase() : '';
      if (!cod) continue;
      catalog[cod] = {
        descripcion: COL.DESCRIPCION !== -1 ? String(row[COL.DESCRIPCION]||'').trim() : '',
        categoria:   COL.CATEGORIA   !== -1 ? String(row[COL.CATEGORIA]  ||'').trim() : '',
        modelos:     COL.MODELOS     !== -1 ? String(row[COL.MODELOS]    ||'').trim() : '',
        precioFOB:   COL.PRECIO_FOB  !== -1 ? (parseFloat(String(row[COL.PRECIO_FOB]||'0').replace(/[^0-9.,\-]/g,'').replace(',','.')) || 0) : 0,
        skuDJI:      COL.SKU_DJI     !== -1 ? String(row[COL.SKU_DJI]   ||'').trim() : ''
      };
    }

    var totalCatalog = Object.keys(catalog).length;
    if (totalCatalog === 0) return { ok: false, msg: 'No se encontraron filas con código válido en el catálogo.' };

    // Upsert en DB_REPUESTOS
    // Schema DB_REPUESTOS: [0]ID [1]SKU [2]DESC [3]MODELOS [4]? [5]CATEGORIA [6]PRECIO_FOB
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (!hojaRep) return { ok: false, msg: 'Hoja DB_REPUESTOS no encontrada.' };
    var dRep = hojaRep.getDataRange().getValues();
    var repIdx = {};
    for (var r = 1; r < dRep.length; r++) {
      var skuR = String(dRep[r][1]||'').trim().toUpperCase();
      if (skuR) repIdx[skuR] = r;
    }

    var nuevos = 0, actualizados = 0, appendRows = [];
    var codsOrdenados = Object.keys(catalog);
    for (var k = 0; k < codsOrdenados.length; k++) {
      var cod2 = codsOrdenados[k];
      var c    = catalog[cod2];
      if (repIdx[cod2] !== undefined) {
        var ri = repIdx[cod2];
        var changed = false;
        // Siempre sobreescribir desde el catálogo (corrige datos previos incorrectos)
        if (c.descripcion && dRep[ri][2] !== c.descripcion) { dRep[ri][2] = c.descripcion; changed = true; }
        if (c.modelos     && dRep[ri][3] !== c.modelos)     { dRep[ri][3] = c.modelos;     changed = true; }
        if (c.precioFOB   && dRep[ri][6] !== c.precioFOB)   { dRep[ri][6] = c.precioFOB;   changed = true; }
        // Categoría: solo si es texto válido (no número) — previene que valores numéricos del catálogo DJI contaminen este campo
        if (c.categoria && _catValida(c.categoria) && dRep[ri][5] !== c.categoria) { dRep[ri][5] = c.categoria; changed = true; }
        // Limpiar categoría si actualmente tiene un número inválido
        if (dRep[ri][5] && !_catValida(String(dRep[ri][5]))) { dRep[ri][5] = ''; changed = true; }
        if (changed) actualizados++;
      } else {
        appendRows.push(['', cod2, c.descripcion, c.modelos, '', _catValida(c.categoria) ? c.categoria : '', c.precioFOB]);
        nuevos++;
      }
    }
    if (actualizados > 0) hojaRep.getDataRange().setValues(dRep);
    for (var na = 0; na < appendRows.length; na++) hojaRep.appendRow(appendRows[na]);
    if (nuevos > 0 || actualizados > 0) invalidateSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);

    // Actualizar STOCK_REPUESTOS: sobreescribir modelos y corregir categorías numéricas
    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    var dStock    = hojaStock.getDataRange().getValues();
    var stockChanged = false;
    for (var s = 1; s < dStock.length; s++) {
      var skuS = String(dStock[s][0]||'').trim().toUpperCase();
      if (!skuS) continue;
      // Limpiar categorías numéricas aunque el SKU no esté en el catálogo
      if (dStock[s][4] && !_catValida(String(dStock[s][4]))) { dStock[s][4] = ''; stockChanged = true; }
      if (!catalog[skuS]) continue;
      var cat = catalog[skuS];
      if (cat.modelos && dStock[s][6] !== cat.modelos)                         { dStock[s][6] = cat.modelos;   stockChanged = true; }
      if (cat.categoria && _catValida(cat.categoria) && !dStock[s][4])         { dStock[s][4] = cat.categoria; stockChanged = true; }
    }
    if (stockChanged) {
      hojaStock.getDataRange().setValues(dStock);
      invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    }

    return { ok: true, total: totalCatalog, nuevos: nuevos, actualizados: actualizados };
  } catch(e) {
    Logger.log('sincronizarCatalogoDJI: ' + e);
    return { ok: false, msg: e.toString() };
  }
}

// ============================================================
//  CREAR REPUESTO MANUAL (para piezas DJI no compradas aÍºn)
// ============================================================
function crearRepuesto(data) {
  try {
    var hojaStock = getSheet(SCHEMA.SHEETS.STOCK);
    var hojaRep   = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    var dStock    = getSheetValues(hojaStock);
    var cod       = String(data.codigo||"").trim().toUpperCase();
    if (!cod) return { ok: false, msg: "Código requerido" };

    // Verificar si ya existe en STOCK_REPUESTOS
    for (var i = 1; i < dStock.length; i++) {
      if (String(dStock[i][0]).trim().toUpperCase() === cod)
        return { ok: false, msg: "El código ya existe en el stock" };
    }

    // Agregar a STOCK_REPUESTOS con stock 0
    hojaStock.appendRow([
      cod,
      String(data.descripcion||""),
      0,  // stock inicial 0
      parseInt(data.minimo)||0,
      String(data.categoria||"D").toUpperCase(),
      String(data.ubicacion||""),
      String(data.modelos||""),
      "", ""
    ]);

    // También agregar a DB_REPUESTOS si no existe
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      var existeEnDB = false;
      for (var j = 1; j < dRep.length; j++) {
        if (String(dRep[j][1]||"").trim().toUpperCase() === cod) { existeEnDB=true; break; }
      }
      if (!existeEnDB) {
        hojaRep.appendRow([
          "", cod,
          String(data.descripcion||""),
          String(data.modelos||""),
          "", "", parseFloat(data.precioFOB)||0
        ]);
      }
    }

    return { ok: true };
  } catch(e) {
    Logger.log("crearRepuesto: "+e);
    return { ok: false, msg: e.toString() };
  }
}

function generarHTMLPendientes() {
  try {
    var ss = SpreadsheetApp.openById(MASTER_SHEET_ID); 
    var hojaSolicitudes = ss.getSheetByName(SCHEMA.SHEETS.SOLICITUDES);
    
    if (!hojaSolicitudes) return "Error: No se encontró la hoja " + SCHEMA.SHEETS.SOLICITUDES;

    var datos = getSheetValues(hojaSolicitudes);
    var pendientes = datos.filter(function(fila, index) {
      return index > 0 && String(fila[8]).trim().toLowerCase() === "pendiente";
    });
    
    if (pendientes.length === 0) return "VACIO";

    var html = '<html><head><style>' +
               'body { font-family: Arial, sans-serif; padding: 20px; }' +
               'table { width: 100%; border-collapse: collapse; margin-top: 15px; }' +
               'th, td { border: 1px solid #000; padding: 6px; font-size: 11px; text-align: left; }' +
               'th { background-color: #eee; }' +
               '.urgente { font-weight: bold; text-decoration: underline; }' +
               '@media print { .no-print { display: none; } }' +
               '</style></head><body>' +
               '<h2>PICKING LIST - BIDCOMAGRO</h2>' +
               '<p>Fecha: ' + new Date().toLocaleString() + '</p>' +
               '<table><thead><tr>' +
               '<th>REF / OT</th><th>RESELLER</th><th>CÓDIGO</th><th>DESCRIPCIÓN</th><th>CANT.</th><th>URG.</th>' +
               '</tr></thead><tbody>';

    pendientes.forEach(function(p) {
      html += '<tr>' +
              '<td>' + p[2] + '</td>' +
              '<td>' + p[3] + '</td>' +
              '<td><strong>' + p[4] + '</strong></td>' +
              '<td>' + p[5] + '</td>' +
              '<td style="text-align:center">' + p[6] + '</td>' +
              '<td>' + p[9] + '</td>' +
              '</tr>';
    });

    html += '</tbody></table></body></html>';
    return html;

  } catch (e) {
    return "Error: " + e.toString();
  }
}

// ============================================================
//  TOP ROTACIÓN — SKUs más movidos en los últimos 7 días
// ============================================================
function obtenerTopRotacion() {
  try {
    var corte = new Date(new Date().getTime() - 30 * 24 * 60 * 60 * 1000);
    var conteo = {}, valor = {}, descs = {};
    var wosSS = SpreadsheetApp.openById(WOS_NOTAS_SS_ID);
    var hojas = [wosSS.getSheetByName('Pedidos_resellers'), wosSS.getSheetByName('Pedidos_OTs')].filter(Boolean);
    for (var h = 0; h < hojas.length; h++) {
      var d = hojas[h].getDataRange().getValues();
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][9] || '').trim() === 'Cancelado') continue;
        var sku = String(d[i][2] || '').trim().toUpperCase();
        if (!sku) continue;
        var fecha = d[i][10]; // COL K — FECHA pedido
        if (!(fecha instanceof Date) || fecha < corte) continue;
        var cant   = Number(d[i][4]) || 1;  // CANT_SOL
        var precio = Number(d[i][7]) || 0;  // PRECIO
        conteo[sku] = (conteo[sku] || 0) + cant;
        valor[sku]  = (valor[sku]  || 0) + cant * precio;
        if (!descs[sku]) descs[sku] = String(d[i][3] || '');
      }
    }
    var keys = Object.keys(conteo);
    var byCant = [], byVal = [];
    for (var k = 0; k < keys.length; k++) {
      var sku = keys[k];
      var desc = descs[sku] || sku;
      byCant.push({ codigo: sku, descripcion: desc, movimientos: conteo[sku] });
      if (valor[sku] > 0) byVal.push({ codigo: sku, descripcion: desc, valor: valor[sku] });
    }
    byCant.sort(function(a, b) { return b.movimientos - a.movimientos; });
    byVal.sort(function(a, b)  { return b.valor - a.valor; });
    return { cantidad: byCant.slice(0, 5), valor: byVal.slice(0, 5) };
  } catch(e) { Logger.log('obtenerTopRotacion: ' + e); return { cantidad: [], valor: [] }; }
}

// ============================================================
//  DETALLE DE ITEMS POR CAS
// ============================================================
function obtenerItemsPorCAS(idCas) {
  try {
    var casB = String(idCas).trim().toUpperCase();
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var R    = SCHEMA.RESERVAS_STOCK;
    var tz   = Session.getScriptTimeZone();

    var dMov     = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var recibidos = [];
    for (var i = 1; i < dMov.length; i++) {
      var f = dMov[i];
      if (String(f[M.TIPO] || '').trim() !== 'ENTRADA_COMPRA') continue;
      if (String(f[M.REFERENCIA] || '').trim().toUpperCase() !== casB) continue;
      recibidos.push({
        sku:         String(f[M.CODIGO]      || ''),
        descripcion: String(f[M.DESCRIPCION] || ''),
        cantidad:    Math.abs(parseInt(f[M.CANTIDAD]) || 0),
        fecha:       f[M.FECHA] instanceof Date
                     ? Utilities.formatDate(f[M.FECHA], tz, 'dd/MM/yyyy')
                     : String(f[M.FECHA] || '')
      });
    }

    var reservados = [];
    var hojaRes = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (hojaRes) {
      var dRes = getSheetValues(hojaRes);
      for (var j = 1; j < dRes.length; j++) {
        var r = dRes[j];
        if (String(r[R.CAS_REF] || '').trim().toUpperCase() !== casB) continue;
        if (String(r[R.ESTADO]  || '') !== 'Activa') continue;
        reservados.push({
          sku:        String(r[R.SKU]          || ''),
          descripcion:String(r[R.DESCRIPCION]  || ''),
          cantidad:   parseInt(r[R.CANTIDAD])  || 0,
          origen:     String(r[R.ORIGEN]       || ''),
          referencia: String(r[R.ID_REFERENCIA]|| '')
        });
      }
    }

    return { cas: idCas, recibidos: recibidos, reservados: reservados };
  } catch(e) {
    return { cas: idCas, recibidos: [], reservados: [], error: e.toString() };
  }
}

// ============================================================
//  COMPRAS DETALLE — manifiesto de ítems por CAS
// ============================================================

function crearHojaComprasDetalle() {
  try {
    var db = getDb();
    var existing = db.getSheetByName(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (existing) return { ok: false, msg: 'La hoja ya existe' };
    var hoja = db.insertSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    hoja.appendRow(['ID_CAS','SKU','DESCRIPCION','CANTIDAD_PEDIDA','CANTIDAD_RECIBIDA','ESTADO']);
    hoja.setFrozenRows(1);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function vincularItemsACAS(idCas, listaItems) {
  // listaItems: [{sku, descripcion, cantPedida}]
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    var hoja = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (!hoja) return { ok: false, msg: 'Hoja COMPRAS_DETALLE no existe. Ejecutá crearHojaComprasDetalle() primero.' };
    var casKey = String(idCas).trim().toUpperCase();
    var CD = SCHEMA.COMPRAS_DETALLE;

    // Full replace: delete existing rows for this CAS (iterate backwards)
    var d = hoja.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--) {
      if (String(d[i][CD.ID_CAS]).trim().toUpperCase() === casKey) {
        hoja.deleteRow(i + 1);
      }
    }

    if (listaItems.length) {
      var newRows = [];
      for (var k = 0; k < listaItems.length; k++) {
        var item = listaItems[k];
        newRows.push([
          casKey,
          String(item.sku || '').trim().toUpperCase(),
          String(item.descripcion || ''),
          parseInt(item.cantPedida) || 0,
          0,
          'Pendiente'
        ]);
      }
      var startRow = hoja.getLastRow() + 1;
      hoja.getRange(startRow, 1, newRows.length, 6).setValues(newRows);
      SpreadsheetApp.flush();
    }

    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS_DETALLE);
    return { ok: true, items: listaItems.length };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function obtenerDetalleCAS(idCas) {
  try {
    var casKey = String(idCas).trim().toUpperCase();
    var CD = SCHEMA.COMPRAS_DETALLE;
    var M  = SCHEMA.MOVIMIENTOS_STOCK;
    var R  = SCHEMA.RESERVAS_STOCK;
    var SD = SCHEMA.SOLICITUDES_DESPACHO;
    var tz = Session.getScriptTimeZone();

    // 1. Manifest
    var manifest = [];
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (hojaCD) {
      var dCD = getSheetValues(hojaCD);
      for (var ci = 1; ci < dCD.length; ci++) {
        if (String(dCD[ci][CD.ID_CAS]).trim().toUpperCase() !== casKey) continue;
        manifest.push({
          sku:          String(dCD[ci][CD.SKU]               || ''),
          descripcion:  String(dCD[ci][CD.DESCRIPCION]       || ''),
          cantPedida:   parseInt(dCD[ci][CD.CANTIDAD_PEDIDA])   || 0,
          cantRecibida: parseInt(dCD[ci][CD.CANTIDAD_RECIBIDA]) || 0,
          estado:       String(dCD[ci][CD.ESTADO]            || 'Pendiente')
        });
      }
    }

    // 2. Received (from MOVIMIENTOS ENTRADA_COMPRA)
    var recibidos = [];
    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    for (var mi = 1; mi < dMov.length; mi++) {
      if (String(dMov[mi][M.TIPO]       || '').trim() !== 'ENTRADA_COMPRA') continue;
      if (String(dMov[mi][M.REFERENCIA] || '').trim().toUpperCase() !== casKey) continue;
      recibidos.push({
        sku:         String(dMov[mi][M.CODIGO]      || ''),
        descripcion: String(dMov[mi][M.DESCRIPCION] || ''),
        cantidad:    Math.abs(parseInt(dMov[mi][M.CANTIDAD]) || 0),
        fecha:       dMov[mi][M.FECHA] instanceof Date
                     ? Utilities.formatDate(dMov[mi][M.FECHA], tz, 'dd/MM/yyyy')
                     : String(dMov[mi][M.FECHA] || '')
      });
    }

    // 3. Active reservations
    var reservados = [];
    var hojaRes = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (hojaRes) {
      var dRes = getSheetValues(hojaRes);
      for (var ri = 1; ri < dRes.length; ri++) {
        if (String(dRes[ri][R.CAS_REF] || '').trim().toUpperCase() !== casKey) continue;
        if (String(dRes[ri][R.ESTADO]  || '') !== 'Activa') continue;
        reservados.push({
          sku:        String(dRes[ri][R.SKU]          || ''),
          descripcion:String(dRes[ri][R.DESCRIPCION]  || ''),
          cantidad:   parseInt(dRes[ri][R.CANTIDAD])  || 0,
          origen:     String(dRes[ri][R.ORIGEN]       || ''),
          referencia: String(dRes[ri][R.ID_REFERENCIA]|| '')
        });
      }
    }

    // 4. Coverage: for each manifest item, how many are pending in SOLICITUDES?
    var cobertura = [];
    if (manifest.length) {
      var dSol   = getSheetValues(SCHEMA.SHEETS.SOLICITUDES);
      var pendMap = {};
      for (var si = 1; si < dSol.length; si++) {
        var est = String(dSol[si][SD.ESTADO] || '');
        if (est === 'Despachado' || est === 'Cancelado') continue;
        var skuSol = String(dSol[si][SD.CODIGO] || '').trim().toUpperCase();
        pendMap[skuSol] = (pendMap[skuSol] || 0) + (parseInt(dSol[si][SD.CANT_SOLICITADA]) || 0);
      }
      for (var ci2 = 0; ci2 < manifest.length; ci2++) {
        var skuUp  = manifest[ci2].sku.trim().toUpperCase();
        var pending = pendMap[skuUp] || 0;
        var enCamino = manifest[ci2].cantPedida - manifest[ci2].cantRecibida;
        cobertura.push({ sku: manifest[ci2].sku, pedido: pending, enCamino: enCamino, cubre: enCamino >= pending });
      }
    }

    return { cas: idCas, manifest: manifest, recibidos: recibidos, reservados: reservados, cobertura: cobertura };
  } catch(e) {
    return { cas: idCas, manifest: [], recibidos: [], reservados: [], cobertura: [], error: e.toString() };
  }
}

// ============================================================
//  ANÁLISIS DE VELOCIDAD DE CONSUMO
// ============================================================
function analizarVelocidadConsumo(leadDias) {
  try {
    if (!leadDias) leadDias = 45;
    var hoy    = new Date();
    var corte  = new Date(hoy.getTime() - 30 * 86400000);
    var dMov   = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var dStock = getSheetValues(SCHEMA.SHEETS.STOCK);
    var M = SCHEMA.MOVIMIENTOS_STOCK;
    var S = SCHEMA.STOCK_REPUESTOS;

    var consumo30 = {};
    for (var m = 1; m < dMov.length; m++) {
      var fm = dMov[m];
      if (!(fm[M.FECHA] instanceof Date) || fm[M.FECHA] < corte) continue;
      var tipo = String(fm[M.TIPO] || '');
      if (tipo !== 'SALIDA_DESPACHO' && tipo !== 'EGRESO') continue;
      var cod = String(fm[M.CODIGO] || '').trim().toUpperCase();
      if (!cod) continue;
      consumo30[cod] = (consumo30[cod] || 0) + Math.abs(parseInt(fm[M.CANTIDAD]) || 0);
    }

    var out = [];
    for (var s = 1; s < dStock.length; s++) {
      if (!dStock[s][S.CODIGO]) continue;
      var codS    = String(dStock[s][S.CODIGO]).trim().toUpperCase();
      var actS    = parseInt(dStock[s][S.STOCK_ACTUAL]) || 0;
      var unids   = consumo30[codS] || 0;
      var burnDay = unids / 30;
      var diasR   = burnDay > 0 ? Math.round(actS / burnDay) : null;
      var flagged = diasR !== null && diasR < leadDias && actS > 0;
      out.push({
        codigo:        codS,
        descripcion:   String(dStock[s][S.DESCRIPCION] || ''),
        stockActual:   actS,
        unidades30d:   unids,
        burnRatePerDay:Math.round(burnDay * 100) / 100,
        diasRestantes: diasR,
        sugeridoCompra:flagged
      });
    }
    out.sort(function(a, b) {
      if (a.diasRestantes === null && b.diasRestantes === null) return 0;
      if (a.diasRestantes === null) return 1;
      if (b.diasRestantes === null) return -1;
      return a.diasRestantes - b.diasRestantes;
    });
    return out;
  } catch(e) { return []; }
}

// ============================================================
//  RESERVAS DE STOCK EN TRÁNSITO
// ============================================================

function crearHojaReservas() {
  try {
    var db = getDb();
    var existing = db.getSheetByName(SCHEMA.SHEETS.RESERVAS);
    if (existing) return { ok: false, msg: 'La hoja ya existe' };
    var hoja = db.insertSheet(SCHEMA.SHEETS.RESERVAS);
    hoja.appendRow(['ID','Fecha','SKU','Descripción','Cantidad','Origen','ID_Referencia','Estado','CAS_Ref','Operador','Observaciones']);
    hoja.setFrozenRows(1);
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function crearReserva(sku, descripcion, cantidad, origen, idReferencia, casRef, operador) {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.RESERVAS);
    if (!hoja) return { ok: false, msg: 'Hoja RESERVAS_STOCK no encontrada. Ejecutar crearHojaReservas() primero.' };
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var id = 'RES-' + new Date().getTime();
      var R  = SCHEMA.RESERVAS_STOCK;
      var row = new Array(11);
      row[R.ID]            = id;
      row[R.FECHA]         = new Date();
      row[R.SKU]           = String(sku).trim().toUpperCase();
      row[R.DESCRIPCION]   = String(descripcion || '');
      row[R.CANTIDAD]      = parseInt(cantidad) || 1;
      row[R.ORIGEN]        = String(origen || 'VENTAS');
      row[R.ID_REFERENCIA] = String(idReferencia || '');
      row[R.ESTADO]        = 'Activa';
      row[R.CAS_REF]       = String(casRef || '');
      row[R.OPERADOR]      = String(operador || '');
      row[R.OBSERVACIONES] = '';
      hoja.appendRow(row);
      invalidateSheetValues(SCHEMA.SHEETS.RESERVAS);
      return { ok: true, id: id };
    } finally {
      if (lock.hasLock()) lock.releaseLock();
    }
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function obtenerReservasActivas() {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.RESERVAS);
    if (!datos || datos.length < 2) return [];
    var R   = SCHEMA.RESERVAS_STOCK;
    var tz  = Session.getScriptTimeZone();
    var out = [];
    for (var i = 1; i < datos.length; i++) {
      var f = datos[i];
      if (String(f[R.ESTADO]) !== 'Activa') continue;
      out.push({
        id:          String(f[R.ID]),
        fecha:       f[R.FECHA] instanceof Date ? Utilities.formatDate(f[R.FECHA], tz, 'dd/MM/yyyy') : String(f[R.FECHA]),
        sku:         String(f[R.SKU]),
        descripcion: String(f[R.DESCRIPCION]),
        cantidad:    parseInt(f[R.CANTIDAD]) || 0,
        origen:      String(f[R.ORIGEN]),
        referencia:  String(f[R.ID_REFERENCIA]),
        casRef:      String(f[R.CAS_REF]),
        operador:    String(f[R.OPERADOR])
      });
    }
    return out;
  } catch(e) { return []; }
}

function obtenerCASEnTransitoSM() {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var out   = [];
    for (var i = 1; i < datos.length; i++) {
      var f      = datos[i];
      var estado = String(f[2] || '').trim();
      if (estado === 'En depósito' || estado.indexOf('Borrador') !== -1) continue;
      out.push({ cas: String(f[0] || ''), estado: estado });
    }
    return out;
  } catch(e) { return []; }
}

// Devuelve OTs en "Espera de repuestos" cruzadas con los CAS que traen sus repuestos
function obtenerOTsBloqueadasConCAS() {
  try {
    var ESTADOS_CERRADOS = ['Entregado_Cerrado', 'Cancelado', 'Entregado_Confirmado'];
    var tz = Session.getScriptTimeZone();

    // Mapa CAS → estado (solo activos en tránsito)
    var dCAS = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var casEstadoMap = {};
    for (var ci = 1; ci < dCAS.length; ci++) {
      var casId  = String(dCAS[ci][0] || '').trim().toUpperCase();
      var casEst = String(dCAS[ci][2] || '').trim();
      if (casId && casEst !== 'En depósito' && casEst.indexOf('Borrador') === -1)
        casEstadoMap[casId] = casEst;
    }

    // Mapa SKU → lista de CAS pendientes que lo traen
    var skuCasMap = {};
    var hojaCD = getSheet(SCHEMA.SHEETS.COMPRAS_DETALLE);
    if (hojaCD) {
      var dCD = getSheetValues(hojaCD);
      var CD  = SCHEMA.COMPRAS_DETALLE;
      for (var cdi = 1; cdi < dCD.length; cdi++) {
        var cdCas = String(dCD[cdi][CD.ID_CAS] || '').trim().toUpperCase();
        var cdSku = String(dCD[cdi][CD.SKU]    || '').trim().toUpperCase();
        var cdPed = parseInt(dCD[cdi][CD.CANTIDAD_PEDIDA])   || 0;
        var cdRec = parseInt(dCD[cdi][CD.CANTIDAD_RECIBIDA]) || 0;
        var cdEst = casEstadoMap[cdCas];
        if (!cdSku || !cdEst) continue;
        var pend = Math.max(0, cdPed - cdRec);
        if (pend > 0) {
          if (!skuCasMap[cdSku]) skuCasMap[cdSku] = [];
          skuCasMap[cdSku].push({ casId: cdCas, estado: cdEst, cant: pend });
        }
      }
    }

    // Fuente de verdad: Pedidos_OTs en WOS
    // Agrupamos por numero; filtramos estados cerrados y cantPend > 0
    var pedMap = {};
    try {
      var wosHoja = SpreadsheetApp.openById(WOS_NOTAS_SS_ID).getSheetByName('Pedidos_OTs');
      if (wosHoja) {
        var wosData = wosHoja.getDataRange().getValues();
        for (var wi = 1; wi < wosData.length; wi++) {
          var wNum  = String(wosData[wi][0] || '').trim();
          var wEst  = String(wosData[wi][9] || '').trim();
          if (!wNum || ESTADOS_CERRADOS.indexOf(wEst) !== -1) continue;
          var wPend = Number(wosData[wi][6]) || 0;
          if (!pedMap[wNum]) {
            var fdRaw = wosData[wi][14];
            pedMap[wNum] = {
              reseller:    String(wosData[wi][1]  || ''),
              estado:      wEst,
              cantDesp:    Number(wosData[wi][5]) || 0,
              cantPend:    0,
              notaEntrega: String(wosData[wi][15] || '').trim(),
              fechaDesp:   (fdRaw instanceof Date) ? Utilities.formatDate(fdRaw, tz, 'dd/MM/yyyy') : '',
              skus:        []
            };
          }
          pedMap[wNum].cantPend += wPend;
          var sku  = String(wosData[wi][2] || '').trim().toUpperCase();
          var desc = String(wosData[wi][3] || '').trim();
          if (sku) {
            pedMap[wNum].skus.push({ sku: sku, desc: desc, cantPend: wPend, cas: skuCasMap[sku] || null });
          }
        }
      }
    } catch(eWOS) { Logger.log('obtenerOTsBloqueadasConCAS WOS: ' + eWOS); }

    var out = [];
    var nums = Object.keys(pedMap);
    for (var ni = 0; ni < nums.length; ni++) {
      var p = pedMap[nums[ni]];
      if (p.cantPend <= 0) continue;
      out.push({
        ot:       nums[ni],
        reseller: p.reseller,
        skus:     p.skus,
        wos:      { estado: p.estado, cantDesp: p.cantDesp, cantPend: p.cantPend,
                    notaEntrega: p.notaEntrega, fechaDesp: p.fechaDesp }
      });
    }
    out.sort(function(a, b) { return b.wos.cantPend - a.wos.cantPend; });
    return out.slice(0, 20);
  } catch(e) {
    Logger.log('obtenerOTsBloqueadasConCAS: ' + e);
    return [];
  }
}

// ============================================================
//  LOOP AUTOMÁTICO — BORRADOR DE COMPRA POR BACKORDER
//  Detecta SKUs en quiebre con OTs bloqueadas y crea un
//  borrador en COMPRAS_DJI + detalle en COMPRAS_DETALLE.
//  Llama cada vez que el operador quiere disparar la compra.
// ============================================================
function crearBorradorCompraAutomatico(operador) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!operador) {
      try { operador = Session.getActiveUser().getEmail(); } catch(eu) { operador = 'Sistema'; }
    }

    var alertas = obtenerAlertasStockCritico();
    // Solo quiebres con OTs bloqueadas
    var candidatos = [];
    for (var a = 0; a < alertas.length; a++) {
      if (alertas[a].estado === 'QUIEBRE' && alertas[a].bloqueadas > 0) {
        candidatos.push(alertas[a]);
      }
    }
    if (!candidatos.length) return { ok: false, msg: 'No hay SKUs en quiebre con OTs bloqueadas.' };

    var tz  = Session.getScriptTimeZone();
    var hoy = new Date();
    var cas = 'BORRADOR-AUTO-' + Utilities.formatDate(hoy, tz, 'yyyyMMdd-HHmm');

    var hojaCompras = getSheet(SCHEMA.SHEETS.COMPRAS);
    var dComp = getSheetValues(hojaCompras);
    // Verificar duplicado del mismo día (evita doble click)
    var prefixHoy = 'BORRADOR-AUTO-' + Utilities.formatDate(hoy, tz, 'yyyyMMdd');
    for (var ci = 1; ci < dComp.length; ci++) {
      var casExist = String(dComp[ci][0] || '').trim().toUpperCase();
      if (casExist.indexOf(prefixHoy.toUpperCase()) === 0) {
        return { ok: false, msg: 'Ya existe un borrador automático para hoy: ' + casExist + '. Eliminalo o editalo desde la sección Compras.' };
      }
    }

    // Crear fila en COMPRAS_DJI con estado "Borrador"
    var obsItems = candidatos.map(function(c) { return c.codigo + ' (' + c.bloqueadas + ' OTs)'; }).join(', ');
    hojaCompras.appendRow([
      cas.toUpperCase(),
      hoy,
      'Borrador',
      '',
      '', '', '', '', '', '', '',
      operador || 'Sistema',
      'Auto-generado. SKUs bloqueados: ' + obsItems
    ]);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);

    // Escribir ítems en COMPRAS_DETALLE
    var listaItems = candidatos.map(function(c) {
      return {
        sku:        c.codigo,
        descripcion:c.descripcion,
        cantPedida: c.bloqueadas  // 1 unidad por OT bloqueada como mínimo
      };
    });
    var resVinc = vincularItemsACAS(cas.toUpperCase(), listaItems);
    if (!resVinc.ok) {
      Logger.log('crearBorradorCompraAutomatico: vincularItemsACAS falló: ' + resVinc.msg);
    }

    // Notificar al supervisor
    _notificarBorradorAutoCompra(cas.toUpperCase(), candidatos, operador);

    return { ok: true, cas: cas.toUpperCase(), items: candidatos.length };
  } catch(e) {
    Logger.log('crearBorradorCompraAutomatico: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

function _notificarBorradorAutoCompra(cas, candidatos, operador) {
  try {
    var filas = candidatos.map(function(c) {
      return "<tr>" +
        "<td style='padding:6px 10px;font-size:12px;font-weight:700;color:#e53935'>" + c.codigo + "</td>" +
        "<td style='padding:6px 10px;font-size:12px'>" + c.descripcion + "</td>" +
        "<td style='padding:6px 10px;font-size:12px;text-align:center'>" + c.bloqueadas + "</td>" +
      "</tr>";
    }).join('');
    GmailApp.sendEmail(SM_CONFIG.EMAIL_SUPERVISOR,
      '[Stock Manager] Borrador de compra generado — ' + candidatos.length + ' SKU(s) bloqueados',
      '', {
        htmlBody:
          "<div style='font-family:sans-serif;max-width:640px'>" +
          "<div style='background:#e53935;padding:16px 20px;border-radius:8px 8px 0 0'>" +
            "<span style='color:#fff;font-size:16px;font-weight:700'>🛒 Borrador de compra auto-generado</span>" +
          "</div>" +
          "<div style='background:#fff;border:1px solid #ddd;padding:18px 20px;border-radius:0 0 8px 8px'>" +
            "<p style='font-size:13px;color:#444;margin:0 0 6px'>CAS: <strong>" + cas + "</strong></p>" +
            "<p style='font-size:13px;color:#444;margin:0 0 14px'>Operador: " + (operador || 'Sistema') + "</p>" +
            "<p style='font-size:13px;color:#444;margin:0 0 10px'>Los siguientes SKUs están en <strong>quiebre total</strong> y tienen OTs esperando esos repuestos:</p>" +
            "<table style='width:100%;border-collapse:collapse;border:1px solid #eee'>" +
              "<thead><tr style='background:#f5f5f5'>" +
                "<th style='padding:6px 10px;font-size:11px;text-align:left'>SKU</th>" +
                "<th style='padding:6px 10px;font-size:11px;text-align:left'>Descripción</th>" +
                "<th style='padding:6px 10px;font-size:11px;text-align:center'>OTs bloqueadas</th>" +
              "</tr></thead>" +
              "<tbody>" + filas + "</tbody>" +
            "</table>" +
            "<p style='font-size:12px;color:#666;margin-top:14px'>Ingresá al <strong>Stock Manager → Compras</strong> para revisar y confirmar el pedido.</p>" +
          "</div></div>",
        name: SM_CONFIG.NOMBRE_REMITENTE,
        replyTo: SM_CONFIG.EMAIL_SUPERVISOR
      });
  } catch(e) { Logger.log('_notificarBorradorAutoCompra: ' + e); }
}


// ============================================================
//  LEDGER ENGINE — EVENT SOURCING LAYER
//
//  Arquitectura CQRS en GAS:
//    Write side : MOVIMIENTOS_STOCK (append-only, fuente de verdad)
//    Read side  : STOCK_REPUESTOS   (vista materializada, write-through
//                                    cache actualizada dentro del lock)
//
//  obtenerSaldoRealEnMemoria → herramienta de AUDITORÍA y reconciliación.
//                              NO es el hot-path de validación transaccional.
//                              En GAS un scan O(n) sobre el ledger completo
//                              toma 2-5s — ejecutarlo en cada transacción de
//                              un lote de 50 ítems agotaría el límite de 6min.
//
//  registrarEventoLedgerSeguro → punto de entrada único para toda mutación
//                                de inventario. Reemplaza registrarMovimientoSeguro.
//
//  reconciliarCacheDesdeMovimientos → mantenimiento nocturno / auditoría manual.
//                                     Jamás llamar desde el hot-path.
// ============================================================

/**
 * Calcula el saldo exacto de un SKU leyendo la historia completa del ledger.
 *
 * CUÁNDO USAR:
 *   - Botón "Auditar" en la UI
 *   - Trigger nocturno de reconciliación
 *   - Post-mortem de una discrepancia detectada
 *
 * CUÁNDO NO USAR:
 *   - Dentro de registrarEventoLedgerSeguro (usa la vista materializada bajo lock)
 *   - En lotes: 50 SKUs × O(n) scan = timeout garantizado
 *
 * @param  {string} sku  Código del repuesto (normalizado a mayúsculas internamente)
 * @return {{ ok: boolean, saldo: number, movimientos: number, error?: string }}
 */
function obtenerSaldoRealEnMemoria(sku) {
  var skuKey = String(sku || '').trim().toUpperCase();
  if (!skuKey) {
    Logger.log('[LEDGER] obtenerSaldoRealEnMemoria: SKU vacío');
    return { ok: false, saldo: 0, movimientos: 0, error: 'WMS_SKU_VACIO' };
  }
  var M = SCHEMA.MOVIMIENTOS_STOCK;
  try {
    var dMov      = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var saldo     = 0;
    var countMov  = 0;
    for (var i = 1; i < dMov.length; i++) {
      if (String(dMov[i][M.CODIGO] || '').trim().toUpperCase() !== skuKey) continue;
      // CANTIDAD ya viene con signo: positivo = entrada, negativo = salida
      var delta = parseFloat(dMov[i][M.CANTIDAD]) || 0;
      saldo    += delta;
      countMov++;
    }
    return { ok: true, saldo: Math.round(saldo), movimientos: countMov };
  } catch(e) {
    var snap = '[LEDGER][obtenerSaldoRealEnMemoria] sku=' + skuKey + ' err=' + e.message;
    Logger.log(snap);
    console.error(snap);
    return { ok: false, saldo: 0, movimientos: 0, error: e.message };
  }
}

/**
 * Punto de entrada ÚNICO para toda mutación de inventario.
 * Patrón: Lock → Verify (cache) → Append (ledger) → Sync (cache) → Release
 *
 * Errores nombrados retornados en { ok: false, error: 'WMS_*' }:
 *   WMS_SKU_VACIO           — parámetro sku ausente
 *   WMS_DELTA_CERO          — delta = 0, operación sin efecto
 *   WMS_SKU_NOT_FOUND       — SKU no existe en STOCK_REPUESTOS
 *   WMS_INSUFFICIENT_FUNDS  — egreso supera el saldo disponible
 *   WMS_LOCK_TIMEOUT        — no se obtuvo el lock en 15 segundos
 *   WMS_LEDGER_WRITE_FAIL   — appendRow al ledger falló después de adquirir lock
 *
 * @param  {Object} params
 *   .sku         {string}  Código del repuesto
 *   .delta       {number}  +N = entrada, −N = salida (con signo)
 *   .tipo        {string}  SALIDA_DESPACHO | ENTRADA_COMPRA | EGRESO | AJUSTE_INVENTARIO | …
 *   .referencia  {string}  OT / CAS / Factura / motivo
 *   .operador    {string}  Email del operador
 *   .descripcion {string}  (opcional) Se resuelve desde la vista si está vacío
 * @return {{ ok: boolean, saldoResultante?: number, error?: string, snapshot?: Object }}
 */
function registrarEventoLedgerSeguro(params) {
  var sku         = String(params.sku         || '').trim().toUpperCase();
  var delta       = parseFloat(params.delta)  || 0;
  var tipo        = String(params.tipo        || 'EGRESO').trim();
  var referencia  = String(params.referencia  || '');
  var operador    = String(params.operador    || '');
  var descripcion = String(params.descripcion || '');

  // Validaciones previas al lock (fail-fast sin costo de I/O)
  if (!sku)     return { ok: false, error: 'WMS_SKU_VACIO' };
  if (delta === 0) return { ok: false, error: 'WMS_DELTA_CERO' };

  var lock     = LockService.getScriptLock();
  var snapshot = {
    ts:         new Date().toISOString(),
    sku:        sku,
    delta:      delta,
    tipo:       tipo,
    referencia: referencia,
    operador:   operador
  };

  try {
    if (!lock.tryLock(15000)) {
      snapshot.error = 'WMS_LOCK_TIMEOUT';
      console.error('[LEDGER][LOCK_TIMEOUT] ' + JSON.stringify(snapshot));
      Logger.log('[LEDGER][LOCK_TIMEOUT] ' + JSON.stringify(snapshot));
      return { ok: false, error: 'WMS_LOCK_TIMEOUT', snapshot: snapshot };
    }

    // 1. Leer vista materializada — O(1) lookup, autoritativa bajo lock
    var hojaStr = getSheet(SCHEMA.SHEETS.STOCK);
    var dStr    = getSheetValues(hojaStr);
    var S       = SCHEMA.STOCK_REPUESTOS;
    var filaIdx = -1;
    for (var i = 1; i < dStr.length; i++) {
      if (String(dStr[i][S.CODIGO] || '').trim().toUpperCase() === sku) {
        filaIdx = i;
        break;
      }
    }
    if (filaIdx === -1) {
      snapshot.error = 'WMS_SKU_NOT_FOUND';
      console.error('[LEDGER][SKU_NOT_FOUND] ' + JSON.stringify(snapshot));
      Logger.log('[LEDGER][SKU_NOT_FOUND] ' + JSON.stringify(snapshot));
      return { ok: false, error: 'WMS_SKU_NOT_FOUND', snapshot: snapshot };
    }

    var saldoCache = parseInt(dStr[filaIdx][S.STOCK_ACTUAL]) || 0;
    if (!descripcion) descripcion = String(dStr[filaIdx][S.DESCRIPCION] || '');
    snapshot.saldoCache = saldoCache;

    // 2. Gatekeeper para egresos — usa la vista materializada (correcta bajo lock)
    //    La vista es write-through: fue escrita por el último lock holder y no
    //    puede divergir mientras este lock esté activo.
    if (delta < 0) {
      var cantSolicitada = Math.abs(delta);
      if (saldoCache < cantSolicitada) {
        snapshot.error         = 'WMS_INSUFFICIENT_FUNDS';
        snapshot.cantSolicitada = cantSolicitada;
        console.error('[LEDGER][INSUFFICIENT_FUNDS] ' + JSON.stringify(snapshot));
        Logger.log('[LEDGER][INSUFFICIENT_FUNDS] ' + JSON.stringify(snapshot));
        return { ok: false, error: 'WMS_INSUFFICIENT_FUNDS', snapshot: snapshot };
      }
    }

    // 3. Calcular saldo resultante
    var saldoNuevo = saldoCache + delta;
    var ahora      = new Date();

    // 4. Append al ledger — operación inmutable, fuente de verdad
    var hojaLedger = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
    if (!hojaLedger) {
      snapshot.error = 'WMS_LEDGER_WRITE_FAIL';
      console.error('[LEDGER][NO_LEDGER_SHEET] ' + JSON.stringify(snapshot));
      Logger.log('[LEDGER][NO_LEDGER_SHEET] ' + JSON.stringify(snapshot));
      return { ok: false, error: 'WMS_LEDGER_WRITE_FAIL', snapshot: snapshot };
    }
    hojaLedger.appendRow([
      ahora,
      tipo,
      sku,
      descripcion,
      delta,
      saldoNuevo,
      referencia,
      operador,
      ''
    ]);

    // 5. Sync write-through: actualizar vista materializada (la misma celda, no rebuild)
    hojaStr.getRange(filaIdx + 1, S.STOCK_ACTUAL + 1).setValue(saldoNuevo);
    if (delta > 0) hojaStr.getRange(filaIdx + 1, S.ULTIMA_ENTRADA + 1).setValue(ahora);
    if (delta < 0) hojaStr.getRange(filaIdx + 1, S.ULTIMA_SALIDA  + 1).setValue(ahora);

    // Carmen se actualiza vía Entregados/Recibidos + fórmula en col C — no tocar con setValue

    // 6. Invalidar caché de getSheetValues para que la próxima lectura sea fresca
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);
    invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    SpreadsheetApp.flush();

    return { ok: true, saldoResultante: saldoNuevo };

  } catch(e) {
    snapshot.error   = e.message;
    snapshot.stack   = e.stack || '(no stack)';
    console.error('[LEDGER][EXCEPTION] ' + JSON.stringify(snapshot));
    Logger.log('[LEDGER][EXCEPTION] ' + JSON.stringify(snapshot));
    return { ok: false, error: e.message, snapshot: snapshot };
  } finally {
    // El lock SIEMPRE se libera — incluso si appendRow falla a mitad
    if (lock.hasLock()) lock.releaseLock();
  }
}

/**
 * Reconcilia la vista materializada (STOCK_REPUESTOS) contra el ledger.
 * Detecta y corrige drift entre caché y fuente de verdad.
 *
 * EJECUTAR: manualmente desde el editor, o vía trigger nocturno.
 * JAMÁS llamar desde el hot-path de una transacción.
 *
 * @return {{ ok: boolean, reconciliados: number, totalSkus: number, discrepancias: Array }}
 */
function reconciliarCacheDesdeMovimientos() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Lock largo solo permitido en operación de mantenimiento
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var S    = SCHEMA.STOCK_REPUESTOS;
    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);

    // Acumular suma de deltas por SKU desde el ledger completo
    var saldosLedger = {};
    for (var i = 1; i < dMov.length; i++) {
      var cod = String(dMov[i][M.CODIGO] || '').trim().toUpperCase();
      if (!cod) continue;
      saldosLedger[cod] = (saldosLedger[cod] || 0) + (parseFloat(dMov[i][M.CANTIDAD]) || 0);
    }

    // Comparar contra la vista materializada y corregir drift
    var hojaStr  = getSheet(SCHEMA.SHEETS.STOCK);
    var dStr     = hojaStr.getDataRange().getValues();
    var discrepancias = [];
    var reconciliados = 0;

    for (var r = 1; r < dStr.length; r++) {
      var sku = String(dStr[r][S.CODIGO] || '').trim().toUpperCase();
      if (!sku || saldosLedger[sku] === undefined) continue;
      var saldoLedger = Math.max(0, Math.round(saldosLedger[sku]));
      var saldoCache  = parseInt(dStr[r][S.STOCK_ACTUAL]) || 0;
      if (saldoCache !== saldoLedger) {
        discrepancias.push({ sku: sku, cache: saldoCache, ledger: saldoLedger, drift: saldoLedger - saldoCache });
        dStr[r][S.STOCK_ACTUAL] = saldoLedger;
        reconciliados++;
      }
    }

    if (reconciliados > 0) {
      hojaStr.getDataRange().setValues(dStr);
      SpreadsheetApp.flush();
      invalidateSheetValues(SCHEMA.SHEETS.STOCK);
      Logger.log('[RECONCILIAR] ' + reconciliados + ' discrepancias corregidas: ' + JSON.stringify(discrepancias));
    } else {
      Logger.log('[RECONCILIAR] Caché consistente con el ledger. Sin drift detectado.');
    }

    return {
      ok:             true,
      reconciliados:  reconciliados,
      totalSkus:      Object.keys(saldosLedger).length,
      discrepancias:  discrepancias
    };
  } catch(e) {
    var snap = '[RECONCILIAR][EXCEPTION] ' + e.message;
    console.error(snap);
    Logger.log(snap);
    return { ok: false, error: e.message };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// Retrocompatibilidad: registrarMovimientoSeguro delega al ledger engine.
// Permite migración gradual sin romper callers existentes.
// ============================================================

// ============================================================
//  HELPERS INTERNOS
// ============================================================

// Devuelve true si el valor es una categoría de texto válida (no un número puro).
// Previene que columnas numéricas del Excel DJI (unit, pcs) contaminen el campo categoría.
function _catValida(v) {
  var s = String(v || '').trim();
  return s.length > 0 && isNaN(Number(s));
}

/**
 * Extrae los nombres de modelos únicos desde DB_REPUESTOS (col 3 = MODELOS).
 * Separa por coma, barra, punto y coma y retorna lista ordenada sin duplicados.
 * Usada para popular el datalist de autocomplete en el generador de pedido.
 */
function getModelosDisponibles() {
  try {
    var dRep = getSheetValues(SCHEMA.SHEETS.DB_REPUESTOS);
    var mapa = {};
    for (var i = 1; i < dRep.length; i++) {
      var raw = String(dRep[i][3] || '').trim();
      if (!raw) continue;
      var partes = raw.split(/[,\/;|]+/);
      for (var p = 0; p < partes.length; p++) {
        var m = partes[p].trim();
        if (m && m.length > 1) mapa[m] = true;
      }
    }
    return Object.keys(mapa).sort();
  } catch(e) {
    Logger.log('getModelosDisponibles: ' + e);
    return [];
  }
}
function _registrarMovimiento(tipo, codigo, desc, cantidad, stockResult, referencia, operador, deposito) {
  var hojaLedger = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
  if (!hojaLedger) {
    var err = '[_registrarMovimiento] Hoja MOVIMIENTOS_STOCK no encontrada. tipo=' + tipo + ' sku=' + codigo;
    Logger.log(err);
    console.error(err);
    throw new Error('WMS_LEDGER_WRITE_FAIL: ' + err);
  }
  hojaLedger.appendRow([
    new Date(), tipo, codigo, desc, cantidad, stockResult,
    referencia || '', operador || '', '', String(deposito || 'BA')
  ]);
}

function _fmtFecha(v) {
  if (!v || !(v instanceof Date)) return "";
  try { return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy"); }
  catch(e) { return ""; }
}

// ============================================================
//  CLASIFICACIÓN ABC — Reclasificación mensual por rotación
// ============================================================

// Recalcula la CATEGORIA (A/B/C/D) de cada SKU en STOCK_REPUESTOS
// aplicando análisis Pareto sobre las salidas de los últimos 90 días.
// A = top 80% del consumo acumulado · B = 80–95% · C = 95–100% · D = sin movimiento.
// Retorna { ok, actualizado, porCategoria: {A,B,C,D}, fechaEjecucion }
function recalcularClasificacionABC() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    var hoy   = new Date();
    var corte = new Date(hoy.getTime() - 90 * 86400000);
    var M     = SCHEMA.MOVIMIENTOS_STOCK;
    var S     = SCHEMA.STOCK_REPUESTOS;

    // 1. Sumar unidades de salida por SKU en los últimos 90 días
    var dMov  = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var consumo = {};
    for (var m = 1; m < dMov.length; m++) {
      var f = dMov[m];
      if (!(f[M.FECHA] instanceof Date) || f[M.FECHA] < corte) continue;
      var tipo = String(f[M.TIPO] || '');
      if (tipo !== 'SALIDA_DESPACHO' && tipo !== 'EGRESO') continue;
      var cod = String(f[M.CODIGO] || '').trim().toUpperCase();
      if (!cod) continue;
      consumo[cod] = (consumo[cod] || 0) + Math.abs(parseInt(f[M.CANTIDAD]) || 0);
    }

    // 2. Leer stock sin caché para garantizar datos frescos
    var hStock = getSheet(SCHEMA.SHEETS.STOCK);
    var dStock = hStock.getDataRange().getValues();
    var nRows  = dStock.length - 1;
    if (nRows <= 0) return { ok: false, msg: 'Sin datos en stock' };

    // 3. Ordenar SKUs por consumo descendente para el corte Pareto
    var skuList = [];
    for (var i = 1; i < dStock.length; i++) {
      var sku = String(dStock[i][S.CODIGO] || '').trim().toUpperCase();
      if (!sku) continue;
      skuList.push({ sku: sku, cons: consumo[sku] || 0 });
    }
    skuList.sort(function(a, b) { return b.cons - a.cons; });

    var totalConsumo = 0;
    for (var k = 0; k < skuList.length; k++) totalConsumo += skuList[k].cons;

    // 4. Asignar categorías acumulando el porcentaje del total
    var mapaCategoria = {};
    var acum = 0;
    for (var j = 0; j < skuList.length; j++) {
      var item = skuList[j];
      if (item.cons === 0 || totalConsumo === 0) { mapaCategoria[item.sku] = 'D'; continue; }
      acum += item.cons;
      var pct = acum / totalConsumo;
      mapaCategoria[item.sku] = pct <= 0.80 ? 'A' : pct <= 0.95 ? 'B' : 'C';
    }

    // 5. Construir columna de categorías y escribir en bulk (solo col 5)
    var colData = [];
    var conteo  = { A: 0, B: 0, C: 0, D: 0 };
    for (var r = 1; r < dStock.length; r++) {
      var skuR = String(dStock[r][S.CODIGO] || '').trim().toUpperCase();
      var cat  = (skuR && mapaCategoria[skuR]) ? mapaCategoria[skuR] : 'D';
      colData.push([cat]);
      conteo[cat] = (conteo[cat] || 0) + 1;
    }

    hStock.getRange(2, S.CATEGORIA + 1, colData.length, 1).setValues(colData);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.STOCK);

    _logClasificacion(hoy, conteo, colData.length);

    return {
      ok:            true,
      actualizado:   colData.length,
      porCategoria:  conteo,
      fechaEjecucion: Utilities.formatDate(hoy, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    };
  } catch(e) {
    Logger.log('recalcularClasificacionABC: ' + e);
    return { ok: false, msg: e.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// Persiste el resultado de la reclasificación en la hoja LOGS.
function _logClasificacion(fecha, conteo, total) {
  try {
    var hLogs = getSheet('LOGS');
    if (!hLogs) {
      hLogs = getSS().insertSheet('LOGS');
      hLogs.appendRow(['Fecha', 'Evento', 'Detalle']);
    }
    hLogs.appendRow([
      fecha, 'RECLASIFICACION_ABC',
      'Total:' + total + ' A:' + (conteo.A||0) + ' B:' + (conteo.B||0) + ' C:' + (conteo.C||0) + ' D:' + (conteo.D||0)
    ]);
  } catch(e) { Logger.log('_logClasificacion: ' + e); }
}

// Retorna la última ejecución de la clasificación ABC desde LOGS, o null.
function obtenerUltimaClasificacion() {
  try {
    var hLogs = getSheet('LOGS');
    if (!hLogs) return null;
    var d = hLogs.getDataRange().getValues();
    for (var i = d.length - 1; i >= 1; i--) {
      if (String(d[i][1]) !== 'RECLASIFICACION_ABC') continue;
      var det = String(d[i][2]);
      var out = {
        fecha: d[i][0] instanceof Date
          ? Utilities.formatDate(d[i][0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
          : String(d[i][0]),
        porCategoria: {}
      };
      var parts = det.split(' ');
      for (var p = 0; p < parts.length; p++) {
        var kv = parts[p].split(':');
        if (kv.length === 2 && 'ABCD'.indexOf(kv[0]) !== -1 && kv[0].length === 1) {
          out.porCategoria[kv[0]] = parseInt(kv[1]) || 0;
        }
      }
      return out;
    }
    return null;
  } catch(e) { return null; }
}

// Instala un trigger que ejecuta recalcularClasificacionABC() el día 1 de cada mes a las 02:00.
// Llamar una sola vez desde el editor GAS o desde el botón de configuración inicial.
function configurarTriggerMensualABC() {
  var existentes = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existentes.length; i++) {
    if (existentes[i].getHandlerFunction() === 'recalcularClasificacionABC') {
      ScriptApp.deleteTrigger(existentes[i]);
    }
  }
  ScriptApp.newTrigger('recalcularClasificacionABC')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();
  Logger.log('Trigger mensual ABC configurado: día 1 de cada mes a las 02:00');
  return { ok: true };
}

// ── TEST — ejecutar desde el editor GAS, eliminar después de verificar ──────
function TEST_generarValeMovimiento() {
  // 1. Insertar fila de prueba en MOVIMIENTOS_STOCK
  var hojaMov = getSheet(SCHEMA.SHEETS.MOVIMIENTOS);
  hojaMov.appendRow([
    new Date(), 'EGRESO', 'TEST-SKU-001', 'Repuesto de prueba PDF',
    -1, 99, 'OT #TEST-001', Session.getActiveUser().getEmail(), 'Fila de prueba — eliminar'
  ]);
  SpreadsheetApp.flush();
  invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);

  // 2. Obtener la fila recién insertada
  var filaTest = hojaMov.getLastRow();
  Logger.log('Fila de prueba insertada: ' + filaTest);

  // 3. Generar el vale PDF
  var resultado = generarValeMovimiento(filaTest);
  Logger.log(JSON.stringify(resultado));

  if (resultado.success) {
    Logger.log('✓ PDF generado correctamente: ' + resultado.url);
  } else {
    Logger.log('✗ Error: ' + resultado.error);
  }

  // 4. Limpiar la fila de prueba
  hojaMov.deleteRow(filaTest);
  invalidateSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
  Logger.log('Fila de prueba eliminada.');
}

// ============================================================
//  LEAD TIME HISTÓRICO REAL
//  Calcula días promedio desde cada estado hasta "En depósito"
//  usando las fechas reales almacenadas en COMPRAS_DJI.
//  Cache 1 hora — solo lectura, no escribe nada.
// ============================================================
function calcularLeadTimeHistorico() {
  return getCachedData('leadtime_hist_v2', function() {
    var dComp = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    var C = SCHEMA.COMPRAS_DJI;

    // estado → array de días que faltaban desde ese estado hasta depósito
    var acum = {
      'Comprado':        [],
      'Pagado':          [],
      'Envío confirmado':[],
      'Forwarder HK':    [],
      'En vuelo':        [],
      'En aduana':       []
    };

    var stateCols = [
      ['Comprado',         C.FECHA_COMPRADO],
      ['Pagado',           C.FECHA_PAGADO],
      ['Envío confirmado', C.FECHA_ENVIO],
      ['Forwarder HK',     C.FECHA_FORWARDER],
      ['En vuelo',         C.FECHA_VUELO],
      ['En aduana',        C.FECHA_ADUANA]
    ];

    for (var i = 1; i < dComp.length; i++) {
      var fDep = dComp[i][C.FECHA_DEPOSITO];
      if (!(fDep instanceof Date)) continue;
      for (var s = 0; s < stateCols.length; s++) {
        var fEst = dComp[i][stateCols[s][1]];
        if (!(fEst instanceof Date)) continue;
        var dias = Math.round((fDep - fEst) / 86400000);
        if (dias > 0 && dias < 400) acum[stateCols[s][0]].push(dias);
      }
    }

    var result = { _n: dComp.length - 1 };
    var estados = Object.keys(acum);
    for (var e = 0; e < estados.length; e++) {
      var arr = acum[estados[e]];
      result[estados[e]] = arr.length > 0
        ? Math.round(arr.reduce(function(a,b){return a+b;},0) / arr.length)
        : null;
    }
    return result;
  }, 3600);
}

// ============================================================
//  SLOTTING AUTOMÁTICO
//  Compara la rotación ABC de cada SKU con la zona de su bin
//  (formato {estante}-{piso}-{lado}) y detecta items mal ubicados.
// ============================================================
function calcularSlottingOptimo() {
  try {
    var dPos = getSheetValues(SCHEMA.SHEETS.TABLA_POSICIONES);
    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    var TP   = SCHEMA.TABLA_POSICIONES;
    var S    = SCHEMA.STOCK_REPUESTOS;

    // Mapa ABC y descripción desde STOCK_REPUESTOS
    var abcMap  = {};
    var descMap = {};
    for (var i = 1; i < dStr.length; i++) {
      var sku = String(dStr[i][S.CODIGO]||'').trim().toUpperCase();
      if (!sku) continue;
      abcMap[sku]  = String(dStr[i][S.CATEGORIA]||'D').trim().toUpperCase() || 'D';
      descMap[sku] = String(dStr[i][S.DESCRIPCION]||'');
    }

    // Construir lista de asignaciones bin → SKU
    var bins = [];
    var maxEstante = 0;
    for (var p = 1; p < dPos.length; p++) {
      var pSku  = String(dPos[p][TP.SKU]||'').trim().toUpperCase();
      var binId = String(dPos[p][TP.BIN_ID]||'').trim();
      if (!pSku || !binId) continue;
      var partes  = binId.split('-');
      var estante = parseInt(partes[0]) || 999;
      if (estante > maxEstante) maxEstante = estante;
      bins.push({
        sku:         pSku,
        descripcion: descMap[pSku] || pSku,
        binId:       binId,
        estante:     estante,
        cantidad:    parseInt(dPos[p][TP.CANTIDAD])||0,
        tipoAlmacen: String(dPos[p][TP.TIPO_ALMACEN]||''),
        categoria:   abcMap[pSku] || 'D'
      });
    }

    if (bins.length === 0) return { ok: true, sugerencias: [], msg: 'Sin datos en TABLA_POSICIONES.' };

    // Umbral: 1/3 del rango de estantes = zona prime
    var thresh = Math.max(2, Math.ceil(maxEstante / 3));

    // Encontrar items A/B mal ubicados y C/D ocupando zonas prime
    var sugerencias = [];
    var skusProcesados = {};

    for (var x = 0; x < bins.length; x++) {
      var item = bins[x];
      if (skusProcesados[item.sku]) continue;
      skusProcesados[item.sku] = true;

      var cat       = item.categoria;
      var isAltaRot = (cat === 'A' || cat === 'B');
      var isFar     = item.estante > thresh * 2;
      var isPrime   = item.estante <= thresh;

      if (isAltaRot && isFar) {
        // A/B en zona lejana — buscar C/D en zona prime para intercambiar
        var swap = null;
        for (var y = 0; y < bins.length; y++) {
          if (!skusProcesados[bins[y].sku] && bins[y].estante <= thresh &&
              (bins[y].categoria === 'C' || bins[y].categoria === 'D')) {
            swap = bins[y];
            break;
          }
        }
        sugerencias.push({
          tipo:        'REUBICAR',
          sku:         item.sku,
          descripcion: item.descripcion,
          binActual:   item.binId,
          categoria:   cat,
          problema:    'Alta rotación (' + cat + ') en estante ' + item.estante + ' — lejos del despacho',
          swapSku:     swap ? swap.sku : null,
          swapDesc:    swap ? swap.descripcion : null,
          swapBin:     swap ? swap.binId : null,
          binSugerido: swap ? swap.binId : 'Zona estante 1–' + thresh
        });
      } else if (!isAltaRot && isPrime) {
        // C/D en zona prime — verificar si algún A/B está lejos
        var hayABLejos = false;
        for (var z = 0; z < bins.length; z++) {
          if ((bins[z].categoria === 'A' || bins[z].categoria === 'B') && bins[z].estante > thresh * 2) {
            hayABLejos = true;
            break;
          }
        }
        if (hayABLejos) {
          sugerencias.push({
            tipo:        'LIBERAR',
            sku:         item.sku,
            descripcion: item.descripcion,
            binActual:   item.binId,
            categoria:   cat,
            problema:    'Baja rotación (' + cat + ') ocupa estante prime ' + item.estante,
            swapSku:     null,
            swapDesc:    null,
            swapBin:     null,
            binSugerido: 'Zona estante ' + (thresh * 2 + 1) + '+'
          });
        }
      }
    }

    return { ok: true, sugerencias: sugerencias, totalBins: bins.length, maxEstante: maxEstante, thresh: thresh };
  } catch(e) {
    Logger.log('calcularSlottingOptimo: ' + e);
    return { ok: false, msg: e.toString(), sugerencias: [] };
  }
}

// ============================================================
//  REABASTECIMIENTO AUTOMÁTICO
//  Trigger diario que crea borradores de compra para items
//  criticos/urgentes sin intervención humana.
// ============================================================

function revisarReabastecimientoAuto() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('auto_restock_habilitado') !== '1') return;
  try {
    var pedido = calcularPedidoDJI();
    var criticos = [];
    for (var p = 0; p < pedido.length; p++) {
      if ((pedido[p].criticidad === 'critico' || pedido[p].criticidad === 'urgente') && pedido[p].sugerido > 0) {
        criticos.push(pedido[p]);
      }
    }
    if (!criticos.length) { Logger.log('revisarReabastecimientoAuto: no hay items criticos/urgentes.'); return; }

    var hoy = new Date();
    var tz  = Session.getScriptTimeZone();
    var prefix = 'AUTO-' + Utilities.formatDate(hoy, tz, 'yyyyMMdd');

    // Verificar si ya existe un borrador automático hoy (para no duplicar)
    var dComp = getSheetValues(SCHEMA.SHEETS.COMPRAS);
    for (var ci = 1; ci < dComp.length; ci++) {
      if (String(dComp[ci][0]||'').toUpperCase().indexOf(prefix) === 0) {
        Logger.log('revisarReabastecimientoAuto: ya existe ' + dComp[ci][0] + ' — omitiendo.');
        return;
      }
    }

    var cas = prefix + '-' + Utilities.formatDate(hoy, tz, 'HHmm');
    var hojaComp = getSheet(SCHEMA.SHEETS.COMPRAS);
    hojaComp.appendRow([
      cas, hoy, 'Borrador', '', '', '', '', '', '', '', '',
      'Sistema · Auto-reabastecimiento',
      'Auto-generado: ' + criticos.length + ' SKU(s) en quiebre/urgente.',
      hoy
    ]);
    SpreadsheetApp.flush();
    invalidateSheetValues(SCHEMA.SHEETS.COMPRAS);

    var listaItems = [];
    for (var li = 0; li < criticos.length; li++) {
      listaItems.push({ sku: criticos[li].codigo, descripcion: criticos[li].descripcion, cantPedida: criticos[li].sugerido });
    }
    vincularItemsACAS(cas, listaItems);
    _notificarAutoRestock(cas, criticos);
    Logger.log('revisarReabastecimientoAuto: creado ' + cas + ' · ' + criticos.length + ' ítems.');
  } catch(e) {
    Logger.log('revisarReabastecimientoAuto error: ' + e);
  }
}

function _notificarAutoRestock(cas, items) {
  try {
    var filas = items.map(function(c) {
      var chipColor = c.criticidad === 'critico' ? '#c0392b' : '#b94600';
      return "<tr style='border-bottom:1px solid #eee'>" +
        "<td style='padding:7px 10px;font-weight:700;color:" + chipColor + ";font-size:12px'>" + c.codigo + "</td>" +
        "<td style='padding:7px 10px;font-size:12px'>" + c.descripcion + "</td>" +
        "<td style='padding:7px 10px;text-align:center;font-size:12px'>" + c.sugerido + "</td>" +
        "<td style='padding:7px 10px;text-align:center'><span style='font-size:10px;font-weight:700;padding:2px 6px;border-radius:4px;" +
          (c.criticidad === 'critico' ? "background:#fde8e8;color:#c0392b" : "background:#ffe5d0;color:#b94600") +
        "'>" + c.criticidad.toUpperCase() + "</span></td></tr>";
    }).join('');
    GmailApp.sendEmail(SM_CONFIG.EMAIL_SUPERVISOR,
      '[Stock Manager] Reabastecimiento automático — ' + items.length + ' SKU(s) críticos',
      '', { htmlBody:
        "<div style='font-family:sans-serif;max-width:640px'>" +
        "<div style='background:#c0392b;padding:16px 20px;border-radius:8px 8px 0 0'>" +
          "<span style='color:#fff;font-size:15px;font-weight:700'>Reabastecimiento automático generado</span></div>" +
        "<div style='background:#fff;border:1px solid #ddd;padding:18px 20px'>" +
          "<p style='font-size:13px;color:#444;margin:0 0 6px'>CAS: <strong>" + cas + "</strong></p>" +
          "<p style='font-size:13px;color:#444;margin:0 0 14px'>El sistema detectó items críticos/urgentes y generó un borrador de compra automáticamente.</p>" +
          "<table style='width:100%;border-collapse:collapse;border:1px solid #eee'>" +
            "<thead><tr style='background:#f5f5f5'><th style='padding:7px 10px;font-size:11px;text-align:left'>SKU</th>" +
            "<th style='padding:7px 10px;font-size:11px;text-align:left'>Descripción</th>" +
            "<th style='padding:7px 10px;font-size:11px'>Sugerido</th>" +
            "<th style='padding:7px 10px;font-size:11px'>Criticidad</th></tr></thead>" +
            "<tbody>" + filas + "</tbody></table>" +
        "</div>" +
        "<div style='background:#f9f9f9;border:1px solid #ddd;border-top:none;padding:12px 20px;border-radius:0 0 8px 8px'>" +
          "<p style='font-size:11px;color:#aaa;margin:0'>Generado automáticamente por Stock Manager · BIDCOMAGRO.<br>Revisá y confirmá el borrador en la sección Compras DJI.</p>" +
        "</div></div>"
      , name: SM_CONFIG.NOMBRE_REMITENTE });
  } catch(e) { Logger.log('_notificarAutoRestock: ' + e); }
}

function configurarAutoReabastecimiento(habilitar) {
  try {
    var props = PropertiesService.getScriptProperties();
    if (habilitar) {
      props.setProperty('auto_restock_habilitado', '1');
      // Crear trigger diario si no existe
      var triggers = ScriptApp.getProjectTriggers();
      for (var t = 0; t < triggers.length; t++) {
        if (triggers[t].getHandlerFunction() === 'revisarReabastecimientoAuto') {
          return { ok: true, msg: 'Ya estaba habilitado. Trigger activo.' };
        }
      }
      ScriptApp.newTrigger('revisarReabastecimientoAuto')
        .timeBased().everyDays(1).atHour(8).create();
      return { ok: true, msg: 'Auto-reabastecimiento habilitado. Corre todos los días a las 8am.' };
    } else {
      props.setProperty('auto_restock_habilitado', '0');
      var triggers2 = ScriptApp.getProjectTriggers();
      for (var t2 = 0; t2 < triggers2.length; t2++) {
        if (triggers2[t2].getHandlerFunction() === 'revisarReabastecimientoAuto') {
          ScriptApp.deleteTrigger(triggers2[t2]);
        }
      }
      return { ok: true, msg: 'Auto-reabastecimiento deshabilitado.' };
    }
  } catch(e) { return { ok: false, msg: e.toString() }; }
}

function getConfigAutoReabastecimiento() {
  try {
    var props = PropertiesService.getScriptProperties();
    var habilitado = props.getProperty('auto_restock_habilitado') === '1';
    var triggerActivo = false;
    var triggers = ScriptApp.getProjectTriggers();
    for (var t = 0; t < triggers.length; t++) {
      if (triggers[t].getHandlerFunction() === 'revisarReabastecimientoAuto') { triggerActivo = true; break; }
    }
    return { ok: true, habilitado: habilitado, triggerActivo: triggerActivo };
  } catch(e) { return { ok: false, habilitado: false, triggerActivo: false }; }
}

// ============================================================
//  MULTI-DEPÓSITO
//  Calcula stock por depósito leyendo MOVIMIENTOS_STOCK col 9.
//  Filas viejas sin depósito cuentan como 'BA' (default).
// ============================================================

function getStockPorDeposito(sku) {
  try {
    var skuKey = String(sku||'').trim().toUpperCase();
    var dMov   = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M      = SCHEMA.MOVIMIENTOS_STOCK;
    var saldos = {};

    for (var i = 1; i < dMov.length; i++) {
      if (String(dMov[i][M.CODIGO]||'').trim().toUpperCase() !== skuKey) continue;
      var dep  = String(dMov[i][M.DEPOSITO]||'BA').trim() || 'BA';
      var cant = parseFloat(dMov[i][M.CANTIDAD]) || 0;
      saldos[dep] = (saldos[dep]||0) + cant;
    }

    // Eliminar depósitos con saldo 0 o negativo
    var resultado = {};
    var deps = Object.keys(saldos);
    for (var d = 0; d < deps.length; d++) {
      if (saldos[deps[d]] > 0) resultado[deps[d]] = Math.round(saldos[deps[d]]);
    }
    return { ok: true, depositos: resultado };
  } catch(e) { return { ok: false, msg: e.toString(), depositos: {} }; }
}

function listarDepositos() {
  try {
    var dMov = getSheetValues(SCHEMA.SHEETS.MOVIMIENTOS);
    var M    = SCHEMA.MOVIMIENTOS_STOCK;
    var deps = {};
    for (var i = 1; i < dMov.length; i++) {
      var dep = String(dMov[i][M.DEPOSITO]||'BA').trim() || 'BA';
      deps[dep] = true;
    }
    var lista = Object.keys(deps).sort();
    if (lista.length === 0) lista = ['BA'];
    return lista;
  } catch(e) { return ['BA']; }
}

function transferirEntreDepositos(sku, cantidad, depositoOrigen, depositoDestino, operador) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var cod = String(sku||'').trim().toUpperCase();
    var cant = parseInt(cantidad)||0;
    var org  = String(depositoOrigen||'BA').trim().toUpperCase();
    var dst  = String(depositoDestino||'BA').trim().toUpperCase();

    if (!cod || cant <= 0) return { ok: false, msg: 'SKU y cantidad requeridos.' };
    if (org === dst) return { ok: false, msg: 'Origen y destino deben ser diferentes.' };

    var saldos = getStockPorDeposito(cod);
    var saldoOrg = (saldos.depositos && saldos.depositos[org]) || 0;
    if (saldoOrg < cant) return { ok: false, msg: 'Stock insuficiente en ' + org + '. Disponible: ' + saldoOrg + ' u.' };

    var dStr = getSheetValues(SCHEMA.SHEETS.STOCK);
    var S    = SCHEMA.STOCK_REPUESTOS;
    var desc = '';
    for (var i = 1; i < dStr.length; i++) {
      if (String(dStr[i][S.CODIGO]||'').trim().toUpperCase() === cod) { desc = String(dStr[i][S.DESCRIPCION]||''); break; }
    }

    var saldoAct = parseInt(dStr[1] && dStr[1][S.STOCK_ACTUAL]||0);
    var ref = 'TRANSF:' + org + '→' + dst;
    _registrarMovimiento('TRANSFERENCIA_SALIDA',  cod, desc,  0, saldoAct, ref, operador||'', org);
    _registrarMovimiento('TRANSFERENCIA_ENTRADA', cod, desc,  0, saldoAct, ref, operador||'', dst);

    return { ok: true, msg: 'Transferido ' + cant + ' u. de ' + org + ' a ' + dst + '.' };
  } catch(e) {
    return { ok: false, msg: e.toString() };
  } finally {
    try { if (lock.hasLock()) lock.releaseLock(); } catch(el) {}
  }
}

// ============================================================
//  LAYOUT DE ALMACÉN — mapa visual de pasillos/estantes/niveles
// ============================================================

function SM_cargarLayout() {
  try {
    var hoja = getSheet(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    if (!hoja) return { ok: true, pasillos: [] };
    var d  = hoja.getDataRange().getValues();
    var LA = SCHEMA.LAYOUT_ALMACEN;
    var map = {};
    for (var i = 1; i < d.length; i++) {
      var pas    = String(d[i][LA.PASILLO]       || '').trim().toUpperCase();
      var ordPas = Number(d[i][LA.ORDEN_PASILLO]) || 0;
      var est    = String(d[i][LA.ESTANTE]        || '').trim();
      var ordEst = Number(d[i][LA.ORDEN_ESTANTE]) || 0;
      var niv    = Math.max(1, Number(d[i][LA.NUM_NIVELES]) || 1);
      if (!pas || !est) continue;
      if (!map[pas]) map[pas] = { pasillo: pas, orden: ordPas, estantes: [] };
      map[pas].estantes.push({ estante: est, orden: ordEst, niveles: niv });
    }
    var pasillos = [];
    var keys = Object.keys(map);
    for (var k = 0; k < keys.length; k++) pasillos.push(map[keys[k]]);
    pasillos.sort(function(a, b) { return a.orden - b.orden || a.pasillo.localeCompare(b.pasillo); });
    for (var p = 0; p < pasillos.length; p++) {
      pasillos[p].estantes.sort(function(a, b) { return a.orden - b.orden || String(a.estante).localeCompare(String(b.estante)); });
    }
    return { ok: true, pasillos: pasillos };
  } catch(e) {
    Logger.log('SM_cargarLayout: ' + e);
    return { ok: false, error: e.message };
  }
}

function SM_guardarLayout(pasillos) {
  try {
    var ss   = getSS();
    var hoja = ss.getSheetByName(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    if (!hoja) hoja = ss.insertSheet(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    hoja.clearContents();
    hoja.appendRow(['PASILLO','ORDEN_PASILLO','ESTANTE','ORDEN_ESTANTE','NUM_NIVELES']);
    for (var i = 0; i < pasillos.length; i++) {
      var p = pasillos[i];
      for (var j = 0; j < p.estantes.length; j++) {
        var e = p.estantes[j];
        hoja.appendRow([
          String(p.pasillo || '').toUpperCase(),
          i + 1,
          String(e.estante || ''),
          j + 1,
          Math.max(1, Number(e.niveles) || 1)
        ]);
      }
    }
    invalidateSheetValues(SCHEMA.SHEETS.LAYOUT_ALMACEN);
    Logger.log('SM_guardarLayout: ' + pasillos.length + ' pasillos guardados');
    return { ok: true };
  } catch(e) {
    Logger.log('SM_guardarLayout: ' + e);
    return { ok: false, error: e.message };
  }
}

