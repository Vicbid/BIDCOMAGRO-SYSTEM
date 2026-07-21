// ── STOCK MANAGER — Dashboard ─────────────────────────────────────
// @version 1.1

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
        metodoPago: String(fct[3] || ''),
        eta:        _fmtFecha(fct[SCHEMA.COMPRAS_DJI.ETA])
      });
    }

    // Cruzar pronóstico de quiebre con unidades en camino (COMPRAS_DETALLE)
    try {
      var casEstMap_d = {};
      var casEtaMap_d = {};
      for (var cip = 1; cip < dCom.length; cip++) {
        var casIdP  = String(dCom[cip][0] || '').trim().toUpperCase();
        var casEstP = String(dCom[cip][2] || '').trim();
        if (casIdP && casEstP !== 'En depósito' && casEstP.indexOf('Borrador') === -1) {
          casEstMap_d[casIdP] = casEstP;
          casEtaMap_d[casIdP] = _fmtFecha(dCom[cip][SCHEMA.COMPRAS_DJI.ETA]);
        }
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
            if (!enCaminoProno[skuP]) { enCaminoProno[skuP] = { total: 0, cas: casP, estado: casEstMap_d[casP], eta: casEtaMap_d[casP] || '' }; }
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
