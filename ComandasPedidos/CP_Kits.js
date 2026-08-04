// @version 1.2
// ============================================================
//  COMANDAS — Detección de kits + carga de comandas (CP_getComandas).
//  Extraído de CP_Código.js 1.43 el 2026-07-30 — reorganización sin
//  cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// Clave para matchear KIT ↔ SKU de Ventas (mayúsculas, espacios colapsados)
function _kitKey(s) {
  return String(s == null ? '' : s).trim().toUpperCase().replace(/\s+/g, ' ');
}


// Mapa de KITS { KIT: { comps:{SKU:{sku,desc,cant}}, orden:[] } }.
// CACHÉ: la tabla de kits es de referencia y casi no cambia → se cachea 5 min (CacheService).
// Sólo se cachea si NO está vacío (un fallo transitorio no queda pegado) y entra en el límite
// de 100 KB. Si editás kits y querés verlo ya, corré CP_olvidarKits().
function _cpKitMap() {
  var cache = null;
  try {
    cache = CacheService.getScriptCache();
    var hit = cache.get('kitmap');
    if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  } catch (_) {}
  var map = _cpKitMapRaw();
  if (cache && map && Object.keys(map).length) {
    try {
      var s = JSON.stringify(map);
      if (s.length <= 90000) cache.put('kitmap', s, 300);  // 5 min, sólo si entra en el límite de CacheService
    } catch (_) {}
  }
  return map;
}


// Olvida el mapa de kits cacheado (por si editaste la tabla y querés que se re-lea ya).
function CP_olvidarKits() {
  try { CacheService.getScriptCache().remove('kitmap'); return { ok: true }; }
  catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}


// Lee el sheet de KITS y arma el mapa (sin caché).
// La cantidad de cada componente = nº de filas repetidas de ese componente para el KIT.
function _cpKitMapRaw() {
  try {
    var ss = _cpSS(CP_KITS_SS_ID);
    var h = ss.getSheetByName(CP_KITS_TAB);
    if (!h) { Logger.log('Kit tab no encontrada: ' + CP_KITS_TAB); return {}; }
    var d = h.getDataRange().getValues();
    if (d.length < 2) return {};

    var H = d[0].map(_norm);
    var cSku  = H.indexOf('sku');         if (cSku  === -1) cSku  = 0; // A
    var cDesc = H.indexOf('descripcion'); if (cDesc === -1) cDesc = 1; // B
    var cKit  = H.indexOf('kit');         if (cKit  === -1) cKit  = 2; // C

    var map = {};
    for (var i = 1; i < d.length; i++) {
      var kit = _s(d[i][cKit]), sku = _s(d[i][cSku]), desc = _s(d[i][cDesc]);
      if (!kit || !sku) continue;
      var ku = _kitKey(kit);
      if (!map[ku]) map[ku] = { comps: {}, orden: [] };
      var su = _kitKey(sku);
      if (!map[ku].comps[su]) { map[ku].comps[su] = { sku: sku, desc: desc, cant: 0 }; map[ku].orden.push(su); }
      map[ku].comps[su].cant += 1;
      if (desc && !map[ku].comps[su].desc) map[ku].comps[su].desc = desc;
    }
    return map;
  } catch (e) {
    Logger.log('_cpKitMap error: ' + e);
    return {};
  }
}


// Acumula un componente a cargar en el mapa de la comanda.
function _addCargar(map, orden, sku, desc, cant, noKit) {
  var key = _kitKey(sku) || ('__' + _kitKey(desc));
  if (!map[key]) { map[key] = { sku: sku, desc: desc, cant: 0, noKit: noKit }; orden.push(key); }
  map[key].cant += cant;
  if (desc && !map[key].desc) map[key].desc = desc;
  if (!noKit) map[key].noKit = false; // si aparece como componente de un kit, no es "carga directa"
}

// Aliases (normalizados) para cada campo que necesitamos leer.
var CP_ALIASES = {
  idVenta:        ['idventa'],
  idEntrega:      ['identrega'],
  operacion:      ['operacion'],
  sku:            ['sku'],
  cantidad:       ['cantidad'],
  descripcion:    ['descripcion'],
  reseller:       ['reseller'],
  razonSocial:    ['razonsocial'],
  totalUSD:       ['totalusd'],
  totalARS:       ['totalars'],
  comentarios:    ['comentarios'],
  rtv:            ['rtv'],
  aprobComercial: ['aprobacioncomercial', 'aprobcomercial', 'aprobacioncom']
};


// Ubica la fila de encabezados (busca la que tenga idventa + identrega en las primeras 6 filas)
// y devuelve { headerRow, col } donde col es {campo: índice}.
function _cpDetectar(rows) {
  var maxScan = Math.min(6, rows.length);
  for (var hr = 0; hr < maxScan; hr++) {
    var normRow = rows[hr].map(_norm);
    if (normRow.indexOf('idventa') > -1 && normRow.indexOf('identrega') > -1) {
      var col = {};
      Object.keys(CP_ALIASES).forEach(function(campo) {
        var idx = -1;
        var alias = CP_ALIASES[campo];
        for (var a = 0; a < alias.length && idx === -1; a++) {
          idx = normRow.indexOf(alias[a]);
        }
        col[campo] = idx; // -1 si no existe esa columna
      });
      return { headerRow: hr, col: col, headers: rows[hr] };
    }
  }
  return null;
}


/* ── API PRINCIPAL ───────────────────────────────────────── */
// Devuelve las ventas con ID_Entrega = CARGAR, agrupadas por ID_Venta.
function CP_getComandas() {
  try {
    var rows = _cpHoja().getDataRange().getValues();
    if (!rows || rows.length < 2) {
      return { ok: true, ts: _cpTs(), ventas: 0, items: 0, comandas: [] };
    }

    var det = _cpDetectar(rows);
    if (!det) {
      return { ok: false, error: 'No encontré los encabezados "ID_Venta" e "ID_Entrega" en las primeras filas de la hoja "' + CP_TAB + '". Revisá que los títulos estén escritos así.' };
    }
    var col = det.col;
    if (col.idEntrega === -1 || col.idVenta === -1) {
      return { ok: false, error: 'Falta la columna ID_Entrega o ID_Venta en la hoja.' };
    }

    // helper para leer una celda por campo (tolera columnas ausentes)
    function get(r, campo) { var i = col[campo]; return i === -1 ? '' : r[i]; }

    var flag = CP_FLAG.toUpperCase();
    var mapa = {};   // idVenta -> comanda
    var orden = [];  // preserva orden de aparición
    var totalItems = 0;
    var logMap  = _cpLogMap();       // { 'IDVENTA||SKU': { ts:Date, origen } }
    var kitMap  = _cpKitMap();       // { KIT: { comps, orden } } → explotar en componentes
    var enviosMap = _cpEnviosMap();  // { IDVENTA: [envios...] } → envíos por venta

    for (var i = det.headerRow + 1; i < rows.length; i++) {
      var r = rows[i];
      var idVenta = _s(get(r, 'idVenta')) || '(sin ID_Venta)';
      var key = idVenta.toUpperCase();
      var esCargar = _s(get(r, 'idEntrega')).toUpperCase() === flag;
      // Incluir la venta si está marcada CARGAR, O si tiene envíos aunque YA haya salido del flag
      // CARGAR (el ERP le saca "CARGAR" al procesarse en Masterchief). Sin esto, una venta con
      // envíos desaparecía de TODAS las solapas al crear el envío.
      if (!esCargar && !(enviosMap[key] && enviosMap[key].length)) continue;

      if (!mapa[key]) {
        mapa[key] = {
          idVenta:        idVenta,
          operacion:      _s(get(r, 'operacion')),
          reseller:       _s(get(r, 'reseller')),
          razonSocial:    _s(get(r, 'razonSocial')),
          rtv:            _s(get(r, 'rtv')),
          aprobComercial: _s(get(r, 'aprobComercial')),
          comentarios:    _s(get(r, 'comentarios')),
          totalUSD:       0,
          totalARS:       0,
          _lin:           {},   // consolidación por SKU
          _linOrden:      [],
          _sinCargar:     !esCargar   // arranca "recuperada por envíos"; baja a false si hay una línea CARGAR
        };
        orden.push(key);
      }
      var c = mapa[key];
      if (esCargar) c._sinCargar = false;

      // Completar cabecera si vino vacía en la primera fila
      if (!c.operacion)      c.operacion      = _s(get(r, 'operacion'));
      if (!c.reseller)       c.reseller       = _s(get(r, 'reseller'));
      if (!c.razonSocial)    c.razonSocial    = _s(get(r, 'razonSocial'));
      if (!c.rtv)            c.rtv            = _s(get(r, 'rtv'));
      if (!c.aprobComercial) c.aprobComercial = _s(get(r, 'aprobComercial'));

      var comLinea = _s(get(r, 'comentarios'));
      if (comLinea && c.comentarios.indexOf(comLinea) === -1) {
        c.comentarios = c.comentarios ? (c.comentarios + ' · ' + comLinea) : comLinea;
      }

      var lUSD = _num(get(r, 'totalUSD'));
      var lARS = _num(get(r, 'totalARS'));
      c.totalUSD += lUSD;
      c.totalARS += lARS;

      // Consolidar por SKU: un mismo SKU se puede partir en varias líneas con
      // cantidades fraccionadas (0.65 + 0.35) por facturación; para despacho
      // interesa la SUMA real. Agrupamos por SKU (o descripción si no hay SKU).
      var skuRaw = _s(get(r, 'sku'));
      var descRaw = _s(get(r, 'descripcion'));
      var lkey = (skuRaw || descRaw || ('__row' + i)).toUpperCase();
      if (!c._lin[lkey]) {
        c._lin[lkey] = { sku: skuRaw, descripcion: descRaw, cant: 0, usd: 0, ars: 0, partes: 0 };
        c._linOrden.push(lkey);
      }
      var lg = c._lin[lkey];
      if (!lg.sku && skuRaw)          lg.sku = skuRaw;
      if (!lg.descripcion && descRaw) lg.descripcion = descRaw;
      lg.cant  += _num(get(r, 'cantidad'));
      lg.usd   += lUSD;
      lg.ars   += lARS;
      lg.partes++;
    }

    var comandas = orden.map(function(k) {
      var c = mapa[k];
      var minTs = null, maxTs = null, minExacto = true;
      var cargarMap = {}, cargarOrden = [];  // explosión agregada a nivel comanda
      c.lineas = c._linOrden.map(function(lk) {
        var lg = c._lin[lk];
        var cant = Math.round(lg.cant * 100) / 100; // limpia ruido de coma flotante
        // momento en que se marcó CARGAR (por ID_Venta + SKU)
        var reg = logMap[_cpKey(c.idVenta, lg.sku)];
        var ts = reg ? reg.ts : null;
        if (ts) {
          if (!minTs || ts < minTs) { minTs = ts; minExacto = (reg.origen === 'edit'); }
          if (!maxTs || ts > maxTs) maxTs = ts;
        }
        // explotar el KIT en sus componentes reales a cargar en Masterchief
        var kit = kitMap[_kitKey(lg.sku)];
        var esKit = !!(kit && kit.orden.length);
        if (esKit) {
          kit.orden.forEach(function(cu) {
            var comp = kit.comps[cu];
            _addCargar(cargarMap, cargarOrden, comp.sku, comp.desc, comp.cant * cant, false);
          });
        } else {
          // no está en la tabla de kits → se carga el propio ítem
          _addCargar(cargarMap, cargarOrden, lg.sku, lg.descripcion, cant, true);
        }
        return {
          sku:         lg.sku,
          cantidad:    _fmtCant(cant),
          descripcion: lg.descripcion,
          partes:      lg.partes,   // >1 => venía dividido en varias líneas de factura
          esKit:       esKit,
          totalUSDStr: _fmtUSD(lg.usd),
          totalARSStr: _fmtARS(lg.ars),
          cargadoTs:   ts ? ts.getTime() : null,
          cargadoStr:  ts ? _fmtTs(ts) : ''
        };
      });
      // lista final "qué cargar en Masterchief"
      c.cargar = cargarOrden.map(function(cu) {
        var it = cargarMap[cu];
        var q = Math.round(it.cant * 100) / 100;
        return { sku: it.sku, descripcion: it.desc, cantidad: _fmtCant(q), cantNum: q, noKit: !!it.noKit };
      });
      totalItems += c.lineas.length;
      c.fueraDeCargar = !!c._sinCargar;   // true = ya no está en CARGAR, se muestra por tener envíos
      delete c._lin; delete c._linOrden; delete c._sinCargar;
      c.totalUSDStr = _fmtUSD(c.totalUSD);
      c.totalARSStr = _fmtARS(c.totalARS);
      // resumen de "marcado CARGAR" a nivel comanda
      c.cargadoTs     = minTs ? minTs.getTime() : null;
      c.cargadoStr    = minTs ? _fmtTs(minTs) : '';
      c.cargadoHasta  = maxTs ? _fmtTs(maxTs) : '';
      c.cargadoRango  = !!(minTs && maxTs && minTs.getTime() !== maxTs.getTime());
      c.cargadoExacto = minExacto;  // true = capturado por onEdit (exacto); false = detección aprox.
      return c;
    });

    // Enriquecer cada venta con sus ENVÍOS + lo que falta enviar
    var masterMap   = _cpMasterMap();
    var cfg         = _cpConfig();
    var ocaBase     = cfg['OCA_TRACKING_URL'] || '';
    var resellerMap = _cpResellerMap();
    var rtvMap      = _cpRtvMailMap();
    var entregaMap  = _cpEntregaMap();   // sync inverso: ENTREGADO marcado en PENDIENTES_ENTREGA
    var hayFijos    = !!_s(cfg['MAIL_DESTINATARIOS']);
    comandas.forEach(function(c) {
      var arr = enviosMap[c.idVenta.toUpperCase()] || [];
      // ¿hay a quién mandarle el mail? (reseller / RTV / fijos)
      var rinfo = resellerMap[_kitKey(c.reseller)] || {};
      var rtvNombre = rinfo.rtv || c.rtv || '';
      var mailT = !!(rtvNombre && rtvMap[_kitKey(rtvNombre)]);
      c.destino = { reseller: !!rinfo.mail, rtv: mailT, fijos: hayFijos, rtvNombre: rtvNombre };
      c.tieneDestino = !!(rinfo.mail || mailT || hayFijos);
      // descripción por SKU
      var descBy = {};
      c.cargar.forEach(function(it) { descBy[String(it.sku).toUpperCase()] = it.descripcion; });
      // total enviado + pendiente
      var enviado = {};
      arr.forEach(function(e) { Object.keys(e.productos || {}).forEach(function(sk) { enviado[sk] = (enviado[sk] || 0) + _num(e.productos[sk]); }); });
      c.pending = [];
      c.cargar.forEach(function(it) {
        var falta = Math.round(((it.cantNum || 0) - _num(enviado[String(it.sku).toUpperCase()])) * 100) / 100;
        if (falta > 0) {
          var eg = entregaMap[_cpKey(c.idVenta, it.sku)];
          c.pending.push({
            sku: it.sku, descripcion: it.descripcion, cantidad: _fmtCant(falta), cantNum: falta,
            entregado: !!(eg && eg.entregado), entregadoFecha: (eg && eg.fechaStr) || ''
          });
        }
      });
      c.entregadosFuera = c.pending.filter(function(p) { return p.entregado; }).length;
      // detalle de cada envío (guía/estado de Comandas Master + estado del mail)
      c.envios = arr.map(function(e) {
        var parts = String(e.comanda || '').split('/').map(function(s) { return s.trim(); }).filter(Boolean);
        var links = [], transs = [], estados = [], pdfs = [], tieneGuia = parts.length > 0;
        parts.forEach(function(p) {
          var m = masterMap[p.toUpperCase()];
          if (!m || !m.guia) tieneGuia = false;
          if (m) {
            if (m.estado && estados.indexOf(m.estado) === -1) estados.push(m.estado);
            if (m.guia) links.push({ guia: m.guia, url: ocaBase ? ocaBase.replace('{GUIA}', encodeURIComponent(m.guia)) : '' });
            if (m.transportista && transs.indexOf(m.transportista) === -1) transs.push(m.transportista);
          }
          var pdf = _cpBuscarPdf(p);
          if (pdf) pdfs.push({ comanda: p, name: pdf.name, url: pdf.url });
        });
        var items = Object.keys(e.productos || {}).map(function(sk) {
          return { sku: sk, descripcion: descBy[sk] || '', cantidad: _fmtCant(_num(e.productos[sk])) };
        });
        var mailErr = (e.estado && e.estado.indexOf('MAIL ERROR') === 0) ? e.estado.replace(/^MAIL ERROR · /, '') : '';
        // tieneGuia = ya autorizado en Masterchief (tiene N° de seguimiento) — NO implica que
        // ya salió. despachado = col F de Comandas Master dice DESPACHADO, recién ahí sale el
        // mail al reseller (ver _cpEnvioListoDespacho en CP_Datos.js).
        var despachado = _cpEnvioListoDespacho(parts, masterMap).listo;
        return {
          envio: e.envio, comanda: e.comanda, fechaStr: e.fechaStr, fechaTs: e.fechaTs, operador: e.operador,
          items: items, guiasLinks: links, estadoDespacho: estados.join(', '),
          transportista: (transs.length ? transs.join(', ') : e.transportista), tieneGuia: tieneGuia, despachado: despachado,
          mailReseller: e.mailReseller, mailAutorizado: e.mailAutorizado, mailAprob: e.mailAprob, mailError: mailErr,
          notaAprob: e.notaAprob, notaReseller: e.notaReseller, pdfs: pdfs
        };
      });
      // clasificación: pendiente (sin envíos) / parcial (algo enviado, falta) / completo (todo enviado)
      c.estadoEnvio = (!arr.length) ? 'pendiente' : (c.pending.length ? 'parcial' : 'completo');
    });

    var pend     = comandas.filter(function(c) { return c.estadoEnvio === 'pendiente'; });
    var parcial  = comandas.filter(function(c) { return c.estadoEnvio === 'parcial'; });
    var completo = comandas.filter(function(c) { return c.estadoEnvio === 'completo'; });
    var itemsPend = pend.reduce(function(a, c) { return a + c.lineas.length; }, 0);
    return {
      ok: true, ts: _cpTs(),
      comandas: comandas,
      ventas: pend.length,               // pendientes (para el hero)
      items: itemsPend,
      parcialCount: parcial.length,
      completoCount: completo.length,
      sla: { warn: _num(cfg['SLA_WARN_HORAS']) || 4, danger: _num(cfg['SLA_DANGER_HORAS']) || 24 }
    };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}
