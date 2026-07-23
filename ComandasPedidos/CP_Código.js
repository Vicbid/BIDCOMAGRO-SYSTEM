// @version 1.43
// ============================================================
//  COMANDAS · CARGA MASTERCHIEF — Backend · modelo N-envíos por venta
// ============================================================

/* ── ENTRY POINT ─────────────────────────────────────────── */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('CP_Index')
    .setTitle('Comandas · Carga Masterchief')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ── HELPERS ─────────────────────────────────────────────── */
// Memoiza los spreadsheets abiertos DENTRO de una misma ejecución. Los globals se reinician
// por request en Apps Script, así que no hay datos viejos: solo evita reabrir el mismo archivo
// 3-5 veces por refresh (getRange/getValues siguen leyendo en vivo). Usa notación de corchetes
// a propósito para no ser reemplazado por el replace de openById.
var _ssCache = {};
function _cpSS(id) {
  if (!_ssCache[id]) _ssCache[id] = SpreadsheetApp['openById'](id);
  return _ssCache[id];
}

function _cpHoja() {
  var ss = _cpSS(CP_SS_ID);
  var h  = ss.getSheetByName(CP_TAB);
  if (!h) throw new Error('No se encontró la hoja "' + CP_TAB + '".');
  return h;
}

function _s(v)  { return String(v == null ? '' : v).trim(); }

// Anti-inyección de fórmulas: si un texto libre del usuario empieza con = + - @ (o tab),
// le antepone ' para que Sheets lo guarde como TEXTO y no como fórmula (evita =IMPORTDATA,
// =HYPERLINK, etc. corriendo en la hoja del dueño cuando la abre).
function _cpSafeCell(s) {
  var v = _s(s);
  return /^[=+\-@\t\r]/.test(v) ? ("'" + v) : v;
}

/* ── SEGURIDAD: usuario, autorización, auditoría, rate-limit ── */
// Email del usuario logueado (con acceso por dominio, ahora sí viene poblado).
function _cpUsuarioActual() {
  try { return _s(Session.getActiveUser().getEmail()); } catch (_) { return ''; }
}

// ¿El usuario puede ejecutar acciones que escriben o mandan mail?
// OPERADORES_AUTORIZADOS vacío = todos (opt-in). Con lista, sólo esos mails.
function _cpUsuarioAutorizado() {
  var lista = _s(_cpConfig()['OPERADORES_AUTORIZADOS']);
  var email = _cpUsuarioActual();
  if (!lista) return { ok: true, email: email };
  var permit = lista.split(',').map(function(x) { return x.trim().toLowerCase(); }).filter(Boolean);
  if (email && permit.indexOf(email.toLowerCase()) > -1) return { ok: true, email: email };
  return { ok: false, noAutorizado: true,
    mensaje: 'No estás autorizado para esta acción' + (email ? ' (' + email + ')' : '') + '. Pedí que agreguen tu mail a OPERADORES_AUTORIZADOS en _CONFIG.' };
}

// Hoja de auditoría (en el sheet de log). Se crea si no existe.
function _cpAuditHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_AUDIT_TAB);
  if (!h) {
    h = ss.insertSheet(CP_AUDIT_TAB);
    h.getRange(1, 1, 1, 6).setValues([['Fecha', 'Usuario', 'Acción', 'ID_Venta', 'Envío', 'Detalle']]);
    h.setFrozenRows(1);
    h.getRange('A1:F1').setFontWeight('bold');
    [150, 220, 150, 130, 60, 380].forEach(function(w, i) { h.setColumnWidth(i + 1, w); });
  }
  return h;
}
// Registra una acción sensible: quién + cuándo + qué + sobre qué. Nunca interrumpe la operación.
function _cpAuditar(accion, idVenta, envio, detalle) {
  try {
    _cpAuditHoja().appendRow([new Date(), _cpUsuarioActual(), _cpSafeCell(accion),
      _cpSafeCell(_s(idVenta)), (envio || envio === 0 ? envio : ''), _cpSafeCell(_s(detalle))]);
  } catch (e) { Logger.log('_cpAuditar: ' + e); }
}

// Rate-limit de mails: contador por ventana de 10 min (CacheService). Protege contra loops
// o abuso y evita agotar la cuota de Gmail. Si el cache falla, NO bloquea (fail-open).
var _mailCapMemo = null;
function _cpMailCap() {
  if (_mailCapMemo == null) { var c = _num(_cpConfig()['MAIL_MAX_POR_10MIN']); _mailCapMemo = c > 0 ? c : 60; }
  return _mailCapMemo;
}
function _cpRateBucketKey() { return 'mailrate_' + Math.floor(Date.now() / (10 * 60000)); }
function _cpRateLimitOk() {
  try { return _num(CacheService.getScriptCache().get(_cpRateBucketKey())) < _cpMailCap(); }
  catch (_) { return true; }
}
function _cpRateLimitInc() {
  try {
    var cache = CacheService.getScriptCache(), k = _cpRateBucketKey();
    cache.put(k, String(_num(cache.get(k)) + 1), 700);  // TTL > ventana de 10 min
  } catch (_) {}
}
// Único punto de envío: aplica rate-limit y manda. Lanza si se pasó el tope (los callers
// ya cachean el error como MAIL ERROR y reintentan cuando se libera la ventana).
// Convierte a ASCII puro toda la parte no-ASCII de un HTML usando referencias numéricas
// (&#NNN;): emojis y acentos viajan como ASCII y el cliente los renderiza igual, sin depender
// de cómo el transporte declare/maneje el charset. FIX de los emojis que llegaban como "������"
// (el transporte de GmailApp mangla los caracteres de plano astral —emojis de 4 bytes—).
// Seguro para nuestro HTML de mail: usa estilos inline (atributos), sin <style>/<script> donde
// las entidades no se decodifican. No toca entidades ya escapadas (&amp; &nbsp; &#233;) porque
// son ASCII. Incluye pares suplentes (flag u) para los emojis de plano astral (📦 → &#128230;).
function _cpHtmlAscii(s) {
  return String(s == null ? '' : s).replace(/[^\x00-\x7F]/gu, function(ch) {
    return '&#' + ch.codePointAt(0) + ';';
  });
}

function _cpGmailSend(to, subject, plain, opts) {
  if (!_cpRateLimitOk()) throw new Error('Límite de envío de mails alcanzado (' + _cpMailCap() + ' por 10 min, anti-abuso). Reintentá en unos minutos.');
  opts = opts || {};
  // MODO PRUEBA: si EMAIL_PRUEBA está seteado en _CONFIG, TODO el correo se redirige a esa
  // dirección (no le llega a resellers/RTV/Sole). Se cortan cc/bcc y el asunto avisa a quién
  // habría ido. Vaciar EMAIL_PRUEBA en _CONFIG para volver a producción.
  var prueba = _s(_cpConfig()['EMAIL_PRUEBA']);
  if (prueba) {
    var destinoReal = [to, opts.cc ? 'cc ' + opts.cc : '', opts.bcc ? 'bcc ' + opts.bcc : '']
      .filter(function(x){ return x; }).join(' · ');
    subject = '[PRUEBA → ' + destinoReal + '] ' + subject;
    to = prueba;
    delete opts.cc;
    delete opts.bcc;
  }
  // Emojis/acentos del cuerpo → referencias numéricas ASCII, así el cliente los renderiza bien
  // (el transporte manglaba los emojis y llegaban como "������"). Se hace acá, en el chokepoint
  // único, para cubrir TODOS los mails (reseller, RTV, Sole/aprobación, recordatorios).
  if (opts.htmlBody) opts.htmlBody = _cpHtmlAscii(opts.htmlBody);
  GmailApp.sendEmail(to, subject, plain, opts);
  _cpRateLimitInc();
}

// Normaliza un encabezado: minúsculas, sin acentos, solo letras/números.
// "ID_Venta" -> "idventa" | "Total USD" -> "totalusd" | "Razon Social" -> "razonsocial"
function _norm(v) {
  return String(v == null ? '' : v)
    .toLowerCase()
    .normalize('NFD')            // separa acentos: "ó" -> "o" + marca combinante
    .replace(/[^a-z0-9]/g, '');  // quita marcas, espacios, guiones, puntos, etc.
}

// Parseo robusto de números (acepta 1.234,56 / 1,234.56 / "USD 1.234")
function _num(v) {
  if (typeof v === 'number') return v;
  var s = _s(v);
  if (!s) return 0;
  s = s.replace(/[^\d,.\-]/g, '');
  if (s.indexOf(',') > -1 && s.indexOf('.') > -1) {
    // el último separador es el decimal
    if (s.lastIndexOf(',') > s.lastIndexOf('.')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.indexOf(',') > -1) {
    s = s.replace(',', '.');
  }
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function _fmtUSD(n) {
  return 'USD ' + n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
function _fmtARS(n) {
  return '$ ' + n.toLocaleString('es-AR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
// Cantidad: entero sin decimales; fracción con hasta 2 decimales (coma es-AR)
function _fmtCant(n) {
  return n.toLocaleString('es-AR', { maximumFractionDigits: 2 });
}

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
        return {
          envio: e.envio, comanda: e.comanda, fechaStr: e.fechaStr, fechaTs: e.fechaTs, operador: e.operador,
          items: items, guiasLinks: links, estadoDespacho: estados.join(', '),
          transportista: (transs.length ? transs.join(', ') : e.transportista), tieneGuia: tieneGuia,
          mailReseller: e.mailReseller, mailAprob: e.mailAprob, mailError: mailErr,
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

function _cpTs() {
  return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', "dd/MM/yyyy HH:mm");
}
function _fmtTs(d) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), 'America/Argentina/Buenos_Aires', "dd/MM/yyyy HH:mm");
}
// Parsea "dd/MM/yyyy HH:mm" (nuestro formato) → Date, o null. Tolera texto alrededor ("SÍ · ...").
function _cpParseFechaAr(s) {
  var m = String(s == null ? '' : s).match(/(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/);
  return m ? new Date(+m[3], +m[2] - 1, +m[1], +m[4], +m[5]) : null;
}

/* ════════════════════════════════════════════════════════════
   REGISTRO DEL MOMENTO EN QUE SE MARCÓ "CARGAR"
   - onEdit (instalable) captura el momento EXACTO cuando alguien
     escribe/pega "CARGAR" en ID_Entrega → origen 'edit'.
   - CP_snapshotCargar() (opcional) estampa "detección" aprox. para
     filas que ya estaban en CARGAR antes de instalar el trigger.
   Se guarda en la hoja CP_LOG_TAB, sin tocar la hoja Ventas.
════════════════════════════════════════════════════════════ */

// clave única por línea: ID_Venta + SKU (normalizados)
function _cpKey(idv, sku) {
  return String(idv == null ? '' : idv).trim().toUpperCase() + '||' +
         String(sku == null ? '' : sku).trim().toUpperCase();
}

// Devuelve/crea la hoja de log (en el sheet SEPARADO, nunca en Ventas).
function _cpLogHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_LOG_TAB);
  if (!h) {
    h = ss.insertSheet(CP_LOG_TAB);
    h.getRange(1, 1, 1, 5).setValues([['ID_Venta', 'SKU', 'Marcado CARGAR', 'Origen', 'Usuario']]);
    h.setFrozenRows(1);
  }
  return h;
}

// Lee el log a un mapa { clave: {ts:Date, origen:string} } (se queda con el más antiguo).
function _cpLogMap() {
  try {
    var ss = _cpSS(CP_LOG_SS_ID);
    var h = ss.getSheetByName(CP_LOG_TAB);
    if (!h) return {};
    var d = h.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var idv = d[i][0], sku = d[i][1], fecha = d[i][2], origen = d[i][3];
      if (idv === '' && sku === '') continue;
      var ts = (fecha instanceof Date) ? fecha : (fecha ? new Date(fecha) : null);
      if (!ts || isNaN(ts.getTime())) continue;
      var key = _cpKey(idv, sku);
      if (!m[key] || ts < m[key].ts) m[key] = { ts: ts, origen: String(origen || 'edit') };
    }
    return m;
  } catch (e) {
    Logger.log('_cpLogMap error: ' + e);
    return {};
  }
}

// Agrega filas [idv, sku, Date, origen, email] al log, evitando duplicados por clave.
function _cpLogStamp(filas) {
  if (!filas || !filas.length) return;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) { /* seguimos igual */ }
  try {
    var h = _cpLogHoja();
    var d = h.getDataRange().getValues();
    var seen = {};
    for (var i = 1; i < d.length; i++) seen[_cpKey(d[i][0], d[i][1])] = true;

    var nuevas = [];
    filas.forEach(function(f) {
      var key = _cpKey(f[0], f[1]);
      if (!seen[key]) { seen[key] = true; nuevas.push(f); }
    });
    if (nuevas.length) {
      h.getRange(h.getLastRow() + 1, 1, nuevas.length, 5).setValues(nuevas);
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// TRIGGER onEdit (instalable). Se dispara con cada edición manual/pegado.
function CP_onEditVentas(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CP_TAB) return;

    var lastCol = sh.getLastColumn();
    var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(_norm);
    var cEnt = headers.indexOf('identrega');
    var cVen = headers.indexOf('idventa');
    var cSku = headers.indexOf('sku');
    if (cEnt === -1 || cVen === -1) return;

    var r0 = e.range.getRow(), c0 = e.range.getColumn();
    var nR = e.range.getNumRows(), nC = e.range.getNumColumns();

    // ¿el rango editado incluye la columna ID_Entrega?
    var colEnt1 = cEnt + 1; // 1-based
    if (colEnt1 < c0 || colEnt1 > c0 + nC - 1) return;

    var block = sh.getRange(r0, 1, nR, lastCol).getValues();
    var flag = CP_FLAG.toUpperCase();
    var now = new Date();
    var email = '';
    try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}

    var filas = [];
    for (var i = 0; i < nR; i++) {
      if (r0 + i === 1) continue; // saltar encabezado
      var row = block[i];
      if (String(row[cEnt] || '').trim().toUpperCase() !== flag) continue;
      var idv = String(row[cVen] || '').trim();
      var sku = cSku > -1 ? String(row[cSku] || '').trim() : '';
      filas.push([idv, sku, now, 'edit', email]);
    }
    _cpLogStamp(filas);
  } catch (err) {
    // nunca interrumpir la edición del usuario
    Logger.log('CP_onEditVentas error: ' + err);
  }
}

// SETUP: correr UNA vez desde el editor. Instala el trigger onEdit y crea la hoja de log.
function CP_setupTrigger() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) {
    return t.getHandlerFunction() === 'CP_onEditVentas';
  });
  if (ya) { Logger.log('El trigger onEdit ya estaba instalado. Nada que hacer.'); }
  else {
    ScriptApp.newTrigger('CP_onEditVentas').forSpreadsheet(CP_SS_ID).onEdit().create();
    Logger.log('✅ Trigger onEdit instalado sobre el spreadsheet.');
  }
  _cpLogHoja();
  _cpEnviosHoja();
  _cpConfigHoja();
  _cpRtvHoja();
  _cpPendHoja();
  _cpAuditHoja();
  try { CP_poblarRtvDesdeResellers(); } catch (e) { Logger.log('poblar RTV: ' + e); }
  Logger.log('✅ Hojas listas (CARGAR_LOG, ENVIOS, _CONFIG, RTV, PENDIENTES_ENTREGA, AUDITORIA) en el sheet de log.');
}

// SELF-CHECK de read-only: verifica que el sheet de ESCRITURA (log) sea distinto del de
// Ventas, y lista las hojas donde el programa escribe. Garantía extra de que Ventas no se toca.
// Correr desde el editor y mirar los logs / el objeto devuelto.
function CP_selfCheckReadOnly() {
  var ok = (CP_LOG_SS_ID !== CP_SS_ID);
  var r = {
    ok: ok,
    ventas_solo_lectura: CP_SS_ID,
    escribe_en: CP_LOG_SS_ID,
    hojas_de_escritura: [CP_LOG_TAB, CP_ENVIOS_TAB, CP_CONFIG_TAB, CP_RTV_TAB, CP_PEND_TAB, CP_AUDIT_TAB],
    mensaje: ok ? 'OK: el archivo de Ventas y el de escritura son distintos; nada se escribe sobre Ventas.'
                : '⚠️ PELIGRO: CP_LOG_SS_ID == CP_SS_ID — la app escribiría sobre la hoja de Ventas.'
  };
  Logger.log(JSON.stringify(r, null, 2));
  return r;
}

// OPCIONAL: estampa una marca de "detección" (aprox = ahora) para todas las filas
// que HOY están en CARGAR y todavía no tienen registro. Útil como línea de base
// para los pedidos que ya estaban marcados antes de instalar el trigger.
function CP_snapshotCargar() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) { Logger.log('No se detectaron encabezados.'); return; }
  var col = det.col, flag = CP_FLAG.toUpperCase();
  var existentes = _cpLogMap();
  var now = new Date();
  var filas = [];
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[col.idEntrega] || '').trim().toUpperCase() !== flag) continue;
    var idv = String(r[col.idVenta] || '').trim();
    var sku = col.sku > -1 ? String(r[col.sku] || '').trim() : '';
    if (existentes[_cpKey(idv, sku)]) continue; // ya registrado
    filas.push([idv, sku, now, 'deteccion', '']);
  }
  _cpLogStamp(filas);
  Logger.log('Snapshot: ' + filas.length + ' fila(s) nuevas marcadas como detección (aprox.).');
}

/* ════════════════════════════════════════════════════════════
   ENVÍOS — una venta puede tener VARIOS envíos. Cada envío tiene su
   propia comanda, su aprobación (mail a Sole) y su guía (mail al
   reseller). Se guardan en CP_ENVIOS_TAB (sheet de log, nunca en Ventas).
   Columnas: ID_Venta | Envío | Comanda | Fecha | Operador |
     Productos(JSON {SKU:cant}) | Guía | Transportista | Estado |
     Mail Aprobador | Mail Reseller | Nota Aprobador | Nota Reseller
════════════════════════════════════════════════════════════ */
function _cpEnviosHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_ENVIOS_TAB);
  if (!h) {
    h = ss.insertSheet(CP_ENVIOS_TAB);
    h.getRange(1, 1, 1, 13).setValues([[
      'ID_Venta', 'Envío', 'Comanda', 'Fecha', 'Operador', 'Productos',
      'Guía', 'Transportista', 'Estado', 'Mail Aprobador', 'Mail Reseller', 'Nota Aprobador', 'Nota Reseller'
    ]]);
    h.setFrozenRows(1);
    h.getRange('A1:M1').setFontWeight('bold');
  }
  return h;
}

// Parsea un JSON {SKU:cant} → { SKU: cant>0 }
function _cpParseJson(raw) {
  try {
    var s = _s(raw); if (!s) return {};
    var o = JSON.parse(s), m = {};
    Object.keys(o).forEach(function(k) { var v = _num(o[k]); if (v > 0) m[String(k).toUpperCase()] = v; });
    return m;
  } catch (e) { return {}; }
}
// Normaliza objeto {SKU:cant} → JSON (sólo > 0) o '' si no hay.
function _cpProductosJson(obj) {
  var limpio = {};
  if (obj && typeof obj === 'object') {
    Object.keys(obj).forEach(function(k) { var v = _num(obj[k]); if (v > 0) limpio[String(k).toUpperCase()] = v; });
  }
  return Object.keys(limpio).length ? JSON.stringify(limpio) : '';
}

// Mapa { IDVENTA: [ {envio, comanda, fecha, fechaStr, fechaTs, operador, productos, guia,
//                    transportista, estado, mailAprob, mailReseller, notaAprob, notaReseller, rowIdx} ] }
function _cpEnviosMap() {
  try {
    var ss = _cpSS(CP_LOG_SS_ID);
    var h = ss.getSheetByName(CP_ENVIOS_TAB);
    if (!h) return {};
    var d = h.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var idv = _s(d[i][0]); if (!idv) continue;
      var key = idv.toUpperCase();
      var fecha = (d[i][3] instanceof Date) ? d[i][3] : (d[i][3] ? new Date(d[i][3]) : null);
      var e = {
        envio:         _num(d[i][1]) || 0,
        comanda:       _s(d[i][2]),
        fecha:         fecha,
        fechaStr:      fecha ? _fmtTs(fecha) : '',
        fechaTs:       (fecha && !isNaN(fecha.getTime())) ? fecha.getTime() : null,
        operador:      _s(d[i][4]),
        productos:     _cpParseJson(d[i][5]),
        guia:          _s(d[i][6]),
        transportista: _s(d[i][7]),
        estado:        _s(d[i][8]),
        mailAprob:     _s(d[i][9]),
        mailReseller:  _s(d[i][10]),
        mailResellerTs:(function() { var dt = _cpParseFechaAr(d[i][10]); return dt ? dt.getTime() : null; })(),
        notaAprob:     _s(d[i][11]),
        notaReseller:  _s(d[i][12]),
        rowIdx:        i + 1
      };
      if (!m[key]) m[key] = [];
      m[key].push(e);
    }
    Object.keys(m).forEach(function(k) { m[k].sort(function(a, b) { return a.envio - b.envio; }); });
    return m;
  } catch (e) { Logger.log('_cpEnviosMap error: ' + e); return {}; }
}

// Total enviado por SKU sumando todos los envíos de una venta.
function _cpEnviadoTotal(enviosArr) {
  var t = {};
  (enviosArr || []).forEach(function(e) {
    Object.keys(e.productos || {}).forEach(function(sk) { t[sk] = (t[sk] || 0) + _num(e.productos[sk]); });
  });
  return t;
}

// Lo que falta enviar de una venta: [{sku, descripcion, cantidad, cantNum}] (pedido - enviado > 0).
function _cpPendingVenta(det, enviosArr) {
  var enviado = _cpEnviadoTotal(enviosArr);
  var pend = [];
  (det.cargar || []).forEach(function(it) {
    var falta = Math.round(((it.cantNum || 0) - _num(enviado[String(it.sku).toUpperCase()])) * 100) / 100;
    if (falta > 0) pend.push({ sku: it.sku, descripcion: it.desc || it.descripcion, cantidad: _fmtCant(falta), cantNum: falta });
  });
  return pend;
}

// Detalle de los productos de un envío: [{sku, desc, cant}] (para los mails).
function _cpDetalleEnvio(det, prodMap) {
  var descBy = {};
  (det.cargar || []).forEach(function(it) { descBy[String(it.sku).toUpperCase()] = it.desc || it.descripcion; });
  return Object.keys(prodMap || {}).map(function(sk) { return { sku: sk, desc: descBy[sk] || '', cant: _fmtCant(_num(prodMap[sk])) }; });
}

// ── PENDIENTES DE ENTREGA (lo que falta enviar: 1 producto por línea) ──
function _cpPendHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_PEND_TAB);
  if (!h) {
    h = ss.insertSheet(CP_PEND_TAB);
    h.getRange(1, 1, 1, 8).setValues([['ID_Venta', 'SKU', 'Descripción', 'Cantidad', 'Reseller', 'Razón Social', 'Fecha', 'Estado']]);
    h.setFrozenRows(1);
    h.getRange('A1:H1').setFontWeight('bold');
    [110, 130, 300, 80, 170, 210, 130, 120].forEach(function(w, i) { h.setColumnWidth(i + 1, w); });
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(['PENDIENTE', 'ENTREGADO'], true).setAllowInvalid(true).build();
    h.getRange('H2:H2000').setDataValidation(rule);
  }
  return h;
}

// Reescribe las líneas pendiente-de-entrega de una venta (1 fila por SKU que falta enviar),
// PRESERVANDO el Estado y la Fecha que ya tuviera cada SKU (para no pisar ediciones manuales).
function _cpSyncPendientesEntrega(idVenta, det, pendArr) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var h = _cpPendHoja();
    var d = h.getDataRange().getValues();
    var key = _s(idVenta).toUpperCase();
    var prev = {};   // SKU -> {estado, fecha}
    for (var i = d.length - 1; i >= 1; i--) {
      if (_s(d[i][0]).toUpperCase() === key) {
        prev[_s(d[i][1]).toUpperCase()] = { estado: _s(d[i][7]), fecha: d[i][6] };
        h.deleteRow(i + 1);
      }
    }
    if (!pendArr || !pendArr.length) return;
    var now = new Date();
    var filas = pendArr.map(function(p) {
      var pv = prev[String(p.sku).toUpperCase()] || {};
      return [_cpSafeCell(idVenta), _cpSafeCell(p.sku), _cpSafeCell(p.descripcion || ''), p.cantNum,
              _cpSafeCell(det.reseller || ''), _cpSafeCell(det.razonSocial || ''),
              (pv.fecha instanceof Date ? pv.fecha : now), (_s(pv.estado) || 'PENDIENTE')];
    });
    h.getRange(h.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

// Lee PENDIENTES_ENTREGA → { 'IDVENTA||SKU': {estado, entregado, fechaStr} }.
// Permite reflejar en la app lo que se marca ENTREGADO manualmente en la hoja (sync inverso).
function _cpEntregaMap() {
  try {
    var ss = _cpSS(CP_LOG_SS_ID);
    var h = ss.getSheetByName(CP_PEND_TAB);
    if (!h) return {};
    var d = h.getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var idv = _s(d[i][0]), sku = _s(d[i][1]); if (!idv) continue;
      var estado = _s(d[i][7]).toUpperCase();
      var fecha = (d[i][6] instanceof Date) ? d[i][6] : null;
      m[_cpKey(idv, sku)] = { estado: estado, entregado: (estado === 'ENTREGADO'), fechaStr: fecha ? _fmtTs(fecha) : '' };
    }
    return m;
  } catch (e) { Logger.log('_cpEntregaMap error: ' + e); return {}; }
}

// El operador crea un ENVÍO: N° de comanda + cuántas unidades de cada SKU manda en este envío.
// productos = { SKU: cantidad }. Dispara el mail de aprobación a Sole (con PDF adjunto).
function CP_crearEnvio(idVenta, comanda, productos, notaAprob, notaReseller, force) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    idVenta = _s(idVenta); comanda = _s(comanda);
    notaAprob = _s(notaAprob); notaReseller = _s(notaReseller);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    if (!comanda) return { ok: false, mensaje: 'Ingresá el número de comanda.' };
    if (!/^\d{5,}(\/\d{5,})*$/.test(comanda)) return { ok: false, mensaje: 'N° de comanda inválido (solo números, o dos separados por /).' };

    var det = _cpDetalleVenta(idVenta);
    var enviosArr = _cpEnviosMap()[idVenta.toUpperCase()] || [];

    // comanda ya usada en otro envío de esta u otra venta → aviso salvo force
    if (!force) {
      var usada = false, dondeV = '';
      var mapAll = _cpEnviosMap();
      Object.keys(mapAll).forEach(function(k) {
        mapAll[k].forEach(function(e) {
          e.comanda.split('/').forEach(function(cp) {
            comanda.split('/').forEach(function(cn) { if (cp.trim() && cp.trim() === cn.trim()) { usada = true; dondeV = k; } });
          });
        });
      });
      if (usada) return { ok: false, yaUsada: true, donde: dondeV, mensaje: 'La comanda ' + comanda + ' ya está usada en un envío (' + dondeV + ').' };
    }

    // lo pendiente antes de este envío
    var pendBy = {};
    _cpPendingVenta(det, enviosArr).forEach(function(p) { pendBy[String(p.sku).toUpperCase()] = p.cantNum; });
    // productos a enviar ahora, topeados a lo pendiente
    var envProd = {};
    Object.keys(productos || {}).forEach(function(k) {
      var ku = String(k).toUpperCase();
      var q = Math.min(_num(productos[k]), pendBy[ku] || 0);
      if (q > 0) envProd[ku] = Math.round(q * 100) / 100;
    });
    if (!Object.keys(envProd).length) return { ok: false, mensaje: 'No hay productos para enviar (revisá las cantidades).' };

    var nextEnvio = 1;
    enviosArr.forEach(function(e) { if (e.envio >= nextEnvio) nextEnvio = e.envio + 1; });
    var email = ''; try { email = Session.getActiveUser().getEmail() || ''; } catch (_) {}
    var now = new Date();

    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch (e) {}
    try {
      _cpEnviosHoja().appendRow([_cpSafeCell(idVenta), nextEnvio, comanda, now, email, _cpProductosJson(envProd),
        '', '', '', '', '', _cpSafeCell(notaAprob), _cpSafeCell(notaReseller)]);
    } finally { try { lock.releaseLock(); } catch (e) {} }

    // recomputar pendiente y espejar (1 producto por línea)
    var pend2 = _cpPendingVenta(det, enviosArr.concat([{ productos: envProd }]));
    try { _cpSyncPendientesEntrega(idVenta, det, pend2); } catch (pe) { Logger.log('syncPend: ' + pe); }

    // mail a Sole (aprobación de ESTE envío) con PDF adjunto
    var res = { ok: true, envio: nextEnvio, completo: pend2.length === 0 };
    var soleRes = _cpEnviarMailSoleCore(idVenta, nextEnvio);
    if (soleRes && soleRes.ok) { res.mailSole = true; if (soleRes.sinPdf) res.mailSoleSinPdf = true; }
    else { res.mailSole = false; if (soleRes && soleRes.mensaje) res.mailSoleError = soleRes.mensaje; }
    _cpAuditar('Crear envío', idVenta, nextEnvio, 'comanda ' + comanda + ' · ' + Object.keys(envProd).length + ' ítem(s)' + (res.mailSole ? ' · mail Sole' : ''));
    return res;
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}

// Envía (o reenvía) a Sole el mail de aprobación de UN envío (con PDF adjunto). Marca col "Mail Aprobador".
function _cpEnviarMailSoleCore(idVenta, envio) {
  idVenta = _s(idVenta); envio = _num(envio);
  var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
  var e = null; arr.forEach(function(x) { if (x.envio === envio) e = x; });
  if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };
  var cfg = _cpConfig();
  var sole = cfg['MAIL_APROBACION'];
  if (!sole) return { ok: false, sinDestino: true, mensaje: 'Falta MAIL_APROBACION en _CONFIG.' };
  var det = _cpDetalleVenta(idVenta);
  var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
  var pdfs = [];
  parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
  var detEnv = _cpDetalleEnvio(det, e.productos);
  var pend = _cpPendingVenta(det, arr);   // faltantes de la venta tras este envío
  var asunto = (cfg['ASUNTO_APROBACION'] || 'APROBAR MC · {COMANDA} · {IDVENTA}')
    .replace('{COMANDA}', parts.join('/')).replace('{IDVENTA}', idVenta);
  var opts = {
    htmlBody: _cpMailAprobacionHtml(idVenta, parts, det, pdfs, e.notaAprob, detEnv, e.envio, pend),
    name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO'
  };
  if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];
  var adj = _cpPdfBlobs(parts);
  if (adj.length) opts.attachments = adj;
  try {
    _cpGmailSend(sole, asunto, 'Autorizar comanda ' + e.comanda + ' en Masterchief.', opts);
  } catch (se) {
    return { ok: false, mailError: true, mensaje: 'No se pudo enviar el mail a Sole: ' + String(se && se.message ? se.message : se) };
  }
  _cpMarcarMailAprob(e.rowIdx);
  return { ok: true, sinPdf: !pdfs.length, destinatario: sole };
}

function _cpMarcarMailAprob(rowIdx) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try { _cpEnviosHoja().getRange(rowIdx, 10).setValue('SÍ · ' + _fmtTs(new Date())); }
  finally { try { lock.releaseLock(); } catch (e) {} }
}

// Reenvía a Sole el mail de aprobación de un envío (por si no lo vio).
function CP_reenviarAprobacion(idVenta, envio) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    var r = _cpEnviarMailSoleCore(_s(idVenta), _num(envio));
    if (r && r.ok) { _cpAuditar('Reenviar aprobación', idVenta, envio, 'a ' + (r.destinatario || '')); return { ok: true, destinatario: r.destinatario, sinPdf: !!r.sinPdf }; }
    return { ok: false, mensaje: (r && r.mensaje) || 'No se pudo reenviar.' };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}

// Chequea si están en Drive los PDF de una comanda (una o varias separadas por /).
// Devuelve { ok, hay:[...], falta:[...], todos:bool }.
function CP_checkPdf(comanda) {
  try {
    var parts = _s(comanda).split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    var hay = [], falta = [];
    parts.forEach(function(p) { if (_cpBuscarPdf(p)) hay.push(p); else falta.push(p); });
    return { ok: true, hay: hay, falta: falta, todos: (parts.length > 0 && falta.length === 0) };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}

// Borra un envío (solo si NO se mandó el mail al reseller). Recalcula pendientes.
function CP_borrarEnvio(idVenta, envio) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch (e) {}
    try {
      var h = _cpEnviosHoja();
      var d = h.getDataRange().getValues();
      var key = idVenta.toUpperCase();
      for (var i = d.length - 1; i >= 1; i--) {
        if (_s(d[i][0]).toUpperCase() === key && (_num(d[i][1]) === envio)) {
          if (_s(d[i][10])) return { ok: false, mensaje: 'Ya se envió el mail al reseller de este envío; no se puede borrar.' };
          h.deleteRow(i + 1);
        }
      }
    } finally { try { lock.releaseLock(); } catch (e) {} }
    var det = _cpDetalleVenta(idVenta);
    var enviosArr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    try { _cpSyncPendientesEntrega(idVenta, det, _cpPendingVenta(det, enviosArr)); } catch (pe) { Logger.log('syncPend: ' + pe); }
    _cpAuditar('Borrar envío', idVenta, envio, '');
    return { ok: true };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}

// Edita la nota para el reseller de un envío puntual.
function CP_setNotaResellerEnvio(idVenta, envio, texto) {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    idVenta = _s(idVenta); envio = _num(envio); texto = _s(texto);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var lock = LockService.getScriptLock();
    try { lock.waitLock(8000); } catch (e) {}
    try {
      var h = _cpEnviosHoja();
      var d = h.getDataRange().getValues();
      var key = idVenta.toUpperCase();
      for (var i = 1; i < d.length; i++) {
        if (_s(d[i][0]).toUpperCase() === key && (_num(d[i][1]) === envio)) { h.getRange(i + 1, 13).setValue(_cpSafeCell(texto)); break; }
      }
    } finally { try { lock.releaseLock(); } catch (e) {} }
    return { ok: true };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}


/* ════════════════════════════════════════════════════════════
   DESPACHO — lee "Comandas Master" (col A idComprobante → estado F,
   guía K, transportista B). Habilita el envío de mail manual.
════════════════════════════════════════════════════════════ */

// Mapa { COMANDA(idComprobante): {estado, guia, transportista, fechaEntrega, idVenta} }
function _cpMasterMap() {
  try {
    var ss = _cpSS(CP_SS_ID);
    var h = ss.getSheetByName(CP_MASTER_TAB);
    if (!h) { Logger.log('Master tab no encontrada: ' + CP_MASTER_TAB); return {}; }
    var d = h.getDataRange().getValues();
    if (d.length < 2) return {};
    var H = d[0].map(_norm);
    var cId    = H.indexOf('idcomprobante');            if (cId    === -1) cId    = 0;  // A
    var cTrans = H.indexOf('nombredespachomedio');      if (cTrans === -1) cTrans = 1;  // B
    var cEst   = H.indexOf('nombredespachoestado');     if (cEst   === -1) cEst   = 5;  // F
    var cFecha = H.indexOf('fechaentregadespacho');     if (cFecha === -1) cFecha = 6;  // G
    var cGuia  = H.indexOf('codigoseguimientodespacho');if (cGuia  === -1) cGuia  = 10; // K
    var cIdV   = H.indexOf('idventa');                  if (cIdV   === -1) cIdV   = 11; // L
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var id = _s(d[i][cId]); if (!id) continue;
      m[id.toUpperCase()] = {
        estado:        _s(d[i][cEst]).toUpperCase(),
        guia:          _s(d[i][cGuia]),
        transportista: _s(d[i][cTrans]),
        fechaEntrega:  (d[i][cFecha] instanceof Date) ? _fmtTs(d[i][cFecha]) : _s(d[i][cFecha]),
        idVenta:       _s(d[i][cIdV])
      };
    }
    return m;
  } catch (e) { Logger.log('_cpMasterMap error: ' + e); return {}; }
}

// Busca el PDF de una comanda en la carpeta de Drive (el nombre empieza con el N°).
// Devuelve {name, url, id} o null.
// CACHÉ (positive-only): sólo se cachean los ACIERTOS — una comanda que ya tiene su PDF no
// cambia, así que evitamos re-buscar en Drive en cada refresh. Los "no encontrado" NUNCA se
// cachean, por lo que un PDF recién subido aparece en el siguiente refresh (no queda pegado).
function _cpBuscarPdf(comanda) {
  comanda = _s(comanda);
  if (!comanda || !CP_PDF_FOLDER_ID) return null;
  var cache = null, ckey = 'pdf_' + comanda.replace(/[^0-9A-Za-z]/g, '');
  try {
    cache = CacheService.getScriptCache();
    var hit = cache.get(ckey);
    if (hit) { try { return JSON.parse(hit); } catch (_) {} }
  } catch (_) {}
  var res = _cpBuscarPdfRaw(comanda);
  if (res && cache) { try { cache.put(ckey, JSON.stringify(res), 3600); } catch (_) {} }  // sólo aciertos, TTL 1 h
  return res;
}

// Búsqueda real en Drive (sin caché).
// Convención: el nombre del PDF EMPIEZA con el N° de comanda (ej. "15861196 - Flo Agro.pdf").
function _cpBuscarPdfRaw(comanda) {
  try {
    var q = _s(comanda).replace(/[^0-9A-Za-z]/g, '');  // sólo alfanumérico → no se puede inyectar en la query de Drive
    if (!q) return null;
    var folder = DriveApp.getFolderById(CP_PDF_FOLDER_ID);
    var norm = function(n) { return _s(n).replace(/[^0-9A-Za-z]/g, ''); };
    var prefijo = function(list) { return list.filter(function(f) { return norm(f.getName()).indexOf(q) === 0; }); };

    // 1) Índice de Drive (rápido). Preferimos los que EMPIEZAN con el número de comanda.
    var idx = [];
    try {
      var it = folder.searchFiles('title contains "' + q + '"');
      while (it.hasNext()) idx.push(it.next());
    } catch (eS) { Logger.log('_cpBuscarPdf searchFiles: ' + eS); }
    var hit = prefijo(idx)[0];

    // 2) Fallback: el índice de Drive TARDA en ver un archivo recién subido → searchFiles no lo
    //    encuentra aunque ya esté en la carpeta. Recorremos la carpeta en vivo (getFiles, listado
    //    directo) y matcheamos por prefijo de nombre. Así el PDF recién subido aparece al instante.
    if (!hit) {
      var it2 = folder.getFiles(), live = [];
      while (it2.hasNext()) live.push(it2.next());
      hit = prefijo(live)[0];
    }
    if (!hit) return null;
    return { name: hit.getName(), url: hit.getUrl(), id: hit.getId() };
  } catch (e) {
    Logger.log('_cpBuscarPdf error: ' + e);
    return null;
  }
}

// Olvida el PDF cacheado de una comanda (por si reemplazaste el archivo y querés que se
// re-resuelva ya, sin esperar el TTL de 1 h). Opcional, se corre desde el editor.
function CP_olvidarPdf(comanda) {
  try {
    var c = _s(comanda);
    if (!c) return { ok: false, mensaje: 'Falta la comanda.' };
    CacheService.getScriptCache().remove('pdf_' + c.replace(/[^0-9A-Za-z]/g, ''));
    return { ok: true };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}

// Devuelve los blobs PDF (para adjuntar) de una lista de comandas.
function _cpPdfBlobs(comandas) {
  var blobs = [];
  (comandas || []).forEach(function(p) {
    try {
      var pdf = _cpBuscarPdf(p);
      if (pdf && pdf.id) blobs.push(DriveApp.getFileById(pdf.id).getBlob());
    } catch (e) { Logger.log('_cpPdfBlobs ' + p + ': ' + e); }
  });
  return blobs;
}

// "Documentos definidos" = TODOS los archivos de la carpeta CP_DOCS_FOLDER_ID. Se adjuntan al
// reseller SOLO en el primer envío de la venta. Carpeta sin configurar o vacía → [] (no adjunta
// nada ni falla el mail).
// _cpDocsDefinidosArchivos: lista liviana [{id, name}] (para nombrarlos en el cuerpo del mail).
function _cpDocsDefinidosArchivos() {
  var out = [];
  try {
    var fid = _s(typeof CP_DOCS_FOLDER_ID !== 'undefined' ? CP_DOCS_FOLDER_ID : '');
    if (!fid) return out;   // carpeta sin configurar
    var it = DriveApp.getFolderById(fid).getFiles();
    while (it.hasNext()) { var f = it.next(); out.push({ id: f.getId(), name: f.getName() }); }
  } catch (e) { Logger.log('_cpDocsDefinidosArchivos: ' + e); }
  return out;
}
// _cpDocsDefinidosBlobs: los blobs (para adjuntar).
function _cpDocsDefinidosBlobs() {
  var blobs = [];
  _cpDocsDefinidosArchivos().forEach(function(a) {
    try { blobs.push(DriveApp.getFileById(a.id).getBlob()); } catch (e) { Logger.log('_cpDocsDefinidosBlobs ' + a.name + ': ' + e); }
  });
  return blobs;
}

// ── _CONFIG (parámetros del mail) ──
var CP_CONFIG_DEFAULTS = [
  ['MAIL_APROBACION',       ''],  // ← mail de Sole (aprobación de comanda en Masterchief)
  ['ASUNTO_APROBACION',     'APROBAR MC · {COMANDA} · {IDVENTA}'],
  ['MAIL_DESTINATARIOS',    ''],  // despacho: destinatarios fijos (separá varios con coma)
  ['MAIL_CC',               ''],
  ['MAIL_BCC',              ''],  // copia oculta (reseller + aprobación); separá varios con coma
  ['MAIL_ASUNTO',           'Despacho {IDVENTA} · Comanda {COMANDA} — {CLIENTE}'],
  ['MAIL_REMITENTE_NOMBRE', 'BIDCOMAGRO'],
  ['EMAIL_PRUEBA',          ''],  // MODO PRUEBA: si ponés un mail acá, TODOS los correos (reseller, RTV, Sole, recordatorios) se mandan SOLO a esa dirección (no a los reales). Vaciar para volver a producción.
  ['OCA_TRACKING_URL',      'https://www1.oca.com.ar/OEPTrackingWeb/trackingdetalle.aspx?numero={GUIA}'],
  ['AUTO_MAIL_DESPACHO',    'NO'],  // SI = manda el mail final al reseller+RTV automáticamente al detectar despacho
  ['RECORDATORIO_HORAS',    '24'],  // cada cuántas horas recordarle a Sole una comanda sin despachar
  ['SLA_WARN_HORAS',        '4'],   // semáforo SLA: a partir de cuántas horas pasa a amarillo
  ['SLA_DANGER_HORAS',      '24'],  // semáforo SLA: a partir de cuántas horas pasa a rojo
  ['OPERADORES_AUTORIZADOS','']  ,  // mails (coma) que pueden crear/borrar envíos y mandar mail; VACÍO = todos
  ['MAIL_MAX_POR_10MIN',    '60']   // tope anti-abuso de mails enviados por ventana de 10 min
];

function _cpConfigHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_CONFIG_TAB);
  if (!h) {
    h = ss.insertSheet(CP_CONFIG_TAB);
    h.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]);
    h.getRange(2, 1, CP_CONFIG_DEFAULTS.length, 2).setValues(CP_CONFIG_DEFAULTS);
    h.setFrozenRows(1);
    h.setColumnWidth(1, 210); h.setColumnWidth(2, 460);
  } else {
    // agregar claves nuevas que falten (para hojas ya creadas)
    var d = h.getDataRange().getValues();
    var existentes = {};
    for (var i = 1; i < d.length; i++) { var k = _s(d[i][0]); if (k) existentes[k.toUpperCase()] = true; }
    var faltan = CP_CONFIG_DEFAULTS.filter(function(kv) { return !existentes[kv[0].toUpperCase()]; });
    if (faltan.length) h.getRange(h.getLastRow() + 1, 1, faltan.length, 2).setValues(faltan);
  }
  return h;
}
function _cpConfig() {
  var d = _cpConfigHoja().getDataRange().getValues();
  var m = {};
  for (var i = 1; i < d.length; i++) { var k = _s(d[i][0]); if (k) m[k.toUpperCase()] = _s(d[i][1]); }
  return m;
}

// Detalle de una venta (para el cuerpo del mail): cliente, totales, qué cargar, pedido.
function _cpDetalleVenta(idVenta) {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) return {};
  var col = det.col;
  function get(r, f) { var i = col[f]; return i === -1 ? '' : r[i]; }
  var kitMap = _cpKitMap();
  var key = _s(idVenta).toUpperCase();
  var razonSocial = '', reseller = '', operacion = '', rtv = '', totalUSD = 0, totalARS = 0;
  var lin = {}, linOrd = [];
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var r = rows[i];
    if (_s(get(r, 'idVenta')).toUpperCase() !== key) continue;
    if (!razonSocial) razonSocial = _s(get(r, 'razonSocial'));
    if (!reseller)    reseller    = _s(get(r, 'reseller'));
    if (!operacion)   operacion   = _s(get(r, 'operacion'));
    if (!rtv)         rtv         = _s(get(r, 'rtv'));
    totalUSD += _num(get(r, 'totalUSD'));
    totalARS += _num(get(r, 'totalARS'));
    var sku = _s(get(r, 'sku')), desc = _s(get(r, 'descripcion'));
    var lk = (sku || desc || ('__' + i)).toUpperCase();
    if (!lin[lk]) { lin[lk] = { sku: sku, desc: desc, cant: 0 }; linOrd.push(lk); }
    if (!lin[lk].sku && sku)   lin[lk].sku = sku;
    if (!lin[lk].desc && desc) lin[lk].desc = desc;
    lin[lk].cant += _num(get(r, 'cantidad'));
  }
  var cargarMap = {}, cargarOrden = [];
  linOrd.forEach(function(lk) {
    var lg = lin[lk], cant = Math.round(lg.cant * 100) / 100;
    var kit = kitMap[_kitKey(lg.sku)];
    if (kit && kit.orden.length) {
      kit.orden.forEach(function(cu) { var comp = kit.comps[cu]; _addCargar(cargarMap, cargarOrden, comp.sku, comp.desc, comp.cant * cant, false); });
    } else { _addCargar(cargarMap, cargarOrden, lg.sku, lg.desc, cant, true); }
  });
  return {
    razonSocial: razonSocial, reseller: reseller, operacion: operacion, rtv: rtv,
    totalUSDStr: _fmtUSD(totalUSD), totalARSStr: _fmtARS(totalARS),
    cargar: cargarOrden.map(function(cu) { var it = cargarMap[cu]; var q = Math.round(it.cant*100)/100; return { sku: it.sku, desc: it.desc, cant: _fmtCant(q), cantNum: q }; }),
    pedido: linOrd.map(function(lk) { var lg = lin[lk]; return { sku: lg.sku, desc: lg.desc, cant: _fmtCant(Math.round(lg.cant*100)/100) }; })
  };
}

// Mapa LIVIANO { IDVENTA: {reseller, rtv} } de UNA sola lectura de Ventas.
// Sirve para chequear destinatarios de muchos envíos sin releer toda la hoja Ventas
// por cada uno (lo que colgaba el diagnóstico cuando había muchos envíos con guía
// pendientes de mail). Usa la misma detección de columnas que _cpDetalleVenta.
function _cpVentaResellerRtvMap() {
  var out = {};
  try {
    var rows = _cpHoja().getDataRange().getValues();
    var det = _cpDetectar(rows);
    if (!det) return out;
    var col = det.col;
    function get(r, f) { var i = col[f]; return i === -1 ? '' : r[i]; }
    for (var i = det.headerRow + 1; i < rows.length; i++) {
      var r = rows[i];
      var key = _s(get(r, 'idVenta')).toUpperCase();
      if (!key) continue;
      if (!out[key]) out[key] = { reseller: '', rtv: '' };
      if (!out[key].reseller) out[key].reseller = _s(get(r, 'reseller'));
      if (!out[key].rtv)      out[key].rtv      = _s(get(r, 'rtv'));
    }
  } catch (e) { Logger.log('_cpVentaResellerRtvMap error: ' + e); }
  return out;
}

// Mapa nombre_reseller(normalizado) → {mail, rtv}, leído de la pestaña "Resellers".
//   B = Reseller (nombre) | C = RTV | J = Email  (se detectan por encabezado)
function _cpResellerMap() {
  try {
    var ss = _cpSS(CP_SS_ID);
    var h = ss.getSheetByName(CP_RESELLERS_TAB);
    if (!h) { Logger.log('Tab "' + CP_RESELLERS_TAB + '" no encontrada.'); return {}; }
    var d = h.getDataRange().getValues();
    if (d.length < 2) return {};
    var H = d[0].map(_norm);
    var cName = _cpFindCol(H, ['reseller', 'nombrereseller', 'empresa', 'nombre']);
    var cMail = _cpFindCol(H, ['email', 'mail', 'correo', 'correoelectronico']);
    var cRtv  = _cpFindCol(H, ['rtv']);
    if (cName === -1) cName = 1;  // col B por defecto
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var nom = _s(d[i][cName]); if (!nom) continue;
      var mail = cMail > -1 ? _s(d[i][cMail]) : '';
      var rtv  = cRtv  > -1 ? _s(d[i][cRtv])  : '';
      m[_kitKey(nom)] = { mail: (mail.indexOf('@') > -1 ? mail : ''), rtv: rtv };
    }
    return m;
  } catch (e) { Logger.log('_cpResellerMap error: ' + e); return {}; }
}

// Hoja RTV (Nombre | Mail) en el sheet de log; se crea si no existe.
function _cpRtvHoja() {
  var ss = _cpSS(CP_LOG_SS_ID);
  var h = ss.getSheetByName(CP_RTV_TAB);
  if (!h) {
    h = ss.insertSheet(CP_RTV_TAB);
    h.getRange(1, 1, 1, 2).setValues([['Nombre RTV', 'Mail']]);
    h.setFrozenRows(1);
    h.setColumnWidth(1, 240); h.setColumnWidth(2, 300);
  }
  return h;
}
function _cpRtvMailMap() {
  try {
    var d = _cpRtvHoja().getDataRange().getValues();
    var m = {};
    for (var i = 1; i < d.length; i++) {
      var nom = _s(d[i][0]), mail = _s(d[i][1]);
      if (!nom || !mail || mail.indexOf('@') === -1) continue;
      m[_kitKey(nom)] = mail;
    }
    return m;
  } catch (e) { Logger.log('_cpRtvMailMap error: ' + e); return {}; }
}

// Pre-carga la hoja RTV con los nombres distintos de RTV que hay en "Resellers"
// (col C), dejando el mail en blanco para completar a mano. Correr una vez.
function CP_poblarRtvDesdeResellers() {
  var rm = _cpResellerMap();
  var nombres = {};
  Object.keys(rm).forEach(function(k) { var n = _s(rm[k].rtv); if (n) nombres[n.toUpperCase()] = n; });
  var h = _cpRtvHoja();
  var d = h.getDataRange().getValues();
  var exist = {};
  for (var i = 1; i < d.length; i++) { var n = _s(d[i][0]); if (n) exist[n.toUpperCase()] = true; }
  var nuevas = [];
  Object.keys(nombres).forEach(function(u) { if (!exist[u]) nuevas.push([nombres[u], '']); });
  if (nuevas.length) h.getRange(h.getLastRow() + 1, 1, nuevas.length, 2).setValues(nuevas);
  // resumen claro (es idempotente: NO re-agrega los que ya están)
  var dd = h.getDataRange().getValues();
  var total = 0, conMail = 0;
  for (var j = 1; j < dd.length; j++) {
    var nn = _s(dd[j][0]); if (!nn) continue;
    total++; if (_s(dd[j][1]).indexOf('@') > -1) conMail++;
  }
  Logger.log('RTV distintos encontrados en Resellers (col C): ' + Object.keys(nombres).length +
    ' · agregados nuevos: ' + nuevas.length +
    ' · total en la hoja "' + CP_RTV_TAB + '": ' + total +
    ' · con mail cargado: ' + conMail +
    (total ? (conMail < total ? ' → faltan ' + (total - conMail) + ' mails por completar (col B)' : ' ✅ todos con mail') : ' → ⚠️ la hoja está vacía y no se encontraron RTV en Resellers'));
}

// Busca el primer encabezado (normalizado) que coincida con alguno de los alias.
function _cpFindCol(H, alias) {
  for (var a = 0; a < alias.length; a++) { var i = H.indexOf(alias[a]); if (i > -1) return i; }
  // por si el header contiene el alias (ej. "Mail Reseller")
  for (var j = 0; j < H.length; j++) {
    for (var b = 0; b < alias.length; b++) { if (H[j].indexOf(alias[b]) > -1) return j; }
  }
  return -1;
}

// Destinatarios del mail al reseller de un envío: reseller (Resellers col J) + RTV (hoja RTV) +
// fijos (_CONFIG MAIL_DESTINATARIOS). Devuelve { to:[...], detalle:[...] } deduplicado.
// resellerMap / rtvMap: opcionales. Si se pasan (ya leídos una vez), evita releer
// las hojas Resellers/RTV por cada envío (clave para no colgar el diagnóstico en lote).
function _cpDestinatariosEnvio(det, cfg, resellerMap, rtvMap) {
  var toList = [], detalle = [];
  var rinfo = (resellerMap || _cpResellerMap())[_kitKey(det.reseller)] || {};
  if (rinfo.mail) { toList.push(rinfo.mail); detalle.push('reseller'); }
  var rtvName = rinfo.rtv || det.rtv;
  var mailRtv = rtvName ? (rtvMap || _cpRtvMailMap())[_kitKey(rtvName)] : '';
  if (mailRtv) { toList.push(mailRtv); detalle.push('RTV (' + rtvName + ')'); }
  if (cfg['MAIL_DESTINATARIOS']) { toList.push(cfg['MAIL_DESTINATARIOS']); detalle.push('fijos'); }
  var seen = {}, to = [];
  toList.join(',').split(',').forEach(function(x) { x = x.trim(); var k = x.toLowerCase(); if (x && !seen[k]) { seen[k] = true; to.push(x); } });
  return { to: to, detalle: detalle };
}

// VISTA PREVIA del mail al reseller de un envío (NO envía, NO marca nada).
// Devuelve { ok, html, asunto, destinatarios, a, sinGuia, sinDestino }.
function CP_previewMailEnvio(idVenta, envio) {
  try {
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    var e = null; arr.forEach(function(x) { if (x.envio === envio) e = x; });
    if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };

    var master = _cpMasterMap();
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    var guias = [], transs = [], tieneGuia = parts.length > 0;
    parts.forEach(function(p) {
      var m = master[p.toUpperCase()];
      if (!m || !m.guia) tieneGuia = false;
      if (m) { if (m.guia) guias.push(m.guia); if (m.transportista && transs.indexOf(m.transportista) === -1) transs.push(m.transportista); }
    });

    var cfg = _cpConfig();
    var det = _cpDetalleVenta(idVenta);
    var dest = _cpDestinatariosEnvio(det, cfg);
    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);
    var pend = _cpPendingVenta(det, arr);
    var ocaBase = cfg['OCA_TRACKING_URL'] || '';
    // Mismos adjuntos que el envío real: comanda(s) siempre; documentos definidos solo en el 1er envío.
    var esPrimerEnvio = !arr.some(function(x) { return x.envio !== e.envio && x.mailReseller; });
    var docNames = esPrimerEnvio ? _cpDocsDefinidosArchivos().map(function(a) { return a.name; }) : [];
    var asunto = (cfg['MAIL_ASUNTO'] || 'Despacho {IDVENTA} · Comanda {COMANDA} — {CLIENTE}')
      .replace('{IDVENTA}', idVenta).replace('{COMANDA}', parts.join('/')).replace('{CLIENTE}', det.razonSocial || det.reseller || '');
    var html = _cpMailHtml(idVenta, parts, det, guias, transs.join(', '), ocaBase, pdfs, detEnv, e.notaReseller, pend, envio, docNames);

    var bcc = _s(cfg['MAIL_BCC']), cc = _s(cfg['MAIL_CC']);
    return {
      ok: true, html: html, asunto: asunto,
      destinatarios: dest.to.join(', '), a: dest.detalle.join(' + '),
      cc: cc, bcc: bcc, sinGuia: !tieneGuia, sinDestino: !dest.to.length
    };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}

// Mail al reseller de UN envío. Manual: force=true → reenvía aunque ya se haya mandado.
function CP_enviarMailEnvio(idVenta, envio) {
  var _auth = _cpUsuarioAutorizado();
  if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
  return _cpEnviarEnvioCore(_s(idVenta), _num(envio), true);
}

// Core del mail al reseller (por envío). force=false → NO reenvía si ya se mandó.
function _cpEnviarEnvioCore(idVenta, envio, force) {
  try {
    idVenta = _s(idVenta); envio = _num(envio);
    if (!idVenta) return { ok: false, mensaje: 'Falta el ID de la venta.' };
    var arr = _cpEnviosMap()[idVenta.toUpperCase()] || [];
    var e = null;
    arr.forEach(function(x) { if (x.envio === envio) e = x; });
    if (!e) return { ok: false, mensaje: 'No se encontró el envío.' };
    if (!force && e.mailReseller) return { ok: false, yaEnviado: true, mensaje: 'El mail de este envío ya fue enviado (' + e.mailReseller + ').' };

    var master = _cpMasterMap();
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    if (!parts.length) return { ok: false, mensaje: 'El envío no tiene comanda.' };

    var guias = [], transs = [], tieneGuia = true;
    parts.forEach(function(p) {
      var m = master[p.toUpperCase()];
      if (!m || !m.guia) tieneGuia = false;
      if (m) { if (m.guia) guias.push(m.guia); if (m.transportista && transs.indexOf(m.transportista) === -1) transs.push(m.transportista); }
    });
    if (!tieneGuia) return { ok: false, mensaje: 'Este envío todavía no tiene número de seguimiento (guía) en Comandas Master.' };

    var cfg = _cpConfig();
    var det = _cpDetalleVenta(idVenta);

    // destinatarios: reseller (Resellers col J) + RTV (col C → hoja RTV) + fijos de _CONFIG
    var dest = _cpDestinatariosEnvio(det, cfg);
    var to = dest.to, detalleDest = dest.detalle;
    if (!to.length) return { ok: false, mensaje: 'No hay destinatarios: falta el mail del reseller (Resellers), del RTV (hoja RTV) o MAIL_DESTINATARIOS en _CONFIG.' };

    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);          // lo que se envió en ESTE envío
    var pend = _cpPendingVenta(det, arr);                    // lo que todavía falta de la venta
    var ocaBase = cfg['OCA_TRACKING_URL'] || '';

    // Adjuntos: el PDF de la comanda de ESTE envío siempre; los "documentos definidos"
    // (carpeta CP_DOCS_FOLDER_ID) SOLO en el primer envío de la venta. "Primer envío" =
    // ningún otro envío de la venta se mandó todavía al reseller (robusto ante borrados/reintentos).
    var esPrimerEnvio = !arr.some(function(x) { return x.envio !== e.envio && x.mailReseller; });
    var docsArch = esPrimerEnvio ? _cpDocsDefinidosArchivos() : [];   // 1 solo listado de la carpeta
    var docNames = docsArch.map(function(a) { return a.name; });      // nombres para el cuerpo del mail

    var asunto = (cfg['MAIL_ASUNTO'] || 'Despacho {IDVENTA} · Comanda {COMANDA} — {CLIENTE}')
      .replace('{IDVENTA}', idVenta).replace('{COMANDA}', parts.join('/')).replace('{CLIENTE}', det.razonSocial || det.reseller || '');
    var html = _cpMailHtml(idVenta, parts, det, guias, transs.join(', '), ocaBase, pdfs, detEnv, e.notaReseller, pend, envio, docNames);

    var opts = { htmlBody: html, name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO' };
    if (cfg['MAIL_CC'])  opts.cc  = cfg['MAIL_CC'];
    if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];

    var adjuntos = _cpPdfBlobs(parts);
    docsArch.forEach(function(a) { try { adjuntos.push(DriveApp.getFileById(a.id).getBlob()); } catch (eD) { Logger.log('adjuntar doc ' + a.name + ': ' + eD); } });
    if (adjuntos.length) opts.attachments = adjuntos;

    try {
      _cpGmailSend(to.join(','), asunto, 'Envío ' + envio + ' · comanda ' + parts.join('/') + ' — ver versión HTML.', opts);
    } catch (se) {
      // registrar el error en la hoja (col Estado) para que quede visible; el auto-trigger reintentará.
      try { _cpEnviosHoja().getRange(e.rowIdx, 9).setValue('MAIL ERROR · ' + _fmtTs(new Date()) + ' · ' + String(se && se.message ? se.message : se)); } catch (_) {}
      return { ok: false, mailError: true, mensaje: 'No se pudo enviar el mail: ' + String(se && se.message ? se.message : se) };
    }
    _cpMarcarMailEnvio(e.rowIdx, guias.join('/'), transs.join(', '));
    _cpAuditar('Mail reseller', idVenta, envio, 'a ' + to.join(', '));
    return { ok: true, destinatarios: to.join(', '), a: detalleDest.join(' + '), envio: envio };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}

function _cpMarcarMailEnvio(rowIdx, guia, transportista) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (e) {}
  try {
    var h = _cpEnviosHoja();
    if (guia)          h.getRange(rowIdx, 7).setValue(guia);
    if (transportista) h.getRange(rowIdx, 8).setValue(transportista);
    h.getRange(rowIdx, 9).setValue(CP_DESPACHADO);
    h.getRange(rowIdx, 11).setValue('SÍ · ' + _fmtTs(new Date()));
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

// AUTOMÁTICO (trigger horario). Manda el mail de los envíos que ya tienen guía y aún no se enviaron.
// Solo actúa si en _CONFIG está AUTO_MAIL_DESPACHO = SI.
function CP_autoMailEnvios() {
  var cfg = _cpConfig();
  var v = _s(cfg['AUTO_MAIL_DESPACHO']).toUpperCase();
  if (v.indexOf('SI') !== 0 && v.indexOf('SÍ') !== 0) { Logger.log('AUTO_MAIL_DESPACHO desactivado en _CONFIG.'); return; }
  var mapAll = _cpEnviosMap(), enviados = 0;
  Object.keys(mapAll).forEach(function(k) {
    mapAll[k].forEach(function(e) {
      if (e.mailReseller) return;
      var r = _cpEnviarEnvioCore(k, e.envio, false);
      if (r && r.ok) enviados++;
    });
  });
  Logger.log('CP_autoMailEnvios: ' + enviados + ' mail(s) enviado(s).');
}
// Alias por si quedó instalado el trigger viejo.
function CP_autoMailDespacho() { return CP_autoMailEnvios(); }

// Reintenta manualmente TODOS los envíos con guía cuyo mail al reseller aún no salió
// (incluye los que quedaron en "MAIL ERROR"). Independiente de AUTO_MAIL_DESPACHO.
// Devuelve { ok, enviados, fallidos:[{idVenta, envio, mensaje}] }.
function CP_reintentarFallidos() {
  try {
    var _auth = _cpUsuarioAutorizado();
    if (!_auth.ok) return { ok: false, noAutorizado: true, mensaje: _auth.mensaje };
    var mapAll = _cpEnviosMap(), enviados = 0, fallidos = [];
    Object.keys(mapAll).forEach(function(k) {
      mapAll[k].forEach(function(e) {
        if (e.mailReseller) return;                 // ya salió
        if (!_s(e.comanda)) return;                 // sin comanda → no aplica
        var r = _cpEnviarEnvioCore(k, e.envio, false);
        if (r && r.ok) enviados++;
        else if (r && r.mailError) fallidos.push({ idVenta: k, envio: e.envio, mensaje: r.mensaje || 'error' });
        // los que aún no tienen guía/destinatario se ignoran (no son "fallidos", sólo no están listos)
      });
    });
    if (enviados || fallidos.length) _cpAuditar('Reintentar fallidos', '', '', enviados + ' enviado(s), ' + fallidos.length + ' fallo(s)');
    return { ok: true, enviados: enviados, fallidos: fallidos };
  } catch (e) {
    return { ok: false, mensaje: String(e && e.message ? e.message : e) };
  }
}

// Recordatorio a Sole cada X horas (RECORDATORIO_HORAS, def 24) para comandas que
// siguen SIN despacharse (sin guía). Usa ScriptProperties para no repetir antes de X horas.
function CP_recordatoriosSole() {
  var cfg = _cpConfig();
  var sole = cfg['MAIL_APROBACION'];
  if (!sole) { Logger.log('CP_recordatoriosSole: falta MAIL_APROBACION.'); return; }
  var horas = _num(cfg['RECORDATORIO_HORAS']) || 24;
  var master = _cpMasterMap();
  var props = PropertiesService.getScriptProperties();
  var now = Date.now(), enviados = 0;
  var mapAll = _cpEnviosMap();
  Object.keys(mapAll).forEach(function(k) {
    mapAll[k].forEach(function(e) {
      if (!e.fechaTs) return;
      var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
      if (!parts.length) return;
      var tieneGuia = true;
      parts.forEach(function(p) { var m = master[p.toUpperCase()]; if (!m || !m.guia) tieneGuia = false; });
      if (tieneGuia) return;                              // ya despachado → no molestar a Sole
      if ((now - e.fechaTs) / 3600000 < horas) return;   // todavía no cumple X horas
      var pk = 'rem_' + k + '_' + e.envio;
      var last = _num(props.getProperty(pk));
      if (last && (now - last) / 3600000 < horas) return; // recordado hace < X horas
      var det = _cpDetalleVenta(k);
      var okr = _cpEnviarRecordatorioSole(sole, k, e, cfg, det, Math.floor((now - e.fechaTs) / 3600000));
      if (okr) { props.setProperty(pk, String(now)); enviados++; }
    });
  });
  Logger.log('CP_recordatoriosSole: ' + enviados + ' recordatorio(s) enviado(s).');
}

function _cpEnviarRecordatorioSole(sole, idVenta, e, cfg, det, horas) {
  try {
    var parts = e.comanda.split('/').map(function(s) { return s.trim(); }).filter(Boolean);
    var pdfs = [];
    parts.forEach(function(p) { var pdf = _cpBuscarPdf(p); if (pdf) pdfs.push({ comanda: p, url: pdf.url, name: pdf.name }); });
    var detEnv = _cpDetalleEnvio(det, e.productos);
    var arr = _cpEnviosMap()[String(idVenta).toUpperCase()] || [];
    var pend = _cpPendingVenta(det, arr);   // faltantes de la venta tras este envío
    var nota = 'RECORDATORIO — esta comanda sigue sin despacharse hace ~' + horas + ' h.' + (e.notaAprob ? ' · ' + e.notaAprob : '');
    var asunto = '⏰ RECORDATORIO · ' + (cfg['ASUNTO_APROBACION'] || 'APROBAR MC · {COMANDA} · {IDVENTA}')
      .replace('{COMANDA}', parts.join('/')).replace('{IDVENTA}', idVenta);
    var opts = {
      htmlBody: _cpMailAprobacionHtml(idVenta, parts, det, pdfs, nota, detEnv, e.envio, pend),
      name: cfg['MAIL_REMITENTE_NOMBRE'] || 'BIDCOMAGRO'
    };
    if (cfg['MAIL_BCC']) opts.bcc = cfg['MAIL_BCC'];
    var adj = _cpPdfBlobs(parts);
    if (adj.length) opts.attachments = adj;
    _cpGmailSend(sole, asunto, 'Recordatorio: autorizar comanda ' + e.comanda + ' en Masterchief.', opts);
    _cpAuditar('Recordatorio Sole', idVenta, e.envio, 'a ' + sole + ' · ~' + horas + ' h sin despachar');
    return true;
  } catch (err) { Logger.log('_cpEnviarRecordatorioSole ' + idVenta + '#' + e.envio + ': ' + err); return false; }
}

// SETUP: instala el trigger de recordatorios a Sole (revisa cada 6 h; recuerda cada RECORDATORIO_HORAS). Correr una vez.
function CP_setupRecordatorios() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'CP_recordatoriosSole'; });
  if (ya) { Logger.log('El trigger CP_recordatoriosSole ya existe.'); return; }
  ScriptApp.newTrigger('CP_recordatoriosSole').timeBased().everyHours(6).create();
  Logger.log('✅ Trigger de recordatorios instalado (revisa cada 6 h; recuerda cada ' + (_cpConfig()['RECORDATORIO_HORAS'] || 24) + ' h una comanda sin despachar).');
}

// SETUP: instala el trigger horario del envío automático. Correr una vez.
function CP_setupAutoMail() {
  var ya = ScriptApp.getProjectTriggers().some(function(t) { return t.getHandlerFunction() === 'CP_autoMailEnvios'; });
  if (ya) { Logger.log('El trigger CP_autoMailEnvios ya existe.'); return; }
  ScriptApp.newTrigger('CP_autoMailEnvios').timeBased().everyMinutes(30).create();
  Logger.log('✅ Trigger automático instalado (cada 30 min). Activá AUTO_MAIL_DESPACHO=SI en _CONFIG para que envíe.');
}

// ── Cuerpos de mail ──
function _cpMailItemsRows(items, esc) {
  return (items || []).map(function(it) {
    return '<tr><td style="padding:5px 8px;border-bottom:1px solid #eee;font-weight:700;text-align:center">' + esc(it.cant) + '×</td>' +
           '<td style="padding:5px 8px;border-bottom:1px solid #eee;font-family:monospace;font-weight:700">' + esc(it.sku || '—') + '</td>' +
           '<td style="padding:5px 8px;border-bottom:1px solid #eee;color:#555">' + esc(it.desc || '') + '</td></tr>';
  }).join('');
}

// Bloque HTML de FALTANTES (lo que queda pendiente de enviar de la venta).
// Devuelve '' si no hay nada pendiente → así no sale nada en el mail.
function _cpPendBloqueHtml(pendientes, esc, titulo) {
  if (!pendientes || !pendientes.length) return '';
  var rows = pendientes.map(function(f){
    return '<tr><td style="padding:5px 8px;border-bottom:1px solid #f3d9d9;font-weight:700;text-align:center;color:#c0392b">'+esc(f.cantidad)+'×</td>'+
           '<td style="padding:5px 8px;border-bottom:1px solid #f3d9d9;font-family:monospace;font-weight:700">'+esc(f.sku||'—')+'</td>'+
           '<td style="padding:5px 8px;border-bottom:1px solid #f3d9d9;color:#555">'+esc(f.descripcion||'')+'</td></tr>';
  }).join('');
  return '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#c0392b;margin:16px 0 8px">'+esc(titulo)+'</div>'+
    '<table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #e6b3b3;border-radius:8px;overflow:hidden;background:#fff7f7">'+rows+'</table>';
}

// Mail al reseller: un ENVÍO (lo despachado en este envío + guía + lo que queda pendiente).
function _cpMailHtml(idVenta, comandas, det, guias, transportista, ocaBase, pdfs, itemsEnviados, notaReseller, pendientes, envioNum, docsAdjuntos) {
  function esc(s){ return String(s==null?'':s).replace(/[&<>]/g,function(m){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[m];}); }
  // Sección explícita de adjuntos: la(s) comanda(s) de este envío + los documentos definidos (1er envío).
  var adjItems = [];
  (pdfs||[]).forEach(function(p){ adjItems.push('📄 Comanda ' + esc(p.comanda) + (p.name ? ' <span style="color:#888">('+esc(p.name)+')</span>' : '')); });
  (docsAdjuntos||[]).forEach(function(n){ adjItems.push('📎 ' + esc(n)); });
  var adjBloque = adjItems.length
    ? '<div style="background:#f0faf4;border:1px solid #b7e3c8;border-left:3px solid #1a9e4a;border-radius:0 8px 8px 0;padding:12px 15px;margin:0 0 16px;font-size:13px;color:#1a1f2e">'
      + '<div style="font-weight:700;margin-bottom:6px">📎 Documentos adjuntos a este correo</div>'
      + '<ul style="margin:0;padding-left:20px;line-height:1.8">' + adjItems.map(function(x){ return '<li>'+x+'</li>'; }).join('') + '</ul>'
      + '</div>'
    : '';
  var guiasHtml = (guias||[]).map(function(g){
    var url = ocaBase ? ocaBase.replace('{GUIA}', encodeURIComponent(g)) : '';
    return url ? '<a href="'+url+'" style="color:#00a3e0;font-weight:700;text-decoration:none">'+esc(g)+' ↗</a>' : '<b>'+esc(g)+'</b>';
  }).join(' &nbsp;·&nbsp; ') || '—';
  var pdfHtml = (pdfs && pdfs.length)
    ? pdfs.map(function(p){ return '<a href="'+esc(p.url)+'" style="color:#00a3e0;font-weight:700;text-decoration:none">📄 '+esc(p.comanda)+' ↗</a>'; }).join(' &nbsp;·&nbsp; ')
    : '—';
  var notaBloque = _s(notaReseller)
    ? '<div style="background:#eef7ff;border:1px solid #cfe6fb;border-left:3px solid #00a3e0;border-radius:0 8px 8px 0;padding:12px 15px;margin:0 0 16px;font-size:13px;color:#1a1f2e;white-space:pre-wrap"><b>Mensaje:</b> '+esc(notaReseller)+'</div>'
    : '';
  var chip = function(l,v){ return '<tr><td style="padding:4px 0;color:#888;width:150px">'+esc(l)+'</td><td style="padding:4px 0;color:#1a1f2e;font-weight:600">'+v+'</td></tr>'; };

  // Los faltantes NO van en el mail de despacho (reseller+RTV): solo lo que se envía en esta
  // comanda. El seguimiento de pendientes sigue vivo en la UI y en PENDIENTES_ENTREGA.
  return ''+
  '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:640px;margin:0 auto;color:#1a1f2e">'+
    '<div style="background:#00a3e0;color:#fff;padding:18px 22px;border-radius:10px 10px 0 0">'+
      '<div style="font-size:12px;opacity:.85;letter-spacing:.05em">BIDCOMAGRO · DESPACHO' + (envioNum ? ' · ENVÍO ' + esc(envioNum) : '') + '</div>'+
      '<div style="font-size:20px;font-weight:800;margin-top:2px">Comanda '+esc((comandas||[]).join(' / '))+' cargada</div>'+
    '</div>'+
    '<div style="border:1px solid #e0e3e8;border-top:none;border-radius:0 0 10px 10px;padding:20px 22px">'+
      notaBloque+
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px">'+
        chip('ID Venta', esc(idVenta))+
        chip('Comanda(s)', esc((comandas||[]).join(' / ')))+
        chip('Razón Social', esc(det.razonSocial||'—'))+
        chip('Reseller', esc(det.reseller||'—'))+
        chip('Operación', esc(det.operacion||'—'))+
        chip('Transportista', esc(transportista||'—'))+
        chip('Guía de seguimiento', guiasHtml)+
        chip('PDF comanda', pdfHtml)+
      '</table>'+
      adjBloque+
      '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#888;margin:6px 0 8px">Detalle despachado (este envío)</div>'+
      '<table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #e0e3e8;border-radius:8px;overflow:hidden">'+
        (_cpMailItemsRows(itemsEnviados, esc) || '<tr><td style="padding:8px;color:#999">Sin detalle</td></tr>')+
      '</table>'+
      '<div style="margin-top:18px;font-size:11px;color:#aaa">Enviado automáticamente desde Comandas · Carga Masterchief.</div>'+
    '</div>'+
  '</div>';
}

// Mail a Sole: aprobar la comanda de un ENVÍO en Masterchief (con detalle de ese envío + PDF adjunto).
function _cpMailAprobacionHtml(idVenta, comandas, det, pdfs, notaAprob, itemsEnviados, envioNum, pendientes) {
  function esc(s){ return String(s==null?'':s).replace(/[&<>]/g,function(m){return {'&':'&amp;','<':'&lt;','>':'&gt;'}[m];}); }
  var pdfHtml = (pdfs && pdfs.length)
    ? pdfs.map(function(p){ return '<a href="'+esc(p.url)+'" style="color:#00a3e0;font-weight:700;text-decoration:none">📄 '+esc(p.comanda)+' ↗</a>'; }).join(' &nbsp;·&nbsp; ')
    : '—';
  var notaBloque = _s(notaAprob)
    ? '<div style="background:#fff6e6;border:1px solid #f3dca6;border-left:3px solid #d4890a;border-radius:0 8px 8px 0;padding:12px 15px;margin:0 0 16px;font-size:13px;color:#1a1f2e;white-space:pre-wrap"><b>Nota del operador:</b> '+esc(notaAprob)+'</div>'
    : '';
  var chip = function(l,v){ return '<tr><td style="padding:4px 0;color:#888;width:150px">'+esc(l)+'</td><td style="padding:4px 0;color:#1a1f2e;font-weight:600">'+v+'</td></tr>'; };

  return ''+
  '<div style="font-family:-apple-system,Segoe UI,Arial,sans-serif;max-width:640px;margin:0 auto;color:#1a1f2e">'+
    '<div style="background:#d4890a;color:#fff;padding:18px 22px;border-radius:10px 10px 0 0">'+
      '<div style="font-size:12px;opacity:.9;letter-spacing:.05em">BIDCOMAGRO · APROBACIÓN MASTERCHIEF' + (envioNum ? ' · ENVÍO ' + esc(envioNum) : '') + '</div>'+
      '<div style="font-size:20px;font-weight:800;margin-top:2px">Autorizar comanda '+esc((comandas||[]).join(' / '))+'</div>'+
    '</div>'+
    '<div style="border:1px solid #e0e3e8;border-top:none;border-radius:0 0 10px 10px;padding:20px 22px">'+
      '<p style="margin:0 0 14px;font-size:13.5px">Se cargó una comanda en Masterchief y necesita tu <b>autorización</b>.</p>'+
      notaBloque+
      '<table style="width:100%;border-collapse:collapse;font-size:13px;margin-bottom:16px">'+
        chip('Comanda(s)', esc((comandas||[]).join(' / ')))+
        chip('ID Venta', esc(idVenta))+
        chip('Razón Social', esc(det.razonSocial||'—'))+
        chip('Reseller', esc(det.reseller||'—'))+
        chip('Operación', esc(det.operacion||'—'))+
        chip('PDF comanda', pdfHtml)+
      '</table>'+
      '<div style="font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#888;margin:6px 0 8px">Detalle a cargar (este envío)</div>'+
      '<table style="width:100%;border-collapse:collapse;font-size:12.5px;border:1px solid #e0e3e8;border-radius:8px;overflow:hidden">'+
        (_cpMailItemsRows(itemsEnviados, esc) || '<tr><td style="padding:8px;color:#999">Sin detalle</td></tr>')+
      '</table>'+
      '<div style="margin-top:18px;font-size:11px;color:#aaa">Enviado automáticamente desde Comandas · Carga Masterchief.</div>'+
    '</div>'+
  '</div>';
}

/* ════════════════════════════════════════════════════════════
   REPORTE DE TIEMPOS — CARGAR → comanda cargada
   delta = fecha del envío (cuando el operador registró la comanda)
           menos el momento MÁS ANTIGUO en que se marcó CARGAR en la venta.
   origen 'edit' = exacto | 'deteccion' = aproximado.
════════════════════════════════════════════════════════════ */

// Mapa liviano leyendo Ventas una sola vez: { IDVENTA: {idVenta, reseller, cliente} }
function _cpInfoVentasMap() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) return {};
  var col = det.col, m = {};
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    var idv = _s(col.idVenta > -1 ? rows[i][col.idVenta] : ''); if (!idv) continue;
    var k = idv.toUpperCase();
    if (!m[k]) m[k] = {
      idVenta:  idv,
      reseller: _s(col.reseller > -1 ? rows[i][col.reseller] : ''),
      cliente:  _s(col.razonSocial > -1 ? rows[i][col.razonSocial] : '')
    };
  }
  return m;
}

// Estadística simple de un array de números: n, promedio, mediana, min, max.
function _cpStats(arr) {
  if (!arr || !arr.length) return { n: 0, prom: 0, mediana: 0, min: 0, max: 0 };
  var s = arr.slice().sort(function(a, b) { return a - b; });
  var n = s.length, sum = s.reduce(function(a, b) { return a + b; }, 0), mid = Math.floor(n / 2);
  var med = (n % 2) ? s[mid] : (s[mid - 1] + s[mid]) / 2;
  var r1 = function(x) { return Math.round(x * 10) / 10; };
  return { n: n, prom: r1(sum / n), mediana: r1(med), min: r1(s[0]), max: r1(s[n - 1]) };
}

// Devuelve { ok, rows:[por envío], pendientes:[marcadas CARGAR sin envío], resumen:{primer, todos} }.
function CP_reporteTiempos() {
  try {
    var logMap = _cpLogMap();
    // CARGAR más antiguo por venta (+ si es exacto)
    var cargarPorVenta = {};
    Object.keys(logMap).forEach(function(key) {
      var idv = key.split('||')[0], reg = logMap[key], cur = cargarPorVenta[idv];
      if (!cur || reg.ts < cur.ts) cargarPorVenta[idv] = { ts: reg.ts, origen: reg.origen };
    });
    var info = _cpInfoVentasMap();
    var enviosMap = _cpEnviosMap();
    var rows = [], dPrim = [], dTodos = [], dDesp = [];
    Object.keys(enviosMap).forEach(function(idvU) {
      var arr = enviosMap[idvU], carg = cargarPorVenta[idvU];
      var cargTs = carg ? carg.ts.getTime() : null, exacto = carg ? (carg.origen === 'edit') : false;
      var iv = info[idvU] || {};
      arr.forEach(function(e) {
        if (!e.fechaTs) return;
        var horas = cargTs != null ? Math.round(((e.fechaTs - cargTs) / 3600000) * 10) / 10 : null;
        if (horas != null && horas >= 0) { dTodos.push(horas); if (e.envio === 1) dPrim.push(horas); }
        // tramo despacho: comanda cargada (fecha del envío) → mail al reseller (proxy del despacho)
        var hDesp = e.mailResellerTs ? Math.round(((e.mailResellerTs - e.fechaTs) / 3600000) * 10) / 10 : null;
        if (hDesp != null && hDesp >= 0) dDesp.push(hDesp);
        rows.push({
          idVenta: iv.idVenta || idvU, reseller: iv.reseller || '', cliente: iv.cliente || '',
          envio: e.envio, comanda: e.comanda,
          cargarStr: cargTs != null ? _fmtTs(new Date(cargTs)) : '',
          origen: carg ? (exacto ? 'exacto' : 'aprox') : 'sin registro',
          cargadaStr: e.fechaStr || '', operador: e.operador || '',
          despachoStr: e.mailResellerTs ? _fmtTs(new Date(e.mailResellerTs)) : '',
          horas: horas, horasDespacho: hDesp, primero: (e.envio === 1)
        });
      });
    });
    rows.sort(function(a, b) { return a.idVenta === b.idVenta ? a.envio - b.envio : (a.idVenta < b.idVenta ? -1 : 1); });
    // pendientes: marcadas CARGAR pero sin ningún envío todavía
    var now = Date.now(), pend = [];
    Object.keys(cargarPorVenta).forEach(function(idvU) {
      if (enviosMap[idvU]) return;
      var carg = cargarPorVenta[idvU], iv = info[idvU] || {};
      pend.push({ idVenta: iv.idVenta || idvU, reseller: iv.reseller || '',
        horasAbierto: Math.round(((now - carg.ts.getTime()) / 3600000) * 10) / 10,
        origen: (carg.origen === 'edit' ? 'exacto' : 'aprox') });
    });
    pend.sort(function(a, b) { return b.horasAbierto - a.horasAbierto; });
    return { ok: true, rows: rows, pendientes: pend, resumen: { primer: _cpStats(dPrim), todos: _cpStats(dTodos), despacho: _cpStats(dDesp) } };
  } catch (e) { return { ok: false, mensaje: String(e && e.message ? e.message : e) }; }
}

/* ── DIAGNÓSTICO (correr desde el editor si algo falla) ──── */
// Ejecutá esta función y mirá los logs: muestra la fila de encabezados
// detectada y a qué índice quedó mapeada cada columna.
function CP_debugColumnas() {
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) { Logger.log('No se detectó fila de encabezados con ID_Venta + ID_Entrega.'); return; }
  Logger.log('Fila de encabezados (0-based): ' + det.headerRow);
  Logger.log('Encabezados: ' + JSON.stringify(det.headers));
  Logger.log('Mapa columnas -> índice: ' + JSON.stringify(det.col));
  var flag = CP_FLAG.toUpperCase(), n = 0;
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    if (_s(rows[i][det.col.idEntrega]).toUpperCase() === flag) n++;
  }
  Logger.log('Filas con ID_Entrega = ' + CP_FLAG + ': ' + n);
}

// Diagnóstico crudo: qué pestañas hay, qué lee la hoja configurada y cómo
// normalizan las primeras filas. Correr desde el editor y mirar los logs.
function CP_debugCrudo() {
  var ss = _cpSS(CP_SS_ID);
  Logger.log('Pestañas en el archivo: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(' | '));
  Logger.log('CP_TAB configurada: "' + CP_TAB + '"');

  var h = ss.getSheetByName(CP_TAB);
  if (!h) { Logger.log('>> getSheetByName("' + CP_TAB + '") devolvió NULL. El nombre no coincide.'); return; }

  var rng = h.getDataRange();
  Logger.log('Dimensiones: ' + rng.getNumRows() + ' filas x ' + rng.getNumColumns() + ' columnas');
  var rows = rng.getValues();

  var maxScan = Math.min(4, rows.length);
  for (var r = 0; r < maxScan; r++) {
    Logger.log('--- Fila ' + r + ' (0-based) ---');
    Logger.log('  CRUDO: ' + JSON.stringify(rows[r]));
    Logger.log('  NORM : ' + JSON.stringify(rows[r].map(_norm)));
  }
}

// Diagnóstico de mails: cuántos resellers/RTV mapeó y ejemplos.
function CP_debugMails() {
  var rm = _cpResellerMap(), rt = _cpRtvMailMap();
  var conMail = Object.keys(rm).filter(function(k){ return rm[k].mail; });
  Logger.log('Resellers leídos: ' + Object.keys(rm).length + ' — con mail: ' + conMail.length);
  Object.keys(rm).slice(0, 6).forEach(function(k) { Logger.log('  ' + k + ' → mail:' + (rm[k].mail||'—') + ' | RTV:' + (rm[k].rtv||'—')); });
  Logger.log('RTV con mail (hoja RTV): ' + Object.keys(rt).length);
  Object.keys(rt).slice(0, 6).forEach(function(k) { Logger.log('  ' + k + ' → ' + rt[k]); });
}

// Diagnóstico de KITS: cuántos kits leyó y la composición de un par de ejemplos.
function CP_debugKits() {
  var map = _cpKitMap();
  var kits = Object.keys(map);
  Logger.log('KITS leídos: ' + kits.length);
  var ejemplos = ['T20COMBO', 'KITRTK3'].concat(kits.slice(0, 3));
  ejemplos.forEach(function(k) {
    var m = map[_kitKey(k)];
    if (!m) { Logger.log('  ' + k + ' → (no está)'); return; }
    Logger.log('  ' + k + ' → ' + m.orden.map(function(cu) {
      var comp = m.comps[cu];
      return comp.cant + 'x ' + comp.sku + ' (' + comp.desc + ')';
    }).join(' | '));
  });
}

/* ════════════════════════════════════════════════════════════
   DIAGNÓSTICO: ¿por qué una venta con envío no pasa a parcial/completo?
   Corré CP_diagEnvios() desde el editor y mirá el registro (Ver → Registro).
   Compara los ID_Venta de la hoja ENVIOS contra los ID_Venta que hoy
   tienen el flag CARGAR en Ventas → revela si el problema es:
     (a) el ID_Venta guardado no matchea (tipo/formato, ceros a la izquierda), o
     (b) la venta ya no está en CARGAR (desapareció de la lista).
════════════════════════════════════════════════════════════ */
function CP_diagEnvios() {
  var out = [];
  var norm = function(v) { return _s(v).toUpperCase(); };          // igual que las claves reales
  var normNum = function(v) { var s = _s(v).replace(/^0+/, ''); return s; }; // sin ceros a la izquierda

  // 1) ID_Venta con envíos (hoja ENVIOS)
  var envMap = _cpEnviosMap();
  var envKeys = Object.keys(envMap);
  out.push('ENVIOS: ' + envKeys.length + ' venta(s) con envío(s) registrados.');

  // 2) ID_Venta con flag CARGAR (hoja Ventas)
  var rows = _cpHoja().getDataRange().getValues();
  var det = _cpDetectar(rows);
  if (!det) { out.push('❌ No detecté encabezados en Ventas.'); Logger.log(out.join('\n')); return out.join('\n'); }
  var col = det.col, flag = CP_FLAG.toUpperCase();
  var cargar = {};      // clave normal → {raw, tipo}
  var cargarNum = {};   // clave sin ceros → clave normal (para detectar near-match)
  for (var i = det.headerRow + 1; i < rows.length; i++) {
    if (norm(rows[i][col.idEntrega]) !== flag) continue;
    var raw = rows[i][col.idVenta];
    var k = norm(raw); if (!k) continue;
    cargar[k] = { raw: raw, tipo: typeof raw };
    cargarNum[normNum(raw)] = k;
  }
  out.push('CARGAR: ' + Object.keys(cargar).length + ' venta(s) marcadas hoy.');
  out.push('');

  // 3) Por cada venta con envíos, ¿matchea?
  envKeys.forEach(function(k) {
    if (cargar[k]) {
      out.push('✅ "' + k + '" (' + envMap[k].length + ' env) → coincide con una venta en CARGAR.');
    } else if (cargarNum[normNum(k)]) {
      out.push('⚠️ "' + k + '" (' + envMap[k].length + ' env) → NO coincide EXACTO, pero sí sin ceros a la izquierda con "' + cargarNum[normNum(k)] + '" (Ventas tipo ' + cargar[cargarNum[normNum(k)]].tipo + '). ⇒ mismatch de formato/tipo.');
    } else {
      out.push('❌ "' + k + '" (' + envMap[k].length + ' env) → NO está en CARGAR (ni exacto ni por número). ⇒ la venta salió del flag CARGAR.');
    }
  });

  var txt = out.join('\n');
  Logger.log(txt);
  return txt;
}

/* ════════════════════════════════════════════════════════════
   DIAGNÓSTICO: ¿por qué NO sale el mail automático al reseller + RTV?
   Corré CP_diagAutoMail() desde el editor y mirá el registro (Ver → Registro).
   Chequea, en orden, TODO lo que tiene que estar bien para que el trigger
   horario CP_autoMailEnvios mande el mail:
     1) EMAIL_PRUEBA vacío (si tiene un mail, TODO se redirige a esa dirección
        y el reseller/RTV NO reciben nada → causa #1 después de usar el modo prueba).
     2) AUTO_MAIL_DESPACHO = SI (si no, el trigger corre pero no manda nada).
     3) El trigger CP_autoMailEnvios instalado (CP_setupAutoMail).
     4) Por cada envío sin mail: ¿tiene guía? ¿tiene destinatarios?
   Es SOLO LECTURA: no manda ningún correo.
════════════════════════════════════════════════════════════ */
function CP_diagAutoMail() {
  var out = [];
  var cfg = _cpConfig();

  // 1) MODO PRUEBA (la causa más común después de testear)
  var prueba = _s(cfg['EMAIL_PRUEBA']);
  if (prueba) {
    out.push('🔴 EMAIL_PRUEBA = "' + prueba + '"  → MODO PRUEBA ACTIVO.');
    out.push('   TODOS los correos (reseller, RTV, Sole) se redirigen a esa dirección.');
    out.push('   El reseller/RTV NO reciben nada. Para volver a producción: VACIÁ EMAIL_PRUEBA en _CONFIG.');
  } else {
    out.push('✅ EMAIL_PRUEBA vacío (producción: los mails van a los destinatarios reales).');
  }

  // 2) Interruptor del automático
  var auto = _s(cfg['AUTO_MAIL_DESPACHO']).toUpperCase();
  var autoOn = (auto.indexOf('SI') === 0 || auto.indexOf('SÍ') === 0);
  out.push((autoOn ? '✅' : '🔴') + ' AUTO_MAIL_DESPACHO = "' + (_s(cfg['AUTO_MAIL_DESPACHO']) || '(vacío)') + '"' +
           (autoOn ? '' : '  → el envío automático está DESACTIVADO. Poné AUTO_MAIL_DESPACHO=SI en _CONFIG.'));

  // 3) Trigger instalado
  var trg = ScriptApp.getProjectTriggers().filter(function(t) { return t.getHandlerFunction() === 'CP_autoMailEnvios'; });
  out.push((trg.length ? '✅' : '🔴') + ' Trigger CP_autoMailEnvios: ' +
           (trg.length ? trg.length + ' instalado(s) (corre cada 30 min).' : 'NO instalado → corré CP_setupAutoMail() una vez.'));

  // 4) Estado de los envíos pendientes de mail
  out.push('');
  var mapAll = _cpEnviosMap();
  var master = _cpMasterMap();
  // Pre-lecturas ÚNICAS: antes se releía Ventas/Resellers/RTV por CADA envío con guía →
  // con un backlog grande de envíos sin mail, el diagnóstico se colgaba y no mostraba nada.
  var ventaMap    = _cpVentaResellerRtvMap();   // { IDVENTA: {reseller, rtv} } (1 lectura de Ventas)
  var resellerMap = _cpResellerMap();           // 1 lectura de Resellers
  var rtvMap      = _cpRtvMailMap();            // 1 lectura de RTV
  var total = 0, yaMail = 0, sinComanda = 0, sinGuia = 0, sinDest = 0, listos = 0;
  var detListos = [], detSinGuia = [], detSinDest = [];

  Object.keys(mapAll).forEach(function(k) {
    mapAll[k].forEach(function(e) {
      total++;
      if (e.mailReseller) { yaMail++; return; }
      var parts = _s(e.comanda).split('/').map(function(s){ return s.trim(); }).filter(Boolean);
      if (!parts.length) { sinComanda++; return; }
      var tieneGuia = parts.every(function(p){ var m = master[p.toUpperCase()]; return m && m.guia; });
      if (!tieneGuia) { sinGuia++; if (detSinGuia.length < 8) detSinGuia.push(k + ' env' + e.envio + ' (' + e.comanda + ')'); return; }
      var vi = ventaMap[k.toUpperCase()] || {};
      var det = { reseller: vi.reseller || '', rtv: vi.rtv || '' };
      var dest = _cpDestinatariosEnvio(det, cfg, resellerMap, rtvMap);
      if (!dest.to.length) { sinDest++; if (detSinDest.length < 8) detSinDest.push(k + ' env' + e.envio + ' — ' + (det.reseller || '?')); return; }
      listos++; if (detListos.length < 8) detListos.push(k + ' env' + e.envio + ' → ' + dest.to.join(', '));
    });
  });

  out.push('ENVÍOS: ' + total + ' total · ' + yaMail + ' con mail ya enviado · ' + (total - yaMail) + ' sin mail.');
  out.push('  De los que están sin mail:');
  out.push('    ⏳ ' + sinGuia + ' esperan la guía (número de seguimiento) en Comandas Master.');
  out.push('    🔴 ' + sinDest + ' listos pero SIN destinatarios (falta mail del reseller/RTV o MAIL_DESTINATARIOS).');
  out.push('    ⚠️ ' + sinComanda + ' sin comanda cargada.');
  out.push('    ✅ ' + listos + ' LISTOS para enviar (con guía y destinatarios).');
  if (detListos.length)  out.push('       listos: ' + detListos.join(' | '));
  if (detSinGuia.length) out.push('       sin guía: ' + detSinGuia.join(' | '));
  if (detSinDest.length) out.push('       sin dest: ' + detSinDest.join(' | '));

  // Veredicto
  out.push('');
  if (prueba)           out.push('👉 CAUSA MÁS PROBABLE: MODO PRUEBA activo (EMAIL_PRUEBA). Vacialo en _CONFIG y probá de nuevo.');
  else if (!autoOn)     out.push('👉 CAUSA: AUTO_MAIL_DESPACHO no está en SI.');
  else if (!trg.length) out.push('👉 CAUSA: falta instalar el trigger (corré CP_setupAutoMail una vez).');
  else if (listos > 0)  out.push('👉 Config OK y hay ' + listos + ' envío(s) listos. Corré CP_autoMailEnvios() a mano para forzar el envío ahora, o esperá al próximo ciclo (≤30 min). Si aun así no llegan, revisá spam/cuota de Gmail.');
  else if (sinDest > 0) out.push('👉 CAUSA: hay envíos con guía pero SIN destinatarios (cargá el mail del reseller en Resellers o del RTV en la hoja RTV).');
  else                  out.push('👉 No hay envíos listos: los pendientes esperan la guía. Cuando aparezca el número de seguimiento el mail sale solo.');

  var txt = out.join('\n');
  Logger.log(txt);
  return txt;
}
