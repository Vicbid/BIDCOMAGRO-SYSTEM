// @version 1.3
// ============================================================
//  PORTAL RESELLER BIDCOM — Autenticación y acceso
// ============================================================

// Cuenta desactivada: col Q (SCHEMA.RESELLERS.ACTIVO).
// Vacío = activo. Se considera de baja si dice NO / BAJA / INACTIVO / FALSE / 0 / booleano false.
function _resellerInactivo(fila) {
  var raw = fila[SCHEMA.RESELLERS.ACTIVO];
  if (raw === false) return true;
  var v = String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
  if (!v) return false;
  return (v === 'NO' || v === 'BAJA' || v === 'INACTIVO' || v === 'INACTIVA'
          || v === 'FALSE' || v === '0' || v === 'DESACTIVADO' || v === 'DESACTIVADA');
}

// Columna "aftersales" — se busca por header (contiene "after") con fallback al índice fijo
// del schema, y se centraliza acá el criterio de qué cuenta como "sí" (antes estaba copiado
// 3 veces, con criterios ligeramente distintos entre copias).
function _resellerColAftersales(header) {
  for (var j = 0; j < header.length; j++) {
    if (header[j].indexOf('after') !== -1) return j;
  }
  return SCHEMA.RESELLERS.AFTERSALES;
}
function _resellerEsAftersales(raw) {
  var s = String(raw === null || raw === undefined ? '' : raw).trim().toUpperCase();
  return (raw === true || s === 'SI' || s === 'SÍ' || s === 'S' || s === 'YES' || s === '1' || s === 'TRUE');
}

// ── PIN: hash + migración transparente ──────────────────────────
// El PIN individual (col K) y el de grupo (col P) se guardan hasheados (SHA-256) — antes se
// guardaban y se reenviaban por mail en texto plano. Las cuentas viejas (PIN todavía en texto
// plano en la hoja) se migran solas la próxima vez que ese reseller entre con éxito: no hace
// falta correr nada a mano ni hay riesgo de dejar a toda la red afuera el día del deploy.
function _hashPin(pin) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(pin),
    Utilities.Charset.UTF_8
  );
  var hex = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    var h = b.toString(16);
    hex += (h.length === 1 ? '0' : '') + h;
  }
  return hex;
}
// true si ya es un hash SHA-256 (64 hex) — para no re-hashear un valor ya migrado.
function _esHashPin(v) { return /^[0-9a-f]{64}$/i.test(String(v || '')); }
// Compara el PIN ingresado contra lo guardado. Si lo guardado todavía está en texto plano
// (cuenta no migrada) y coincide, lo hashea ahí mismo — `getHoja` es un getter perezoso
// (solo abre la hoja para escribir cuando de verdad hace falta migrar).
function _pinCoincideYMigrar(getHoja, fila1based, col0based, guardado, ingresado) {
  var guardadoStr  = String(guardado  || '').trim();
  var ingresadoStr = String(ingresado || '').trim();
  if (!guardadoStr || !ingresadoStr) return false;
  if (_esHashPin(guardadoStr)) return guardadoStr === _hashPin(ingresadoStr);
  if (guardadoStr !== ingresadoStr) return false;
  try {
    getHoja().getRange(fila1based, col0based + 1).setValue(_hashPin(ingresadoStr));
    invalidateSheetValues(SCHEMA.SHEETS.RESELLERS);
  } catch (e) { Logger.log('_pinCoincideYMigrar: no se pudo migrar a hash — ' + e); }
  return true;
}
function _generarPinNuevo() { return String(Math.floor(1000 + Math.random() * 9000)); }

// ── Rate limiting simple contra fuerza bruta del PIN (4 dígitos = 10.000 combinaciones) ──
// Bloquea por nombre de empresa/grupo tras varios intentos fallidos seguidos, ventana de
// 10 minutos. CacheService no pide ningún scope nuevo — sin re-auth.
var _AUTH_MAX_INTENTOS = 6;
var _AUTH_VENTANA_SEG  = 600;
function _authIntentosKey(nombre) { return 'auth_fail_' + String(nombre || '').trim().toLowerCase(); }
function _authBloqueado(nombre) {
  var n = parseInt(CacheService.getScriptCache().get(_authIntentosKey(nombre)) || '0', 10);
  return n >= _AUTH_MAX_INTENTOS;
}
function _authRegistrarFallo(nombre) {
  var cache = CacheService.getScriptCache();
  var key   = _authIntentosKey(nombre);
  var n     = parseInt(cache.get(key) || '0', 10) + 1;
  cache.put(key, String(n), _AUTH_VENTANA_SEG);
}
function _authLimpiarFallos(nombre) {
  CacheService.getScriptCache().remove(_authIntentosKey(nombre));
}

// Busca el email de sesión activa en Resellers. Retorna { ok, nombre, aftersales }.
function obtenerDatosReseller() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch(e) {}
    if (!email) return { ok: false, nombre: null, aftersales: false };

    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    if (!datos || datos.length < 2) return { ok: false, nombre: null, aftersales: false };

    var header = datos[0].map(function(c){ return String(c || '').toLowerCase().trim(); });
    var colAftersales = _resellerColAftersales(header);

    var emailLow = email.toLowerCase().trim();
    for (var i = 1; i < datos.length; i++) {
      var rowEmail = String(datos[i][SCHEMA.RESELLERS.EMAIL] || '').toLowerCase().trim();
      if (!rowEmail || rowEmail !== emailLow) continue;

      var nombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
      return { ok: true, nombre: nombre, aftersales: _resellerEsAftersales(datos[i][colAftersales]) };
    }
    return { ok: false, nombre: null, aftersales: false };
  } catch(e) {
    Logger.log("obtenerDatosReseller: " + e);
    return { ok: false, nombre: null, aftersales: false };
  }
}

// Login por nombre de empresa (Col A) + PIN (Col K). Retorna { ok, nombre, aftersales }.
// Si el nombre tiene prefijo "★ " es un grupo: valida contra PIN_GRUPO (Col P) y devuelve esGrupo:true.
function validarAccesoInicial(nombre, clave) {
  try {
    var nombreRaw = String(nombre || '').trim();
    var claveB    = String(clave  || '').trim();
    if (!nombreRaw || !claveB) return { ok: false, motivo: 'campos_vacios' };

    if (_authBloqueado(nombreRaw)) return { ok: false, motivo: 'bloqueado' };

    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    if (!datos || datos.length < 2) return { ok: false, motivo: 'sin_datos' };

    var getHoja = function() { return getSheet(SCHEMA.SHEETS.RESELLERS); };

    // ── LOGIN DE GRUPO ────────────────────────────────────────
    var esLoginGrupo = (nombreRaw.indexOf('★ ') === 0);
    if (esLoginGrupo) {
      var grupoNombre = nombreRaw.slice(2).trim(); // quitar "★ "
      var resellersDelGrupo = [];
      var pinGrupoEncontrado = false;
      for (var gi = 1; gi < datos.length; gi++) {
        var rowGrupo    = String(datos[gi][SCHEMA.RESELLERS.GRUPO]     || '').trim();
        var rowPinGrupo = String(datos[gi][SCHEMA.RESELLERS.PIN_GRUPO] || '').trim();
        var rowNombreG  = String(datos[gi][SCHEMA.RESELLERS.NOMBRE]    || '').trim();
        if (rowGrupo.toLowerCase() !== grupoNombre.toLowerCase()) continue;
        if (!pinGrupoEncontrado) {
          if (!_pinCoincideYMigrar(getHoja, gi + 1, SCHEMA.RESELLERS.PIN_GRUPO, rowPinGrupo, claveB)) {
            _authRegistrarFallo(nombreRaw);
            return { ok: false, motivo: 'clave_incorrecta' };
          }
          pinGrupoEncontrado = true;
        }
        if (_resellerInactivo(datos[gi])) continue; // sucursal de baja: fuera del grupo
        if (rowNombreG) resellersDelGrupo.push(rowNombreG);
      }
      if (!pinGrupoEncontrado) { _authRegistrarFallo(nombreRaw); return { ok: false, motivo: 'grupo_no_encontrado' }; }
      if (!resellersDelGrupo.length) return { ok: false, motivo: 'inactivo' };
      _authLimpiarFallos(nombreRaw);
      return {
        ok:        true,
        nombre:    grupoNombre,
        esGrupo:   true,
        resellers: resellersDelGrupo,
        aftersales: false
      };
    }

    // ── LOGIN INDIVIDUAL (flujo original) ─────────────────────
    var nombreB = nombreRaw.toLowerCase();
    var header = datos[0].map(function(c){ return String(c || '').toLowerCase().trim(); });
    var colAft = _resellerColAftersales(header);

    for (var i = 1; i < datos.length; i++) {
      var rowNombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase();
      var rowClave  = String(datos[i][SCHEMA.RESELLERS.PIN] || '').trim();
      if (!rowNombre || !rowClave) continue;
      if (rowNombre !== nombreB) continue;
      if (!_pinCoincideYMigrar(getHoja, i + 1, SCHEMA.RESELLERS.PIN, rowClave, claveB)) {
        _authRegistrarFallo(nombreRaw);
        return { ok: false, motivo: 'clave_incorrecta' };
      }

      if (_resellerInactivo(datos[i])) return { ok: false, motivo: 'inactivo' };

      _authLimpiarFallos(nombreRaw);
      return {
        ok:         true,
        nombre:     String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim(),
        esGrupo:    false,
        aftersales: _resellerEsAftersales(datos[i][colAft])
      };
    }
    _authRegistrarFallo(nombreRaw);
    return { ok: false, motivo: 'reseller_no_encontrado' };
  } catch(e) {
    Logger.log('validarAccesoInicial: ' + e);
    return { ok: false, motivo: 'error_interno' };
  }
}

function obtenerResellers() {
  try {
    var d = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var resellers = [];
    var grupos = {};
    for (var i = 1; i < d.length; i++) {
      if (_resellerInactivo(d[i])) continue; // cuentas de baja no aparecen en el selector
      var nombre = String(d[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
      var grupo  = String(d[i][SCHEMA.RESELLERS.GRUPO]  || '').trim();
      if (nombre) resellers.push(nombre);
      if (grupo)  grupos[grupo] = 1;
    }
    resellers.sort();
    var gruposArr = Object.keys(grupos).sort();
    // Grupos van al tope con prefijo "★ " para distinguirlos visualmente
    var lista = gruposArr.map(function(g){ return '★ ' + g; }).concat(resellers);
    return lista;
  } catch(e) { return []; }
}

function obtenerListaResellers() {
  return obtenerResellers();
}

// Template compartido de los mails de acceso (antes había 3 copias casi idénticas).
// opts: { nombre, mensaje, pin (opcional — si viene, dibuja el recuadro con el PIN),
//         ctaUrl (opcional), footer (opcional) }
function _authEmailHtml(opts) {
  var pinBlock = opts.pin ? (
    "<div style='background:#f5f8fc;border:1px solid #dde3ea;border-radius:8px;padding:18px 22px;margin-bottom:24px'>" +
      "<div style='margin-bottom:10px'>" +
        "<span style='font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.06em'>Empresa</span>" +
        "<div style='font-size:15px;font-weight:700;color:#1a1a2e;margin-top:3px'>" + opts.nombre + "</div>" +
      "</div>" +
      "<div>" +
        "<span style='font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.06em'>PIN</span>" +
        "<div style='font-size:26px;font-weight:800;color:#00a3e0;letter-spacing:6px;margin-top:3px'>" + opts.pin + "</div>" +
      "</div>" +
    "</div>"
  ) : '';
  var cta = opts.ctaUrl
    ? "<a href='" + opts.ctaUrl + "' style='display:inline-block;background:#00a3e0;color:#fff;text-decoration:none;padding:11px 28px;border-radius:7px;font-weight:700;font-size:14px'>Ingresar al portal →</a>"
    : '';
  var footer = opts.footer
    ? "<p style='font-size:11px;color:#aaa;margin-top:24px;margin-bottom:0'>" + opts.footer + "</p>"
    : '';
  return "<div style='font-family:sans-serif;max-width:520px;margin:0 auto'>" +
    "<div style='background:#00a3e0;padding:20px 24px;border-radius:10px 10px 0 0'>" +
      "<span style='color:#fff;font-size:18px;font-weight:700'>Portal Resellers \xb7 BIDCOMAGRO</span>" +
    "</div>" +
    "<div style='background:#fff;border:1px solid #dde3ea;border-top:none;padding:28px 24px;border-radius:0 0 10px 10px'>" +
      "<p style='font-size:14px;color:#444;margin:0 0 18px'>Hola <strong>" + opts.nombre + "</strong>,</p>" +
      "<p style='font-size:13px;color:#555;margin:0 0 22px'>" + opts.mensaje + "</p>" +
      pinBlock + cta + footer +
    "</div>" +
  "</div>";
}

// ── Alta de accesos (PIN) para resellers nuevos ──────────────────
// Ejecutar desde el editor de Apps Script cada vez que se cargan resellers nuevos sin PIN
// (columna K vacía): genera un PIN de 4 dígitos por cada uno, lo hashea y lo guarda, y se lo
// manda por mail. No toca a nadie que ya tenga algo cargado en la columna K (hasheado, o de
// una cuenta vieja sin migrar — eso se migra solo en su próximo login, ver
// _pinCoincideYMigrar) — se puede volver a correr las veces que hagan falta sin pisar nada.
// Retorna resumen: { enviados, omitidos, errores }
function enviarCredencialesResellers() {
  var hoja      = getSheet(SCHEMA.SHEETS.RESELLERS);
  var datos     = hoja.getDataRange().getValues();
  var portalUrl = '';
  try { portalUrl = ScriptApp.getService().getUrl(); } catch(eu) {}
  var enviados  = 0, omitidos = 0, errores = [];

  for (var i = 1; i < datos.length; i++) {
    var nombre       = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
    var email        = String(datos[i][SCHEMA.RESELLERS.EMAIL]  || '').trim();
    var pinYaCargado = String(datos[i][SCHEMA.RESELLERS.PIN]    || '').trim();

    if (!nombre || !email || pinYaCargado) { omitidos++; continue; }

    var pin = _generarPinNuevo();
    try {
      hoja.getRange(i + 1, SCHEMA.RESELLERS.PIN + 1).setValue(_hashPin(pin));
      var html = _authEmailHtml({
        nombre: nombre, pin: pin,
        mensaje: 'Ya podés acceder al <strong>Portal de Resellers de BIDCOMAGRO</strong> con las siguientes credenciales:',
        ctaUrl: portalUrl,
        footer: 'Si tenés algún problema para acceder, respondé este email y te ayudamos.'
      });
      GmailApp.sendEmail(email, 'Tus credenciales — Portal Resellers BIDCOMAGRO', '', {
        htmlBody: html,
        name:     'BIDCOMAGRO · Portal Resellers',
        replyTo:  'soporteagrasdji@bidcom.com.ar'
      });
      enviados++;
      Logger.log('Credenciales enviadas → ' + nombre + ' <' + email + '>');
    } catch(e) {
      errores.push(nombre + ' (' + email + '): ' + e.message);
      Logger.log('ERROR enviando a ' + email + ': ' + e);
    }
  }
  invalidateSheetValues(SCHEMA.SHEETS.RESELLERS);

  var resumen = 'Enviados: ' + enviados + ' | Omitidos (sin nombre/email, o ya tenían PIN cargado): ' + omitidos;
  if (errores.length) resumen += ' | Errores: ' + errores.join(' / ');
  Logger.log('enviarCredencialesResellers → ' + resumen);
  return { enviados: enviados, omitidos: omitidos, errores: errores };
}

// ── Perfil self-service: leer datos de contacto del reseller ──
function RS_obtenerPerfil(resellerNombre) {
  try {
    var nombre = String(resellerNombre || '').trim().toLowerCase();
    if (!nombre) return { ok: false, error: 'Nombre inválido' };
    var d = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === nombre) {
        return {
          ok:        true,
          direccion: String(d[i][SCHEMA.RESELLERS.DIRECCION] || '').trim(),
          cp:        String(d[i][SCHEMA.RESELLERS.CP]        || '').trim(),
          localidad: String(d[i][SCHEMA.RESELLERS.LOCALIDAD] || '').trim(),
          provincia: String(d[i][SCHEMA.RESELLERS.PROVINCIA] || '').trim(),
          telefono:  String(d[i][SCHEMA.RESELLERS.TELEFONO]  || '').trim(),
          email:     String(d[i][SCHEMA.RESELLERS.EMAIL]     || '').trim()
        };
      }
    }
    return { ok: false, error: 'Reseller no encontrado' };
  } catch(e) {
    Logger.log('RS_obtenerPerfil: ' + e);
    return { ok: false, error: 'Error interno' };
  }
}

// ── Perfil self-service: actualizar datos del reseller ──────────
function RS_actualizarPerfil(resellerNombre, telefono, email, direccion, cp, localidad, provincia) {
  try {
    var nombre = String(resellerNombre || '').trim().toLowerCase();
    if (!nombre) return { ok: false, error: 'Nombre inválido' };
    var hoja = getSheet(SCHEMA.SHEETS.RESELLERS);
    var d    = hoja.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (String(d[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === nombre) {
        hoja.getRange(i + 1, SCHEMA.RESELLERS.DIRECCION + 1).setValue(String(direccion || '').trim());
        hoja.getRange(i + 1, SCHEMA.RESELLERS.CP        + 1).setValue(String(cp        || '').trim());
        hoja.getRange(i + 1, SCHEMA.RESELLERS.LOCALIDAD + 1).setValue(String(localidad || '').trim());
        hoja.getRange(i + 1, SCHEMA.RESELLERS.PROVINCIA + 1).setValue(String(provincia || '').trim());
        hoja.getRange(i + 1, SCHEMA.RESELLERS.TELEFONO  + 1).setValue(String(telefono  || '').trim());
        hoja.getRange(i + 1, SCHEMA.RESELLERS.EMAIL     + 1).setValue(String(email     || '').trim());
        invalidateSheetValues(SCHEMA.SHEETS.RESELLERS);
        return { ok: true };
      }
    }
    return { ok: false, error: 'Reseller no encontrado' };
  } catch(e) {
    Logger.log('RS_actualizarPerfil: ' + e);
    return { ok: false, error: 'Error al guardar' };
  }
}

// ── Recordatorio de clave (self-service desde el login) ──────
// El PIN ya no se guarda en texto plano, así que no se puede "recordar" el viejo: se genera
// uno nuevo, se hashea y se guarda, y se manda por mail — mismo criterio que un reset de
// contraseña común. El anterior queda sin efecto.
function RS_recordarClave(nombre) {
  try {
    var nombreB = String(nombre || '').trim().toLowerCase();
    if (!nombreB) return { ok: false, error: 'Seleccioná tu empresa primero.' };

    var hoja  = getSheet(SCHEMA.SHEETS.RESELLERS);
    var datos = hoja.getDataRange().getValues();
    var portalUrl = '';
    try { portalUrl = ScriptApp.getService().getUrl(); } catch(eu) {}

    for (var i = 1; i < datos.length; i++) {
      var rowNombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
      if (rowNombre.toLowerCase() !== nombreB) continue;

      if (_resellerInactivo(datos[i])) return { ok: false, error: 'Esta cuenta está desactivada. Contactá a soporte.' };

      var email = String(datos[i][SCHEMA.RESELLERS.EMAIL] || '').trim();
      if (!email) return { ok: false, error: 'No tenés email registrado. Contactá a soporte.' };

      var pinNuevo = _generarPinNuevo();
      hoja.getRange(i + 1, SCHEMA.RESELLERS.PIN + 1).setValue(_hashPin(pinNuevo));
      invalidateSheetValues(SCHEMA.SHEETS.RESELLERS);

      var html = _authEmailHtml({
        nombre: rowNombre, pin: pinNuevo,
        mensaje: 'Generamos un PIN nuevo para tu acceso al <strong>Portal de Resellers de BIDCOMAGRO</strong> — el anterior queda sin efecto:',
        ctaUrl: portalUrl,
        footer: 'Si no lo pediste vos, contactá a soporte — alguien más intentó entrar con el nombre de tu empresa.'
      });
      GmailApp.sendEmail(email, 'Tu PIN nuevo — Portal Resellers BIDCOMAGRO', '', {
        htmlBody: html,
        name:     'BIDCOMAGRO · Portal Resellers',
        replyTo:  'soporteagrasdji@bidcom.com.ar'
      });
      Logger.log('RS_recordarClave: PIN reseteado → ' + rowNombre + ' <' + email + '>');
      return { ok: true };
    }
    return { ok: false, error: 'Empresa no encontrada. Verificá la selección.' };
  } catch(e) {
    Logger.log('RS_recordarClave: ' + e);
    return { ok: false, error: 'Error al enviar. Intentá de nuevo.' };
  }
}

// ── Corrección de URL: ejecutar UNA VEZ desde el editor ──────
// Envía un email corto a todos los resellers con email + PIN cargado indicando el link
// correcto (/exec en vez de /dev).
function enviarCorreccionUrl() {
  var datos     = getSheetValues(SCHEMA.SHEETS.RESELLERS);
  var portalUrl = 'https://script.google.com/a/macros/bidcom.com.ar/s/AKfycbwyg2uTFTNjYGxfk1htu8Yk5xaO2cOI5xRpyDKEaeA5_URuP7_GbB3cKcE2C8-QRXCt/exec';
  var enviados  = 0, omitidos = 0, errores = [];

  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
    var email  = String(datos[i][SCHEMA.RESELLERS.EMAIL]  || '').trim();
    var pin    = String(datos[i][SCHEMA.RESELLERS.PIN]    || '').trim();

    if (!nombre || !email || !pin) { omitidos++; continue; }

    var html = _authEmailHtml({
      nombre: nombre,
      mensaje: 'Te enviamos un email anterior con el acceso al portal, pero el link estaba incorrecto. El link correcto es el siguiente:',
      ctaUrl: portalUrl,
      footer: 'Tu empresa y contraseña son los mismos que te enviamos antes. Si no los encontrás, respondé este email.'
    });

    try {
      GmailApp.sendEmail(email, 'Corrección de acceso — Portal Resellers BIDCOMAGRO', '', {
        htmlBody: html,
        name:     'BIDCOMAGRO · Portal Resellers',
        replyTo:  'soporteagrasdji@bidcom.com.ar'
      });
      enviados++;
      Logger.log('Corrección enviada → ' + nombre + ' <' + email + '>');
    } catch(e) {
      errores.push(nombre + ' (' + email + '): ' + e.message);
    }
  }

  var resumen = 'Enviados: ' + enviados + ' | Omitidos: ' + omitidos;
  if (errores.length) resumen += ' | Errores: ' + errores.join(' / ');
  Logger.log('enviarCorreccionUrl → ' + resumen);
  return resumen;
}
