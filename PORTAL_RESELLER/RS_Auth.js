// ============================================================
//  PORTAL RESELLER BIDCOM — Autenticación y acceso
// ============================================================

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

function autenticarReseller(pin) {
  try {
    var d       = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var pinHash = _hashPin(String(pin).trim());
    for (var i = 1; i < d.length; i++) {
      var nombre  = String(d[i][0]  || "").trim();
      var pinHoja = String(d[i][9]  || "").trim();
      if (!nombre || !pinHoja) continue;
      if (pinHoja === pinHash) return { ok: true, nombre: nombre };
    }
    return { ok: false };
  } catch(e) { Logger.log("autenticarReseller: " + e); return { ok: false }; }
}

// Valida email + clave (PIN) del portal. Retorna { autorizado, nombre, aftersales }.
function validarAccesoReseller(mail, clave) {
  try {
    var datos  = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    if (!datos || datos.length < 2) return { autorizado: false };

    var mailB  = String(mail  || '').trim().toLowerCase();
    var claveB = String(clave || '').trim();
    if (!mailB || !claveB) return { autorizado: false };

    var header = datos[0].map(function(c){ return String(c || '').toLowerCase().trim(); });
    var colAft = 11;
    for (var j = 0; j < header.length; j++) {
      if (header[j].indexOf('after') !== -1) { colAft = j; break; }
    }

    for (var i = 1; i < datos.length; i++) {
      var rowMail  = String(datos[i][SCHEMA.RESELLERS.EMAIL] || '').trim().toLowerCase();
      var rowClave = String(datos[i][10] || '').trim();
      var nombre   = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
      if (!rowMail || !rowClave || !nombre) continue;
      if (rowMail !== mailB || rowClave !== claveB) continue;

      var rawAft  = datos[i][colAft];
      var aftsStr = String(rawAft === null || rawAft === undefined ? '' : rawAft).trim().toUpperCase();
      var aftersales = (rawAft === true || aftsStr === 'SI' || aftsStr === 'SÍ' || aftsStr === 'S'
                        || aftsStr === 'YES' || aftsStr === '1' || aftsStr === 'TRUE');
      return { autorizado: true, nombre: nombre, aftersales: aftersales };
    }
    return { autorizado: false };
  } catch(e) {
    Logger.log("validarAccesoReseller: " + e);
    return { autorizado: false };
  }
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
    var colAftersales = SCHEMA.RESELLERS.AFTERSALES;
    for (var j = 0; j < header.length; j++) {
      if (header[j].indexOf('after') !== -1) { colAftersales = j; break; }
    }

    var emailLow = email.toLowerCase().trim();
    for (var i = 1; i < datos.length; i++) {
      var rowEmail = String(datos[i][SCHEMA.RESELLERS.EMAIL] || '').toLowerCase().trim();
      if (!rowEmail || rowEmail !== emailLow) continue;

      var nombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
      var raw    = String(datos[i][colAftersales] || '').trim().toUpperCase();
      var aftersales = (raw === 'S' || raw === 'SI' || raw === 'SÍ' || raw === 'YES' || raw === '1');
      return { ok: true, nombre: nombre, aftersales: aftersales };
    }
    return { ok: false, nombre: null, aftersales: false };
  } catch(e) {
    Logger.log("obtenerDatosReseller: " + e);
    return { ok: false, nombre: null, aftersales: false };
  }
}

// Login por nombre de empresa (Col A) + clave (Col K). Retorna { ok, nombre, aftersales }.
// Si el nombre tiene prefijo "★ " es un grupo: valida contra PIN_GRUPO (Col P) y devuelve esGrupo:true.
function validarAccesoInicial(nombre, clave) {
  try {
    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    if (!datos || datos.length < 2) return { ok: false, motivo: 'sin_datos' };

    var nombreRaw = String(nombre || '').trim();
    var claveB    = String(clave  || '').trim();
    if (!nombreRaw || !claveB) return { ok: false, motivo: 'campos_vacios' };

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
          if (rowPinGrupo !== claveB) return { ok: false, motivo: 'clave_incorrecta' };
          pinGrupoEncontrado = true;
        }
        if (rowNombreG) resellersDelGrupo.push(rowNombreG);
      }
      if (!pinGrupoEncontrado) return { ok: false, motivo: 'grupo_no_encontrado' };
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
    var colAft = 11;
    for (var j = 0; j < header.length; j++) {
      if (header[j].indexOf('after') !== -1) { colAft = j; break; }
    }

    for (var i = 1; i < datos.length; i++) {
      var rowNombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase();
      var rowClave  = String(datos[i][10] || '').trim();
      if (!rowNombre || !rowClave) continue;
      if (rowNombre !== nombreB) continue;
      if (rowClave  !== claveB) return { ok: false, motivo: 'clave_incorrecta' };

      var rawAft  = datos[i][colAft];
      var aftsStr = String(rawAft === null || rawAft === undefined ? '' : rawAft).trim().toUpperCase();
      var aftersales = (rawAft === true || aftsStr === 'SI' || aftsStr === 'SÍ' || aftsStr === 'S'
                        || aftsStr === 'YES' || aftsStr === '1' || aftsStr === 'TRUE');
      return {
        ok:         true,
        nombre:     String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim(),
        esGrupo:    false,
        aftersales: aftersales
      };
    }
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

// ── Envío masivo de credenciales ──────────────────────────────
// Ejecutar UNA VEZ desde el editor de Apps Script.
// Manda email a cada reseller que tenga nombre + email + PIN.
// Retorna resumen: { enviados, omitidos, errores }
function enviarCredencialesResellers() {
  var datos     = getSheetValues(SCHEMA.SHEETS.RESELLERS);
  var portalUrl = 'https://script.google.com/a/macros/bidcom.com.ar/s/AKfycbwyg2uTFTNjYGxfk1htu8Yk5xaO2cOI5xRpyDKEaeA5_URuP7_GbB3cKcE2C8-QRXCt/exec';
  var enviados  = 0, omitidos = 0, errores = [];

  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE]   || '').trim();
    var email  = String(datos[i][SCHEMA.RESELLERS.EMAIL]    || '').trim();
    var pin    = String(datos[i][10]                        || '').trim();

    if (!nombre || !email || !pin) { omitidos++; continue; }

    var html =
      "<div style='font-family:sans-serif;max-width:520px;margin:0 auto'>" +
        "<div style='background:#00a3e0;padding:20px 24px;border-radius:10px 10px 0 0'>" +
          "<img src='https://bidcomagro.com.ar/logo.png' height='32' style='vertical-align:middle;margin-right:10px' onerror='this.style.display=\"none\"'>" +
          "<span style='color:#fff;font-size:18px;font-weight:700;vertical-align:middle'>Portal Resellers · BIDCOMAGRO</span>" +
        "</div>" +
        "<div style='background:#fff;border:1px solid #dde3ea;border-top:none;padding:28px 24px;border-radius:0 0 10px 10px'>" +
          "<p style='font-size:14px;color:#444;margin:0 0 18px'>Hola <strong>" + nombre + "</strong>,</p>" +
          "<p style='font-size:13px;color:#555;margin:0 0 22px'>Ya podés acceder al <strong>Portal de Resellers de BIDCOMAGRO</strong> con las siguientes credenciales:</p>" +
          "<div style='background:#f5f8fc;border:1px solid #dde3ea;border-radius:8px;padding:18px 22px;margin-bottom:24px'>" +
            "<div style='margin-bottom:10px'>" +
              "<span style='font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.06em'>Empresa</span>" +
              "<div style='font-size:15px;font-weight:700;color:#1a1a2e;margin-top:3px'>" + nombre + "</div>" +
            "</div>" +
            "<div>" +
              "<span style='font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.06em'>Contraseña</span>" +
              "<div style='font-size:26px;font-weight:800;color:#00a3e0;letter-spacing:6px;margin-top:3px'>" + pin + "</div>" +
            "</div>" +
          "</div>" +
          "<a href='" + portalUrl + "' style='display:inline-block;background:#00a3e0;color:#fff;text-decoration:none;padding:11px 28px;border-radius:7px;font-weight:700;font-size:14px'>Ingresar al portal →</a>" +
          "<p style='font-size:11px;color:#aaa;margin-top:24px;margin-bottom:0'>Si tenés algún problema para acceder, respondé este email y te ayudamos.</p>" +
        "</div>" +
      "</div>";

    try {
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

  var resumen = 'Enviados: ' + enviados + ' | Omitidos (sin email/PIN): ' + omitidos;
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
function RS_recordarClave(nombre) {
  try {
    var nombreB = String(nombre || '').trim().toLowerCase();
    if (!nombreB) return { ok: false, error: 'Seleccioná tu empresa primero.' };

    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var portalUrl = '';
    try { portalUrl = ScriptApp.getService().getUrl(); } catch(eu) {}

    for (var i = 1; i < datos.length; i++) {
      var rowNombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
      if (rowNombre.toLowerCase() !== nombreB) continue;

      var email = String(datos[i][SCHEMA.RESELLERS.EMAIL] || '').trim();
      var pin   = String(datos[i][10] || '').trim();

      if (!email) return { ok: false, error: 'No tenés email registrado. Contactá a soporte.' };
      if (!pin)   return { ok: false, error: 'No tenés PIN registrado. Contactá a soporte.' };

      var html =
        "<div style='font-family:sans-serif;max-width:520px;margin:0 auto'>" +
          "<div style='background:#00a3e0;padding:20px 24px;border-radius:10px 10px 0 0'>" +
            "<span style='color:#fff;font-size:18px;font-weight:700;vertical-align:middle'>Portal Resellers \xb7 BIDCOMAGRO</span>" +
          "</div>" +
          "<div style='background:#fff;border:1px solid #dde3ea;border-top:none;padding:28px 24px;border-radius:0 0 10px 10px'>" +
            "<p style='font-size:14px;color:#444;margin:0 0 18px'>Hola <strong>" + rowNombre + "</strong>,</p>" +
            "<p style='font-size:13px;color:#555;margin:0 0 22px'>Solicitaste un recordatorio de acceso al <strong>Portal de Resellers de BIDCOMAGRO</strong>. Ac\xe1 van tus credenciales:</p>" +
            "<div style='background:#f5f8fc;border:1px solid #dde3ea;border-radius:8px;padding:18px 22px;margin-bottom:24px'>" +
              "<div style='margin-bottom:10px'>" +
                "<span style='font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.06em'>Empresa</span>" +
                "<div style='font-size:15px;font-weight:700;color:#1a1a2e;margin-top:3px'>" + rowNombre + "</div>" +
              "</div>" +
              "<div>" +
                "<span style='font-size:11px;color:#888;text-transform:uppercase;letter-spacing:.06em'>Contrase\xf1a</span>" +
                "<div style='font-size:26px;font-weight:800;color:#00a3e0;letter-spacing:6px;margin-top:3px'>" + pin + "</div>" +
              "</div>" +
            "</div>" +
            (portalUrl ? "<a href='" + portalUrl + "' style='display:inline-block;background:#00a3e0;color:#fff;text-decoration:none;padding:11px 28px;border-radius:7px;font-weight:700;font-size:14px'>Ingresar al portal →</a>" : '') +
            "<p style='font-size:11px;color:#aaa;margin-top:24px;margin-bottom:0'>Si no solicitaste este recordatorio, ignor\xe1 este email. Si ten\xe9s problemas para acceder, respond\xe9 este correo.</p>" +
          "</div>" +
        "</div>";

      GmailApp.sendEmail(email, 'Recordatorio de acceso — Portal Resellers BIDCOMAGRO', '', {
        htmlBody: html,
        name:     'BIDCOMAGRO \xb7 Portal Resellers',
        replyTo:  'soporteagrasdji@bidcom.com.ar'
      });
      Logger.log('RS_recordarClave: enviado → ' + rowNombre + ' <' + email + '>');
      return { ok: true };
    }
    return { ok: false, error: 'Empresa no encontrada. Verific\xe1 la selecci\xf3n.' };
  } catch(e) {
    Logger.log('RS_recordarClave: ' + e);
    return { ok: false, error: 'Error al enviar. Intent\xe1 de nuevo.' };
  }
}

// ── Corrección de URL: ejecutar UNA VEZ desde el editor ──────
// Envía un email corto a todos los resellers con email + PIN
// indicando el link correcto (/exec en vez de /dev).
function enviarCorreccionUrl() {
  var datos     = getSheetValues(SCHEMA.SHEETS.RESELLERS);
  var portalUrl = 'https://script.google.com/a/macros/bidcom.com.ar/s/AKfycbwyg2uTFTNjYGxfk1htu8Yk5xaO2cOI5xRpyDKEaeA5_URuP7_GbB3cKcE2C8-QRXCt/exec';
  var enviados  = 0, omitidos = 0, errores = [];

  for (var i = 1; i < datos.length; i++) {
    var nombre = String(datos[i][SCHEMA.RESELLERS.NOMBRE] || '').trim();
    var email  = String(datos[i][SCHEMA.RESELLERS.EMAIL]  || '').trim();
    var pin    = String(datos[i][10]                      || '').trim();

    if (!nombre || !email || !pin) { omitidos++; continue; }

    var html =
      "<div style='font-family:sans-serif;max-width:520px;margin:0 auto'>" +
        "<div style='background:#00a3e0;padding:20px 24px;border-radius:10px 10px 0 0'>" +
          "<span style='color:#fff;font-size:18px;font-weight:700'>Portal Resellers · BIDCOMAGRO</span>" +
        "</div>" +
        "<div style='background:#fff;border:1px solid #dde3ea;border-top:none;padding:28px 24px;border-radius:0 0 10px 10px'>" +
          "<p style='font-size:14px;color:#444;margin:0 0 12px'>Hola <strong>" + nombre + "</strong>,</p>" +
          "<p style='font-size:13px;color:#555;margin:0 0 18px'>Te enviamos un email anterior con el acceso al portal, pero el link estaba incorrecto. El link correcto es el siguiente:</p>" +
          "<a href='" + portalUrl + "' style='display:inline-block;background:#00a3e0;color:#fff;text-decoration:none;padding:11px 28px;border-radius:7px;font-weight:700;font-size:14px'>Ingresar al portal →</a>" +
          "<p style='font-size:12px;color:#888;margin-top:20px'>Tu empresa y contraseña son los mismos que te enviamos antes. Si no los encontrás, respondé este email.</p>" +
        "</div>" +
      "</div>";

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
