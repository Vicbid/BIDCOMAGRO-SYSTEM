// ============================================================
// @version 1.0
//  PORTAL RESELLER — Módulo de Registro de Nuevos Resellers
//  Flujo: El reseller selecciona su empresa (ya existente en la hoja)
//         → completa los campos que faltan →
//         Email admin con links 1-click →
//         Aprobación: campos escritos en fila existente + PIN generado →
//         Email de bienvenida con credenciales
// ============================================================

var _REG_HOJA = 'SOLICITUDES_RESELLER';

var _REG_COL = {
  FECHA:     0,  // A
  EMPRESA:   1,  // B
  CUIT:      2,  // C  (solo si faltaba en el sheet)
  EMAIL:     3,  // D  (solo si faltaba en el sheet)
  TELEFONO:  4,  // E  (solo si faltaba en el sheet)
  ESTADO:    5,  // F  Pendiente / Aprobado / Rechazado
  TOKEN:     6,  // G  Token único por solicitud (SHA-256, un solo uso)
  FECHA_RES: 7,  // H  Fecha de resolución
  NOTAS:     8,  // I  Mensaje del solicitante
  DIRECCION: 9,  // J  (solo si faltaba en el sheet)
  CP:        10, // K  (solo si faltaba en el sheet)
  LOCALIDAD: 11, // L  (solo si faltaba en el sheet)
  PROVINCIA: 12  // M  (solo si faltaba en el sheet)
};

// ── Asegura que la hoja existe con encabezados ────────────────
function _REG_asegurarHoja() {
  var hoja = getSheet(_REG_HOJA);
  if (!hoja) {
    hoja = getDb().insertSheet(_REG_HOJA);
    var hdrs = ['Fecha','Empresa','CUIT','Email','Teléfono','Estado','Token','Fecha resolución','Notas','Dirección','CP','Localidad','Provincia'];
    hoja.appendRow(hdrs);
    hoja.getRange(1, 1, 1, hdrs.length).setFontWeight('bold').setBackground('#f0f0f0');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// ── Token único por solicitud (SHA-256 truncado a 32 chars) ──
function _REG_generarToken(empresa, ts) {
  var secret = PropertiesService.getScriptProperties().getProperty('APPROVAL_SECRET') || 'bidcomagro-reg-secret';
  var input  = String(empresa) + '|' + String(ts) + '|reg|' + secret;
  var bytes  = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  var hex    = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i] < 0 ? bytes[i] + 256 : bytes[i];
    hex += ('0' + b.toString(16)).slice(-2);
  }
  return hex.substring(0, 32);
}

// ── PIN de 4 dígitos único entre los Resellers existentes ────
function _REG_generarPin() {
  var d      = getSheetValues(SCHEMA.SHEETS.RESELLERS);
  var usados = {};
  for (var i = 1; i < d.length; i++) {
    var p = String(d[i][SCHEMA.RESELLERS.PIN] || '').trim();
    if (p) usados[p] = true;
  }
  var pin, intentos = 0;
  do {
    pin = String(Math.floor(1000 + Math.random() * 9000));
    intentos++;
  } while (usados[pin] && intentos < 200);
  return pin;
}

// ── Qué campos le faltan a esta empresa ──────────────────────
// Retorna: { ok, empresa, campos: { email, telefono, cuit }, error }
// campos.X = true  → campo vacío, debe completarse
// campos.X = false → campo ya tiene dato, no se muestra
function REG_obtenerCamposVacios(empresaNombre) {
  try {
    var empNorm = String(empresaNombre || '').trim().toLowerCase();
    if (!empNorm) return { ok: false, error: 'Seleccioná tu empresa de la lista.' };

    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var RS    = SCHEMA.RESELLERS;

    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][RS.NOMBRE] || '').trim().toLowerCase() !== empNorm) continue;

      var pin       = String(datos[i][RS.PIN]       || '').trim();
      var email     = String(datos[i][RS.EMAIL]     || '').trim();
      var telefono  = String(datos[i][RS.TELEFONO]  || '').trim();
      var cuit      = String(datos[i][RS.CUIT]      || '').trim();
      var direccion = String(datos[i][RS.DIRECCION] || '').trim();
      var cp        = String(datos[i][RS.CP]        || '').trim();
      var localidad = String(datos[i][RS.LOCALIDAD] || '').trim();
      var provincia = String(datos[i][RS.PROVINCIA] || '').trim();

      // Si ya tiene PIN puede ingresar directamente
      if (pin) return { ok: false, error: 'Esta empresa ya tiene acceso al portal. Ingresá con tu empresa y PIN.' };

      // Verificar solicitud pendiente
      var hoja = _REG_asegurarHoja();
      var sol  = hoja.getDataRange().getValues();
      for (var si = 1; si < sol.length; si++) {
        if (String(sol[si][_REG_COL.EMPRESA] || '').trim().toLowerCase() === empNorm &&
            String(sol[si][_REG_COL.ESTADO])                             === 'Pendiente') {
          return { ok: false, error: 'Ya hay una solicitud pendiente para esta empresa. Revisá tu email.' };
        }
      }

      return {
        ok:      true,
        empresa: String(datos[i][RS.NOMBRE] || '').trim(),
        campos: {
          email:     !email,
          telefono:  !telefono,
          cuit:      !cuit,
          direccion: !direccion,
          cp:        !cp,
          localidad: !localidad,
          provincia: !provincia
        }
      };
    }

    return { ok: false, error: 'Empresa no encontrada. Contactate con BIDCOMAGRO.' };
  } catch(e) {
    Logger.log('REG_obtenerCamposVacios: ' + e);
    return { ok: false, error: 'Error interno.' };
  }
}

// ── Punto de entrada público — llamado desde el cliente ───────
// params: { empresa, cuit?, email?, telefono?, notas?, direccion?, cp?, localidad?, provincia? }
// Solo se envían los campos que estaban vacíos en el sheet.
function REG_solicitarAcceso(params) {
  try {
    var empresa   = String(params.empresa   || '').trim();
    var cuit      = String(params.cuit      || '').trim();
    var email     = String(params.email     || '').trim().toLowerCase();
    var telefono  = String(params.telefono  || '').trim();
    var notas     = String(params.notas     || '').trim();
    var direccion = String(params.direccion || '').trim();
    var cp        = String(params.cp        || '').trim();
    var localidad = String(params.localidad || '').trim();
    var provincia = String(params.provincia || '').trim();

    if (!empresa) return { ok: false, error: 'Empresa requerida.' };

    // Verificar que la empresa existe y no tiene PIN todavía
    var datos  = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var RS     = SCHEMA.RESELLERS;
    var filaR  = -1;
    for (var ri = 1; ri < datos.length; ri++) {
      if (String(datos[ri][RS.NOMBRE] || '').trim().toLowerCase() === empresa.toLowerCase()) {
        filaR = ri;
        break;
      }
    }
    if (filaR === -1) return { ok: false, error: 'Empresa no encontrada.' };
    if (String(datos[filaR][RS.PIN] || '').trim()) return { ok: false, error: 'Esta empresa ya tiene acceso al portal.' };

    // Verificar solicitud pendiente duplicada
    var hoja = _REG_asegurarHoja();
    var sol  = hoja.getDataRange().getValues();
    for (var si = 1; si < sol.length; si++) {
      if (String(sol[si][_REG_COL.EMPRESA] || '').trim().toLowerCase() === empresa.toLowerCase() &&
          String(sol[si][_REG_COL.ESTADO])                             === 'Pendiente') {
        return { ok: false, error: 'Ya hay una solicitud pendiente para esta empresa.' };
      }
    }

    // Al menos email o algún dato nuevo debe ser aportado
    if (!email && !telefono && !cuit && !direccion && !cp && !localidad && !provincia)
      return { ok: false, error: 'Completá al menos el email de contacto.' };
    if (email && email.indexOf('@') < 1) return { ok: false, error: 'Email inválido.' };

    // Guardar solicitud
    var ahora = new Date();
    var token = _REG_generarToken(empresa, ahora.getTime());
    hoja.appendRow([ahora, empresa, cuit, email, telefono, 'Pendiente', token, '', notas, direccion, cp, localidad, provincia]);

    // Notificar al admin
    _REG_notificarAdmin(empresa, cuit, email, telefono, notas, token, direccion, cp, localidad, provincia);

    return { ok: true };
  } catch(e) {
    Logger.log('REG_solicitarAcceso: ' + e);
    return { ok: false, error: 'Error interno. Intentá de nuevo.' };
  }
}

// ── Email al admin con botones Aprobar / Rechazar 1-click ─────
function _REG_notificarAdmin(empresa, cuit, email, telefono, notas, token, direccion, cp, localidad, provincia) {
  try {
    var base  = ScriptApp.getService().getUrl();
    var urlOk = base + '?action=reg-ok&token=' + token;
    var urlNo = base + '?action=reg-no&token=' + token;

    var tdL = 'style="padding:9px 0;font-size:11px;color:#999;width:120px"';
    var tdR = 'style="padding:9px 0;font-size:13px;color:#222"';
    var sep = 'style="border-bottom:1px solid #f5f5f5"';
    var dirStr = [direccion, (cp ? 'CP ' + cp : ''), localidad, provincia].filter(Boolean).join(' · ');

    var filasTbl =
      '<tr ' + sep + '><td ' + tdL + '>Empresa</td><td ' + tdR + ' style="font-weight:700">' + empresa + '</td></tr>' +
      (cuit    ? '<tr ' + sep + '><td ' + tdL + '>CUIT</td><td ' + tdR + '>'     + cuit    + '</td></tr>' : '') +
      (email   ? '<tr ' + sep + '><td ' + tdL + '>Email</td><td ' + tdR + '>'    + email   + '</td></tr>' : '') +
      (telefono? '<tr ' + sep + '><td ' + tdL + '>Teléfono</td><td ' + tdR + '>' + telefono+ '</td></tr>' : '') +
      (dirStr  ? '<tr ' + sep + '><td ' + tdL + '>Dirección</td><td ' + tdR + '>'+ dirStr  + '</td></tr>' : '') +
      (notas   ? '<tr><td ' + tdL + '>Mensaje</td><td style="padding:9px 0;font-size:13px;color:#555;font-style:italic">' + notas + '</td></tr>' : '');

    var html =
      "<div style='font-family:sans-serif;max-width:560px;margin:0 auto'>" +
        "<div style='background:#009ee3;padding:18px 24px;border-radius:10px 10px 0 0'>" +
          "<span style='color:#fff;font-size:16px;font-weight:700'>Solicitud de acceso al Portal Resellers</span>" +
        "</div>" +
        "<div style='background:#fff;border:1px solid #e8e8e8;border-top:none;padding:24px;border-radius:0 0 10px 10px'>" +
          "<p style='font-size:13px;color:#444;margin:0 0 20px'>Un distribuidor ya registrado en el sistema completó sus datos de contacto y solicita acceso al portal.</p>" +
          "<table style='width:100%;border-collapse:collapse;margin-bottom:24px'>" + filasTbl + "</table>" +
          "<div style='display:flex;gap:12px;margin-bottom:20px'>" +
            "<a href='" + urlOk + "' style='flex:1;display:block;text-align:center;background:#00a650;color:#fff;text-decoration:none;padding:13px 0;border-radius:8px;font-weight:700;font-size:14px'>&#10003; &nbsp;Aprobar y enviar PIN</a>" +
            "<a href='" + urlNo + "' style='flex:1;display:block;text-align:center;background:#f5f5f5;color:#666;text-decoration:none;padding:13px 0;border-radius:8px;font-weight:600;font-size:14px;border:1px solid #e0e0e0'>&#10007; &nbsp;Rechazar</a>" +
          "</div>" +
          "<p style='font-size:11px;color:#bbb;margin:0'>Cada link es de un solo uso. Una vez procesado no puede volver a usarse.</p>" +
        "</div>" +
      "</div>";

    GmailApp.sendEmail(
      PORTAL_CONFIG.EMAIL_SUPERVISOR,
      '[Portal Resellers] Solicitud de acceso: ' + empresa,
      '',
      { htmlBody: html, name: 'Portal Resellers · BIDCOMAGRO', replyTo: email || PORTAL_CONFIG.EMAIL_SUPERVISOR }
    );
  } catch(e) {
    Logger.log('_REG_notificarAdmin: ' + e);
  }
}

// ── Procesa el click del link del admin (aprobación / rechazo) ──
// Llamado desde doGet con action=reg-ok|reg-no y token
function _REG_procesarDecision(token, accion) {
  var tokenB    = String(token  || '').trim();
  var esAprobar = accion === 'reg-ok';

  if (!tokenB) return _REG_paginaResultado('Link inválido', 'El link no contiene información válida.', '#e74c3c', '&#9888;');

  try {
    var hoja  = _REG_asegurarHoja();
    var datos = hoja.getDataRange().getValues();
    var fila  = -1;
    for (var i = 1; i < datos.length; i++) {
      if (String(datos[i][_REG_COL.TOKEN] || '').trim() === tokenB) { fila = i; break; }
    }

    if (fila === -1)
      return _REG_paginaResultado('Link inválido', 'No se encontró la solicitud. Puede que ya haya sido procesada.', '#f39c12', 'ℹ️');

    var estado = String(datos[fila][_REG_COL.ESTADO] || '');
    if (estado !== 'Pendiente')
      return _REG_paginaResultado('Ya procesada', 'Esta solicitud ya fue <strong>' + estado.toLowerCase() + '</strong> anteriormente.', '#f39c12', 'ℹ️');

    var empresa   = String(datos[fila][_REG_COL.EMPRESA]   || '').trim();
    var cuit      = String(datos[fila][_REG_COL.CUIT]      || '').trim();
    var email     = String(datos[fila][_REG_COL.EMAIL]     || '').trim();
    var telefono  = String(datos[fila][_REG_COL.TELEFONO]  || '').trim();
    var direccion = String(datos[fila][_REG_COL.DIRECCION] || '').trim();
    var cp        = String(datos[fila][_REG_COL.CP]        || '').trim();
    var localidad = String(datos[fila][_REG_COL.LOCALIDAD] || '').trim();
    var provincia = String(datos[fila][_REG_COL.PROVINCIA] || '').trim();
    var ahora     = new Date();

    if (esAprobar) {
      // Buscar la fila existente del reseller y actualizar solo campos vacíos
      var hojaRes  = getSheet(SCHEMA.SHEETS.RESELLERS);
      var resRows  = hojaRes.getDataRange().getValues();
      var RS       = SCHEMA.RESELLERS;
      var filaRes  = -1;
      for (var r = 1; r < resRows.length; r++) {
        if (String(resRows[r][RS.NOMBRE] || '').trim().toLowerCase() === empresa.toLowerCase()) {
          filaRes = r; break;
        }
      }

      var pin;
      if (filaRes === -1) {
        // Fallback: crear fila nueva (no debería ocurrir)
        pin = _REG_generarPin();
        hojaRes.appendRow([empresa, cuit, '', '', '', '', telefono, '', '', email, pin]);
      } else {
        // Actualizar solo los campos que estaban vacíos
        pin = String(resRows[filaRes][RS.PIN] || '').trim();
        if (!pin) {
          pin = _REG_generarPin();
          hojaRes.getRange(filaRes + 1, RS.PIN + 1).setValue(pin);
        }
        if (cuit      && !String(resRows[filaRes][RS.CUIT]      || '').trim())
          hojaRes.getRange(filaRes + 1, RS.CUIT      + 1).setValue(cuit);
        if (email     && !String(resRows[filaRes][RS.EMAIL]     || '').trim())
          hojaRes.getRange(filaRes + 1, RS.EMAIL     + 1).setValue(email);
        if (telefono  && !String(resRows[filaRes][RS.TELEFONO]  || '').trim())
          hojaRes.getRange(filaRes + 1, RS.TELEFONO  + 1).setValue(telefono);
        if (direccion && !String(resRows[filaRes][RS.DIRECCION] || '').trim())
          hojaRes.getRange(filaRes + 1, RS.DIRECCION + 1).setValue(direccion);
        if (cp        && !String(resRows[filaRes][RS.CP]        || '').trim())
          hojaRes.getRange(filaRes + 1, RS.CP        + 1).setValue(cp);
        if (localidad && !String(resRows[filaRes][RS.LOCALIDAD] || '').trim())
          hojaRes.getRange(filaRes + 1, RS.LOCALIDAD + 1).setValue(localidad);
        if (provincia && !String(resRows[filaRes][RS.PROVINCIA] || '').trim())
          hojaRes.getRange(filaRes + 1, RS.PROVINCIA + 1).setValue(provincia);
      }

      invalidateSheetValues(SCHEMA.SHEETS.RESELLERS);

      hoja.getRange(fila + 1, _REG_COL.ESTADO    + 1).setValue('Aprobado');
      hoja.getRange(fila + 1, _REG_COL.FECHA_RES + 1).setValue(ahora);

      // Solo mandar bienvenida si tenemos email
      var emailDest = email || String((filaRes !== -1 ? resRows[filaRes][RS.EMAIL] : '') || '').trim();
      if (emailDest) _REG_emailBienvenida(empresa, emailDest, pin);

      return _REG_paginaResultado(
        'Acceso aprobado &#10003;',
        '<strong>' + empresa + '</strong> ya puede acceder al portal.<br>' +
        (emailDest ? 'Enviamos las credenciales a <em>' + emailDest + '</em>.' : 'No hay email registrado — entregá el PIN manualmente.'),
        '#00a650', '&#10003;'
      );
    } else {
      hoja.getRange(fila + 1, _REG_COL.ESTADO    + 1).setValue('Rechazado');
      hoja.getRange(fila + 1, _REG_COL.FECHA_RES + 1).setValue(ahora);

      if (email) _REG_emailRechazo(empresa, email);

      return _REG_paginaResultado(
        'Solicitud rechazada',
        'La solicitud de <strong>' + empresa + '</strong> fue rechazada.' + (email ? ' Notificamos al solicitante.' : ''),
        '#e74c3c', '&#10007;'
      );
    }
  } catch(e) {
    Logger.log('_REG_procesarDecision: ' + e);
    return _REG_paginaResultado('Error', 'Ocurrió un error al procesar la solicitud. Revisá los logs.', '#e74c3c', '&#9888;');
  }
}

// ── Email de bienvenida con credenciales ─────────────────────
function _REG_emailBienvenida(empresa, email, pin) {
  try {
    var url  = ScriptApp.getService().getUrl();
    var html =
      "<div style='font-family:sans-serif;max-width:520px;margin:0 auto'>" +
        "<div style='background:#009ee3;padding:20px 24px;border-radius:10px 10px 0 0'>" +
          "<span style='color:#fff;font-size:18px;font-weight:700'>Portal Resellers · BIDCOMAGRO</span>" +
        "</div>" +
        "<div style='background:#fff;border:1px solid #e8e8e8;border-top:none;padding:28px 24px;border-radius:0 0 10px 10px'>" +
          "<p style='font-size:14px;color:#444;margin:0 0 18px'>Hola <strong>" + empresa + "</strong>, tu solicitud fue aprobada. Ya podés ingresar al portal:</p>" +
          "<div style='background:#f7f9fc;border:1px solid #e8e8e8;border-radius:8px;padding:18px 22px;margin-bottom:24px'>" +
            "<div style='margin-bottom:12px'>" +
              "<div style='font-size:10px;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px'>Empresa</div>" +
              "<div style='font-size:15px;font-weight:700;color:#222'>" + empresa + "</div>" +
            "</div>" +
            "<div>" +
              "<div style='font-size:10px;color:#aaa;text-transform:uppercase;letter-spacing:.06em;margin-bottom:3px'>Contraseña (PIN)</div>" +
              "<div style='font-size:32px;font-weight:800;color:#009ee3;letter-spacing:8px'>" + pin + "</div>" +
            "</div>" +
          "</div>" +
          "<a href='" + url + "' style='display:block;text-align:center;background:#009ee3;color:#fff;text-decoration:none;padding:13px 0;border-radius:8px;font-weight:700;font-size:14px;margin-bottom:20px'>Ingresar al portal &rarr;</a>" +
          "<p style='font-size:11px;color:#bbb;margin:0'>Si tenés dudas escribinos a " + PORTAL_CONFIG.EMAIL_SUPERVISOR + "</p>" +
        "</div>" +
      "</div>";
    GmailApp.sendEmail(email, 'Tu acceso al Portal Resellers — BIDCOMAGRO', '', {
      htmlBody: html, name: 'BIDCOMAGRO · Portal Resellers', replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR
    });
  } catch(e) { Logger.log('_REG_emailBienvenida: ' + e); }
}

// ── Email de rechazo al solicitante ──────────────────────────
function _REG_emailRechazo(empresa, email) {
  try {
    var html =
      "<div style='font-family:sans-serif;max-width:520px;margin:0 auto'>" +
        "<div style='background:#009ee3;padding:20px 24px;border-radius:10px 10px 0 0'>" +
          "<span style='color:#fff;font-size:18px;font-weight:700'>Portal Resellers · BIDCOMAGRO</span>" +
        "</div>" +
        "<div style='background:#fff;border:1px solid #e8e8e8;border-top:none;padding:28px 24px;border-radius:0 0 10px 10px'>" +
          "<p style='font-size:14px;color:#444;margin:0 0 14px'>Hola <strong>" + empresa + "</strong>,</p>" +
          "<p style='font-size:13px;color:#555;margin:0 0 20px;line-height:1.6'>Por el momento no podemos habilitar el acceso al Portal Resellers. Si creés que es un error o querés más información, respondé este email y te contactamos.</p>" +
          "<p style='font-size:11px;color:#bbb;margin:0'>BIDCOMAGRO &middot; Soporte Técnico DJI Agriculture</p>" +
        "</div>" +
      "</div>";
    GmailApp.sendEmail(email, 'Solicitud de acceso — Portal Resellers BIDCOMAGRO', '', {
      htmlBody: html, name: 'BIDCOMAGRO · Portal Resellers', replyTo: PORTAL_CONFIG.EMAIL_SUPERVISOR
    });
  } catch(e) { Logger.log('_REG_emailRechazo: ' + e); }
}

// ── Página HTML de resultado para el admin ────────────────────
function _REG_paginaResultado(titulo, mensaje, color, icono) {
  var html =
    "<!DOCTYPE html><html><head><meta charset='UTF-8'>" +
    "<meta name='viewport' content='width=device-width,initial-scale=1'>" +
    "<title>" + titulo + " — BIDCOMAGRO</title>" +
    "<style>*{box-sizing:border-box}body{margin:0;background:#f2f4f7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}" +
    ".card{background:#fff;border-radius:14px;box-shadow:0 4px 24px rgba(0,0,0,.08);max-width:440px;width:100%;padding:40px 36px;text-align:center}" +
    ".ico{width:56px;height:56px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:24px;font-weight:700;color:#fff;background:" + color + ";margin:0 auto 20px}" +
    "h1{font-size:20px;font-weight:800;margin:0 0 12px;color:#222}" +
    "p{font-size:14px;color:#666;line-height:1.6;margin:0 0 10px}" +
    ".brand{font-size:11px;color:#ccc;margin-top:28px;padding-top:16px;border-top:1px solid #f0f0f0}" +
    "</style></head><body>" +
    "<div class='card'>" +
      "<div class='ico'>" + icono + "</div>" +
      "<h1>" + titulo + "</h1>" +
      "<p>" + mensaje + "</p>" +
      "<p style='font-size:12px;color:#bbb'>Podés cerrar esta ventana.</p>" +
      "<div class='brand'>BIDCOMAGRO &middot; Portal Resellers</div>" +
    "</div></body></html>";
  return HtmlService.createHtmlOutput(html)
    .setTitle(titulo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
