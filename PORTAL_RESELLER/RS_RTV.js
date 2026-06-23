// ============================================================
//  PORTAL RESELLER BIDCOM — Vista RTV (solo lectura)
// ============================================================

// Detecta si el usuario logueado es un RTV verificando su email
// contra columna EMAIL_RTV de la hoja Resellers.
// Retorna { ok, resellers: [nombres asignados] }
function obtenerDatosRTV() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail(); } catch(e) {}
    if (!email) return { ok: false };

    var datos   = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var RS      = SCHEMA.RESELLERS;
    var emailLow = email.toLowerCase().trim();
    var resellers = [];

    for (var i = 1; i < datos.length; i++) {
      var rtvEmail = String(datos[i][RS.EMAIL_RTV] || '').trim().toLowerCase();
      if (rtvEmail !== emailLow) continue;
      var nombre = String(datos[i][RS.NOMBRE] || '').trim();
      if (nombre) resellers.push(nombre);
    }

    if (!resellers.length) return { ok: false };
    return { ok: true, email: email, resellers: resellers };
  } catch(e) {
    Logger.log('obtenerDatosRTV: ' + e);
    return { ok: false };
  }
}

// Retorna los pedidos de los resellers asignados al RTV llamante.
// La verificación se hace server-side: no se acepta la lista del cliente,
// se re-calcula desde el email de la sesión activa.
function obtenerPedidosRTV() {
  try {
    var email = '';
    try { email = Session.getActiveUser().getEmail().toLowerCase().trim(); } catch(e) {}
    if (!email) return { ok: false, error: 'Sin sesión', pedidos: [] };

    // Determinar resellers autorizados para este RTV
    var datos = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var RS    = SCHEMA.RESELLERS;
    var autorizado = {};
    for (var i = 1; i < datos.length; i++) {
      var rtvEmail = String(datos[i][RS.EMAIL_RTV] || '').trim().toLowerCase();
      if (rtvEmail !== email) continue;
      var nombre = String(datos[i][RS.NOMBRE] || '').trim();
      if (nombre) autorizado[nombre.toLowerCase()] = nombre;
    }
    if (!Object.keys(autorizado).length) return { ok: false, error: 'Sin resellers asignados', pedidos: [] };

    var dPed  = getSheetValues(SCHEMA.SHEETS.PEDIDOS_REPUESTOS);
    var P     = SCHEMA.PEDIDOS_REPUESTOS;
    var pedidos = [];

    for (var j = dPed.length - 1; j >= 1; j--) {
      var reseller = String(dPed[j][P.RESELLER] || '').trim();
      if (!autorizado[reseller.toLowerCase()]) continue;

      var fecha = dPed[j][P.FECHA];
      var fechaStr = (fecha instanceof Date)
        ? Utilities.formatDate(fecha, 'GMT-3', 'dd/MM/yyyy HH:mm')
        : String(fecha || '');

      var itemsJson = String(dPed[j][P.ITEMS_JSON] || '[]');
      var items = [];
      try { items = JSON.parse(itemsJson); } catch(e) {}

      pedidos.push({
        id:        String(dPed[j][P.ID]           || ''),
        fecha:     fechaStr,
        reseller:  reseller,
        cantItems: Number(dPed[j][P.CANT_ITEMS]   || 0),
        estado:    String(dPed[j][P.ESTADO]        || ''),
        total:     Number(dPed[j][P.TOTAL_USD]     || 0),
        obs:       String(dPed[j][P.OBSERVACIONES] || ''),
        pdfUrl:    String(dPed[j][P.PDF_URL]       || ''),
        formaPago: String(dPed[j][P.FORMA_PAGO]    || ''),
        envio:     String(dPed[j][P.ENVIO]         || ''),
        items:     items
      });

      if (pedidos.length >= 200) break;
    }

    return { ok: true, pedidos: pedidos };
  } catch(e) {
    Logger.log('obtenerPedidosRTV: ' + e);
    return { ok: false, error: e.toString(), pedidos: [] };
  }
}
