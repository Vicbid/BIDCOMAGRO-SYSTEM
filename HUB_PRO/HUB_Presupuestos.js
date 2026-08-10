// @version 1.1
// ============================================================
//  HUB PRO — Presupuestos: armado del HTML, tokens de aprobación/
//  rechazo por link de mail, envío al reseller.
//  Extraído de HUB_Código.js 2.26 el 2026-07-30 — reorganización
//  sin cambios funcionales (más archivos, mismo comportamiento).
// ============================================================


// Token HMAC-SHA256 para aprobación de presupuesto 1-click.
// El mismo secreto debe estar en las Script Properties de HUB y Portal con clave APPROVAL_SECRET.
function _tokenAprobacion(ot, action) {
  var secret = PropertiesService.getScriptProperties().getProperty('APPROVAL_SECRET') || 'bidcomagro-default';
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(ot) + '|' + String(action) + '|' + secret
  );
  return bytes.map(function(b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('').substring(0, 40);
}



// ============================================================
//  TRIGGER DIARIO — SLA

function _bloquesBotonesPresupuesto(ot) {
  var urlBase   = CONFIG.PORTAL_URL;
  var tokAprob  = _tokenAprobacion(ot, 'aprobar');
  var tokRechaz = _tokenAprobacion(ot, 'rechazar');
  var urlAprob  = urlBase + '?action=aprobar&ot='  + encodeURIComponent(ot) + '&token=' + tokAprob;
  var urlRechaz = urlBase + '?action=rechazar&ot=' + encodeURIComponent(ot) + '&token=' + tokRechaz;
  return "<div style='background:#f9f9f9;border:1px solid #e0e0e0;border-radius:8px;padding:20px 24px;margin-bottom:16px;text-align:center'>" +
    "<p style='font-size:13px;color:#555;margin:0 0 16px;line-height:1.5'>Por favor revisá el detalle y tomá una decisión:</p>" +
    "<a href='" + urlAprob + "' style='display:inline-block;background:#1a9e4a;color:#fff;font-weight:700;font-size:14px;padding:12px 28px;border-radius:6px;text-decoration:none;margin:0 8px'>✅ Aprobar presupuesto</a>" +
    "<a href='" + urlRechaz + "' style='display:inline-block;background:#e74c3c;color:#fff;font-weight:700;font-size:14px;padding:12px 28px;border-radius:6px;text-decoration:none;margin:0 8px'>❌ Rechazar presupuesto</a>" +
    "<p style='font-size:11px;color:#aaa;margin:14px 0 0'>Este link es de uso único y está asociado a tu cuenta. Ante dudas respondé este email.</p>" +
  "</div>";
}


// ============================================================
//  ENVIAR PRESUPUESTO AL RESELLER — manual desde el HUB
// ============================================================
function enviarPresupuesto(data) {
  try {
    var emailReseller = obtenerEmailReseller(data.reseller);
    if (!emailReseller) return { ok: false, msg: "El reseller no tiene email registrado." };

    // Precios repuestos desde DB_REPUESTOS col F
    var precioMap = {};
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      for (var pr = 1; pr < dRep.length; pr++) {
        var codR = String(dRep[pr][1]||"").trim().toUpperCase();
        if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5]||"0").replace(",",".")) || 0;
      }
    }

    // Tabla repuestos con precios
    var tablaRepHTML = "";
    var totalRep = 0;
    if (data.repuestos && data.repuestos !== "Sin consumo de repuestos") {
      var ls = data.repuestos.split(" ; ");
      tablaRepHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border:1px solid #e0e0e0'>Cant.</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>P.Unit (USD)</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Subtotal (USD)</th>" +
        "</tr></thead><tbody>";
      for (var ri = 0; ri < ls.length; ri++) {
        var p = ls[ri].split(" | ");
        if (p.length < 3) continue;
        var cod     = p[0].trim();
        var desc    = p[1].trim();
        var ped     = parseInt((p[2].split(" E:")[0] || "").replace("P:","")) || 0;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var descPct = (data.tipoDestinatario === "cliente") ? 0.25 : 0.40;
        var pvp     = pvpBase > 0 ? pvpBase * (1 - descPct) : 0;
        var sub     = pvp * ped;
        totalRep   += sub;
        var labelDesc = (data.tipoDestinatario === "cliente") ? "25% desc." : "40% desc.";
        tablaRepHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + cod + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + desc + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:center'>" + ped + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + (pvpBase ? "USD " + _fmtNum(pvpBase) + " <span style='color:#888;font-size:10px'>(-" + labelDesc + ")</span>" : "—") + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right;font-weight:700'>" + (pvp ? "USD " + _fmtNum(pvp) : "—") + "</td>" +
          "</tr>";
      }
      if (totalRep > 0) {
        tablaRepHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='4' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL REPUESTOS</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalRep) + "</td></tr>";
      }
      tablaRepHTML += "</tbody></table>";
    }

    // Tabla mano de obra
    var tablaMOHTML = "";
    var totalMO = 0;
    if (data.manoObra && data.manoObra.length) {
      tablaMOHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Precio</th>" +
        "</tr></thead><tbody>";
      for (var mi = 0; mi < data.manoObra.length; mi++) {
        var mo = data.manoObra[mi];
        var pMO = parseFloat(String(mo.precio).replace(/[^0-9.]/g,"")) || 0;
        totalMO += pMO;
        tablaMOHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + mo.codigo + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + mo.descripcion + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + mo.precio + "</td></tr>";
      }
      if (totalMO > 0) {
        tablaMOHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='2' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL MANO DE OBRA</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalMO) + "</td></tr>";
      }
      tablaMOHTML += "</tbody></table>";
    }

    var cuerpoDetalle =
      filaDetalle("Orden de Trabajo", data.ot) +
      filaDetalle("Equipo / Modelo",  data.equipo) +
      filaDetalle("Nº de Serie",      data.sn || "—") +
      (data.cas ? filaDetalle("Caso DJI (CAS/FWR)", data.cas) : "") +
      filaDetalle("Garantía", "Fuera de garantía (OOW)");

    var bloques =
      bloqueCard("📋 Detalle de la Orden", cuerpoDetalle, "#00a3e0") +
      (data.trabajo ? bloqueCard("📝 Falla reportada",
        "<p style='margin:0;font-size:13px;line-height:1.6;color:#555'>" + data.trabajo.replace(/\n/g,"<br>") + "</p>",
        "#6b7280") : "") +
      (data.informeTecnico ? bloqueCard("🔧 Diagnóstico Técnico",
        "<p style='margin:0;font-size:13px;line-height:1.6;color:#555'>" + data.informeTecnico.replace(/\n/g,"<br>") + "</p>",
        "#27ae60") : "") +
      (tablaRepHTML ? "<div style='margin-bottom:16px'><p style='font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase'>Repuestos</p>" + tablaRepHTML + "</div>" : "") +
      (tablaMOHTML  ? "<div style='margin-bottom:16px'><p style='font-size:12px;font-weight:700;color:#333;margin-bottom:8px;text-transform:uppercase'>Mano de Obra</p>" + tablaMOHTML + "</div>" : "") +
      _bloquesBotonesPresupuesto(data.ot);

    var htmlEmail = construirEmailHTML(
      "Presupuesto de Reparación — " + data.ot,
      "Estimado equipo de <strong>" + data.reseller + "</strong>,<br>Le enviamos el presupuesto para la siguiente orden:",
      bloques,
      "Ante cualquier consulta no dude en contactarnos."
    );

    var asunto = "OT " + data.ot + " — " + data.equipo;
    var tidPres = _enviarConHilo(data.ot, emailReseller, asunto, htmlEmail);
    registrarEmailLog(data.ot, emailReseller, "Reseller", asunto, "OK", tidPres || "");
    var tidCopia = _enviarConHilo(data.ot, CONFIG.EMAIL_SUPERVISOR, "[COPIA] " + asunto, htmlEmail);
    registrarEmailLog(data.ot, CONFIG.EMAIL_SUPERVISOR, "Supervisor", "[COPIA] " + asunto, "OK", tidCopia || "");

    // Actualizar estado a "Presupuesto enviado" automáticamente
    try {
      var hoja = getSheet(SCHEMA.SHEETS.OT);
      var todos = getSheetValues(hoja);
      for (var ri = 1; ri < todos.length; ri++) {
        if (String(todos[ri][SCHEMA.OT.OT] || "").trim() === String(data.ot).trim()) {
          var fila = ri + 1;
          hoja.getRange(fila, SCHEMA.OT.ESTADO + 1).setValue("Presupuesto enviado");
          hoja.getRange(fila, SCHEMA.OT.FECHA_ESTADO + 1).setValue(new Date());
          hoja.getRange(fila, SCHEMA.OT.ULTIMA_MODIFICACION + 1).setValue(new Date());
          invalidateSheetValues(SCHEMA.SHEETS.OT);
          registrarLog(data.ot, "", Session.getActiveUser().getEmail(), "ACTUALIZACIÓN", "En Revision", "Presupuesto enviado", "Presupuesto enviado via botón");
          break;
        }
      }
    } catch(eUpd) { Logger.log("enviarPresupuesto: no se pudo actualizar estado: " + eUpd); }

    return { ok: true, emailEnviado: emailReseller };

  } catch(e) {
    Logger.log("enviarPresupuesto ERROR: " + e);
    return { ok: false, msg: e.toString() };
  }
}


function _getPresupNum() {
  var props = PropertiesService.getScriptProperties();
  var n = parseInt(props.getProperty('PRESUP_COUNTER') || '1000') + 1;
  props.setProperty('PRESUP_COUNTER', String(n));
  return '0002-' + (Array(8).join('0') + n).slice(-7);
}


// Devuelve el HTML del presupuesto sin enviarlo (para vista previa / descarga)
function obtenerPresupuestoHTML(data) {
  try {
    // Reseller data
    var rsData = getSheetValues(SCHEMA.SHEETS.RESELLERS);
    var rLow = String(data.reseller || '').trim().toLowerCase();
    var rCuit = '', rDir = '', rCp = '', rLoc = '', rTel = '';
    for (var ri = 1; ri < rsData.length; ri++) {
      if (String(rsData[ri][SCHEMA.RESELLERS.NOMBRE] || '').trim().toLowerCase() === rLow) {
        rCuit = String(rsData[ri][SCHEMA.RESELLERS.CUIT]      || '').trim();
        rDir  = String(rsData[ri][SCHEMA.RESELLERS.DIRECCION] || '').trim();
        rCp   = String(rsData[ri][SCHEMA.RESELLERS.CP]        || '').trim();
        rLoc  = String(rsData[ri][SCHEMA.RESELLERS.LOCALIDAD] || '').trim();
        rTel  = String(rsData[ri][SCHEMA.RESELLERS.TELEFONO]  || '').trim();
        break;
      }
    }

    // Presupuesto number
    var presupNum = data.presupNum || (data.modoPreview ? 'BORRADOR' : _getPresupNum());

    // Prices from DB_REPUESTOS
    var precioMap = {};
    var hojaRep = getSheet(SCHEMA.SHEETS.DB_REPUESTOS);
    if (hojaRep) {
      var dRep = getSheetValues(hojaRep);
      for (var pr = 1; pr < dRep.length; pr++) {
        var codR = String(dRep[pr][1]||"").trim().toUpperCase();
        if (codR) precioMap[codR] = parseFloat(String(dRep[pr][5]||"0").replace(",",".")) || 0;
      }
    }

    var tablaRepHTML = "", totalRep = 0;
    if (data.repuestos && data.repuestos !== "Sin consumo de repuestos") {
      var ls = data.repuestos.split(" ; ");
      tablaRepHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:center;border:1px solid #e0e0e0'>Cant.</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>P.Unit (USD)</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Subtotal (USD)</th>" +
        "</tr></thead><tbody>";
      for (var ri = 0; ri < ls.length; ri++) {
        var p = ls[ri].split(" | ");
        if (p.length < 3) continue;
        var cod  = p[0].trim();
        var desc = p[1].trim();
        var ped  = parseInt((p[2].split(" E:")[0]||"").replace("P:","")) || 0;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var descPct = (data.tipoDestinatario === "cliente") ? 0.25 : 0.40;
        var pvp  = pvpBase > 0 ? pvpBase * (1 - descPct) : 0;
        var sub  = pvp * ped;
        totalRep += sub;
        var labelDesc = (data.tipoDestinatario === "cliente") ? "25% desc." : "40% desc.";
        tablaRepHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + cod + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + desc + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:center'>" + ped + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + (pvpBase ? "USD " + _fmtNum(pvpBase) + " <span style='color:#888;font-size:10px'>(-" + labelDesc + ")</span>" : "—") + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right;font-weight:700'>" + (pvp ? "USD " + _fmtNum(pvp) : "—") + "</td></tr>";
      }
      if (totalRep > 0) {
        tablaRepHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='4' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL REPUESTOS</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalRep) + "</td></tr>";
      }
      tablaRepHTML += "</tbody></table>";
    }

    var tablaMOHTML = "", totalMO = 0;
    if (data.manoObra && data.manoObra.length) {
      tablaMOHTML = "<table style='width:100%;border-collapse:collapse;font-size:12px;margin-bottom:12px'>" +
        "<thead><tr style='background:#f5f5f5'>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Código</th>" +
        "<th style='padding:7px 10px;text-align:left;border:1px solid #e0e0e0'>Descripción</th>" +
        "<th style='padding:7px 10px;text-align:right;border:1px solid #e0e0e0'>Precio</th>" +
        "</tr></thead><tbody>";
      for (var mi = 0; mi < data.manoObra.length; mi++) {
        var mo = data.manoObra[mi];
        var pMO = parseFloat(String(mo.precio).replace(/[^0-9.]/g,"")) || 0;
        totalMO += pMO;
        tablaMOHTML += "<tr>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;font-weight:600;color:#00a3e0'>" + mo.codigo + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0'>" + mo.descripcion + "</td>" +
          "<td style='padding:6px 10px;border:1px solid #e0e0e0;text-align:right'>" + mo.precio + "</td></tr>";
      }
      if (totalMO > 0) {
        tablaMOHTML += "<tr style='background:#f5f5f5;font-weight:700'>" +
          "<td colspan='2' style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>SUBTOTAL MANO DE OBRA</td>" +
          "<td style='padding:7px 10px;border:1px solid #e0e0e0;text-align:right'>USD " + _fmtNum(totalMO) + "</td></tr>";
      }
      tablaMOHTML += "</tbody></table>";
    }

    var totalGeneral = totalRep + totalMO;
    var totalFila = totalGeneral > 0
      ? "<table style='width:100%;border-collapse:collapse;margin-top:8px'><tr style='background:#00a3e0;color:#fff;font-weight:700;font-size:14px'>" +
        "<td style='padding:10px 14px;text-align:right'>TOTAL GENERAL</td>" +
        "<td style='padding:10px 14px;text-align:right;white-space:nowrap'>USD " + _fmtNum(totalGeneral) + "</td></tr></table>"
      : "";

    var cuerpoDetalle =
      filaDetalle("Orden de Trabajo", data.ot) +
      filaDetalle("Equipo / Modelo",  data.equipo) +
      filaDetalle("Nº de Serie",      data.sn || "—") +
      (data.cas ? filaDetalle("Caso DJI", data.cas) : "") +
      filaDetalle("Reseller",         data.reseller) +
      filaDetalle("Garantía",         "Fuera de garantía (OOW)") +
      filaDetalle("Fecha",            Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"));

    // Build items table rows
    var descPct   = (data.tipoDestinatario === 'cliente') ? 0.25 : 0.40;
    var descLabel = (data.tipoDestinatario === 'cliente') ? '25,00%' : '40,00%';
    var tablaFilas = '', totalGeneral = 0;

    if (data.repuestos && data.repuestos !== 'Sin consumo de repuestos') {
      var ls = data.repuestos.split(' ; ');
      for (var li = 0; li < ls.length; li++) {
        var p = ls[li].split(' | ');
        if (p.length < 3) continue;
        var cod  = p[0].trim();
        var desc = p[1].trim();
        var ped  = parseInt((p[2].split(' E:')[0]||'').replace('P:','')) || 0;
        var pvpBase = precioMap[cod.toUpperCase()] || 0;
        var pvpNeto = pvpBase > 0 ? pvpBase * (1 - descPct) : 0;
        var sub  = pvpNeto * ped;
        totalGeneral += sub;
        tablaFilas +=
          '<tr>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px;color:#00a3e0;font-weight:700">' + cod + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px">' + desc + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px">' + (pvpBase ? _fmtNum(pvpBase) : '—') + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:center;font-size:10px">' + descLabel + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px">' + (pvpNeto ? _fmtNum(pvpNeto) : '—') + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:center;font-size:10px">' + ped + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px;font-weight:700">' + (sub ? _fmtNum(sub) : '—') + '</td>' +
          '</tr>';
      }
    }

    if (data.manoObra && data.manoObra.length) {
      for (var mi2 = 0; mi2 < data.manoObra.length; mi2++) {
        var mo = data.manoObra[mi2];
        var pMO = parseFloat(String(mo.precio).replace(/[^0-9.,]/g,'').replace(',','.')) || 0;
        totalGeneral += pMO;
        tablaFilas +=
          '<tr style="background:#f8fff8">' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px;color:#27ae60;font-weight:700">' + mo.codigo + '</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;font-size:10px">' + mo.descripcion + '</td>' +
          '<td colspan="3" style="padding:4px 7px;border:1px solid #ccc;font-size:9px;color:#888;text-align:center">Revisión y mano de obra</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:center;font-size:10px">1</td>' +
          '<td style="padding:4px 7px;border:1px solid #ccc;text-align:right;font-size:10px;font-weight:700">' + _fmtNum(pMO) + '</td>' +
          '</tr>';
      }
    }

    // Blank filler rows
    for (var ei = 0; ei < 4; ei++) {
      tablaFilas += '<tr><td style="border:1px solid #ccc;height:18px" colspan="7">&nbsp;</td></tr>';
    }

    // Date parts
    var tz  = Session.getScriptTimeZone();
    var now = new Date();
    var dia  = Utilities.formatDate(now, tz, 'dd');
    var mes2 = Utilities.formatDate(now, tz, 'MM');
    var anio = Utilities.formatDate(now, tz, 'yyyy');

    // Technician initials
    var tecNombre = data.tecnico || '';
    var tecIni = tecNombre ? tecNombre.split(' ').map(function(w){ return w.charAt(0).toUpperCase(); }).join('') : '—';

    // Borrador watermark
    var wm = data.modoPreview
      ? '<div style="position:fixed;top:45%;left:50%;transform:translate(-50%,-50%) rotate(-35deg);font-size:90px;font-weight:900;color:rgba(200,0,0,0.06);pointer-events:none;white-space:nowrap;z-index:999">BORRADOR</div>'
      : '';

    var html =
      '<!DOCTYPE html><html><head><meta charset="UTF-8">' +
      '<style>' +
      'body{font-family:Arial,sans-serif;margin:0;padding:14px 18px;font-size:11px;color:#222}' +
      'table{border-collapse:collapse}' +
      '.it td,.it th{border:1px solid #ccc}' +
      '.it th{background:#e6e6e6;padding:5px 7px;font-size:10px}' +
      '.total-r td{background:#111;color:#fff;font-weight:700;font-size:13px;padding:6px 9px;text-align:right;border:1px solid #111}' +
      '.fbox{border:1px solid #444;display:inline-block;min-width:28px;padding:2px 6px;text-align:center;font-size:17px;font-weight:700}' +
      '.lbl{font-size:9px;color:#888}' +
      '.nota{font-size:8.5px;color:#555;line-height:1.5}' +
      '.footer-bar{text-align:center;font-size:10px;color:#00a3e0;font-weight:700;padding-top:7px;border-top:2px solid #00a3e0;margin-top:8px}' +
      '@media print{body{padding:8px 10px}.print-bar{display:none}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}' +
      '</style></head><body>' +
      wm +
      // ── HEADER ──
      '<table style="width:100%;margin-bottom:6px"><tr>' +
        '<td style="width:38%;vertical-align:top">' +
          '<div style="font-size:24px;font-weight:900;color:#00a3e0;line-height:1">Bidcom<span style="color:#222">Agro</span></div>' +
          '<div style="font-size:8px;color:#666;line-height:1.5;margin-top:3px">Parque Industrial Carmen de Areco Lote 28<br>B6725 Carmen de Areco, Prov. Buenos Aires<br>Tel: 2325-65-6826</div>' +
        '</td>' +
        '<td style="width:24%;text-align:center;vertical-align:middle">' +
          '<div style="border:3px solid #333;display:inline-block;padding:4px 12px;font-size:34px;font-weight:900;line-height:1">P</div><br>' +
          '<div style="border:2px solid #e74c3c;color:#e74c3c;font-size:8px;font-weight:700;padding:3px 6px;margin-top:4px;display:inline-block;line-height:1.3;text-align:center">DOCUMENTO NO VALIDO<br>COMO FACTURA</div>' +
        '</td>' +
        '<td style="width:38%;text-align:right;vertical-align:top">' +
          '<div style="font-size:19px;font-weight:900;letter-spacing:.04em">PRESUPUESTO</div>' +
          '<div style="font-size:12px;font-weight:700;color:#555;margin-bottom:6px">N° ' + presupNum + '</div>' +
          '<span class="fbox">' + dia + '</span>' +
          '<span style="font-size:11px;color:#aaa;margin:0 3px">|</span>' +
          '<span class="fbox">' + mes2 + '</span>' +
          '<span style="font-size:11px;color:#aaa;margin:0 3px">|</span>' +
          '<span class="fbox">' + anio + '</span>' +
          '<div style="font-size:8px;color:#aaa;margin-top:2px">Día &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mes &nbsp;&nbsp;&nbsp;&nbsp; Año</div>' +
        '</td>' +
      '</tr></table>' +
      // ── RESELLER INFO ──
      '<table class="it" style="width:100%;margin-bottom:5px"><tr>' +
        '<td style="width:12%;padding:4px 7px" class="lbl">Señores:</td>' +
        '<td style="width:50%;padding:4px 7px;font-weight:700;font-size:12px">' + data.reseller + '</td>' +
        '<td style="width:38%;padding:4px 7px" class="lbl">CUIT N°: <strong style="color:#222">' + (rCuit||'—') + '</strong></td>' +
      '</tr><tr>' +
        '<td style="padding:4px 7px" class="lbl">Domicilio:</td>' +
        '<td style="padding:4px 7px;font-size:11px">' + (rDir||'—') + '</td>' +
        '<td style="padding:4px 7px" class="lbl">Localidad: <strong style="color:#222">' + (rLoc||'—') + '</strong></td>' +
      '</tr><tr>' +
        '<td style="padding:4px 7px" class="lbl">Provincia:</td>' +
        '<td style="padding:4px 7px;font-size:10px">Prov. de Buenos Aires &nbsp; <span class="lbl">Cód. Postal:</span> ' + (rCp||'—') + ' &nbsp; <span class="lbl">Descuentos (Repuestos):</span> <strong>' + descLabel + '</strong></td>' +
        '<td style="padding:4px 7px" class="lbl">Técnico Asignado: <strong style="color:#222">' + tecIni + '</strong></td>' +
      '</tr><tr>' +
        '<td style="padding:4px 7px" class="lbl">Teléfono:</td>' +
        '<td style="padding:4px 7px;font-size:11px">' + (rTel||'—') + '</td>' +
        '<td style="padding:4px 7px" class="lbl">Referencia: <strong style="color:#222">' + data.ot + '</strong>' + (data.cas ? ' &nbsp; <span class="lbl">CAS:</span> <strong style="color:#222">' + data.cas + '</strong>' : '') + '</td>' +
      '</tr>' +
      (data.trabajo ? '<tr><td style="padding:4px 7px" class="lbl">Falla reportada:</td><td colspan="2" style="padding:4px 7px;font-size:10px;color:#555;background:#f7f7f7">' + data.trabajo.replace(/\n/g,' ') + '</td></tr>' : '') +
      (data.informeTecnico ? '<tr><td style="padding:4px 7px" class="lbl">Diagnóstico:</td><td colspan="2" style="padding:4px 7px;font-size:10px;color:#2d6a3f;background:#f5fff8">' + data.informeTecnico.replace(/\n/g,' ') + '</td></tr>' : '') +
      '</table>' +
      // ── ITEMS TABLE ──
      '<table class="it" style="width:100%;margin-bottom:5px">' +
        '<thead><tr>' +
          '<th style="width:14%">Cód</th>' +
          '<th style="width:30%">Descripción</th>' +
          '<th style="width:13%;text-align:right">Precio (USD) s/i</th>' +
          '<th style="width:9%;text-align:center">Dto.</th>' +
          '<th style="width:13%;text-align:right">P. Neto (USD)</th>' +
          '<th style="width:8%;text-align:center">Cantidad</th>' +
          '<th style="width:13%;text-align:right">Subtotal (USD)</th>' +
        '</tr></thead>' +
        '<tbody>' + tablaFilas + '</tbody>' +
        '<tfoot><tr class="total-r"><td colspan="6">Total USD (Sin impuestos)</td><td>' + (totalGeneral ? _fmtNum(totalGeneral) : '—') + '</td></tr></tfoot>' +
      '</table>' +
      // ── NOTAS ──
      '<table style="width:100%;border:1px solid #ccc"><tr><td style="padding:6px 9px">' +
        '<div class="nota">NOTA: Este presupuesto se basa en los datos visibles e informados por el cliente durante la inspección. No incluye costos por daños ocultos o no visibles. Si se descubren daños adicionales durante la reparación, se procederá a informar un nuevo presupuesto. Los costos finales pueden variar según los datos identificados durante el proceso de reparación. Este presupuesto tiene una validez de 5 días hábiles a partir de la fecha de su realización.</div>' +
        '<div class="nota" style="margin-top:3px">*Cotización expresada en dólares estadounidenses; los pagos en pesos argentinos se tomarán al TC Banco Nación Venta Billetes al momento de recepción de valores, efectivo o depósito.<br>' +
        '*La presente oferta no constituye un compromiso y está sujeta a aprobación y evaluación de riesgo crediticio.<br>' +
        '*Lugar de Entrega: Depósito Parque Industrial Carmen de Areco. No incluye envío.<br>' +
        '*Condición de pago: 10 días a partir del efectivo pago.</div>' +
      '</td></tr></table>' +
      // ── FOOTER ──
      '<div class="footer-bar">www.dji.bidcomagro.com.ar &nbsp;|&nbsp; www.brumby.com.ar &nbsp;|&nbsp; 0810-345-9002</div>' +
      '</body></html>';

    return { ok: true, html: html };
  } catch(e) {
    Logger.log("obtenerPresupuestoHTML ERROR: " + e);
    return { ok: false, msg: e.toString() };
  }
}
