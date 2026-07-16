// @version 1.0
// ══════════════════════════════════════════════════════════════
//  STOCK MANAGER — EQUIPOS EN DEPÓSITO
//  Inventario de EQUIPOS COMPLETOS (por N° de serie), distinto del
//  stock de repuestos (piezas sueltas). Sirve para conciliaciones:
//  qué equipos tenemos, su destino y su estado.
//  FASE 1: alta / lista / edición + reporte. El "qué le falta a cada
//  equipo" se lleva por ahora como nota de texto (FALTANTES); la Fase 2
//  lo estructura en un ledger de partes por equipo.
// ══════════════════════════════════════════════════════════════

var EQ = SCHEMA.EQUIPOS_DEPOSITO;
var EQ_HEADER = ['N\xb0 de serie', 'Modelo', 'Condici\xf3n', 'Destino', 'Estado',
  'Ubicaci\xf3n', 'Fecha ingreso', 'Origen', 'Valor', 'Faltantes (nota)',
  'Observaciones', 'Situaci\xf3n', '\xdaltima actualizaci\xf3n', 'Operador'];

// Crea (si falta) la hoja EQUIPOS_DEPOSITO con encabezado y la devuelve.
function _getHojaEquipos() {
  var db = getDb();
  var hoja = db.getSheetByName(SCHEMA.SHEETS.EQUIPOS_DEPOSITO);
  if (!hoja) {
    hoja = db.insertSheet(SCHEMA.SHEETS.EQUIPOS_DEPOSITO);
    hoja.appendRow(EQ_HEADER);
    hoja.setFrozenRows(1);
    hoja.getRange(1, 1, 1, EQ_HEADER.length).setFontWeight('bold').setBackground('#00a3e0').setFontColor('#fff');
    hoja.setColumnWidth(1, 160); hoja.setColumnWidth(10, 240); hoja.setColumnWidth(11, 240);
  }
  return hoja;
}

// Lista completa de equipos (el front filtra por destino / estado / situación).
function getEquiposDeposito() {
  try {
    var hoja = _getHojaEquipos();
    var d = hoja.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < d.length; i++) {
      var sn = String(d[i][EQ.SN] || '').trim();
      if (!sn) continue;
      out.push({
        sn:            sn,
        modelo:        String(d[i][EQ.MODELO] || ''),
        condicion:     String(d[i][EQ.CONDICION] || ''),
        destino:       String(d[i][EQ.DESTINO] || ''),
        estado:        String(d[i][EQ.ESTADO] || ''),
        ubicacion:     String(d[i][EQ.UBICACION] || ''),
        fechaIngreso:  _fmtFecha(d[i][EQ.FECHA_INGRESO]),
        origen:        String(d[i][EQ.ORIGEN] || ''),
        valor:         Number(d[i][EQ.VALOR]) || 0,
        faltantes:     String(d[i][EQ.FALTANTES] || ''),
        observaciones: String(d[i][EQ.OBSERVACIONES] || ''),
        situacion:     String(d[i][EQ.SITUACION] || '') || 'En dep\xf3sito'
      });
    }
    return { ok: true, equipos: out };
  } catch(e) {
    Logger.log('getEquiposDeposito: ' + e);
    return { ok: false, equipos: [], error: e.toString() };
  }
}

// Alta o edición (upsert por N° de serie).
// data: {sn, snOriginal?, modelo, condicion, destino, estado, ubicacion,
//        fechaIngreso?('YYYY-MM-DD'), origen, valor, faltantes, observaciones, situacion, operador}
function guardarEquipo(data) {
  try {
    data = data || {};
    var sn = String(data.sn || '').trim();
    if (!sn) return { ok: false, error: 'Falta el N\xb0 de serie.' };

    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var hoja = _getHojaEquipos();
      var d = hoja.getDataRange().getValues();
      var buscar = String(data.snOriginal || sn).trim().toUpperCase();
      var fila = 0;
      for (var i = 1; i < d.length; i++) {
        if (String(d[i][EQ.SN] || '').trim().toUpperCase() === buscar) { fila = i + 1; break; }
      }
      // El SN (nuevo o cambiado) no puede pisar a otro equipo existente.
      for (var j = 1; j < d.length; j++) {
        if ((j + 1) !== fila && String(d[j][EQ.SN] || '').trim().toUpperCase() === sn.toUpperCase()) {
          return { ok: false, error: 'Ya existe un equipo con el N\xb0 de serie ' + sn + '.' };
        }
      }

      var ahora = new Date();
      var esAlta = !fila;
      if (esAlta) {
        var fi = data.fechaIngreso ? new Date(String(data.fechaIngreso) + 'T12:00:00') : ahora;
        hoja.appendRow([sn, '', '', '', '', '', fi, '', '', '', '', '', ahora, '']);
        fila = hoja.getLastRow();
      } else if (data.fechaIngreso) {
        hoja.getRange(fila, EQ.FECHA_INGRESO + 1).setValue(new Date(String(data.fechaIngreso) + 'T12:00:00'));
      }

      var set = function(col, val) { hoja.getRange(fila, col + 1).setValue(val); };
      set(EQ.SN,            sn);
      set(EQ.MODELO,        data.modelo || '');
      set(EQ.CONDICION,     data.condicion || '');
      set(EQ.DESTINO,       data.destino || '');
      set(EQ.ESTADO,        data.estado || '');
      set(EQ.UBICACION,     data.ubicacion || '');
      set(EQ.ORIGEN,        data.origen || '');
      set(EQ.VALOR,         (data.valor !== undefined && data.valor !== '' && data.valor !== null) ? (Number(data.valor) || 0) : '');
      set(EQ.FALTANTES,     data.faltantes || '');
      set(EQ.OBSERVACIONES, data.observaciones || '');
      set(EQ.SITUACION,     data.situacion || 'En dep\xf3sito');
      set(EQ.FECHA_ACT,     ahora);
      if (data.operador) set(EQ.OPERADOR, data.operador);

      SpreadsheetApp.flush();
      return { ok: true, alta: esAlta };
    } finally {
      try { lock.releaseLock(); } catch(eL) {}
    }
  } catch(e) {
    Logger.log('guardarEquipo: ' + e);
    return { ok: false, error: e.toString() };
  }
}
