var HOJA_TARIFARIO = 'Tarifario estandar';
var HOJA_CATALOGO = 'CAT_ESTADOS';

// Activa o desactiva cascadas.
var ACTIVAR_ORIGEN = true;
var ACTIVAR_DESTINO = true;

// Columnas ORIGEN (K, L, M).
var ORIGEN_PAIS_COL = 11;
var ORIGEN_ESTADO_COL = 12;
var ORIGEN_CIUDAD_COL = 13;

// Columnas DESTINO (P, Q, R). Ajusta si tu hoja usa otras posiciones.
var DESTINO_PAIS_COL = 16;
var DESTINO_ESTADO_COL = 17;
var DESTINO_CIUDAD_COL = 18;

function onEdit(e) {
  if (!e || !e.range) return;

  var sh = e.range.getSheet();
  if (sh.getName() !== HOJA_TARIFARIO) return;

  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row < 2) return;

  if (ACTIVAR_ORIGEN) {
    manejarCascada(row, col, sh, ORIGEN_PAIS_COL, ORIGEN_ESTADO_COL, ORIGEN_CIUDAD_COL);
  }

  if (ACTIVAR_DESTINO) {
    manejarCascada(row, col, sh, DESTINO_PAIS_COL, DESTINO_ESTADO_COL, DESTINO_CIUDAD_COL);
  }
}

function manejarCascada(row, col, sh, paisCol, estadoCol, ciudadCol) {
  if (col === paisCol) {
    actualizarEstados(sh, row, paisCol, estadoCol, ciudadCol);
    return;
  }

  if (col === estadoCol) {
    actualizarCiudades(sh, row, paisCol, estadoCol, ciudadCol);
  }
}

function actualizarEstados(sh, row, paisCol, estadoCol, ciudadCol) {
  var pais = limpiarTexto(sh.getRange(row, paisCol).getDisplayValue());
  var estadoCell = sh.getRange(row, estadoCol);
  var ciudadCell = sh.getRange(row, ciudadCol);

  estadoCell.clearContent().clearDataValidations();
  ciudadCell.clearContent().clearDataValidations();
  if (!pais) return;

  var cat = SpreadsheetApp.getActive().getSheetByName(HOJA_CATALOGO);
  if (!cat) return;

  var lastRow = cat.getLastRow();
  if (lastRow < 2) return;

  var data = cat.getRange(2, 1, lastRow - 1, 2).getValues();
  var estados = [];
  for (var i = 0; i < data.length; i++) {
    var filaPais = limpiarTexto(data[i][0]);
    var estado = limpiarTexto(data[i][1]);
    if (filaPais === pais && estado && estados.indexOf(estado) === -1) {
      estados.push(estado);
    }
  }

  if (estados.length === 0) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(estados, true)
    .setAllowInvalid(false)
    .build();

  estadoCell.setDataValidation(rule);
}

function actualizarCiudades(sh, row, paisCol, estadoCol, ciudadCol) {
  var pais = limpiarTexto(sh.getRange(row, paisCol).getDisplayValue());
  var estado = limpiarTexto(sh.getRange(row, estadoCol).getDisplayValue());
  var ciudadCell = sh.getRange(row, ciudadCol);

  ciudadCell.clearContent().clearDataValidations();
  if (!pais || !estado) return;

  var cat = SpreadsheetApp.getActive().getSheetByName(HOJA_CATALOGO);
  if (!cat) return;

  var lastRow = cat.getLastRow();
  if (lastRow < 2) return;

  var data = cat.getRange(2, 1, lastRow - 1, 3).getValues();
  var ciudades = [];
  for (var i = 0; i < data.length; i++) {
    var filaPais = limpiarTexto(data[i][0]);
    var filaEstado = limpiarTexto(data[i][1]);
    var ciudad = limpiarTexto(data[i][2]);
    if (filaPais === pais && filaEstado === estado && ciudad && ciudades.indexOf(ciudad) === -1) {
      ciudades.push(ciudad);
    }
  }

  if (ciudades.length === 0) return;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ciudades, true)
    .setAllowInvalid(false)
    .build();

  ciudadCell.setDataValidation(rule);
}

function limpiarTexto(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}
