/**
 * Backend para el escáner de inventario.
 *
 * IMPORTANTE: este script NUNCA borra ni reescribe toda la hoja.
 * Cada fila se identifica por un ID único generado en la app. Al
 * escanear o editar, solo se agrega o actualiza esa fila puntual.
 * Al borrar un producto en la app, solo se elimina esa fila puntual.
 *
 * INSTALACIÓN:
 * 1. Abre tu Google Sheet.
 * 2. Menú Extensiones > Apps Script.
 * 3. Borra el código de ejemplo y pega TODO este archivo.
 * 4. Guarda (icono de disquete).
 * 5. Implementar > Nueva implementación (o "Administrar implementaciones"
 *    y edita la existente con el lápiz si ya tenías una).
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Autoriza los permisos cuando Google lo pida (es tu propio script).
 * 7. Copia la URL que termina en /exec y pégala en la app.
 *
 * Si ya tenías una implementación, usa "Nueva versión" al editar para
 * que la misma URL /exec use este código actualizado.
 */

var SHEET_NAME = 'Inventario';
var HEADERS = ['ID', 'Codigo', 'Origen salida', 'Origen entrada', 'Cantidad', 'Stock', 'Ultima actualizacion'];

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function findRowById_(sheet, id) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function doPost(e) {
  var sheet = getSheet_();
  var data = JSON.parse(e.postData.contents);
  var action = data.action;
  var item = data.item;

  if (action === 'delete') {
    var row = findRowById_(sheet, item.id);
    if (row !== -1) sheet.deleteRow(row);
  } else {
    var row2 = findRowById_(sheet, item.id);
    var now = new Date();
    var values = [item.id, item.code, item.origenSalida, item.origenEntrada, item.qty, item.stock, now];
    if (row2 === -1) {
      sheet.appendRow(values);
    } else {
      sheet.getRange(row2, 1, 1, values.length).setValues([values]);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var sheet = getSheet_();
  var items = [];
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    rows.forEach(function (r) {
      if (r[1] !== '') {
        items.push({
          id: String(r[0]),
          code: String(r[1]),
          origenSalida: r[2],
          origenEntrada: r[3],
          qty: r[4],
          stock: r[5]
        });
      }
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', items: items }))
    .setMimeType(ContentService.MimeType.JSON);
}
