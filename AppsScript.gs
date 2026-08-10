/**
 * Script para Google Sheets: recibe el inventario escaneado desde la app
 * web y lo escribe en la hoja "Inventario" (la crea si no existe).
 *
 * INSTALACIÓN:
 * 1. Abre tu Google Sheet.
 * 2. Menú Extensiones > Apps Script.
 * 3. Borra el código de ejemplo y pega TODO este archivo.
 * 4. Guarda (icono de disquete).
 * 5. Implementar > Nueva implementación.
 *    - Tipo: Aplicación web
 *    - Ejecutar como: Yo (tu cuenta)
 *    - Quién tiene acceso: Cualquier usuario
 * 6. Autoriza los permisos cuando Google lo pida (es tu propio script,
 *    es normal que diga "app no verificada" - clic en Avanzado > Ir a...).
 * 7. Copia la URL que termina en /exec y pégala en la app del celular.
 */

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Inventario') || ss.insertSheet('Inventario');
  var data = JSON.parse(e.postData.contents);
  var items = data.items || [];

  sheet.clearContents();
  sheet.appendRow(['Codigo', 'Origen salida', 'Origen entrada', 'Cantidad', 'Stock', 'Ultima actualizacion']);

  var now = new Date();
  items.forEach(function (i) {
    sheet.appendRow([i.code, i.origenSalida, i.origenEntrada, i.qty, i.stock, now]);
  });

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', count: items.length }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: 'API de inventario activa' }))
    .setMimeType(ContentService.MimeType.JSON);
}
