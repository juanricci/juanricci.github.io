/**
 * BACKEND - Control de Stock + Lista de Compras
 * Este script se pega en Extensions > Apps Script del Google Sheet
 * y se publica como Web App (Deploy > New deployment > Web app).
 *
 * Deploy settings recomendados:
 *   - Execute as: Me
 *   - Who has access: Anyone
 *
 * IMPORTANTE - configurar el token antes de usar:
 *   1. Ejecutá una vez la función setToken() de acá abajo (Run > setToken),
 *      o andá a Project Settings > Script Properties y agregá manualmente
 *      una property llamada TOKEN con el mismo valor que pusiste en el
 *      front-end (constante API_TOKEN en index.html).
 *   2. El valor sugerido generado para vos es:
 *      91e25726f3efbc14634750657c43dcddc2d41d3786be2ee1
 */

const SHEET_STOCK = 'Stock';
const SHEET_LISTA = 'ListaCompras';

const STOCK_HEADERS = ['id', 'producto', 'categoria', 'cantidad', 'unidad', 'stockMinimo', 'actualizado'];
const LISTA_HEADERS = ['id', 'producto', 'categoria', 'cantidad', 'origen', 'comprado'];

// Ejecutar UNA VEZ manualmente desde el editor (Run > setToken) para guardar
// el token en Script Properties. No queda en el código, así no lo ve nadie
// que lea este archivo si alguna vez lo compartís.
function setToken() {
  PropertiesService.getScriptProperties().setProperty(
    'TOKEN',
    '91e25726f3efbc14634750657c43dcddc2d41d3786be2ee1'
  );
}

function checkToken(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!expected) throw new Error('Token no configurado en el servidor. Ejecutá setToken() primero.');
  if (token !== expected) throw new Error('No autorizado');
}

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getOrCreateSheet(name, headers) {
  const ss = getSS();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureSheets() {
  getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
}

function sheetToObjects(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.map((row, i) => {
    const obj = {};
    headers.forEach((h, idx) => (obj[h] = row[idx]));
    obj._row = i + 2; // fila real en el sheet, útil para updates
    return obj;
  });
}

function newId() {
  return Utilities.getUuid();
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- LECTURA ----------

function getAllData() {
  ensureSheets();
  const stockSheet = getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  const listaSheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const stock = sheetToObjects(stockSheet, STOCK_HEADERS);
  const lista = sheetToObjects(listaSheet, LISTA_HEADERS);
  return { stock, lista };
}

// ---------- STOCK ----------

function addProduct(payload) {
  const sheet = getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  const id = newId();
  sheet.appendRow([
    id,
    payload.producto || '',
    payload.categoria || 'Sin categoría',
    Number(payload.cantidad) || 0,
    payload.unidad || 'u',
    Number(payload.stockMinimo) || 0,
    new Date().toISOString()
  ]);
  return getAllData();
}

function updateProduct(payload) {
  const sheet = getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  const items = sheetToObjects(sheet, STOCK_HEADERS);
  const item = items.find(i => i.id === payload.id);
  if (!item) throw new Error('Producto no encontrado');
  const row = item._row;
  if (payload.producto !== undefined) sheet.getRange(row, 2).setValue(payload.producto);
  if (payload.categoria !== undefined) sheet.getRange(row, 3).setValue(payload.categoria);
  if (payload.cantidad !== undefined) sheet.getRange(row, 4).setValue(Number(payload.cantidad));
  if (payload.unidad !== undefined) sheet.getRange(row, 5).setValue(payload.unidad);
  if (payload.stockMinimo !== undefined) sheet.getRange(row, 6).setValue(Number(payload.stockMinimo));
  sheet.getRange(row, 7).setValue(new Date().toISOString());
  return getAllData();
}

// delta puede ser positivo (repongo) o negativo (consumo)
function adjustStock(payload) {
  const sheet = getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  const items = sheetToObjects(sheet, STOCK_HEADERS);
  const item = items.find(i => i.id === payload.id);
  if (!item) throw new Error('Producto no encontrado');
  const nuevaCantidad = Math.max(0, Number(item.cantidad) + Number(payload.delta));
  sheet.getRange(item._row, 4).setValue(nuevaCantidad);
  sheet.getRange(item._row, 7).setValue(new Date().toISOString());
  return getAllData();
}

function deleteProduct(payload) {
  const sheet = getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  const items = sheetToObjects(sheet, STOCK_HEADERS);
  const item = items.find(i => i.id === payload.id);
  if (!item) throw new Error('Producto no encontrado');
  sheet.deleteRow(item._row);
  return getAllData();
}

// ---------- LISTA DE COMPRAS ----------

function addToShoppingList(payload) {
  const sheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const id = newId();
  sheet.appendRow([
    id,
    payload.producto || '',
    payload.categoria || 'Sin categoría',
    Number(payload.cantidad) || 1,
    payload.origen || 'manual',
    false
  ]);
  return getAllData();
}

function updateShoppingListItem(payload) {
  const sheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const items = sheetToObjects(sheet, LISTA_HEADERS);
  const item = items.find(i => i.id === payload.id);
  if (!item) throw new Error('Item no encontrado');
  const row = item._row;
  if (payload.producto !== undefined) sheet.getRange(row, 2).setValue(payload.producto);
  if (payload.categoria !== undefined) sheet.getRange(row, 3).setValue(payload.categoria);
  if (payload.cantidad !== undefined) sheet.getRange(row, 4).setValue(Number(payload.cantidad));
  return getAllData();
}

function toggleComprado(payload) {
  const sheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const items = sheetToObjects(sheet, LISTA_HEADERS);
  const item = items.find(i => i.id === payload.id);
  if (!item) throw new Error('Item no encontrado');
  sheet.getRange(item._row, 6).setValue(!item.comprado);
  return getAllData();
}

function deleteFromShoppingList(payload) {
  const sheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const items = sheetToObjects(sheet, LISTA_HEADERS);
  const item = items.find(i => i.id === payload.id);
  if (!item) throw new Error('Item no encontrado');
  sheet.deleteRow(item._row);
  return getAllData();
}

function clearComprados() {
  const sheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const items = sheetToObjects(sheet, LISTA_HEADERS);
  // borrar de abajo hacia arriba para no romper índices de fila
  items
    .filter(i => i.comprado === true)
    .sort((a, b) => b._row - a._row)
    .forEach(i => sheet.deleteRow(i._row));
  return getAllData();
}

// Genera sugerencias: agrega a la lista los productos de Stock por debajo
// del mínimo que todavía no están (no comprados), y quita de la lista los
// items 'auto' pendientes cuyo stock ya dejó de estar bajo mínimo.
function generateSuggestions() {
  const stockSheet = getOrCreateSheet(SHEET_STOCK, STOCK_HEADERS);
  const listaSheet = getOrCreateSheet(SHEET_LISTA, LISTA_HEADERS);
  const stock = sheetToObjects(stockSheet, STOCK_HEADERS);
  const lista = sheetToObjects(listaSheet, LISTA_HEADERS);

  const stockPorNombre = {};
  stock.forEach(p => { stockPorNombre[p.producto] = p; });

  const yaEnLista = new Set(
    lista.filter(i => !i.comprado).map(i => i.producto)
  );

  const bajos = stock.filter(
    p => Number(p.cantidad) <= Number(p.stockMinimo) && !yaEnLista.has(p.producto)
  );

  bajos.forEach(p => {
    listaSheet.appendRow([
      newId(),
      p.producto,
      p.categoria,
      1,
      'auto',
      false
    ]);
  });

  // productos auto, pendientes, que ya no están bajo mínimo (o fueron repuestos)
  const yaNoBajos = lista.filter(i => {
    if (i.origen !== 'auto' || i.comprado) return false;
    const p = stockPorNombre[i.producto];
    if (!p) return false; // el producto ya no existe en Stock, no se toca
    return Number(p.cantidad) > Number(p.stockMinimo);
  });
  yaNoBajos
    .sort((a, b) => b._row - a._row) // de abajo hacia arriba, para no romper índices
    .forEach(i => listaSheet.deleteRow(i._row));

  return getAllData();
}

// ---------- ENTRY POINTS ----------

function doGet(e) {
  try {
    checkToken(e.parameter.token);
    return jsonOut({ ok: true, data: getAllData() });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    checkToken(body.token);
    const action = body.action;
    const payload = body.payload || {};
    let data;

    switch (action) {
      case 'addProduct':
        data = addProduct(payload);
        break;
      case 'updateProduct':
        data = updateProduct(payload);
        break;
      case 'adjustStock':
        data = adjustStock(payload);
        break;
      case 'deleteProduct':
        data = deleteProduct(payload);
        break;
      case 'addToShoppingList':
        data = addToShoppingList(payload);
        break;
      case 'updateShoppingListItem':
        data = updateShoppingListItem(payload);
        break;
      case 'toggleComprado':
        data = toggleComprado(payload);
        break;
      case 'deleteFromShoppingList':
        data = deleteFromShoppingList(payload);
        break;
      case 'clearComprados':
        data = clearComprados();
        break;
      case 'generateSuggestions':
        data = generateSuggestions();
        break;
      default:
        throw new Error('Acción desconocida: ' + action);
    }

    return jsonOut({ ok: true, data });
  } catch (err) {
    return jsonOut({ ok: false, error: err.message });
  }
}
