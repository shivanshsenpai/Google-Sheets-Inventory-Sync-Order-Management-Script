/* =============================================================================
 * Inventory management sync
 *
 * Attach this file to the spreadsheet that contains:
 * - MatchingTable
 * - Inventory
 *
 * It keeps the external All Orders spreadsheet stock status updated when
 * Inventory or MatchingTable is edited, and keeps Inventory images loaded from
 * IMAGE_URL.
 * ============================================================================= */

var ORDER_SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // Replace with your actual Spreadsheet ID
var ALL_ORDERS_SPREADSHEET_ID = ORDER_SPREADSHEET_ID;
var NEW_ORDERS_SPREADSHEET_ID = ORDER_SPREADSHEET_ID;

var SHEET_NAMES = {
  NEW_ORDERS:     'New_Orders',
  ALL_ORDERS:     'All Orders',
  MATCHING_TABLE: 'MatchingTable',
  INVENTORY:      'Inventory'
};

var HEADER_ROW = 2;
var DATA_START_ROW = 3;
var CLEAR_COLOR = '#ffffff';
var INVENTORY_EDIT_TRIGGER_HANDLER = 'handleInventoryManagementEdit';

var HEADERS = {
  SKU:          10,
  QTY:          15,
  Product_Cost: 36,
  STOCK:        41,
  what_to_fix:  42
};

var MATCHING_COLS = {
  VendorSKU: 1,
  PortalSKU: 2
};

var INVENTORY_COLS = {
  VendorSKU:    1,
  Image:        2,
  Item_Name:    3,
  Stock:        4,
  Product_Cost: 5,
  Location:     6,
  IMAGE_URL:    7
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Inventory Sync')
    .addItem('Install edit sync trigger', 'installInventoryEditTrigger')
    .addItem('Refresh All Orders stock', 'refreshAllOrdersStockAvailability')
    .addItem('Refresh New Orders SKUs', 'refreshNewOrdersProductInfo')
    .addItem('Refresh Inventory images', 'refreshAllInventoryImages')
    .addToUi();
}

/**
 * Simple trigger: safe local-only behavior for the Inventory spreadsheet.
 * Cross-spreadsheet sync needs the installable trigger created below.
 */
function onEdit(e) {
  handleInventoryImageEdit_(e);
}

function installInventoryEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === INVENTORY_EDIT_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger(INVENTORY_EDIT_TRIGGER_HANDLER)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    'Inventory sync trigger installed.\n\nEdits in Inventory or MatchingTable will now refresh All Orders stock and New Orders SKU status.'
  );
}

/**
 * Installable trigger handler. Run installInventoryEditTrigger() once so this
 * can update the separate All Orders spreadsheet by ID.
 */
function handleInventoryManagementEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();
  if (sheetName !== SHEET_NAMES.INVENTORY && sheetName !== SHEET_NAMES.MATCHING_TABLE) return;

  handleInventoryImageEdit_(e);
  refreshAllOrdersStockAvailability_();
  refreshNewOrdersProductInfo_();
}

function handleInventoryImageEdit_(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  if (sheet.getName() !== SHEET_NAMES.INVENTORY) return;

  var startRow = Math.max(2, e.range.getRow());
  var endRow = e.range.getLastRow();
  if (endRow < 2) return;

  var startCol = e.range.getColumn();
  var endCol = e.range.getLastColumn();
  if (isColumnInRange_(INVENTORY_COLS.IMAGE_URL, startCol, endCol)) {
    updateInventoryImageFormulas_(sheet, startRow, endRow);
  }
}

function refreshAllOrdersStockAvailability() {
  refreshAllOrdersStockAvailability_();
  SpreadsheetApp.getActiveSpreadsheet().toast('All Orders stock status refreshed.', 'Inventory Sync', 5);
}

function refreshNewOrdersProductInfo() {
  refreshNewOrdersProductInfo_();
  SpreadsheetApp.getActiveSpreadsheet().toast('New Orders SKU status refreshed.', 'Inventory Sync', 5);
}

function refreshAllInventoryImages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.INVENTORY);
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  updateInventoryImageFormulas_(sheet, 2, lastRow);
  SpreadsheetApp.getActiveSpreadsheet().toast('Inventory images refreshed.', 'Inventory Sync', 5);
}

function updateInventoryImageFormulas_(sheet, startRow, endRow) {
  var numRows = endRow - startRow + 1;
  if (numRows <= 0) return;

  var urls = sheet.getRange(startRow, INVENTORY_COLS.IMAGE_URL, numRows, 1).getValues();
  var formulas = [];

  for (var i = 0; i < numRows; i++) {
    var url = String(urls[i][0]).trim();
    if (!url) {
      formulas.push(['']);
    } else {
      formulas.push(['=IMAGE("' + url.replace(/"/g, '""') + '")']);
    }
  }

  sheet.getRange(startRow, INVENTORY_COLS.Image, numRows, 1).setFormulas(formulas);
}

function refreshAllOrdersStockAvailability_() {
  var allOrdersSheet = getAllOrdersSheet_();
  if (!allOrdersSheet) return;

  var lastRow = allOrdersSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  var skuCol = findHeaderColByAliases_(allOrdersSheet, HEADER_ROW, [
    'sku',
    'portal sku',
    'portalsku'
  ]) || HEADERS.SKU;

  var qtyCol = findHeaderColByAliases_(allOrdersSheet, HEADER_ROW, [
    'qty',
    'quantity'
  ]) || HEADERS.QTY;

  var stockStatusCol = findHeaderColByAliases_(allOrdersSheet, HEADER_ROW, [
    'stock',
    'stock_availibilty',
    'stock_availability',
    'stockavailability'
  ]);
  if (!stockStatusCol) return;

  var numRows = lastRow - DATA_START_ROW + 1;
  var skuVals = allOrdersSheet.getRange(DATA_START_ROW, skuCol, numRows, 1).getValues();
  var qtyVals = allOrdersSheet.getRange(DATA_START_ROW, qtyCol, numRows, 1).getValues();

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var out = [];
  var bg = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuVals[i][0]).trim();
    var qty = Number(qtyVals[i][0]) || 0;

    if (!sku) {
      out.push(['']);
      bg.push([CLEAR_COLOR]);
      continue;
    }

    var vendor = portalMap[sku];
    if (!vendor || !invMap[vendor]) {
      out.push(['']);
      bg.push([CLEAR_COLOR]);
      continue;
    }

    var status = getStockAvailabilityStatus_(invMap[vendor].Stock, qty);
    out.push([status]);

    if (status.indexOf('Out of stock') === 0) {
      bg.push(['#ffe0b2']);
    } else {
      bg.push([CLEAR_COLOR]);
    }
  }

  var stockRange = allOrdersSheet.getRange(DATA_START_ROW, stockStatusCol, numRows, 1);
  stockRange.setValues(out);
  stockRange.setBackgrounds(bg);
  stockRange.setWrap(true);
}

function refreshNewOrdersProductInfo_() {
  var sheet = getNewOrdersSheet_();
  if (!sheet) return;

  var lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return;

  var numRows = lastRow - DATA_START_ROW + 1;
  var skuValues = sheet.getRange(DATA_START_ROW, HEADERS.SKU, numRows, 1).getValues();
  var existingErrVals = sheet.getRange(DATA_START_ROW, HEADERS.what_to_fix, numRows, 1).getValues();

  var portalMap = getPortalToVendorMap_();
  var invMap = getInventoryMap_();
  var costOut = [];
  var stockOut = [];
  var errOut = [];

  for (var i = 0; i < numRows; i++) {
    var sku = String(skuValues[i][0]).trim();
    var errorsArr = removeSkuLookupErrors_(existingErrVals[i][0]);

    if (!sku) {
      costOut.push(['']);
      stockOut.push(['']);
      errOut.push([errorsArr.join(' | ')]);
      continue;
    }

    var info = lookupProductInfo_(sku, portalMap, invMap);
    if (info.ok) {
      costOut.push([info.Product_Cost]);
      stockOut.push([info.Stock]);
      errOut.push([errorsArr.join(' | ')]);
    } else {
      costOut.push(['']);
      stockOut.push(['']);
      if (info.reason === 'vendor_sku_missing') {
        errorsArr.push('Vendor SKU not found: ' + sku);
      } else if (info.reason === 'vendor_details_missing') {
        errorsArr.push('Vendor details not found in Inventory: ' + info.vendorSKU);
      } else {
        errorsArr.push('SKU not found: ' + sku);
      }
      errOut.push([errorsArr.join(' | ')]);
    }
  }

  sheet.getRange(DATA_START_ROW, HEADERS.Product_Cost, numRows, 1).setValues(costOut);
  sheet.getRange(DATA_START_ROW, HEADERS.STOCK, numRows, 1).setValues(stockOut);
  sheet.getRange(DATA_START_ROW, HEADERS.what_to_fix, numRows, 1).setValues(errOut);
}

function getAllOrdersSheet_() {
  var ss = SpreadsheetApp.openById(ALL_ORDERS_SPREADSHEET_ID);
  return ss ? ss.getSheetByName(SHEET_NAMES.ALL_ORDERS) : null;
}

function getNewOrdersSheet_() {
  var ss = SpreadsheetApp.openById(NEW_ORDERS_SPREADSHEET_ID);
  return ss ? ss.getSheetByName(SHEET_NAMES.NEW_ORDERS) : null;
}

function getPortalToVendorMap_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.MATCHING_TABLE);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var vendorSKU = String(data[i][MATCHING_COLS.VendorSKU - 1]).trim();
    var portalSKU = String(data[i][MATCHING_COLS.PortalSKU - 1]).trim();
    if (portalSKU && vendorSKU) {
      map[portalSKU] = vendorSKU;
    }
  }

  return map;
}

function getInventoryMap_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAMES.INVENTORY);
  if (!sheet) return {};

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  var data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var map = {};

  for (var i = 0; i < data.length; i++) {
    var vsku = String(data[i][INVENTORY_COLS.VendorSKU - 1]).trim();
    if (vsku) {
      map[vsku] = {
        Product_Cost: Number(data[i][INVENTORY_COLS.Product_Cost - 1]) || 0,
        Stock: Number(data[i][INVENTORY_COLS.Stock - 1]) || 0
      };
    }
  }

  return map;
}

function lookupProductInfo_(sku, portalMap, invMap) {
  var vendorSKU = portalMap[sku];
  if (!vendorSKU) {
    return { ok: false, reason: 'vendor_sku_missing', sku: sku };
  }

  var inv = invMap[vendorSKU];
  if (!inv) {
    return { ok: false, reason: 'vendor_details_missing', sku: sku, vendorSKU: vendorSKU };
  }

  return {
    ok: true,
    reason: 'ok',
    Product_Cost: inv.Product_Cost,
    Stock: inv.Stock,
    VendorSKU: vendorSKU
  };
}

function removeSkuLookupErrors_(value) {
  var parts = String(value || '').split('|');
  var out = [];

  for (var i = 0; i < parts.length; i++) {
    var err = String(parts[i] || '').trim();
    if (!err) continue;
    if (isSkuLookupErrorMessage_(err)) continue;
    if (err.toLowerCase().indexOf('all fields complete') !== -1) continue;
    if (err === 'Stock is 0') continue;
    out.push(err);
  }

  return out;
}

function isSkuLookupErrorMessage_(msg) {
  var m = String(msg || '').toLowerCase();
  return m.indexOf('sku not found') !== -1 ||
         m.indexOf('vendor sku not found') !== -1 ||
         m.indexOf('vendor details not found in inventory') !== -1;
}

function getStockAvailabilityStatus_(stock, qty) {
  var s = Number(stock);
  var q = Number(qty);
  if (isNaN(s) || isNaN(q)) return '';

  if (s >= q && s > 0) return 'Available : ' + s;

  var needed = Math.max(0, q - s);
  return 'Out of stock : ' + s + '\nNeed : ' + needed;
}

function findHeaderColByAliases_(sheet, headerRow, aliases) {
  if (!sheet) return 0;

  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return 0;

  var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  var aliasSet = {};

  for (var i = 0; i < aliases.length; i++) {
    aliasSet[normalizeHeaderKey_(aliases[i])] = true;
  }

  for (var c = 0; c < headers.length; c++) {
    if (aliasSet[normalizeHeaderKey_(headers[c])]) return c + 1;
  }

  return 0;
}

function normalizeHeaderKey_(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function isColumnInRange_(col, startCol, endCol) {
  return col >= startCol && col <= endCol;
}
