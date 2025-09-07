/**
 * DATA VALIDATION + SYNC for RepairDesk
 * v1.0 (2025-09-08)
 * Author: Tekfix Assistant
 */

// ==== CONFIG — EDIT THESE TO MATCH YOUR SHEET NAMES / COLUMNS ====
const CFG = {
  tabs: {
    products: 'Products',
    manufacturers: 'Manufacturers',
    devices: 'Devices',
    log: 'Edits_Log'
  },
  cols: {
    // 1-based indexes in Products sheet
    manufacturer: 4, // e.g., column D
    device: 5       // e.g., column E
  },
  api: {
    enabled: false, // set true when ready
    baseUrl: 'https://api.repairdesk.co/', // confirm exact base URL
    apiKey: 'PUT_YOUR_API_KEY_HERE' // keep in PropertiesService for safety
  },
  alias: {
    // Manufacturer aliases → canonical name
    'apple inc': 'Apple',
    'apple pty ltd': 'Apple',
    'samsung electronics': 'Samsung',
    'oppo': 'OPPO',
    'vivo': 'VIVO',
    'xiaomi corp': 'Xiaomi'
  }
};

// ==== HELPERS ====
function logAction(sheet, action, value, extra='') {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CFG.tabs.log) || ss.insertSheet(CFG.tabs.log);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Sheet', 'Action', 'Value', 'Extra']);
  }
  sh.appendRow([new Date(), sheet, action, value, extra]);
}

function toCanonical(s) {
  if (!s) return '';
  let x = String(s).trim().replace(/\s+/g, ' ');
  // Title Case (basic)
  x = x.toLowerCase();
  if (CFG.alias[x]) return CFG.alias[x];
  return x.replace(/\b\w/g, c => c.toUpperCase());
}

function getColRange_(sheetName, colIndex) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  const lastRow = Math.max(2, sh.getMaxRows());
  return sh.getRange(2, colIndex, lastRow - 1, 1); // from row 2 downwards
}

function uniqueAppend_(sheetName, colIndex, value, extraCols=[]) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  if (sh.getLastRow() === 0) sh.appendRow(['Value', 'Extra1', 'Extra2', 'Extra3']);
  const col = sh.getRange(2, colIndex, Math.max(1, sh.getLastRow()-1)).getValues().flat();
  if (!col.map(String).map(s=>s.trim().toLowerCase()).includes(String(value).trim().toLowerCase())) {
    const row = [value, ...extraCols];
    sh.appendRow(row);
    logAction(sheetName, 'APPEND', value, JSON.stringify(extraCols));
    return true;
  }
  return false;
}

function findOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) sh.appendRow(headers);
  } else if (sh.getLastRow() === 0 && headers && headers.length) {
    sh.appendRow(headers);
  }
  return sh;
}

// ==== MENU ====
function dataOpsBuildMenu() {
  SpreadsheetApp.getUi()
    .createMenu('DataOps')
    .addItem('Rebuild Validation', 'rebuildValidation')
    .addItem('Normalize Names (Products)', 'normalizeProductsNames')
    .addSeparator()
    .addItem('Push New to RepairDesk', 'pushNewToRepairDesk')
    .addToUi();
}

// ==== VALIDATION ====
function rebuildValidation() {
  const ss = SpreadsheetApp.getActive();
  const prod = ss.getSheetByName(CFG.tabs.products);
  const man = ss.getSheetByName(CFG.tabs.manufacturers);
  const dev = ss.getSheetByName(CFG.tabs.devices);

  if (!prod || !man || !dev) throw new Error('Missing required tabs. Check CFG.tabs.');

  // Manufacturers dropdown
  const manRange = man.getRange('A2:A');
  const manRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(manRange, true)
    .setAllowInvalid(false)
    .build();
  prod.getRange(2, CFG.cols.manufacturer, prod.getMaxRows()-1, 1).setDataValidation(manRule);

  // Devices dropdown
  const devRange = dev.getRange('A2:A');
  const devRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(devRange, true)
    .setAllowInvalid(false)
    .build();
  prod.getRange(2, CFG.cols.device, prod.getMaxRows()-1, 1).setDataValidation(devRule);

  logAction('Products', 'REBUILD_VALIDATION', 'OK');
}

// ==== NORMALIZATION ====
function normalizeProductsNames() {
  const ss = SpreadsheetApp.getActive();
  const prod = ss.getSheetByName(CFG.tabs.products);
  const mCol = CFG.cols.manufacturer;
  const dCol = CFG.cols.device;
  const mVals = getColRange_(CFG.tabs.products, mCol).getValues();
  const dVals = getColRange_(CFG.tabs.products, dCol).getValues();
  let changed = 0;

  for (let i = 0; i < mVals.length; i++) {
    const m = mVals[i][0];
    const d = dVals[i][0];
    const m2 = toCanonical(m);
    const d2 = toCanonical(d);
    if (m2 !== m) { mVals[i][0] = m2; changed++; }
    if (d2 !== d) { dVals[i][0] = d2; changed++; }
  }
  getColRange_(CFG.tabs.products, mCol).setValues(mVals);
  getColRange_(CFG.tabs.products, dCol).setValues(dVals);
  logAction('Products', 'NORMALIZE', String(changed)+' cells');
}

// ==== AUTO-APPEND ON EDIT ====
function onEdit(e) {
  try {
    const rng = e.range;
    const sh = rng.getSheet();
    if (sh.getName() !== CFG.tabs.products) return;
    const col = rng.getColumn();
    const val = toCanonical(rng.getValue());
    if (!val) return;

    if (col === CFG.cols.manufacturer) {
      const added = uniqueAppend_(CFG.tabs.manufacturers, 1, val);
      if (added) rebuildValidation();
    }
    if (col === CFG.cols.device) {
      // Optional: also capture manufacturer in same row for cross-check
      const man = toCanonical(sh.getRange(rng.getRow(), CFG.cols.manufacturer).getValue());
      const extra = man ? [man] : [];
      const added = uniqueAppend_(CFG.tabs.devices, 1, val, extra);
      if (added) rebuildValidation();
    }
  } catch (err) {
    logAction('Products', 'onEdit_ERROR', String(err));
  }
}

// ==== REPAIRDESK SYNC (Manufacturers + Devices) ====
function pushNewToRepairDesk() {
  if (!CFG.api.enabled) {
    SpreadsheetApp.getUi().alert('API sync disabled. Set CFG.api.enabled=true when ready.');
    return;
  }
  const props = PropertiesService.getDocumentProperties();
  const lastSyncKey = 'lastSyncTs';
  const lastSync = Number(props.getProperty(lastSyncKey) || 0);
  const since = new Date(lastSync || 0);

  const ss = SpreadsheetApp.getActive();
  const man = ss.getSheetByName(CFG.tabs.manufacturers);
  const dev = ss.getSheetByName(CFG.tabs.devices);
  const newM = collectNew_(man, since, 1);
  const newD = collectNew_(dev, since, 1);

  const payload = { manufacturers: newM, devices: newD };
  if (!newM.length && !newD.length) {
    SpreadsheetApp.getUi().alert('Nothing new to push.');
    return;
  }

  const url = CFG.api.baseUrl.replace(/\/$/, '') + '/v1/sync'; // TODO: set actual endpoints
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: { 'Authorization': 'Bearer ' + CFG.api.apiKey, 'Content-Type': 'application/json' },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  logAction('API', 'PUSH', res.getResponseCode(), res.getContentText());
  if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
    props.setProperty(lastSyncKey, String(Date.now()));
    SpreadsheetApp.getUi().alert('Pushed to RepairDesk: ' + JSON.stringify(payload));
  } else {
    SpreadsheetApp.getUi().alert('API error: ' + res.getResponseCode());
  }
}

function collectNew_(sheet, since, colIndex) {
  // naive: treat newly appended rows since last run as new
  // refine by reading Edits_Log entries if needed
  const values = sheet.getRange(2, colIndex, Math.max(0, sheet.getLastRow()-1)).getValues();
  // return non-empty canonical list
  return values
    .map(r => toCanonical(r[0]))
    .filter(v => v)
    .filter((v, i, a) => a.indexOf(v) === i);
}

function setupDataOpsOpenTrigger() {
  // Remove any prior DataOps onOpen triggers to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'dataOpsBuildMenu') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Create a dedicated onOpen trigger for DataOps
  ScriptApp.newTrigger('dataOpsBuildMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}