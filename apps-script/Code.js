// === Store Master Setup (Apps Script) ===
// Creates a Google Sheet structure that mirrors RepairDesk product import format
// and adds Manufacturers & Devices tabs with validation + quality checks.
// Author: ChatGPT (for your store project)
//
// Usage: In your Sheet → Extensions → Apps Script → paste this file, Save → reload the Sheet.
// A custom menu "Store Setup" will appear with actions.

const PRODUCT_SHEET = "Products";
const MANUFACTURERS_SHEET = "Manufacturers";
const DEVICES_SHEET = "Devices";

const PRODUCT_HEADERS = ['Item ID', 'Parent ID', 'Serial Number', 'Category', 'Item Name', 'Description', 'Manufacturer', 'Device', 'SKU', 'Supplier', 'Multiple Supplier SKUs', 'UPC', 'Manage Inventory level', 'Valuation Method', 'Manage Serials', 'On Hand Qty', 'New Stock Adjustment', 'Cost Price', 'Retail Price', 'Online Price', 'Promotional Price', 'Minimum Price', 'Tax Class', 'Tax Inclusive', 'Stock Warning', 'Re-Order Level', 'Condition', 'Physical Location', 'Warranty', 'Warranty Time Frame', 'IMEI', 'Display On Point of Sale', 'Commission Percentage', 'Commission Amount', 'Size', 'Color', 'Network'];

function onOpen_old() {
  SpreadsheetApp.getUi()
    .createMenu('Store Setup')
    .addItem('1) Create/Reset Master Template', 'createMasterTemplate')
    .addItem('2) Install Data Validation', 'installValidation')
    .addSeparator()
    .addItem('Pull from Notion → Sheet', 'pullNotionToSheet')
    .addItem('Push from Sheet → Notion', 'pushSheetToNotion')
    .addSeparator()
    .addItem('Two-way Sync (Push then Pull)', 'twoWaySync') // fixed order
    .addToUi();
}


// --- STEP 1: Create/Reset Master Template ---
function createMasterTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Ensure sheets exist
  const products = getOrCreateSheet_(ss, PRODUCT_SHEET);
  const mans = getOrCreateSheet_(ss, MANUFACTURERS_SHEET);
  const devs = getOrCreateSheet_(ss, DEVICES_SHEET);

  // Reset headers
  resetHeaders_(products, PRODUCT_HEADERS);
  resetHeaders_(mans, ["Manufacturer","Show on POS","Show on widgets"]);
  resetHeaders_(devs, ["Manufacturer","Device","Show on POS","Show on widgets"]);

  // Freeze, filter, width
  prepareSheet_(products);
  prepareSheet_(mans);
  prepareSheet_(devs);

  // Protect header rows
  protectHeader_(products);
  protectHeader_(mans);
  protectHeader_(devs);
}

// --- STEP 2: Data Validation for dropdowns ---
function installValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const products = ss.getSheetByName('Products');
  const mans = ss.getSheetByName('Manufacturers');
  const devs = ss.getSheetByName('Devices');
  if (!products || !mans || !devs) throw new Error("Run 'Create/Reset Master Template' first.");

  // Build header index for Products
  const lastCol = products.getLastColumn();
  const headers = products.getRange(1,1,1,lastCol).getValues()[0];
  const H = Object.fromEntries(headers.map((h,i)=>[h,i])); // name -> index

  // Clear existing validations on Products data area
  if (products.getLastRow() > 1) {
    products.getRange(2,1,products.getMaxRows()-1,lastCol).clearDataValidations();
  }

  // Validation ranges
  const lastRow = Math.max(1000, products.getMaxRows()); // room to grow
  const ynRule = SpreadsheetApp.newDataValidation().requireValueInList(["Yes","No"], true).build();

  // Manufacturer dropdown on correct column (by name)
  if (H['Manufacturer'] != null) {
    const manufacturerRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(mans.getRange("A2:A"), true).setAllowInvalid(false).build();
    products.getRange(2, H['Manufacturer']+1, lastRow-1, 1).setDataValidation(manufacturerRule);
  }

  // Device dropdown on correct column (by name)
  if (H['Device'] != null) {
    const deviceRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(devs.getRange("B2:B"), true).setAllowInvalid(false).build();
    products.getRange(2, H['Device']+1, lastRow-1, 1).setDataValidation(deviceRule);
  }

  // Yes/No dropdowns on Manufacturers & Devices tabs
  mans.getRange("B2:C").setDataValidation(ynRule);
  devs.getRange("C2:D").setDataValidation(ynRule);

  SpreadsheetApp.getActive().toast('✅ Data validation installed (by header names).', 'Store Setup', 6);
}

// --- STEP 3: Quality Rules (conditional formatting) ---
function addQualityRules() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Products');
  if (!sh) throw new Error('Products sheet missing.');

  // Clear existing rules on Products
  sh.setConditionalFormatRules([]);

  const rules = [];
  // Required fields: Category (2), Item Name (3), Manufacturer (7), Device (8), SKU (9), Retail Price (19)
  const requiredCols = [2,3,7,8,9,19];

  const lastRow = Math.max(sh.getLastRow(), 1000); // give room if empty
  const numRows = Math.max(1, lastRow - 1);        // rows starting at 2

  requiredCols.forEach(col => {
    const colA1 = colLetter_(col);
    // In conditional formatting, the formula is evaluated relative to the top-left cell of the range.
    // $ locks the column; row stays relative starting at 2.
    const formula = '=LEN($' + colA1 + '2)=0';
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula)
        .setBackground('#ffe5e5')
        .setRanges([ sh.getRange(2, col, numRows, 1) ])
        .build()
    );
  });

  // Warn if Retail Price (col 19 = S) <= Cost Price (col 18 = R)
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$S2<=$R2')
      .setBackground('#fff3cd')
      .setRanges([ sh.getRange(2, 19, numRows, 1) ])
      .build()
  );

  sh.setConditionalFormatRules(rules);
}

// Helper already in your script:
function colLetter_(n) {
  let s = '';
  while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
  }
  return s;
}


// --- Helpers ---
function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function resetHeaders_(sh, headers) {
  sh.clear();
  sh.getRange(1,1,1,headers.length).setValues([headers]);
}

function prepareSheet_(sh) {
  sh.setFrozenRows(1);
  const range = sh.getDataRange();
  range.createFilter();
  // Auto-fit columns (limit to first 40 columns to be safe)
  for (let c = 1; c <= Math.min(40, sh.getLastColumn()); c++) {
    sh.autoResizeColumn(c);
  }
}

function protectHeader_(sh) {
  const protection = sh.getRange(1,1,1,sh.getLastColumn()).protect();
  protection.setDescription(sh.getName() + " header protected");
  protection.removeEditors(protection.getEditors()); // only owner editable
}

// Utility: convert row/col to A1 for conditional formatting formulas
function getA1_(row, col) {
  return SpreadsheetApp.getActive().getRange(row, col).getA1Notation().replace(/\d+/g,"") + row;
}

function getNotionTitleName_() {
  const P = PropertiesService.getScriptProperties();
  const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + P.getProperty('PRODUCTS_DB_ID'), {
    method:'get',
    headers:{ Authorization:'Bearer '+P.getProperty('NOTION_API_KEY'), 'Notion-Version':'2022-06-28' }
  });
  const props = JSON.parse(resp.getContentText()).properties || {};
  const titleName = Object.keys(props).find(k => props[k].type === 'title') || 'Name';
  Logger.log('Title property = ' + titleName);
}