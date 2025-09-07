// === Notion ⇄ Google Sheets Sync (v2) ===
// Works with your "Products" tab. Adds pull/push + twoWaySync().
// Setup (Apps Script → Project Settings → Script properties):
//   NOTION_API_KEY        = secret_... (or ntn_...)
//   NOTION_DATABASE_ID    = your Notion DB ID
//   SHEET_ID              = Spreadsheet ID (from URL)
//   SHEET_NAME            = Products   (or your tab name)
//
// Property Map: Adjust 'notionProp' to match your Notion DB property names.
// Types supported: title, rich_text, select, multi_select, number, checkbox, url
//
// Safety: Test in a copy of your Notion DB first.
function logNotionSchema() {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty('NOTION_DATABASE_ID');
  if (!key || !db) throw new Error('Missing NOTION_API_KEY or NOTION_DATABASE_ID.');

  const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'get',
    headers: {
      Authorization: 'Bearer ' + key,
      'Notion-Version': '2022-06-28'
    },
    muteHttpExceptions: true
  });

  const code = resp.getResponseCode();
  if (code !== 200) throw new Error(resp.getContentText());

  const props = JSON.parse(resp.getContentText()).properties;
  Object.keys(props).forEach(name => {
    const type = props[name].type;
    Logger.log(name + ' → ' + type);
  });

  SpreadsheetApp.getActive().toast('Schema logged to View → Logs.', 'Notion', 6);
}
function pushNow() {
  try {
    const msg = pushSheetToNotion();
    SpreadsheetApp.getActive().toast('✅ ' + msg, 'Notion Push', 8);
    Logger.log(msg);
  } catch (e) {
    SpreadsheetApp.getActive().toast('❌ ' + e.message, 'Notion Push Error', 10);
    Logger.log(e.stack || e);
  }
}
const PROP = PropertiesService.getScriptProperties();
const NOTION_API_KEY = PROP.getProperty('NOTION_API_KEY');
const DATABASE_ID    = PROP.getProperty('NOTION_DATABASE_ID');
const SHEET_ID       = PROP.getProperty('SHEET_ID');
const SHEET_NAME     = PROP.getProperty('SHEET_NAME') || 'Products';
// Ensure Notion DB has a property for every Sheet header (except title)
// Creates missing properties in Notion as rich_text (Option A).
function ensureNotionHasAllProperties() {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty('NOTION_DATABASE_ID');
  const ss  = SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
  const sheetName = P.getProperty('SHEET_NAME') || 'Products';
  const sh  = ss.getSheetByName(sheetName);
  if (!key || !db) throw new Error('Missing NOTION_API_KEY or NOTION_DATABASE_ID.');
  if (!sh) throw new Error('Sheet "' + sheetName + '" not found.');

  // 1) Sheet headers
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].filter(Boolean);
  // We map Item ID -> Name (title), so skip "Name" because Notion already has the title
  const desired = new Set(headers.filter(h => h !== 'Name'));

  // 2) Get current Notion properties
  const get = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + key, 'Notion-Version': '2022-06-28' },
    muteHttpExceptions: true
  });
  if (get.getResponseCode() !== 200) throw new Error('Read DB failed: ' + get.getContentText());
  const currentProps = JSON.parse(get.getContentText()).properties || {};
  const currentNames = new Set(Object.keys(currentProps));

  // 3) Figure out what’s missing (by exact name)
  const missing = headers.filter(h => !currentNames.has(h) && h !== 'Name');
  if (!missing.length) {
    SpreadsheetApp.getActive().toast('✅ Notion already has all properties.', 'Notion', 5);
    Logger.log('No missing properties.');
    return;
  }

  // 4) Build PATCH payload: add each missing as rich_text
  const propsToAdd = {};
  missing.forEach(name => { propsToAdd[name] = { rich_text: {} }; });

  const patch = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'patch',
    headers: {
      Authorization: 'Bearer ' + key,
      'Notion-Version': '2022-06-28',
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({ properties: propsToAdd }),
    muteHttpExceptions: true
  });

  if (patch.getResponseCode() >= 200 && patch.getResponseCode() < 300) {
    SpreadsheetApp.getActive().toast('✅ Added ' + missing.length + ' properties to Notion.', 'Notion', 6);
    Logger.log('Added properties: ' + missing.join(', '));
  } else {
    throw new Error('Add props failed: ' + patch.getContentText());
  }
}
// --- Mapping between Sheet headers (RepairDesk format) and Notion properties ---
// Edit the right-hand 'notionProp' to match your Notion database property names.

const PROPERTY_MAP = [
  { sheetHeader: 'Item ID', notionProp: 'Item ID', type: 'title' },

  // all other fields are rich_text because your schema is text-only
  { sheetHeader: 'Parent ID', notionProp: 'Parent ID', type: 'rich_text' },
  { sheetHeader: 'Serial Number', notionProp: 'Serial Number', type: 'rich_text' },
  { sheetHeader: 'Category', notionProp: 'Category', type: 'rich_text' },
  { sheetHeader: 'Item Name', notionProp: 'Item Name', type: 'rich_text' },
  { sheetHeader: 'Description', notionProp: 'Description', type: 'rich_text' },
  { sheetHeader: 'Manufacturer', notionProp: 'Manufacturer', type: 'rich_text' },
  { sheetHeader: 'Device', notionProp: 'Device', type: 'rich_text' },
  { sheetHeader: 'SKU', notionProp: 'SKU', type: 'rich_text' },
  { sheetHeader: 'Supplier', notionProp: 'Supplier', type: 'rich_text' },
  { sheetHeader: 'Multiple Supplier SKUs', notionProp: 'Multiple Supplier SKUs', type: 'rich_text' },
  { sheetHeader: 'UPC', notionProp: 'UPC', type: 'rich_text' },
  { sheetHeader: 'Manage Inventory level', notionProp: 'Manage Inventory level', type: 'rich_text' },
  { sheetHeader: 'Valuation Method', notionProp: 'Valuation Method', type: 'rich_text' },
  { sheetHeader: 'Manage Serials', notionProp: 'Manage Serials', type: 'rich_text' },
  { sheetHeader: 'On Hand Qty', notionProp: 'On Hand Qty', type: 'rich_text' },
  { sheetHeader: 'New Stock Adjustment', notionProp: 'New Stock Adjustment', type: 'rich_text' },
  { sheetHeader: 'Cost Price', notionProp: 'Cost Price', type: 'rich_text' },
  { sheetHeader: 'Retail Price', notionProp: 'Retail Price', type: 'rich_text' },
  { sheetHeader: 'Online Price', notionProp: 'Online Price', type: 'rich_text' },
  { sheetHeader: 'Promotional Price', notionProp: 'Promotional Price', type: 'rich_text' },
  { sheetHeader: 'Minimum Price', notionProp: 'Minimum Price', type: 'rich_text' },
  { sheetHeader: 'Tax Class', notionProp: 'Tax Class', type: 'rich_text' },
  { sheetHeader: 'Tax Inclusive', notionProp: 'Tax Inclusive', type: 'rich_text' },
  { sheetHeader: 'Stock Warning', notionProp: 'Stock Warning', type: 'rich_text' },
  { sheetHeader: 'Re-Order Level', notionProp: 'Re-Order Level', type: 'rich_text' },
  { sheetHeader: 'Condition', notionProp: 'Condition', type: 'rich_text' },
  { sheetHeader: 'Physical Location', notionProp: 'Physical Location', type: 'rich_text' },
  { sheetHeader: 'Warranty', notionProp: 'Warranty', type: 'rich_text' },
  { sheetHeader: 'Warranty Time Frame', notionProp: 'Warranty Time Frame', type: 'rich_text' },
  { sheetHeader: 'IMEI', notionProp: 'IMEI', type: 'rich_text' },
  { sheetHeader: 'Display On Point of Sale', notionProp: 'Display On Point of Sale', type: 'rich_text' },
  { sheetHeader: 'Commission Percentage', notionProp: 'Commission Percentage', type: 'rich_text' },
  { sheetHeader: 'Commission Amount', notionProp: 'Commission Amount', type: 'rich_text' },
  { sheetHeader: 'Size', notionProp: 'Size', type: 'rich_text' },
  { sheetHeader: 'Color', notionProp: 'Color', type: 'rich_text' },
  { sheetHeader: 'Network', notionProp: 'Network', type: 'rich_text' },
  { sheetHeader: 'Notes', notionProp: 'Notes', type: 'rich_text' }, // helper
];

const SYS_HEADERS = ['__page_id','__last_pulled_at','__last_pushed_at','__status'];

// --- HTTP helper for Notion ---
function notionRequest_(endpoint, method, payload) {
  const url = 'https://api.notion.com/v1' + endpoint;
  const options = {
    method: method || 'get',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + NOTION_API_KEY,
      'Notion-Version': '2022-06-28',
    },
    muteHttpExceptions: true
  };
  if (payload) options.payload = JSON.stringify(payload);
  const resp = UrlFetchApp.fetch(url, options);
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code >= 200 && code < 300) return JSON.parse(text);
  throw new Error('Notion API ' + code + ': ' + text);
}

// --- Sheet helpers ---
function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  // Ensure all headers exist (append system headers if needed)
  const needed = PROPERTY_MAP.map(p => p.sheetHeader).concat(SYS_HEADERS);
  const firstRow = sh.getRange(1,1,1,Math.max(needed.length, sh.getLastColumn() || 1)).getValues()[0];
  if (firstRow.filter(String).length === 0) {
    sh.getRange(1,1,1,needed.length).setValues([needed]);
  } else {
    // Fill any missing headers at the end
    let headers = firstRow.slice();
    needed.forEach(h => { if (!headers.includes(h)) headers.push(h); });
    sh.getRange(1,1,1,headers.length).setValues([headers]);
  }
  return sh;
}

function readRows_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
  return values.map(row => {
    const obj = {};
    headers.forEach((h,i) => obj[h] = row[i]);
    return obj;
  });
}

function writeRows_(sh, rows) {
  if (!rows.length) return;
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const matrix = rows.map(r => headers.map(h => (r[h] === undefined ? '' : r[h])));
  sh.getRange(2,1,matrix.length,headers.length).setValues(matrix);
}

function clearData_(sh) {
  const r = sh.getLastRow();
  if (r > 1) sh.getRange(2,1,r-1,sh.getLastColumn()).clearContent();
}

// --- Converters ---
function fromNotionProps_(props) {
  const row = {};
  PROPERTY_MAP.forEach(m => {
    const p = props[m.notionProp];
    let val = '';
    switch (m.type) {
      case 'title':      val = (p && p.title) ? p.title.map(t => t.plain_text).join('') : ''; break;
      case 'rich_text':  val = (p && p.rich_text) ? p.rich_text.map(t => t.plain_text).join('') : ''; break;
      case 'select':     val = p && p.select ? p.select.name : ''; break;
      case 'multi_select': val = p && p.multi_select ? p.multi_select.map(s => s.name).join(', ') : ''; break;
      case 'number':     val = (p && typeof p.number === 'number') ? p.number : ''; break;
      case 'checkbox':   val = p && p.checkbox ? true : false; break;
      case 'url':        val = p && p.url ? p.url : ''; break;
      default:           val = '';
    }
    row[m.sheetHeader] = val;
  });
  return row;
}

function toNotionProps_(row) {
  const props = {};
  PROPERTY_MAP.forEach(m => {
    const v = row[m.sheetHeader];
    switch (m.type) {
      case 'title':
        props[m.notionProp] = { title: [{ type: 'text', text: { content: String(v || '') } }] };
        break;
      case 'rich_text':
        props[m.notionProp] = { rich_text: v ? [{ type: 'text', text: { content: String(v) } }] : [] };
        break;
      case 'select':
        props[m.notionProp] = v ? { select: { name: String(v) } } : { select: null };
        break;
      case 'multi_select':
        props[m.notionProp] = { multi_select: String(v||'').split(',').map(s => s.trim()).filter(Boolean).map(name => ({ name })) };
        break;
      case 'number':
        props[m.notionProp] = { number: (v === '' || v === null) ? null : Number(v) };
        break;
      case 'checkbox':
        props[m.notionProp] = { checkbox: Boolean(v) };
        break;
      case 'url':
        props[m.notionProp] = { url: v ? String(v) : null };
        break;
    }
  });
  return props;
}

// --- Pull: Notion → Sheet ---
function pullNotionToSheet() {
  if (!NOTION_API_KEY || !DATABASE_ID || !SHEET_ID) {
    throw new Error('Missing NOTION_API_KEY, NOTION_DATABASE_ID, or SHEET_ID in Script Properties.');
  }
  const sh = getSheet_();
  clearData_(sh);

  let hasMore = true;
  let cursor = null;
  const out = [];
  while (hasMore) {
    const payload = { page_size: 100 };
    if (cursor) payload.start_cursor = cursor;
    const data = notionRequest_('/databases/' + DATABASE_ID + '/query', 'post', payload);
    data.results.forEach(page => {
      const row = fromNotionProps_(page.properties);
      row['__page_id'] = page.id;
      row['__last_pulled_at'] = new Date();
      row['__last_pushed_at'] = '';
      row['__status'] = 'pulled';
      out.push(row);
    });
    hasMore = data.has_more;
    cursor = data.next_cursor;
  }
  writeRows_(sh, out);
  return 'Pulled ' + out.length + ' rows from Notion.';
}

// --- Push: Sheet → Notion ---
function pushSheetToNotion() {
  if (!NOTION_API_KEY || !DATABASE_ID || !SHEET_ID) {
    throw new Error('Missing NOTION_API_KEY, NOTION_DATABASE_ID, or SHEET_ID in Script Properties.');
  }
  const sh = getSheet_();
  const rows = readRows_(sh);
  const updated = [];
  const created = [];
  const errors = [];

  // decide create vs update
  rows.forEach(r => {
    const hasName = !!(r['Item Name'] && String(r['Item Name']).trim());
    if (r['__page_id']) updated.push(r);
    else if (hasName) created.push(r);
  });

  // updates
  updated.forEach(r => {
    const props = toNotionProps_(r);
    try {
      notionRequest_('/pages/' + r['__page_id'], 'patch', { properties: props });
      markRow_(sh, r, 'updated');
      Utilities.sleep(200);
    } catch (e) {
      errors.push('update ' + r['Item Name'] + ': ' + e.message);
      markRow_(sh, r, 'update_error');
    }
  });

  // creates
  created.forEach(r => {
    const props = toNotionProps_(r);
    try {
      const res = notionRequest_('/pages', 'post', { parent: { database_id: DATABASE_ID }, properties: props });
      r['__page_id'] = res.id;
      markRow_(sh, r, 'created');
      Utilities.sleep(200);
    } catch (e) {
      errors.push('create ' + r['Item Name'] + ': ' + e.message);
      markRow_(sh, r, 'create_error');
    }
  });

  if (errors.length) return 'Done with errors: ' + errors.join('; ');
  return 'Pushed ' + (updated.length + created.length) + ' rows to Notion.';
}

function markRow_(sh, rowObj, status) {
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const data = readRows_(sh);
  // match by __page_id if present, else by Item Name
  let idx = -1;
  if (rowObj['__page_id']) {
    idx = data.findIndex(r => r['__page_id'] === rowObj['__page_id']);
  }
  if (idx < 0) {
    idx = data.findIndex(r => r['Item Name'] === rowObj['Item Name']);
  }
  const rowNum = idx >= 0 ? idx + 2 : 2; // 1-based with header

  const statusCol = headers.indexOf('__status') + 1;
  const pushedCol = headers.indexOf('__last_pushed_at') + 1;
  if (statusCol > 0) sh.getRange(rowNum, statusCol).setValue(status);
  if (pushedCol > 0) sh.getRange(rowNum, pushedCol).setValue(new Date());
}

// --- Convenience combo ---
function twoWaySync() {
  const pushMsg = pushSheetToNotion();   // push first to preserve your edits
  const pullMsg = pullNotionToSheet();   // then refresh from Notion
  SpreadsheetApp.getActive().toast('✅ ' + pushMsg + ' | ' + pullMsg, 'Two-way Sync', 6);
  return pushMsg + ' | ' + pullMsg;
}

function logNotionSchema() {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty('NOTION_DATABASE_ID');
  if (!key || !db) throw new Error('Missing NOTION_API_KEY or NOTION_DATABASE_ID.');

  const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + key, 'Notion-Version': '2022-06-28' },
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const body = JSON.parse(resp.getContentText());
  if (code !== 200) throw new Error('Notion ' + code + ': ' + resp.getContentText());

  const props = body.properties;
  Object.keys(props).forEach(name => {
    const p = props[name];
    const type = p.type; // e.g., title, rich_text, number, select, checkbox
    Logger.log(name + ' → ' + type);
  });
  SpreadsheetApp.getActive().toast('Logged Notion schema to View → Logs.', 'Notion', 5);
}

function exportRepairDeskCSV() {
  const RD_HEADERS = [
    "Item ID","Parent ID","Serial Number","Category","Item Name","Description",
    "Manufacturer","Device","SKU","Supplier","Multiple Supplier SKUs","UPC",
    "Manage Inventory level","Valuation Method","Manage Serials","On Hand Qty",
    "New Stock Adjustment","Cost Price","Retail Price","Online Price","Promotional Price",
    "Minimum Price","Tax Class","Tax Inclusive","Stock Warning","Re-Order Level","Condition",
    "Physical Location","Warranty","Warranty Time Frame","IMEI","Display On Point of Sale",
    "Commission Percentage","Commission Amount","Size","Color","Network"
  ];

  const sh = SpreadsheetApp.getActive().getSheetByName('Products');
  if (!sh) throw new Error('Products sheet not found.');
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) throw new Error('No data to export.');

  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const values  = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  // Build rows in the exact RD header order; missing columns become blank
  const headerIndex = Object.fromEntries(headers.map((h,i)=>[h,i]));
  const dataOrdered = values.map(row =>
    RD_HEADERS.map(h => (headerIndex[h] != null ? row[headerIndex[h]] : ""))
  );

  // CSV encode
  const needsQuote = s => /[,"\n]/.test(s);
  const toCSVcell = v => {
    const s = (v == null ? "" : String(v));
    return needsQuote(s) ? `"${s.replace(/"/g,'""')}"` : s;
    };
  const csv = [RD_HEADERS.map(toCSVcell).join(',')]
    .concat(dataOrdered.map(r => r.map(toCSVcell).join(',')))
    .join('\n');

  // Save
  const folderName = 'RepairDesk Exports';
  const folders = DriveApp.getFoldersByName(folderName);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  const ts = new Date();
  const name = `Products_Export_${ts.getFullYear()}-${String(ts.getMonth()+1).padStart(2,'0')}-${String(ts.getDate()).padStart(2,'0')}_${String(ts.getHours()).padStart(2,'0')}${String(ts.getMinutes()).padStart(2,'0')}.csv`;
  const file = folder.createFile(name, csv, MimeType.CSV);

  SpreadsheetApp.getActive().toast('CSV created: ' + file.getUrl(), 'RepairDesk Export', 6);
  Logger.log('CSV URL: ' + file.getUrl());
}


function validateMapAgainstNotion() {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty('NOTION_DATABASE_ID');
  if (!key || !db) throw new Error('Missing NOTION_API_KEY or NOTION_DATABASE_ID.');

  const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + key, 'Notion-Version': '2022-06-28' },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error(resp.getContentText());
  const notionNames = new Set(Object.keys(JSON.parse(resp.getContentText()).properties || {}));

  const missing = [];
  const ok = [];
  PROPERTY_MAP.forEach(m => {
    if (m.type === 'title') {
      // Notion has exactly one title; ensure we’re pointing at the real title name
      if (!notionNames.has(m.notionProp)) missing.push(`[TITLE] ${m.notionProp}`);
      else ok.push(m.notionProp);
    } else {
      if (!notionNames.has(m.notionProp)) missing.push(m.notionProp);
      else ok.push(m.notionProp);
    }
  });

  Logger.log('✅ Found in Notion: ' + JSON.stringify(ok));
  Logger.log('❌ Missing in Notion: ' + JSON.stringify(missing));
  const msg = missing.length ? `Missing in Notion: ${missing.join(', ')}` : 'All mapped properties exist in Notion.';
  SpreadsheetApp.getActive().toast(msg, 'Schema check', 8);
}