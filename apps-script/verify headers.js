/***********************
 * Expected header sets
 ***********************/
const EXPECTED = {
  Products: [
    "Item ID","Category","Item Name","Description","Manufacturer","Device","SKU","Supplier","Multiple Supplier SKUs","UPC",
    "Manage Inventory level","Valuation Method","Manage Serials","On Hand Qty","New Stock Adjustment","Cost Price",
    "New Inventory Item Cost","Retail Price","Online Price","Promotional Price","Minimum Price","Tax Class","Tax Inclusive",
    "Stock Warning","Re-Order Level","Condition","Physical Location","Warranty","Warranty Time Frame","IMEI",
    "Display On Point of Sale","Commission Percentage","Commission Amount"
  ],
  Manufacturers: ["Manufacturer","Show on POS","Show on widgets"],
  Devices: ["Manufacturer","Device","Show on POS","Show on widgets"],
};

const PRODUCTS_SYNC_COLS = ["__page_id","__last_pulled_at","__last_pushed_at","__status"];

/************************************
 * Utils: normalize + array helpers *
 ************************************/
const norm = s => String(s||"").trim().replace(/\s+/g," ").toLowerCase();

function compareHeaders_(actual, expected, {allowSyncCols=false}={}) {
  const a = actual.map(norm);
  const e = expected.map(norm);

  // Build sets for contains checks (by normalized name)
  const aSet = new Set(a);
  const eSet = new Set(e);

  // Missing (expected not found in actual)
  const missing = expected.filter(h => !aSet.has(norm(h)));

  // Extras (actual not in expected)
  let extras = actual.filter(h => !eSet.has(norm(h)));

  // If products tab & sync cols allowed, ignore them in "extras"
  if (allowSyncCols) {
    const syncSet = new Set(PRODUCTS_SYNC_COLS.map(norm));
    extras = extras.filter(h => !syncSet.has(norm(h)));
  }



  // Out-of-order (only check those that exist in both)
  const common = expected.filter(h => aSet.has(norm(h)));
  const orderMismatch = common.some((h, i) => a.indexOf(norm(h)) !== actual.indexOf(h));

  return { missing, extras, orderMismatch };
}

/************************************
 * Validate all three tabs (no edit) *
 ************************************/
function validateHeaders() {
  const ss = SpreadsheetApp.getActive();
  const tabs = ["Products","Manufacturers","Devices"];
  const logs = [];

  tabs.forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (!sh) { logs.push(`${tab}: ❌ Sheet not found`); return; }

    const lastCol = sh.getLastColumn() || 0;
    const headers = lastCol ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
    if (!headers.length) { logs.push(`${tab}: ❌ No headers in row 1`); return; }

    const allowSyncCols = (tab === "Products");
    const {missing, extras, orderMismatch} =
      compareHeaders_(headers, EXPECTED[tab], {allowSyncCols});

    logs.push(
      `${tab}: ${missing.length||extras.length||orderMismatch ? '⚠️ Issues' : '✅ OK'}`
      + (missing.length ? ` | Missing: ${missing.join(', ')}` : '')
      + (extras.length ? ` | Extra: ${extras.join(', ')}` : '')
      + (orderMismatch ? ` | Order differs` : '')
    );

    // Check Products sync cols
    if (tab === "Products") {
      const haveSync = PRODUCTS_SYNC_COLS.every(c => headers.map(norm).includes(norm(c)));
      if (!haveSync) logs.push(`Products: ℹ️ Will add missing sync columns at the end on fix: ${PRODUCTS_SYNC_COLS.filter(c => !headers.map(norm).includes(norm(c))).join(', ')}`);
    }
  });

  Logger.log(logs.join('\n'));
  SpreadsheetApp.getActive().toast('Header validation done. See View → Logs.', 'Header Check', 6);
  return logs.join('\n');
}

/*****************************************
 * Auto-fix headers (reorder/add/remove) *
 *****************************************/
function fixHeaders() {
  const ss = SpreadsheetApp.getActive();

  // Products
  fixTabHeaders_(ss, "Products", EXPECTED.Products, {appendSyncCols:true});

  // Manufacturers
  fixTabHeaders_(ss, "Manufacturers", EXPECTED.Manufacturers, {appendSyncCols:false});

  // Devices
  fixTabHeaders_(ss, "Devices", EXPECTED.Devices, {appendSyncCols:false});

  SpreadsheetApp.getActive().toast('✅ Headers fixed (see logs for details).', 'Header Fix', 6);
}

/*************************************************************
 * Fix one tab: overwrite row 1 to exact expected order/names
 * - Adds missing
 * - Drops extras (except sync cols in Products)
 * - Appends sync cols at end for Products
 *************************************************************/
function fixTabHeaders_(ss, tabName, expected, {appendSyncCols}) {
  const sh = ss.getSheetByName(tabName);
  if (!sh) { Logger.log(`${tabName}: sheet not found`); return; }

  // Read all headers
  const lastCol = sh.getLastColumn() || 0;
  const current = lastCol ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];

  // Build the final header line
  let final = expected.slice(); // copy
  if (appendSyncCols) {
    // Ensure sync cols appear at the end (no duplicates)
    PRODUCTS_SYNC_COLS.forEach(c => { if (!final.includes(c)) final.push(c); });
  }

  // Overwrite row 1 with the final header order
  if (final.length === 0) { Logger.log(`${tabName}: no expected headers?`); return; }
  sh.getRange(1,1,1,Math.max(final.length, lastCol || final.length)).clearContent();
  sh.getRange(1,1,1,final.length).setValues([final]);

  // Optional: auto-size first N columns
  for (let c = 1; c <= Math.min(final.length, 40); c++) sh.autoResizeColumn(c);

  // Log diffs
  const {missing, extras, orderMismatch} =
    compareHeaders_(current, expected, {allowSyncCols: tabName==='Products'});
  Logger.log(`${tabName}: fixed. Missing that were added: ${missing.join(', ') || 'none'}. Extras that were dropped: ${extras.join(', ') || 'none'}. Order was ${orderMismatch ? 'changed' : 'already correct'}.`);
}


/***** CONFIG *****/
const PRODUCTS_BATCH_SIZE = 80;     // rows per run
const NOTION_MIN_DELAY_MS = 120;    // polite delay between requests (~8 rps max; Notion ~3 rps cap per integration)
const ONLY_CHANGED_ROWS = true;     // push only new/dirty/error rows (recommended)

/***** CURSOR KEYS *****/
const K_CURSOR_ROW = 'PUSH_CURSOR_ROW';
const K_TARGET_ROWS = 'PUSH_TARGET_ROWS_JSON';

/***** ENTRYPOINTS (add to your menu if you like) *****/
function pushProductsBatched() {
  return pushProductsBatchInternal_(false);  // normal run
}
function pushProductsBatchedResume() {
  return pushProductsBatchInternal_(true);   // resume if cursor exists
}

/***** CORE *****/
function pushProductsBatchInternal_(resume) {
  const P = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
  const sh = ss.getSheetByName(P.getProperty('SHEET_NAME') || 'Products');
  if (!sh) throw new Error('Products sheet not found.');
  ensureNotionHasAllPropertiesForTab('Products'); // make sure schema exists

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const H = Object.fromEntries(headers.map((h,i)=>[h,i]));
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) {
    SpreadsheetApp.getActive().toast('No product rows to push.', 'Notion Push', 5);
    return 'No rows.';
  }

  // Build target row indexes (1-based sheet rows) once, then store
  let targetRows;
  if (resume && P.getProperty(K_TARGET_ROWS)) {
    targetRows = JSON.parse(P.getProperty(K_TARGET_ROWS));
  } else {
    const values = sh.getRange(2,1,lastRow-1,lastCol).getValues();
    targetRows = [];
    for (let i=0;i<values.length;i++){
      const r = values[i];
      const status = H['__status'] != null ? String(r[H['__status']]).trim().toLowerCase() : '';
      const pageId = H['__page_id'] != null ? String(r[H['__page_id']]||'').trim() : '';
      const itemId = H['Item ID'] != null ? String(r[H['Item ID']]||'').trim() : '';
      if (!itemId) continue; // skip empty
      if (ONLY_CHANGED_ROWS) {
        const isNew = !pageId;
        const isDirty = ['dirty','update_error','create_error'].includes(status);
        if (!(isNew || isDirty)) continue;
      }
      targetRows.push(i+2); // convert to absolute row number in sheet
    }
    P.setProperty(K_TARGET_ROWS, JSON.stringify(targetRows));
    P.deleteProperty(K_CURSOR_ROW); // fresh run
  }

  if (targetRows.length === 0) {
    SpreadsheetApp.getActive().toast('Nothing to push (no new/dirty rows).', 'Notion Push', 5);
    P.deleteProperty(K_TARGET_ROWS);
    P.deleteProperty(K_CURSOR_ROW);
    return 'Nothing to push.';
  }

  // Cursor / window
  let cursor = Number(P.getProperty(K_CURSOR_ROW) || '0'); // index into targetRows
  const end = Math.min(cursor + PRODUCTS_BATCH_SIZE, targetRows.length);
  const batch = targetRows.slice(cursor, end);

  // Live Notion property names (prevents “property doesn’t exist” 400s)
  const notionProps = getNotionPropertySet_();

  // Do the work
  let created=0, updated=0, errors=0;
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty('PRODUCTS_DB_ID');

  batch.forEach(absRow => {
    try {
      const rowVals = sh.getRange(absRow,1,1,lastCol).getValues()[0];
      const row = Object.fromEntries(headers.map((h,i)=>[h,rowVals[i]]));
      if (!String(row['Item ID']||'').trim()) return;

      // Build properties from your PROPERTY_MAP but only those that exist in Notion
      const props = {};
      PROPERTY_MAP.forEach(m => {
        if (!notionProps.has(m.notionProp)) return;
        const v = row[m.sheetHeader];
        switch (m.type) {
          case 'title':
            props[m.notionProp] = { title:[{ type:'text', text:{ content:String(v||'') } }] };
            break;
          case 'rich_text':
            props[m.notionProp] = v ? { rich_text:[{ type:'text', text:{ content:String(v) } }] } : { rich_text:[] };
            break;
          case 'number':
            props[m.notionProp] = { number:(v===''||v==null)?null:Number(v) };
            break;
          case 'checkbox':
            props[m.notionProp] = { checkbox: Boolean(v) };
            break;
          case 'select':
            props[m.notionProp] = v ? { select:{ name:String(v) } } : { select:null };
            break;
          case 'multi_select':
            props[m.notionProp] = { multi_select:String(v||'').split(',').map(s=>s.trim()).filter(Boolean).map(name=>({name})) };
            break;
        }
      });

      const pageId = String(row['__page_id']||'').trim();
      let url, method, payload;
      if (pageId) {
        url = 'https://api.notion.com/v1/pages/' + pageId;
        method = 'patch';
        payload = { properties: props };
      } else {
        url = 'https://api.notion.com/v1/pages';
        method = 'post';
        payload = { parent:{ database_id: db }, properties: props };
      }

      const resp = UrlFetchApp.fetch(url, {
        method,
        headers: { Authorization:'Bearer '+key, 'Notion-Version':'2022-06-28', 'Content-Type':'application/json' },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code === 200 || code === 201) {
        const body = JSON.parse(resp.getContentText());
        if (!pageId) {
          // write back new page id
          sh.getRange(absRow, H['__page_id']+1).setValue(body.id);
          created++;
        } else {
          updated++;
        }
        if (H['__status'] != null) sh.getRange(absRow, H['__status']+1).setValue(pageId ? 'updated' : 'created');
        if (H['__last_pushed_at'] != null) sh.getRange(absRow, H['__last_pushed_at']+1).setValue(new Date());
      } else {
        errors++;
        const msg = resp.getContentText();
        if (H['__status'] != null) sh.getRange(absRow, H['__status']+1).setValue('push_error');
        Logger.log(`Row ${absRow} push error: ${msg}`);
      }

      Utilities.sleep(NOTION_MIN_DELAY_MS);
    } catch (e) {
      errors++;
      if (H['__status'] != null) sh.getRange(absRow, H['__status']+1).setValue('push_error');
      Logger.log(`Row ${absRow} error: ${e.message}`);
    }
  });

  // Advance cursor & decide whether to resume
  cursor = end;
  if (cursor < targetRows.length) {
    P.setProperty(K_CURSOR_ROW, String(cursor));
    SpreadsheetApp.getActive().toast(`Pushed ${created+updated} (errors ${errors}). Resuming… ${cursor}/${targetRows.length}`, 'Notion Push', 6);
    // Schedule a continuation trigger shortly (avoids manual rerun spam)
    ScriptApp.newTrigger('pushProductsBatchedResume').timeBased().after(10 * 1000).create();
  } else {
    // Done
    P.deleteProperty(K_CURSOR_ROW);
    P.deleteProperty(K_TARGET_ROWS);
    SpreadsheetApp.getActive().toast(`✅ Done. Created ${created}, updated ${updated}, errors ${errors}.`, 'Notion Push', 6);
  }

  return `Batch: created ${created}, updated ${updated}, errors ${errors}. Cursor ${cursor}/${targetRows.length}`;
}

/***** helper: live Notion property names *****/
function getNotionPropertySet_() {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty('PRODUCTS_DB_ID');
  const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + key, 'Notion-Version': '2022-06-28' },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error(resp.getContentText());
  const props = JSON.parse(resp.getContentText()).properties || {};
  return new Set(Object.keys(props));
}