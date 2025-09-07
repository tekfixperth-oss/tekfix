/***************
 * TAB CONFIGS *
 ***************/
const TABS = {
  Products: {
    dbProp: 'PRODUCTS_DB_ID',
    // You already have PROPERTY_MAP for Products (Option A all rich_text + title)
    // We'll keep using that.
  },
  Manufacturers: {
    dbProp: 'MANUFACTURERS_DB_ID',
    // Treat all headers as rich_text except these checkboxes
    checkboxHeaders: new Set(['Show on POS','Show on widgets']),
  },
  Devices: {
    dbProp: 'DEVICES_DB_ID',
    checkboxHeaders: new Set(['Show on POS','Show on widgets']),
  }
};

/**************************************
 * SHARED: build props for a row/tab  *
 **************************************/
function buildPropsForTab_(tabName, headers, row) {
  const props = {};
  if (tabName === 'Products') {
    // Use your existing PROPERTY_MAP for Products
    PROPERTY_MAP.forEach(m => {
      const v = row[headers.indexOf(m.sheetHeader)];
      switch (m.type) {
        case 'title':
          props[m.notionProp] = { title: [{ type:'text', text:{ content:String(v||'') } }] };
          break;
        case 'rich_text':
          props[m.notionProp] = v ? { rich_text: [{ type:'text', text:{ content:String(v) } }] } : { rich_text: [] };
          break;
        case 'number':
          props[m.notionProp] = { number: (v===''||v==null) ? null : Number(v) };
          break;
        case 'checkbox':
          props[m.notionProp] = { checkbox: Boolean(v) };
          break;
        case 'select':
          props[m.notionProp] = v ? { select: { name: String(v) } } : { select: null };
          break;
        case 'multi_select':
          props[m.notionProp] = { multi_select: String(v||'').split(',').map(s=>s.trim()).filter(Boolean).map(name=>({name})) };
          break;
      }
    });
    return props;
  }

  // Manufacturers / Devices: Option A = make everything rich_text except a few checkboxes
  const checkboxHeaders = TABS[tabName].checkboxHeaders || new Set();
  headers.forEach((h, i) => {
    const v = row[i];
    if (h === 'Item ID' || h === '__page_id' || h === '__last_pulled_at' || h === '__last_pushed_at' || h === '__status') return; // skip meta in these tabs
    if (!h) return;
    if (checkboxHeaders.has(h)) {
      props[h] = { checkbox: Boolean(v) };
    } else {
      props[h] = v ? { rich_text: [{ type:'text', text:{ content:String(v) } }] } : { rich_text: [] };
    }
  });
  return props;
}

/*****************************************
 * Ensure Notion has all properties (tab) *
 *****************************************/
function ensureNotionHasAllPropertiesForTab(tabName) {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty(TABS[tabName].dbProp);
  const ss  = SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
  const sh  = ss.getSheetByName(tabName);
  if (!key || !db) throw new Error(`Missing NOTION_API_KEY or ${TABS[tabName].dbProp}.`);
  if (!sh) throw new Error(`Sheet "${tabName}" not found.`);

  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].filter(Boolean);
  // Fetch current Notion props
  const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + key, 'Notion-Version': '2022-06-28' },
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error(resp.getContentText());
  const currentProps = JSON.parse(resp.getContentText()).properties || {};
  const currentNames = new Set(Object.keys(currentProps));

  // Build patch: any missing header => add prop (rich_text or checkbox for known boolean headers)
  const checkboxHeaders = TABS[tabName].checkboxHeaders || new Set();
  const toAdd = {};
  headers.forEach(h => {
    if (!h) return;
    if (tabName === 'Products' && h === 'Item ID') return; // Products title handled separately
    if (h === 'Name') return; // generic skip
    if (!currentNames.has(h)) {
      toAdd[h] = checkboxHeaders.has(h) ? { checkbox: {} } : { rich_text: {} };
    }
  });

  if (Object.keys(toAdd).length === 0) {
    SpreadsheetApp.getActive().toast(`✅ ${tabName}: No missing properties.`, 'Notion', 5);
    return;
  }

  const patch = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db, {
    method: 'patch',
    headers: {
      Authorization: 'Bearer ' + key,
      'Notion-Version': '2022-06-28',
      'Content-Type':'application/json'
    },
    payload: JSON.stringify({ properties: toAdd }),
    muteHttpExceptions: true
  });
  if (patch.getResponseCode() >= 200 && patch.getResponseCode() < 300) {
    SpreadsheetApp.getActive().toast(`✅ ${tabName}: Added ${Object.keys(toAdd).length} properties.`, 'Notion', 6);
  } else {
    throw new Error(`${tabName} add props failed: ` + patch.getContentText());
  }
}

/*************************
 * PUSH tab → Notion DB  *
 *************************/
function pushTabToNotion(tabName) {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty(TABS[tabName].dbProp);
  const ss  = SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
  const sh  = ss.getSheetByName(tabName);
  if (!key || !db || !sh) throw new Error(`Missing key/db/sheet for ${tabName}.`);

  // Make sure Notion has all props first
  ensureNotionHasAllPropertiesForTab(tabName);

  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return `${tabName}: No data to push.`;
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];
  const values  = sh.getRange(2,1,lastRow-1,lastCol).getValues();

  // If this is Products, locate __page_id; for others we’ll create fresh (optional)
  const H = Object.fromEntries(headers.map((h,i)=>[h,i]));
  const hasPageId = H.hasOwnProperty('__page_id');

  let created=0, updated=0, errors=0;
  values.forEach((row, idx) => {
    try {
      let url, method, payload;
      const props = buildPropsForTab_(tabName, headers, row);

      if (tabName === 'Products' && hasPageId && row[H['__page_id']]) {
        // Update by page id
        url = 'https://api.notion.com/v1/pages/' + row[H['__page_id']];
        method = 'patch';
        payload = { properties: props };
      } else {
        // Create new page
        url = 'https://api.notion.com/v1/pages';
        method = 'post';
        payload = { parent: { database_id: db }, properties: props };
      }

      const resp = UrlFetchApp.fetch(url, {
        method,
        headers: {
          Authorization: 'Bearer ' + key,
          'Notion-Version':'2022-06-28',
          'Content-Type':'application/json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code === 200 || code === 201) {
        const body = JSON.parse(resp.getContentText());
        if (tabName === 'Products' && !(hasPageId && row[H['__page_id']])) {
          // write back new page id
          sh.getRange(idx+2, H['__page_id']+1).setValue(body.id);
          created++;
        } else {
          updated++;
        }
      } else {
        errors++;
        Logger.log(`${tabName} push error row ${idx+2}: ${resp.getContentText()}`);
      }
      Utilities.sleep(120); // keep it gentle
    } catch (e) {
      errors++;
      Logger.log(`${tabName} push error row ${idx+2}: ${e.message}`);
    }
  });

  const msg = `${tabName}: created ${created}, updated ${updated}, errors ${errors}`;
  SpreadsheetApp.getActive().toast('✅ ' + msg, 'Notion Push', 6);
  return msg;
}

/***********************
 * PULL Notion → Sheet *
 ***********************/
function pullTabFromNotion(tabName) {
  const P = PropertiesService.getScriptProperties();
  const key = P.getProperty('NOTION_API_KEY');
  const db  = P.getProperty(TABS[tabName].dbProp);
  const ss  = SpreadsheetApp.openById(P.getProperty('SHEET_ID'));
  const sh  = ss.getSheetByName(tabName);
  if (!key || !db || !sh) throw new Error(`Missing key/db/sheet for ${tabName}.`);

  // fetch all pages
  let cursor = null, pages = [];
  do {
    const payload = { page_size: 100 };
    if (cursor) payload.start_cursor = cursor;
    const resp = UrlFetchApp.fetch('https://api.notion.com/v1/databases/' + db + '/query', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + key, 'Notion-Version': '2022-06-28', 'Content-Type':'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() !== 200) throw new Error(resp.getContentText());
    pages = pages.concat(data.results);
    cursor = data.has_more ? data.next_cursor : null;
  } while (cursor);

  // Build a header set from the sheet (we don’t rebuild headers here)
  const lastCol = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,lastCol).getValues()[0];

  // Convert Notion properties -> row objects using header names
  const rows = pages.map(p => {
    const obj = {};
    headers.forEach(h => {
      if (!h) return;
      const prop = p.properties[h];
      if (!prop) { obj[h] = ''; return; }
      switch (prop.type) {
        case 'title':
          obj[h] = (prop.title||[]).map(x=>x.plain_text).join('');
          break;
        case 'rich_text':
          obj[h] = (prop.rich_text||[]).map(x=>x.plain_text).join('');
          break;
        case 'number':
          obj[h] = prop.number ?? '';
          break;
        case 'checkbox':
          obj[h] = prop.checkbox ? true : false;
          break;
        case 'select':
          obj[h] = prop.select ? prop.select.name : '';
          break;
        case 'multi_select':
          obj[h] = (prop.multi_select||[]).map(x=>x.name).join(', ');
          break;
        default:
          obj[h] = '';
      }
    });
    // also write __page_id if present in sheet
    if (headers.includes('__page_id')) obj['__page_id'] = p.id;
    return obj;
  });

  // Write to sheet (clear existing body)
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  if (rows.length) {
    const data = rows.map(r => headers.map(h => r[h] ?? ''));
    sh.getRange(2,1,rows.length,headers.length).setValues(data);
  }

  const msg = `${tabName}: pulled ${pages.length} rows from Notion.`;
  SpreadsheetApp.getActive().toast('✅ ' + msg, 'Notion Pull', 6);
  return msg;
}

/**********************
 * Convenience menus  *
 **********************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Store Setup')
    .addItem('1) Create/Reset Master Template', 'createMasterTemplate')
    .addItem('2) Install Data Validation', 'installValidation')
    .addSeparator()
    .addItem('Pull Products (Notion → Sheet)', 'pullNotionToSheet')
    .addItem('Push Products (Sheet → Notion)', 'pushSheetToNotion')
    .addSeparator()
    .addItem('Pull Manufacturers', 'pullManufacturers')
    .addItem('Push Manufacturers', 'pushManufacturers')
    .addSeparator()
    .addItem('Pull Devices', 'pullDevices')
    .addItem('Push Devices', 'pushDevices')
    .addSeparator()
    .addItem('Initial Load: Push All', 'initialLoadPushAll')
    .addToUi();
}

// Thin wrappers for menu
function pushManufacturers(){ return pushTabToNotion('Manufacturers'); }
function pullManufacturers(){ return pullTabFromNotion('Manufacturers'); }
function pushDevices(){ return pushTabToNotion('Devices'); }
function pullDevices(){ return pullTabFromNotion('Devices'); }

function initialLoadPushAll() {
  const msgs = [];
  msgs.push(pushManufacturers());
  msgs.push(pushDevices());
  msgs.push(pushSheetToNotion()); // Products last, since it’s the heavy one
  SpreadsheetApp.getActive().toast('✅ ' + msgs.join(' | '), 'Initial Load', 8);
  return msgs.join(' | ');
}