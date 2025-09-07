/*** Edits tab helpers (status-aware) — SKU/UPC-first key resolution ***/
function ensureEditsTab_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('Edits');
  if (!sh) {
    sh = ss.insertSheet('Edits');
    sh.getRange(1,1,1,6).setValues([['Action','Key (SKU/UPC/Item ID)','Column','New Value','Status','Notes']]);
    sh.setFrozenRows(1);
  } else {
    // Ensure headers exist in the right order (non-destructive)
    const want = ['Action','Key (SKU/UPC/Item ID)','Column','New Value','Status','Notes'];
    const hdr = sh.getRange(1,1,1,Math.max(want.length, sh.getLastColumn())).getValues()[0];
    // Preserve any existing header text if present; otherwise apply desired header
    const out = want.map((w,i)=> hdr[i] || w);
    sh.getRange(1,1,1,out.length).setValues([out]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function applyEditsFromEditsTab() {
  const ss = SpreadsheetApp.getActive();
  const prod = ss.getSheetByName('Products');
  const edits = ensureEditsTab_();
  if (!prod) throw new Error('Products tab not found.');

  // ---- Build header map for Products
  const pLastCol = prod.getLastColumn();
  const pHeaders = prod.getRange(1,1,1,pLastCol).getValues()[0];
  const H = Object.fromEntries(pHeaders.map((h,i)=>[h,i])); // header -> index

  const eLastRow = edits.getLastRow();
  if (eLastRow < 2) {
    SpreadsheetApp.getActive().toast('No edits to apply. Add rows to Edits.', 'Edits', 5);
    return;
  }

  // Read edits rows (A:F = Action, Key, Column, New Value, Status, Notes)
  const eVals = edits.getRange(2,1,eLastRow-1,6).getValues();

  // Preload Products data once
  const pVals = prod.getRange(2,1,Math.max(prod.getLastRow()-1,0), pLastCol).getValues();

  // Indices in Products (if present)
  const idxSKU    = H['SKU'];
  const idxUPC    = H['UPC'];
  const idxItemID = H['Item ID'];  // read-only, fallback match only

  // Build lookup maps value -> sheetRow for fast matching
  const mapSKU    = new Map();
  const mapUPC    = new Map();
  const mapItemID = new Map();
  for (let i=0; i<pVals.length; i++) {
    const sheetRow = i + 2; // 1-based row (with header)
    if (idxSKU != null) {
      const sku = String(pVals[i][idxSKU] || '').trim();
      if (sku && !mapSKU.has(sku)) mapSKU.set(sku, sheetRow);
    }
    if (idxUPC != null) {
      const upc = String(pVals[i][idxUPC] || '').trim();
      if (upc && !mapUPC.has(upc)) mapUPC.set(upc, sheetRow);
    }
    if (idxItemID != null) {
      const itemId = String(pVals[i][idxItemID] || '').trim();
      if (itemId && !mapItemID.has(itemId)) mapItemID.set(itemId, sheetRow);
    }
  }

  let applied=0, noMatch=0, badCol=0, noChange=0, errors=0;

  for (let r=0; r<eVals.length; r++){
    const [actionRaw, keyRaw, colRaw, newValRaw, statusRaw] = eVals[r];

    const status = String(statusRaw||'').toLowerCase().trim();
    // Process if blank or 'ready'. (If you want ONLY 'ready', change to: if (status !== 'ready') continue;)
    if (status && status !== 'ready') continue;

    const action  = String(actionRaw||'set').toLowerCase().trim();
    const key     = String(keyRaw||'').trim();            // SKU or UPC or Item ID
    const colName = String(colRaw||'').trim();
    const newVal  = newValRaw;

    // Prepare where to write Status/Notes back
    const eRow = r + 2; // edits sheet row number
    const setStatus = (s,note) => {
      edits.getRange(eRow, 5).setValue(s);         // Status
      edits.getRange(eRow, 6).setValue(note||'');  // Notes
    };

    if (!key || !colName) {
      errors++;
      setStatus('error','Missing Key or Column');
      continue;
    }
    if (H[colName] == null) {
      badCol++;
      setStatus('bad_column', `Column not found in Products: ${colName}`);
      continue;
    }

    // Resolve product row by trying SKU -> UPC -> Item ID
    let pRow = null;
    if (key) {
      if (pRow == null && mapSKU.has(key))    pRow = mapSKU.get(key);
      if (pRow == null && mapUPC.has(key))    pRow = mapUPC.get(key);
      if (pRow == null && mapItemID.has(key)) pRow = mapItemID.get(key);
    }
    if (!pRow) {
      noMatch++;
      setStatus('no_match', `Key not found (checked SKU, UPC, Item ID): ${key}`);
      continue;
    }

    try {
      if (action !== 'set') {
        setStatus('error', `Unsupported action: ${action}`);
        errors++;
        continue;
      }

      // Read current value and write new value
      const colIndex = H[colName] + 1; // 1-based for Range
      const curVal = prod.getRange(pRow, colIndex).getValue();

      // If unchanged, mark as no_change
      const same = String(curVal) === String(newVal);
      if (same) {
        noChange++;
        setStatus('no_change', 'Already up-to-date');
        continue;
      }

      prod.getRange(pRow, colIndex).setValue(newVal);

      // Optional dirty markers if your schema has these columns
      const statusCol = (H['__status'] != null) ? (H['__status']+1) : null;
      const pushedCol = (H['__last_pushed_at'] != null) ? (H['__last_pushed_at']+1) : null;
      if (statusCol) prod.getRange(pRow, statusCol).setValue('dirty');
      if (pushedCol) prod.getRange(pRow, pushedCol).setValue(''); // clear last pushed timestamp

      setStatus('applied', `Was: ${curVal} → Now: ${newVal}`);
      applied++;
    } catch (e) {
      errors++;
      setStatus('error', e.message);
    }
  }

  SpreadsheetApp.getActive().toast(
    `Edits: applied ${applied} • no_change ${noChange} • no_match ${noMatch} • bad_column ${badCol} • errors ${errors}`,
    'Edits', 8
  );
}

// ---- Edits menu (installable trigger, won't collide with other menus) ----
function editsBuildMenu() {
  SpreadsheetApp.getUi()
    .createMenu('Edits')
    .addItem('Open “Edits” tab', 'editsOpenTab')
    .addItem('Apply Edits (A→F)', 'applyEditsFromEditsTab')
    .addToUi();
}

function setupEditsOpenTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === 'editsBuildMenu') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('editsBuildMenu')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
}

// Small helper to jump to the Edits sheet
function editsOpenTab() {
  const sh = ensureEditsTab_();
  SpreadsheetApp.setActiveSheet(sh);
}