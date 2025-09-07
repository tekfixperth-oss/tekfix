/** Build a RepairDesk Inventory Import CSV from one invoice.
 *  Creates a Drive file named: RD_InventoryImport_<invoice>.csv
 *  Columns: SKU, Product Name, Cost Price, UPC, Retail Price (optional), Tax, Warranty (days)
 */
function buildRdProductImportFromInvoice() {
  const ui = SpreadsheetApp.getUi();
  const invoice = ui.prompt('RD Inventory Import', 'Enter Invoice Number to export:', ui.ButtonSet.OK_CANCEL);
  if (invoice.getSelectedButton() !== ui.Button.OK) return;
  const invNo = invoice.getResponseText().trim();
  if (!invNo) return;

  const ss = SpreadsheetApp.getActive();
  const invSh = ss.getSheetByName(INV_CFG.tabs.invoiceLines);
  const pendSh = ss.getSheetByName(INV_CFG.tabs.pending);
  if (!invSh || !pendSh) { ui.alert('Missing sheets (Invoice_Lines or Pending_Products).'); return; }

  const Hin = headersMap_(invSh);
  const Pin = headersMap_(pendSh);

  // Read all lines for this invoice
  const invVals = invSh.getRange(2,1, Math.max(0, invSh.getLastRow()-1), invSh.getLastColumn()).getValues()
    .filter(r => String(r[Hin['Invoice Number']]).trim() === invNo);

  if (!invVals.length) { ui.alert('No lines found for that invoice.'); return; }

  // Build map of suggestions from Pending_Products (for names/category if queued)
  const pendVals = pendSh.getRange(2,1, Math.max(0, pendSh.getLastRow()-1), pendSh.getLastColumn()).getValues();
  const pendBySku = new Map();
  pendVals.forEach(r => {
    const sku = String(r[Pin['SKU']] || '').trim().toUpperCase();
    if (!sku) return;
    pendBySku.set(sku, {
      name: String(r[Pin['Product Name']] || '').trim(),
      warranty: Number(r[Pin['Warranty (days)']] || INV_CFG.defaults.warrantyDays),
      tax: String(r[Pin['Tax Code']] || INV_CFG.defaults.taxCode).trim()
    });
  });

  // Build CSV rows
  const rows = [];
  rows.push(['SKU','Product Name','Cost Price','UPC','Retail Price','Tax','Warranty (days)']);
  invVals.forEach(r => {
    const sku = String(r[Hin['SKU']] || '').trim();
    if (!sku) return;
    const upc = String(r[Hin['UPC']] || '').trim();
    const supplierName = String(r[Hin['Supplier Name']] || '').trim();
    const cost = Number(r[Hin['Unit Price']] || 0);
    const pend = pendBySku.get(sku.toUpperCase()) || {};
    const name = pend.name || supplierName || sku;
    const retail = Math.round(cost * INV_CFG.defaults.priceMarkup * 100) / 100;
    const tax = pend.tax || INV_CFG.defaults.taxCode;
    const warranty = pend.warranty || INV_CFG.defaults.warrantyDays;

    rows.push([sku, name, cost, upc, retail, tax, warranty]);
  });

  // Write to Drive
  const csv = rows.map(row => row.map(v => {
    const s = (v==null ? '' : String(v));
    return /[",\n]/.test(s) ? `"${s.replace(/"/g,'""')}"` : s;
  }).join(',')).join('\n');
  const fileName = `RD_InventoryImport_${invNo}.csv`;
  const file = DriveApp.createFile(fileName, csv, MimeType.CSV);

  ui.alert(`Created ${fileName} in your Drive.`);
}

/** Build a GRN helper CSV (manual entry checklist) for one invoice.
 *  Creates: GRN_Helper_<invoice>.csv with SKU, Qty, Unit Cost, Subtotal, Name
 */
function buildGrnHelperFromInvoice() {
  const ui = SpreadsheetApp.getUi();
  const invoice = ui.prompt('GRN Helper Export', 'Enter Invoice Number to export:', ui.ButtonSet.OK_CANCEL);
  if (invoice.getSelectedButton() !== ui.Button.OK) return;
  const invNo = invoice.getResponseText().trim();
  if (!invNo) return;

  const ss = SpreadsheetApp.getActive();
  const invSh = ss.getSheetByName(INV_CFG.tabs.invoiceLines);
  if (!invSh) { ui.alert('Invoice_Lines sheet not found.'); return; }

  const H = headersMap_(invSh);
  const invVals = invSh.getRange(2,1, Math.max(0, invSh.getLastRow()-1), invSh.getLastColumn()).getValues()
    .filter(r => String(r[H['Invoice Number']]).trim() === invNo);

  if (!invVals.length) { ui.alert('No lines found for that invoice.'); return; }

  const rows = [];
  rows.push(['SKU','Qty','Unit Cost','Subtotal','Supplier Name']);
  invVals.forEach(r => {
    rows.push([
      String(r[H['SKU']] || '').trim(),
      Number(r[H['Qty']] || 0),
      Number(r[H['Unit Price']] || 0),
      Number(r[H['Subtotal']] || 0),
      String(r[H['Supplier Name']] || '').trim()
    ]);
  });

  const csv = rows.map(row => row.join(',')).join('\n');
  const fileName = `GRN_Helper_${invNo}.csv`;
  DriveApp.createFile(fileName, csv, MimeType.CSV);
  ui.alert(`Created ${fileName} in your Drive.`);
}