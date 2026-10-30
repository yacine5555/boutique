/**
 * BOUTIQUE PRODUCT SYNC — Apps Script
 * ------------------------------------------------------------
 * What this does:
 *   - doGet():  Reads "Product Master" and serves it as JSON (used by
 *               the Scanner app and the Barcode Generator to check for
 *               existing products/barcodes).
 *   - doPost(): Writes a new barcode into an existing product row, OR
 *               appends a brand-new product row — called automatically
 *               by the Barcode Generator's "Add New Product" button.
 *               Also handles inventory counts (Counter app) and, as of
 *               this version, nightly stock movements (Stock Ledger app).
 *
 * SETUP (one-time):
 *   1. Open your Google Sheet (Store_Inventory...).
 *   2. Extensions → Apps Script.
 *   3. Delete any existing code in Code.gs and paste this whole file in.
 *   4. Click Save (disk icon).
 *   5. Set your write key below (WRITE_KEY) to something only you know —
 *      this stops random people from writing to your sheet if they ever
 *      guessed your Web App URL. Keep it simple, e.g. your store name + a number.
 *   6. Deploy → Manage deployments → pencil icon (edit) → Version: "New version" → Deploy.
 *      (If this is your very first deployment, use Deploy → New deployment → Web app →
 *      Execute as "Me" → Who has access "Anyone" → Deploy, then authorize when prompted.)
 *   7. Copy the Web app URL (ends in /exec) — same URL goes in the Scanner app's
 *      Product Sync field, the Barcode Generator's Live Sync field, the Inventory
 *      Counter's sync field, AND the new Stock Ledger app's sync field.
 *
 * IMPORTANT — Re-deploying after edits:
 *   Editing this script and clicking Save is NOT enough to apply changes.
 *   You must create a "New version" under Manage deployments every time
 *   you change this file, or the live URL keeps running the old code.
 *
 * Column layout expected in "Product Master" (row 3 = headers, data from row 4):
 *   A:#  B:Product Name  C:Brand  D:Category  E:Sub-Category
 *   F:Unit  G:Pack Size  H:Barcode  I:Cost (USD)  J:Price (CFA)  K:Notes
 *
 * Column layout expected in each location tab ("Store 1", "Store 2", "Warehouse"),
 * same header row (3) / first data row (4) convention:
 *   B:Product Name  E:Qty Counted  G:Status  M:Barcode
 *   "Qty On Hand" (new, self-creating — see STOCK LEDGER section below) is
 *   added automatically the first time it's needed, right after whatever
 *   the last used column is. You do not need to add it by hand.
 */
const SHEET_NAME = "Product Master";
const HEADER_ROW = 3;
const FIRST_DATA_ROW = 4;
// Change this to anything private only you know. The Barcode Generator
// must send this same value for writes to be accepted.
const WRITE_KEY = "PUT-YOUR-WRITE-KEY-HERE"; // the real key is only in Apps Script, never on GitHub
/**
 * handleUpdateCounts — called by the Inventory Counter app when the user
 * taps "Envoyer au Sheet". Expects:
 *   {
 *     "key": "...",
 *     "action": "updateCounts",
 *     "location": "Store 1" | "Store 2" | "Warehouse",
 *     "counts": [ { "barcode": "...", "count": 3 }, ... ]
 *   }
 *
 * Writes into the matching location tab's "Qty Counted" (col E) and
 * "Status" (col G) columns, matched by Barcode (col M). Never touches
 * Product Master directly. If a counted barcode doesn't have a row in
 * this location tab yet — e.g. a product that was added to Product
 * Master via the Barcode Generator but never given a row on this
 * specific store's tab — a new row is created for it (name pulled from
 * Product Master) instead of the count being silently dropped. Only a
 * barcode that isn't in Product Master EITHER is reported back as
 * genuinely not found.
 */
function handleUpdateCounts(body){
  const validLocations = ["Store 1", "Store 2", "Warehouse"];
  if (validLocations.indexOf(body.location) === -1) {
    return jsonResponse({ success: false, error: "Unknown location: " + body.location }, 400);
  }
  if (!Array.isArray(body.counts) || !body.counts.length) {
    return jsonResponse({ success: false, error: "No counts provided" }, 400);
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(body.location);
  if (!sheet) {
    return jsonResponse({ success: false, error: "Sheet '" + body.location + "' not found" }, 500);
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) {
    return jsonResponse({ success: false, error: "Destination sheet is empty" });
  }
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const destBarcodes = sheet.getRange(FIRST_DATA_ROW, 13, numRows, 1).getValues(); // M: Barcode
  // Build barcode -> row lookup once
  const rowByBarcode = {};
  for (let i = 0; i < destBarcodes.length; i++) {
    const bc = barcodeKey(destBarcodes[i][0]);
    if (bc) rowByBarcode[bc] = FIRST_DATA_ROW + i;
  }
  // Only pay the cost of reading Product Master if we actually need it
  // to name a self-created row.
  let masterByBarcode = null;
  const needsMaster = body.counts.some(entry => !rowByBarcode[barcodeKey(entry.barcode)]);
  if (needsMaster) masterByBarcode = getProductMasterByBarcode();
  let updated = 0;
  let added = 0;
  const notFound = [];
  body.counts.forEach(entry => {
    const bc = normalizeBarcode(entry.barcode);
    const bk = barcodeKey(entry.barcode);
    let row = rowByBarcode[bk];
    if (!row) {
      const master = masterByBarcode[bk];
      if (!master) { notFound.push(bc); return; } // not in Product Master either — nothing to name it with
      row = sheet.getLastRow() + 1;
      sheet.getRange(row, 2).setValue(master.name); // B: Product Name
      const bcCell = sheet.getRange(row, 13);        // M: Barcode
      bcCell.setNumberFormat("@");
      bcCell.setValue(bc);
      rowByBarcode[bk] = row; // guard against the same new barcode appearing twice in this batch
      added++;
    } else {
      updated++;
    }
    sheet.getRange(row, 5).setValue(entry.count);                              // E: Qty Counted
    sheet.getRange(row, 7).setValue(entry.count > 0 ? "✅ Counted" : "⚠️ Zero"); // G: Status
  });
  try {
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    sheet.getRange(2, 1).setValue("  Last count date: " + dateStr);
  } catch (e) { /* non-critical */ }
  // Bring counted items to the top of the list right away, so it's obvious
  // at a glance what's left to count. See sortLocationByCountedStatus below.
  try { sortLocationByCountedStatus(body.location); } catch (e) { Logger.log('Auto-sort skipped: ' + e); }

  // Zone-level counts, if the Counter app sent any (body.zones), get written
  // to the "Zone Stock" tab's "Zone Qty Counted" column — same two-phase
  // pattern as the location-level Qty Counted -> Qty On Hand above: this
  // step only records what a zone-count found, it never touches Zone Qty On
  // Hand directly. That happens later, when seedQtyOnHand (below) is called
  // after "Envoyer au Sheet" — mirroring the existing location workflow.
  let zonesUpdated = 0;
  if (body.zones && typeof body.zones === 'object') {
    try { zonesUpdated = applyZoneCounts(body.location, body.zones, body.zoneMeta || {}); }
    catch (e) { Logger.log('Zone count write skipped: ' + e); }
  }

  // Shelf-level counts (Counter app "emplacements"): body.spots =
  // { "1B-3-2": { zone: "1B", items: [{barcode, count, name}] }, ... }.
  // Written to the "Spot Stock" tab — where each product sits, shelf by shelf.
  let spotsUpdated = 0;
  if (body.spots && typeof body.spots === 'object' && Object.keys(body.spots).length) {
    try { spotsUpdated = applySpotCounts(body.location, body.spots, body.session || '', body.countedBy || ''); }
    catch (e) { Logger.log('Spot count write skipped: ' + e); }
  }

  return jsonResponse({
    success: true,
    updated: updated,
    added: added,
    notFound: notFound,
    zonesUpdated: zonesUpdated,
    spotsUpdated: spotsUpdated,
    location: body.location
  });
}
/**
 * getProductMasterByBarcode — barcode -> {name, category} lookup built
 * from Product Master. Used by handleUpdateCounts to self-heal a
 * location tab that's missing a row for a product that legitimately
 * exists in the master list.
 */
function getProductMasterByBarcode() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const map = {};
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return map;
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const data = sheet.getRange(FIRST_DATA_ROW, 1, numRows, 8).getValues(); // through H: Barcode
  data.forEach(row => {
    const name = row[1];      // B
    const category = row[3];  // D
    const barcode = row[7];   // H
    if (!barcode || !name) return;
    map[barcodeKey(barcode)] = { name: String(name).trim(), category: category ? String(category).trim() : "" };
  });
  return map;
}
function doGet(e) {
  try {
    // Route by action parameter
    if (e.parameter && e.parameter.action === 'getCustomers') {
      return handleGetCustomers();
    }
    if (e.parameter && e.parameter.action === 'getStockSummary') {
      return handleGetStockSummary();
    }
    if (e.parameter && e.parameter.action === 'getCounterUsers') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return jsonResponse({ success: true, users: readCounterUsers_(e.parameter.app === 'formation') });
    }
    if (e.parameter && e.parameter.action === 'getFormation') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return handleGetFormation();
    }
    if (e.parameter && e.parameter.action === 'getMapLayouts') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return jsonResponse({ success: true, layouts: readMapLayouts_() });
    }
    if (e.parameter && e.parameter.action === 'getCatalogue') {
      return jsonResponse({ success: true, lists: readCatalogue_() });
    }
    if (e.parameter && e.parameter.action === 'getZonePhotos') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return jsonResponse({ success: true, traces: readZonePhotoTraces_() });
    }
    if (e.parameter && e.parameter.action === 'getControlReport') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return jsonResponse({ success: true, text: buildControlReport_(e.parameter.date || '').text });
    }
    if (e.parameter && e.parameter.action === 'getZoneStock') {
      return handleGetZoneStock(e.parameter.location);
    }
    if (e.parameter && e.parameter.action === 'getPriceList') {
      return handleGetPriceList(e.parameter.key === WRITE_KEY);
    }
    if (e.parameter && e.parameter.action === 'getPriceTasks') {
      return handleGetPriceTasks(e.parameter.location || '');
    }
    if (e.parameter && e.parameter.action === 'getLosses') {
      return handleGetLosses(e.parameter.days);
    }
    if (e.parameter && e.parameter.action === 'getExpiryOverview') {
      return handleGetExpiryOverview();
    }
    if (e.parameter && e.parameter.action === 'getDateTodo') {
      return handleGetDateTodo(e.parameter.location);
    }
    if (e.parameter && e.parameter.action === 'getExpiryWatchlist') {
      return handleGetExpiryWatchlist(e.parameter.location);
    }
    if (e.parameter && e.parameter.action === 'getPrintQueue') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return handleGetPrintQueue(e.parameter.store || '');
    }
    if (e.parameter && e.parameter.action === 'getStock1981') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return handleGetStock1981();
    }
    if (e.parameter && e.parameter.action === 'getTransit') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return handleGetTransit();
    }
    if (e.parameter && e.parameter.action === 'getArrivals') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return handleGetArrivals();
    }
    if (e.parameter && e.parameter.action === 'getMovements') {
      if (e.parameter.key !== WRITE_KEY) return jsonResponse({ error: "Invalid key" }, 403);
      return handleGetMovements(e.parameter.days);
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ error: "Sheet '" + SHEET_NAME + "' not found" }, 500);
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < FIRST_DATA_ROW) {
      return jsonResponse({ products: [], count: 0, generatedAt: new Date().toISOString() });
    }
    const numRows = lastRow - FIRST_DATA_ROW + 1;
    const data = sheet.getRange(FIRST_DATA_ROW, 1, numRows, lastCol).getValues();
    const reqLoc = (e.parameter && e.parameter.location) || '';
    const spByKey = (reqLoc === 'Store 1' || reqLoc === 'Store 2') ? _spMap(_spSheet()) : null;
    const products = [];
    const aliasMap = loadAliasMap_();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[7] && aliasMap[barcodeKey(row[7])]) continue; // replaced by its real barcode (second code)
      const name = row[1];     // B: Product Name
      const category = row[3]; // D: Category
      const subCategory = row[4]; // E: Sub-Category
      const unit = row[5];     // F: Unit
      const barcode = row[7];  // H: Barcode
      const price = row[9];    // J: Price (CFA)
      if (!barcode || !name) continue; // skip rows without a barcode or name
      products.push({
        barcode: normalizeBarcode(barcode),
        name: String(name).trim(),
        category: category ? String(category).trim() : "",
        unit: unit ? String(unit).trim() : "",
        subCategory: subCategory !== "" && subCategory !== null ? String(subCategory).trim() : "",
        price: (spByKey && spByKey[barcodeKey(barcode)] && spByKey[barcodeKey(barcode)][reqLoc === 'Store 1' ? 's1' : 's2'] != null)
          ? spByKey[barcodeKey(barcode)][reqLoc === 'Store 1' ? 's1' : 's2']
          : ((price === "" || price === null) ? null : Number(price))
      });
    }
    return jsonResponse({
      products: products,
      aliases: aliasMap,
      count: products.length,
      generatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}
/**
 * doPost — called by the Barcode Generator when a new barcode is
 * created, by the Counter app when sending counts, and by the Stock
 * Ledger app for nightly stock movements. Expects a JSON body whose
 * "action" field routes to the right handler; the default (no action)
 * is the original Barcode Generator behavior.
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.key !== WRITE_KEY) {
      return jsonResponse({ error: "Invalid key" }, 403);
    }
    if (body.action === "setCounterUser") {
      return handleSetCounterUser(body);
    }
    if (body.action === "formationResult") {
      return handleFormationResult(body);
    }
    if (body.action === "formationSuivi") {
      return handleFormationSuivi(body);
    }
    if (body.action === "setMapLayout") {
      return handleSetMapLayout(body);
    }
    if (body.action === "setCatalogueList") {
      return handleSetCatalogueList(body);
    }
    if (body.action === "setZonePhotoTraces") {
      return handleSetZonePhotoTraces(body);
    }
    if (body.action === "getControlPlan") {
      return handleGetControlPlan(body);
    }
    if (body.action === "controlCount") {
      return handleControlCount(body);
    }
    if (body.action === "logRestock") {
      return handleLogRestock(body);
    }
    if (body.action === "resolveEcart") {
      return handleResolveEcart(body);
    }
    if (body.action === "updateCounts") {
      return handleUpdateCounts(body);
    }
    if (body.action === "saveCustomer") {
      return handleSaveCustomer(body);
    }
    if (body.action === "applyStockMovements") {
      return handleApplyStockMovements(body);
    }
    if (body.action === "seedQtyOnHand") {
      return handleSeedQtyOnHand(body);
    }
    if (body.action === "addNewProductWithStock") {
      return handleAddNewProductWithStock(body);
    }
    if (body.action === "addAliasedBarcode") {
      return handleAddAliasedBarcode(body);
    }
    if (body.action === "setStorePrice") {
      return handleSetStorePrice(body);
    }
    if (body.action === "copyStorePrices") {
      return handleCopyStorePrices(body);
    }
    if (body.action === "donePriceTask") {
      return handleDonePriceTask(body);
    }
    if (body.action === "updatePrice") {
      return handleUpdatePrice(body);
    }
    if (body.action === "logContainerIntake") {
      return handleLogContainerIntake(body);
    }
    if (body.action === "saveExpiryCheck") {
      return handleSaveExpiryCheck(body);
    }
    if (body.action === "logExpiryBatch") {
      return handleLogExpiryBatch(body);
    }
    if (body.action === "applyExpiryAction") {
      return handleApplyExpiryAction(body);
    }
    if (body.action === "resolveExpiryBatch") {
      return handleResolveExpiryBatch(body);
    }
    if (body.action === "stock1981Adjust") {
      return handleStock1981Adjust(body);
    }
    if (body.action === "stock1981Count") {
      return handleStock1981Count(body);
    }
    if (body.action === "stock1981Labels") {
      return handleStock1981Labels(body);
    }
    if (body.action === "updateCost") {
      return handleUpdateCost(body);
    }
    if (body.action === "applyCodeReplacements") {
      return jsonResponse(applyCodeReplacements());
    }
    if (body.action === "markLabelsPrinted") {
      return handleMarkLabelsPrinted(body);
    }
    if (body.action === "recvArrival") {
      return handleRecvArrival(body);
    }
    if (body.action === "recvUnit") {
      return handleRecvUnit(body);
    }
    if (body.action === "recvSetDate") {
      return handleRecvSetDate(body);
    }
    if (body.action === "importZoneStock") {
      return handleImportZoneStock(body);
    }
    // Default action: add/update a single product's barcode (existing behavior)
    if (!body.name || !body.barcode) {
      return jsonResponse({ error: "Missing name or barcode" }, 400);
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ error: "Sheet '" + SHEET_NAME + "' not found" }, 500);
    }
    const lastRow = sheet.getLastRow();
    const targetName = normalizeNameForMatch(body.name);
    // 1. Try to find an existing row with this name and an empty barcode.
    if (lastRow >= FIRST_DATA_ROW) {
      const numRows = lastRow - FIRST_DATA_ROW + 1;
      const names = sheet.getRange(FIRST_DATA_ROW, 2, numRows, 1).getValues(); // col B
      const barcodes = sheet.getRange(FIRST_DATA_ROW, 8, numRows, 1).getValues(); // col H
      for (let i = 0; i < names.length; i++) {
        const cellName = names[i][0];
        const cellBarcode = barcodes[i][0];
        if (!cellName) continue;
        if (normalizeNameForMatch(String(cellName)) === targetName) {
          const row = FIRST_DATA_ROW + i;
          if (!cellBarcode || String(cellBarcode).trim() === "") {
            const bcCell = sheet.getRange(row, 8);
            bcCell.setNumberFormat("@"); // force plain text so leading zeros survive
            bcCell.setValue(body.barcode); // H: Barcode
            if (body.qty) {
              const subCatCell = sheet.getRange(row, 5);
              if (!subCatCell.getValue()) subCatCell.setValue(body.qty);
            }
            return jsonResponse({ success: true, action: "updated_existing_row", row: row });
          } else {
            // Row exists and ALREADY has a different barcode — don't overwrite silently.
            return jsonResponse({
              success: false,
              error: "A product with this name already has barcode " + cellBarcode,
              existingBarcode: normalizeBarcode(cellBarcode)
            }, 409);
          }
        }
      }
    }
    // 2. No existing row found — append a new one.
    let maxNum = 0;
    if (lastRow >= FIRST_DATA_ROW) {
      const numRows = lastRow - FIRST_DATA_ROW + 1;
      const nums = sheet.getRange(FIRST_DATA_ROW, 1, numRows, 1).getValues();
      nums.forEach(r => {
        const n = parseInt(r[0], 10);
        if (!isNaN(n) && n > maxNum) maxNum = n;
      });
    }
    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1).setValue(maxNum + 1);           // A: #
    sheet.getRange(newRow, 2).setValue(body.name);            // B: Product Name
    sheet.getRange(newRow, 4).setValue(body.category || "Cosmetics"); // D: Category
    if (body.qty) sheet.getRange(newRow, 5).setValue(body.qty); // E: Sub-Category (qty convention)
    const newBcCell = sheet.getRange(newRow, 8);
    newBcCell.setNumberFormat("@"); // force plain text so leading zeros survive
    newBcCell.setValue(body.barcode);                          // H: Barcode
    return jsonResponse({ success: true, action: "added_new_row", row: newRow });
  } catch (err) {
    return jsonResponse({ error: err.message }, 500);
  }
}
function normalizeNameForMatch(s) {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
/**
 * Normalizes a barcode cell value to a clean string.
 * Handles cases where Sheets stored the barcode as a number
 * (e.g. 667558209752.0 -> "667558209752").
 */
function normalizeBarcode(value) {
  let s;
  if (typeof value === "number") {
    s = String(Math.trunc(value));
  } else {
    s = String(value).trim();
  }
  // Self-heal our own generated codes: they're always 13 digits starting
  // with "0610000...". If Sheets stored the cell as a number, the leading
  // zero gets silently dropped, leaving a 12-digit "610000..." value.
  if (s.length === 12 && s.startsWith("610000")) {
    s = "0" + s;
  }
  return s;
}
// Comparison key for barcodes: same as normalizeBarcode() but with leading
// zeros removed, so "010181145551" and "10181145551" count as the SAME product.
// Scanners and labels often add or drop a leading zero (UPC-A vs EAN-13).
// Use this ONLY to compare/look up — what gets stored stays as typed.
function barcodeKey(value) {
  const s = normalizeBarcode(value);
  const t = s.replace(/^0+/, "");
  return t || s;
}
function jsonResponse(obj, statusCode) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
// ══════════════════════════════════════════════════════════════════════════
//  CLIENTS TAB — shared customer database across both stores
// ══════════════════════════════════════════════════════════════════════════
/**
 * Ensures the Clients tab exists with the correct headers.
 * Called automatically before any read/write on the Clients tab.
 * Columns: Phone | Name | Store | First Visit | Last Visit | Last Purchase | Total Visits
 */
function ensureClientsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Clients')) {
    const ws = ss.insertSheet('Clients');
    const headers = ['Phone','Name','Store','First Visit','Last Visit','Last Purchase','Total Visits'];
    ws.appendRow(headers);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    ws.setColumnWidth(1, 160);
    ws.setColumnWidth(2, 200);
    ws.setColumnWidth(3, 120);
    ws.setColumnWidth(4, 130);
    ws.setColumnWidth(5, 130);
    ws.setColumnWidth(6, 320);
    ws.setColumnWidth(7, 100);
  }
}
/**
 * Returns all customers as JSON.
 * Called by the Scanner app to pre-fill known customers.
 */
function handleGetCustomers() {
  ensureClientsTab();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('Clients');
  const lastRow = ws.getLastRow();
  if (lastRow < 2) return jsonResponse({ customers: [] });
  const data = ws.getRange(2, 1, lastRow - 1, 7).getValues();
  const customers = data
    .filter(r => r[0])
    .map(r => ({
      phone:       String(r[0]).trim(),
      name:        String(r[1] || '').trim(),
      store:       String(r[2] || '').trim(),
      firstVisit:  r[3] ? Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      lastVisit:   r[4] ? Utilities.formatDate(new Date(r[4]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
      lastPurchase:String(r[5] || '').trim(),
      totalVisits: Number(r[6]) || 1,
    }));
  return jsonResponse({ customers, count: customers.length });
}
/**
 * Saves or updates a customer record matched by phone number.
 * If the phone already exists, updates name (if provided), last visit,
 * last purchase, and increments visit count.
 * If new, appends a fresh row.
 */
function handleSaveCustomer(body) {
  ensureClientsTab();
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const ws  = ss.getSheetByName('Clients');
  const now = new Date();
  const tz  = Session.getScriptTimeZone();
  const phone       = String(body.phone       || '').replace(/\s/g, '').trim();
  const name        = String(body.name        || '').trim();
  const store       = String(body.store       || '').trim();
  const lastPurchase= String(body.lastPurchase|| '').trim();
  if (!phone) return jsonResponse({ success: false, error: 'Phone required' });
  const lastRow = ws.getLastRow();
  if (lastRow >= 2) {
    const phones = ws.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < phones.length; i++) {
      if (String(phones[i][0]).replace(/\s/g,'').trim() === phone) {
        const row = i + 2;
        if (name) ws.getRange(row, 2).setValue(name);
        ws.getRange(row, 5).setValue(Utilities.formatDate(now, tz, 'dd/MM/yyyy'));
        if (lastPurchase) ws.getRange(row, 6).setValue(lastPurchase);
        const prev = ws.getRange(row, 7).getValue();
        ws.getRange(row, 7).setValue((Number(prev) || 1) + 1);
        return jsonResponse({ success: true, action: 'updated', phone });
      }
    }
  }
  // New customer
  ws.appendRow([
    phone,
    name,
    store,
    Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
    Utilities.formatDate(now, tz, 'dd/MM/yyyy'),
    lastPurchase,
    1,
  ]);
  return jsonResponse({ success: true, action: 'created', phone });
}
// ══════════════════════════════════════════════════════════════════════════
//  STOCK LEDGER — nightly sold/received movements + running "Qty On Hand"
// ══════════════════════════════════════════════════════════════════════════
//
// Design notes (why it works this way):
//   - "Qty Counted" (col E on each location tab) is a periodic PHYSICAL
//     COUNT snapshot — it gets overwritten each time you run the Counter
//     app, on purpose, because that's its job: tell you what's really on
//     the shelf on inventory day.
//   - "Qty On Hand" is a different, NEW thing: a running number that only
//     this Stock Ledger section touches, moving up and down every night
//     as things sell or arrive. It lives on the same location tabs
//     (Store 1 / Store 2 / Warehouse), matched by Barcode (col M), same
//     as Qty Counted — but it's its own column, added automatically the
//     first time it's needed (see getOrCreateColumnByHeader below), so
//     you never have to edit the sheet by hand.
//   - Nothing is ever silently overwritten. Every movement — a sale, a
//     restock, a new product, or a reset from a physical count — appends
//     one row to the "Stock Ledger" tab (auto-created) with a timestamp,
//     so you always have a permanent trail of what happened and when.
//   - Because Qty On Hand starts blank (=0) until you seed it, drift is
//     expected until you run a physical count and reset it — that's what
//     handleSeedQtyOnHand is for. Think of physical counts as the
//     periodic "trust but verify" step for the nightly ledger.
const LEDGER_SHEET_NAME = "Stock Ledger";
const QTY_ON_HAND_HEADER = "Qty On Hand";
const LOCATION_NAME_COL = 2;      // B — Product Name (mirrors Product Master)
const LOCATION_BARCODE_COL = 13;  // M — Barcode (same column handleUpdateCounts already matches on)
const LOCATION_COUNTED_COL = 5;   // E — Qty Counted (same column handleUpdateCounts already writes)
const VALID_LOCATIONS = ["Store 1", "Store 2", "Warehouse"];
function getLocationSheetOrNull(location) {
  if (VALID_LOCATIONS.indexOf(location) === -1) return null;
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(location);
}

// ══════════════════════════════════════════════════════════════════════════
//  SORT BY COUNTED STATUS — counted items float to the top of the list
// ══════════════════════════════════════════════════════════════════════════
// Reorders a location tab's data rows (row FIRST_DATA_ROW downward) so
// rows with a Status (col G) already set — "✅ Counted" or "⚠️ Zero" —
// appear above rows that have never been counted (Status still blank).
// Within each group, original relative order is kept (stable sort) —
// nothing about a row's own content changes, only its position.
//
// Called automatically at the end of every handleUpdateCounts run (so the
// list re-sorts itself right after "Envoyer au Sheet" in the Counter
// app), and also runnable by hand any time via sortAllLocationsByCountedStatus
// below — e.g. after editing counts directly in the sheet.
//
// Moves every column present on the row, not just B/E/G/M, so extra
// columns — including the self-created "Qty On Hand" — travel with their
// row intact. The Barcode column's Plain Text format is re-applied after
// the move, since re-ordering rows overwrites cell formatting along with
// values, and a re-formatted barcode cell is exactly what normalizeBarcode
// above exists to guard against.
function sortLocationByCountedStatus(location) {
  const sheet = getLocationSheetOrNull(location);
  if (!sheet) return { success: false, error: "Unknown location: " + location };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < FIRST_DATA_ROW + 1) return { success: true, sorted: 0, location: location }; // 0 or 1 row — nothing to sort

  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const range = sheet.getRange(FIRST_DATA_ROW, 1, numRows, lastCol);
  const values = range.getValues();

  const STATUS_COL_INDEX = 6; // G, 0-indexed within the row array
  const counted = [];
  const notCounted = [];
  values.forEach(row => {
    const status = row[STATUS_COL_INDEX];
    if (status !== "" && status !== null && status !== undefined) counted.push(row);
    else notCounted.push(row);
  });
  const sortedRows = counted.concat(notCounted);

  range.setValues(sortedRows);
  sheet.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, numRows, 1).setNumberFormat('@');

  return { success: true, sorted: counted.length, location: location };
}

// Manual entry point — select this in the Apps Script editor's function
// dropdown and click Run any time you want to re-sort by hand (e.g. after
// editing counts directly in the sheet rather than through the Counter
// app, which already auto-sorts after every "Envoyer au Sheet").
function sortAllLocationsByCountedStatus() {
  VALID_LOCATIONS.forEach(loc => {
    const result = sortLocationByCountedStatus(loc);
    Logger.log(loc + ': ' + JSON.stringify(result));
  });
}
/**
 * Finds a column by its header text in the given header row, creating it
 * (appended right after the last used column) if it doesn't exist yet.
 * Lets "Qty On Hand" self-install on first use instead of requiring you
 * to manually edit every location tab.
 */
function getOrCreateColumnByHeader(sheet, headerRow, headerName) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim() === headerName) return i + 1;
  }
  const newCol = lastCol + 1;
  sheet.getRange(headerRow, newCol).setValue(headerName);
  return newCol;
}
function ensureLedgerTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(LEDGER_SHEET_NAME);
    const headers = ['Timestamp','Location','Barcode','Product','Type','Delta','Qty After','Note'];
    ws.appendRow(headers);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    ws.setColumnWidth(1, 140);
    ws.setColumnWidth(4, 220);
    ws.setColumnWidth(8, 260);
  }
  return ws;
}
function appendLedgerRows(rows) {
  if (!rows.length) return;
  const ledger = ensureLedgerTab();
  ledger.getRange(ledger.getLastRow() + 1, 1, rows.length, 8).setValues(rows);
}
/**
 * handleApplyStockMovements — called by the Stock Ledger app at the end
 * of the night. Applies a batch of +/- adjustments to "Qty On Hand" for
 * one location, and logs one row per movement to "Stock Ledger".
 *
 * Expects:
 *   {
 *     "key": "...",
 *     "action": "applyStockMovements",
 *     "location": "Store 1" | "Store 2" | "Warehouse",
 *     "movements": [
 *       { "barcode": "...", "delta": -3, "type": "sold", "note": "" },
 *       { "barcode": "...", "delta": 12, "type": "restock", "note": "" }
 *     ]
 *   }
 * Negative delta = stock going down (sold, damaged, etc). Positive =
 * stock going up (restocked). The caller (the app) decides the sign;
 * this function just applies it and records it.
 *
 * Barcodes that don't exist in the location tab are skipped and
 * reported back in "notFound" — never silently dropped — same
 * philosophy as handleUpdateCounts.
 */
function handleApplyStockMovements(body) {
  const sheet = getLocationSheetOrNull(body.location);
  if (!sheet) return jsonResponse({ success: false, error: "Unknown location: " + body.location }, 400);
  if (!Array.isArray(body.movements) || !body.movements.length) {
    return jsonResponse({ success: false, error: "No movements provided" }, 400);
  }
  body.movements = body.movements.map(m => Object.assign({}, m, { barcode: resolveAlias_(m.barcode) }));
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) {
    return jsonResponse({ success: false, error: "Destination sheet is empty" });
  }
  const qtyCol = getOrCreateColumnByHeader(sheet, HEADER_ROW, QTY_ON_HAND_HEADER);
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const barcodes = sheet.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, numRows, 1).getValues();
  const names = sheet.getRange(FIRST_DATA_ROW, LOCATION_NAME_COL, numRows, 1).getValues();
  const rowByBarcode = {};
  for (let i = 0; i < barcodes.length; i++) {
    const bc = barcodeKey(barcodes[i][0]);
    if (bc) rowByBarcode[bc] = FIRST_DATA_ROW + i;
  }
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  let applied = 0;
  const notFound = [];
  const ledgerRows = [];
  body.movements.forEach(m => {
    const bc = normalizeBarcode(m.barcode);
    const row = rowByBarcode[barcodeKey(m.barcode)];
    if (!row) { notFound.push(bc); return; }
    const cell = sheet.getRange(row, qtyCol);
    const current = Number(cell.getValue()) || 0;
    const delta = Number(m.delta) || 0;
    const next = current + delta;
    cell.setValue(next);
    const name = names[row - FIRST_DATA_ROW][0] || bc;
    ledgerRows.push([timestamp, body.location, bc, name, m.type || 'adjustment', delta, next, m.note || '']);
    applied++;
  });
  appendLedgerRows(ledgerRows);

  // Cascade the same movements down to zone level. This is what makes a
  // sale decrement "the specific zone the item belongs to" (wishlist item
  // 1) rather than just the store-wide total above. Barcodes with no Zone
  // Stock rows yet (never counted into a zone) are silently skipped here —
  // the location-level number above already applied correctly either way.
  try { applyZoneMovements(body.location, body.movements); }
  catch (e) { Logger.log('Zone movement cascade skipped: ' + e); }

  return jsonResponse({ success: true, applied: applied, notFound: notFound, location: body.location });
}
/**
 * handleSeedQtyOnHand — resets "Qty On Hand" to match the most recent
 * physical count ("Qty Counted") for a location. Meant to be called
 * right after finishing a Counter session, so the nightly ledger starts
 * from a number you just verified by hand instead of drifting forever.
 * Logs one "recount-reset" row per changed product, so the reset itself
 * is part of the permanent history, not an invisible jump.
 *
 * Expects: { key, action: "seedQtyOnHand", location }
 */
function handleSeedQtyOnHand(body) {
  const sheet = getLocationSheetOrNull(body.location);
  if (!sheet) return jsonResponse({ success: false, error: "Unknown location: " + body.location }, 400);
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return jsonResponse({ success: false, error: "Destination sheet is empty" });
  const qtyCol = getOrCreateColumnByHeader(sheet, HEADER_ROW, QTY_ON_HAND_HEADER);
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const barcodes = sheet.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, numRows, 1).getValues();
  const names = sheet.getRange(FIRST_DATA_ROW, LOCATION_NAME_COL, numRows, 1).getValues();
  const counted = sheet.getRange(FIRST_DATA_ROW, LOCATION_COUNTED_COL, numRows, 1).getValues();
  const onHandRange = sheet.getRange(FIRST_DATA_ROW, qtyCol, numRows, 1);
  const onHand = onHandRange.getValues();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const ledgerRows = [];
  let changed = 0;
  for (let i = 0; i < numRows; i++) {
    const newQty = Number(counted[i][0]) || 0;
    const oldQty = Number(onHand[i][0]) || 0;
    if (newQty === oldQty) continue;
    onHand[i][0] = newQty;
    const bc = normalizeBarcode(barcodes[i][0]);
    const name = names[i][0] || bc;
    ledgerRows.push([timestamp, body.location, bc, name, 'recount-reset', newQty - oldQty, newQty, 'Reset from physical count']);
    changed++;
  }
  onHandRange.setValues(onHand);
  appendLedgerRows(ledgerRows);

  // Same reset, one level down: every Zone Stock row for this location gets
  // its Zone Qty On Hand pulled up to match Zone Qty Counted. A zone that
  // was recounted to zero (barcode no longer scanned in it this session)
  // correctly zeroes out here too, since applyZoneCounts (below) already
  // set its Zone Qty Counted to 0 when it wasn't in the latest zone payload.
  let zonesChanged = 0;
  try { zonesChanged = seedZoneStockFromCounts(body.location); }
  catch (e) { Logger.log('Zone seed skipped: ' + e); }

  return jsonResponse({ success: true, changed: changed, zonesChanged: zonesChanged, location: body.location });
}
/**
 * addProductToMaster — shared "append or fill in a product on Product
 * Master" logic. Mirrors doPost's original default-action behavior, but
 * factored out into its own function so handleAddNewProductWithStock can
 * reuse it without duplicating the row-matching logic. The original
 * inline version in doPost is left untouched on purpose (don't disturb
 * a working path — this is a separate copy used only by the new flow).
 */
function addProductToMaster(name, barcode, category) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return { success: false, error: "Sheet '" + SHEET_NAME + "' not found" };
  const lastRow = sheet.getLastRow();
  const targetName = normalizeNameForMatch(name);
  if (lastRow >= FIRST_DATA_ROW) {
    const numRows = lastRow - FIRST_DATA_ROW + 1;
    const names = sheet.getRange(FIRST_DATA_ROW, 2, numRows, 1).getValues();
    const barcodes = sheet.getRange(FIRST_DATA_ROW, 8, numRows, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      const cellName = names[i][0];
      if (!cellName) continue;
      if (normalizeNameForMatch(String(cellName)) === targetName) {
        const row = FIRST_DATA_ROW + i;
        const cellBarcode = barcodes[i][0];
        if (!cellBarcode || String(cellBarcode).trim() === "") {
          const bcCell = sheet.getRange(row, 8);
          bcCell.setNumberFormat("@");
          bcCell.setValue(barcode);
          return { success: true, action: "updated_existing_row", row: row };
        }
        return { success: false, error: "Already has barcode " + cellBarcode, existingBarcode: normalizeBarcode(cellBarcode) };
      }
    }
  }
  let maxNum = 0;
  if (lastRow >= FIRST_DATA_ROW) {
    const numRows = lastRow - FIRST_DATA_ROW + 1;
    const nums = sheet.getRange(FIRST_DATA_ROW, 1, numRows, 1).getValues();
    nums.forEach(r => { const n = parseInt(r[0], 10); if (!isNaN(n) && n > maxNum) maxNum = n; });
  }
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1).setValue(maxNum + 1);
  sheet.getRange(newRow, 2).setValue(name);
  sheet.getRange(newRow, 4).setValue(category || "Cosmetics");
  const newBcCell = sheet.getRange(newRow, 8);
  newBcCell.setNumberFormat("@");
  newBcCell.setValue(barcode);
  return { success: true, action: "added_new_row", row: newRow };
}
/**
 * handleAddNewProductWithStock — for products that don't exist anywhere
 * yet. Adds the product to "Product Master" (so it shows up in the
 * Scanner/Generator/Counter product list going forward) AND appends a
 * matching row to the given location tab with a starting Qty On Hand,
 * so it's immediately trackable by tonight's ledger too.
 *
 * Expects:
 *   {
 *     key, action: "addNewProductWithStock",
 *     name, barcode, category, location, qty
 *   }
 */
function handleAddNewProductWithStock(body) {
  if (!body.name || !body.barcode || !body.location) {
    return jsonResponse({ success: false, error: "Missing name, barcode, or location" }, 400);
  }
  const locSheet = getLocationSheetOrNull(body.location);
  if (!locSheet) return jsonResponse({ success: false, error: "Unknown location: " + body.location }, 400);
  const masterResult = addProductToMaster(body.name, body.barcode, body.category);
  const newRow = locSheet.getLastRow() + 1;
  locSheet.getRange(newRow, LOCATION_NAME_COL).setValue(body.name);
  const bcCell = locSheet.getRange(newRow, LOCATION_BARCODE_COL);
  bcCell.setNumberFormat("@");
  bcCell.setValue(body.barcode);
  const qtyCol = getOrCreateColumnByHeader(locSheet, HEADER_ROW, QTY_ON_HAND_HEADER);
  const startQty = Number(body.qty) || 0;
  locSheet.getRange(newRow, qtyCol).setValue(startQty);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  appendLedgerRows([[timestamp, body.location, normalizeBarcode(body.barcode), body.name, 'new-product', startQty, startQty, 'Added via Stock Ledger app']]);
  return jsonResponse({ success: true, masterResult: masterResult, location: body.location, row: newRow });
}

// ══════════════════════════════════════════════════════════════════════════
//  BARCODE ALIASES — wishlist item 5's "Nido case": a manufacturer changes
//  a product's barcode while old-barcode stock is still on the shelf.
//  Kept deliberately simple, per the ask: the new barcode gets its OWN
//  Product Master row and its OWN stock tracking (old and new barcodes
//  each keep counting independently — nothing is merged or migrated), and
//  a row on this tab just records that the two are the same product, so
//  nobody looking at two different-barcode "Nido" rows is left confused
//  about whether that's a duplicate or a barcode change.
// ══════════════════════════════════════════════════════════════════════════
const ALIASES_SHEET_NAME = "Barcode Aliases";
const ALIASES_HEADERS = ['Date', 'Product Name', 'Old Barcode', 'New Barcode', 'Note'];

function ensureAliasesTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(ALIASES_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(ALIASES_SHEET_NAME);
    ws.appendRow(ALIASES_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, ALIASES_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(2, 220);
  }
  return ws;
}

/**
 * handleAddAliasedBarcode — called by the Counter app when the person
 * confirms a name collision is the SAME product under a NEW barcode.
 * Adds the new barcode as its own Product Master row (so it's recognized
 * everywhere going forward, same as any other product) and its own
 * starting stock row on the given location tab, then logs the link.
 *
 * Expects: { key, action: "addAliasedBarcode", name, oldBarcode,
 *            newBarcode, category, location, qty }
 */
function handleAddAliasedBarcode(body) {
  if (!body.name || !body.oldBarcode || !body.newBarcode || !body.location) {
    return jsonResponse({ success: false, error: "Missing name, oldBarcode, newBarcode, or location" }, 400);
  }
  const locSheet = getLocationSheetOrNull(body.location);
  if (!locSheet) return jsonResponse({ success: false, error: "Unknown location: " + body.location }, 400);

  const newBc = normalizeBarcode(body.newBarcode);
  const oldBc = normalizeBarcode(body.oldBarcode);

  // Always append a fresh Product Master row for the new barcode — unlike
  // addProductToMaster, this is expected to collide on name, so it doesn't
  // try to reuse the old row (that one keeps its own barcode untouched).
  const masterSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  let masterRow = null;
  if (masterSheet) {
    let maxNum = 0;
    const lastRow = masterSheet.getLastRow();
    if (lastRow >= FIRST_DATA_ROW) {
      const numRows = lastRow - FIRST_DATA_ROW + 1;
      const nums = masterSheet.getRange(FIRST_DATA_ROW, 1, numRows, 1).getValues();
      nums.forEach(r => { const n = parseInt(r[0], 10); if (!isNaN(n) && n > maxNum) maxNum = n; });
    }
    masterRow = masterSheet.getLastRow() + 1;
    masterSheet.getRange(masterRow, 1).setValue(maxNum + 1);
    masterSheet.getRange(masterRow, 2).setValue(body.name);
    masterSheet.getRange(masterRow, 4).setValue(body.category || "Cosmetics");
    const bcCell = masterSheet.getRange(masterRow, 8);
    bcCell.setNumberFormat("@");
    bcCell.setValue(newBc);
  }

  const newRow = locSheet.getLastRow() + 1;
  locSheet.getRange(newRow, LOCATION_NAME_COL).setValue(body.name);
  const bcCell = locSheet.getRange(newRow, LOCATION_BARCODE_COL);
  bcCell.setNumberFormat("@");
  bcCell.setValue(newBc);
  const qtyCol = getOrCreateColumnByHeader(locSheet, HEADER_ROW, QTY_ON_HAND_HEADER);
  const startQty = Number(body.qty) || 0;
  locSheet.getRange(newRow, qtyCol).setValue(startQty);

  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  appendLedgerRows([[timestamp, body.location, newBc, body.name, 'new-product-aliased', startQty, startQty,
    'Nouveau code-barres pour "' + body.name + '" (ancien : ' + oldBc + ')']]);

  const aliasSheet = ensureAliasesTab();
  aliasSheet.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
    body.name, oldBc, newBc, 'Lié depuis l\'app Comptage Inventaire']);

  return jsonResponse({ success: true, location: body.location, row: newRow, masterRow: masterRow });
}

// ══════════════════════════════════════════════════════════════════════════
//  PRICE EDITS — wishlist item 4: set/manage price from inside the Scanner
//  and Counter apps (Product Master stays the single source of truth), so
//  it syncs to both stores automatically via the existing 15-minute
//  Product Sync pull. Per 1981's call: no PIN prompt on every single edit —
//  just a log entry, attributed to whatever name/session was entered when
//  that work session started (the Counter app's session name), so it's
//  still clear afterward who was working when a price changed, without
//  interrupting counting to re-authenticate on every price fix.
// ══════════════════════════════════════════════════════════════════════════
const PRICE_EDITS_SHEET_NAME = "Price Edits";
const PRICE_EDITS_HEADERS = ['Timestamp', 'Barcode', 'Product', 'Old Price', 'New Price', 'Location', 'Session', 'App'];
const MASTER_PRICE_COL = 10; // J: Price (CFA)

function ensurePriceEditsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(PRICE_EDITS_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(PRICE_EDITS_SHEET_NAME);
    ws.appendRow(PRICE_EDITS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, PRICE_EDITS_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(3, 220);
  }
  return ws;
}

/**
 * handleUpdatePrice — writes a new price to Product Master (col J) for the
 * given barcode and logs the change. Expects:
 *   { key, action: "updatePrice", barcode, newPrice, location, session, app }
 * `session` is free text (whatever the Counter app's session name was, or
 * the Scanner's logged-in employee) — just an attribution string, no PIN.
 */
function handleUpdatePrice(body) {
  const bc = normalizeBarcode(body.barcode);
  const bk = barcodeKey(body.barcode);
  if (!bc || body.newPrice === undefined || body.newPrice === null || body.newPrice === '') {
    return jsonResponse({ success: false, error: "Missing barcode or newPrice" }, 400);
  }
  const newPrice = Number(body.newPrice);
  if (isNaN(newPrice)) return jsonResponse({ success: false, error: "newPrice is not a number" }, 400);

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResponse({ success: false, error: "Sheet '" + SHEET_NAME + "' not found" }, 500);
  const lastRow = sheet.getLastRow();
  if (lastRow < FIRST_DATA_ROW) return jsonResponse({ success: false, error: "Product Master is empty" });

  const numRows = lastRow - FIRST_DATA_ROW + 1;
  const barcodes = sheet.getRange(FIRST_DATA_ROW, 8, numRows, 1).getValues(); // H: Barcode
  let row = null;
  for (let i = 0; i < barcodes.length; i++) {
    if (barcodeKey(barcodes[i][0]) === bk) { row = FIRST_DATA_ROW + i; break; }
  }
  if (!row) return jsonResponse({ success: false, error: "Barcode not found in Product Master: " + bc });

  const nameCell = sheet.getRange(row, 2).getValue();
  const priceCell = sheet.getRange(row, MASTER_PRICE_COL);
  const oldPrice = Number(priceCell.getValue()) || 0;
  priceCell.setValue(newPrice);

  const ledger = ensurePriceEditsTab();
  ledger.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
    bc, nameCell || bc, oldPrice, newPrice, body.location || '', body.session || '', body.app || ''
  ]);

  return jsonResponse({ success: true, barcode: bc, name: nameCell, oldPrice: oldPrice, newPrice: newPrice });
}

// ══════════════════════════════════════════════════════════════════════════
//  STORE PRICES — a price per store, plus tag-change tasks.
//  Tab "Store Prices": Barcode | Product | Store 1 | Store 2 | Updated | By.
//  Effective price of a store = its own price if set, else Product Master
//  col J (the "liste principale", which is also what the Counter writes).
//  Tab "Price Tasks": one row per tag to change at a store when the admin
//  changes a price (Status open / done).
//  setStorePrice mode 'session' (employees): refused when the store already
//  has a price for that product. mode 'admin': always allowed, creates tasks.
// ══════════════════════════════════════════════════════════════════════════
const STORE_PRICES_SHEET = 'Store Prices';
const STORE_PRICES_HEADERS = ['Barcode', 'Product', 'Store 1', 'Store 2', 'Updated', 'By'];
const PRICE_TASKS_SHEET = 'Price Tasks';
const PRICE_TASKS_HEADERS = ['ID', 'Time', 'Location', 'Barcode', 'Product', 'Old Price', 'New Price', 'By', 'Status', 'Done By', 'Done At'];
const SP_COL = { 'Store 1': 3, 'Store 2': 4 };

function _spSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(STORE_PRICES_SHEET);
  if (!ws) {
    ws = ss.insertSheet(STORE_PRICES_SHEET);
    ws.appendRow(STORE_PRICES_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, STORE_PRICES_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(2, 220);
  }
  return ws;
}
function _ptSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(PRICE_TASKS_SHEET);
  if (!ws) {
    ws = ss.insertSheet(PRICE_TASKS_SHEET);
    ws.appendRow(PRICE_TASKS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, PRICE_TASKS_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(5, 220);
  }
  return ws;
}
function _num(v) { return (v === '' || v === null || v === undefined || isNaN(Number(v))) ? null : Number(v); }

// barcodeKey -> { row, s1, s2 }
function _spMap(ws) {
  const out = {};
  const last = ws.getLastRow();
  if (last < 2) return out;
  ws.getRange(2, 1, last - 1, 4).getValues().forEach((r, i) => {
    if (!r[0]) return;
    out[barcodeKey(r[0])] = { row: i + 2, s1: _num(r[2]), s2: _num(r[3]) };
  });
  return out;
}
// barcodeKey -> { row, name, price }  from Product Master
function _masterMap() {
  const out = {};
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return out;
  const last = sheet.getLastRow();
  if (last < FIRST_DATA_ROW) return out;
  sheet.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, 10).getValues().forEach((r, i) => {
    if (!r[7] || !r[1]) return;
    out[barcodeKey(r[7])] = { row: FIRST_DATA_ROW + i, name: String(r[1]).trim(), price: _num(r[9]), barcode: normalizeBarcode(r[7]), category: r[3] ? String(r[3]).trim() : '' };
  });
  return out;
}
// Effective price for a location (Warehouse and unknown fall back to the main list).
function _effectivePrice(loc, master, sp) {
  const col = loc === 'Store 1' ? 's1' : (loc === 'Store 2' ? 's2' : '');
  if (col && sp && sp[col] != null) return sp[col];
  return master ? master.price : null;
}

/** getPriceList — for the Admin list and the price session. */
function handleGetPriceList(withCost) {
  const mm = _masterMap(), sp = _spMap(_spSheet());
  const aliases = loadAliasMap_();
  const cost = withCost ? _costMap_() : {};
  const products = [];
  Object.keys(mm).forEach(k => {
    if (aliases[k]) return; // old made-up code, replaced by its real barcode
    const m = mm[k], s = sp[k];
    const p = { barcode: m.barcode, name: m.name, category: m.category, m: m.price, s1: s ? s.s1 : null, s2: s ? s.s2 : null };
    if (withCost) p.cost = cost[k] != null ? cost[k] : null;
    products.push(p);
  });
  products.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return jsonResponse({ success: true, products: products, generatedAt: new Date().toISOString() });
}

function _ptOpenFor(rows, loc, key) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r[8] === 'open' && r[2] === loc && barcodeKey(r[3]) === key) return i;
  }
  return -1;
}

// Creates / updates / cancels the "change the tag" task for one store.
function _ptRecord(loc, bc, name, oldP, newP, by) {
  if (loc !== 'Store 1' && loc !== 'Store 2') return;
  if (oldP == null || oldP === newP) { // nothing to relabel (or back to the original)
    _ptCancel(loc, bc, newP);
    return;
  }
  const ws = _ptSheet();
  const last = ws.getLastRow();
  const rows = last >= 2 ? ws.getRange(2, 1, last - 1, PRICE_TASKS_HEADERS.length).getValues() : [];
  const key = barcodeKey(bc);
  const idx = _ptOpenFor(rows, loc, key);
  const tz = Session.getScriptTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  if (idx >= 0) {
    const rowNum = idx + 2, origOld = _num(rows[idx][5]);
    if (origOld === newP) { ws.getRange(rowNum, 9, 1, 3).setValues([['cancelled', by || '', stamp]]); return; }
    ws.getRange(rowNum, 2).setValue(stamp);
    ws.getRange(rowNum, 7).setValue(newP);
    ws.getRange(rowNum, 8).setValue(by || '');
    return;
  }
  const id = 'PT' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
  const nr = last + 1;
  const rng = ws.getRange(nr, 1, 1, PRICE_TASKS_HEADERS.length);
  rng.setNumberFormat('@');
  rng.setValues([[id, stamp, loc, normalizeBarcode(bc), name, oldP, newP, by || '', 'open', '', '']]);
}
function _ptCancel(loc, bc, newP) {
  const ws = _ptSheet();
  const last = ws.getLastRow();
  if (last < 2) return;
  const rows = ws.getRange(2, 1, last - 1, PRICE_TASKS_HEADERS.length).getValues();
  const idx = _ptOpenFor(rows, loc, barcodeKey(bc));
  if (idx >= 0) ws.getRange(idx + 2, 9).setValue('cancelled');
}

function _spLedger(bc, name, oldP, newP, loc, by, app) {
  try {
    ensurePriceEditsTab().appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      bc, name || bc, oldP == null ? '' : oldP, newP, loc || '', by || '', app || '']);
  } catch (e) {}
}

/**
 * setStorePrice — { key, action, mode: 'session'|'admin', barcode, name?, by,
 *   price (session: for `location`) | prices: {'Store 1': n, 'Store 2': n} (admin) , location }
 */
function handleSetStorePrice(body) {
  const bc = normalizeBarcode(body.barcode || '');
  if (!bc) return jsonResponse({ success: false, error: "Missing barcode" }, 400);
  const mode = body.mode === 'admin' ? 'admin' : 'session';
  const changes = {};
  if (mode === 'admin') {
    const p = body.prices || {};
    ['Store 1', 'Store 2'].forEach(l => { if (p[l] !== undefined && p[l] !== null && p[l] !== '') changes[l] = Number(p[l]); });
  } else if (body.location === 'Store 1' || body.location === 'Store 2') {
    changes[body.location] = Number(body.price);
  }
  const locs = Object.keys(changes);
  if (!locs.length) return jsonResponse({ success: false, error: "Aucun magasin / prix" }, 400);
  for (let i = 0; i < locs.length; i++) {
    const v = changes[locs[i]];
    if (isNaN(v) || v < 0) return jsonResponse({ success: false, error: "Prix invalide" }, 400);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const mm = _masterMap();
    const key = barcodeKey(bc);
    const m = mm[key];
    if (!m) return jsonResponse({ success: false, error: "Barcode not found in Product Master: " + bc });
    const ws = _spSheet();
    const sp = _spMap(ws)[key];
    const by = body.by || '';
    const tz = Session.getScriptTimeZone();
    const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    const applied = {};
    if (mode === 'session') {
      const l = locs[0];
      const cur = sp ? (l === 'Store 1' ? sp.s1 : sp.s2) : null;
      if (cur != null) return jsonResponse({ success: false, error: "already-set", current: cur, location: l });
    }
    let rowNum = sp ? sp.row : -1;
    if (rowNum < 0) {
      rowNum = ws.getLastRow() + 1;
      const rng = ws.getRange(rowNum, 1, 1, 6);
      rng.setNumberFormat('@');
      rng.setValues([[bc, m.name, '', '', stamp, by]]);
      ws.getRange(rowNum, 3, 1, 2).setNumberFormat('0');
    }
    locs.forEach(l => {
      const oldEff = _effectivePrice(l, m, sp);
      const newP = changes[l];
      ws.getRange(rowNum, SP_COL[l]).setValue(newP);
      applied[l] = newP;
      _spLedger(bc, m.name, oldEff, newP, l, by, mode === 'admin' ? 'Admin (liste des prix)' : 'Scanner (session prix)');
      if (mode === 'admin') _ptRecord(l, bc, m.name, oldEff, newP, by);
    });
    ws.getRange(rowNum, 5, 1, 2).setValues([[stamp, by]]);
    // Baseline for products that had no main price at all (keeps sales/other apps working).
    if (m.price == null) {
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME).getRange(m.row, MASTER_PRICE_COL).setValue(changes[locs[0]]);
    }
    return jsonResponse({ success: true, barcode: bc, name: m.name, applied: applied });
  } finally { lock.releaseLock(); }
}

/** copyStorePrices — { from: 'Main'|'Store 1'|'Store 2', to: 'Store 1'|'Store 2', onlyMissing, by } */
function handleCopyStorePrices(body) {
  const from = String(body.from || ''), to = String(body.to || '');
  if (['Main', 'Store 1', 'Store 2'].indexOf(from) === -1 || ['Store 1', 'Store 2'].indexOf(to) === -1 || from === to) {
    return jsonResponse({ success: false, error: "Source / destination invalide" }, 400);
  }
  const onlyMissing = body.onlyMissing !== false;
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const mm = _masterMap();
    const ws = _spSheet();
    const spm = _spMap(ws);
    const tz = Session.getScriptTimeZone();
    const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    let copied = 0, skipped = 0;
    Object.keys(mm).forEach(k => {
      const m = mm[k], sp = spm[k];
      const srcP = from === 'Main' ? m.price : (sp ? (from === 'Store 1' ? sp.s1 : sp.s2) : null);
      if (srcP == null) return;
      const curP = sp ? (to === 'Store 1' ? sp.s1 : sp.s2) : null;
      if (onlyMissing && curP != null) { skipped++; return; }
      if (curP === srcP) { skipped++; return; }
      let rowNum = sp ? sp.row : -1;
      if (rowNum < 0) {
        rowNum = ws.getLastRow() + 1;
        ws.getRange(rowNum, 1, 1, 6).setNumberFormat('@').setValues([[m.barcode, m.name, '', '', stamp, body.by || '']]);
        spm[k] = { row: rowNum, s1: null, s2: null };
      }
      ws.getRange(rowNum, SP_COL[to]).setValue(srcP);
      ws.getRange(rowNum, 5, 1, 2).setValues([[stamp, body.by || '']]);
      if (to === 'Store 1') spm[k].s1 = srcP; else spm[k].s2 = srcP;
      copied++;
    });
    return jsonResponse({ success: true, copied: copied, skipped: skipped });
  } finally { lock.releaseLock(); }
}

/** getPriceTasks — open tasks (optionally for one location) + the 40 most recently done. */
function handleGetPriceTasks(location) {
  const ws = _ptSheet();
  const last = ws.getLastRow();
  const open = [], done = [];
  if (last >= 2) {
    ws.getRange(2, 1, last - 1, PRICE_TASKS_HEADERS.length).getValues().forEach(r => {
      if (location && r[2] !== location) return;
      const o = { id: r[0], time: _exStr(r[1]), location: r[2], barcode: String(r[3]), name: r[4], oldPrice: _num(r[5]), newPrice: _num(r[6]), by: r[7], status: r[8], doneBy: r[9], doneAt: _exStr(r[10]) };
      if (r[8] === 'open') open.push(o); else if (r[8] === 'done') done.push(o);
    });
  }
  return jsonResponse({ success: true, open: open, done: done.slice(-40).reverse() });
}
function handleDonePriceTask(body) {
  const id = String(body.id || '');
  if (!id) return jsonResponse({ success: false, error: "Missing id" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ws = _ptSheet();
    const last = ws.getLastRow();
    if (last < 2) return jsonResponse({ success: false, error: "Introuvable" });
    const ids = ws.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === id) {
        const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
        ws.getRange(i + 2, 9, 1, 3).setValues([['done', body.by || '', stamp]]);
        return jsonResponse({ success: true });
      }
    }
    return jsonResponse({ success: false, error: "Introuvable" });
  } finally { lock.releaseLock(); }
}

// ══════════════════════════════════════════════════════════════════════════
//  CONTAINER INTAKE — wishlist item 8: two-phase container arrival. This is
//  ONLY the fast phase-1 log (qty + rough location, no barcode/price) —
//  kept as its own permanent per-shipment record, deliberately NOT merged
//  into Product Master or any location tab. The detailed barcode/price/
//  count pass a day or two later is just an ordinary Counter zone-count
//  session (handleUpdateCounts above); nothing special ties it back to a
//  shipment beyond the person naming that session after it, by eye.
// ══════════════════════════════════════════════════════════════════════════
const CONTAINER_INTAKE_SHEET_NAME = "Container Intake";
const CONTAINER_INTAKE_HEADERS = ['Timestamp', 'Shipment', 'Location', 'Description', 'Qty', 'Rough Location', 'Contenant Vendable', 'Status'];

function ensureContainerIntakeTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(CONTAINER_INTAKE_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(CONTAINER_INTAKE_SHEET_NAME);
    ws.appendRow(CONTAINER_INTAKE_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, CONTAINER_INTAKE_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(4, 240);
    ws.setColumnWidth(6, 200);
  }
  return ws;
}

/**
 * handleLogContainerIntake — appends one row per entry to "Container
 * Intake". Expects:
 *   { key, action: "logContainerIntake", shipment, location,
 *     entries: [{ description, qty, roughLocation, container }, ...] }
 * `container: true` flags stock sitting INSIDE a sellable container
 * itself (e.g. a cooler) — a reminder that it needs emptying before that
 * container can be sold, surfaced here just by being filterable on this
 * tab (kept simple, per the ask — no separate workflow built for it).
 */
function handleLogContainerIntake(body) {
  if (!body.shipment || !Array.isArray(body.entries) || !body.entries.length) {
    return jsonResponse({ success: false, error: "Missing shipment or entries" }, 400);
  }
  const sheet = ensureContainerIntakeTab();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const rows = body.entries.map(e => [
    timestamp, body.shipment, body.location || '', e.description || '', Number(e.qty) || 0,
    e.roughLocation || '', !!e.container, 'En attente du passage détaillé'
  ]);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, CONTAINER_INTAKE_HEADERS.length).setValues(rows);
  return jsonResponse({ success: true, logged: rows.length, shipment: body.shipment });
}

// ══════════════════════════════════════════════════════════════════════════
//  EXPIRY TRACKING — wishlist item 11: items expiring without being
//  noticed in time. Tracked per batch/shipment (varies batch to batch,
//  not one fixed date per product), ties into item 8's shipment records
//  only loosely — by the shipment name matching, typed by eye, same as
//  the rest of this tool's deliberately-simple design.
// ══════════════════════════════════════════════════════════════════════════
const EXPIRY_SHEET_NAME = "Expiry Batches";
const EXPIRY_HEADERS = ['Timestamp', 'Location', 'Zone', 'Barcode', 'Product', 'Qty', 'Expiry Date', 'Shipment', 'Status'];
const EX_COL_LOCATION = 2, EX_COL_ZONE = 3, EX_COL_BARCODE = 4, EX_COL_PRODUCT = 5,
      EX_COL_QTY = 6, EX_COL_EXPIRY = 7, EX_COL_SHIPMENT = 8, EX_COL_STATUS = 9;

function ensureExpiryTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(EXPIRY_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(EXPIRY_SHEET_NAME);
    ws.appendRow(EXPIRY_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, EXPIRY_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(5, 200);
  }
  return ws;
}

// Parses "DD/MM/YYYY" (what the Counter app's prompt collects) into a Date.
function parseFrDate(s) {
  const m = String(s).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

/**
 * handleLogExpiryBatch — records one batch's expiry date. Expects:
 *   { key, action: "logExpiryBatch", barcode, name, qty, location, zone,
 *     expiryDate: "DD/MM/YYYY", shipment }
 */
function handleLogExpiryBatch(body) {
  const bc = normalizeBarcode(body.barcode);
  if (!bc || !body.expiryDate || !body.location) {
    return jsonResponse({ success: false, error: "Missing barcode, expiryDate, or location" }, 400);
  }
  if (!parseFrDate(body.expiryDate)) {
    return jsonResponse({ success: false, error: "expiryDate must be DD/MM/YYYY" }, 400);
  }
  const sheet = ensureExpiryTab();
  sheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
    body.location, body.zone || '', bc, body.name || bc, Number(body.qty) || 0,
    body.expiryDate, body.shipment || '', 'watch'
  ]);
  return jsonResponse({ success: true });
}

/**
 * handleGetExpiryWatchlist — every non-resolved batch for a location,
 * sorted soonest-expiring first, with daysLeft computed relative to today
 * (negative = already past) so the Counter app can color-code urgency
 * without doing its own date math.
 */
function handleGetExpiryWatchlist(location) {
  if (["Store 1", "Store 2", "Warehouse"].indexOf(location) === -1) {
    return jsonResponse({ success: false, error: "Unknown location: " + location }, 400);
  }
  const sheet = ensureExpiryTab();
  const lastRow = sheet.getLastRow();
  const entries = [];
  if (lastRow >= 2) {
    const numRows = lastRow - 1;
    const data = sheet.getRange(2, 1, numRows, EXPIRY_HEADERS.length).getValues();
    const today = new Date(); today.setHours(0, 0, 0, 0);
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (row[EX_COL_LOCATION - 1] !== location) continue;
      { const st = row[EX_COL_STATUS - 1]; if (st && st !== 'watch') continue; }
      const expiry = parseFrDate(_exStr(row[EX_COL_EXPIRY - 1]));
      if (!expiry) continue;
      const daysLeft = Math.round((expiry - today) / 86400000);
      entries.push({
        row: i + 2, barcode: row[EX_COL_BARCODE - 1], name: row[EX_COL_PRODUCT - 1],
        qty: row[EX_COL_QTY - 1], zone: row[EX_COL_ZONE - 1], shipment: row[EX_COL_SHIPMENT - 1],
        expiryDate: _exStr(row[EX_COL_EXPIRY - 1]), daysLeft: daysLeft
      });
    }
    entries.sort((a, b) => a.daysLeft - b.daysLeft);
  }
  return jsonResponse({ success: true, location: location, entries: entries });
}

/** handleResolveExpiryBatch — marks one row 'resolved' so it drops off the watch list. */
function handleResolveExpiryBatch(body) {
  const row = Number(body.row);
  if (!row) return jsonResponse({ success: false, error: "Missing row" }, 400);
  const sheet = ensureExpiryTab();
  if (row < 2 || row > sheet.getLastRow()) return jsonResponse({ success: false, error: "Row out of range" });
  sheet.getRange(row, EX_COL_STATUS).setValue('resolved');
  return jsonResponse({ success: true });
}

// ══════════════════════════════════════════════════════════════════════════
//  EXPIRY v2 — several batches per product, a "to date" to-do list built from
//  stock, and an all-locations overview with value at risk. Everything here is
//  additive: the Counter's logExpiryBatch / getExpiryWatchlist / resolveExpiryBatch
//  keep working. New columns on the Expiry Batches tab: Checked By, Checked At.
//  Status values: 'watch' (active) | 'noexpiry' (product has no date, counts as
//  done) | 'resolved' (dealt with) | 'replaced' (superseded by a newer check).
// ══════════════════════════════════════════════════════════════════════════
const EX_COL_CHECKEDBY = 10, EX_COL_CHECKEDAT = 11;
const EX_ALL_HEADERS = EXPIRY_HEADERS.concat(['Checked By', 'Checked At']);

// Sheets may auto-convert "15/03/2027" into a real Date; accept both.
function _exStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v == null ? '' : v).trim();
}
function _exEnsureCols(ws) {
  if (String(ws.getRange(1, EX_COL_CHECKEDBY).getValue() || '') !== 'Checked By') {
    ws.getRange(1, EX_COL_CHECKEDBY, 1, 2).setValues([['Checked By', 'Checked At']]);
    ws.getRange(1, EX_COL_CHECKEDBY, 1, 2).setFontWeight('bold');
  }
}
function _exIsActive(status) { return !status || status === 'watch'; }
function _exIsDated(status) { return !status || status === 'watch' || status === 'noexpiry'; }

function _exPriceMap() {
  const out = {};
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return out;
    const last = sheet.getLastRow();
    if (last < FIRST_DATA_ROW) return out;
    const data = sheet.getRange(FIRST_DATA_ROW, 1, last - FIRST_DATA_ROW + 1, 10).getValues();
    data.forEach(r => {
      if (!r[7]) return;
      const p = r[9];
      if (p !== '' && p !== null && !isNaN(Number(p))) out[barcodeKey(r[7])] = Number(p);
    });
  } catch (e) { Logger.log('price map skipped: ' + e); }
  return out;
}

/** getExpiryOverview — every active batch at every location, with daysLeft and price. */
function handleGetExpiryOverview() {
  const ws = ensureExpiryTab();
  _exEnsureCols(ws);
  const last = ws.getLastRow();
  const prices = _exPriceMap();
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const batches = [];
  if (last >= 2) {
    const data = ws.getRange(2, 1, last - 1, 16).getValues();
    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      if (!_exIsActive(r[EX_COL_STATUS - 1])) continue;
      const expStr = _exStr(r[EX_COL_EXPIRY - 1]);
      const expiry = parseFrDate(expStr);
      if (!expiry) continue;
      const bc = String(r[EX_COL_BARCODE - 1]);
      batches.push({
        row: i + 2, location: r[EX_COL_LOCATION - 1], zone: String(r[EX_COL_ZONE - 1] || ''),
        barcode: bc, name: r[EX_COL_PRODUCT - 1], qty: Number(r[EX_COL_QTY - 1]) || 0,
        expiryDate: expStr, daysLeft: Math.round((expiry - today) / 86400000),
        checkedBy: r[EX_COL_CHECKEDBY - 1] || '', checkedAt: _exStr(r[EX_COL_CHECKEDAT - 1]),
        action: r[EX_COL_ACTION - 1] || '', actionPrice: r[EX_COL_ACTPRICE - 1] === '' || r[EX_COL_ACTPRICE - 1] == null ? null : Number(r[EX_COL_ACTPRICE - 1]),
        price: prices[barcodeKey(bc)] != null ? prices[barcodeKey(bc)] : null
      });
    }
    batches.sort((a, b) => a.daysLeft - b.daysLeft);
  }
  return jsonResponse({ success: true, batches: batches, generatedAt: new Date().toISOString() });
}

/**
 * getDateTodo — for one location: per zone, which stocked products already have
 * a recorded expiry check and which don't yet. Falls back to the location tab
 * (zone "(sans zone)") when Zone Stock has no rows for that location.
 */
function handleGetDateTodo(location) {
  if (VALID_LOCATIONS.indexOf(location) === -1) {
    return jsonResponse({ success: false, error: "Unknown location: " + location }, 400);
  }
  const ws = ensureExpiryTab();
  _exEnsureCols(ws);
  const last = ws.getLastRow();
  const datedZone = {}, datedAny = {}, lastChecked = {};
  const today = new Date(); today.setHours(0, 0, 0, 0);
  if (last >= 2) {
    const data = ws.getRange(2, 1, last - 1, EX_ALL_HEADERS.length).getValues();
    data.forEach(r => {
      if (r[EX_COL_LOCATION - 1] !== location) return;
      if (!_exIsDated(r[EX_COL_STATUS - 1])) return;
      const k = barcodeKey(r[EX_COL_BARCODE - 1]);
      const z = String(r[EX_COL_ZONE - 1] || '');
      datedZone[k + '|' + z] = true; datedAny[k] = true;
      const ca = parseFrDate(_exStr(r[EX_COL_CHECKEDAT - 1]));
      if (ca && (!lastChecked[z] || ca > lastChecked[z])) lastChecked[z] = ca;
    });
  }
  const zones = {};
  const addItem = (zone, bc, name, qty, isDated) => {
    if (!zones[zone]) zones[zone] = { total: 0, dated: 0, todo: [], daysSince: null };
    zones[zone].total++;
    if (isDated) zones[zone].dated++; else zones[zone].todo.push({ barcode: String(bc), name: String(name || bc), qty: qty });
  };
  const zs = ensureZoneStockTab();
  const zl = zs.getLastRow();
  let usedZoneStock = false;
  if (zl >= 2) {
    const zd = zs.getRange(2, 1, zl - 1, ZONE_STOCK_HEADERS.length).getValues();
    zd.forEach(r => {
      if (r[ZS_COL_LOCATION - 1] !== location) return;
      const qty = Number(r[ZS_COL_ONHAND - 1]) || Number(r[4]) || 0;
      if (qty <= 0) return;
      usedZoneStock = true;
      const z = String(r[ZS_COL_ZONE - 1] || '');
      addItem(z, r[ZS_COL_BARCODE - 1], r[ZS_COL_PRODUCT - 1], qty, !!datedZone[barcodeKey(r[ZS_COL_BARCODE - 1]) + '|' + z]);
    });
  }
  if (!usedZoneStock) {
    const sheet = getLocationSheetOrNull(location);
    if (sheet && sheet.getLastRow() >= FIRST_DATA_ROW) {
      const lastCol = sheet.getLastColumn();
      const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
      const qi = headers.indexOf(QTY_ON_HAND_HEADER);
      const n = sheet.getLastRow() - FIRST_DATA_ROW + 1;
      const bcs = sheet.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, n, 1).getValues();
      const names = sheet.getRange(FIRST_DATA_ROW, LOCATION_NAME_COL, n, 1).getValues();
      const onh = qi >= 0 ? sheet.getRange(FIRST_DATA_ROW, qi + 1, n, 1).getValues() : null;
      const cnt = sheet.getRange(FIRST_DATA_ROW, LOCATION_COUNTED_COL, n, 1).getValues();
      for (let i = 0; i < n; i++) {
        if (!bcs[i][0]) continue;
        const qty = (onh ? Number(onh[i][0]) || 0 : 0) || Number(cnt[i][0]) || 0;
        if (qty <= 0) continue;
        addItem('(sans zone)', bcs[i][0], names[i][0], qty, !!datedAny[barcodeKey(bcs[i][0])]);
      }
    }
  }
  let total = 0, dated = 0;
  Object.keys(zones).forEach(z => {
    total += zones[z].total; dated += zones[z].dated;
    const lc = lastChecked[z];
    zones[z].daysSince = lc ? Math.round((today - lc) / 86400000) : null;
    zones[z].todo.sort((a, b) => a.name.localeCompare(b.name));
  });
  return jsonResponse({ success: true, location: location, zones: zones, totals: { total: total, dated: dated }, generatedAt: new Date().toISOString() });
}

/**
 * saveExpiryCheck — records what one person found for ONE product in ONE zone:
 * a list of batches (date + qty), or noExpiry:true for products with no date.
 * Replaces that product's previous active batches in that zone.
 *   { key, action, location, zone, barcode, name, by, noExpiry, batches:[{expiry:"DD/MM/YYYY", qty}] }
 */
function handleSaveExpiryCheck(body) {
  const bc = normalizeBarcode(body.barcode || '');
  if (!bc || VALID_LOCATIONS.indexOf(body.location) === -1) {
    return jsonResponse({ success: false, error: "Missing barcode or unknown location" }, 400);
  }
  const batches = Array.isArray(body.batches) ? body.batches : [];
  if (!body.noExpiry && !batches.length) {
    return jsonResponse({ success: false, error: "No batches provided" }, 400);
  }
  for (let i = 0; i < batches.length; i++) {
    if (!parseFrDate(batches[i].expiry)) {
      return jsonResponse({ success: false, error: "Each expiry must be DD/MM/YYYY" }, 400);
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ws = ensureExpiryTab();
    _exEnsureCols(ws);
    const zone = String(body.zone || '');
    const key = barcodeKey(bc);
    const last = ws.getLastRow();
    let replaced = 0;
    if (last >= 2) {
      const data = ws.getRange(2, 1, last - 1, EX_ALL_HEADERS.length).getValues();
      for (let i = 0; i < data.length; i++) {
        const r = data[i];
        if (r[EX_COL_LOCATION - 1] !== body.location) continue;
        if (String(r[EX_COL_ZONE - 1] || '') !== zone) continue;
        if (barcodeKey(r[EX_COL_BARCODE - 1]) !== key) continue;
        if (!_exIsDated(r[EX_COL_STATUS - 1])) continue;
        ws.getRange(i + 2, EX_COL_STATUS).setValue('replaced');
        replaced++;
      }
    }
    const tz = Session.getScriptTimeZone();
    const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    const day = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');
    const name = body.name || bc;
    const rows = body.noExpiry
      ? [[stamp, body.location, zone, bc, name, 0, '', '', 'noexpiry', body.by || '', day]]
      : batches.map(b => [stamp, body.location, zone, bc, name, Number(b.qty) || 0, String(b.expiry), '', 'watch', body.by || '', day]);
    const start = ws.getLastRow() + 1;
    const rng = ws.getRange(start, 1, rows.length, EX_ALL_HEADERS.length);
    rng.setNumberFormat('@');
    rng.setValues(rows);
    return jsonResponse({ success: true, saved: rows.length, replaced: replaced });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════════════════════════════════════
//  EXPIRY v2.1 — actions on a batch + loss log.
//  New columns on Expiry Batches: Action, Action Price, Action Note, Action By, Action At.
//  Actions: 'discount' (batch stays tracked, tagged with the sale price),
//           'sold' (batch resolved), 'discard' (stock reduced, loss logged,
//           batch reduced or resolved), 'transfer' (batch reduced/resolved and a
//           copy created at the destination; the stock itself moves through the
//           normal transfer flow in the Scanner).
//  Losses go to the "Pertes" tab, valued at the selling price. Each discard
//  carries an opId so a retry can never log or deduct twice.
// ══════════════════════════════════════════════════════════════════════════
const EX_COL_ACTION = 12, EX_COL_ACTPRICE = 13, EX_COL_ACTNOTE = 14, EX_COL_ACTBY = 15, EX_COL_ACTAT = 16;
const LOSS_SHEET_NAME = 'Pertes';
const LOSS_HEADERS = ['Timestamp', 'Date', 'Location', 'Zone', 'Barcode', 'Product', 'Qty', 'Unit Price', 'Loss Value', 'Reason', 'Expiry Date', 'By', 'Note', 'OpId'];

function _exEnsureActionCols(ws) {
  if (String(ws.getRange(1, EX_COL_ACTION).getValue() || '') !== 'Action') {
    ws.getRange(1, EX_COL_ACTION, 1, 5).setValues([['Action', 'Action Price', 'Action Note', 'Action By', 'Action At']]);
    ws.getRange(1, EX_COL_ACTION, 1, 5).setFontWeight('bold');
  }
}
function _lossSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(LOSS_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(LOSS_SHEET_NAME);
    ws.appendRow(LOSS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, LOSS_HEADERS.length).setFontWeight('bold');
  }
  return ws;
}

function handleApplyExpiryAction(body) {
  const row = Number(body.row);
  const action = String(body.exAction || '');
  if (!row || ['discount', 'sold', 'discard', 'transfer'].indexOf(action) === -1) {
    return jsonResponse({ success: false, error: "Missing row or unknown action" }, 400);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ws = ensureExpiryTab();
    _exEnsureCols(ws); _exEnsureActionCols(ws);
    if (row < 2 || row > ws.getLastRow()) return jsonResponse({ success: false, error: "Row out of range" });
    const r = ws.getRange(row, 1, 1, 16).getValues()[0];
    if (!_exIsActive(r[EX_COL_STATUS - 1])) return jsonResponse({ success: false, error: "Ce lot est déjà traité" });
    const loc = r[EX_COL_LOCATION - 1], zone = String(r[EX_COL_ZONE - 1] || ''), bc = String(r[EX_COL_BARCODE - 1] || ''),
          name = r[EX_COL_PRODUCT - 1], batchQty = Number(r[EX_COL_QTY - 1]) || 0, expStr = _exStr(r[EX_COL_EXPIRY - 1]);
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const stamp = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');
    const by = body.by || '', note = String(body.note || '');
    const qty = Math.max(0, Math.floor(Number(body.qty) || 0)) || batchQty;
    const setAct = (act, price) => ws.getRange(row, EX_COL_ACTION, 1, 5).setValues([[act, price == null ? '' : price, note, by, stamp]]);

    if (action === 'discount') {
      const price = Math.round(Number(body.price));
      if (!price || price <= 0) return jsonResponse({ success: false, error: "Prix invalide" }, 400);
      setAct('discount', price);
      return jsonResponse({ success: true });
    }
    if (action === 'sold') {
      ws.getRange(row, EX_COL_STATUS).setValue('resolved');
      setAct('sold', null);
      return jsonResponse({ success: true });
    }
    if (action === 'discard') {
      const opId = String(body.opId || '');
      const lw = _lossSheet();
      if (opId) {
        const ll = lw.getLastRow();
        if (ll >= 2) {
          const ids = lw.getRange(2, LOSS_HEADERS.length, ll - 1, 1).getValues();
          for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === opId) return jsonResponse({ success: true, duplicate: true });
        }
      }
      if (batchQty && qty > batchQty) return jsonResponse({ success: false, error: "Quantité supérieure au lot (" + batchQty + ")" }, 400);
      const price = _exPriceMap()[barcodeKey(bc)];
      let stockMsg = 'ok';
      try {
        const out = handleApplyStockMovements({ location: loc, movements: [{ barcode: bc, delta: -qty, type: 'writeoff', note: 'Perte (' + (body.reason || 'expiré') + ') ' + expStr }] });
        const d = JSON.parse(out.getContent());
        if (!d.success) return jsonResponse({ success: false, error: d.error || 'Stock non mis à jour' });
        if (d.notFound && d.notFound.length) stockMsg = 'nf';
      } catch (e) { return jsonResponse({ success: false, error: 'Stock non mis à jour : ' + e }); }
      const lr = lw.getLastRow() + 1;
      lw.getRange(lr, 1, 1, LOSS_HEADERS.length).setValues([[stamp, Utilities.formatDate(now, tz, 'yyyy-MM-dd'), loc, zone, bc, name, qty,
        price == null ? '' : price, price == null ? '' : price * qty, body.reason || 'expiré', expStr, by, note, opId]]);
      const left = batchQty - qty;
      if (left > 0) ws.getRange(row, EX_COL_QTY).setValue(left);
      else { ws.getRange(row, EX_COL_QTY).setValue(0); ws.getRange(row, EX_COL_STATUS).setValue('resolved'); }
      setAct('discard', null);
      return jsonResponse({ success: true, left: Math.max(0, left), stock: stockMsg });
    }
    // transfer
    const dest = String(body.to || '');
    if (VALID_LOCATIONS.indexOf(dest) === -1 || dest === loc) return jsonResponse({ success: false, error: "Destination invalide" }, 400);
    if (batchQty && qty > batchQty) return jsonResponse({ success: false, error: "Quantité supérieure au lot (" + batchQty + ")" }, 400);
    const left = batchQty - qty;
    if (left > 0) ws.getRange(row, EX_COL_QTY).setValue(left);
    else { ws.getRange(row, EX_COL_QTY).setValue(0); ws.getRange(row, EX_COL_STATUS).setValue('resolved'); }
    setAct('transfer', null);
    const day = Utilities.formatDate(now, tz, 'dd/MM/yyyy');
    const nr = ws.getLastRow() + 1;
    const rng = ws.getRange(nr, 1, 1, 11);
    rng.setNumberFormat('@');
    rng.setValues([[stamp, dest, '', bc, name, qty, expStr, '', 'watch', by, day]]);
    return jsonResponse({ success: true, left: Math.max(0, left) });
  } finally { lock.releaseLock(); }
}

/** getLosses — rows of the Pertes tab from the last N days (default 100). */
function handleGetLosses(days) {
  days = Math.max(7, Math.min(400, Number(days) || 100));
  const ws = _lossSheet();
  const last = ws.getLastRow();
  const out = [];
  if (last >= 2) {
    const cutoff = new Date(Date.now() - days * 86400000);
    ws.getRange(2, 1, last - 1, LOSS_HEADERS.length).getValues().forEach(r => {
      const ds = String(r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : r[1]);
      if (new Date(ds) < cutoff) return;
      out.push({ date: ds, location: r[2], zone: r[3], barcode: String(r[4]), name: r[5], qty: Number(r[6]) || 0,
        price: r[7] === '' ? null : Number(r[7]), value: r[8] === '' ? null : Number(r[8]), reason: r[9], expiry: _exStr(r[10]), by: r[11], note: r[12] });
    });
  }
  return jsonResponse({ success: true, losses: out });
}

// ══════════════════════════════════════════════════════════════════════════
//  ZONE STOCK — wishlist item 1: live stock visibility PER ZONE, not just
//  per store. Everything below is additive: nothing here changes how the
//  location-level "Qty Counted" / "Qty On Hand" columns behave, it just
//  layers a finer-grained twin on top, on its own tab.
// ══════════════════════════════════════════════════════════════════════════
//
// New tab "Zone Stock" (auto-created), one row per (Location, Zone, Barcode):
//   A: Location   B: Zone   C: Barcode   D: Product
//   E: Zone Qty Counted   F: Zone Qty On Hand   G: Clearance (TRUE/FALSE)
//
// Same two-phase philosophy as the existing location tabs:
//   - "Zone Qty Counted" is written by applyZoneCounts(), called from
//     handleUpdateCounts() whenever the Counter app's payload includes a
//     `zones` breakdown — i.e. every time counting happens with a zone
//     selected. A zone-count REPLACES what that zone is known to hold:
//     a barcode previously counted into a zone but absent from this
//     session's scan of that same zone gets its Zone Qty Counted reset to
//     0 (not deleted — so it's still visible that it *used* to live there,
//     until the row ages out on its own the next time the zone is fully
//     recounted and stays empty).
//   - "Zone Qty On Hand" only moves in two ways: (a) seedZoneStockFromCounts(),
//     called from handleSeedQtyOnHand() right after "Envoyer au Sheet" +
//     "Sceller stock de départ", which pulls it up to match Zone Qty
//     Counted — exactly like the location-level seed step; (b)
//     applyZoneMovements(), called from handleApplyStockMovements() every
//     time a sale/return/void hits the location level, which nudges it by
//     the same delta, at whichever zone actually owns that barcode.
//
// Dual-zone tie-break (a product in both its normal section AND the
// sale/clearance zone at once — a real but small exception per the
// wishlist): a sale pulls from whichever zone currently has stock; if both
// do, the zone flagged "Clearance" is drawn down first. A restock/return
// goes back to the non-clearance ("home") zone first. No manual zone
// selection is ever required from the cashier — the Scanner app doesn't
// even need to know zones exist; this all happens on the backend.
// ══════════════════════════════════════════
//  SPOT STOCK — shelf-level location of every product (Counter app)
//  One row per (Location, Spot, Barcode). A shelf counted again replaces
//  everything that shelf held before; shelves not in the payload are kept.
// ══════════════════════════════════════════
// ══════════════════════════════════════════
//  COUNTER USERS — personal sign-in codes for the Inventory Counter
//  Only a hash of each code is stored (the app hashes name + code before
//  sending). Roles: "Admin" (everything) and "Contrôle" (counts and checks).
//  The very first user created must be an Admin; after that only an Admin
//  (proving their own code) can add or change users.
// ══════════════════════════════════════════
const COUNTER_USERS_SHEET = "Counter Users";
const COUNTER_USERS_HEADERS = ['Name', 'Role', 'Code Hash', 'Active', 'Updated At', 'Updated By'];

function counterUsersSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(COUNTER_USERS_SHEET);
  if (!ws) {
    ws = ss.insertSheet(COUNTER_USERS_SHEET);
    ws.appendRow(COUNTER_USERS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, COUNTER_USERS_HEADERS.length).setFontWeight('bold');
  }
  return ws;
}
// Training-only staff (role « Formation », added from formation.html) are
// left out unless includeTraining is true, so they can never sign in to
// Comptage / Stock1981 or sign stock changes.
function readCounterUsers_(includeTraining) {
  const ws = counterUsersSheet_();
  const last = ws.getLastRow();
  if (last < 2) return [];
  return ws.getRange(2, 1, last - 1, COUNTER_USERS_HEADERS.length).getValues()
    .filter(r => String(r[0]).trim())
    .map(r => ({ name: String(r[0]).trim(), role: String(r[1] || 'Contrôle'), hash: String(r[2] || ''), active: r[3] !== false && String(r[3]).toUpperCase() !== 'FALSE' }))
    .filter(u => includeTraining || u.role !== 'Formation');
}
function handleSetCounterUser(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const users = readCounterUsers_(true);
    const name = String(body.name || '').trim();
    if (!name) return jsonResponse({ success: false, error: "Nom manquant" }, 400);
    const role = body.role === 'Admin' ? 'Admin' : (body.role === 'Formation' ? 'Formation' : 'Contrôle');
    if (users.length) {
      const admin = users.find(u => u.active && u.role === 'Admin' && u.name.toLowerCase() === String(body.adminName || '').toLowerCase());
      if (!admin || admin.hash !== String(body.adminHash || '')) {
        return jsonResponse({ success: false, error: "Code administrateur incorrect" }, 403);
      }
    } else if (role !== 'Admin') {
      return jsonResponse({ success: false, error: "Le premier utilisateur doit être Admin" }, 400);
    }
    const ws = counterUsersSheet_();
    const idx = users.findIndex(u => u.name.toLowerCase() === name.toLowerCase());
    const existing = idx >= 0 ? users[idx] : null;
    const hash = body.hash ? String(body.hash) : (existing ? existing.hash : '');
    if (!hash) return jsonResponse({ success: false, error: "Code manquant" }, 400);
    const active = body.active === false ? false : true;
    const row = [name, role, hash, active, new Date(), String(body.adminName || name)];
    if (idx >= 0 && role === 'Formation' && existing.role !== 'Formation') {
      return jsonResponse({ success: false, error: "Ce nom est déjà utilisé dans Comptage / Stock1981" }, 400);
    }
    if (idx >= 0) {
      // rows are found by position in the full list (training users included)
      const all = ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues().map(r => String(r[0]).trim().toLowerCase());
      const rowNum = all.indexOf(name.toLowerCase()) + 2;
      ws.getRange(rowNum, 1, 1, row.length).setValues([row]);
    }
    else ws.appendRow(row);
    return jsonResponse({ success: true, users: readCounterUsers_(body.app === 'formation') });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════
//  ZONE PHOTOS — tappable photos in the Counter's shelf picker
//  One row per traced space: where it sits on the zone's photo, as % of
//  the photo (x, y = top-left corner; w, h = size). Photo W/H = the size of
//  the photo when it was traced, so the app can tell when a photo has been
//  replaced and the boxes no longer fit. A zone saved with no boxes at all
//  gets one row with Spot "(aucune)" = "use the drawing for this zone".
//  Only an Admin (proving their own code) can change the traces.
// ══════════════════════════════════════════
const ZONE_PHOTOS_SHEET = "Zone Photos";
const ZONE_PHOTOS_HEADERS = ['Location', 'Zone', 'Spot', 'X %', 'Y %', 'W %', 'H %', 'Photo W', 'Photo H', 'Traced By', 'Traced At'];
const ZONE_PHOTOS_NONE = '(aucune)';

function zonePhotosSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(ZONE_PHOTOS_SHEET);
  if (!ws) {
    ws = ss.insertSheet(ZONE_PHOTOS_SHEET);
    ws.appendRow(ZONE_PHOTOS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, ZONE_PHOTOS_HEADERS.length).setFontWeight('bold');
  }
  return ws;
}
// -> { "Store 2": { "A1": { pw, ph, by, at, spots: { "A1-5": [x, y, w, h], ... } } } }
function readZonePhotoTraces_() {
  const ws = zonePhotosSheet_();
  const last = ws.getLastRow();
  const out = {};
  if (last < 2) return out;
  ws.getRange(2, 1, last - 1, ZONE_PHOTOS_HEADERS.length).getValues().forEach(r => {
    const loc = String(r[0]).trim(), zone = String(r[1]).trim(), spot = String(r[2]).trim();
    if (!loc || !zone) return;
    out[loc] = out[loc] || {};
    const z = out[loc][zone] = out[loc][zone] || { pw: Number(r[7]) || 0, ph: Number(r[8]) || 0, by: String(r[9] || ''), at: r[10] instanceof Date ? r[10].toISOString() : String(r[10] || ''), spots: {} };
    if (spot && spot !== ZONE_PHOTOS_NONE) z.spots[spot] = [Number(r[3]) || 0, Number(r[4]) || 0, Number(r[5]) || 0, Number(r[6]) || 0];
  });
  return out;
}
function counterAdminOk_(body) {
  const users = readCounterUsers_();
  const admin = users.find(u => u.active && u.role === 'Admin' && u.name.toLowerCase() === String(body.adminName || '').toLowerCase());
  return !!(admin && admin.hash === String(body.adminHash || ''));
}
// body: { location, zone, photoW, photoH, spots: { spot: [x,y,w,h] }, reset, adminName, adminHash }
// Replaces every row of that location+zone. reset = remove the zone's rows
// entirely (the app then goes back to its built-in first tracing).
function handleSetZonePhotoTraces(body) {
  if (!counterAdminOk_(body)) return jsonResponse({ success: false, error: "Code administrateur incorrect" }, 403);
  const loc = String(body.location || '').trim(), zone = String(body.zone || '').trim();
  if (!loc || !zone) return jsonResponse({ success: false, error: "Zone manquante" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ws = zonePhotosSheet_();
    const last = ws.getLastRow();
    if (last >= 2) {
      const vals = ws.getRange(2, 1, last - 1, 2).getValues();
      for (let i = vals.length - 1; i >= 0; i--) {
        if (String(vals[i][0]).trim() === loc && String(vals[i][1]).trim() === zone) ws.deleteRow(i + 2);
      }
    }
    if (!body.reset) {
      const pw = Number(body.photoW) || 0, ph = Number(body.photoH) || 0;
      const by = String(body.adminName || ''), at = new Date();
      const clamp = v => Math.round(Math.max(0, Math.min(100, Number(v) || 0)) * 100) / 100;
      const spots = body.spots || {};
      let rows = Object.keys(spots).filter(k => Array.isArray(spots[k]) && spots[k].length === 4)
        .map(k => [loc, zone, k, clamp(spots[k][0]), clamp(spots[k][1]), clamp(spots[k][2]), clamp(spots[k][3]), pw, ph, by, at]);
      if (!rows.length) rows = [[loc, zone, ZONE_PHOTOS_NONE, '', '', '', '', pw, ph, by, at]];
      ws.getRange(ws.getLastRow() + 1, 1, rows.length, ZONE_PHOTOS_HEADERS.length).setValues(rows);
    }
    return jsonResponse({ success: true, traces: readZonePhotoTraces_() });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════
//  MAP LAYOUTS — the store maps drawn in the Counter (🗺 Modifier la carte)
//  Every save is a NEW row (Saved At, Location, Version, Saved By, JSON), so
//  the tab is also the history: to go back, delete the newest rows of that
//  location. The Counter uses the newest row per location; with no row it
//  uses the map built into the app. Only an Admin (own code) can save.
// ══════════════════════════════════════════
const MAP_LAYOUTS_SHEET = 'Map Layouts';
const MAP_LAYOUTS_HEADERS = ['Saved At', 'Location', 'Version', 'Saved By', 'Layout JSON'];
function mapLayoutsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(MAP_LAYOUTS_SHEET);
  if (!ws) {
    ws = ss.insertSheet(MAP_LAYOUTS_SHEET);
    ws.appendRow(MAP_LAYOUTS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, MAP_LAYOUTS_HEADERS.length).setFontWeight('bold');
  }
  return ws;
}
function readMapLayouts_() {
  const ws = mapLayoutsSheet_();
  const last = ws.getLastRow();
  const out = {};
  if (last < 2) return out;
  ws.getRange(2, 1, last - 1, 5).getValues().forEach(r => {
    const loc = String(r[1]).trim();
    if (!loc) return;
    let layout = null;
    try { layout = JSON.parse(String(r[4] || '')); } catch (e) { return; }
    const v = Number(r[2]) || 0;
    if (!out[loc] || v >= out[loc].version) {
      out[loc] = { version: v, by: String(r[3] || ''), at: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''), layout: layout };
    }
  });
  return out;
}
// body: { location, layout: {viewBox, items:[...]}, adminName, adminHash }
function handleSetMapLayout(body) {
  if (!counterAdminOk_(body)) return jsonResponse({ success: false, error: "Code administrateur incorrect" }, 403);
  const loc = String(body.location || '').trim();
  const layout = body.layout;
  if (!loc || !layout || !Array.isArray(layout.items)) return jsonResponse({ success: false, error: "Carte invalide" }, 400);
  const json = JSON.stringify(layout);
  if (json.length > 45000) return jsonResponse({ success: false, error: "Carte trop grande" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const cur = readMapLayouts_()[loc];
    const v = (cur ? cur.version : 0) + 1;
    mapLayoutsSheet_().appendRow([new Date(), loc, v, String(body.adminName || ''), json]);
    return jsonResponse({ success: true, layouts: readMapLayouts_() });
  } finally {
    lock.releaseLock();
  }
}

// ══════════════════════════════════════════
//  CATALOGUE — named product lists for the Scanner's 📚 picker (products
//  without a printed barcode yet, chosen by name/photo). One row per
//  product: List, Order, Barcode, Name, Photo (file name in
//  images/products/ on GitHub). Lists are edited in the Scanner
//  (Admin → Catalogue); saving a list replaces all of its rows.
// ══════════════════════════════════════════
const CATALOGUE_SHEET = 'Catalogue';
const CATALOGUE_HEADERS = ['List', 'Order', 'Barcode', 'Name', 'Photo', 'Updated At'];
function catalogueSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(CATALOGUE_SHEET);
  if (!ws) {
    ws = ss.insertSheet(CATALOGUE_SHEET);
    ws.appendRow(CATALOGUE_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, CATALOGUE_HEADERS.length).setFontWeight('bold');
    ws.getRange('C:C').setNumberFormat('@'); // barcodes stay text (leading zeros)
  }
  return ws;
}
function readCatalogue_() {
  const ws = catalogueSheet_();
  const last = ws.getLastRow();
  const lists = {};
  if (last < 2) return lists;
  ws.getRange(2, 1, last - 1, 5).getValues().forEach(r => {
    const list = String(r[0]).trim();
    if (!list) return;
    (lists[list] = lists[list] || []).push({ order: Number(r[1]) || 0, barcode: String(r[2]).trim(), name: String(r[3]).trim(), photo: String(r[4]).trim() });
  });
  Object.keys(lists).forEach(k => lists[k].sort((a, b) => a.order - b.order));
  return lists;
}
// body: { list, items: [{barcode, name, photo}], remove: true to delete the list }
function handleSetCatalogueList(body) {
  const list = String(body.list || '').trim();
  if (!list) return jsonResponse({ success: false, error: 'Nom de liste manquant' }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ws = catalogueSheet_();
    const last = ws.getLastRow();
    if (last >= 2) {
      const vals = ws.getRange(2, 1, last - 1, 1).getValues();
      for (let i = vals.length - 1; i >= 0; i--) if (String(vals[i][0]).trim() === list) ws.deleteRow(i + 2);
    }
    if (!body.remove) {
      const now = new Date();
      const rows = (Array.isArray(body.items) ? body.items : []).filter(it => it && (it.barcode || it.name))
        .map((it, i) => [list, i + 1, String(it.barcode || ''), String(it.name || ''), String(it.photo || ''), now]);
      if (rows.length) {
        const start = ws.getLastRow() + 1;
        ws.getRange(start, 3, rows.length, 1).setNumberFormat('@');
        ws.getRange(start, 1, rows.length, rows[0].length).setValues(rows);
      }
    }
    return jsonResponse({ success: true, lists: readCatalogue_() });
  } finally {
    lock.releaseLock();
  }
}

const SPOT_STOCK_SHEET_NAME = "Spot Stock";
const SPOT_STOCK_HEADERS = ['Location', 'Zone', 'Spot', 'Barcode', 'Product', 'Qty Counted', 'Counted At', 'Session', 'Counted By'];

function ensureSpotStockTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(SPOT_STOCK_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(SPOT_STOCK_SHEET_NAME);
    ws.appendRow(SPOT_STOCK_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, SPOT_STOCK_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(5, 220);
  } else if (ws.getLastColumn() < SPOT_STOCK_HEADERS.length) {
    ws.getRange(1, 1, 1, SPOT_STOCK_HEADERS.length).setValues([SPOT_STOCK_HEADERS]).setFontWeight('bold');
  }
  return ws;
}

function applySpotCounts(location, spotsPayload, sessionName, countedBy) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return applySpotCountsUnlocked_(location, spotsPayload, sessionName, countedBy);
  } finally {
    lock.releaseLock();
  }
}
// Same as applySpotCounts, for callers that already hold the script lock.
// "Counted At" = when the shelf was actually counted (entry.at, sent by the
// Counter), not when it was sent — the spot-check subtracts sales made AFTER
// that moment, so a count sent hours later must keep its real time.
function applySpotCountsUnlocked_(location, spotsPayload, sessionName, countedBy) {
  {
    const sheet = ensureSpotStockTab();
    const width = SPOT_STOCK_HEADERS.length;
    const lastRow = sheet.getLastRow();
    const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];
    const replaced = {};
    Object.keys(spotsPayload).forEach(sp => { replaced[sp] = true; });
    const kept = existing.filter(r => !(String(r[0]) === String(location) && replaced[String(r[2])]));
    const now = new Date();
    const added = [];
    Object.keys(spotsPayload).forEach(sp => {
      const entry = spotsPayload[sp] || {};
      const at = entry.at ? new Date(entry.at) : null;
      const when = (at && !isNaN(at.getTime()) && at.getTime() <= now.getTime() + 60000) ? at : now;
      (entry.items || []).forEach(it => {
        const bc = normalizeBarcode(it.barcode);
        if (!bc || !(Number(it.count) > 0)) return;
        added.push([location, entry.zone || '', sp, bc, it.name || '', Number(it.count), when, sessionName, entry.by || countedBy || '']);
      });
    });
    const all = kept.concat(added);
    all.sort((a, b) => (String(a[0]) + '|' + String(a[2])).localeCompare(String(b[0]) + '|' + String(b[2]), undefined, { numeric: true }));
    if (lastRow >= 2) sheet.getRange(2, 1, lastRow - 1, width).clearContent();
    if (all.length) {
      sheet.getRange(2, 4, all.length, 1).setNumberFormat('@'); // keep leading zeros on barcodes
      sheet.getRange(2, 1, all.length, width).setValues(all);
    }
    return Object.keys(spotsPayload).length;
  }
}

const ZONE_STOCK_SHEET_NAME = "Zone Stock";
const ZONE_STOCK_HEADERS = ['Location', 'Zone', 'Barcode', 'Product', 'Zone Qty Counted', 'Zone Qty On Hand', 'Clearance'];
const ZS_COL_LOCATION = 1, ZS_COL_ZONE = 2, ZS_COL_BARCODE = 3, ZS_COL_PRODUCT = 4,
      ZS_COL_COUNTED = 5, ZS_COL_ONHAND = 6, ZS_COL_CLEARANCE = 7;

function ensureZoneStockTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(ZONE_STOCK_SHEET_NAME);
  if (!ws) {
    ws = ss.insertSheet(ZONE_STOCK_SHEET_NAME);
    ws.appendRow(ZONE_STOCK_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, ZONE_STOCK_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(4, 200);
  }
  return ws;
}

/**
 * Reads the whole Zone Stock tab once and returns it as an in-memory index,
 * so the functions below don't each re-scan the sheet. Cheap enough for a
 * boutique-scale product list (a few hundred to a couple thousand rows).
 */
function loadZoneStockIndex() {
  const sheet = ensureZoneStockTab();
  const lastRow = sheet.getLastRow();
  const byLocZoneBarcode = {}; // "Location||Zone||Barcode" -> rowNum
  const byLocBarcode = {};     // "Location||Barcode" -> [rowNum, ...]
  if (lastRow < 2) return { sheet, byLocZoneBarcode, byLocBarcode, lastRow };
  const numRows = lastRow - 1;
  const data = sheet.getRange(2, 1, numRows, ZONE_STOCK_HEADERS.length).getValues();
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const loc = row[ZS_COL_LOCATION - 1], zone = row[ZS_COL_ZONE - 1], bc = barcodeKey(row[ZS_COL_BARCODE - 1]);
    if (!loc || !zone || !bc) continue;
    const rowNum = i + 2;
    byLocZoneBarcode[loc + '||' + zone + '||' + bc] = rowNum;
    const lbKey = loc + '||' + bc;
    if (!byLocBarcode[lbKey]) byLocBarcode[lbKey] = [];
    byLocBarcode[lbKey].push(rowNum);
  }
  return { sheet, byLocZoneBarcode, byLocBarcode, lastRow };
}

/**
 * applyZoneCounts — writes this session's per-zone counts into Zone Qty
 * Counted. `zonesPayload` = { zoneCode: [{barcode, count}, ...], ... }.
 * `zoneMeta` = { zoneCode: { clearance: true|false }, ... } (optional —
 * only sent when known; existing Clearance flags are left alone otherwise).
 * Returns the number of zones touched.
 */
function applyZoneCounts(location, zonesPayload, zoneMeta) {
  const idx = loadZoneStockIndex();
  const sheet = idx.sheet;
  const newRows = []; // rows to append: [loc, zone, bc, name, counted, onHand, clearance]
  let touched = 0;

  Object.keys(zonesPayload).forEach(zone => {
    const items = zonesPayload[zone] || [];
    const clearanceFlag = (zoneMeta && zoneMeta[zone] && typeof zoneMeta[zone].clearance === 'boolean')
      ? zoneMeta[zone].clearance : null;
    const seenBarcodes = {};
    touched++;

    items.forEach(entry => {
      const bc = normalizeBarcode(entry.barcode);
      if (!bc) return;
      const bk = barcodeKey(entry.barcode);
      seenBarcodes[bk] = true;
      const key = location + '||' + zone + '||' + bk;
      const rowNum = idx.byLocZoneBarcode[key];
      if (rowNum) {
        sheet.getRange(rowNum, ZS_COL_COUNTED).setValue(entry.count);
        if (entry.name) sheet.getRange(rowNum, ZS_COL_PRODUCT).setValue(entry.name);
        if (clearanceFlag !== null) sheet.getRange(rowNum, ZS_COL_CLEARANCE).setValue(clearanceFlag);
      } else {
        newRows.push([location, zone, bc, entry.name || '', entry.count, 0, clearanceFlag === true]);
      }
    });

    // Anything previously counted into this zone but not scanned this
    // session gets reset to 0 counted — see the two-phase note above.
    const lbPrefix = location + '||';
    Object.keys(idx.byLocZoneBarcode).forEach(key => {
      if (key.indexOf(lbPrefix) !== 0) return;
      const parts = key.split('||'); // [loc, zone, bc]
      if (parts[1] !== zone) return;
      if (seenBarcodes[parts[2]]) return;
      const rowNum = idx.byLocZoneBarcode[key];
      sheet.getRange(rowNum, ZS_COL_COUNTED).setValue(0);
    });
  });

  if (newRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, ZONE_STOCK_HEADERS.length).setValues(newRows);
  }
  return touched;
}

/**
 * seedZoneStockFromCounts — pulls Zone Qty On Hand up to match Zone Qty
 * Counted for every row belonging to `location`. Mirrors
 * handleSeedQtyOnHand's location-level behavior exactly. Returns the
 * number of rows whose On Hand value actually changed.
 */
function seedZoneStockFromCounts(location) {
  const sheet = ensureZoneStockTab();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const numRows = lastRow - 1;
  const data = sheet.getRange(2, 1, numRows, ZONE_STOCK_HEADERS.length).getValues();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const ledgerRows = [];
  let changed = 0;
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    if (row[ZS_COL_LOCATION - 1] !== location) continue;
    const counted = Number(row[ZS_COL_COUNTED - 1]) || 0;
    const onHand = Number(row[ZS_COL_ONHAND - 1]) || 0;
    if (counted === onHand) continue;
    const rowNum = i + 2;
    sheet.getRange(rowNum, ZS_COL_ONHAND).setValue(counted);
    changed++;
    ledgerRows.push([timestamp, location, row[ZS_COL_BARCODE - 1], row[ZS_COL_PRODUCT - 1] || row[ZS_COL_BARCODE - 1],
      'zone-recount-reset', counted - onHand, counted, 'Zone ' + row[ZS_COL_ZONE - 1] + ' — reset from physical count']);
  }
  appendLedgerRows(ledgerRows);
  return changed;
}

/**
 * applyZoneMovements — cascades the same sale/return/void deltas that just
 * hit the location-level Qty On Hand down to whichever zone actually owns
 * each barcode. See the tie-break rule in the header comment above.
 */
function applyZoneMovements(location, movements) {
  const idx = loadZoneStockIndex();
  const sheet = idx.sheet;
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  const ledgerRows = [];

  movements.forEach(m => {
    const bc = normalizeBarcode(m.barcode);
    const delta = Number(m.delta) || 0;
    if (!bc || !delta) return;
    // A movement can name its zone (transfers arrive in « Non placé »): the row
    // is created if that product was never in that zone.
    if (m.zone) {
      const zk = location + '||' + m.zone + '||' + barcodeKey(m.barcode);
      let rn = idx.byLocZoneBarcode[zk], before = 0, prod = String(m.name || '');
      if (rn) {
        const rv = sheet.getRange(rn, 1, 1, ZONE_STOCK_HEADERS.length).getValues()[0];
        before = Number(rv[ZS_COL_ONHAND - 1]) || 0; prod = prod || String(rv[ZS_COL_PRODUCT - 1] || '');
        sheet.getRange(rn, ZS_COL_ONHAND).setValue(before + delta);
      } else {
        rn = sheet.getLastRow() + 1;
        sheet.getRange(rn, ZS_COL_BARCODE).setNumberFormat('@');
        sheet.getRange(rn, 1, 1, ZONE_STOCK_HEADERS.length).setValues([[location, m.zone, bc, prod, 0, delta, false]]);
        idx.byLocZoneBarcode[zk] = rn;
        const lb = location + '||' + barcodeKey(m.barcode);
        (idx.byLocBarcode[lb] = idx.byLocBarcode[lb] || []).push(rn);
      }
      ledgerRows.push([timestamp, location, bc, prod || bc, m.type || 'adjustment', delta, before + delta, 'Zone ' + m.zone + (m.note ? ' — ' + m.note : '')]);
      return;
    }
    const rowNums = idx.byLocBarcode[location + '||' + barcodeKey(m.barcode)];
    if (!rowNums || !rowNums.length) return; // never counted into any zone — nothing to cascade

    const candidates = rowNums.map(rowNum => {
      const rowVals = sheet.getRange(rowNum, 1, 1, ZONE_STOCK_HEADERS.length).getValues()[0];
      return {
        rowNum: rowNum,
        zone: rowVals[ZS_COL_ZONE - 1],
        onHand: Number(rowVals[ZS_COL_ONHAND - 1]) || 0,
        clearance: rowVals[ZS_COL_CLEARANCE - 1] === true,
        product: rowVals[ZS_COL_PRODUCT - 1]
      };
    });

    let chosen;
    if (candidates.length === 1) {
      chosen = candidates[0];
    } else if (delta < 0) {
      // Sale: prefer a clearance zone that still has stock, then any zone
      // with stock, then just the first candidate (records the movement
      // even if it drives that zone's count negative, same as the
      // location-level ledger already tolerates).
      chosen = candidates.find(c => c.clearance && c.onHand > 0)
        || candidates.find(c => c.onHand > 0)
        || candidates[0];
    } else {
      // Return/restock: prefer the non-clearance ("home") zone.
      chosen = candidates.find(c => !c.clearance) || candidates[0];
    }

    const next = chosen.onHand + delta;
    sheet.getRange(chosen.rowNum, ZS_COL_ONHAND).setValue(next);
    ledgerRows.push([timestamp, location, bc, chosen.product || bc, m.type || 'adjustment', delta, next,
      'Zone ' + chosen.zone + (m.note ? ' — ' + m.note : '')]);
  });

  appendLedgerRows(ledgerRows);
}

/**
 * handleGetZoneStock — read-only snapshot for the Counter app's "live
 * view" screen: current Zone Qty On Hand for every zone at a location,
 * independent of any counting session. Called via GET
 * (?action=getZoneStock&location=Store%201).
 */
/**
 * handleGetStockSummary — current "Qty On Hand" per product at each
 * location (Store 1 / Store 2 / Warehouse), keyed by barcode with leading
 * zeros removed so it can be matched with sales data. Read-only.
 */
function handleGetStockSummary() {
  const field = { "Store 1": "s1", "Store 2": "s2", "Warehouse": "wh" };
  const out = {};
  VALID_LOCATIONS.forEach(loc => {
    const sheet = getLocationSheetOrNull(loc);
    if (!sheet) return;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < FIRST_DATA_ROW || lastCol < 1) return;
    const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const qtyIdx = headers.indexOf(QTY_ON_HAND_HEADER);
    if (qtyIdx < 0) return;
    const n = lastRow - FIRST_DATA_ROW + 1;
    const bcs = sheet.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, n, 1).getValues();
    const names = sheet.getRange(FIRST_DATA_ROW, LOCATION_NAME_COL, n, 1).getValues();
    const qtys = sheet.getRange(FIRST_DATA_ROW, qtyIdx + 1, n, 1).getValues();
    for (let i = 0; i < n; i++) {
      if (!bcs[i][0]) continue;
      const k = barcodeKey(bcs[i][0]);
      if (!out[k]) out[k] = { name: String(names[i][0] || ""), s1: null, s2: null, wh: null };
      const q = qtys[i][0];
      out[k][field[loc]] = (q === "" || q === null) ? null : Number(q) || 0;
    }
  });
  return jsonResponse({ success: true, stock: out, generatedAt: new Date().toISOString() });
}

function handleGetZoneStock(location) {
  if (["Store 1", "Store 2", "Warehouse"].indexOf(location) === -1) {
    return jsonResponse({ success: false, error: "Unknown location: " + location }, 400);
  }
  const sheet = ensureZoneStockTab();
  const lastRow = sheet.getLastRow();
  const zones = {}; // zoneCode -> {itemCount, totalQty, items:[{barcode,name,qty}]}
  if (lastRow >= 2) {
    const numRows = lastRow - 1;
    const data = sheet.getRange(2, 1, numRows, ZONE_STOCK_HEADERS.length).getValues();
    data.forEach(row => {
      if (row[ZS_COL_LOCATION - 1] !== location) return;
      const qty = Number(row[ZS_COL_ONHAND - 1]) || 0;
      if (qty <= 0) return; // live view shows what's actually there right now
      const zone = row[ZS_COL_ZONE - 1];
      if (!zones[zone]) zones[zone] = { itemCount: 0, totalQty: 0, items: [] };
      zones[zone].itemCount++;
      zones[zone].totalQty += qty;
      zones[zone].items.push({ barcode: row[ZS_COL_BARCODE - 1], name: row[ZS_COL_PRODUCT - 1], qty: qty });
    });
  }
  return jsonResponse({ success: true, location: location, zones: zones, generatedAt: new Date().toISOString() });
}

// ══════════════════════════════════════════════════════════════════════════
//  CONTRÔLE DU JOUR — anti-theft shelf spot-check (Counter app)
// ──────────────────────────────────────────────────────────────────────────
//  Each day the Counter asks for a plan (getControlPlan): 4 shelves drawn at
//  random, weighted toward expensive zones, shelves that had gaps and shelves
//  not checked for a while. The draw is stored in "Tirages" so it is the same
//  on every phone and can't be re-rolled. Patronne can add her own shelves.
//
//  When a shelf is counted (controlCount), it is compared with what SHOULD be
//  there:   expected = last count of that shelf
//                      − items sold since (Sales_yyyy-MM-dd tabs of the sales
//                        Sheet, exact sale time) + returns
//                      ± transfers / write-offs since (Stock Ledger)
//  A product that also sits on other shelves only gets its share of the sales
//  (marked "≈"). Differences go to the "Écarts" tab, with who was at the
//  register during the period. The new count then becomes the shelf's
//  reference for next time.
//
//  Each gap is then marked: Retrouvé ailleurs / Perte connue / Non expliqué.
//  sendControlReport() emails the day's summary (run installControlReport()
//  once to schedule it every evening at 21:30).
// ══════════════════════════════════════════════════════════════════════════
// Sales Sheet ("Tableau de bords": SalesLog, Sales_yyyy-MM-dd, ShiftReconciliation…).
// If this ID is wrong or empty, the check falls back to the Stock Ledger
// (sales synced every 15 min — slightly less precise, no cashier names).
const SALES_SHEET_ID = 'PUT-THE-SALES-SHEET-ID-HERE';
const CONTROL_TZ = 'Africa/Ouagadougou';
const CONTROL_REPORT_EMAIL = '1981estvideo@gmail.com';
const CONTROL_DRAWS_PER_DAY = 4;
const CONTROL_LOC_NAMES = { 'Store 1': 'Balkuy', 'Store 2': 'Ouaga 2000', 'Warehouse': 'Entrepôt' };

const TIRAGES_SHEET = 'Tirages';
const TIRAGES_HEADERS = ['Date', 'Location', 'Units', 'Created At'];
const CONTROLS_SHEET = 'Contrôles';
const CONTROLS_HEADERS = ['ID', 'Date', 'Time', 'Location', 'Unit', 'Zone', 'Label', 'Type', 'Spots', 'Counted By',
  'Products', 'Missing Units', 'Missing Value', 'Extra Units', 'First Count', 'Baseline JSON', 'Status', 'Source'];
const ECARTS_SHEET = 'Écarts';
const ECARTS_HEADERS = ['ID', 'Control ID', 'Date', 'Time', 'Location', 'Zone', 'Unit', 'Label', 'Barcode', 'Product',
  'Reference Qty', 'Reference At', 'Mouvements', 'Expected', 'Counted', 'Écart', 'Unit Price', 'Value (CFA)', 'Approx',
  'Type', 'Counted By', 'Caissiers (période)', 'Statut', 'Note', 'Resolved By', 'Resolved At'];
const EC = {}; ECARTS_HEADERS.forEach((h, i) => { EC[h] = i; });
const CT = {}; CONTROLS_HEADERS.forEach((h, i) => { CT[h] = i; });

function ctlSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
    ws.appendRow(headers);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return ws;
}
function ctlRows_(ws, width) {
  const last = ws.getLastRow();
  return last >= 2 ? ws.getRange(2, 1, last - 1, width).getValues() : [];
}
function ctlDay_(d) { return Utilities.formatDate(d || new Date(), CONTROL_TZ, 'yyyy-MM-dd'); }
function ctlTime_(d) { return Utilities.formatDate(d || new Date(), CONTROL_TZ, 'HH:mm'); }
function ctlDayStr_(v) { return v instanceof Date ? ctlDay_(v) : String(v || ''); }
function ctlId_(p) { return p + Utilities.formatDate(new Date(), CONTROL_TZ, 'yyMMddHHmmss') + Math.floor(Math.random() * 900 + 100); }
function ctlFmt_(n) { return Math.round(Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' F'; }
// "dd/MM/yyyy" + "HH:mm[:ss]" (text or Date cells) -> Date, or null.
function ctlParseDT_(dv, tv, tz) {
  let ds = dv instanceof Date ? Utilities.formatDate(dv, tz, 'dd/MM/yyyy') : String(dv || '').trim();
  let ts = tv instanceof Date ? Utilities.formatDate(tv, tz, 'HH:mm:ss') : String(tv || '').trim();
  const dm = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(ds);
  if (!dm) return null;
  const tm = /(\d{1,2})[:hH](\d{2})(?::(\d{2}))?/.exec(ts) || [0, '0', '00', '00'];
  const s = ('0' + dm[1]).slice(-2) + '/' + ('0' + dm[2]).slice(-2) + '/' + dm[3] + ' ' +
            ('0' + tm[1]).slice(-2) + ':' + tm[2] + ':' + (tm[3] || '00');
  try { return Utilities.parseDate(s, tz, 'dd/MM/yyyy HH:mm:ss'); } catch (e) { return null; }
}

// ── Sales of one store between two moments, from the sales Sheet ──
// Returns { ok, rows:[{bk, qty, t, emp}], voids:{emp:n} } — qty signed (returns < 0).
let SALES_SS_CACHE_ = null;
function ctlSalesSS_() {
  if (SALES_SS_CACHE_ !== null) return SALES_SS_CACHE_;
  SALES_SS_CACHE_ = false;
  if (SALES_SHEET_ID) {
    try {
      const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
      if (ss.getSheetByName('SalesLog') || ss.getSheets().some(s => /^Sales_\d{4}-\d{2}-\d{2}$/.test(s.getName()))) SALES_SS_CACHE_ = ss;
    } catch (e) { Logger.log('Sales Sheet not reachable: ' + e); }
  }
  return SALES_SS_CACHE_;
}
function ctlSalesBetween_(location, from, to) {
  const ss = ctlSalesSS_();
  const out = { ok: !!ss, rows: [], voids: {} };
  if (!ss) return out;
  const start = new Date(Math.max(from.getTime(), to.getTime() - 120 * 86400000));
  for (let d = new Date(start.getTime() - 86400000); d.getTime() <= to.getTime() + 86400000; d = new Date(d.getTime() + 86400000)) {
    const sh = ss.getSheetByName('Sales_' + ctlDay_(d));
    if (!sh || sh.getLastRow() < 2) continue;
    sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues().forEach(r => {
      if (String(r[12]) !== location) return;
      const t = ctlParseDT_(r[0], r[1], CONTROL_TZ);
      if (!t || t.getTime() <= from.getTime() || t.getTime() > to.getTime()) return;
      const type = String(r[2]), emp = String(r[11] || '').trim();
      if (type === 'Voided') { if (emp) out.voids[emp] = (out.voids[emp] || 0) + 1; return; }
      const qty = Number(r[6]) || 0;
      out.rows.push({ bk: r[3] ? barcodeKey(r[3]) : '', qty: qty, t: t.getTime(), emp: emp });
    });
  }
  return out;
}
// Non-sale stock movements (transfers, write-offs…) from the Stock Ledger.
// When the sales Sheet can't be read, sales/returns/voids come from here too.
function ctlLedgerMoves_(location, from, to, includeSales) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LEDGER_SHEET_NAME);
  const out = [];
  if (!ws || ws.getLastRow() < 2) return out;
  const last = ws.getLastRow(), n = Math.min(last - 1, 8000);
  const tz = Session.getScriptTimeZone();
  const skip = { 'recount-reset': 1, 'zone-recount-reset': 1 };
  const saleTypes = { 'sold': 1, 'return': 1, 'void': 1 };
  ws.getRange(last - n + 1, 1, n, 8).getValues().forEach(r => {
    if (String(r[1]) !== location) return;
    const type = String(r[4] || ''), note = String(r[7] || '');
    if (skip[type] || note.indexOf('Zone ') === 0) return;
    if (saleTypes[type] && !includeSales) return;
    let t = r[0] instanceof Date ? r[0] : null;
    if (!t) { try { t = Utilities.parseDate(String(r[0]), tz, 'dd/MM/yyyy HH:mm'); } catch (e) { t = null; } }
    if (!t || t.getTime() <= from.getTime() || t.getTime() > to.getTime()) return;
    out.push({ bk: barcodeKey(r[2]), delta: Number(r[5]) || 0, t: t.getTime(), type: type });
  });
  return out;
}

// ── Daily plan: 4 random shelves (weighted) + today's results + open gaps ──
function handleGetControlPlan(body) {
  const location = String(body.location || '');
  if (VALID_LOCATIONS.indexOf(location) === -1) return jsonResponse({ success: false, error: 'Emplacement inconnu' }, 400);
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : ctlDay_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const tw = ctlSheet_(TIRAGES_SHEET, TIRAGES_HEADERS);
    const tir = ctlRows_(tw, TIRAGES_HEADERS.length);
    let units = null;
    tir.forEach(r => { if (ctlDayStr_(r[0]) === date && String(r[1]) === location) { try { units = JSON.parse(r[2]); } catch (e) {} } });
    const controls = ctlRows_(ctlSheet_(CONTROLS_SHEET, CONTROLS_HEADERS), CONTROLS_HEADERS.length).filter(r => String(r[CT['Location']]) === location);
    const ecarts = ctlRows_(ctlSheet_(ECARTS_SHEET, ECARTS_HEADERS), ECARTS_HEADERS.length).filter(r => String(r[EC['Location']]) === location);
    const cands = Array.isArray(body.candidates) ? body.candidates.filter(c => c && c.id && c.zone) : [];
    if (!units && cands.length) {
      const now = Date.now();
      const lastCtl = {}, gapUnit = {};
      controls.forEach(r => { const d = new Date(ctlDayStr_(r[CT['Date']])).getTime(); const u = String(r[CT['Unit']]); if (!lastCtl[u] || d > lastCtl[u]) lastCtl[u] = d; });
      ecarts.forEach(r => {
        const st = String(r[EC['Statut']]);
        if ((st === 'Non expliqué' || st === 'À vérifier') && now - new Date(ctlDayStr_(r[EC['Date']])).getTime() < 30 * 86400000) gapUnit[String(r[EC['Unit']])] = true;
      });
      const pool = cands.map(c => {
        let w = 1 + (c.risk ? 3 : 0) + (gapUnit[c.id] ? 3 : 0);
        w += lastCtl[c.id] ? Math.min((now - lastCtl[c.id]) / (7 * 86400000), 2) : 1.5;
        return { c: c, w: w };
      });
      const picked = [], usedZones = {};
      const draw = (allowSameZone) => {
        const avail = pool.filter(p => !p.taken && (allowSameZone || !usedZones[p.c.zone]));
        const total = avail.reduce((a, p) => a + p.w, 0);
        if (!avail.length || total <= 0) return false;
        let x = Math.random() * total;
        for (const p of avail) { x -= p.w; if (x <= 0) { p.taken = true; usedZones[p.c.zone] = true; picked.push(p.c); return true; } }
        const p = avail[avail.length - 1]; p.taken = true; usedZones[p.c.zone] = true; picked.push(p.c); return true;
      };
      while (picked.length < CONTROL_DRAWS_PER_DAY && draw(false)) {}
      while (picked.length < CONTROL_DRAWS_PER_DAY && draw(true)) {}
      units = picked.map(c => ({ id: c.id, zone: c.zone, label: c.label || c.id }));
      tw.appendRow([date, location, JSON.stringify(units), new Date()]);
      tw.getRange(tw.getLastRow(), 1).setNumberFormat('@').setValue(date);
    }
    const latestByUnit = {};
    controls.forEach(r => { if (ctlDayStr_(r[CT['Date']]) === date && String(r[CT['Status']]) !== 'Recompté') latestByUnit[String(r[CT['Unit']])] = r; });
    const todays = Object.keys(latestByUnit).map(k => latestByUnit[k])
      .map(r => ({ id: String(r[CT['ID']]), unit: String(r[CT['Unit']]), zone: String(r[CT['Zone']]), label: String(r[CT['Label']]),
        type: String(r[CT['Type']]), time: String(r[CT['Time']]), by: String(r[CT['Counted By']]), first: r[CT['First Count']] === true || String(r[CT['First Count']]).toUpperCase() === 'TRUE',
        missingUnits: Number(r[CT['Missing Units']]) || 0, missingValue: Number(r[CT['Missing Value']]) || 0,
        lines: ecarts.filter(e => String(e[EC['Control ID']]) === String(r[CT['ID']])).map(ctlEcartObj_) }));
    const cutoff = Date.now() - 14 * 86400000;
    const pending = ecarts.filter(e => String(e[EC['Statut']]) === 'À vérifier' && new Date(ctlDayStr_(e[EC['Date']])).getTime() >= cutoff).map(ctlEcartObj_);
    return jsonResponse({ success: true, date: date, units: units || [], controls: todays, pending: pending });
  } finally {
    lock.releaseLock();
  }
}
function ctlEcartObj_(e) {
  return { id: String(e[EC['ID']]), date: ctlDayStr_(e[EC['Date']]), unit: String(e[EC['Unit']]), label: String(e[EC['Label']]),
    zone: String(e[EC['Zone']]), barcode: String(e[EC['Barcode']]), name: String(e[EC['Product']]),
    ref: Number(e[EC['Reference Qty']]) || 0, moves: String(e[EC['Mouvements']]), expected: Number(e[EC['Expected']]) || 0,
    counted: Number(e[EC['Counted']]) || 0, diff: Number(e[EC['Écart']]) || 0, price: e[EC['Unit Price']] === '' ? null : Number(e[EC['Unit Price']]),
    value: Number(e[EC['Value (CFA)']]) || 0, approx: e[EC['Approx']] === true || String(e[EC['Approx']]).toUpperCase() === 'TRUE',
    status: String(e[EC['Statut']]), cashiers: String(e[EC['Caissiers (période)']]), note: String(e[EC['Note']] || '') };
}

// ── One shelf counted in "Contrôle du jour" ──
// body: { location, date, unit:{id, zone, label, kind:'Tirage'|'Choisi'},
//         spots:{ spotCode:{ zone, at, items:[{barcode,count,name}] } }, by }
// Every spot of the shelf is a key in `spots`, even when counted empty.
function handleControlCount(body) {
  const location = String(body.location || '');
  if (VALID_LOCATIONS.indexOf(location) === -1) return jsonResponse({ success: false, error: 'Emplacement inconnu' }, 400);
  const unit = body.unit || {};
  const spots = body.spots || {};
  const spotList = Object.keys(spots);
  if (!unit.id || !spotList.length) return jsonResponse({ success: false, error: 'Étagère ou comptage manquant' }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const now = new Date();
    let countedAt = 0;
    spotList.forEach(sp => { const t = new Date(spots[sp].at || 0).getTime(); if (t > countedAt) countedAt = t; });
    const countDate = new Date(countedAt && countedAt <= now.getTime() + 60000 ? countedAt : now.getTime());
    const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || '')) ? String(body.date) : ctlDay_(countDate);
    const inUnit = {}; spotList.forEach(sp => { inUnit[sp] = true; });

    // 1. Reference = what the Spot Stock tab says these spots held (or, for a
    //    "Recompter" the same day, the reference that earlier count used).
    const cw = ctlSheet_(CONTROLS_SHEET, CONTROLS_HEADERS);
    const ctlAll = ctlRows_(cw, CONTROLS_HEADERS.length);
    const ew = ctlSheet_(ECARTS_SHEET, ECARTS_HEADERS);
    let baseline = null, replacedId = '';
    for (let i = body.recount ? ctlAll.length - 1 : -1; i >= 0; i--) {
      const r = ctlAll[i];
      if (String(r[CT['Location']]) === location && String(r[CT['Unit']]) === String(unit.id) && ctlDayStr_(r[CT['Date']]) === date && String(r[CT['Status']]) !== 'Recompté') {
        try { baseline = JSON.parse(r[CT['Baseline JSON']] || '{}'); } catch (e) { baseline = {}; }
        replacedId = String(r[CT['ID']]);
        cw.getRange(i + 2, CT['Status'] + 1).setValue('Recompté');
        break;
      }
    }
    const spotWs = ensureSpotStockTab();
    const spotRows = ctlRows_(spotWs, SPOT_STOCK_HEADERS.length).filter(r => String(r[0]) === location);
    if (!baseline) {
      baseline = {};
      spotRows.forEach(r => {
        if (!inUnit[String(r[2])]) return;
        const bk = barcodeKey(r[3]);
        const at = r[6] instanceof Date ? r[6].getTime() : new Date(r[6]).getTime();
        if (!baseline[bk]) baseline[bk] = { barcode: normalizeBarcode(r[3]), name: String(r[4] || ''), qty: 0, at: at || now.getTime() };
        baseline[bk].qty += Number(r[5]) || 0;
        if (at && at < baseline[bk].at) baseline[bk].at = at;
      });
    }
    // A shelf refilled ("Réassort") since its reference can't be compared:
    // today's count simply becomes its new reference.
    let restockReset = false;
    if (!replacedId && Object.keys(baseline).length) {
      let minRef = Infinity;
      Object.keys(baseline).forEach(bk => { if (baseline[bk].at < minRef) minRef = baseline[bk].at; });
      ctlRows_(ctlSheet_(RESTOCK_SHEET, RESTOCK_HEADERS), RESTOCK_HEADERS.length).forEach(r => {
        if (String(r[2]) !== location) return;
        const t = new Date(r[7]).getTime();
        if (!t || t <= minRef || t > countDate.getTime()) return;
        if (String(r[4]).split(' ').some(sp => inUnit[sp])) restockReset = true;
      });
      if (restockReset) baseline = {};
    }
    if (replacedId) {
      const er = ctlRows_(ew, ECARTS_HEADERS.length);
      er.forEach((e, i) => { if (String(e[EC['Control ID']]) === replacedId) ew.getRange(i + 2, EC['Statut'] + 1).setValue('Remplacé (recompté)'); });
    }
    const firstCount = Object.keys(baseline).length === 0;

    // 2. What was counted now.
    const counted = {};
    spotList.forEach(sp => (spots[sp].items || []).forEach(it => {
      const bk = barcodeKey(it.barcode);
      if (!bk || !(Number(it.count) > 0)) return;
      if (!counted[bk]) counted[bk] = { barcode: normalizeBarcode(it.barcode), name: it.name || '', qty: 0 };
      counted[bk].qty += Number(it.count);
    }));

    // 3. Movements since the reference.
    const lines = [];
    let source = 'ventes', cashiers = '';
    if (!firstCount) {
      let minAt = countDate.getTime();
      Object.keys(baseline).forEach(bk => { if (baseline[bk].at < minAt) minAt = baseline[bk].at; });
      const sales = ctlSalesBetween_(location, new Date(minAt), countDate);
      if (!sales.ok) source = 'registre';
      const moves = ctlLedgerMoves_(location, new Date(minAt), countDate, !sales.ok);
      const emps = {};
      sales.rows.forEach(s => { if (s.emp) emps[s.emp] = true; });
      cashiers = Object.keys(emps).join(', ');
      // Share of each product held by this shelf (others = the other shelves' counts).
      const others = {};
      spotRows.forEach(r => { if (inUnit[String(r[2])]) return; const bk = barcodeKey(r[3]); others[bk] = (others[bk] || 0) + (Number(r[5]) || 0); });
      const mm = _masterMap(), sp = _spMap(_spSheet());
      const keys = {}; Object.keys(baseline).forEach(k => { keys[k] = 1; }); Object.keys(counted).forEach(k => { keys[k] = 1; });
      Object.keys(keys).forEach(bk => {
        const b = baseline[bk], c = counted[bk];
        const got = c ? c.qty : 0;
        const name = (c && c.name) || (b && b.name) || (mm[bk] ? mm[bk].name : bk);
        if (!b) { lines.push({ bk: bk, barcode: c.barcode, name: name, ref: 0, expected: null, counted: got, diff: 0, kind: 'Nouveau' }); return; }
        let sold = 0, other = 0;
        sales.rows.forEach(s => { if (s.bk === bk && s.t > b.at) sold += s.qty; });
        moves.forEach(m => {
          if (m.bk !== bk || m.t <= b.at) return;
          if (m.type === 'sold' || m.type === 'return' || m.type === 'void') sold -= m.delta; // fallback: sales from the ledger
          else other += m.delta;
        });
        const share = (others[bk] || 0) > 0 ? b.qty / (b.qty + others[bk]) : 1;
        const net = -sold + other;
        const expected = Math.max(0, Math.round(b.qty + share * net));
        const price = _effectivePrice(location, mm[bk], sp[bk]);
        const diff = got - expected;
        const parts = [];
        if (sold) parts.push((sold > 0 ? 'vendu ' : 'retour ') + Math.abs(sold));
        if (other) parts.push((other > 0 ? '+' : '') + other + ' transfert/perte');
        lines.push({ bk: bk, barcode: b.barcode, name: name, ref: b.qty, refAt: b.at, moves: parts.join(', ') + (share < 1 && net ? ' (part de ce rayon ≈' + Math.round(share * 100) + '%)' : ''),
          expected: expected, counted: got, diff: diff, price: price, value: diff < 0 && price != null ? -diff * price : 0, approx: share < 1 && net !== 0,
          kind: diff < 0 ? 'Manque' : (diff > 0 ? 'En plus' : 'OK') });
      });
    }

    // 4. Write the gaps and the control record.
    const controlId = ctlId_('C');
    const time = ctlTime_(countDate);
    const by = String(body.by || '');
    const newRows = [];
    lines.filter(l => l.kind === 'Manque' || l.kind === 'En plus').forEach(l => {
      newRows.push([ctlId_('E') + newRows.length, controlId, date, time, location, unit.zone || '', unit.id, unit.label || unit.id,
        l.barcode, l.name, l.ref, l.refAt ? Utilities.formatDate(new Date(l.refAt), CONTROL_TZ, 'dd/MM/yyyy HH:mm') : '', l.moves || '',
        l.expected, l.counted, l.diff, l.price == null ? '' : l.price, l.value, l.approx, l.kind, by, cashiers,
        l.kind === 'Manque' ? 'À vérifier' : 'En plus', '', '', '']);
    });
    if (newRows.length) {
      const start = ew.getLastRow() + 1;
      ew.getRange(start, 1, newRows.length, ECARTS_HEADERS.length).setNumberFormat('@');
      ew.getRange(start, 1, newRows.length, ECARTS_HEADERS.length).setValues(newRows.map(r => r.map(v => typeof v === 'number' || typeof v === 'boolean' ? String(v) : v)));
      // keep numbers numeric for sums in the Sheet
      [EC['Reference Qty'], EC['Expected'], EC['Counted'], EC['Écart'], EC['Unit Price'], EC['Value (CFA)']].forEach(ci => {
        const rng = ew.getRange(start, ci + 1, newRows.length, 1);
        rng.setNumberFormat('0');
        rng.setValues(newRows.map(r => [r[ci] === '' ? '' : Number(r[ci])]));
      });
      ew.getRange(start, EC['Approx'] + 1, newRows.length, 1).setNumberFormat('General').setValues(newRows.map(r => [r[EC['Approx']] === true]));
    }
    const missUnits = lines.reduce((a, l) => a + (l.diff < 0 && l.kind === 'Manque' ? -l.diff : 0), 0);
    const missValue = lines.reduce((a, l) => a + (l.kind === 'Manque' ? l.value : 0), 0);
    const extraUnits = lines.reduce((a, l) => a + (l.kind === 'En plus' ? l.diff : 0), 0);
    const crow = [controlId, date, time, location, unit.id, unit.zone || '', unit.label || unit.id, unit.kind === 'Choisi' ? 'Choisi' : 'Tirage',
      spotList.join(' '), by, Object.keys(counted).length, missUnits, missValue, extraUnits, firstCount, JSON.stringify(baseline), 'OK', firstCount ? '' : source];
    const cr = cw.getLastRow() + 1;
    cw.getRange(cr, 2, 1, 2).setNumberFormat('@');
    cw.getRange(cr, 1, 1, crow.length).setValues([crow]);

    // 5. Today's count becomes the shelf's new reference.
    const payload = {};
    spotList.forEach(spc => { payload[spc] = { zone: spots[spc].zone || unit.zone || '', by: by, at: countDate.toISOString(), items: spots[spc].items || [] }; });
    applySpotCountsUnlocked_(location, payload, 'Contrôle ' + date, by);

    return jsonResponse({ success: true, controlId: controlId, firstCount: firstCount, restockReset: restockReset, source: firstCount ? '' : source, cashiers: cashiers,
      missingUnits: missUnits, missingValue: missValue, extraUnits: extraUnits,
      lines: ctlRows_(ew, ECARTS_HEADERS.length).filter(e => String(e[EC['Control ID']]) === controlId).map(ctlEcartObj_) });
  } finally {
    lock.releaseLock();
  }
}

// ── "Réassort": a shelf was refilled from the back / warehouse ──
// body: { location, zone, unit, label, spots:[spot codes], by, note }
const RESTOCK_SHEET = 'Réassorts';
const RESTOCK_HEADERS = ['Date', 'Time', 'Location', 'Zone', 'Spots', 'Label', 'By', 'At', 'Note'];
function handleLogRestock(body) {
  const location = String(body.location || '');
  const spots = Array.isArray(body.spots) ? body.spots.map(String).filter(Boolean) : [];
  if (VALID_LOCATIONS.indexOf(location) === -1 || !spots.length) return jsonResponse({ success: false, error: 'Emplacement ou étagère manquant' }, 400);
  const now = new Date();
  const ws = ctlSheet_(RESTOCK_SHEET, RESTOCK_HEADERS);
  const r = ws.getLastRow() + 1;
  ws.getRange(r, 1, 1, RESTOCK_HEADERS.length).setNumberFormat('@')
    .setValues([[ctlDay_(now), ctlTime_(now), location, String(body.zone || ''), spots.join(' '), String(body.label || ''), String(body.by || ''), now.toISOString(), String(body.note || '')]]);
  return jsonResponse({ success: true });
}

// ── Mark a gap: retrouve | perte | non_explique ──
function handleResolveEcart(body) {
  const map = { retrouve: 'Retrouvé ailleurs', perte: 'Perte connue', non_explique: 'Non expliqué', a_verifier: 'À vérifier' };
  const status = map[String(body.status || '')];
  if (!status || !body.id) return jsonResponse({ success: false, error: 'Statut ou écart manquant' }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ew = ctlSheet_(ECARTS_SHEET, ECARTS_HEADERS);
    const rows = ctlRows_(ew, ECARTS_HEADERS.length);
    const i = rows.findIndex(r => String(r[EC['ID']]) === String(body.id));
    if (i < 0) return jsonResponse({ success: false, error: 'Écart introuvable' });
    const r = rows[i];
    const prev = String(r[EC['Statut']]);
    const by = String(body.by || ''), note = String(body.note || '');
    ew.getRange(i + 2, EC['Statut'] + 1, 1, 4).setValues([[status, note, by, Utilities.formatDate(new Date(), CONTROL_TZ, 'dd/MM/yyyy HH:mm')]]);
    if (status === 'Perte connue' && prev !== 'Perte connue') {
      const qty = -(Number(r[EC['Écart']]) || 0);
      const price = r[EC['Unit Price']] === '' ? '' : Number(r[EC['Unit Price']]);
      const now = new Date();
      _lossSheet().appendRow([Utilities.formatDate(now, CONTROL_TZ, 'dd/MM/yyyy HH:mm'), ctlDay_(now), r[EC['Location']], r[EC['Zone']],
        String(r[EC['Barcode']]), r[EC['Product']], qty, price, price === '' ? '' : price * qty, 'Contrôle étagère — perte connue', '', by,
        (note ? note + ' · ' : '') + 'Écart ' + r[EC['ID']] + ' (stock non modifié)', 'ECART-' + r[EC['ID']]]);
    }
    return jsonResponse({ success: true, id: String(body.id), status: status });
  } finally {
    lock.releaseLock();
  }
}

// ── Evening summary (email + WhatsApp text in the Counter) ──
function buildControlReport_(dateStr) {
  const date = /^\d{4}-\d{2}-\d{2}$/.test(String(dateStr || '')) ? String(dateStr) : ctlDay_();
  const dayStart = Utilities.parseDate(date + ' 00:00:00', CONTROL_TZ, 'yyyy-MM-dd HH:mm:ss');
  const dayEnd = new Date(dayStart.getTime() + 86400000 - 1);
  const frDate = Utilities.formatDate(dayStart, CONTROL_TZ, 'dd/MM/yyyy');
  const tir = ctlRows_(ctlSheet_(TIRAGES_SHEET, TIRAGES_HEADERS), TIRAGES_HEADERS.length);
  const ctl = ctlRows_(ctlSheet_(CONTROLS_SHEET, CONTROLS_HEADERS), CONTROLS_HEADERS.length);
  const ec = ctlRows_(ctlSheet_(ECARTS_SHEET, ECARTS_HEADERS), ECARTS_HEADERS.length);
  const sss = ctlSalesSS_();
  const L = [];
  const counts = { missing: 0, value: 0 };
  L.push('🛡 Contrôle du jour — ' + Utilities.formatDate(dayStart, CONTROL_TZ, 'dd/MM/yyyy'));
  ['Store 2', 'Store 1'].forEach(loc => {
    const name = CONTROL_LOC_NAMES[loc] || loc;
    let units = [];
    tir.forEach(r => { if (ctlDayStr_(r[0]) === date && String(r[1]) === loc) { try { units = JSON.parse(r[2]); } catch (e) {} } });
    // Latest count of each shelf today (an earlier first count + a later check = one shelf).
    const latest = {};
    ctl.forEach(r => { if (String(r[CT['Location']]) === loc && ctlDayStr_(r[CT['Date']]) === date && String(r[CT['Status']]) !== 'Recompté') latest[String(r[CT['Unit']])] = r; });
    const done = Object.keys(latest).map(k => latest[k]);
    const doneUnits = {}; done.forEach(r => { doneUnits[String(r[CT['Unit']])] = true; });
    const gaps = ec.filter(e => String(e[EC['Location']]) === loc && ctlDayStr_(e[EC['Date']]) === date && String(e[EC['Type']]) === 'Manque' && String(e[EC['Statut']]) !== 'Remplacé (recompté)');
    L.push('');
    L.push('━━ ' + name + ' ━━');
    if (!units.length && !done.length) {
      L.push('Aucun contrôle aujourd\'hui.');
    } else {
      const tDone = units.filter(u => doneUnits[u.id]).length;
      if (units.length) L.push('Tirage : ' + tDone + '/' + units.length + ' fait(s)' + (tDone < units.length ? ' ⚠️ pas faits : ' + units.filter(u => !doneUnits[u.id]).map(u => u.label).join(', ') : ''));
      const chosen = done.filter(r => String(r[CT['Type']]) === 'Choisi').length;
      if (chosen) L.push('Choisis par la contrôleuse : ' + chosen);
      const firsts = done.filter(r => r[CT['First Count']] === true || String(r[CT['First Count']]).toUpperCase() === 'TRUE').length;
      if (firsts) L.push('Premier comptage (référence posée) : ' + firsts);
      const by = {}; done.forEach(r => { const b = String(r[CT['Counted By']]); if (b) by[b] = 1; });
      if (Object.keys(by).length) L.push('Compté par : ' + Object.keys(by).join(', '));
      const group = (st) => gaps.filter(e => String(e[EC['Statut']]) === st);
      const sum = (arr) => arr.reduce((a, e) => ({ u: a.u - (Number(e[EC['Écart']]) || 0), v: a.v + (Number(e[EC['Value (CFA)']]) || 0) }), { u: 0, v: 0 });
      const line = (e) => '   • ' + e[EC['Label']] + ' — ' + e[EC['Product']] + ' : ' + e[EC['Écart']] + (e[EC['Approx']] === true || String(e[EC['Approx']]).toUpperCase() === 'TRUE' ? ' ≈' : '') +
        ' (' + ctlFmt_(e[EC['Value (CFA)']]) + ')' + (e[EC['Caissiers (période)']] ? ' — caisse : ' + e[EC['Caissiers (période)']] : '');
      const nonExp = group('Non expliqué'), toCheck = group('À vérifier');
      const ok = gaps.filter(e => ['Retrouvé ailleurs', 'Perte connue'].indexOf(String(e[EC['Statut']])) >= 0);
      if (!gaps.length) L.push('✅ Aucun article manquant.');
      if (nonExp.length) { const s = sum(nonExp); counts.missing += s.u; counts.value += s.v; L.push('❌ Non expliqués : ' + s.u + ' article(s) · ' + ctlFmt_(s.v)); nonExp.forEach(e => L.push(line(e))); }
      if (toCheck.length) { const s = sum(toCheck); counts.missing += s.u; counts.value += s.v; L.push('⏳ À vérifier : ' + s.u + ' article(s) · ' + ctlFmt_(s.v)); toCheck.forEach(e => L.push(line(e))); }
      if (ok.length) { const s = sum(ok); L.push('↩️ Retrouvés / pertes connues : ' + s.u + ' article(s)'); }
      const extra = ec.filter(e => String(e[EC['Location']]) === loc && ctlDayStr_(e[EC['Date']]) === date && String(e[EC['Statut']]) === 'En plus');
      if (extra.length) L.push('➕ En plus que prévu sur ' + extra.length + ' produit(s) (réassort non noté ?)');
    }
    // Register side: cash / Orange Money differences, cash checks, cancelled lines.
    if (sss) {
      try {
        const rec = sss.getSheetByName('ShiftReconciliation');
        if (rec && rec.getLastRow() >= 2) {
          rec.getRange(2, 1, rec.getLastRow() - 1, 14).getValues().forEach(r => {
            const d = r[0] instanceof Date ? Utilities.formatDate(r[0], CONTROL_TZ, 'dd/MM/yyyy') : String(r[0]);
            if (d !== frDate || String(r[3]) !== loc) return;
            const cash = Number(r[9]) || 0, om = Number(r[13]) || 0;
            L.push('💵 Fin de service ' + r[2] + ' : espèces ' + (cash ? (cash > 0 ? '+' : '') + ctlFmt_(cash) : 'OK') + ' · Orange Money ' + (om ? (om > 0 ? '+' : '') + ctlFmt_(om) : 'OK'));
          });
        }
        const cc = sss.getSheetByName('Cash Checks');
        if (cc && cc.getLastRow() >= 2) {
          cc.getRange(2, 1, cc.getLastRow() - 1, 14).getValues().forEach(r => {
            const d = r[1] instanceof Date ? Utilities.formatDate(r[1], CONTROL_TZ, 'dd/MM/yyyy') : String(r[1]);
            if (d !== frDate || String(r[2]) !== loc) return;
            const diff = Number(r[8]) || 0, miss = Number(r[11]) || 0;
            L.push('🧾 Vérif caisse (' + (r[3] || '?') + ') : espèces ' + (diff ? (diff > 0 ? '+' : '') + ctlFmt_(diff) : 'OK') + (miss ? ' · paiements non confirmés ' + ctlFmt_(miss) : ''));
          });
        }
        const v = ctlSalesBetween_(loc, dayStart, dayEnd).voids;
        const vk = Object.keys(v);
        if (vk.length) L.push('✂️ Lignes annulées : ' + vk.map(k => k + ' ' + v[k]).join(', '));
      } catch (e) { L.push('(données de caisse illisibles : ' + e + ')'); }
    }
  });
  // 7-day trends (unexplained + still to check).
  const since = dayStart.getTime() - 6 * 86400000;
  const recent = ec.filter(e => String(e[EC['Type']]) === 'Manque' && ['Non expliqué', 'À vérifier'].indexOf(String(e[EC['Statut']])) >= 0 &&
    new Date(ctlDayStr_(e[EC['Date']]) + 'T12:00:00').getTime() >= since && new Date(ctlDayStr_(e[EC['Date']]) + 'T12:00:00').getTime() <= dayEnd.getTime());
  if (recent.length) {
    const top = (keyFn) => {
      const m = {};
      recent.forEach(e => keyFn(e).forEach(k => { if (!k) return; if (!m[k]) m[k] = { n: 0, v: 0 }; m[k].n += 1; m[k].v += Number(e[EC['Value (CFA)']]) || 0; }));
      return Object.keys(m).sort((a, b) => m[b].v - m[a].v).slice(0, 5).map(k => k + ' (' + m[k].n + '×, ' + ctlFmt_(m[k].v) + ')');
    };
    L.push('');
    L.push('📈 7 derniers jours (non expliqués + à vérifier)');
    L.push('Rayons : ' + top(e => [(CONTROL_LOC_NAMES[e[EC['Location']]] || e[EC['Location']]) + ' ' + e[EC['Label']]]).join(' · '));
    L.push('Produits : ' + top(e => [String(e[EC['Product']])]).join(' · '));
    const cs = top(e => String(e[EC['Caissiers (période)']] || '').split(',').map(s => s.trim()));
    if (cs.length) L.push('En caisse pendant ces écarts : ' + cs.join(' · '));
  }
  return { text: L.join('\n'), missing: counts.missing, value: counts.value };
}

function sendControlReport() {
  const r = buildControlReport_(ctlDay_());
  const subject = '🛡 Contrôle ' + Utilities.formatDate(new Date(), CONTROL_TZ, 'dd/MM') +
    (r.missing ? ' — ' + r.missing + ' article(s) manquant(s), ' + ctlFmt_(r.value) : ' — RAS');
  MailApp.sendEmail(CONTROL_REPORT_EMAIL, subject, r.text);
}
// Run ONCE from the editor: schedules sendControlReport every evening at 21:30.
function installControlReport() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'sendControlReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendControlReport').timeBased().everyDays(1).atHour(21).nearMinute(30).inTimezone(CONTROL_TZ).create();
}
// Run from the editor to see today's report in the log without emailing.
function testControlReport() { Logger.log(buildControlReport_(ctlDay_()).text); }

// ══════════════════════════════════════════
//  WEEKLY BACKUP — every Sunday night (installWeeklyJobs)
//  Into Drive › "Boutique 1981 — Sauvegardes" › "<date> (semaine NN)":
//   1. A full copy of BOTH Sheets (this inventory Sheet + the sales Sheet):
//      the real safety net, restorable in one click.
//   2. One CSV per tab, ready for analysis:
//      - every Sales_yyyy-MM-dd day tab merged into ventes_lignes.csv
//        (one row per sale line, + "Jour" column),
//      - every other tab as its own file (sign-in codes, PINs and settings
//        tabs are left out on purpose), plus lisez-moi.txt describing them.
//      Conventions: comma-separated UTF-8, dates 2026-10-17 (and 09:30 /
//      2026-10-17 09:30), plain numbers, barcodes as text plus a
//      "BarcodeKey" column without leading zeros (same product = same key
//      everywhere).
//   3. A short email: OK, or what failed.
//  Full Sheet copies are kept 12 weeks (older ones go to the Drive bin);
//  CSV files are kept for good.
// ══════════════════════════════════════════
const BACKUP_FOLDER_NAME = 'Boutique 1981 — Sauvegardes';
const BACKUP_KEEP_COPIES_WEEKS = 12;
// Tabs never exported: sign-in code hashes, PINs, settings, the live dashboard.
const BACKUP_SKIP_TABS = ['Counter Users', 'Employees', 'GlobalSettings', 'StoreSettings', 'Dashboard'];

function weeklyBackup() {
  const started = Date.now();
  const tz = CONTROL_TZ;
  const now = new Date();
  const day = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const week = Utilities.formatDate(now, tz, 'w');
  const errors = [], made = [];
  let folder = null;
  try {
    const root = backupRootFolder_();
    folder = root.createFolder(day + ' (semaine ' + week + ')');
  } catch (e) {
    MailApp.sendEmail(CONTROL_REPORT_EMAIL, '⚠ Sauvegarde ' + day + ' — ÉCHEC', 'Impossible de créer le dossier de sauvegarde dans Drive : ' + e);
    return;
  }
  const inv = SpreadsheetApp.getActiveSpreadsheet();
  let sales = null;
  try { sales = SpreadsheetApp.openById(SALES_SHEET_ID); } catch (e) { errors.push('Sheet des ventes inaccessible : ' + e); }

  // 1. full copies
  [[inv, 'Inventaire'], [sales, 'Ventes']].forEach(([ss, label]) => {
    if (!ss) return;
    try { DriveApp.getFileById(ss.getId()).makeCopy('Copie ' + label + ' — ' + day, folder); made.push('Copie complète : ' + label); }
    catch (e) { errors.push('Copie ' + label + ' : ' + e); }
  });

  // 2. CSVs
  const readme = [];
  const add = (name, rows, from) => {
    try {
      folder.createFile(name, backupCsv_(rows), MimeType.CSV);
      readme.push(name + ' — ' + Math.max(rows.length - 1, 0) + ' ligne(s), depuis « ' + from + ' »\n    colonnes : ' + (rows[0] || []).join(', '));
    } catch (e) { errors.push(name + ' : ' + e); }
  };
  const dayRe = /^Sales_(\d{4}-\d{2}-\d{2})$/;
  [[inv, 'inventaire'], [sales, 'ventes']].forEach(([ss, prefix]) => {
    if (!ss) return;
    ss.getSheets().forEach(sh => {
      const nm = sh.getName();
      if (dayRe.test(nm) || BACKUP_SKIP_TABS.indexOf(nm) >= 0) return;
      if (Date.now() - started > 300000) { errors.push('Temps dépassé avant « ' + nm + ' »'); return; }
      try {
        const rows = backupTidy_(sh.getDataRange().getValues());
        if (rows.length) add(prefix + '_' + backupSlug_(nm) + '.csv', rows, nm);
      } catch (e) { errors.push(nm + ' : ' + e); }
    });
  });
  // every day tab merged into one file
  if (sales) {
    try {
      let header = null; const out = [];
      sales.getSheets().filter(sh => dayRe.test(sh.getName())).sort((a, b) => a.getName().localeCompare(b.getName())).forEach(sh => {
        const jour = sh.getName().match(dayRe)[1];
        const vals = sh.getDataRange().getValues();
        if (vals.length < 2) return;
        const h = vals[0].slice(0, 18).map(String);
        if (!header || h.length > header.length) header = h;
        vals.slice(1).forEach(r => { if (String(r[0]).trim() || String(r[3]).trim()) out.push([jour].concat(r.slice(0, 18))); });
      });
      if (header) {
        const rows = backupTidy_([['Jour'].concat(header)].concat(out.map(r => { while (r.length < header.length + 1) r.push(''); return r; })));
        add('ventes_lignes.csv', rows, 'tous les onglets Sales_aaaa-mm-jj');
      }
    } catch (e) { errors.push('ventes_lignes.csv : ' + e); }
  }
  try {
    folder.createFile('lisez-moi.txt',
      'Sauvegarde Boutique 1981 du ' + day + '\n\n' +
      'Conventions : CSV séparé par des virgules, UTF-8 ; dates au format 2026-10-17 ; heures 09:30 ;\n' +
      'nombres sans espaces ni « CFA » ; BarcodeKey = code-barres sans les zéros du début\n' +
      '(même produit = même BarcodeKey dans tous les fichiers).\n' +
      'ventes_lignes.csv = une ligne par article vendu/retourné (Type = Sale / Return / Remise / Voided ;\n' +
      'Voided = ligne annulée, quantité et montant 0). UnitPrice = prix normal, LineTotal = montant payé.\n\n' +
      readme.join('\n'), MimeType.PLAIN_TEXT);
  } catch (e) { errors.push('lisez-moi.txt : ' + e); }

  // 3. keep full copies 12 weeks
  try { backupPruneCopies_(); } catch (e) { errors.push('Nettoyage des anciennes copies : ' + e); }

  const subject = (errors.length ? '⚠ Sauvegarde ' + day + ' — incomplète' : '💾 Sauvegarde ' + day + ' — OK') +
    ' (' + made.length + ' copie(s), ' + readme.length + ' fichier(s) CSV)';
  const body = 'Dossier : ' + folder.getUrl() + '\n\n' + made.join('\n') + '\n' + readme.length + ' fichier(s) CSV.\n' +
    (errors.length ? '\nProblèmes :\n- ' + errors.join('\n- ') + '\n' : '\nTout est sauvegardé.\n');
  MailApp.sendEmail(CONTROL_REPORT_EMAIL, subject, body);
}
function backupRootFolder_() {
  const it = DriveApp.getFoldersByName(BACKUP_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(BACKUP_FOLDER_NAME);
}
function backupSlug_(s) {
  return String(s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '') || 'onglet';
}
// Dates → 2026-10-17 (+ time), French text dates → same, trailing blank
// rows dropped, and a BarcodeKey column after every barcode column.
function backupTidy_(vals) {
  if (!vals.length) return [];
  const tz = CONTROL_TZ;
  const fmtCell = v => {
    if (Object.prototype.toString.call(v) === '[object Date]') {
      if (v.getFullYear() <= 1900) return Utilities.formatDate(v, tz, 'HH:mm');
      const hm = Utilities.formatDate(v, tz, 'HH:mm');
      return Utilities.formatDate(v, tz, 'yyyy-MM-dd') + (hm !== '00:00' ? ' ' + hm : '');
    }
    if (typeof v === 'string') {
      const m = v.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:[ T](\d{1,2}:\d{2}(?::\d{2})?))?$/);
      if (m) return m[3] + '-' + m[2] + '-' + m[1] + (m[4] ? ' ' + m[4] : '');
    }
    return v;
  };
  let rows = vals.map(r => r.map(fmtCell));
  while (rows.length > 1 && rows[rows.length - 1].every(c => String(c) === '')) rows.pop();
  const head = rows[0].map(String);
  const bcCols = [];
  head.forEach((h, i) => { if (/^(barcode|code[- ]?barres?|old barcode|new barcode)$/i.test(h.trim())) bcCols.push(i); });
  if (!bcCols.length) return rows;
  return rows.map((r, ri) => {
    const out = [];
    r.forEach((c, i) => {
      out.push(c);
      if (bcCols.indexOf(i) >= 0) out.push(ri === 0 ? (bcCols.length > 1 ? head[i] + ' Key' : 'BarcodeKey') : (String(c).trim().replace(/^0+/, '') || String(c).trim()));
    });
    return out;
  });
}
function backupCsv_(rows) {
  const cell = v => {
    const s = v == null ? '' : String(v);
    return /[",\n\r]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  };
  return rows.map(r => r.map(cell).join(',')).join('\r\n');
}
function backupPruneCopies_() {
  const root = backupRootFolder_();
  const folders = [];
  const it = root.getFolders();
  while (it.hasNext()) { const f = it.next(); if (/^\d{4}-\d{2}-\d{2}/.test(f.getName())) folders.push(f); }
  folders.sort((a, b) => b.getName().localeCompare(a.getName()));
  folders.slice(BACKUP_KEEP_COPIES_WEEKS).forEach(f => {
    const files = f.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) files.next().setTrashed(true);
  });
}

// ══════════════════════════════════════════
//  WEEKLY BEST SELLERS — emailed every Monday ~7:00 (installWeeklyJobs)
//  Last week, Monday → Sunday, from the sales Sheet's Sales_yyyy-MM-dd day
//  tabs: Sale and Return lines (Voided = cancelled, left out; whole-sale
//  'Remise' rows have no product, left out). Same product = same barcode
//  without leading zeros; no-barcode items (00000) are grouped by name.
//  Contents: week totals per store vs the week before · top 20 by units ·
//  top 20 by revenue · top 10 per store · biggest rises and falls · last
//  week's best sellers that sold nothing · days with no sale recorded.
// ══════════════════════════════════════════
const BESTSELLER_TOP = 20;
const BESTSELLER_STORE_TOP = 10;
// Shown under the title while a store's data is known to be incomplete.
const BESTSELLER_NOTE = 'Ouaga 2000 : chiffres peut-être incomplets (synchro en cours de vérification).';

function bsWeekDays_(endSunday) {
  const out = [];
  for (let i = 6; i >= 0; i--) out.push(Utilities.formatDate(new Date(endSunday.getTime() - i * 86400000), CONTROL_TZ, 'yyyy-MM-dd'));
  return out;
}
function bsLastSunday_(now) {
  // the Sunday that ends the last complete week (Monday → Sunday)
  const dow = Number(Utilities.formatDate(now, CONTROL_TZ, 'u')); // 1 = Monday … 7 = Sunday
  return new Date(now.getTime() - (dow % 7 === 0 ? 7 : dow) * 86400000);
}
// Reads Sale/Return lines of the given days → { products: key → {...}, stores: name → {...}, days: day → {store → lines} }
function bsCollect_(ss, days) {
  const products = {}, stores = {}, daySeen = {};
  days.forEach(d => { daySeen[d] = {}; });
  days.forEach(d => {
    const sh = ss.getSheetByName('Sales_' + d);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, 13).getValues();
    vals.forEach(r => {
      const type = String(r[2]).trim();
      if (type !== 'Sale' && type !== 'Return') return;
      const bc = String(r[3]).trim(), name = String(r[4]).trim() || bc || '(sans nom)';
      const qty = Number(r[6]) || 0, rev = Number(r[8]) || 0, store = String(r[12]).trim() || '(magasin ?)';
      const bk = bc.replace(/^0+/, '');
      const key = (!bk || /^0*$/.test(bc)) ? 'nom:' + name.toLowerCase() : bk;
      const p = products[key] = products[key] || { key, name, barcode: bc, units: 0, rev: 0, byStore: {} };
      p.units += qty; p.rev += rev;
      const ps = p.byStore[store] = p.byStore[store] || { units: 0, rev: 0 };
      ps.units += qty; ps.rev += rev;
      const st = stores[store] = stores[store] || { units: 0, rev: 0, lines: 0 };
      st.units += qty; st.rev += rev; st.lines++;
      daySeen[d][store] = (daySeen[d][store] || 0) + 1;
    });
  });
  return { products, stores, days: daySeen };
}
function bsMoney_(n) { return Math.round(Number(n) || 0).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' F'; }
function bsPct_(a, b) { if (!b) return a ? 'nouveau' : '—'; const p = Math.round((a - b) / Math.abs(b) * 100); return (p > 0 ? '+' : '') + p + ' %'; }
function bsEsc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function buildBestSellers_(now) {
  const end = bsLastSunday_(now || new Date());
  const days = bsWeekDays_(end), prevDays = bsWeekDays_(new Date(end.getTime() - 7 * 86400000));
  const ss = SpreadsheetApp.openById(SALES_SHEET_ID);
  const cur = bsCollect_(ss, days), prev = bsCollect_(ss, prevDays);
  const fr = d => d.slice(8, 10) + '/' + d.slice(5, 7);
  const title = 'Meilleures ventes — semaine du ' + fr(days[0]) + ' au ' + fr(days[6]);
  const list = Object.values(cur.products);
  const prevUnits = k => (prev.products[k] || {}).units || 0;
  const storeNames = Object.keys(cur.stores).concat(Object.keys(prev.stores)).filter((v, i, a) => a.indexOf(v) === i).sort();
  const T = [], H = [];  // plain text + html
  const h2 = t => { T.push('', '== ' + t + ' =='); H.push('<h3 style="margin:18px 0 6px;font:600 15px sans-serif">' + bsEsc_(t) + '</h3>'); };
  const table = (head, rows) => {
    rows.forEach(r => T.push(r.join('  |  ')));
    H.push('<table style="border-collapse:collapse;font:13px sans-serif">' +
      '<tr>' + head.map(x => '<th style="text-align:left;padding:3px 8px;border-bottom:1px solid #ccc">' + bsEsc_(x) + '</th>').join('') + '</tr>' +
      rows.map((r, i) => '<tr style="background:' + (i % 2 ? '#fff' : '#f6f3ee') + '">' + r.map((x, j) => '<td style="padding:3px 8px;' + (j && /^[-+\d]/.test(String(x)) ? 'text-align:right;' : '') + '">' + bsEsc_(x) + '</td>').join('') + '</tr>').join('') +
      '</table>');
  };
  T.push(title); H.push('<h2 style="font:700 18px sans-serif;margin:0 0 4px">' + bsEsc_(title) + '</h2>');
  if (BESTSELLER_NOTE) { T.push(BESTSELLER_NOTE); H.push('<div style="font:12px sans-serif;color:#9a5b00">' + bsEsc_(BESTSELLER_NOTE) + '</div>'); }

  h2('Totaux de la semaine');
  const totRows = storeNames.map(s => { const c = cur.stores[s] || {}, p = prev.stores[s] || {};
    return [s, bsMoney_(c.rev), String(c.units || 0) + ' art.', bsPct_(c.rev || 0, p.rev || 0) + ' vs sem. préc.']; });
  const cr = Object.values(cur.stores).reduce((a, s) => a + s.rev, 0), pr = Object.values(prev.stores).reduce((a, s) => a + s.rev, 0);
  const cu = Object.values(cur.stores).reduce((a, s) => a + s.units, 0);
  totRows.push(['Les deux magasins', bsMoney_(cr), cu + ' art.', bsPct_(cr, pr) + ' vs sem. préc.']);
  table(['Magasin', 'Chiffre', 'Articles', 'Évolution'], totRows);

  const split = p => storeNames.filter(s => p.byStore[s]).map(s => s.split(' ').slice(-2).join(' ') + ' ' + p.byStore[s].units).join(' · ');
  h2('Top ' + BESTSELLER_TOP + ' en quantité (les deux magasins)');
  table(['#', 'Produit', 'Qté', 'Chiffre', 'vs sem. préc.', 'Par magasin'],
    list.slice().sort((a, b) => b.units - a.units || b.rev - a.rev).slice(0, BESTSELLER_TOP)
      .map((p, i) => [String(i + 1), p.name, String(p.units), bsMoney_(p.rev), bsPct_(p.units, prevUnits(p.key)), split(p)]));
  h2('Top ' + BESTSELLER_TOP + ' en chiffre d\'affaires');
  table(['#', 'Produit', 'Chiffre', 'Qté'],
    list.slice().sort((a, b) => b.rev - a.rev).slice(0, BESTSELLER_TOP).map((p, i) => [String(i + 1), p.name, bsMoney_(p.rev), String(p.units)]));

  Object.keys(cur.stores).sort().forEach(s => {
    h2('Top ' + BESTSELLER_STORE_TOP + ' — ' + s);
    table(['#', 'Produit', 'Qté', 'Chiffre'],
      list.filter(p => p.byStore[s]).sort((a, b) => b.byStore[s].units - a.byStore[s].units)
        .slice(0, BESTSELLER_STORE_TOP).map((p, i) => [String(i + 1), p.name, String(p.byStore[s].units), bsMoney_(p.byStore[s].rev)]));
  });

  const keys = Object.keys(cur.products).concat(Object.keys(prev.products)).filter((v, i, a) => a.indexOf(v) === i);
  const delta = keys.map(k => ({ k, name: (cur.products[k] || prev.products[k]).name, now: (cur.products[k] || {}).units || 0, before: prevUnits(k) }))
    .map(x => Object.assign(x, { d: x.now - x.before }));
  h2('Plus fortes hausses');
  const ups = delta.filter(x => x.d > 0).sort((a, b) => b.d - a.d).slice(0, 5);
  if (ups.length) table(['Produit', 'Cette sem.', 'Sem. préc.', 'Écart'], ups.map(x => [x.name, String(x.now), String(x.before), '+' + x.d]));
  else { T.push('(aucune)'); H.push('<div style="font:13px sans-serif">(aucune)</div>'); }
  h2('Plus fortes baisses');
  const downs = delta.filter(x => x.d < 0).sort((a, b) => a.d - b.d).slice(0, 5);
  if (downs.length) table(['Produit', 'Cette sem.', 'Sem. préc.', 'Écart'], downs.map(x => [x.name, String(x.now), String(x.before), String(x.d)]));
  else { T.push('(aucune)'); H.push('<div style="font:13px sans-serif">(aucune)</div>'); }

  h2('Meilleures ventes de la semaine précédente, ZÉRO vente cette semaine');
  const gone = Object.values(prev.products).sort((a, b) => b.units - a.units).slice(0, 30)
    .filter(p => !((cur.products[p.key] || {}).units > 0)).slice(0, 10);
  if (gone.length) table(['Produit', 'Sem. préc.', 'Cette sem.'], gone.map(p => [p.name, String(p.units), '0']));
  else { T.push('(aucun — bon signe)'); H.push('<div style="font:13px sans-serif">(aucun — bon signe)</div>'); }
  T.push('À vérifier : rupture de stock, produit déplacé, ou ventes non enregistrées ?');
  H.push('<div style="font:12px sans-serif;color:#666">À vérifier : rupture de stock, produit déplacé, ou ventes non enregistrées ?</div>');

  h2('Jours sans aucune vente enregistrée');
  const known = storeNames.filter(s => cur.stores[s] || prev.stores[s]);
  const empty = [];
  days.forEach(d => known.forEach(s => { if (!cur.days[d][s]) empty.push([fr(d), s]); }));
  if (empty.length) table(['Jour', 'Magasin'], empty);
  else { T.push('(aucun)'); H.push('<div style="font:13px sans-serif">(aucun)</div>'); }

  return { subject: '🏆 ' + title + ' — ' + bsMoney_(cr) + ' (' + bsPct_(cr, pr) + ')', text: T.join('\n'), html: '<div style="max-width:720px">' + H.join('\n') + '</div>' };
}
function weeklyBestSellers() {
  const r = buildBestSellers_(new Date());
  MailApp.sendEmail({ to: CONTROL_REPORT_EMAIL, subject: r.subject, body: r.text, htmlBody: r.html });
}
// Run from the editor to see last week's report in the log without emailing.
function testWeeklyBestSellers() { Logger.log(buildBestSellers_(new Date()).text); }

// Run ONCE from the editor (and accept the Drive permission it asks for):
// weekly backup every Sunday ~23:00, best sellers every Monday ~7:00.
function installWeeklyJobs() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['weeklyBackup', 'weeklyBestSellers'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('weeklyBackup').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).inTimezone(CONTROL_TZ).create();
  ScriptApp.newTrigger('weeklyBestSellers').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).inTimezone(CONTROL_TZ).create();
  DriveApp.getRootFolder(); // makes sure the Drive permission is asked now
}

/**
 * 23 Oct 2026 — Balkuy re-arranged: the M3 cabinet no longer exists.
 * Run retireBalkuyM3 ONCE from the Apps Script editor (▶ Exécuter): every
 * Zone Stock row for Store 1 / M3 moves to the zone "Non placé" (merged
 * with a row already there for the same product), so no stock is lost —
 * it simply waits to be counted into its new shelf.
 */
function retireBalkuyM3() {
  const n = moveZoneToUnplaced_('Store 1', 'M3');
  Logger.log(n + ' produit(s) de M3 déplacé(s) vers « Non placé ».');
  return n;
}
function moveZoneToUnplaced_(location, zone) {
  const UNPLACED = 'Non placé';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const idx = loadZoneStockIndex();
    const sheet = idx.sheet;
    const toDelete = [];
    let moved = 0;
    Object.keys(idx.byLocZoneBarcode).forEach(key => {
      const parts = key.split('||');
      if (parts[0] !== location || parts[1] !== zone) return;
      const rowNum = idx.byLocZoneBarcode[key];
      const row = sheet.getRange(rowNum, 1, 1, ZONE_STOCK_HEADERS.length).getValues()[0];
      const target = idx.byLocZoneBarcode[location + '||' + UNPLACED + '||' + parts[2]];
      if (target) {
        const t = sheet.getRange(target, 1, 1, ZONE_STOCK_HEADERS.length).getValues()[0];
        sheet.getRange(target, ZS_COL_COUNTED).setValue((Number(t[ZS_COL_COUNTED - 1]) || 0) + (Number(row[ZS_COL_COUNTED - 1]) || 0));
        sheet.getRange(target, ZS_COL_ONHAND).setValue((Number(t[ZS_COL_ONHAND - 1]) || 0) + (Number(row[ZS_COL_ONHAND - 1]) || 0));
        toDelete.push(rowNum);
      } else {
        sheet.getRange(rowNum, ZS_COL_ZONE).setValue(UNPLACED);
        sheet.getRange(rowNum, ZS_COL_CLEARANCE).setValue(false);
      }
      moved++;
    });
    toDelete.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));
    return moved;
  } finally {
    lock.releaseLock();
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  STOCK1981 — simple inventory app (Stock1981.html), added 24 Oct 2026
// ──────────────────────────────────────────────────────────────────────────
//  getStock1981 (GET, key): one snapshot for the whole app, cached offline:
//    zones   = every Zone Stock row (location, zone, barcode, name, onHand, clearance)
//    spots   = every Spot Stock row (last count per shelf space)
//    prices  = Product Master price (CFA) per barcode key
//  stock1981Adjust (POST): a batch of stock changes, each with a reason.
//    { key, action, by, userHash, ops: [ { opId, location, zone, barcode,
//      name, mode: 'delta'|'set', qty, reason, note } ] }
//    - Signed by the user's code hash (Counter Users), like the Counter.
//    - Each op changes Zone Qty On Hand of that zone (row created if missing)
//      AND the store's Qty On Hand (location tab), so both levels agree.
//    - Logged in "Stock1981 Log" (before / change / after / reason / who) and
//      in "Stock Ledger". Losses (Abîmé, Périmé, Offert, Vol) also go to "Pertes".
//    - opId makes a retried op (offline queue) apply only once.
// ══════════════════════════════════════════════════════════════════════════
const S1981_LOG_SHEET = 'Stock1981 Log';
const S1981_LOG_HEADERS = ['Timestamp', 'Location', 'Zone', 'Barcode', 'Product', 'Before', 'Change', 'After', 'Reason', 'By', 'Note', 'OpId'];
const S1981_REASONS = ['Vendu (hors caisse)', 'Abîmé', 'Périmé', 'Offert / usage boutique', 'Vol / inexpliqué', 'Recomptage', 'Ajout / réception'];
const S1981_LOSS_REASONS = ['Abîmé', 'Périmé', 'Offert / usage boutique', 'Vol / inexpliqué'];
// Accepted from the app but not offered as a manual reason: one move = a
// "Transfert" out of one zone + a "Transfert" into another (25 Oct 2026).
const S1981_INTERNAL_REASONS = ['Transfert'];

function s1981LogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(S1981_LOG_SHEET);
  if (!ws) {
    ws = ss.insertSheet(S1981_LOG_SHEET);
    ws.appendRow(S1981_LOG_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, S1981_LOG_HEADERS.length).setFontWeight('bold');
    ws.setColumnWidth(5, 220);
  }
  return ws;
}

function handleGetStock1981() {
  s1981PhotoCatchUp_();
  const zs = ensureZoneStockTab();
  const zLast = zs.getLastRow();
  const zones = zLast >= 2 ? zs.getRange(2, 1, zLast - 1, ZONE_STOCK_HEADERS.length).getValues()
    .filter(r => r[0] && r[1] && r[2])
    .map(r => [String(r[0]), String(r[1]), normalizeBarcode(r[2]), String(r[3] || ''), Number(r[5]) || 0, r[6] === true ? 1 : 0]) : [];
  const ss = ensureSpotStockTab();
  const sLast = ss.getLastRow();
  const tz = Session.getScriptTimeZone();
  const spots = sLast >= 2 ? ss.getRange(2, 1, sLast - 1, SPOT_STOCK_HEADERS.length).getValues()
    .filter(r => r[0] && r[2] && r[3])
    .map(r => [String(r[0]), String(r[1] || ''), String(r[2]), normalizeBarcode(r[3]), String(r[4] || ''), Number(r[5]) || 0,
      r[6] instanceof Date ? Utilities.formatDate(r[6], tz, "yyyy-MM-dd'T'HH:mm") : String(r[6] || ''), String(r[8] || '')]) : [];
  const prices = _exPriceMap();
  // Product Master names (and prices) by barcode key, so the app can name a
  // scanned product even if it has never been placed on a shelf.
  const names = {};
  try {
    const pm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (pm && pm.getLastRow() >= FIRST_DATA_ROW) {
      pm.getRange(FIRST_DATA_ROW, 1, pm.getLastRow() - FIRST_DATA_ROW + 1, 8).getValues().forEach(r => {
        if (r[7] && r[1]) names[barcodeKey(r[7])] = String(r[1]);
      });
    }
  } catch (e) { Logger.log('names skipped: ' + e); }
  return jsonResponse({ success: true, zones: zones, spots: spots, prices: prices, names: names, aliases: loadAliasMap_(), images: s1981Images_(), reasons: S1981_REASONS, generatedAt: new Date().toISOString() });
}

function s1981UserOk_(body) {
  const users = readCounterUsers_();
  const u = users.find(x => x.active && x.name.toLowerCase() === String(body.by || '').toLowerCase());
  return !!(u && u.hash && u.hash === String(body.userHash || ''));
}

function handleStock1981Adjust(body) {
  if (!s1981UserOk_(body)) return jsonResponse({ success: false, error: "Utilisateur ou code non reconnu — reconnectez-vous." }, 403);
  const ops = Array.isArray(body.ops) ? body.ops : [];
  if (!ops.length) return jsonResponse({ success: false, error: "Aucun changement" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const log = s1981LogSheet_();
    const done = {};
    const lLast = log.getLastRow();
    if (lLast >= 2) {
      const from = Math.max(2, lLast - 2000);
      log.getRange(from, S1981_LOG_HEADERS.length, lLast - from + 1, 1).getValues().forEach(r => { if (r[0]) done[String(r[0])] = true; });
    }
    const idx = loadZoneStockIndex();
    const zsheet = idx.sheet;
    const tz = Session.getScriptTimeZone();
    const now = new Date();
    const stamp = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');
    const prices = _exPriceMap();
    const logRows = [], ledgerRows = [], lossRows = [], results = [];
    const locSheets = {};

    ops.forEach(op => {
      const opId = String(op.opId || '');
      const loc = String(op.location || '');
      const zone = String(op.zone || '').trim();
      const bc = normalizeBarcode(resolveAlias_(op.barcode));
      const reason = String(op.reason || '');
      if (opId && done[opId]) { results.push({ opId: opId, ok: true, duplicate: true }); return; }
      if (VALID_LOCATIONS.indexOf(loc) === -1 || !zone || !bc) { results.push({ opId: opId, ok: false, error: 'Données incomplètes' }); return; }
      if (S1981_REASONS.indexOf(reason) === -1 && S1981_INTERNAL_REASONS.indexOf(reason) === -1) { results.push({ opId: opId, ok: false, error: 'Raison inconnue' }); return; }
      const qty = Math.round(Number(op.qty));
      if (isNaN(qty)) { results.push({ opId: opId, ok: false, error: 'Quantité invalide' }); return; }

      // Zone level
      const key = loc + '||' + zone + '||' + barcodeKey(bc);
      let rowNum = idx.byLocZoneBarcode[key];
      let before = 0, name = String(op.name || '');
      if (rowNum) {
        const r = zsheet.getRange(rowNum, 1, 1, ZONE_STOCK_HEADERS.length).getValues()[0];
        before = Number(r[ZS_COL_ONHAND - 1]) || 0;
        name = name || String(r[ZS_COL_PRODUCT - 1] || '');
      }
      const delta = op.mode === 'set' ? qty - before : qty;
      if (!delta) { results.push({ opId: opId, ok: true, unchanged: true, after: before }); return; }
      const after = before + delta;
      if (rowNum) {
        zsheet.getRange(rowNum, ZS_COL_ONHAND).setValue(after);
        if (op.mode === 'set') zsheet.getRange(rowNum, ZS_COL_COUNTED).setValue(after);
      } else {
        rowNum = zsheet.getLastRow() + 1;
        zsheet.getRange(rowNum, ZS_COL_BARCODE).setNumberFormat('@');
        zsheet.getRange(rowNum, 1, 1, ZONE_STOCK_HEADERS.length)
          .setValues([[loc, zone, bc, name, op.mode === 'set' ? after : 0, after, false]]);
        idx.byLocZoneBarcode[key] = rowNum;
      }

      // Store level (location tab "Qty On Hand")
      let storeAfter = '';
      try {
        if (!locSheets[loc]) {
          const sh = getLocationSheetOrNull(loc);
          if (sh && sh.getLastRow() >= FIRST_DATA_ROW) {
            const qtyCol = getOrCreateColumnByHeader(sh, HEADER_ROW, QTY_ON_HAND_HEADER);
            const n = sh.getLastRow() - FIRST_DATA_ROW + 1;
            const bcs = sh.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, n, 1).getValues();
            const map = {};
            bcs.forEach((r, i) => { const k = barcodeKey(r[0]); if (k) map[k] = FIRST_DATA_ROW + i; });
            locSheets[loc] = { sheet: sh, qtyCol: qtyCol, map: map };
          } else locSheets[loc] = null;
        }
        const L = locSheets[loc];
        const lr = L && L.map[barcodeKey(bc)];
        if (lr) {
          const cell = L.sheet.getRange(lr, L.qtyCol);
          storeAfter = (Number(cell.getValue()) || 0) + delta;
          cell.setValue(storeAfter);
        }
      } catch (e) { Logger.log('store-level skipped: ' + e); }

      const who = String(body.by || '');
      const note = String(op.note || '');
      logRows.push([stamp, loc, zone, bc, name, before, delta, after, reason, who, note, opId]);
      ledgerRows.push([stamp, loc, bc, name || bc, 'stock1981: ' + reason, delta, storeAfter === '' ? after : storeAfter,
        'Zone ' + zone + ' — ' + who + (note ? ' — ' + note : '')]);
      if (delta < 0 && S1981_LOSS_REASONS.indexOf(reason) !== -1) {
        const price = prices[barcodeKey(bc)];
        lossRows.push([stamp, Utilities.formatDate(now, tz, 'yyyy-MM-dd'), loc, zone, bc, name, -delta,
          price == null ? '' : price, price == null ? '' : price * -delta, reason, '', who, note, opId]);
      }
      if (opId) done[opId] = true;
      results.push({ opId: opId, ok: true, before: before, after: after, storeLevel: storeAfter !== '' });
    });

    if (logRows.length) {
      const r0 = log.getLastRow() + 1;
      log.getRange(r0, 4, logRows.length, 1).setNumberFormat('@');
      log.getRange(r0, 1, logRows.length, S1981_LOG_HEADERS.length).setValues(logRows);
    }
    appendLedgerRows(ledgerRows);
    if (lossRows.length) {
      const lw = _lossSheet();
      lw.getRange(lw.getLastRow() + 1, 1, lossRows.length, LOSS_HEADERS.length).setValues(lossRows);
    }
    s1981QueuePhotos_(ops.filter(op => Number(op.qty) > 0 && op.mode !== 'set' && op.reason === 'Ajout / réception').map(op => ({ bc: normalizeBarcode(resolveAlias_(op.barcode)), name: String(op.name || '') })));
    return jsonResponse({ success: true, results: results });
  } finally {
    lock.releaseLock();
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  STOCK1981 — shelf count (action "stock1981Count"), added 24 Oct 2026
// ──────────────────────────────────────────────────────────────────────────
//  { key, action, by, userHash, opId, location, zone, spot, at,
//    items: [ { barcode, name, count } ] }
//  - The shelf's list in "Spot Stock" is replaced by this count (what was on
//    that shelf before and isn't scanned now = no longer there).
//  - For every product on that shelf (before or now), the zone's stock
//    ("Zone Qty On Hand") becomes the sum of its latest counts on all the
//    zone's shelves. A product kept on one shelf (most of them) is therefore
//    exactly what was just counted.
//  - Each change is logged in "Stock1981 Log" with reason "Comptage" and the
//    stock before / after, so gaps can be reviewed.
// ══════════════════════════════════════════════════════════════════════════
function handleStock1981Count(body) {
  if (!s1981UserOk_(body)) return jsonResponse({ success: false, error: "Utilisateur ou code non reconnu — reconnectez-vous." }, 403);
  const loc = String(body.location || ''), zone = String(body.zone || '').trim(), spot = String(body.spot || '').trim();
  if (VALID_LOCATIONS.indexOf(loc) === -1 || !zone || !spot) return jsonResponse({ success: false, error: "Magasin, zone ou étagère manquant" }, 400);
  const items = (Array.isArray(body.items) ? body.items : [])
    .map(it => ({ bc: normalizeBarcode(resolveAlias_(it.barcode)), name: String(it.name || ''), count: Math.max(0, Math.round(Number(it.count) || 0)) }))
    .filter(it => it.bc && it.count > 0);
  const opId = String(body.opId || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const log = s1981LogSheet_();
    const lLast = log.getLastRow();
    if (opId && lLast >= 2) {
      const from = Math.max(2, lLast - 3000);
      const ids = log.getRange(from, S1981_LOG_HEADERS.length, lLast - from + 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === opId) return jsonResponse({ success: true, duplicate: true });
    }
    // 1. Shelf contents before this count
    const ss = ensureSpotStockTab();
    const sLast = ss.getLastRow();
    const spotRows = sLast >= 2 ? ss.getRange(2, 1, sLast - 1, SPOT_STOCK_HEADERS.length).getValues() : [];
    const touched = {};   // barcodeKey -> {bc, name}
    spotRows.forEach(r => {
      if (String(r[0]) === loc && String(r[2]) === spot && r[3]) touched[barcodeKey(r[3])] = { bc: normalizeBarcode(r[3]), name: String(r[4] || '') };
    });
    items.forEach(it => { touched[barcodeKey(it.bc)] = { bc: it.bc, name: it.name || (touched[barcodeKey(it.bc)] || {}).name || '' }; });

    // 2. Replace the shelf in Spot Stock
    const when = body.at ? new Date(body.at) : new Date();
    const payload = {};
    payload[spot] = { zone: zone, at: isNaN(when.getTime()) ? new Date().toISOString() : when.toISOString(), by: String(body.by || ''),
      items: items.map(it => ({ barcode: it.bc, name: it.name, count: it.count })) };
    applySpotCountsUnlocked_(loc, payload, 'Stock1981', String(body.by || ''));

    // 3. Sum of latest counts on every shelf of this zone, for the touched products
    const after = ss.getRange(2, 1, Math.max(1, ss.getLastRow() - 1), SPOT_STOCK_HEADERS.length).getValues();
    const sum = {};
    Object.keys(touched).forEach(k => { sum[k] = 0; });
    after.forEach(r => {
      if (String(r[0]) !== loc || String(r[1]) !== zone || !r[3]) return;
      const k = barcodeKey(r[3]);
      if (k in sum) sum[k] += Number(r[5]) || 0;
    });

    // 4. Zone Stock + log
    const idx = loadZoneStockIndex();
    const zs = idx.sheet;
    const tz = Session.getScriptTimeZone();
    const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    const who = String(body.by || '');
    const logRows = [], newRows = [];
    let changed = 0;
    Object.keys(touched).forEach(k => {
      const t = touched[k], target = sum[k];
      const rowNum = idx.byLocZoneBarcode[loc + '||' + zone + '||' + k];
      let before = 0;
      if (rowNum) {
        before = Number(zs.getRange(rowNum, ZS_COL_ONHAND).getValue()) || 0;
        zs.getRange(rowNum, ZS_COL_COUNTED, 1, 2).setValues([[target, target]]);
        if (t.name) zs.getRange(rowNum, ZS_COL_PRODUCT).setValue(t.name);
      } else if (target > 0) {
        newRows.push([loc, zone, t.bc, t.name, target, target, false]);
      } else return;
      if (before !== target) changed++;
      logRows.push([stamp, loc, zone, t.bc, t.name, before, target - before, target, 'Comptage', who, 'Étagère ' + spot, opId]);
    });
    if (newRows.length) {
      const r0 = zs.getLastRow() + 1;
      zs.getRange(r0, ZS_COL_BARCODE, newRows.length, 1).setNumberFormat('@');
      zs.getRange(r0, 1, newRows.length, ZONE_STOCK_HEADERS.length).setValues(newRows);
    }
    if (!logRows.length) logRows.push([stamp, loc, zone, '', '(étagère vide)', 0, 0, 0, 'Comptage', who, 'Étagère ' + spot, opId]);
    const r1 = log.getLastRow() + 1;
    log.getRange(r1, 4, logRows.length, 1).setNumberFormat('@');
    log.getRange(r1, 1, logRows.length, S1981_LOG_HEADERS.length).setValues(logRows);
    s1981QueuePhotos_(items.map(it => ({ bc: it.bc, name: it.name })));
    return jsonResponse({ success: true, products: items.length, changed: changed });
  } finally {
    lock.releaseLock();
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  IMPORT ZONE STOCK (action "importZoneStock"), added 24 Oct 2026
//  Called by the sales script's "Nouveaux produits" tool (server to server,
//  write key only) after a product from a scanned list is added or known:
//  { key, action, location, zone, by, items: [ { barcode, name, qty } ] }
//  Sets that zone's stock to the scanned quantity. The store total is NOT
//  touched here (the caller already set it with updateCounts /
//  addNewProductWithStock). Logged in "Stock1981 Log" (Ajout / réception).
// ══════════════════════════════════════════════════════════════════════════
function handleImportZoneStock(body) {
  const loc = String(body.location || ''), zone = String(body.zone || '').trim();
  if (VALID_LOCATIONS.indexOf(loc) === -1 || !zone) return jsonResponse({ success: false, error: "Magasin ou zone manquant" }, 400);
  const items = (Array.isArray(body.items) ? body.items : []).filter(it => normalizeBarcode(it.barcode));
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const idx = loadZoneStockIndex();
    const zs = idx.sheet;
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const logRows = [], newRows = [];
    items.forEach(it => {
      const bc = normalizeBarcode(it.barcode), k = barcodeKey(bc), qty = Math.max(0, Math.round(Number(it.qty) || 0));
      const rowNum = idx.byLocZoneBarcode[loc + '||' + zone + '||' + k];
      let before = 0;
      if (rowNum) {
        before = Number(zs.getRange(rowNum, ZS_COL_ONHAND).getValue()) || 0;
        zs.getRange(rowNum, ZS_COL_COUNTED, 1, 2).setValues([[qty, qty]]);
        if (it.name) zs.getRange(rowNum, ZS_COL_PRODUCT).setValue(String(it.name));
      } else {
        newRows.push([loc, zone, bc, String(it.name || ''), qty, qty, false]);
        idx.byLocZoneBarcode[loc + '||' + zone + '||' + k] = -1;
      }
      logRows.push([stamp, loc, zone, bc, String(it.name || ''), before, qty - before, qty, 'Ajout / réception', String(body.by || 'Import'), 'Import des scans', '']);
    });
    if (newRows.length) {
      const r0 = zs.getLastRow() + 1;
      zs.getRange(r0, ZS_COL_BARCODE, newRows.length, 1).setNumberFormat('@');
      zs.getRange(r0, 1, newRows.length, ZONE_STOCK_HEADERS.length).setValues(newRows);
    }
    if (logRows.length) {
      const log = s1981LogSheet_();
      const r1 = log.getLastRow() + 1;
      log.getRange(r1, 4, logRows.length, 1).setNumberFormat('@');
      log.getRange(r1, 1, logRows.length, S1981_LOG_HEADERS.length).setValues(logRows);
    }
    return jsonResponse({ success: true, items: logRows.length });
  } finally {
    lock.releaseLock();
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  LABELS TO PRINT — added 25 Oct 2026
// ──────────────────────────────────────────────────────────────────────────
//  Stock1981 (on the phone) moves stock from the Entrepôt to a store and asks
//  for that many stickers; it can also ask for one box label (code, name,
//  quantity in the box). Those requests wait in "Étiquettes à imprimer"
//  until the Scanner on the computer with the label printer prints them.
//  stock1981Labels   (POST, user-signed) { by, userHash, labels:[{id, barcode,
//                     name, qty, store, kind:'unit'|'box'}] } — id makes a
//                     retried request count once.
//  getPrintQueue     (GET, key, store?) — labels still to print.
//  markLabelsPrinted (POST, key) { ids:[…] } — after printing.
// ══════════════════════════════════════════════════════════════════════════
const S1981_LABEL_SHEET = 'Étiquettes à imprimer';
const S1981_LABEL_HEADERS = ['Id', 'Date', 'Barcode', 'Produit', 'Qté', 'Magasin', 'Type', 'Par', 'Statut', 'Imprimé le'];
const S1981_LABEL_TODO = 'À imprimer', S1981_LABEL_DONE = 'Imprimé';

function s1981LabelSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(S1981_LABEL_SHEET);
  if (!ws) {
    ws = ss.insertSheet(S1981_LABEL_SHEET);
    ws.appendRow(S1981_LABEL_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, S1981_LABEL_HEADERS.length).setFontWeight('bold');
    ws.getRange('C:C').setNumberFormat('@');
    ws.setColumnWidth(4, 240);
  }
  return ws;
}

function handleStock1981Labels(body) {
  if (!s1981UserOk_(body)) return jsonResponse({ success: false, error: "Utilisateur ou code non reconnu — reconnectez-vous." }, 403);
  const labels = Array.isArray(body.labels) ? body.labels : [];
  if (!labels.length) return jsonResponse({ success: false, error: "Aucune étiquette" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const ws = s1981LabelSheet_();
    const last = ws.getLastRow();
    const seen = {};
    if (last >= 2) {
      const from = Math.max(2, last - 3000);
      ws.getRange(from, 1, last - from + 1, 1).getValues().forEach(r => { if (r[0]) seen[String(r[0])] = true; });
    }
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const rows = [];
    labels.forEach(l => {
      const id = String(l.id || '');
      const bc = normalizeBarcode(l.barcode);
      const qty = Math.round(Number(l.qty));
      if (!id || seen[id] || !bc || !(qty > 0) || qty > 2000) return;
      seen[id] = true;
      rows.push([id, stamp, bc, String(l.name || ''), qty, String(l.store || ''), l.kind === 'box' ? 'Carton' : 'Unités',
        String(body.by || ''), S1981_LABEL_TODO, '']);
    });
    if (rows.length) ws.getRange(ws.getLastRow() + 1, 1, rows.length, S1981_LABEL_HEADERS.length).setValues(rows);
    return jsonResponse({ success: true, added: rows.length });
  } finally {
    lock.releaseLock();
  }
}

function handleGetPrintQueue(store) {
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S1981_LABEL_SHEET);
  const out = [];
  if (ws && ws.getLastRow() >= 2) {
    ws.getRange(2, 1, ws.getLastRow() - 1, S1981_LABEL_HEADERS.length).getDisplayValues().forEach(r => {
      if (r[8] !== S1981_LABEL_TODO) return;
      if (store && r[5] !== store) return;
      out.push({ id: r[0], date: r[1], barcode: r[2], name: r[3], qty: Number(r[4]) || 0, store: r[5], kind: r[6] === 'Carton' ? 'box' : 'unit', by: r[7] });
    });
  }
  return jsonResponse({ success: true, labels: out });
}

function handleMarkLabelsPrinted(body) {
  const ids = {};
  (Array.isArray(body.ids) ? body.ids : []).forEach(i => { ids[String(i)] = true; });
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S1981_LABEL_SHEET);
  if (!ws || ws.getLastRow() < 2) return jsonResponse({ success: true, marked: 0 });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const n = ws.getLastRow() - 1;
    const vals = ws.getRange(2, 1, n, S1981_LABEL_HEADERS.length).getValues();
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let marked = 0;
    vals.forEach(r => { if (ids[String(r[0])] && r[8] === S1981_LABEL_TODO) { r[8] = S1981_LABEL_DONE; r[9] = stamp; marked++; } });
    if (marked) ws.getRange(2, 9, n, 2).setValues(vals.map(r => [r[8], r[9]]));
    return jsonResponse({ success: true, marked: marked });
  } finally {
    lock.releaseLock();
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  SECOND CODES (aliases) — added 26 Oct 2026
// ──────────────────────────────────────────────────────────────────────────
//  A made-up code (0610000…) can be replaced by the product's real barcode
//  while the old stickers keep scanning: the old code becomes a "second
//  code" of the real one. Kept in the "Codes secondaires" tab.
//  - The product feed (doGet) leaves the old Product Master row out and
//    sends `aliases` {oldKey: realBarcode}; the Scanner and Stock1981 swap a
//    scanned old code for the real one before anything is recorded.
//  - Stock movements, Stock1981 changes and shelf counts map old codes too,
//    in case an old device or an offline queue still sends them.
//  - applyCodeReplacements() applies the rows ticked « Remplacer par le code
//    réel » in the « Codes maison → codes réels » tab: adds the second code,
//    copies the price if the real product has none, and moves stock rows
//    (store tabs, Zone Stock, Spot Stock) from the old code to the real one.
//    Runnable from the editor or POST {action:"applyCodeReplacements"}.
// ══════════════════════════════════════════════════════════════════════════
const ALIAS_SHEET = 'Codes secondaires';
const ALIAS_HEADERS = ['Code secondaire', 'Code principal', 'Produit', 'Ajouté le', 'Par', 'Note'];
const REPL_SHEET = 'Codes maison → codes réels';
const REPL_CHOICE = 'Remplacer par le code réel';
let _aliasCache = null;

function aliasSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(ALIAS_SHEET);
  if (!ws) {
    ws = ss.insertSheet(ALIAS_SHEET);
    ws.appendRow(ALIAS_HEADERS);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, ALIAS_HEADERS.length).setFontWeight('bold');
    ws.getRange('A:B').setNumberFormat('@');
    ws.setColumnWidth(3, 260);
  }
  return ws;
}

function loadAliasMap_() {
  if (_aliasCache) return _aliasCache;
  const map = {};
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ALIAS_SHEET);
  if (ws && ws.getLastRow() >= 2) {
    ws.getRange(2, 1, ws.getLastRow() - 1, 2).getValues().forEach(r => {
      if (r[0] && r[1] && barcodeKey(r[0]) !== barcodeKey(r[1])) map[barcodeKey(r[0])] = normalizeBarcode(r[1]);
    });
  }
  _aliasCache = map;
  return map;
}

function resolveAlias_(bc) {
  if (bc === null || bc === undefined || bc === '') return bc;
  const real = loadAliasMap_()[barcodeKey(bc)];
  return real || bc;
}

function applyCodeReplacements() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rs = ss.getSheetByName(REPL_SHEET);
  if (!rs || rs.getLastRow() < 5) return { success: false, error: "Onglet « " + REPL_SHEET + " » introuvable ou vide." };
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const n = rs.getLastRow() - 4;
    const rows = rs.getRange(5, 1, n, 11).getValues();
    const tz = Session.getScriptTimeZone();
    const stamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
    const pm = ss.getSheetByName(SHEET_NAME);
    const pmData = pm.getRange(FIRST_DATA_ROW, 1, pm.getLastRow() - FIRST_DATA_ROW + 1, 10).getValues();
    const pmRow = {};
    pmData.forEach((r, i) => { if (r[7]) { const k = barcodeKey(r[7]); if (!pmRow[k]) pmRow[k] = { row: FIRST_DATA_ROW + i, name: String(r[1] || ''), price: r[9] }; } });
    const aws = aliasSheet_();
    const existing = loadAliasMap_();
    const done = [], notes = [];
    rows.forEach((r, i) => {
      const choice = String(r[9] || ''), note = String(r[10] || '');
      if (choice !== REPL_CHOICE || /Appliqué/.test(note)) return;
      const oldBc = normalizeBarcode(r[1]), realBc = normalizeBarcode(r[5]);
      if (!oldBc || !realBc) return;
      const oldK = barcodeKey(oldBc), realK = barcodeKey(realBc);
      const real = pmRow[realK], old = pmRow[oldK];
      if (!real) { notes.push([i, 'Code réel absent de Product Master']); return; }
      // 1. second code
      if (!existing[oldK]) {
        aws.getRange(aws.getLastRow() + 1, 1, 1, ALIAS_HEADERS.length)
          .setValues([[oldBc, realBc, real.name, stamp, 'Codes maison → codes réels', old ? 'Ancien nom : ' + old.name : '']]);
        existing[oldK] = realBc;
      }
      // 2. price: keep the real product's, else take the old one's
      if ((real.price === '' || real.price === null) && old && old.price !== '' && old.price !== null) {
        pm.getRange(real.row, 10).setValue(old.price);
      }
      try { _moveStoreStock_(oldK, realBc, real.name); } catch (e) { Logger.log('store stock: ' + e); }
      try { _moveZoneRows_(oldK, realBc, real.name); } catch (e) { Logger.log('zone stock: ' + e); }
      notes.push([i, 'Appliqué ' + stamp]);
      done.push(oldBc + ' → ' + realBc);
    });
    notes.forEach(([i, txt]) => rs.getRange(5 + i, 11).setValue(txt));
    _aliasCache = null;
    return { success: true, applied: done.length, details: done };
  } finally {
    lock.releaseLock();
  }
}

// Store tabs: the old code's Qty On Hand goes onto the real code's row.
function _moveStoreStock_(oldK, realBc, realName) {
  VALID_LOCATIONS.forEach(loc => {
    const sh = getLocationSheetOrNull(loc);
    if (!sh || sh.getLastRow() < FIRST_DATA_ROW) return;
    const qtyCol = getOrCreateColumnByHeader(sh, HEADER_ROW, QTY_ON_HAND_HEADER);
    const n = sh.getLastRow() - FIRST_DATA_ROW + 1;
    const bcs = sh.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, n, 1).getValues();
    let oldRow = 0, realRow = 0;
    bcs.forEach((r, i) => { const k = barcodeKey(r[0]); if (k === oldK && !oldRow) oldRow = FIRST_DATA_ROW + i; if (k === barcodeKey(realBc) && !realRow) realRow = FIRST_DATA_ROW + i; });
    if (!oldRow) return;
    const oldQty = Number(sh.getRange(oldRow, qtyCol).getValue()) || 0;
    if (realRow) {
      if (oldQty) {
        const c = sh.getRange(realRow, qtyCol);
        c.setValue((Number(c.getValue()) || 0) + oldQty);
        sh.getRange(oldRow, qtyCol).setValue(0);
      }
    } else {
      sh.getRange(oldRow, LOCATION_BARCODE_COL).setNumberFormat('@').setValue(realBc);
      sh.getRange(oldRow, LOCATION_NAME_COL).setValue(realName);
    }
  });
}

// Zone Stock and Spot Stock rows of the old code are relabelled to the real
// code (merged into an existing row of the same zone / shelf when there is one).
function _moveZoneRows_(oldK, realBc, realName) {
  const realK = barcodeKey(realBc);
  const zs = ensureZoneStockTab();
  if (zs.getLastRow() >= 2) {
    const n = zs.getLastRow() - 1;
    const v = zs.getRange(2, 1, n, ZONE_STOCK_HEADERS.length).getValues();
    const at = {};
    v.forEach((r, i) => { if (barcodeKey(r[2]) === realK) at[r[0] + '|' + r[1]] = i; });
    v.forEach((r, i) => {
      if (barcodeKey(r[2]) !== oldK) return;
      const key = r[0] + '|' + r[1];
      if (at[key] !== undefined) {
        const t = v[at[key]];
        t[ZS_COL_COUNTED - 1] = (Number(t[ZS_COL_COUNTED - 1]) || 0) + (Number(r[ZS_COL_COUNTED - 1]) || 0);
        t[ZS_COL_ONHAND - 1] = (Number(t[ZS_COL_ONHAND - 1]) || 0) + (Number(r[ZS_COL_ONHAND - 1]) || 0);
        r[ZS_COL_COUNTED - 1] = 0; r[ZS_COL_ONHAND - 1] = 0;
      } else {
        r[2] = realBc; r[3] = realName; at[key] = i;
      }
    });
    zs.getRange(2, ZS_COL_BARCODE, n, 1).setNumberFormat('@');
    zs.getRange(2, 1, n, ZONE_STOCK_HEADERS.length).setValues(v);
  }
  const sp = ensureSpotStockTab();
  if (sp.getLastRow() >= 2) {
    const n = sp.getLastRow() - 1;
    const v = sp.getRange(2, 1, n, SPOT_STOCK_HEADERS.length).getValues();
    let changed = false;
    v.forEach(r => { if (barcodeKey(r[3]) === oldK) { r[3] = realBc; r[4] = realName; changed = true; } });
    if (changed) {
      sp.getRange(2, 4, n, 1).setNumberFormat('@');
      sp.getRange(2, 1, n, SPOT_STOCK_HEADERS.length).setValues(v);
    }
  }
}


// ══════════════════════════════════════════════════════════════════════════
//  COST PRICE (USD) — added 26 Oct 2026
//  Product Master column I « Cost (USD) ». Read with getPriceList when the
//  write key is given (the Scanner's Admin list); written with POST
//  {action:"updateCost", barcode, cost, by}. Logged in "Cost Edits".
// ══════════════════════════════════════════════════════════════════════════
const COST_COL = 9; // I
function _costMap_() {
  const pm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const out = {};
  if (!pm || pm.getLastRow() < FIRST_DATA_ROW) return out;
  pm.getRange(FIRST_DATA_ROW, 8, pm.getLastRow() - FIRST_DATA_ROW + 1, 2).getValues().forEach(r => {
    if (!r[0]) return;
    const k = barcodeKey(r[0]);
    const v = Number(String(r[1]).replace(',', '.'));
    if (r[1] !== '' && r[1] !== null && isFinite(v) && out[k] == null) out[k] = v;
  });
  return out;
}
function handleUpdateCost(body) {
  const k = barcodeKey(resolveAlias_(body.barcode));
  const cost = body.cost === '' || body.cost === null || body.cost === undefined ? '' : Number(String(body.cost).replace(',', '.'));
  if (!k || (cost !== '' && (!isFinite(cost) || cost < 0 || cost > 100000))) return jsonResponse({ success: false, error: "Code ou coût invalide" }, 400);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pm = ss.getSheetByName(SHEET_NAME);
  const n = pm.getLastRow() - FIRST_DATA_ROW + 1;
  const vals = pm.getRange(FIRST_DATA_ROW, 1, n, 9).getValues();
  let hit = 0, name = '', old = '';
  vals.forEach((r, i) => {
    if (r[7] && barcodeKey(r[7]) === k) {
      if (!hit) { name = String(r[1] || ''); old = r[8]; }
      pm.getRange(FIRST_DATA_ROW + i, COST_COL).setValue(cost);
      hit++;
    }
  });
  if (!hit) return jsonResponse({ success: false, error: "Produit introuvable dans Product Master" }, 404);
  let log = ss.getSheetByName('Cost Edits');
  if (!log) { log = ss.insertSheet('Cost Edits'); log.appendRow(['Timestamp', 'Barcode', 'Product', 'Old Cost (USD)', 'New Cost (USD)', 'By']); log.setFrozenRows(1); log.getRange('B:B').setNumberFormat('@'); }
  log.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'), normalizeBarcode(body.barcode), name, old, cost, String(body.by || '')]);
  return jsonResponse({ success: true, cost: cost, rows: hit });
}


// ══════════════════════════════════════════════════════════════════════════
//  MOUVEMENTS (action "getMovements", GET, key required), added 27 Oct 2026
// ──────────────────────────────────────────────────────────────────────────
//  Read-only, for Scanner Admin → « 📦 Mouvements ».
//  - moves: every Déplacer from Stock1981 of the last N days (default 30).
//    A Déplacer is logged in "Stock1981 Log" as two "Transfert" rows written
//    together (out of one zone, then into another); they are paired back here.
//  - unplaced: what is still in the zone « Non placé » (Zone Stock), with the
//    time it last arrived there (Stock1981 Log or Stock Ledger), so the app can
//    flag stock waiting more than 24 h to be shelved.
//  Nothing is written.
// ═════════════════════════════════════════════════════════════════════════
const MV_UNPLACED = 'Non placé';

function mvParse_(v) {
  if (v instanceof Date) return v;
  const m = String(v || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0));
}

function handleGetMovements(days) {
  const nDays = Math.max(1, Math.min(180, Number(days) || 30));
  const since = new Date(Date.now() - nDays * 86400000);
  const tz = Session.getScriptTimeZone();
  const iso = d => Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lastIn = {}; // loc||barcodeKey → latest Date it arrived in « Non placé »
  const noteIn = (loc, bc, d) => { const k = loc + '||' + barcodeKey(bc); if (!lastIn[k] || d > lastIn[k]) lastIn[k] = d; };

  // 1) Déplacer pairs from Stock1981 Log
  const moves = [];
  const log = ss.getSheetByName(S1981_LOG_SHEET);
  if (log && log.getLastRow() >= 2) {
    const from = Math.max(2, log.getLastRow() - 6000);
    const rows = log.getRange(from, 1, log.getLastRow() - from + 1, S1981_LOG_HEADERS.length).getValues();
    // columns: 0 Timestamp, 1 Location, 2 Zone, 3 Barcode, 4 Product, 5 Before, 6 Change, 7 After, 8 Reason, 9 By, 10 Note, 11 OpId
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (String(r[8]) !== 'Transfert') continue;
      const d = mvParse_(r[0]);
      const ch = Number(r[6]) || 0;
      if (ch > 0 && String(r[2]) === MV_UNPLACED && d) noteIn(String(r[1]), r[3], d);
      if (!d || d < since) continue;
      const bc = normalizeBarcode(r[3]);
      const n = rows[i + 1];
      const paired = ch < 0 && n && String(n[8]) === 'Transfert' && barcodeKey(n[3]) === barcodeKey(r[3]) && (Number(n[6]) || 0) === -ch;
      if (paired) {
        moves.push({ at: iso(d), bc: bc, name: String(r[4] || n[4] || ''), qty: -ch,
          fromLoc: String(r[1]), fromZone: String(r[2]), toLoc: String(n[1]), toZone: String(n[2]), by: String(r[9] || '') });
        const dn = mvParse_(n[0]);
        if ((Number(n[6]) || 0) > 0 && String(n[2]) === MV_UNPLACED && dn) noteIn(String(n[1]), n[3], dn);
        i++;
      } else {
        // a half on its own (e.g. the other half failed) — still shown
        moves.push({ at: iso(d), bc: bc, name: String(r[4] || ''), qty: Math.abs(ch), half: true,
          fromLoc: ch < 0 ? String(r[1]) : '', fromZone: ch < 0 ? String(r[2]) : '',
          toLoc: ch > 0 ? String(r[1]) : '', toZone: ch > 0 ? String(r[2]) : '', by: String(r[9] || ''), note: String(r[10] || '') });
      }
    }
  }

  // 2) Transfers received from the Scanner land in « Non placé » through the Stock Ledger
  const led = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (led && led.getLastRow() >= 2) {
    const from = Math.max(2, led.getLastRow() - 8000);
    led.getRange(from, 1, led.getLastRow() - from + 1, 8).getValues().forEach(r => {
      if ((Number(r[5]) || 0) <= 0) return;
      if (String(r[7] || '').indexOf('Zone ' + MV_UNPLACED) !== 0) return;
      const d = mvParse_(r[0]); if (d) noteIn(String(r[1]), r[2], d);
    });
  }

  // 3) What is still waiting in « Non placé »
  const unplaced = [];
  const zs = ss.getSheetByName(ZONE_STOCK_SHEET_NAME);
  if (zs && zs.getLastRow() >= 2) {
    zs.getRange(2, 1, zs.getLastRow() - 1, ZONE_STOCK_HEADERS.length).getValues().forEach(r => {
      if (String(r[ZS_COL_ZONE - 1]) !== MV_UNPLACED) return;
      const q = Number(r[ZS_COL_ONHAND - 1]) || 0;
      if (q <= 0) return;
      const loc = String(r[ZS_COL_LOCATION - 1]);
      const d = lastIn[loc + '||' + barcodeKey(r[ZS_COL_BARCODE - 1])];
      unplaced.push({ loc: loc, bc: normalizeBarcode(r[ZS_COL_BARCODE - 1]), name: String(r[ZS_COL_PRODUCT - 1] || ''), qty: q, since: d ? iso(d) : '' });
    });
  }

  moves.sort((a, b) => a.at < b.at ? 1 : (a.at > b.at ? -1 : 0));
  return jsonResponse({ success: true, days: nDays, moves: moves, unplaced: unplaced, generatedAt: new Date().toISOString() });
}


// ══════════════════════════════════════════════════════════════════════════
//  RÉCEPTION D'UN ARRIVAGE (container), added 28 Oct 2026 — Stock1981 « 📦 Réception »
// ──────────────────────────────────────────────────────────────────────────
//  An arrival (« Arrivage ») is emptied one cooler / box at a time, blind:
//  the phone scans everything inside, then sends the whole unit at once.
//    POST recvArrival  { by, userHash, id, name, note }         create / rename
//    POST recvUnit     { by, userHash, opId, arrivalId, arrivalName, unit, kind,
//                        cooler: { barcode, name } | null,
//                        lines: [ { lineId, barcode, name, qty, mode: 'date'|'none'|'later',
//                                   expiry: 'DD/MM/YYYY', solde: bool } ] }
//    POST recvSetDate  { by, userHash, lineId, mode, expiry }   fill a « plus tard » date
//    GET  getArrivals                                            arrivals + units + lines
//  Effects of recvUnit:
//    - stock goes into Warehouse / « Entrepôt » (zone + store level + ledger +
//      Stock1981 Log) through handleStock1981Adjust, reason « Ajout / réception »;
//    - an emptied cooler that is for sale is added as +1 of its own product;
//    - products the Sheet doesn't know are added to Product Master (no price yet)
//      and to the Warehouse tab;
//    - dates go to « Expiry Batches » (Shipment = arrival name), same tab the
//      Scanner's Péremptions card reads; « ne périme pas » is recorded once per
//      product so it is never asked again;
//    - every line is kept in « Réception » and every unit in « Réception Unités ».
//  opId makes a retried unit (offline queue) apply only once.
// ══════════════════════════════════════════════════════════════════════════
const RECV_LOC = 'Warehouse', RECV_ZONE = 'Entrepôt';
const RECV_ARR_SHEET = 'Arrivages';
const RECV_ARR_HEADERS = ['ID', 'Nom', 'Créé le', 'Lieu', 'Statut', 'Par', 'Note'];
const RECV_UNIT_SHEET = 'Réception Unités';
const RECV_UNIT_HEADERS = ['Timestamp', 'Arrivage ID', 'Arrivage', 'Unité', 'Type', 'Glacière code', 'Glacière nom', 'Produits', 'Articles', 'Par', 'OpId'];
const RECV_LINE_SHEET = 'Réception';
const RECV_LINE_HEADERS = ['Timestamp', 'Arrivage ID', 'Arrivage', 'Unité', 'Barcode', 'Produit', 'Qté', 'Date', 'Péremption', 'Solde', 'Par', 'LineId'];
const RL = { TS: 0, AID: 1, ANAME: 2, UNIT: 3, BC: 4, NAME: 5, QTY: 6, MODE: 7, EXP: 8, SOLDE: 9, BY: 10, LID: 11 };

function recvSheet_(name, headers, textCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(name);
  if (!ws) {
    ws = ss.insertSheet(name);
    ws.appendRow(headers);
    ws.setFrozenRows(1);
    ws.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    (textCols || []).forEach(c => ws.getRange(1, c, ws.getMaxRows(), 1).setNumberFormat('@'));
  }
  return ws;
}
function recvRows_(ws, width) {
  const last = ws.getLastRow();
  return last >= 2 ? ws.getRange(2, 1, last - 1, width).getValues() : [];
}
function recvStamp_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'); }
function recvDay_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy'); }

// Product Master: barcode keys already known
function recvMasterKeys_() {
  const keys = {};
  const pm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (pm && pm.getLastRow() >= FIRST_DATA_ROW) {
    pm.getRange(FIRST_DATA_ROW, 8, pm.getLastRow() - FIRST_DATA_ROW + 1, 1).getValues().forEach(r => { if (r[0]) keys[barcodeKey(r[0])] = true; });
  }
  return keys;
}
// Adds new products to Product Master (no price) and to the Warehouse tab, so
// the Scanner can sell them and the store level of the Warehouse moves with them.
function recvEnsureProducts_(products) {
  const added = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pm = ss.getSheetByName(SHEET_NAME);
  const known = recvMasterKeys_();
  if (pm) {
    let maxNum = 0;
    if (pm.getLastRow() >= FIRST_DATA_ROW) {
      pm.getRange(FIRST_DATA_ROW, 1, pm.getLastRow() - FIRST_DATA_ROW + 1, 1).getValues()
        .forEach(r => { const n = parseInt(r[0], 10); if (!isNaN(n) && n > maxNum) maxNum = n; });
    }
    products.forEach(p => {
      const k = barcodeKey(p.bc);
      if (!k || known[k]) return;
      const row = pm.getLastRow() + 1;
      pm.getRange(row, 8).setNumberFormat('@');
      pm.getRange(row, 1, 1, 2).setValues([[++maxNum, p.name || p.bc]]);
      pm.getRange(row, 4).setValue('Cosmetics');
      pm.getRange(row, 8).setValue(p.bc);
      known[k] = true;
      added.push(p.bc);
    });
  }
  const loc = getLocationSheetOrNull(RECV_LOC);
  if (loc) {
    const have = {};
    if (loc.getLastRow() >= FIRST_DATA_ROW) {
      loc.getRange(FIRST_DATA_ROW, LOCATION_BARCODE_COL, loc.getLastRow() - FIRST_DATA_ROW + 1, 1).getValues()
        .forEach(r => { if (r[0]) have[barcodeKey(r[0])] = true; });
    }
    const qtyCol = getOrCreateColumnByHeader(loc, HEADER_ROW, QTY_ON_HAND_HEADER);
    products.forEach(p => {
      const k = barcodeKey(p.bc);
      if (!k || have[k]) return;
      const row = loc.getLastRow() + 1;
      loc.getRange(row, LOCATION_NAME_COL).setValue(p.name || p.bc);
      loc.getRange(row, LOCATION_BARCODE_COL).setNumberFormat('@').setValue(p.bc);
      loc.getRange(row, qtyCol).setValue(0);
      have[k] = true;
    });
  }
  return added;
}

function recvParseDate_(s) {
  const d = parseFrDate(String(s || ''));
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
}

// Writes dates for received lines into Expiry Batches. A product+date+arrival
// already recorded from another cooler gets its quantity added, not a new row.
function recvWriteExpiry_(lines, arrivalName, by) {
  const ws = ensureExpiryTab();
  _exEnsureCols(ws);
  const data = recvRows_(ws, EX_ALL_HEADERS.length);
  const noExp = {}, sameBatch = {};
  data.forEach((r, i) => {
    const k = barcodeKey(r[EX_COL_BARCODE - 1]);
    if (r[EX_COL_STATUS - 1] === 'noexpiry') noExp[k] = true;
    if (_exIsActive(r[EX_COL_STATUS - 1]) && r[EX_COL_LOCATION - 1] === RECV_LOC && String(r[EX_COL_ZONE - 1]) === RECV_ZONE &&
        String(r[EX_COL_SHIPMENT - 1]) === arrivalName) sameBatch[k + '|' + _exStr(r[EX_COL_EXPIRY - 1])] = { row: i + 2, qty: Number(r[EX_COL_QTY - 1]) || 0 };
  });
  const stamp = recvStamp_(), day = recvDay_();
  const add = [], pending = {};
  lines.forEach(l => {
    const k = barcodeKey(l.bc);
    if (l.mode === 'none') {
      if (noExp[k]) return;
      noExp[k] = true;
      add.push([stamp, RECV_LOC, RECV_ZONE, l.bc, l.name || l.bc, 0, '', '', 'noexpiry', by, day]);
    } else if (l.mode === 'date' && l.expiry) {
      const bk = k + '|' + l.expiry;
      if (pending[bk] != null) { add[pending[bk]][5] += l.qty; return; }
      const sb = sameBatch[bk];
      if (sb) { sb.qty += l.qty; ws.getRange(sb.row, EX_COL_QTY).setValue(sb.qty); return; }
      pending[bk] = add.length;
      add.push([stamp, RECV_LOC, RECV_ZONE, l.bc, l.name || l.bc, l.qty, l.expiry, arrivalName, 'watch', by, day]);
    }
  });
  if (add.length) {
    const r0 = ws.getLastRow() + 1;
    const rng = ws.getRange(r0, 1, add.length, EX_ALL_HEADERS.length);
    rng.setNumberFormat('@');
    rng.setValues(add);
  }
}

function handleRecvArrival(body) {
  if (!s1981UserOk_(body)) return jsonResponse({ success: false, error: "Utilisateur ou code non reconnu — reconnectez-vous." }, 403);
  const id = String(body.id || '').trim(), name = String(body.name || '').trim();
  if (!id || !name) return jsonResponse({ success: false, error: "Nom de l'arrivage manquant" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ws = recvSheet_(RECV_ARR_SHEET, RECV_ARR_HEADERS);
    const rows = recvRows_(ws, RECV_ARR_HEADERS.length);
    const i = rows.findIndex(r => String(r[0]) === id);
    const status = body.status === 'closed' ? 'closed' : 'open';
    if (i >= 0) {
      ws.getRange(i + 2, 2).setValue(name);
      if (body.status) ws.getRange(i + 2, 5).setValue(status);
    } else ws.appendRow([id, name, recvStamp_(), RECV_LOC, status, String(body.by || ''), String(body.note || '')]);
    return jsonResponse({ success: true });
  } finally { lock.releaseLock(); }
}

function handleRecvUnit(body) {
  if (!s1981UserOk_(body)) return jsonResponse({ success: false, error: "Utilisateur ou code non reconnu — reconnectez-vous." }, 403);
  const opId = String(body.opId || ''), aid = String(body.arrivalId || ''), aname = String(body.arrivalName || '').trim();
  const unit = String(body.unit || '').trim(), kind = String(body.kind || 'Glacière');
  if (!opId || !aid || !aname || !unit) return jsonResponse({ success: false, error: "Arrivage ou numéro de glacière manquant" }, 400);
  const by = String(body.by || '');
  const lines = (Array.isArray(body.lines) ? body.lines : []).map((l, i) => ({
    lineId: String(l.lineId || (opId + '-' + i)), bc: normalizeBarcode(resolveAlias_(l.barcode)), name: String(l.name || '').trim(),
    qty: Math.max(0, Math.round(Number(l.qty) || 0)), mode: ['date', 'none', 'later'].indexOf(l.mode) >= 0 ? l.mode : 'later',
    expiry: recvParseDate_(l.expiry), solde: !!l.solde
  })).filter(l => l.bc && l.qty > 0);
  lines.forEach(l => { if (l.mode === 'date' && !l.expiry) l.mode = 'later'; });
  const cooler = body.cooler && normalizeBarcode(resolveAlias_(body.cooler.barcode)) ?
    { bc: normalizeBarcode(resolveAlias_(body.cooler.barcode)), name: String(body.cooler.name || '').trim() } : null;

  // 1. Already applied? (offline retry)
  const unitWs = recvSheet_(RECV_UNIT_SHEET, RECV_UNIT_HEADERS, [6]);
  if (recvRows_(unitWs, RECV_UNIT_HEADERS.length).some(r => String(r[10]) === opId)) return jsonResponse({ success: true, duplicate: true });

  // 2. Arrival row + unknown products (under the lock)
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let added = [];
  try {
    const aws = recvSheet_(RECV_ARR_SHEET, RECV_ARR_HEADERS);
    if (!recvRows_(aws, RECV_ARR_HEADERS.length).some(r => String(r[0]) === aid)) aws.appendRow([aid, aname, recvStamp_(), RECV_LOC, 'open', by, '']);
    const prods = {};
    lines.forEach(l => { prods[barcodeKey(l.bc)] = prods[barcodeKey(l.bc)] || { bc: l.bc, name: l.name }; });
    if (cooler) prods[barcodeKey(cooler.bc)] = prods[barcodeKey(cooler.bc)] || { bc: cooler.bc, name: cooler.name };
    added = recvEnsureProducts_(Object.keys(prods).map(k => prods[k]));
  } finally { lock.releaseLock(); }

  // 3. Stock (each op has its own opId, so a retry never adds twice)
  const note = aname + ' · ' + kind + ' ' + unit;
  const ops = lines.map((l, i) => ({ opId: opId + '-s' + i, location: RECV_LOC, zone: RECV_ZONE, barcode: l.bc, name: l.name, mode: 'delta', qty: l.qty, reason: 'Ajout / réception', note: note }));
  if (cooler) ops.push({ opId: opId + '-c', location: RECV_LOC, zone: RECV_ZONE, barcode: cooler.bc, name: cooler.name, mode: 'delta', qty: 1, reason: 'Ajout / réception', note: note + ' (glacière vidée, à vendre)' });
  if (ops.length) {
    const res = JSON.parse(handleStock1981Adjust({ by: by, userHash: body.userHash, ops: ops }).getContent());
    if (!res.success) return jsonResponse({ success: false, error: res.error || 'Stock non mis à jour' });
    const bad = (res.results || []).filter(r => !r.ok);
    if (bad.length) return jsonResponse({ success: false, error: bad.length + ' ligne(s) refusée(s) : ' + bad[0].error });
  }

  // 4. Dates, lines, unit summary
  lock.waitLock(25000);
  try {
    if (recvRows_(unitWs, RECV_UNIT_HEADERS.length).some(r => String(r[10]) === opId)) return jsonResponse({ success: true, duplicate: true });
    recvWriteExpiry_(lines, aname, by);
    const stamp = recvStamp_();
    if (lines.length) {
      const lws = recvSheet_(RECV_LINE_SHEET, RECV_LINE_HEADERS, [5, 9]);
      const rows = lines.map(l => [stamp, aid, aname, unit, l.bc, l.name, l.qty, l.mode, l.mode === 'date' ? l.expiry : '', l.solde ? 'oui' : '', by, l.lineId]);
      lws.getRange(lws.getLastRow() + 1, 1, rows.length, RECV_LINE_HEADERS.length).setValues(rows);
    }
    const units = lines.reduce((a, l) => a + l.qty, 0);
    const prodCount = Object.keys(lines.reduce((o, l) => (o[barcodeKey(l.bc)] = 1, o), {})).length;
    unitWs.appendRow([stamp, aid, aname, unit, kind, cooler ? cooler.bc : '', cooler ? cooler.name : '', prodCount, units, by, opId]);
    return jsonResponse({ success: true, products: prodCount, units: units, newProducts: added.length });
  } finally { lock.releaseLock(); }
}

function handleRecvSetDate(body) {
  if (!s1981UserOk_(body)) return jsonResponse({ success: false, error: "Utilisateur ou code non reconnu — reconnectez-vous." }, 403);
  const lineId = String(body.lineId || '');
  const mode = body.mode === 'none' ? 'none' : 'date';
  const expiry = recvParseDate_(body.expiry);
  if (!lineId || (mode === 'date' && !expiry)) return jsonResponse({ success: false, error: "Date invalide" }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const lws = recvSheet_(RECV_LINE_SHEET, RECV_LINE_HEADERS, [5, 9]);
    const rows = recvRows_(lws, RECV_LINE_HEADERS.length);
    const i = rows.findIndex(r => String(r[RL.LID]) === lineId);
    if (i < 0) return jsonResponse({ success: false, error: "Ligne introuvable" });
    const r = rows[i];
    if (r[RL.MODE] !== 'later') return jsonResponse({ success: true, already: true });
    lws.getRange(i + 2, RL.MODE + 1, 1, 2).setValues([[mode, mode === 'date' ? expiry : '']]);
    recvWriteExpiry_([{ bc: normalizeBarcode(r[RL.BC]), name: String(r[RL.NAME]), qty: Number(r[RL.QTY]) || 0, mode: mode, expiry: expiry }],
      String(r[RL.ANAME]), String(body.by || ''));
    return jsonResponse({ success: true });
  } finally { lock.releaseLock(); }
}

function handleGetArrivals() {
  const aws = recvSheet_(RECV_ARR_SHEET, RECV_ARR_HEADERS);
  const uws = recvSheet_(RECV_UNIT_SHEET, RECV_UNIT_HEADERS, [6]);
  const lws = recvSheet_(RECV_LINE_SHEET, RECV_LINE_HEADERS, [5, 9]);
  const prices = _exPriceMap();
  const arrivals = recvRows_(aws, RECV_ARR_HEADERS.length).filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]), created: r[2] instanceof Date ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), 'dd/MM/yyyy') : String(r[2]).slice(0, 10), status: String(r[4] || 'open'), by: String(r[5] || ''), units: [], lines: [] }));
  const byId = {}; arrivals.forEach(a => { byId[a.id] = a; });
  recvRows_(uws, RECV_UNIT_HEADERS.length).forEach(r => {
    const a = byId[String(r[1])]; if (!a) return;
    a.units.push({ unit: String(r[3]), kind: String(r[4]), cooler: String(r[6] || r[5] || ''), products: Number(r[7]) || 0, units: Number(r[8]) || 0, by: String(r[9] || ''), at: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : String(r[0]) });
  });
  recvRows_(lws, RECV_LINE_HEADERS.length).forEach(r => {
    const a = byId[String(r[RL.AID])]; if (!a) return;
    const p = prices[barcodeKey(r[RL.BC])];
    // [lineId, unit, barcode, name, qty, mode, expiry, solde, price]
    a.lines.push([String(r[RL.LID]), String(r[RL.UNIT]), normalizeBarcode(r[RL.BC]), String(r[RL.NAME]), Number(r[RL.QTY]) || 0,
      String(r[RL.MODE]), _exStr(r[RL.EXP]), r[RL.SOLDE] === 'oui' ? 1 : 0, p == null ? null : p]);
  });
  // products recorded once as « ne périme pas » (any place) are never asked again
  const noExpiry = {};
  recvRows_(ensureExpiryTab(), EXPIRY_HEADERS.length).forEach(r => { if (r[EX_COL_STATUS - 1] === 'noexpiry') noExpiry[barcodeKey(r[EX_COL_BARCODE - 1])] = 1; });
  arrivals.sort((a, b) => (a.status === b.status ? 0 : a.status === 'open' ? -1 : 1) || (a.id < b.id ? 1 : -1));
  return jsonResponse({ success: true, arrivals: arrivals, noExpiry: Object.keys(noExpiry), generatedAt: new Date().toISOString() });
}


// ══════════════════════════════════════════════════════════════════════════
//  STOCK1981 PRODUCT PHOTOS (29 Oct 2026) — added to getStock1981 as `images`
//  { barcodeKey: url }. Same photos the Scanner shows at the till: the sales
//  Sheet's « Product Images » tab, rows « Trouvé » / « Validé » with a photo
//  link (read through SALES_SHEET_ID, like the controls), plus the Catalogue
//  photos kept on the site (images/products/…). Read-only.
// ══════════════════════════════════════════════════════════════════════════
function s1981Images_() {
  const out = {};
  try {
    const cat = readCatalogue_();
    Object.keys(cat).forEach(l => cat[l].forEach(it => {
      const k = barcodeKey(it.barcode);
      if (k && it.photo && !out[k]) out[k] = 'images/products/' + encodeURIComponent(it.photo);
    }));
  } catch (e) { Logger.log('catalogue photos skipped: ' + e); }
  try {
    const ss = ctlSalesSS_();
    const sh = ss && ss.getSheetByName('Product Images');
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(r => {
        if (['Trouvé', 'Validé'].indexOf(String(r[2])) < 0 || !r[5]) return;
        const k = barcodeKey(r[0]);
        if (k) out[k] = String(r[5]);
      });
    }
  } catch (e) { Logger.log('sales photos skipped: ' + e); }
  return out;
}


// ══════════════════════════════════════════════════════════════════════════
//  PHOTO SEARCH FOR CONFIRMED STOCK (29 Oct 2026)
//  When stock is confirmed in Stock1981 (shelf count sent, cooler emptied in
//  Réception, product added to a zone), every product that has no row yet in
//  the sales Sheet's « Product Images » tab gets one, « À chercher ». The
//  sales script's 15-min runImageLookup() then searches its photo as usual
//  (free databases, UPCitemdb, Brave suggestions). Our own 0610000… codes are
//  marked « Code interne » (no website knows them). Never blocks the count:
//  any problem is only logged.
//  The first getStock1981 after this update also queues everything already
//  counted with stock (once; Script property S1981_PHOTO_CATCHUP).
// ══════════════════════════════════════════════════════════════════════════
function s1981QueuePhotos_(products) {
  try {
    if (!products || !products.length) return 0;
    const ss = ctlSalesSS_();
    const sh = ss && ss.getSheetByName('Product Images');
    if (!sh) return 0;
    const have = {};
    if (sh.getLastRow() >= 2) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(r => { const k = barcodeKey(r[0]); if (k) have[k] = true; });
    const rows = [];
    products.forEach(p => {
      const bc = normalizeBarcode(p.bc), k = barcodeKey(bc);
      if (!k || have[k] || k.length < 6) return;
      have[k] = true;
      rows.push([bc, String(p.name || ''), /^610000/.test(k) ? 'Code interne' : 'À chercher', '', '', '', '', '', '', 0, '']);
    });
    if (!rows.length) return 0;
    const r0 = sh.getLastRow() + 1;
    sh.getRange(r0, 1, rows.length, 1).setNumberFormat('@');
    sh.getRange(r0, 1, rows.length, rows[0].length).setValues(rows);
    return rows.length;
  } catch (e) { Logger.log('photo queue skipped: ' + e); return 0; }
}
function s1981PhotoCatchUp_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('S1981_PHOTO_CATCHUP') === '1') return;
    const zs = ensureZoneStockTab();
    const prods = [];
    if (zs.getLastRow() >= 2) zs.getRange(2, 1, zs.getLastRow() - 1, ZONE_STOCK_HEADERS.length).getValues().forEach(r => {
      if ((Number(r[ZS_COL_ONHAND - 1]) || 0) > 0 && r[ZS_COL_BARCODE - 1]) prods.push({ bc: r[ZS_COL_BARCODE - 1], name: r[ZS_COL_PRODUCT - 1] });
    });
    const n = s1981QueuePhotos_(prods);
    props.setProperty('S1981_PHOTO_CATCHUP', '1');
    Logger.log('photo catch-up: ' + n + ' product(s) queued');
  } catch (e) { Logger.log('photo catch-up skipped: ' + e); }
}
// Can also be run by hand from the editor to queue everything in stock again.
function queueStockPhotosNow() { PropertiesService.getScriptProperties().deleteProperty('S1981_PHOTO_CATCHUP'); s1981PhotoCatchUp_(); }


// ══════════════════════════════════════════════════════════════════════════
//  EN ROUTE (action "getTransit", GET, key required), added 29 Oct 2026
//  Read-only, for Stock1981 « 🚚 En route ». Reads the sales Sheet's
//  « Transfers » tab (written by the Scanner's 🔄 Transferts): lines sent and
//  not yet received (status 'transit'), and requests not yet prepared
//  ('requested'). The JSON column holds the shipment (grp), carrier, boxes,
//  sentTs / sentBy, ts / by.
// ══════════════════════════════════════════════════════════════════════════
function handleGetTransit() {
  const out = { transit: [], requested: [] };
  try {
    const ss = ctlSalesSS_();
    const sh = ss && ss.getSheetByName('Transfers');
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 14).getValues().forEach(r => {
        const st = String(r[3] || '');
        if (st !== 'transit' && st !== 'requested') return;
        let o = {};
        try { o = r[13] ? JSON.parse(String(r[13])) : {}; } catch (e) { o = {}; }
        const line = { id: String(r[2] || o.id || ''), grp: String(o.grp || o.gid || r[2] || ''), from: String(r[4] || ''), to: String(r[5] || ''),
          bc: normalizeBarcode(r[6]), name: String(r[7] || ''), qty: Number(r[8]) || 0, by: String(r[10] || o.by || ''),
          carrier: String(o.carrier || ''), boxes: Number(o.boxes) || 0, sentTs: String(o.sentTs || ''), sentBy: String(o.sentBy || ''), ts: String(o.ts || '') };
        (st === 'transit' ? out.transit : out.requested).push(line);
      });
    }
  } catch (e) { return jsonResponse({ success: false, error: 'Transferts illisibles : ' + e }); }
  return jsonResponse(Object.assign({ success: true, generatedAt: new Date().toISOString() }, out));
}


// ══════════════════════════════════════════════════════════════════════════
//  FORMATION — staff cosmetics training (formation.html, 30 Oct 2026)
//  Content lives in two tabs you edit yourself:
//    « Formation Contenu »   one row per ingredient (level, effect, notes,
//                            what customers say, our products' barcodes)
//    « Formation Questions » your own extra quiz questions
//  Results are written by the page:
//    « Formation Résultats » one row per game played (score, exam passed)
//    « Formation Réponses »  one row per answer (topic, right / wrong)
//  Both content tabs are created with a first Level 1 the first time the
//  page loads. Set « Actif » to NON to hide a row; add rows with Niveau 2
//  for next week's level. Training staff are « Formation » users in
//  Counter Users: they can only sign in to the training page.
// ══════════════════════════════════════════════════════════════════════════
const FORMATION_PASS = 80;        // % needed on the exam to pass a level
const FORMATION_EXAM_SIZE = 20;   // questions per exam
const FM_CONTENT_SHEET = 'Formation Contenu';
const FM_QUESTIONS_SHEET = 'Formation Questions';
const FM_RESULTS_SHEET = 'Formation Résultats';
const FM_ANSWERS_SHEET = 'Formation Réponses';
const FM_CONTENT_HEADERS = ['Niveau', 'Ordre', 'Ingrédient', 'Pour', 'Effet court', 'Explication', 'À dire au client', 'Attention', 'Le client dit (séparer par |)', 'Codes-barres de nos produits (séparer par ,)', 'Actif'];
const FM_QUESTIONS_HEADERS = ['Niveau', 'Question', 'Bonne réponse', 'Faux 1', 'Faux 2', 'Faux 3', 'Sujet (ingrédient)', 'Explication', 'Actif'];
const FM_RESULTS_HEADERS = ['Date', 'Semaine', 'Employé', 'Niveau', 'Jeu', 'Score', 'Total', '%', 'Examen réussi', 'Durée (s)', 'Session'];
const FM_ANSWERS_HEADERS = ['Date', 'Semaine', 'Employé', 'Niveau', 'Jeu', 'Session', 'Sujet', 'Question', 'Réponse donnée', 'Bonne réponse', 'Juste'];

const FM_SEED_CONTENT = [
  [1, 1, 'Vitamine C', 'Peau', 'Éclat du teint et taches',
    "Antioxydant. Rend le teint plus lumineux et aide à atténuer les taches brunes, avec le temps. On la trouve dans des crèmes, des sérums, des savons, et aussi en comprimés à avaler.",
    "Pour un teint plus lumineux et des taches moins visibles, un produit à la vitamine C, tous les jours.",
    "Résultats en quelques semaines, pas en un jour. Protection solaire le jour. Peut picoter un peu les peaux sensibles.",
    "J'ai des taches brunes | Mon teint est terne, sans éclat",
    '811170034395, 75609201059, 75609215339, 19100109995, 810400035447, 8859587314114, 74312715105, 76314302024'],
  [1, 2, 'Vitamine E', 'Peau et cheveux', 'Nourrit et protège',
    "Antioxydant qui nourrit, adoucit et protège la peau du dessèchement. Souvent associée au beurre de cacao, au karité ou à la vitamine C.",
    "La vitamine E nourrit et protège la peau : très bien pour une peau sèche ou abîmée.",
    "Ce n'est pas un produit contre les taches : elle nourrit et protège.",
    "Je cherche une crème qui protège et nourrit ma peau",
    '610000084353, 10181140082, 810107180075, 19100109995, 8859587314114, 610000010260'],
  [1, 3, 'Rétinol', 'Peau', 'Rides et renouvellement de la peau',
    "Dérivé de la vitamine A. Accélère le renouvellement des cellules : rides et ridules moins marquées, peau plus lisse. Aide aussi contre les boutons.",
    "Le rétinol s'utilise le soir. On commence 2 ou 3 soirs par semaine.",
    "Seulement le soir, et protection solaire le jour. Peut irriter au début (rougeurs, peau qui pèle). Déconseillé pendant la grossesse et l'allaitement.",
    "J'ai des rides, je veux une crème anti-âge",
    ''],
  [1, 4, 'Niacinamide (vitamine B3)', 'Peau', 'Pores et peau grasse',
    "Resserre l'aspect des pores, régule le sébum (peau qui brille), apaise les rougeurs et aide à unifier le teint. Douce : convient à presque toutes les peaux.",
    "Pour la peau grasse et les pores visibles, la niacinamide est douce, matin et soir.",
    "Très bien tolérée. Peut s'utiliser avec la vitamine C.",
    "J'ai la peau grasse, je brille | Mes pores sont très visibles",
    '3600551183992'],
  [1, 5, 'Acide hyaluronique', 'Peau', 'Hydratation',
    "Attire et retient l'eau dans la peau, comme une éponge. La peau est repulpée et plus souple. Convient à tous les types de peau, même grasse.",
    "Pour une peau qui tire et manque d'eau, l'acide hyaluronique hydrate sans graisser.",
    "Malgré son nom, il ne pique pas et n'exfolie pas. Mettre une crème par-dessus pour garder l'eau.",
    "Ma peau tire, elle manque d'eau",
    '3606000537668, 8851932427722'],
  [1, 6, 'Collagène', 'Peau et cheveux', 'Fermeté et souplesse',
    "Protéine qui donne sa fermeté à la peau. Dans une crème, un savon ou un shampoing, il hydrate et adoucit la surface. En poudre ou en gélules, c'est un complément alimentaire qu'on avale.",
    "En crème, le collagène hydrate et adoucit. Il existe aussi en poudre à boire, comme complément.",
    "Une crème au collagène ne remplace pas le collagène de la peau : ne pas promettre de miracle.",
    "Ma peau est moins ferme, elle se relâche",
    '8851437881982, 8859587314114, 19100110083, 610000020092, 8859587324113, 816559019031, 610000010260, 16185129481, 857273008666, 817598001032'],
  [1, 7, 'Beurre de karité', 'Peau et cheveux', 'Nourrit peaux et cheveux secs',
    "Beurre de la noix de l'arbre à karité, un produit de chez nous. Très nourrissant : protège la peau sèche et adoucit les cheveux secs ou crépus.",
    "Le karité nourrit en profondeur la peau et les cheveux très secs.",
    "Très riche : il peut être lourd sur un visage gras (risque de boutons).",
    "Ma peau et mes cheveux sont très secs",
    '810107180075'],
  [1, 8, 'Beurre de cacao', 'Peau', 'Adoucit et assouplit la peau',
    "Beurre de la fève de cacao. Nourrit et rend la peau souple et douce. Très utilisé sur le corps, souvent choisi pendant la grossesse pour les vergetures.",
    "Le beurre de cacao garde la peau souple : beaucoup de clientes l'utilisent sur le ventre pendant la grossesse.",
    "Il aide la peau à rester souple mais n'efface pas les vergetures déjà là.",
    "J'ai des vergetures",
    '610000084353, 10181140082'],
  [1, 9, 'AHA (acides de fruits)', 'Peau', 'Exfolie : peau lisse et éclatante',
    "Acide glycolique, lactique, citrique… Ils enlèvent les cellules mortes à la surface : peau plus lisse, teint plus éclatant, petites taches atténuées.",
    "Les AHA exfolient en douceur, pour une peau plus lisse et plus nette.",
    "Rendent la peau plus sensible au soleil : protection solaire le jour. Pas tous les jours au début.",
    "Ma peau est rugueuse, pas lisse",
    ''],
  [1, 10, 'BHA (acide salicylique)', 'Peau', 'Boutons et points noirs',
    "Exfoliant qui entre dans les pores et les débouche. Idéal contre les boutons, les points noirs et la peau grasse.",
    "Pour les boutons et les points noirs, un produit à l'acide salicylique (BHA).",
    "Peut dessécher : ajouter une crème hydratante. Protection solaire le jour. Pas avec un autre exfoliant le même soir.",
    "J'ai des boutons et des points noirs",
    '70501025901'],
  [1, 11, 'Céramides', 'Peau', 'Répare la barrière de la peau',
    "Graisses naturellement présentes dans la peau. Elles jouent le rôle du ciment entre les briques : moins de sécheresse, moins d'irritations.",
    "Pour une peau sèche qui gratte ou s'irrite, une crème aux céramides répare et protège.",
    "Très douces, conviennent aussi aux peaux sensibles et aux enfants.",
    "Ma peau gratte et s'irrite facilement",
    '3606000537668, 52800483286'],
  [1, 12, 'Acide kojique / Arbutine', 'Peau', 'Unifie le teint, atténue les taches',
    "Actifs qui freinent la fabrication des taches brunes (mélanine). On les trouve dans beaucoup de nos savons et sérums « éclat » de Thaïlande.",
    "Pour les taches : utiliser régulièrement et se protéger du soleil, sinon les taches reviennent.",
    "Peut irriter : essayer d'abord sur une petite zone. Protection solaire indispensable. Pas sur une peau blessée.",
    "J'ai des taches brunes",
    '610000020078, 8851437881982'],
  [1, 13, 'Glutathion (« gluta »)', 'Peau', 'Antioxydant pour un teint lumineux',
    "Antioxydant présent naturellement dans le corps. Dans les savons et laits « gluta », il est utilisé pour un teint plus lumineux et plus uniforme.",
    "Le gluta aide à avoir un teint plus lumineux, avec une utilisation régulière.",
    "En crème ou en savon, les effets restent modestes : ne pas promettre un changement de couleur de peau.",
    '',
    '610000020023, 8851932427722']
];
const FM_SEED_QUESTIONS = [
  [1, "Une cliente enceinte demande une crème anti-rides. Quel actif faut-il éviter ?", 'Rétinol', 'Acide hyaluronique', 'Beurre de cacao', 'Vitamine E', 'Rétinol', "Le rétinol est déconseillé pendant la grossesse et l'allaitement."],
  [1, "Quand faut-il utiliser le rétinol ?", 'Le soir', 'Le matin avant de sortir', 'Juste avant d’aller au soleil', 'N’importe quand, sans précaution', 'Rétinol', "Le soir seulement, avec une protection solaire le jour."],
  [1, "Quels actifs rendent la peau plus sensible au soleil ?", 'AHA, BHA et rétinol', 'Karité et beurre de cacao', 'Acide hyaluronique et céramides', 'Vitamine E seule', 'AHA (acides de fruits)', "Les exfoliants et le rétinol : toujours conseiller une protection solaire le jour."],
  [1, "D'où vient le beurre de karité ?", "De la noix d'un arbre de chez nous", "De la fève de cacao", "D'un fruit de Thaïlande", "Uniquement d'un laboratoire", 'Beurre de karité', "Le karité est un produit d'Afrique de l'Ouest, du Burkina notamment."],
  [1, "Malgré son nom, l'acide hyaluronique…", "Hydrate et ne pique pas", "Exfolie très fort", "Débouche les pores", "Fait des taches", 'Acide hyaluronique', "Ce n'est pas un acide qui exfolie : il retient l'eau dans la peau."],
  [1, "Une crème au collagène…", "Hydrate et adoucit la peau", "Remplace le collagène perdu de la peau", "Efface les rides en un jour", "Fait pousser les cheveux", 'Collagène', "Elle hydrate en surface. Ne pas promettre de miracle."],
  [1, "Avec un produit contre les taches, que faut-il toujours conseiller ?", "Une protection solaire le jour", "De l'utiliser une seule fois", "De le rincer à l'eau très chaude", "Rien de plus", 'Acide kojique / Arbutine', "Sans protection solaire, les taches reviennent."],
  [1, "Le BHA, c'est quel actif ?", "L'acide salicylique", "La vitamine C", "Le collagène", "Le beurre de cacao", 'BHA (acide salicylique)', "BHA = acide salicylique, contre les boutons et les points noirs."]
];

function fmSheet_(name, headers, seed, textCols) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ws = ss.getSheetByName(name);
  if (ws) return ws;
  ws = ss.insertSheet(name);
  ws.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3e8ee');
  ws.setFrozenRows(1);
  (textCols || []).forEach(c => ws.getRange(1, c, ws.getMaxRows(), 1).setNumberFormat('@'));
  if (seed && seed.length) {
    const rows = seed.map(r => { const o = r.slice(); while (o.length < headers.length - 1) o.push(''); o.push('OUI'); return o; });
    ws.getRange(2, 1, rows.length, headers.length).setValues(rows).setVerticalAlignment('top').setWrap(true);
    if (name === FM_CONTENT_SHEET) [60, 50, 160, 110, 180, 360, 280, 300, 240, 260, 60].forEach((w, i) => ws.setColumnWidth(i + 1, w));
    if (name === FM_QUESTIONS_SHEET) [60, 340, 200, 180, 180, 180, 160, 300, 60].forEach((w, i) => ws.setColumnWidth(i + 1, w));
  }
  return ws;
}
function fmRows_(ws, n) {
  const last = ws.getLastRow();
  return last < 2 ? [] : ws.getRange(2, 1, last - 1, n).getValues();
}
function fmActive_(v) { return String(v).trim().toUpperCase() !== 'NON'; }
function fmWeek_(d) {
  // ISO week, e.g. 2026-S44 (weeks start on Monday)
  const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const day = t.getUTCDay() || 7;
  t.setUTCDate(t.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
  const w = Math.ceil(((t - y0) / 86400000 + 1) / 7);
  return t.getUTCFullYear() + '-S' + (w < 10 ? '0' : '') + w;
}
function fmUser_(body) {
  const u = readCounterUsers_(true).find(x => x.active && x.name.toLowerCase() === String(body.by || '').toLowerCase());
  return (u && u.hash && u.hash === String(body.userHash || '')) ? u : null;
}
function fmProgress_() {
  const out = {};
  const ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FM_RESULTS_SHEET);
  if (!ws) return out;
  fmRows_(ws, FM_RESULTS_HEADERS.length).forEach(r => {
    if (r[4] !== 'Examen' || !r[2]) return;
    const emp = String(r[2]), lvl = String(r[3]);
    const at = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    const p = (out[emp] = out[emp] || {})[lvl] = (out[emp] || {})[lvl] || { passedAt: '', lastExamAt: '', lastPassed: false, best: 0, exams: 0 };
    p.exams++;
    p.best = Math.max(p.best, Number(r[7]) || 0);
    if (!p.lastExamAt || at > p.lastExamAt) { p.lastExamAt = at; p.lastPassed = r[8] === 'OUI'; }
    if (r[8] === 'OUI' && (!p.passedAt || at < p.passedAt)) p.passedAt = at;
  });
  return out;
}

function handleGetFormation() {
  try {
    const cws = fmSheet_(FM_CONTENT_SHEET, FM_CONTENT_HEADERS, FM_SEED_CONTENT, [10]);
    const qws = fmSheet_(FM_QUESTIONS_SHEET, FM_QUESTIONS_HEADERS, FM_SEED_QUESTIONS);
    const content = fmRows_(cws, FM_CONTENT_HEADERS.length)
      .filter(r => String(r[2]).trim() && fmActive_(r[10]))
      .map(r => ({
        level: Number(r[0]) || 1, order: Number(r[1]) || 0, name: String(r[2]).trim(), for: String(r[3]).trim(),
        effect: String(r[4]).trim(), text: String(r[5]).trim(), say: String(r[6]).trim(), warn: String(r[7]).trim(),
        says: String(r[8]).split('|').map(x => x.trim()).filter(String),
        products: String(r[9]).split(/[,;\s]+/).map(x => barcodeKey(x)).filter(String)
      }))
      .sort((a, b) => a.level - b.level || a.order - b.order);
    const questions = fmRows_(qws, FM_QUESTIONS_HEADERS.length)
      .filter(r => String(r[1]).trim() && String(r[2]).trim() && fmActive_(r[8]))
      .map(r => ({ level: Number(r[0]) || 1, q: String(r[1]).trim(), a: String(r[2]).trim(),
        wrong: [r[3], r[4], r[5]].map(x => String(x).trim()).filter(String), topic: String(r[6]).trim(), why: String(r[7]).trim() }));
    // names + photos for the barcodes used in the content
    const keys = {};
    content.forEach(c => c.products.forEach(k => keys[k] = true));
    const products = {};
    const pm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (pm && pm.getLastRow() >= FIRST_DATA_ROW) {
      pm.getRange(FIRST_DATA_ROW, 1, pm.getLastRow() - FIRST_DATA_ROW + 1, 8).getValues().forEach(r => {
        const k = barcodeKey(r[7]);
        if (k && keys[k] && r[1] && !products[k]) products[k] = { name: String(r[1]).trim(), img: '' };
      });
    }
    let images = {};
    try { images = s1981Images_(); } catch (e) { Logger.log('formation photos skipped: ' + e); }
    Object.keys(keys).forEach(k => {
      if (!products[k]) products[k] = { name: '', img: '' };
      if (images[k]) products[k].img = images[k];
    });
    return jsonResponse({ success: true, content: content, questions: questions, products: products,
      progress: fmProgress_(), passMark: FORMATION_PASS, examSize: FORMATION_EXAM_SIZE, generatedAt: new Date().toISOString() });
  } catch (err) {
    return jsonResponse({ success: false, error: 'Formation : ' + err.message });
  }
}

// body: { by, userHash, session: { id, level, game, score, total, seconds, at,
//         answers: [ { topic, q, given, correct, ok } ] } }
function handleFormationResult(body) {
  const u = fmUser_(body);
  if (!u) return jsonResponse({ success: false, error: 'Utilisateur ou code incorrect' }, 403);
  const s = body.session || {};
  if (!s.id || !s.game) return jsonResponse({ success: false, error: 'Session incomplète' }, 400);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rws = fmSheet_(FM_RESULTS_SHEET, FM_RESULTS_HEADERS);
    const aws = fmSheet_(FM_ANSWERS_SHEET, FM_ANSWERS_HEADERS);
    const last = rws.getLastRow();
    if (last >= 2) {
      const ids = rws.getRange(2, 11, last - 1, 1).getValues().map(r => String(r[0]));
      if (ids.indexOf(String(s.id)) >= 0) return jsonResponse({ success: true, duplicate: true, progress: fmProgress_() });
    }
    const at = s.at ? new Date(s.at) : new Date();
    const when = isNaN(at.getTime()) ? new Date() : at;
    const week = fmWeek_(when);
    const total = Math.max(0, Number(s.total) || 0), score = Math.max(0, Number(s.score) || 0);
    const pct = total ? Math.round(score * 100 / total) : 0;
    const passed = s.game === 'Examen' ? (pct >= FORMATION_PASS ? 'OUI' : 'NON') : '';
    rws.appendRow([when, week, u.name, Number(s.level) || 1, String(s.game), score, total, pct, passed, Math.round(Number(s.seconds) || 0), String(s.id)]);
    const ans = (s.answers || []).slice(0, 200).map(a => [when, week, u.name, Number(s.level) || 1, String(s.game), String(s.id),
      String(a.topic || ''), String(a.q || '').slice(0, 300), String(a.given || '').slice(0, 200), String(a.correct || '').slice(0, 200), a.ok ? 'OUI' : 'NON']);
    if (ans.length) aws.getRange(aws.getLastRow() + 1, 1, ans.length, FM_ANSWERS_HEADERS.length).setValues(ans);
    return jsonResponse({ success: true, passed: passed, pct: pct, progress: fmProgress_() });
  } finally {
    lock.releaseLock();
  }
}

// Admin only. Everything the « Suivi » screen needs, already grouped.
function handleFormationSuivi(body) {
  const u = fmUser_(body);
  if (!u || u.role !== 'Admin') return jsonResponse({ success: false, error: 'Réservé à l’administrateur' }, 403);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const results = [];
  const rws = ss.getSheetByName(FM_RESULTS_SHEET);
  if (rws) fmRows_(rws, FM_RESULTS_HEADERS.length).forEach(r => {
    if (!r[2]) return;
    results.push({ at: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), week: String(r[1]), emp: String(r[2]),
      level: Number(r[3]) || 1, game: String(r[4]), score: Number(r[5]) || 0, total: Number(r[6]) || 0, pct: Number(r[7]) || 0,
      passed: r[8] === 'OUI', seconds: Number(r[9]) || 0 });
  });
  const topics = {}, weeks = {};
  const aws = ss.getSheetByName(FM_ANSWERS_SHEET);
  if (aws) fmRows_(aws, FM_ANSWERS_HEADERS.length).forEach(r => {
    const emp = String(r[2]); if (!emp) return;
    const ok = r[10] === 'OUI', t = String(r[6]) || '—', w = String(r[1]);
    const tp = (topics[emp] = topics[emp] || {})[t] = (topics[emp] || {})[t] || [0, 0];
    tp[0]++; if (!ok) tp[1]++;
    const wk = (weeks[emp] = weeks[emp] || {})[w] = (weeks[emp] || {})[w] || [0, 0];
    wk[0]++; if (ok) wk[1]++;
  });
  const users = readCounterUsers_(true).filter(x => x.active).map(x => ({ name: x.name, role: x.role }));
  return jsonResponse({ success: true, users: users, results: results, topics: topics, weeks: weeks,
    progress: fmProgress_(), passMark: FORMATION_PASS, generatedAt: new Date().toISOString() });
}
