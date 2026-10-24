/**
 * BOUTIQUE POS — SALES SYNC + NIGHTLY REPORT (Apps Script Web App)
 * ==================================================================
 * This is the server-side counterpart to the "Sales Sync" section of the
 * scanner's Admin panel. It receives real-time pushes of every scan, sale
 * and shift, and — because it runs on Google's servers, not in the
 * browser — it can also send the 9:15 PM end-of-day report automatically,
 * even if the scanner device is turned off.
 *
 * SETUP (do once):
 *   1. Create a new Google Sheet (or use one you've set aside for sales data).
 *   2. Extensions > Apps Script. Delete the starter code, paste this whole file.
 *   3. Edit the CONFIG block right below — set WRITE_KEY to any secret string
 *      you make up, and confirm REPORT_EMAIL. Also confirm INVENTORY_SYNC_URL
 *      and INVENTORY_SYNC_KEY match the Store Inventory Sheet's script (see
 *      the INVENTORY SYNC section further down for what these do).
 *   4. Deploy > New deployment > type "Web app".
 *        - Execute as: Me
 *        - Who has access: Anyone
 *      Click Deploy, authorize when prompted, and copy the resulting URL
 *      (it ends in /exec).
 *   5. In the scanner app: Admin > 📤 Sales Sync (Google Sheets) > paste that
 *      URL into "Web App URL", and paste the SAME string you used for
 *      WRITE_KEY below into "Write Key" > Save.
 *   6. Back in this Apps Script editor: pick createNightlyTrigger from the
 *      function dropdown at the top and click Run (once). Authorize again
 *      if asked. This installs the daily 9:15 PM trigger — you only do this
 *      one time, not every day.
 *   7. Also run setupDashboard() once (same way — pick it from the dropdown,
 *      click Run). This creates a "Dashboard" tab with live sales figures
 *      you can check any time from your phone's Google Sheets app — no
 *      notifications, just open it when you want a pulse check. Safe to
 *      re-run any time later (e.g. after adding new sheets) — it just
 *      resets the tab's formulas, it never touches your actual sales data.
 *   8. Run createInventorySyncTrigger() once (same way). This installs a
 *      trigger that checks every 15 minutes for newly-sold or returned
 *      items and applies them to the Store Inventory Sheet's live stock
 *      automatically — see the INVENTORY SYNC section below for details.
 *      The first time it runs you may be asked to authorize an additional
 *      permission ("Connect to an external service") — that's expected,
 *      it's what lets this script call the Store Inventory script.
 *   9. Also run createDashboardRefreshTrigger() once (same way). The
 *      Dashboard normally updates the instant a sale/return/hold/Mensualité
 *      event happens — but on a slow morning (nothing closed yet today) it
 *      would otherwise keep showing yesterday's numbers. This adds a
 *      15-minute timer so it never sits stale for long even with no events.
 *  10. If you're pasting this over an older version of the script that was
 *      already writing real data (not a brand-new spreadsheet): also run
 *      fixDateTimeFormatting() once. Google Sheets can silently auto-convert
 *      Date/Time/Phone text into real Date/Number values on write, which
 *      breaks the "is this today" checks Dashboard and the nightly report
 *      rely on. This repairs anything already affected and locks those
 *      columns so it can't happen again. New spreadsheets never need this —
 *      every sheet created from here on is protected automatically.
 *  11. To use Boutique_Admin_Remote.html (view + manage both stores from any
 *      computer, without visiting either store): set REMOTE_ADMIN_PASSPHRASE
 *      below to a real passphrase (12+ characters — NOT the 4-digit staff
 *      PIN), redeploy (Deploy > Manage deployments > pencil > New version),
 *      then open Boutique_Admin_Remote.html and enter the same Web App URL
 *      + Write Key as the scanner, plus this passphrase. Everything it shows
 *      or changes (branding, employees, admin PIN, hold duration) is pulled
 *      down by the scanner apps automatically within ~15 minutes.
 *
 * That's the whole setup. From then on: scans/sales/shifts push here live,
 * every 15 minutes sold/returned items get applied to live stock on the
 * Store Inventory Sheet, every night at 9:15 PM this script reads today's
 * rows and emails the report to REPORT_EMAIL, and the Dashboard tab always
 * shows current numbers whenever you open it — independent of whether the
 * scanner is open.
 *
 * SALES_YYYY-MM-DD TABS — feed for the inventory-reduction sync below:
 * every item sold or returned lands in a tab named for its own day (e.g.
 * "Sales_2026-08-10"), one row per item, always with Barcode + Product, and
 * a signed Qty (negative for returns). The "Processed" column starts blank
 * and is filled in by syncInventoryFromSales() once it has applied that row
 * to Store Inventory, so a row is never applied twice. See _dayTab() below.
 */
const CONFIG = {
  WRITE_KEY: 'PUT-YOUR-WRITE-KEY-HERE',   // must exactly match the "Write Key" field in the scanner Admin panel
  // Separate, stronger secret for the Remote Admin page only (Boutique_Admin_Remote.html).
  // Change this to a real passphrase (12+ characters, not the 4-digit staff PIN) before
  // using the remote admin page — it guards both reading store data remotely and every
  // remote "manage" action (branding, employees, settings). Staff PINs are unaffected.
  REMOTE_ADMIN_PASSPHRASE: 'nassourou1235',
  REPORT_EMAIL: '1981estvideo@gmail.com',
  REPORT_HOUR: 21,                       // 24h clock — 21 = 9 PM
  REPORT_MINUTE: 15,
  TIMEZONE: 'Africa/Dakar',              // change if the store is in a different timezone
  // ---- Inventory sync (see INVENTORY SYNC section near the bottom) ----
  // This is the Store Inventory Sheet's OWN script URL/key (the one behind
  // the Scanner/Generator/Counter/Stock-Ledger apps) — NOT this script's
  // own WRITE_KEY above. Copy them from that script's CONFIG.
  INVENTORY_SYNC_URL: 'PUT-THE-INVENTORY-WEB-APP-URL-HERE',
  INVENTORY_SYNC_KEY: 'PUT-YOUR-INVENTORY-KEY-HERE',
  INVENTORY_SYNC_LOOKBACK_DAYS: 2,       // how many day-tabs back to re-check for unprocessed rows
};

const SHEETS = {
  SCANS: 'ScanLog', // legacy — no longer written to; per-item rows now go to the Sales_yyyy-MM-dd day-tabs (see _dayTab below). Old rows left in place, untouched.
  SALES: 'SalesLog',
  SHIFTS: 'ShiftLog',
  CUSTOMERS: 'Customers',
  RETURNS: 'Returns',
  INST_PLANS: 'InstallmentPlans',
  INST_PAYMENTS: 'InstallmentPayments',
  HOLDS: 'Holds',
  DISCOUNTS: 'Discounts',
  RECONCILIATION: 'ShiftReconciliation',
  EXPENSES: 'Expenses',
  STORE_SETTINGS: 'StoreSettings',
  GLOBAL_SETTINGS: 'GlobalSettings',
  EMPLOYEES: 'Employees',
};

const PAY_LABELS = {
  cash: '💵 Cash', orange: '🟠 Orange Money', wave: '🌊 Wave',
  bank: '🏦 Virement', monthly: '📅 Mensualité',
};

// ══════════════════════════════════════════════════════════
//  ENTRY POINTS
// ══════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.key !== CONFIG.WRITE_KEY) {
      return _json({ error: 'Unauthorized — write key does not match.' });
    }
    switch (data.action) {
      case 'logScan':              return _logScan(data);
      case 'logActivity':          return _logActivity(data);
      case 'savePricePin':         return _savePricePin(data);
      case 'saveAdminPin':         return _saveAdminPin(data);
      case 'voidScan':              return _voidScan(data);
      case 'finalizeOrphanScans':   return _finalizeOrphanScans(data);
      case 'editSalePayMethod':     return _editSalePayMethod(data);
      case 'editSaleTotal':         return _editSaleTotal(data);
      case 'logSale':               return _logSale(data);
      case 'logPaperSale':          return _logPaperSale(data);
      case 'logShift':              return _logShift(data);
      case 'setProductImage':       return _setProductImage(data);
      case 'processScannedCodes':   return _json(processScannedCodes());
      case 'addNewCode':            return _addNewCode(data);
      case 'skipNewCode':           return _skipNewCode(data);
      case 'saveCustomer':          return _saveCustomer(data);
      case 'saveDelivery':          return _saveDelivery(data);
      case 'saveTransfer':          return _saveTransfer(data);
      case 'retryTransfers':        return _retryTransfers();
      case 'logCashCheck':          return _logCashCheck(data);
      case 'logReturn':             return _logReturn(data);
      case 'createInstallmentPlan': return _createInstallmentPlan(data);
      case 'updatePlanMeta':        return _updatePlanMeta(data);
      case 'logInstallmentPayment': return _logInstallmentPayment(data);
      case 'createHold':            return _createOrUpdateHold(data);
      case 'logDiscount':           return _logDiscount(data);
      case 'logExpense':            return _logExpense(data);
      case 'uploadReceiptImage':    return _saveReceiptImage(data);
      // ---- Remote Admin (Boutique_Admin_Remote.html) writes — each of
      // these additionally requires the REMOTE_ADMIN_PASSPHRASE below,
      // on top of the WRITE_KEY already checked above.
      case 'saveStoreSettings':
        if (!_checkRemotePass(data)) return _json({ error: 'Unauthorized — wrong remote admin passphrase.' });
        return _saveStoreSettings(data);
      case 'saveGlobalSettings':
        if (!_checkRemotePass(data)) return _json({ error: 'Unauthorized — wrong remote admin passphrase.' });
        return _saveGlobalSettings(data);
      case 'saveEmployees':
        if (!_checkRemotePass(data)) return _json({ error: 'Unauthorized — wrong remote admin passphrase.' });
        return _saveEmployees(data);
      default:                       return _json({ error: 'Unknown action: ' + data.action });
    }
  } catch (err) {
    return _json({ error: String(err) });
  }
}

// Shared gate for every Remote Admin write action. Checked in addition to
// (not instead of) the normal WRITE_KEY check every doPost request already
// goes through above.
function _checkRemotePass(d) {
  return !!CONFIG.REMOTE_ADMIN_PASSPHRASE &&
    CONFIG.REMOTE_ADMIN_PASSPHRASE !== 'CHANGE-ME-TO-A-LONG-REMOTE-PASSPHRASE' &&
    d.remotePass === CONFIG.REMOTE_ADMIN_PASSPHRASE;
}

function doGet(e) {
  const action = e.parameter.action;
  // Hosted receipt page — <Web App URL>?r=<TxnID>&s=<Store>. This is what
  // the "Reçu détaillé" link in the WhatsApp receipt message opens. It's a
  // real HTML page (not JSON), so it's handled before the action dispatch
  // below and returns straight away. No key required, same trust level as
  // getCustomers/getInstallmentPlans just below: a receipt/transaction ID
  // isn't meaningfully more sensitive than the number printed on a paper
  // receipt, and every field shown here already appears on the printed
  // receipt or the WhatsApp text it's attached to.
  if (e.parameter.r) return _renderReceiptPage(e.parameter.r, e.parameter.s || '');
  // Cosmetic link-shortening for the "Reçu détaillé" line in the WhatsApp
  // message — till-only (WRITE_KEY gated, same as the settings reads just
  // below), called in the background right when a sale closes. See
  // prefetchShortReceiptLink() in Boutique_Scanner.html.
  if (action === 'shortenReceiptLink') {
    if (e.parameter.key !== CONFIG.WRITE_KEY) return _json({ error: 'Unauthorized — wrong write key.' });
    const longUrl = ScriptApp.getService().getUrl() +
      '?r=' + encodeURIComponent(e.parameter.r || '') + '&s=' + encodeURIComponent(e.parameter.s || '');
    return _json({ shortUrl: _shortenUrl(longUrl) || longUrl });
  }
  if (action === 'getCustomers') return _getCustomers();
  if (action === 'getActivityLog') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getActivityLog(e.parameter.limit) : _json({ error: 'Unauthorized' });
  if (action === 'getRestoreData') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getRestoreData(e.parameter.store, e.parameter.days) : _json({ error: 'Unauthorized' });
  if (action === 'getTransfers') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getTransfers() : _json({ error: 'Unauthorized' });
  if (action === 'getTopSellers') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getTopSellers(Number(e.parameter.days) || 30) : _json({ error: 'Unauthorized' });
  if (action === 'getSalesCalendar') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getSalesCalendar(e.parameter.month || '') : _json({ error: 'Unauthorized' });
  if (action === 'getDeliveries') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getDeliveries() : _json({ error: 'Unauthorized' });
  if (action === 'getProductImages') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getProductImages() : _json({ error: 'Unauthorized' });
  if (action === 'lookupBarcodeName') return (e.parameter.key === CONFIG.WRITE_KEY) ? _lookupBarcodeName(e.parameter.barcode || '') : _json({ error: 'Unauthorized' });
  if (action === 'getNewCodes') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getNewCodes() : _json({ error: 'Unauthorized' });
  if (action === 'getProductImageList') return (e.parameter.key === CONFIG.WRITE_KEY) ? _getProductImageList() : _json({ error: 'Unauthorized' });
  if (action === 'getInstallmentPlans') return _getInstallmentPlans(e.parameter.phone || '');
  // ---- Remote Admin (Boutique_Admin_Remote.html) reads, split into two
  // tiers below. Settings reads (branding/hold-duration/employees) only need the same
  // WRITE_KEY every store device already holds — they're what let the
  // scanner apps pull down changes made remotely, so gating them behind
  // the separate remote passphrase would mean putting that stronger
  // secret on every store computer too, which defeats the point of it.
  if (action === 'getPlansDashboard') {
    if (e.parameter.key !== CONFIG.WRITE_KEY) return _json({ error: 'Unauthorized — wrong write key.' });
    return _getPlansDashboard();
  }
  const settingsActions = ['getStoreSettings', 'getGlobalSettings', 'getEmployees', 'getSyncedKeys'];
  if (settingsActions.indexOf(action) !== -1) {
    if (e.parameter.key !== CONFIG.WRITE_KEY) return _json({ error: 'Unauthorized — wrong write key.' });
    switch (action) {
      case 'getStoreSettings':  return _getStoreSettings();
      case 'getGlobalSettings': return _getGlobalSettings();
      case 'getEmployees':      return _getEmployees();
      // Every TxnID|Store and ReturnID|Store key already in the Sheet — lets
      // a till device work out, entirely client-side, which of its own local
      // sales/returns (going back as far as its local history goes, not
      // just one date) never made it to the Sheet, so only the missing ones
      // get replayed. See runFullResync() in Boutique_Scanner.html.
      case 'getSyncedKeys':     return _getSyncedKeys();
    }
  }
  // Everything else remote-admin-only (financial/monitoring views) requires
  // the stronger REMOTE_ADMIN_PASSPHRASE — only Boutique_Admin_Remote.html
  // needs to know this one, never the in-store scanner devices.
  const remoteOnlyActions = ['getReconciliation', 'getExpenses', 'getDailySummary', 'getHolds',
    'getDashboardData', 'getAllInstallmentPlans'];
  if (remoteOnlyActions.indexOf(action) !== -1) {
    if (!_checkRemotePass({ remotePass: e.parameter.pass })) {
      return _json({ error: 'Unauthorized — wrong remote admin passphrase.' });
    }
    switch (action) {
      case 'getReconciliation':      return _getReconciliation();
      case 'getExpenses':            return _getExpenses();
      // month: optional "yyyy-MM" (native <input type="month"> value) —
      // defaults to the current calendar month when absent, same as before
      // this became selectable in the Remote Admin app.
      case 'getDailySummary':        return _getDailySummary(e.parameter.month || '');
      case 'getHolds':               return _getHoldsRemote();
      case 'getDashboardData':       return _getDashboardData();
      case 'getAllInstallmentPlans': return _getAllInstallmentPlans();
    }
  }
  return _json({ error: 'Unknown GET action.' });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ══════════════════════════════════════════════════════════
//  HOSTED RECEIPT PAGE
// ══════════════════════════════════════════════════════════
// Replaces the old "render a PNG on the till, copy it to the clipboard,
// paste it into WhatsApp" flow. The scanner now sends the customer a short
// text message with a link to this page instead — nothing to paste, and no
// risk of a glyph rendering as a black box on the customer's phone, since
// this is rendered by their own browser rather than plain WhatsApp text.
//
// Two ways this page is built, tried in order:
//   1. The exact receipt PNG the till already generates for print/preview,
//      uploaded via 'uploadReceiptImage' (see _saveReceiptImage below) —
//      pixel-identical to what used to be pasted into WhatsApp by hand,
//      including delivery details/discount tags/barcode that only exist
//      client-side. This is the common case: the till uploads it in the
//      background the moment "Send" is clicked.
//   2. If no image has been uploaded yet for this TxnID (upload still in
//      flight, failed, or sync wasn't configured), a plain HTML receipt is
//      reconstructed here instead from SalesLog + the day-tab + Discounts +
//      StoreSettings — not pixel-identical (no delivery specifics, no
//      barcode) but always available the instant the sale itself has synced.
//
// TxnID alone is NOT guaranteed globally unique — each till device keeps
// its own local "T-0001, T-0002…" counter, so two different stores can
// produce the same ID. The "s" (store) query param disambiguates that;
// when present, only a match on BOTH TxnID and Store is used, for both the
// image lookup and the SalesLog lookup.
function _renderReceiptPage(txnId, store) {
  const imgBase64 = _findReceiptImageBase64(txnId, store);
  const html = imgBase64 ? _receiptImagePageShell(imgBase64, txnId) : _buildReceiptHtml(txnId, store);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Reçu ' + txnId)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Receipt image storage (Drive) ───────────────────────────────────────
// One PNG per receipt, named deterministically from TxnID+Store so a
// lookup never needs an index sheet. Stored in a single Drive folder this
// script creates on first use (same account the spreadsheet lives in,
// since this runs "as me" regardless of who opens the receipt link).
const RECEIPT_IMAGE_FOLDER_NAME = 'Boutique Scanner - Reçus';

function _receiptImageFileName(txnId, store) {
  const safe = function (v) { return String(v || '').replace(/[^a-zA-Z0-9_-]+/g, '_'); };
  return 'receipt_' + safe(store || 'store') + '_' + safe(txnId || 'txn') + '.png';
}

function _receiptImageFolder() {
  const it = DriveApp.getFoldersByName(RECEIPT_IMAGE_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(RECEIPT_IMAGE_FOLDER_NAME);
}

// Called from doPost's 'uploadReceiptImage' action, fired by the till in
// the background right when a WhatsApp receipt is sent (see
// uploadReceiptImageForLink() in Boutique_Scanner.html) — never blocks the
// send itself. Overwrites any previous upload for the same TxnID+Store so
// re-sending the same receipt (or a customer-modal retry) doesn't pile up
// duplicate files in the folder over time.
function _saveReceiptImage(d) {
  try {
    const folder = _receiptImageFolder();
    const fileName = _receiptImageFileName(d.txnId, d.store);
    const existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) existing.next().setTrashed(true);
    const bytes = Utilities.base64Decode(d.imageBase64);
    folder.createFile(Utilities.newBlob(bytes, 'image/png', fileName));
    return _json({ ok: true });
  } catch (err) {
    return _json({ error: String(err) });
  }
}

// Returns the stored receipt's base64 PNG data, or null if none has been
// uploaded yet for this TxnID+Store (caller falls back to the plain HTML
// receipt in that case — see _renderReceiptPage above).
function _findReceiptImageBase64(txnId, store) {
  try {
    const files = _receiptImageFolder().getFilesByName(_receiptImageFileName(txnId, store));
    if (files.hasNext()) return Utilities.base64Encode(files.next().getBlob().getBytes());
  } catch (err) {
    Logger.log('Receipt image lookup failed: ' + err);
  }
  return null;
}

function _receiptImagePageShell(imgBase64, txnId) {
  return '<!doctype html><html lang="fr"><head><meta charset="utf-8">' +
    '<title>Reçu ' + _escHtml(txnId) + '</title><style>' +
    'body{margin:0;padding:28px 12px;background:#f2f2f2;display:flex;justify-content:center;' +
    'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif}' +
    '.wrap{max-width:380px;width:100%}' +
    'img{width:100%;display:block;border-radius:10px;box-shadow:0 2px 10px rgba(0,0,0,.12)}' +
    '</style></head><body><div class="wrap">' +
    '<img src="data:image/png;base64,' + imgBase64 + '" alt="Reçu ' + _escHtml(txnId) + '">' +
    '</div></body></html>';
}

// Optional cosmetic shortening of the receipt link via TinyURL's free,
// keyless endpoint (server-side UrlFetchApp call — no browser CORS concern
// the way a client-side fetch to a third-party host would have). Best
// effort only: on any failure this returns null and the caller just keeps
// using the full (longer, but fully working) link — see doGet's
// 'shortenReceiptLink' branch and receiptLinkLine() in Boutique_Scanner.html.
function _shortenUrl(longUrl) {
  try {
    const resp = UrlFetchApp.fetch(
      'https://tinyurl.com/api-create.php?url=' + encodeURIComponent(longUrl),
      { muteHttpExceptions: true }
    );
    const text = (resp.getContentText() || '').trim();
    if (resp.getResponseCode() === 200 && /^https?:\/\//.test(text)) return text;
  } catch (err) {
    Logger.log('URL shortening failed: ' + err);
  }
  return null;
}

function _findSaleRow(txnId, store) {
  const sh = _sheet(SHEETS.SALES,
    ['TxnID', 'Date', 'Time', 'Employee', 'Store', 'ItemCount', 'Total', 'PayMethod',
      'PayMethodLabel', 'SplitPayments', 'IsDelivery', 'DeliveryID', 'CustomerName', 'Items']);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) { // newest first
    if (String(rows[i][0]) === String(txnId) && (!store || String(rows[i][4]) === String(store))) {
      return rows[i];
    }
  }
  return null;
}

// Pulls the itemized breakdown from that sale's day-tab (see SALES_ITEM_HEADERS
// near _dayTab below) — the only place unit price / qty / line total per item
// is recorded; SalesLog itself only keeps a flattened name summary.
function _findSaleItems(txnId, store, frDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(_dayTabName(frDate));
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const items = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; // Date,Time,Type,Barcode,Product,Category,Qty,UnitPrice,LineTotal,PayMethod,PayMethodLabel,Employee,Store,TxnID,RefID,Processed
    if (r[2] === 'Sale' && String(r[13]) === String(txnId) && (!store || String(r[12]) === String(store))) {
      items.push({ name: r[4], qty: Number(r[6]) || 1, unitPrice: Number(r[7]) || 0, lineTotal: Number(r[8]) || 0 });
    }
  }
  return items;
}

function _findSaleDiscount(txnId, store) {
  const sh = _sheet(SHEETS.DISCOUNTS, DISCOUNT_HEADERS);
  const rows = sh.getDataRange().getValues();
  let amountOff = 0;
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i]; // Date,Time,TxnID,Scope,ItemName,DiscountType,DiscountValue,AmountOff,AppliedBy,Employee,Store
    if (String(r[2]) === String(txnId) && r[3] === 'sale' && (!store || String(r[10]) === String(store))) {
      amountOff += Number(r[7]) || 0;
    }
  }
  return amountOff;
}

function _getStoreBrandingFor(store) {
  const sh = _sheet(SHEETS.STORE_SETTINGS, STORE_SETTINGS_HEADERS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === store) {
      return { icon: rows[i][1], name: rows[i][2], tagline: rows[i][3], phone: rows[i][4], address: rows[i][5], closingMsg: rows[i][6] };
    }
  }
  return null;
}

// French-style grouping (space-separated thousands) written by hand rather
// than Number.prototype.toLocaleString('fr-FR') — Apps Script's V8 runtime
// doesn't reliably ship full ICU locale data server-side, so relying on it
// here risked silently falling back to plain unformatted numbers.
function _fmtCFA(n) {
  const num = Math.round(Number(n) || 0);
  const grouped = String(Math.abs(num)).replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
  return (num < 0 ? '-' : '') + grouped + ' CFA';
}

function _escHtml(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function _buildReceiptHtml(txnId, store) {
  const saleRow = _findSaleRow(txnId, store);
  if (!saleRow) return _receiptPageShell(
    '<div class="r-center"><h1 style="font-size:17px;margin:0 0 8px">Reçu introuvable</h1>' +
    '<p style="font-size:13px;color:#666;margin:0">Le reçu N° ' + _escHtml(txnId) +
    ' n’a pas encore été synchronisé, ou le lien est incorrect. ' +
    'Réessayez dans quelques instants, ou contactez la boutique.</p></div>', txnId);

  // SalesLog: TxnID,Date,Time,Employee,Store,ItemCount,Total,PayMethod,PayMethodLabel,SplitPayments,IsDelivery,DeliveryID,CustomerName,Items
  const date = _frDateStr(saleRow[1], CONFIG.TIMEZONE);
  const time = saleRow[2];
  const employee = saleRow[3];
  const storeName = saleRow[4];
  const total = Number(saleRow[6]) || 0;
  const payMethodLabel = saleRow[8] || '';
  const isDelivery = !!saleRow[10];
  const deliveryId = saleRow[11] || '';

  const items = _findSaleItems(txnId, storeName, date);
  const discount = _findSaleDiscount(txnId, storeName);
  const s = _getStoreBrandingFor(storeName) || {};

  let itemsHtml = '';
  if (items.length) {
    itemsHtml += '<div class="r-item-head"><span>Désignation</span><span class="r-col-qte">Qté</span>' +
      '<span class="r-col-pu">Prix U.</span><span class="r-col-mt">Montant</span></div>';
    items.forEach(function (it) {
      itemsHtml += '<div class="r-item-row"><span class="r-name">' + _escHtml(it.name) + '</span>' +
        '<span class="r-col-qte">' + it.qty + '</span>' +
        '<span class="r-col-pu">' + _fmtCFA(it.unitPrice).replace(' CFA', '') + '</span>' +
        '<span class="r-col-mt">' + _fmtCFA(it.lineTotal).replace(' CFA', '') + '</span></div>';
    });
  } else if (saleRow[13]) {
    String(saleRow[13]).split(' | ').filter(Boolean).forEach(function (n) {
      itemsHtml += '<div class="r-row"><span class="r-name">' + _escHtml(n) + '</span><span></span></div>';
    });
  }

  const discountHtml = discount > 0
    ? '<div class="r-row"><span>Remise</span><span>-' + _fmtCFA(discount) + '</span></div>' : '';
  const deliveryLine = isDelivery
    ? '<div class="r-row"><span>Livraison N°</span><span>' + _escHtml(deliveryId) + '</span></div>' : '';

  let header = '<div class="r-header">';
  if (s.icon) header += '<div class="r-icon">' + _escHtml(s.icon) + '</div>';
  if (s.name) header += '<div class="r-title">' + _escHtml(s.name) + '</div>';
  if (s.tagline) header += '<div class="r-tagline">' + _escHtml(s.tagline) + '</div>';
  if (s.phone) header += '<div class="r-sub">Tél: ' + _escHtml(s.phone) + '</div>';
  if (s.address) header += '<div class="r-sub">' + _escHtml(s.address) + '</div>';
  header += '</div>';

  const closing = s.closingMsg || 'Revenez bientôt !';

  const body = header +
    '<div class="r-div"></div>' +
    '<div class="r-row"><span>Reçu N°</span><span>' + _escHtml(txnId) + '</span></div>' +
    deliveryLine +
    '<div class="r-row"><span>Date</span><span>' + _escHtml(date) + ' ' + _escHtml(time) + '</span></div>' +
    '<div class="r-row"><span>Vendeur</span><span>' + _escHtml(employee || '—') + '</span></div>' +
    '<div class="r-div"></div>' +
    itemsHtml +
    '<div class="r-div"></div>' +
    discountHtml +
    '<div class="r-total-row"><span>TOTAL</span><span>' + _fmtCFA(total) + '</span></div>' +
    '<div class="r-div"></div>' +
    (payMethodLabel ? '<div class="r-row"><span>Mode de paiement</span><span>' + _escHtml(payMethodLabel) + '</span></div>' : '') +
    '<div class="r-center" style="margin-top:14px;font-style:italic">' + _escHtml(closing) + '</div>';

  return _receiptPageShell(body, txnId);
}

// Same visual language as the printed/on-screen receipt in Boutique_Scanner.html
// (.receipt-paper / .r-* classes) — kept as one shared inline stylesheet here
// since this page is served standalone, with no access to the scanner's CSS.
function _receiptPageShell(bodyHtml, txnId) {
  return '<!doctype html><html lang="fr"><head><meta charset="utf-8">' +
    '<title>Reçu ' + _escHtml(txnId) + '</title><style>' +
    'body{margin:0;padding:28px 12px;background:#f2f2f2;' +
    'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Arial,sans-serif;' +
    'display:flex;justify-content:center}' +
    '.receipt-paper{background:#fff;max-width:360px;width:100%;border-radius:10px;' +
    'box-shadow:0 2px 10px rgba(0,0,0,.08);padding:20px;font-size:13px;line-height:1.7;' +
    'font-family:"Courier New",Courier,monospace;color:#222}' +
    '.r-center{text-align:center}.r-header{text-align:center;margin-bottom:4px}' +
    '.r-icon{font-size:32px;line-height:1;margin-bottom:2px}' +
    '.r-title{font-family:"Brush Script MT","Segoe Script","Bradley Hand",cursive;' +
    'font-size:24px;font-weight:400;text-align:center;margin-bottom:2px;color:#222}' +
    '.r-tagline{text-align:center;font-size:10px;color:#888;letter-spacing:1.5px;' +
    'text-transform:uppercase;margin-bottom:4px}' +
    '.r-sub{text-align:center;font-size:11px;color:#666;margin-bottom:2px}' +
    '.r-div{border-top:1px dashed #aaa;margin:8px 0}' +
    '.r-row{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}' +
    '.r-name{flex:1;min-width:0;overflow:hidden;white-space:nowrap;text-overflow:ellipsis}' +
    '.r-item-head{display:grid;grid-template-columns:1fr 28px 55px 60px;gap:6px;font-size:10px;' +
    'font-weight:700;color:#888;text-transform:uppercase;letter-spacing:.5px;padding-bottom:3px;' +
    'border-bottom:1px solid #ddd;margin-bottom:3px}' +
    '.r-item-row{display:grid;grid-template-columns:1fr 28px 55px 60px;gap:6px;margin:3px 0}' +
    '.r-col-qte,.r-col-pu,.r-col-mt{text-align:right;white-space:nowrap}' +
    '.r-total-row{display:flex;justify-content:space-between;font-weight:700;font-size:15px}' +
    '</style></head><body><div class="receipt-paper">' + bodyHtml + '</div></body></html>';
}

// Called at the end of every write handler below. Wrapped defensively —
// recording the actual sale/return/payment/hold always matters more than
// the Dashboard tab being current, so a dashboard hiccup must never make
// the real write look like it failed.
function _refreshDashboard() {
  try { updateDashboard(); } catch (e) { Logger.log('Dashboard refresh skipped: ' + e); }
}

// ══════════════════════════════════════════════════════════
//  SHEET HELPERS
// ══════════════════════════════════════════════════════════
// Google Sheets auto-detects text that "looks like" a date/time/number and
// silently converts it to a real Date or Number value — even when written
// via Apps Script, not just typed by hand. That breaks every r[i] === today
// string comparison this whole file relies on (Dashboard totals, nightly
// report filters, day-tab lookups) the moment it happens, with no error —
// the row is just invisible to those checks. These columns hold values that
// must stay literal text for that reason: 1-indexed column numbers, per
// sheet name.
const TEXT_COLUMNS = {
  'SalesLog': [2, 3],              // Date, Time
  'Returns': [3, 4],               // Date, Time
  'Customers': [1, 4],             // Phone, LastPurchase
  'InstallmentPlans': [3, 4, 5, 14, 16], // Date, Time, Phone, LastPaymentDate, NextDue
  'InstallmentPayments': [2, 3, 4],  // Phone, Date, Time
  'Holds': [2, 6, 7, 8],           // Phone, Date, Time, ExpiryDate
  'Discounts': [1, 2],             // Date, Time
  'ShiftReconciliation': [1, 2],   // Date, Time
  'Expenses': [1, 2],              // Date, Time
  'Employees': [2],                // Pin (keep leading zeros / exact digits)
  'GlobalSettings': [1],           // AdminPin
};

// Forces Plain Text ('@') number format on a column for a generous number
// of rows, so any value written into it from here on is kept exactly as
// typed — Sheets only auto-converts on write when a cell's format is still
// "Automatic," so doing this once per column is enough to stop it for good.
function _forceTextColumns(sh, cols) {
  cols.forEach(col => {
    sh.getRange(1, col, Math.max(sh.getMaxRows(), 2000), 1).setNumberFormat('@');
  });
}

// Second line of defense against the exact drift described above: even if
// a Date/Time column's format silently reverts to Automatic again and a
// cell comes back as a real Date object instead of text, every "is this
// today / this month" comparison in the file should still work instead of
// silently excluding the row. Normalizes a SalesLog/Returns Date-column
// cell to the canonical 'dd/MM/yyyy' text form used everywhere else.
function _frDateStr(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz || CONFIG.TIMEZONE, 'dd/MM/yyyy');
  return String(v == null ? '' : v);
}

function _sheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    if (TEXT_COLUMNS[name]) _forceTextColumns(sh, TEXT_COLUMNS[name]);
  }
  return sh;
}

// ══════════════════════════════════════════════════════════
//  PER-DAY ITEM FEED — for the Inventory Sync below
// ══════════════════════════════════════════════════════════
// One tab per calendar day (e.g. "Sales_2026-08-10"), one row per item sold
// or returned, always carrying Barcode + Product so a downstream inventory
// script can match on barcode alone. Qty is signed — positive for a sale,
// negative for a return — so that script can apply every row with the same
// stock -= Qty logic regardless of type. A blank "Processed" column is left
// for that script to mark once it has applied a row, so repeated reads never
// double-count.
//
// The tab name is derived from the item's own Date (not the server clock)
// so an item scanned late at night but synced after midnight (offline
// queue/retry) still lands on the day it actually happened.
const SALES_ITEM_HEADERS = ['Date', 'Time', 'Type', 'Barcode', 'Product', 'Category',
  'Qty', 'UnitPrice', 'LineTotal', 'PayMethod', 'PayMethodLabel', 'Employee', 'Store',
  'TxnID', 'RefID', 'Processed', 'ScanID', 'Remise'];
// Remise (col 18) = amount taken off the line (negative for a majoration).
// UnitPrice stays the normal price; LineTotal is what was actually paid.
// ScanID (col 17) added so a scanned-but-not-yet-paid item can be voided
// (removed from the cart, or the whole sale cancelled) before checkout —
// see _voidScan() below. Appended at the end so every existing hardcoded
// column index above (Processed = col 16, etc.) stays correct unchanged.

function _dayTabName(frDate) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(frDate || '').trim());
  if (m) return 'Sales_' + m[3] + '-' + m[2] + '-' + m[1]; // dd/MM/yyyy -> Sales_yyyy-MM-dd
  return 'Sales_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

// ══════════════════════════════════════════════════════════
//  DAY-TAB TIDINESS
//  1. Sales_yyyy-MM-dd tabs are kept in date order (newest first, right after
//     Dashboard) — Sheets would otherwise drop each new tab next to whatever
//     tab was last in use.
//  2. "Voided" lines (the undo record left when a quantity or price is changed
//     or an item removed) are shown in light grey.
//  3. A live totals block sits at the top right of every day tab (columns S:Z):
//     net total per store, and per store × payment method. These are plain
//     QUERY formulas over the tab's own rows, so they follow every new sale
//     with no script run. Voided lines count for 0; returns are negative.
//  Everything here is best-effort: it can never block a sale from being logged.
//  Run organizeTabs() once from the Apps Script editor to tidy existing tabs.
// ══════════════════════════════════════════════════════════
function _sortDayTabs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dayRe = /^Sales_(\d{4}-\d{2}-\d{2})$/;
  const days = ss.getSheets().filter(s => dayRe.test(s.getName()))
    .sort((a, b) => b.getName().localeCompare(a.getName())); // newest first
  if (!days.length) return 0;
  const prev = ss.getActiveSheet();
  let moved = 0;
  const dash = ss.getSheetByName('Dashboard');
  if (dash && dash.getIndex() !== 1) { ss.setActiveSheet(dash); ss.moveActiveSheet(1); moved++; }
  let pos = dash ? 2 : 1;
  days.forEach(s => {
    if (s.getIndex() !== pos) { ss.setActiveSheet(s); ss.moveActiveSheet(pos); moved++; }
    pos++;
  });
  if (prev) ss.setActiveSheet(prev);
  return moved;
}

// Formulas are written in the usual comma style, then read back: if the
// Sheet shows #ERROR! (a formula it cannot read — e.g. a Sheet set to a
// language that separates arguments with ";"), the same formula is written
// again with ";" and checked again. Whichever works is remembered.
function _fxSemi(f) {
  let out = '', q = false;
  for (let i = 0; i < f.length; i++) {
    const ch = f[i];
    if (ch === '"') q = !q;
    out += (!q && ch === ',') ? ';' : ch;
  }
  return out;
}
function _setFormulaSafe(rng, f) {
  const props = PropertiesService.getScriptProperties();
  const pref = props.getProperty('FORMULA_SEP');
  const order = pref === 'semi' ? [true, false] : [false, true];
  for (let k = 0; k < order.length; k++) {
    rng.setFormula(order[k] ? _fxSemi(f) : f);
    SpreadsheetApp.flush();
    if (String(rng.getDisplayValue()).indexOf('#ERROR') !== 0) {
      if (pref !== (order[k] ? 'semi' : 'comma')) props.setProperty('FORMULA_SEP', order[k] ? 'semi' : 'comma');
      return true;
    }
  }
  Logger.log('Formule illisible dans ' + rng.getSheet().getName() + '!' + rng.getA1Notation());
  return false;
}
const _TOTALS_S2 = '=IFERROR(QUERY($A$1:$Q,"select M, sum(I) where C=\'Sale\' or C=\'Return\' or C=\'Remise\' group by M label M \'Magasin\', sum(I) \'Total\'",1),"Aucune vente")';
const _TOTALS_W2 = '=IFERROR(QUERY($A$1:$Q,"select M, sum(I) where (C=\'Sale\' or C=\'Return\' or C=\'Remise\') group by M pivot J label M \'Magasin\'",1),"")';
const _REMISE_T8 = '=SUMIF($C:$C,"Sale",$R:$R)+SUMIF($C:$C,"Remise",$R:$R)';
function _setTotalsFormulas(sh) {
  _setFormulaSafe(sh.getRange('S2'), _TOTALS_S2);
  _setFormulaSafe(sh.getRange('W2'), _TOTALS_W2);
}

function _setRemiseTotal(sh, force) {
  if (!force && String(sh.getRange('S8').getValue() || '')) return;
  sh.getRange('S8').setValue('Remises du jour (CFA)').setFontWeight('bold');
  sh.getRange('T8').setNumberFormat('#,##0');
  _setFormulaSafe(sh.getRange('T8'), _REMISE_T8);
}
// Does this day tab show #ERROR! in its totals block?
function _totalsBroken(sh) {
  if (!/^Totaux du jour par magasin/.test(String(sh.getRange('S1').getValue() || ''))) return false;
  return ['S2', 'W2', 'T8'].some(a => String(sh.getRange(a).getDisplayValue()).indexOf('#ERROR') === 0);
}
// Run ONCE from the editor: rewrites the totals block of every day tab that
// shows #ERROR! (and of the 60 most recent tabs), so they all calculate.
function repairDayTotals() {
  const started = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dayRe = /^Sales_(\d{4}-\d{2}-\d{2})$/;
  let fixed = 0, still = 0;
  ss.getSheets().filter(s => dayRe.test(s.getName())).sort((a, b) => b.getName().localeCompare(a.getName())).forEach(s => {
    if (Date.now() - started > 300000) return;
    try {
      if (!String(s.getRange('S1').getValue() || '')) { _decorateDayTab(s); fixed++; return; }
      if (!_totalsBroken(s)) return;
      _setTotalsFormulas(s);
      _setRemiseTotal(s, true);
      if (_totalsBroken(s)) still++; else fixed++;
    } catch (e) { Logger.log('repair ' + s.getName() + ': ' + e); still++; }
  });
  Logger.log('Totaux réparés : ' + fixed + ' onglet(s)' + (still ? ' — encore en erreur : ' + still : '') +
    ' — séparateur utilisé : ' + (PropertiesService.getScriptProperties().getProperty('FORMULA_SEP') || '?'));
}

function _hasFormatRule(sh, needle) {
  try {
    return sh.getConditionalFormatRules().some(r => {
      const c = r.getBooleanCondition();
      return c && (c.getCriteriaValues() || []).some(v => String(v).indexOf(needle) >= 0);
    });
  } catch (e) { return false; }
}

function _addRemiseFormats(sh) {
  const rules = sh.getConditionalFormatRules();
  const rng = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 18);
  let added = false;
  if (!_hasFormatRule(sh, '$R2>0')) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($R2),$R2>0,$C2<>"Voided")').setBackground('#ffe0b2').setRanges([rng]).build()); added = true;
  }
  if (!_hasFormatRule(sh, '$R2<0')) {
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(ISNUMBER($R2),$R2<0,$C2<>"Voided")').setBackground('#dbeafe').setRanges([rng]).build()); added = true;
  }
  if (added) sh.setConditionalFormatRules(rules);
}

function _decorateDayTab(sh) {
  // Totals block (only if the spot is free, so a hand-made block is never overwritten)
  if (!String(sh.getRange('S1').getValue() || '')) {
    sh.getRange('S1').setValue('Totaux du jour par magasin (net, retours déduits)').setFontWeight('bold');
    sh.getRange('W1').setValue('Par mode de paiement').setFontWeight('bold');
    _setTotalsFormulas(sh);
    sh.getRange('T2:T12').setNumberFormat('#,##0');
    sh.getRange('X2:AD12').setNumberFormat('#,##0');
    sh.setColumnWidth(19, 150);
    _setRemiseTotal(sh);
  }
  // Grey out Voided lines
  const rules = sh.getConditionalFormatRules();
  const rng = sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 17);
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="Voided"').setFontColor('#a0a0a0').setRanges([rng]).build());
  sh.setConditionalFormatRules(rules);
  _addRemiseFormats(sh);
}

// Brings an already-decorated (older) day tab up to date: Remise header and
// colours, the totals formulas that count whole-sale remise rows, and a
// "Remises du jour" line. Only touches the totals block if it is ours.
function _upgradeDayTab(sh) {
  if (sh.getMaxColumns() < 18) sh.insertColumnsAfter(sh.getMaxColumns(), 18 - sh.getMaxColumns());
  if (!sh.getRange(1, 18).getValue()) sh.getRange(1, 18).setValue('Remise');
  sh.getRange(2, 18, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('#,##0');
  if (/^Totaux du jour par magasin/.test(String(sh.getRange('S1').getValue() || ''))) {
    _setTotalsFormulas(sh);
    _setRemiseTotal(sh, true);
  }
  if (!_hasFormatRule(sh, '$C2="Voided"')) {
    const rules = sh.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$C2="Voided"').setFontColor('#a0a0a0').setRanges([sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), 17)]).build());
    sh.setConditionalFormatRules(rules);
  }
  _addRemiseFormats(sh);
}

// Voided lines are the undo record of a cancelled/edited line. They are hidden
// by default (rows are hidden, never deleted; totals and stock are unaffected).
// show=true brings them back into view.
function _setVoidedVisible(sh, show) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const types = sh.getRange(2, 3, last - 1, 1).getValues();
  let n = 0, i = 0;
  while (i < types.length) {
    if (types[i][0] === 'Voided') {
      let j = i;
      while (j + 1 < types.length && types[j + 1][0] === 'Voided') j++;
      if (show) sh.showRows(i + 2, j - i + 1); else sh.hideRows(i + 2, j - i + 1);
      n += j - i + 1; i = j + 1;
    } else i++;
  }
  return n;
}

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('Boutique 1981')
      .addItem('Afficher les lignes annulées (onglet actuel)', 'menuShowVoided')
      .addItem('Masquer les lignes annulées (onglet actuel)', 'menuHideVoided')
      .addItem('Afficher les lignes annulées (tous les onglets récents)', 'menuShowVoidedAll')
      .addItem('Masquer les lignes annulées (tous les onglets récents)', 'menuHideVoidedAll')
      .addToUi();
  } catch (e) {}
}
function menuShowVoided() { _setVoidedVisible(SpreadsheetApp.getActiveSheet(), true); }
function menuHideVoided() { _setVoidedVisible(SpreadsheetApp.getActiveSheet(), false); }
function _eachRecentDayTab(fn) {
  const started = Date.now();
  SpreadsheetApp.getActiveSpreadsheet().getSheets().filter(s => /^Sales_\d{4}-\d{2}-\d{2}$/.test(s.getName())).slice(0, 60).forEach(s => {
    if (Date.now() - started > 240000) return;
    try { fn(s); } catch (e) { Logger.log(s.getName() + ': ' + e); }
  });
}
function menuShowVoidedAll() { _eachRecentDayTab(s => _setVoidedVisible(s, true)); }
function menuHideVoidedAll() { _eachRecentDayTab(s => _setVoidedVisible(s, false)); }

/** Run once from the editor: sorts every day tab and decorates the most recent ones. */
function organizeTabs() {
  const started = Date.now();
  const moved = _sortDayTabs();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dayRe = /^Sales_(\d{4}-\d{2}-\d{2})$/;
  let done = 0, upgraded = 0, hidden = 0;
  ss.getSheets().filter(s => dayRe.test(s.getName())).slice(0, 60).forEach(s => {
    if (Date.now() - started > 240000) return; // stay within the run-time limit
    try {
      const has = String(s.getRange('S1').getValue() || '');
      if (!has) { _decorateDayTab(s); done++; }
      if (sheetNeedsRemise_(s)) { _upgradeDayTab(s); upgraded++; }
      else if (_totalsBroken(s)) { _setTotalsFormulas(s); _setRemiseTotal(s, true); upgraded++; }
      hidden += _setVoidedVisible(s, false);
    } catch (e) { Logger.log('decorate ' + s.getName() + ': ' + e); }
  });
  Logger.log('Tabs moved: ' + moved + ' — decorated: ' + done + ' — upgraded: ' + upgraded + ' — voided rows hidden: ' + hidden);
}
function sheetNeedsRemise_(s) { return !String(s.getRange(1, 18).getValue() || '') || !String(s.getRange('S8').getValue() || ''); }

function _dayTab(frDate) {
  const name = _dayTabName(frDate);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const isNew = !ss.getSheetByName(name);
  const sh = _sheet(name, SALES_ITEM_HEADERS);
  if (isNew) {
    _forceTextColumns(sh, [1, 2, 4, 17]); // Date, Time, Barcode, ScanID
    try { _decorateDayTab(sh); } catch (e) { Logger.log('day-tab decoration skipped: ' + e); }
    try { _sortDayTabs(); } catch (e) { Logger.log('day-tab sort skipped: ' + e); }
  } else if (sh.getLastColumn() < 17 || !sh.getRange(1, 17).getValue()) {
    // Self-heal: a day-tab created before ScanID existed. Add the header
    // now so rows appended from here on line up under it correctly.
    sh.getRange(1, 17).setValue('ScanID');
    _forceTextColumns(sh, [17]);
  }
  if (!isNew) {
    // Self-heal: Remise column (18) for tabs created before remises were shown here.
    try {
      if (sh.getMaxColumns() < 18) sh.insertColumnsAfter(sh.getMaxColumns(), 18 - sh.getMaxColumns());
      if (!sh.getRange(1, 18).getValue()) sh.getRange(1, 18).setValue('Remise');
    } catch (e) { Logger.log('Remise header skipped: ' + e); }
  }
  return sh;
}

// ══════════════════════════════════════════════════════════
//  WRITE HANDLERS (called from the scanner in real time)
// ══════════════════════════════════════════════════════════
function _logScan(d) {
  // Written the instant an item is scanned — before payment method is known
  // — so PayMethod starts blank here and _backfillPayMethod() (called from
  // _logSale below) fills it in once the sale closes.
  const sh = _dayTab(d.date);
  sh.appendRow([d.date, d.time, 'Sale', d.barcode, d.name, d.category,
    d.qty, d.price, d.lineTotal, '', '', d.employee, d.store, d.txnId, d.txnId, '', d.scanId || '']);
  return _json({ ok: true });
}

function _logSale(d) {
  const sh = _sheet(SHEETS.SALES,
    ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  sh.appendRow([d.txnId, d.date, d.time, d.employee, d.store, d.itemCount, d.total,
    d.payMethod, d.payMethodLabel, d.splitPayments || '', d.isDelivery, d.deliveryId, d.customerName, d.items]);
  _backfillPayMethod(d);
  _applyRemisesToDayTab(d);
  _refreshDashboard();
  return _json({ ok: true });
}

// ══════════════════════════════════════════════════════════
//  PAPER SALES — sales written on paper (cashier ill, power cut…) typed in
//  afterwards from Admin → "Calendrier des ventes". Each one is written as
//  a normal closed sale on ITS OWN date: a SalesLog row (TxnID PAPIER-…,
//  CustomerName "Papier") + one 'Sale' line per item in that day's tab, so
//  the Dashboard, day totals, weekly email and backups all include it.
//  The day-tab lines are born with Processed = "PAPIER (pas de stock)": the
//  15-min stock sync skips any line whose Processed is filled, so paper
//  sales never change stock (owner's choice, 17 Oct).
//  Idempotent: the same paperId sent twice (queue retry) is written once.
// ══════════════════════════════════════════════════════════
const PAPER_MARK = 'PAPIER (pas de stock)';
function _logPaperSale(d) {
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(String(d.date || ''))) return _json({ error: 'Date invalide (jj/mm/aaaa attendu).' });
  const items = (Array.isArray(d.items) ? d.items : []).filter(it => it && String(it.name || it.barcode || '').trim() && Number(it.qty));
  if (!items.length) return _json({ error: 'Aucun article.' });
  if (!d.paperId) return _json({ error: 'paperId manquant.' });
  const txnId = 'PAPIER-' + String(d.paperId);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const salesSh = _sheet(SHEETS.SALES,
      ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
    if (salesSh.getLastRow() > 1 && salesSh.getRange(2, 1, salesSh.getLastRow() - 1, 1).createTextFinder(txnId).matchEntireCell(true).findNext()) {
      return _json({ ok: true, txnId: txnId, duplicate: true });
    }
    const time = /^\d{1,2}:\d{2}$/.test(String(d.time || '')) ? String(d.time) : '12:00';
    const label = d.payMethodLabel || PAY_LABELS[d.payMethod] || d.payMethod || '';
    const store = String(d.store || '').trim(), emp = String(d.employee || '').trim() || 'Papier';
    let total = 0;
    const lines = items.map(it => {
      const qty = Number(it.qty) || 0, unit = Number(it.unitPrice) || 0, lt = Math.round(qty * unit);
      total += lt;
      return [d.date, time, 'Sale', String(it.barcode || ''), String(it.name || '').trim() || String(it.barcode), String(it.category || ''),
        qty, unit, lt, d.payMethod || '', label, emp, store, txnId, txnId, PAPER_MARK, '', 0];
    });
    const day = _dayTab(d.date);
    day.getRange(day.getLastRow() + 1, 1, lines.length, lines[0].length).setValues(lines);
    const summary = items.map(it => String(it.name || it.barcode) + (Number(it.qty) !== 1 ? ' x' + it.qty : '')).join(', ');
    salesSh.appendRow([txnId, d.date, time, emp, store, items.length, total, d.payMethod || '', label, '', false, '',
      'Papier' + (d.note ? ' — ' + String(d.note).slice(0, 120) : ''), summary]);
    try { _refreshDashboard(); } catch (e) {}
    return _json({ ok: true, txnId: txnId, total: total });
  } finally {
    lock.releaseLock();
  }
}
// Month view for the calendar: per day and per store, number of sales,
// total and how many of them are paper sales. month = "yyyy-MM".
function _getSalesCalendar(month) {
  const tz = CONFIG.TIMEZONE;
  if (!/^\d{4}-\d{2}$/.test(month)) month = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  const want = month.slice(5, 7) + '/' + month.slice(0, 4);
  const salesSh = _sheet(SHEETS.SALES,
    ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  const days = {}, stores = {};
  const vals = salesSh.getLastRow() > 1 ? salesSh.getRange(2, 1, salesSh.getLastRow() - 1, 7).getValues() : [];
  vals.forEach(r => {
    const store = String(r[4] || '').trim();
    if (store) stores[store] = true;
    const ds = _frDateStr(r[1], tz);
    if (ds.slice(3) !== want) return;
    const k = store || '(magasin ?)';
    const day = days[ds] = days[ds] || {};
    const c = day[k] = day[k] || { n: 0, total: 0, paper: 0 };
    c.n++; c.total += Number(r[6]) || 0;
    if (/^PAPIER-/.test(String(r[0]))) c.paper++;
  });
  return _json({ ok: true, month: month, days: days, stores: Object.keys(stores).sort() });
}

// Writes the remises of a closed sale into its day-tab lines: Remise (col 18)
// = amount taken off, LineTotal (col 9) = what was actually paid, UnitPrice
// stays the normal price. A whole-sale remise becomes its own 'Remise' row
// (negative LineTotal) so the day totals match the money really taken.
// Safe to run twice for the same sale (retries): it sets values, never adds.
function _applyRemisesToDayTab(d) {
  try {
    const lines = d.discLines || [];
    const sale = d.saleDiscount && Number(d.saleDiscount.amountOff) ? d.saleDiscount : null;
    if (!lines.length && !sale) return;
    const sh = _dayTab(d.date);
    const rows = sh.getDataRange().getValues();
    const used = {};
    lines.forEach(l => {
      let hit = -1;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][2] !== 'Sale' || String(rows[i][13]) !== String(d.txnId) || used[i]) continue;
        if (l.scanId && String(rows[i][16]) === String(l.scanId)) { hit = i; break; }
      }
      if (hit < 0) {
        for (let i = 1; i < rows.length; i++) {
          if (rows[i][2] !== 'Sale' || String(rows[i][13]) !== String(d.txnId) || used[i]) continue;
          if (String(rows[i][3]) === String(l.barcode || '') && String(rows[i][4]) === String(l.name || '')) { hit = i; break; }
        }
      }
      if (hit < 0) return;
      used[hit] = true;
      sh.getRange(hit + 1, 9).setValue(Number(l.lineTotal) || 0);
      sh.getRange(hit + 1, 18).setValue(Number(l.amountOff) || 0);
    });
    if (sale) {
      let exists = false;
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][2] === 'Remise' && String(rows[i][13]) === String(d.txnId)) { exists = true; break; }
      }
      if (!exists) {
        const off = Number(sale.amountOff) || 0;
        const label = d.payMethodLabel || PAY_LABELS[d.payMethod] || d.payMethod || '';
        sh.appendRow([d.date, d.time, 'Remise', '', (off < 0 ? 'Majoration sur la vente' : 'Remise sur la vente') + (sale.by ? ' (' + sale.by + ')' : ''), '',
          0, '', -off, d.payMethod || '', label, d.employee, d.store, d.txnId, d.txnId, '', '', off]);
      }
    }
  } catch (e) { Logger.log('Remise write skipped: ' + e); }
}

// Fills PayMethod/PayMethodLabel into every day-tab row from this TxnID
// (they're left blank by _logScan since payment isn't chosen until checkout).
function _backfillPayMethod(d) {
  try {
    const sh = _dayTab(d.date);
    const rows = sh.getDataRange().getValues();
    const label = d.payMethodLabel || PAY_LABELS[d.payMethod] || d.payMethod || '';
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][13] === d.txnId && rows[i][2] === 'Sale' && !rows[i][9]) {
        sh.getRange(i + 1, 10).setValue(d.payMethod || '');  // col 10 = PayMethod
        sh.getRange(i + 1, 11).setValue(label);               // col 11 = PayMethodLabel
      }
    }
  } catch (e) { Logger.log('Pay-method backfill skipped: ' + e); }
}

// ══════════════════════════════════════════════════════════
//  EDIT SALE PAYMENT METHOD — admin-only correction (Boutique_Scanner.html's
//  PIN-gated Admin panel, "Historique des ventes" card). Fixes the common
//  cashier mistake of logging Cash vs Orange Money vs Wave incorrectly,
//  after the sale has already closed.
// ══════════════════════════════════════════════════════════
// d = { key, action: 'editSalePayMethod', txnId, payMethod, payMethodLabel }
// Only ever sets one of the four simple methods (cash/orange/wave/bank) —
// switching a sale into/out of a split payment or a Mensualité plan isn't
// supported here, those have their own dedicated structures.
function _editSalePayMethod(d) {
  try {
    if (!d.txnId || !d.payMethod) return _json({ error: 'Missing txnId or payMethod' });
    const sh = _sheet(SHEETS.SALES,
      ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return _json({ error: 'SalesLog is empty' });

    const rows = sh.getRange(2, 1, lastRow - 1, 14).getValues();
    let saleDate = '';
    let oldLabel = '';
    let found = false;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] !== d.txnId) continue;
      const sheetRow = i + 2;
      oldLabel = rows[i][8] || rows[i][7] || '';
      saleDate = rows[i][1];
      sh.getRange(sheetRow, 8).setValue(d.payMethod);
      sh.getRange(sheetRow, 9).setValue(d.payMethodLabel || d.payMethod);
      sh.getRange(sheetRow, 10).setValue(''); // clear SplitPayments — this always sets one simple method
      found = true;
      break;
    }
    if (!found) return _json({ error: 'Sale not found: ' + d.txnId });

    // Mirror into that day's per-item feed too, so the barcode-level feed
    // agrees with the sale summary. Unlike _backfillPayMethod (which only
    // fills blanks right after a brand-new sale), this overwrites whatever
    // was there — that's the whole point of an edit.
    try {
      const dayTab = _dayTab(saleDate);
      const dLastRow = dayTab.getLastRow();
      if (dLastRow >= 2) {
        const dRows = dayTab.getRange(2, 1, dLastRow - 1, SALES_ITEM_HEADERS.length).getValues();
        for (let i = 0; i < dRows.length; i++) {
          if (dRows[i][13] !== d.txnId) continue; // col 14 = TxnID
          dayTab.getRange(i + 2, 10).setValue(d.payMethod);       // PayMethod
          dayTab.getRange(i + 2, 11).setValue(d.payMethodLabel || d.payMethod); // PayMethodLabel
        }
      }
    } catch (e) { Logger.log('Day-tab pay-method mirror skipped: ' + e); }

    _logSaleEdit(d.txnId, 'Mode de paiement', oldLabel, d.payMethodLabel || d.payMethod);
    _refreshDashboard();
    return _json({ ok: true });
  } catch (e) {
    return _json({ error: String(e) });
  }
}

// d = { key, action: 'editSaleTotal', txnId, total }
// Corrects only the sale's overall Total (typo, forgotten discount, etc).
// Deliberately does NOT touch item quantities, LineTotal on the per-item
// day-tab rows, or Store Inventory stock — those stay exactly as scanned.
// If the original sale was a split payment, the individual split amounts
// are left as-is too (they may no longer sum to the new Total) — fixing
// that is a per-item/split edit, out of scope for this quick correction.
function _editSaleTotal(d) {
  try {
    if (!d.txnId || d.total === undefined || d.total === null || isNaN(Number(d.total))) {
      return _json({ error: 'Missing or invalid txnId/total' });
    }
    const sh = _sheet(SHEETS.SALES,
      ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return _json({ error: 'SalesLog is empty' });

    const rows = sh.getRange(2, 1, lastRow - 1, 14).getValues();
    let oldTotal = null;
    let found = false;
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] !== d.txnId) continue;
      const sheetRow = i + 2;
      oldTotal = rows[i][6]; // Total, col 7
      sh.getRange(sheetRow, 7).setValue(Number(d.total));
      found = true;
      break;
    }
    if (!found) return _json({ error: 'Sale not found: ' + d.txnId });

    _logSaleEdit(d.txnId, 'Total', oldTotal, Number(d.total));
    _refreshDashboard();
    return _json({ ok: true });
  } catch (e) {
    return _json({ error: String(e) });
  }
}

// Lightweight audit trail — same philosophy as "Sync Issues" / "Stock
// Ledger": every admin correction to a past sale is visible in its own
// tab, not buried in execution logs, so the fix is itself never invisible.
// Shared by every kind of sale correction (payment method, total, and any
// future one) — "field" names which part of the sale changed.
function _logSaleEdit(txnId, field, oldVal, newVal) {
  try {
    const headers = ['Timestamp', 'TxnID', 'Champ modifié', 'Avant', 'Après'];
    const sh = _sheet('Sale Edits', headers);
    // Self-heal: upgrade a tab created by the earlier payment-method-only
    // version of this function (3 columns, no "Champ modifié").
    if (sh.getLastColumn() < headers.length || String(sh.getRange(1, 3).getValue()) !== headers[2]) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sh.appendRow([Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'), txnId, field, oldVal, newVal]);
  } catch (e) { Logger.log('Could not log sale edit: ' + e); }
}

function _logShift(d) {
  const sh = _sheet(SHEETS.SHIFTS, ['Employee','Store','ClockIn','ClockOut']);
  if (d.shiftAction === 'open') {
    sh.appendRow([d.employee, d.store, d.clockIn, '']);
  } else {
    // Close: find the most recent open row (no ClockOut yet) for this employee.
    const rows = sh.getDataRange().getValues();
    for (let i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0] === d.employee && !rows[i][3]) {
        sh.getRange(i + 1, 4).setValue(d.clockOut || new Date());
        break;
      }
    }
    _logShiftReconciliation(d);
  }
  return _json({ ok: true });
}

// Cash + Orange Money reconciliation — separate sheet (not extending the
// existing 4-column ShiftLog above) to avoid a live-schema migration risk,
// same reasoning as the Discounts tab. Only writes a row when the client
// actually sent reconciliation fields (older client versions still just
// clock in/out with no balances, and that keeps working exactly as before).
const RECONCILIATION_HEADERS = ['Date', 'Time', 'Employee', 'Store', 'ClockIn', 'ClockOut',
  'StartingCash', 'EndingCash', 'ExpectedCash', 'CashDiscrepancy',
  'StartingOrange', 'EndingOrange', 'ExpectedOrange', 'OrangeDiscrepancy'];

function _logShiftReconciliation(d) {
  if (d.endingCash === undefined && d.endingOrange === undefined) return; // older client, nothing to log
  const sh = _sheet(SHEETS.RECONCILIATION, RECONCILIATION_HEADERS);
  const tz = CONFIG.TIMEZONE;
  const now = new Date();
  sh.appendRow([
    Utilities.formatDate(now, tz, 'dd/MM/yyyy'), Utilities.formatDate(now, tz, 'HH:mm'),
    d.employee, d.store, d.clockIn, d.clockOut,
    d.startingCash || 0, d.endingCash || 0, d.expectedCash || 0, d.cashDiscrepancy || 0,
    d.startingOrange || 0, d.endingOrange || 0, d.expectedOrange || 0, d.orangeDiscrepancy || 0,
  ]);
}

// Customers tab with the NotesJSON column (customer requests / notes) added to
// a sheet created before it existed.
function _customersSheet() {
  const sh = _sheet(SHEETS.CUSTOMERS, ['Phone','Name','Store','LastPurchase','NotesJSON']);
  if (String(sh.getRange(1, 5).getValue() || '') !== 'NotesJSON') sh.getRange(1, 5).setValue('NotesJSON');
  return sh;
}

// Union of two JSON note lists by note id; a note marked done on either side stays done.
function _mergeNotesJson(a, b) {
  const parse = s => { try { const v = s ? JSON.parse(s) : []; return Array.isArray(v) ? v : []; } catch (e) { return []; } };
  const map = {};
  parse(a).concat(parse(b)).forEach(n => {
    if (!n || !n.id) return;
    if (!map[n.id]) map[n.id] = n; else map[n.id].done = !!(map[n.id].done || n.done);
  });
  return JSON.stringify(Object.keys(map).map(k => map[k]));
}

function _saveCustomer(d) {
  const sh = _customersSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.phone) {
      sh.getRange(i + 1, 2, 1, 3).setValues([[d.name || rows[i][1], d.store || rows[i][2], d.lastPurchase || rows[i][3]]]);
      if (d.notesJson) sh.getRange(i + 1, 5).setValue(_mergeNotesJson(rows[i][4], d.notesJson));
      return _json({ ok: true });
    }
  }
  sh.appendRow([d.phone, d.name || '', d.store || '', d.lastPurchase || '', d.notesJson || '']);
  return _json({ ok: true });
}

function _logCashCheck(d) {
  const sh = _sheet('Cash Checks', ['Timestamp','Date','Store','By','BillsJSON','Coins','CashCounted','CashExpected','CashDiff','DigitalTotal','DigitalConfirmed','DigitalUnconfirmed','Unconfirmed','Note']);
  sh.appendRow([d.ts || new Date().toISOString(), d.date, d.store || '', d.by || '', d.billsJson || '', d.coins || 0,
    d.cashCounted || 0, d.cashExpected || 0, d.cashDiff || 0, d.digitalTotal || 0, d.digitalConfirmed || 0, d.digitalMissing || 0, d.unconfirmed || '', d.note || '']);
  return _json({ ok: true });
}

// Units + revenue per product over the last N days, read from the per-day
// item tabs (all stores). Voided rows have Qty 0; returns are negative, so
// summing nets them off. "qty7" = the last 7 days; "qtyPrev" = the days before that.
function _topKey(barcode, name) {
  const t = String(barcode || '').trim().replace(/^0+/, '');
  return t ? t : 'n:' + String(name || '').trim().toLowerCase();
}
function _getTopSellers(days) {
  days = Math.max(8, Math.min(90, days || 30));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const map = {};
  const now = new Date();
  for (let i = 0; i < days; i++) {
    const d = new Date(now.getTime() - i * 86400000);
    const sh = ss.getSheetByName('Sales_' + Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd'));
    if (!sh) continue;
    const last = sh.getLastRow();
    if (last < 2) continue;
    const rows = sh.getRange(2, 1, last - 1, 9).getValues();
    rows.forEach(r => {
      const type = r[2];
      if (type === 'Voided') return;
      const qty = Number(r[6]) || 0;
      if (!qty) return;
      const k = _topKey(r[3], r[4]);
      if (!map[k]) map[k] = { barcode: String(r[3] || ''), name: String(r[4] || ''), qty30: 0, rev30: 0, qty7: 0, qtyPrev: 0 };
      const it = map[k];
      it.qty30 += qty;
      it.rev30 += Number(r[8]) || 0;
      if (i < 7) it.qty7 += qty; else it.qtyPrev += qty;
    });
  }
  const items = Object.keys(map).map(k => map[k]).filter(it => it.qty30 > 0);
  items.forEach(it => { it.days = days; });
  return _json({ days: days, items: items, generatedAt: new Date().toISOString() });
}

// ══════════════════════════════════════════════════════════
//  TRANSFERS (stock actually moves, in two steps)
//  Sent  → stock leaves the "from" location right away (status 'transit').
//  Received → stock is added at the "to" location (the quantity that really
//  arrived). Cancelled while in transit → the stock goes back.
//  Each movement is applied once (flags OutApplied / InApplied), by calling
//  the Store Inventory script's applyStockMovements — same route used when a
//  sale is voided. Anything that fails stays pending and is retried on the
//  next save or with "Relancer le stock" in Admin.
// ══════════════════════════════════════════════════════════
const TRANSFER_HEADERS = ['UID','Updated','ID','Status','From','To','Barcode','Product','Qty','QtyReceived','By','ReceivedBy','Date','JSON','OutApplied','InApplied'];
function _transfersSheet() { return _sheet('Transfers', TRANSFER_HEADERS); }

// 'ok' | 'nf' (barcode not in that location's tab, or no barcode) | 'err' (try again later)
function _stockMove(location, barcode, delta, type, note) {
  const bc = String(barcode || '').trim();
  if (!bc || bc === '(manual)') return 'nf';
  try {
    const res = UrlFetchApp.fetch(CONFIG.INVENTORY_SYNC_URL, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ key: CONFIG.INVENTORY_SYNC_KEY, action: 'applyStockMovements', location: location,
        movements: [{ barcode: bc, delta: delta, type: type, note: note || '' }] })
    });
    const data = JSON.parse(res.getContentText());
    if (!data.success) { _logSyncIssue(location, 'Transfert : échec pour ' + bc + ' : ' + (data.error || 'erreur inconnue')); return 'err'; }
    if (data.notFound && data.notFound.length) { _logSyncIssue(location, 'Transfert : code ' + bc + ' introuvable dans ' + location + ' — stock à ajuster à la main'); return 'nf'; }
    return 'ok';
  } catch (e) {
    _logSyncIssue(location, 'Transfert : impossible de contacter le Store Inventory Sheet : ' + e);
    return 'err';
  }
}

// Applies whatever movements are still owed for one Transfers row (1-based sheet row).
function _applyTransferRow(sh, rowNum) {
  const r = sh.getRange(rowNum, 1, 1, TRANSFER_HEADERS.length).getValues()[0];
  const status = String(r[3]), from = r[4], to = r[5], bc = r[6], qty = Number(r[8]) || 0;
  const qtyRecv = (r[9] === '' || r[9] === null) ? qty : (Number(r[9]) || 0);
  let out = String(r[14] || ''), inn = String(r[15] || '');
  const label = 'Transfert ' + (r[2] || r[0]);
  if ((status === 'transit' || status === 'received') && !out && qty > 0) {
    const m = _stockMove(from, bc, -qty, 'transfer-out', label + ' → ' + to);
    if (m !== 'err') out = (m === 'ok') ? 'yes' : 'nf';
  }
  if (status === 'received' && !inn && qtyRecv > 0) {
    const m = _stockMove(to, bc, qtyRecv, 'transfer-in', label + ' ← ' + from);
    if (m !== 'err') inn = (m === 'ok') ? 'yes' : 'nf';
  }
  if (status === 'cancelled' && out) {
    if (out === 'yes') {
      const m = _stockMove(from, bc, qty, 'transfer-cancel', label + ' annulé');
      if (m !== 'err') out = 'returned';
    } else if (out === 'nf') out = 'cancelled';
  } else if (status === 'cancelled' && !out) { out = 'cancelled'; }
  sh.getRange(rowNum, 15, 1, 2).setValues([[out, inn]]);
}

function _saveTransfer(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = _transfersSheet();
    const rows = sh.getDataRange().getValues();
    const upd = Number(d.upd) || 0;
    let rowNum = -1;
    for (let i = 1; i < rows.length; i++) { if (rows[i][0] === d.uid) { rowNum = i + 1; break; } }
    const vals = [d.uid, upd, d.id || '', d.status || '', d.from || '', d.to || '', String(d.barcode || ''), d.name || '',
      Number(d.qty) || 0, (d.qtyReceived === null || d.qtyReceived === undefined || d.qtyReceived === '') ? '' : Number(d.qtyReceived),
      d.by || '', d.receivedBy || '', d.date || '', d.json || ''];
    if (rowNum < 0) {
      rowNum = sh.getLastRow() + 1;
      const rng = sh.getRange(rowNum, 1, 1, TRANSFER_HEADERS.length);
      rng.setNumberFormat('@');
      rng.setValues([vals.concat(['', ''])]);
    } else if ((Number(rows[rowNum - 1][1]) || 0) < upd) {
      sh.getRange(rowNum, 1, 1, 14).setValues([vals]);
    }
    _applyTransferRow(sh, rowNum);
    _sweepPendingTransfers(sh, rowNum);
    return _json({ ok: true });
  } finally { lock.releaseLock(); }
}

// Retries other rows whose stock movement failed earlier (bounded, so a save stays fast).
function _sweepPendingTransfers(sh, skipRow) {
  const rows = sh.getDataRange().getValues();
  let n = 0;
  for (let i = rows.length - 1; i >= 1 && n < 5; i--) {
    if (i + 1 === skipRow) continue;
    const r = rows[i], st = String(r[3]), out = String(r[14] || ''), inn = String(r[15] || '');
    const pending = ((st === 'transit' || st === 'received') && !out) || (st === 'received' && !inn) || (st === 'cancelled' && out === 'yes');
    if (pending) { _applyTransferRow(sh, i + 1); n++; }
  }
}

function _retryTransfers() {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const sh = _transfersSheet();
    const rows = sh.getDataRange().getValues();
    let n = 0;
    for (let i = 1; i < rows.length && n < 40; i++) {
      const r = rows[i], st = String(r[3]), out = String(r[14] || ''), inn = String(r[15] || '');
      const pending = ((st === 'transit' || st === 'received') && !out) || (st === 'received' && !inn) || (st === 'cancelled' && out === 'yes');
      if (pending) { _applyTransferRow(sh, i + 1); n++; }
    }
    return _json({ ok: true, retried: n });
  } finally { lock.releaseLock(); }
}

function _getTransfers() {
  const sh = _transfersSheet();
  const rows = sh.getDataRange().getValues().slice(1);
  const out = [];
  rows.slice(-400).forEach(r => {
    try {
      const o = r[13] ? JSON.parse(r[13]) : {};
      o.uid = r[0]; o.upd = Number(r[1]) || 0; o.status = r[3]; o.id = r[2] || o.id;
      o.from = r[4]; o.to = r[5]; o.barcode = String(r[6] || ''); o.name = r[7]; o.qty = Number(r[8]) || 0;
      o.qtyReceived = (r[9] === '' ? null : Number(r[9])); o.by = r[10]; o.receivedBy = r[11];
      o.outApplied = String(r[14] || ''); o.inApplied = String(r[15] || '');
      out.push(o);
    } catch (e) {}
  });
  return _json({ transfers: out });
}

// ══════════════════════════════════════════════════════════
//  ACTIVITY LOG + RESTORE
//  "ActivityLog" tab: who did what on which device (admin sign-ins, deletions,
//  app started with empty storage, failed local saves…). Written by the Scanner.
//  getRestoreData: lets a device rebuild its local sales history for the last
//  N days from SalesLog / Returns / the day tabs (the Sheet is the master copy).
// ══════════════════════════════════════════════════════════
const ACTIVITY_HEADERS = ['Timestamp', 'Store', 'Device', 'Employee', 'Type', 'Detail'];
function _logActivity(d) {
  const sh = _sheet('ActivityLog', ACTIVITY_HEADERS);
  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  sh.appendRow([d.ts ? String(d.ts) : stamp, d.store || '', d.device || '', d.employee || '', d.type || '', String(d.detail || '').slice(0, 500)]);
  return _json({ ok: true });
}
function _getActivityLog(limit) {
  const sh = _sheet('ActivityLog', ACTIVITY_HEADERS);
  const last = sh.getLastRow();
  if (last < 2) return _json({ entries: [] });
  const n = Math.min(Number(limit) || 150, last - 1);
  const rows = sh.getRange(last - n + 1, 1, n, 6).getValues();
  const tz = CONFIG.TIMEZONE;
  return _json({ entries: rows.map(r => ({
    ts: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'dd/MM/yyyy HH:mm:ss') : String(r[0]),
    store: r[1], device: r[2], employee: r[3], type: r[4], detail: r[5] })).reverse() });
}
function _timeStr(v) {
  if (v instanceof Date) return Utilities.formatDate(v, CONFIG.TIMEZONE, 'HH:mm:ss');
  return String(v == null ? '' : v);
}
function _frToDate(s) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(s || '').trim());
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}
function _getRestoreData(store, days) {
  days = Math.max(1, Math.min(60, Number(days) || 7));
  store = String(store || '');
  const tz = CONFIG.TIMEZONE;
  const cutoff = new Date(); cutoff.setHours(0, 0, 0, 0); cutoff.setDate(cutoff.getDate() - (days - 1));
  const salesSh = _sheet(SHEETS.SALES, ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  const retSh = _sheet(SHEETS.RETURNS, ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);
  const sales = [], returns = [], scans = [];
  const sl = salesSh.getLastRow();
  if (sl >= 2) salesSh.getRange(2, 1, sl - 1, 14).getValues().forEach(r => {
    if (!r[0] || String(r[4]) !== store) return;
    const ds = _frDateStr(r[1], tz), d = _frToDate(ds);
    if (!d || d < cutoff) return;
    sales.push({ id: String(r[0]), date: ds, time: _timeStr(r[2]), employee: r[3], itemCount: Number(r[5]) || 0, total: Number(r[6]) || 0,
      payMethod: r[7], payMethodLabel: r[8], splitPayments: r[9], isDelivery: r[10] === true || String(r[10]).toLowerCase() === 'true',
      deliveryId: r[11], customerName: r[12], items: r[13] });
  });
  const rl = retSh.getLastRow();
  if (rl >= 2) retSh.getRange(2, 1, rl - 1, 11).getValues().forEach(r => {
    if (!r[0] || String(r[5]) !== store) return;
    const ds = _frDateStr(r[2], tz), d = _frToDate(ds);
    if (!d || d < cutoff) return;
    returns.push({ id: String(r[0]), txnId: String(r[1]), date: ds, time: _timeStr(r[3]), employee: r[4], refundTotal: Number(r[6]) || 0,
      refundMethod: r[7], refundMethodLabel: r[8], reason: r[9], items: r[10] });
  });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let i = 0; i < days; i++) {
    const day = new Date(); day.setDate(day.getDate() - i);
    const sh = ss.getSheetByName('Sales_' + Utilities.formatDate(day, tz, 'yyyy-MM-dd'));
    if (!sh) continue;
    const last = sh.getLastRow();
    if (last < 2) continue;
    sh.getRange(2, 1, last - 1, 17).getValues().forEach(r => {
      if (r[2] !== 'Sale' || String(r[12]) !== store) return;
      scans.push({ date: _frDateStr(r[0], tz), time: _timeStr(r[1]), barcode: String(r[3] || ''), name: r[4], category: r[5],
        qty: Number(r[6]) || 1, price: Number(r[7]) || 0, lineTotal: Number(r[8]) || 0, employee: r[11], txnId: String(r[13] || ''), scanId: String(r[16] || '') });
    });
  }
  return _json({ store: store, days: days, sales: sales, returns: returns, scans: scans });
}

function _deliveriesSheet() {
  return _sheet('Livraisons', ['UID','Updated','DlvID','Status','Name','Phone','Store','JSON']);
}

// Upsert by UID; a stale copy (older Updated) never overwrites a newer one.
function _saveDelivery(d) {
  const sh = _deliveriesSheet();
  const rows = sh.getDataRange().getValues();
  const row = [d.uid, Number(d.upd) || 0, d.dlvId || '', d.status || '', d.name || '', String(d.phone || ''), d.store || '', d.json || ''];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.uid) {
      if ((Number(rows[i][1]) || 0) >= row[1]) return _json({ ok: true, stale: true });
      sh.getRange(i + 1, 1, 1, 8).setValues([row]);
      return _json({ ok: true });
    }
  }
  sh.appendRow(row);
  return _json({ ok: true });
}

function _getDeliveries() {
  const rows = _deliveriesSheet().getDataRange().getValues().slice(1);
  const out = [];
  rows.forEach(r => { try { if (r[7]) out.push(JSON.parse(r[7])); } catch (e) {} });
  return _json({ deliveries: out });
}

function _getCustomers() {
  const sh = _customersSheet();
  const rows = sh.getDataRange().getValues().slice(1);
  const customers = rows.filter(r => r[0]).map(r => {
    let notes = [];
    try { notes = r[4] ? JSON.parse(r[4]) : []; } catch (e) { notes = []; }
    return { phone: r[0], name: r[1], store: r[2], lastPurchase: r[3], notes: Array.isArray(notes) ? notes : [] };
  });
  return _json({ customers });
}

function _logReturn(d) {
  const sh = _sheet(SHEETS.RETURNS,
    ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);
  sh.appendRow([d.returnId, d.txnId, d.date, d.time, d.employee, d.store,
    d.refundTotal, d.refundMethod, d.refundMethodLabel, d.reason || '', d.items || '']);
  _logReturnItems(d);
  _refreshDashboard();
  return _json({ ok: true });
}

// Mirrors each returned item into the same per-day feed as the original sale,
// as a negative-Qty row (barcode-keyed, same columns), so the inventory sync
// can apply sales and returns with identical stock -= Qty logic from one feed.
function _logReturnItems(d) {
  try {
    let items = [];
    try { items = d.itemsJson ? JSON.parse(d.itemsJson) : []; } catch (e2) { items = []; }
    if (!items.length) return; // older client without itemsJson — Returns tab above still has the text summary
    const sh = _dayTab(d.date);
    const label = d.refundMethodLabel || PAY_LABELS[d.refundMethod] || d.refundMethod || '';
    items.forEach(it => {
      sh.appendRow([d.date, d.time, 'Return', it.barcode || '', it.name || '', it.category || '',
        -Math.abs(it.qty || 0), it.unitPrice || 0, -Math.abs(it.lineTotal || 0),
        d.refundMethod || '', label, d.employee, d.store, d.txnId, d.returnId, '', '']);
    });
  } catch (e) { Logger.log('Return item mirror skipped: ' + e); }
}

// ══════════════════════════════════════════════════════════
//  VOID / CANCEL — called from the scanner when an item is removed from
//  the cart, or the whole in-progress sale is cancelled, BEFORE checkout.
// ══════════════════════════════════════════════════════════
// A scanned item is logged into the day-tab immediately (_logScan, above)
// so nothing is lost if the till crashes mid-sale. That means a scan the
// cashier later removes — wrong item, customer changed their mind, sale
// abandoned entirely — is still sitting in the day-tab. Left alone, the
// next syncInventoryFromSales() run would treat it as a real sale and
// wrongly decrement Store Inventory.
//
// d = { key, action: 'voidScan', date, scanId }
//
// Two cases, found by looking at that row's own "Processed" column:
//   - Still blank (syncInventoryFromSales hasn't looked at it yet): just
//     zero the row out in place. It keeps its place in the day-tab as a
//     visible "Voided" record, but a zero Qty means the sync loop's own
//     `!qty` check (see syncInventoryFromSales above) skips it forever —
//     nothing further to do.
//   - Already has a timestamp (stock was already decremented): zeroing
//     the row alone isn't enough — Store Inventory's Qty On Hand already
//     reflects the sale. Send one compensating movement through the same
//     "applyStockMovements" action syncInventoryFromSales itself uses, to
//     put that unit's stock back, then zero the row for a clean record.
function _voidScan(d) {
  try {
    if (!d.scanId) return _json({ error: 'Missing scanId' });
    const sh = _dayTab(d.date);
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return _json({ ok: true, note: 'day-tab empty' });
    const numRows = lastRow - 1;
    const values = sh.getRange(2, 1, numRows, SALES_ITEM_HEADERS.length).getValues();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (row[16] !== d.scanId) continue; // col 17 = ScanID
      const sheetRow = i + 2;
      const processed = row[15];   // Processed
      const barcode = row[3];
      const qty = Number(row[6]) || 0;
      const type = row[2];         // 'Sale' | 'Return'
      const storeRaw = row[12];

      if (processed && qty && barcode) {
        // Stock was already moved by syncInventoryFromSales — reverse it.
        const location = STORE_NAME_MAP[storeRaw] || storeRaw;
        const delta = type === 'Return' ? -Math.abs(qty) : Math.abs(qty);
        try {
          const res = UrlFetchApp.fetch(CONFIG.INVENTORY_SYNC_URL, {
            method: 'post',
            contentType: 'application/json',
            payload: JSON.stringify({
              key: CONFIG.INVENTORY_SYNC_KEY,
              action: 'applyStockMovements',
              location: location,
              movements: [{ barcode: String(barcode), delta: delta, type: 'void', note: 'Annulation de ' + (row[13] || '') }],
            }),
            muteHttpExceptions: true,
          });
          const data = JSON.parse(res.getContentText());
          if (!data.success) {
            _logSyncIssue(location, 'Échec de l\'annulation pour ' + barcode + ' : ' + (data.error || 'erreur inconnue'));
          }
        } catch (e) {
          _logSyncIssue(storeRaw || '(inconnu)', 'Impossible de contacter le Store Inventory Sheet pour annuler ' + barcode + ' : ' + e);
        }
      }

      // Zero the row out either way, and mark it Voided so the day-tab
      // still shows what happened instead of silently vanishing.
      sh.getRange(sheetRow, 3).setValue('Voided');  // Type
      sh.getRange(sheetRow, 7).setValue(0);          // Qty
      sh.getRange(sheetRow, 9).setValue(0);          // LineTotal
      sh.getRange(sheetRow, 18).setValue('');        // Remise
      try { sh.hideRows(sheetRow); } catch (e) {}    // hidden by default; menu "Boutique 1981" shows them
      if (!processed) {
        sh.getRange(sheetRow, 16).setValue(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm')); // Processed
      }
      return _json({ ok: true, wasAlreadySynced: !!processed });
    }
    return _json({ ok: true, note: 'scanId not found (already voided or never logged)' });
  } catch (e) {
    return _json({ error: String(e) });
  }
}

// ══════════════════════════════════════════════════════════
//  FINALIZE ORPHANED SCANS — admin recovery after a power outage or app
//  crash mid-sale. See Boutique_Scanner.html's "⚡ Ventes interrompues"
//  admin card: a scan is written to the day-tab feed the instant it
//  happens (_logScan above), independent of whether the sale containing
//  it ever got closed. The in-progress cart only ever lives in the
//  device's memory (never saved to disk) — if power cuts before "Close
//  Sale", that grouping is gone forever, but the day-tab rows already
//  sent stay there, with no sale ever attached: never in SalesLog (so
//  invisible to the Dashboard), even though the 15-minute inventory
//  sync already deducted their stock regardless of any of this.
//
// Creates exactly one new SalesLog row for the admin's chosen group of
// orphaned scans — dated to when they actually happened, not today —
// and stamps those specific day-tab rows (found by ScanID, same lookup
// _voidScan uses above) with the payment method and this new TxnID.
// Deliberately does NOT call the usual _backfillPayMethod: that matches
// rows by TxnID, and these rows still carry their ORIGINAL abandoned
// TxnID, not the new consolidated one, so it would find nothing.
// Never touches Store Inventory stock — that was already correctly
// applied once, at scan time; this only fixes the revenue record.
//
// d = { key, action: 'finalizeOrphanScans', txnId, date, time, employee,
//       store, itemCount, total, items, payMethod, payMethodLabel, scanIds }
function _finalizeOrphanScans(d) {
  try {
    if (!d.txnId || !Array.isArray(d.scanIds) || !d.scanIds.length) {
      return _json({ error: 'Missing txnId or scanIds' });
    }
    const sh = _sheet(SHEETS.SALES,
      ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
    sh.appendRow([d.txnId, d.date, d.time, d.employee, d.store, d.itemCount, d.total,
      d.payMethod, d.payMethodLabel || d.payMethod, '', false, '', '', d.items || '']);

    let matched = 0;
    try {
      const dayTab = _dayTab(d.date);
      const lastRow = dayTab.getLastRow();
      if (lastRow >= 2) {
        const wanted = {};
        d.scanIds.forEach(id => { if (id) wanted[id] = true; });
        const numRows = lastRow - 1;
        const rows = dayTab.getRange(2, 1, numRows, SALES_ITEM_HEADERS.length).getValues();
        for (let i = 0; i < rows.length; i++) {
          const scanId = rows[i][16]; // col 17 = ScanID
          if (!scanId || !wanted[scanId]) continue;
          const sheetRow = i + 2;
          dayTab.getRange(sheetRow, 10).setValue(d.payMethod);                      // PayMethod
          dayTab.getRange(sheetRow, 11).setValue(d.payMethodLabel || d.payMethod);  // PayMethodLabel
          dayTab.getRange(sheetRow, 14).setValue(d.txnId);                           // TxnID
          dayTab.getRange(sheetRow, 15).setValue(d.txnId);                           // RefID
          matched++;
        }
      }
    } catch (e) { Logger.log('Orphan scan row stamping skipped: ' + e); }

    _logSaleEdit(d.txnId, 'Vente récupérée (coupure)', d.scanIds.length + ' scan(s) sélectionné(s)', matched + ' ligne(s) rattachée(s)');
    _refreshDashboard();
    return _json({ ok: true, matched: matched, requested: d.scanIds.length });
  } catch (e) {
    return _json({ error: String(e) });
  }
}

// ══════════════════════════════════════════════════════════
//  MENSUALITÉ (installment plans) — layaway-style customer credit ledger
// ══════════════════════════════════════════════════════════
// InstallmentPlans columns (1-indexed for getRange, 0-indexed for arrays):
//  0 PlanID, 1 TxnID, 2 Date, 3 Time, 4 Phone, 5 CustomerName, 6 Employee,
//  7 Store, 8 ItemsText, 9 ItemsJSON, 10 TotalDue, 11 AmountPaid, 12 Balance,
//  13 LastPaymentDate, 14 Status
//  14 Status, 15 NextDue (dd/mm/yyyy text; added later — older rows have it blank)
const INST_PLAN_HEADERS = ['PlanID','TxnID','Date','Time','Phone','CustomerName','Employee','Store',
  'ItemsText','ItemsJSON','TotalDue','AmountPaid','Balance','LastPaymentDate','Status','NextDue'];
// InstallmentPayments columns:
//  0 PlanID, 1 Phone, 2 Date, 3 Time, 4 Employee, 5 Amount, 6 Method, 7 MethodLabel, 8 BalanceAfter, 9 Status
const INST_PAYMENT_HEADERS = ['PlanID','Phone','Date','Time','Employee','Amount','Method','MethodLabel','BalanceAfter','Status'];

// Same as _sheet() for the plans tab, but also adds the NextDue column to a
// sheet created before that column existed (header + plain-text format).
function _instPlansSheet() {
  const sh = _sheet(SHEETS.INST_PLANS, INST_PLAN_HEADERS);
  if (String(sh.getRange(1, 16).getValue() || '') !== 'NextDue') {
    sh.getRange(1, 16).setValue('NextDue');
    sh.getRange(1, 16, Math.max(sh.getMaxRows(), 2000), 1).setNumberFormat('@');
    sh.getRange(1, 16).setNumberFormat('General');
  }
  return sh;
}

function _normPhone(p) {
  return String(p || '').replace(/[\s\-().+]/g, '').trim();
}

function _createInstallmentPlan(d) {
  const sh = _instPlansSheet();
  sh.appendRow([d.planId, d.txnId, d.date, d.time, d.phone, d.customerName || '', d.employee || '', d.store || '',
    d.items || '', d.itemsJson || '', d.totalDue || 0, d.amountPaid || 0, d.balance || 0,
    d.lastPaymentDate || '', d.status || 'open', d.nextDue || '']);
  _refreshDashboard();
  return _json({ ok: true });
}

function _logInstallmentPayment(d) {
  const paySh = _sheet(SHEETS.INST_PAYMENTS, INST_PAYMENT_HEADERS);
  paySh.appendRow([d.planId, d.phone, d.date, d.time, d.employee || '', d.amount || 0,
    d.method || '', d.methodLabel || '', d.balanceAfter || 0, d.status || '']);

  // Also update the running totals on the plan's own row, so the Sheet is
  // browsable on its own without cross-referencing the payments tab.
  const plansSh = _instPlansSheet();
  const rows = plansSh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.planId) {
      const newAmountPaid = (Number(rows[i][11]) || 0) + (d.amount || 0);
      plansSh.getRange(i + 1, 12).setValue(newAmountPaid);        // AmountPaid
      plansSh.getRange(i + 1, 13).setValue(d.balanceAfter || 0);  // Balance
      plansSh.getRange(i + 1, 14).setValue(d.date || '');         // LastPaymentDate
      plansSh.getRange(i + 1, 15).setValue(d.status || '');       // Status
      if (d.nextDue !== undefined) plansSh.getRange(i + 1, 16).setValue(d.nextDue || ''); // NextDue
      break;
    }
  }
  _refreshDashboard();
  return _json({ ok: true });
}

// Admin ticks an item "Récupéré" and/or changes the next due date. Sends the
// full ItemsJSON (with each item's pickedUp flag) so the call is idempotent.
function _updatePlanMeta(d) {
  const sh = _instPlansSheet();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.planId) {
      if (d.itemsJson !== undefined && d.itemsJson !== '') sh.getRange(i + 1, 10).setValue(d.itemsJson);
      if (d.nextDue !== undefined) sh.getRange(i + 1, 16).setValue(d.nextDue || '');
      return _json({ ok: true });
    }
  }
  return _json({ error: 'Plan not found: ' + d.planId });
}

// Every plan, with parsed items (incl. pickedUp), next due date and payments —
// for the Mensualités dashboard in the scanner's Admin screen. Gated by the
// normal WRITE_KEY in doGet (the scanner itself is behind the admin PIN).
function _getPlansDashboard() {
  const plansSh = _instPlansSheet();
  const paySh = _sheet(SHEETS.INST_PAYMENTS, INST_PAYMENT_HEADERS);
  const planRows = plansSh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const payRows = paySh.getDataRange().getValues().slice(1);
  const plans = planRows.map(r => {
    let items = [];
    try { items = r[9] ? JSON.parse(r[9]) : []; } catch (e) { items = []; }
    const payments = payRows.filter(p => p[0] === r[0]).map(p => ({
      date: p[2], time: p[3], employee: p[4], amount: Number(p[5]) || 0, method: p[6], methodLabel: p[7],
      balanceAfter: Number(p[8]) || 0,
    }));
    return {
      id: r[0], txnId: r[1], date: r[2], time: r[3], phone: r[4], customerName: r[5], employee: r[6], store: r[7],
      items, totalDue: Number(r[10]) || 0, amountPaid: Number(r[11]) || 0, balance: Number(r[12]) || 0,
      lastPaymentDate: r[13], status: r[14], nextDue: r[15] ? String(r[15]) : '', payments,
    };
  });
  return _json({ plans });
}

// Cross-device fallback: the scanner checks this if a phone number isn't
// found in its own local storage (e.g. the plan was created on another
// device, or the browser's local storage was cleared).
function _getInstallmentPlans(phone) {
  const normTarget = _normPhone(phone);
  const plansSh = _instPlansSheet();
  const paySh = _sheet(SHEETS.INST_PAYMENTS, INST_PAYMENT_HEADERS);
  const planRows = plansSh.getDataRange().getValues().slice(1)
    .filter(r => r[0] && _normPhone(r[4]) === normTarget);
  const payRows = paySh.getDataRange().getValues().slice(1);

  const plans = planRows.map(r => {
    let items = [];
    try { items = r[9] ? JSON.parse(r[9]) : []; } catch (e) { items = []; }
    const payments = payRows
      .filter(p => p[0] === r[0])
      .map(p => ({ date: p[2], time: p[3], employee: p[4], amount: Number(p[5]) || 0,
        method: p[6], methodLabel: p[7], balanceAfter: Number(p[8]) || 0 }));
    return {
      id: r[0], txnId: r[1], date: r[2], time: r[3], phone: r[4], customerName: r[5], employee: r[6],
      items, totalDue: Number(r[10]) || 0, amountPaid: Number(r[11]) || 0, balance: Number(r[12]) || 0,
      lastPaymentDate: r[13], status: r[14], nextDue: r[15] ? String(r[15]) : '', payments,
    };
  });
  return _json({ plans });
}

// Remote-admin-only variant of the above — ALL plans (every phone number),
// for the "Mensualités" view in Boutique_Admin_Remote.html. Gated by
// REMOTE_ADMIN_PASSPHRASE in doGet, unlike _getInstallmentPlans (phone
// lookup) which the scanner itself uses and needs no extra secret for.
function _getAllInstallmentPlans() {
  const plansSh = _instPlansSheet();
  const paySh = _sheet(SHEETS.INST_PAYMENTS, INST_PAYMENT_HEADERS);
  const planRows = plansSh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const payRows = paySh.getDataRange().getValues().slice(1);
  const plans = planRows.map(r => {
    const payments = payRows.filter(p => p[0] === r[0]).map(p => ({
      date: p[2], time: p[3], employee: p[4], amount: Number(p[5]) || 0, method: p[6], methodLabel: p[7],
    }));
    return {
      id: r[0], txnId: r[1], date: r[2], time: r[3], phone: r[4], customerName: r[5], employee: r[6], store: r[7],
      itemsText: r[8], totalDue: Number(r[10]) || 0, amountPaid: Number(r[11]) || 0, balance: Number(r[12]) || 0,
      lastPaymentDate: r[13], status: r[14], payments,
    };
  });
  return _json({ plans });
}

// ══════════════════════════════════════════════════════════
//  HOLDS / RÉSERVATIONS
// ══════════════════════════════════════════════════════════
// Columns: 0 HoldID, 1 Phone, 2 CustomerName, 3 Employee, 4 Store, 5 Date,
// 6 Time, 7 ExpiryDate, 8 ItemsText, 9 ItemsJSON, 10 Total, 11 Status
const HOLD_HEADERS = ['HoldID','Phone','CustomerName','Employee','Store','Date','Time',
  'ExpiryDate','ItemsText','ItemsJSON','Total','Status'];

// The scanner reuses the single 'createHold' action both when a hold is
// first created and whenever its status changes (confirmed/cancelled) —
// this upserts by HoldID rather than needing a separate update action.
function _createOrUpdateHold(d) {
  const sh = _sheet(SHEETS.HOLDS, HOLD_HEADERS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.holdId) {
      sh.getRange(i + 1, 12).setValue(d.status || 'active'); // Status
      _refreshDashboard();
      return _json({ ok: true });
    }
  }
  sh.appendRow([d.holdId, d.phone, d.customerName || '', d.employee || '', d.store || '',
    d.date, d.time, d.expiryDate, d.items || '', d.itemsJson || '', d.total || 0, d.status || 'active']);
  _refreshDashboard();
  return _json({ ok: true });
}

// ══════════════════════════════════════════════════════════
//  DISCOUNTS — audit trail of every discount applied at the POS, whole-sale
//  or per-item, with who authorized it (PIN-matched on the scanner side
//  against the employees list, or the admin PIN as a fallback).
// ══════════════════════════════════════════════════════════
const DISCOUNT_HEADERS = ['Date', 'Time', 'TxnID', 'Scope', 'ItemName',
  'DiscountType', 'DiscountValue', 'AmountOff', 'AppliedBy', 'Employee', 'Store'];

function _logDiscount(d) {
  const sh = _sheet(SHEETS.DISCOUNTS, DISCOUNT_HEADERS);
  sh.appendRow([d.date, d.time, d.txnId, d.scope, d.itemName || '',
    d.discountType, d.discountValue, d.amountOff, d.appliedBy, d.employee || '', d.store || '']);
  return _json({ ok: true });
}

// ══════════════════════════════════════════════════════════
//  PETTY CASH / TILL EXPENSES — logged by any employee from the POS
//  ("💸 Sortie de caisse"), cash only. Recorded here for the admin's
//  records; the POS already deducts these locally from expected cash
//  for shift reconciliation, this sheet is just the durable record.
// ══════════════════════════════════════════════════════════
const EXPENSE_HEADERS = ['Date', 'Time', 'Employee', 'Store', 'Amount', 'Note'];

function _logExpense(d) {
  const sh = _sheet(SHEETS.EXPENSES, EXPENSE_HEADERS);
  sh.appendRow([d.date, d.time, d.employee || '', d.store || '', d.amount || 0, d.note || '']);
  return _json({ ok: true });
}

// ══════════════════════════════════════════════════════════
//  REMOTE ADMIN — backs Boutique_Admin_Remote.html, the single-file page
//  the owner opens on their own computer to view AND manage both stores
//  without visiting either one. Every function below is only reachable
//  through the passphrase-gated routes in doGet/doPost above.
//
//  Settings (store branding, hold duration, admin PIN, employee roster)
//  used to live only in each device's local storage. They're now ALSO
//  centralized here so a change made remotely actually reaches the store
//  computers — Boutique_Scanner.html pulls these down periodically the
//  same way it already pulls the product catalog (cache-and-layer, so it
//  still works offline from the last-synced values).
// ══════════════════════════════════════════════════════════

const STORE_SETTINGS_HEADERS = ['Store', 'Icon', 'Name', 'Tagline', 'Phone', 'Address', 'ClosingMsg'];

function _getStoreSettings() {
  const sh = _sheet(SHEETS.STORE_SETTINGS, STORE_SETTINGS_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const stores = rows.map(r => ({
    store: r[0], icon: r[1], name: r[2], tagline: r[3], phone: r[4], address: r[5], closingMsg: r[6],
  }));
  return _json({ stores });
}

// Upserts one store's branding row, keyed by Store name. Payload: { store, icon, name, tagline, phone, address, closingMsg }.
function _saveStoreSettings(d) {
  const sh = _sheet(SHEETS.STORE_SETTINGS, STORE_SETTINGS_HEADERS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === d.store) {
      sh.getRange(i + 1, 1, 1, 7).setValues([[d.store, d.icon || '', d.name || '', d.tagline || '',
        d.phone || '', d.address || '', d.closingMsg || '']]);
      return _json({ ok: true });
    }
  }
  sh.appendRow([d.store, d.icon || '', d.name || '', d.tagline || '', d.phone || '', d.address || '', d.closingMsg || '']);
  return _json({ ok: true });
}

// Single-row settings shared across both stores: admin PIN, hold duration,
// and the EmailJS config used by the (separate, client-side) email-receipt
// feature. Always row 2 (row 1 is headers) — created on first save/read.
const GLOBAL_SETTINGS_HEADERS = ['AdminPin', 'HoldDurationDays', 'EmailJSService', 'EmailJSTemplate', 'EmailJSKey', 'EmailJSEmail', 'PricePin'];

function _getGlobalSettings() {
  const sh = _sheet(SHEETS.GLOBAL_SETTINGS, GLOBAL_SETTINGS_HEADERS);
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return _json({ settings: null });
  const r = rows[1];
  return _json({ settings: {
    adminPin: r[0], pricePin: String(r[6] || ''), holdDurationDays: Number(r[1]) || 3,
    emailJs: { service: r[2], template: r[3], key: r[4], email: r[5] },
  }});
}

// Narrow writers so one device can change a single PIN without touching the
// rest of the shared settings row.
function _saveSettingsCell(col, value) {
  const sh = _sheet(SHEETS.GLOBAL_SETTINGS, GLOBAL_SETTINGS_HEADERS);
  if (String(sh.getRange(1, 7).getValue() || '') !== 'PricePin') sh.getRange(1, 7).setValue('PricePin');
  if (sh.getLastRow() < 2) sh.appendRow(['1234', 3, '', '', '', '', '']);
  const c = sh.getRange(2, col);
  c.setNumberFormat('@');
  c.setValue(String(value));
}
function _savePricePin(d) { _saveSettingsCell(7, d.pricePin || ''); return _json({ ok: true }); }
function _saveAdminPin(d) {
  if (!/^\d{4,}$/.test(String(d.adminPin || ''))) return _json({ error: 'PIN invalide' });
  _saveSettingsCell(1, d.adminPin);
  return _json({ ok: true });
}

function _saveGlobalSettings(d) {
  const sh = _sheet(SHEETS.GLOBAL_SETTINGS, GLOBAL_SETTINGS_HEADERS);
  const row = [d.adminPin || '1234', d.holdDurationDays || 3,
    (d.emailJs && d.emailJs.service) || '', (d.emailJs && d.emailJs.template) || '',
    (d.emailJs && d.emailJs.key) || '', (d.emailJs && d.emailJs.email) || ''];
  if (sh.getLastRow() < 2) sh.appendRow(row);
  else sh.getRange(2, 1, 1, row.length).setValues([row]);
  return _json({ ok: true });
}

// Employees are a single shared roster (not split per store), matching how
// the scanner already stores them today. saveEmployees REPLACES the whole
// list each time — simplest to reason about, and the remote page always
// sends the full current list back.
const EMPLOYEE_HEADERS = ['Name', 'Pin'];

function _getEmployees() {
  const sh = _sheet(SHEETS.EMPLOYEES, EMPLOYEE_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  return _json({ employees: rows.map(r => ({ name: r[0], pin: String(r[1]) })) });
}

function _saveEmployees(d) {
  const sh = _sheet(SHEETS.EMPLOYEES, EMPLOYEE_HEADERS);
  const list = Array.isArray(d.employees) ? d.employees : [];
  sh.getDataRange().clear();
  sh.appendRow(EMPLOYEE_HEADERS);
  sh.setFrozenRows(1);
  _forceTextColumns(sh, TEXT_COLUMNS['Employees']);
  list.forEach(e => sh.appendRow([e.name, e.pin]));
  return _json({ ok: true });
}

// Recent shift-close discrepancies — mirrors the filter the POS app's own
// Admin panel already applies locally (non-zero cash or Orange discrepancy),
// just read from the Sheet instead of one device's local storage.
function _getReconciliation() {
  const sh = _sheet(SHEETS.RECONCILIATION, RECONCILIATION_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1);
  const flagged = rows
    .map(r => ({
      date: r[0], time: r[1], employee: r[2], store: r[3], clockIn: r[4], clockOut: r[5],
      startingCash: Number(r[6]) || 0, endingCash: Number(r[7]) || 0, expectedCash: Number(r[8]) || 0,
      cashDiscrepancy: Number(r[9]) || 0, startingOrange: Number(r[10]) || 0, endingOrange: Number(r[11]) || 0,
      expectedOrange: Number(r[12]) || 0, orangeDiscrepancy: Number(r[13]) || 0,
    }))
    .filter(r => r.cashDiscrepancy !== 0 || r.orangeDiscrepancy !== 0)
    .slice(-60).reverse();
  return _json({ reconciliations: flagged });
}

function _getExpenses() {
  const sh = _sheet(SHEETS.EXPENSES, EXPENSE_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const expenses = rows.map(r => ({
    date: r[0], time: r[1], employee: r[2], store: r[3], amount: Number(r[4]) || 0, note: r[5],
  })).slice(-100).reverse();
  return _json({ expenses });
}

function _getHoldsRemote() {
  const sh = _sheet(SHEETS.HOLDS, HOLD_HEADERS);
  const rows = sh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const holds = rows.map(r => ({
    holdId: r[0], phone: r[1], customerName: r[2], employee: r[3], store: r[4],
    date: r[5], time: r[6], expiryDate: r[7], itemsText: r[8], total: Number(r[10]) || 0, status: r[11],
  }));
  return _json({ holds });
}

// Simplified per-day totals (current calendar month) for the "Résumé
// quotidien" view — Total / Cash / Orange / Wave / Bank, all stores
// combined, computed from SalesLog minus Returns. Dates are read as the
// plain-text 'dd/MM/yyyy' strings TEXT_COLUMNS already enforces, so no
// timezone/Date-object ambiguity.
function _getDailySummary(monthParam) {
  const salesSh = _sheet(SHEETS.SALES,
    ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  const returnsSh = _sheet(SHEETS.RETURNS,
    ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);
  const tz = CONFIG.TIMEZONE;
  const now = new Date();
  // monthParam: optional "yyyy-MM" (native <input type="month"> value).
  // Falls back to the current calendar month when absent/malformed, exactly
  // the old always-current-month behavior.
  let curMonth = Utilities.formatDate(now, tz, 'MM/yyyy');
  let labelDate = now;
  if (monthParam && /^\d{4}-\d{2}$/.test(monthParam)) {
    const [y, m] = monthParam.split('-');
    curMonth = m + '/' + y;
    labelDate = new Date(Number(y), Number(m) - 1, 1);
  }
  const days = {}; // 'dd/MM/yyyy' -> {total,cash,orange,wave,bank,monthly}
  const blank = () => ({ total: 0, cash: 0, orange: 0, wave: 0, bank: 0, monthly: 0 });
  const inThisMonth = dateStr => dateStr.slice(3) === curMonth;

  salesSh.getDataRange().getValues().slice(1).forEach(r => {
    const dateStr = _frDateStr(r[1], tz);
    if (!inThisMonth(dateStr)) return;
    if (!days[dateStr]) days[dateStr] = blank();
    const total = Number(r[6]) || 0;
    days[dateStr].total += total;
    const method = r[7];
    if (method === 'split' && r[9]) {
      try {
        JSON.parse(r[9]).forEach(sp => {
          if (days[dateStr][sp.method] !== undefined) days[dateStr][sp.method] += Number(sp.amount) || 0;
        });
      } catch (e) { /* malformed split payload — total still counted above */ }
    } else if (days[dateStr][method] !== undefined) {
      days[dateStr][method] += total;
    }
  });
  returnsSh.getDataRange().getValues().slice(1).forEach(r => {
    const dateStr = _frDateStr(r[2], tz);
    if (!inThisMonth(dateStr)) return;
    if (!days[dateStr]) days[dateStr] = blank();
    const refund = Number(r[6]) || 0;
    days[dateStr].total -= refund;
    const method = r[7];
    if (days[dateStr][method] !== undefined) days[dateStr][method] -= refund;
  });

  const dayList = Object.keys(days).map(dateStr => Object.assign({ date: dateStr }, days[dateStr]));
  // Sort by actual day-of-month (dd/MM/yyyy — first two chars) descending.
  dayList.sort((a, b) => Number(b.date.slice(0, 2)) - Number(a.date.slice(0, 2)));
  const month = dayList.reduce((m, d) => {
    m.total += d.total; m.cash += d.cash; m.orange += d.orange; m.wave += d.wave; m.bank += d.bank; m.monthly += d.monthly;
    return m;
  }, blank());
  return _json({ days: dayList, month, monthLabel: Utilities.formatDate(labelDate, tz, 'MMMM yyyy') });
}

// Every TxnID|Store and ReturnID|Store key currently in the Sheet, so a till
// device can diff its own full local history against this list and find out
// exactly which of its sales/returns — from any date, not just today — never
// made it across, without risking re-sending (and duplicating) ones that
// already did. See runFullResync() in Boutique_Scanner.html.
function _getSyncedKeys() {
  const salesSh = _sheet(SHEETS.SALES,
    ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  const returnsSh = _sheet(SHEETS.RETURNS,
    ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);
  const saleKeys = salesSh.getDataRange().getValues().slice(1)
    .filter(r => r[0]).map(r => r[0] + '|' + (r[4] || ''));
  const returnKeys = returnsSh.getDataRange().getValues().slice(1)
    .filter(r => r[0]).map(r => r[0] + '|' + (r[5] || ''));
  return _json({ saleKeys, returnKeys });
}

// Reads the already-computed Dashboard tab (today's totals + per-store
// breakdown) rather than recomputing from raw rows — same numbers the
// in-store Dashboard tab shows, one function call away.
function _getDashboardData() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!sh) return _json({ error: 'Dashboard tab not found — run setupDashboard() once.' });
  const todayDate = sh.getRange('B4').getValue();
  const salesCount = sh.getRange('B5').getValue();
  const grossRevenue = sh.getRange('B6').getValue();
  const returnsCount = sh.getRange('B7').getValue();
  const netRevenue = sh.getRange('B9').getValue();
  const payBreakdown = {};
  Object.keys(PAY_LABELS).forEach((m, i) => { payBreakdown[m] = sh.getRange(12 + i, 2).getValue(); });
  // Per-store table starts at row 39 in setupDashboard()'s layout; read
  // until an empty Store-name cell is hit.
  const storeRows = [];
  for (let i = 0; i < 10; i++) {
    const name = sh.getRange(40 + i, 1).getValue();
    if (!name) break;
    storeRows.push({
      store: name, sales: sh.getRange(40 + i, 2).getValue(), gross: sh.getRange(40 + i, 3).getValue(),
      returns: sh.getRange(40 + i, 4).getValue(), refunds: sh.getRange(40 + i, 5).getValue(),
      net: sh.getRange(40 + i, 6).getValue(),
    });
  }
  return _json({ today: todayDate, salesCount, grossRevenue, returnsCount, netRevenue, payBreakdown, stores: storeRows });
}

// ══════════════════════════════════════════════════════════
//  INVENTORY SYNC — applies sold/returned items to live "Qty On Hand"
//  on the Store Inventory Sheet, automatically, every 15 minutes.
// ══════════════════════════════════════════════════════════
//
// How it works:
//   1. Every 15 minutes, syncInventoryFromSales() looks at today's and
//      yesterday's Sales_yyyy-MM-dd day-tabs (INVENTORY_SYNC_LOOKBACK_DAYS
//      controls how far back — 2 days by default, which comfortably covers
//      a scan that synced late from an offline device).
//   2. It collects every row whose "Processed" column is still blank — a
//      row it hasn't applied yet.
//   3. It groups those rows by Store (Store 1 / Store 2 / Warehouse — must
//      match those names exactly, see STORE_NAME_MAP below if it doesn't)
//      and sends one batch per store to the Store Inventory Sheet's own
//      script, using the SAME "applyStockMovements" action the nightly
//      Stock Ledger app already uses. A sale (positive Qty) becomes a
//      negative delta (stock down); a return (negative Qty) becomes a
//      positive delta (stock up) — same "stock -= Qty" convention this
//      day-tab was already designed around.
//   4. Every row it attempted — whether the Store Inventory Sheet accepted
//      it or not — gets its "Processed" column filled with a timestamp, so
//      it's never re-applied on the next run. A barcode the Store
//      Inventory Sheet doesn't recognize is a data problem (that product
//      was never added to that location's tab), not something retrying
//      every 15 minutes would fix — those get logged to a "Sync Issues"
//      tab instead, so you can see and fix them without digging through
//      execution logs.
//
// This only touches Store Inventory's live "Qty On Hand" — it never
// touches this spreadsheet's own SalesLog/Returns/Dashboard data, and it
// never touches Store Inventory's "Qty Counted" (that's still only ever
// written by a physical count in the Counter app).

// Only needed if a POS "Store" value doesn't already match the Store
// Inventory Sheet's location names exactly. Currently identity (confirmed
// the POS already uses "Store 1" / "Store 2") — add entries here if that
// ever changes, e.g. STORE_NAME_MAP['Boutique 1'] = 'Store 1'.
const STORE_NAME_MAP = {
  'Store 1': 'Store 1',
  'Store 2': 'Store 2',
  'Warehouse': 'Warehouse',
};

function _logSyncIssue(location, message) {
  try {
    const sh = _sheet('Sync Issues', ['Timestamp', 'Location', 'Message']);
    sh.appendRow([Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'), location, message]);
  } catch (e) { Logger.log('Could not log sync issue: ' + e); }
}

function syncInventoryFromSales() {
  const tz = CONFIG.TIMEZONE;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const dayTabNames = [];
  for (let i = 0; i < CONFIG.INVENTORY_SYNC_LOOKBACK_DAYS; i++) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    dayTabNames.push('Sales_' + Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
  }

  const byLocation = {}; // location -> [{barcode, delta, type, note}]
  const pending = [];    // { sheet, row } to mark Processed after this run

  dayTabNames.forEach(tabName => {
    const sh = ss.getSheetByName(tabName);
    if (!sh) return; // no sales that day yet — nothing to do
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const values = sh.getRange(2, 1, lastRow - 1, SALES_ITEM_HEADERS.length).getValues();
    values.forEach((row, i) => {
      const processed = row[15]; // Processed
      const barcode = row[3];    // Barcode
      const qty = Number(row[6]) || 0; // Qty (signed)
      const type = row[2];       // 'Sale' | 'Return'
      const storeRaw = row[12];  // Store
      if (processed || !barcode || !qty) return;

      const location = STORE_NAME_MAP[storeRaw] || storeRaw;
      pending.push({ sheet: sh, row: i + 2 }); // +2: header row offset + 1-indexed

      if (!location) {
        _logSyncIssue('(inconnu)', 'Barcode ' + barcode + ' a un magasin vide/non reconnu ("' + storeRaw + '") — ligne ignorée.');
        return;
      }
      if (!byLocation[location]) byLocation[location] = [];
      byLocation[location].push({
        barcode: String(barcode),
        delta: -qty, // sale (qty>0) reduces stock, return (qty<0) restores it
        type: type === 'Return' ? 'return' : 'sold',
        note: row[13] || '', // TxnID, for traceability in the Stock Ledger tab
      });
    });
  });

  const locations = Object.keys(byLocation);
  if (!locations.length && !pending.length) {
    Logger.log('syncInventoryFromSales: nothing new to sync.');
    return;
  }

  let totalApplied = 0;
  locations.forEach(location => {
    const movements = byLocation[location];
    try {
      const res = UrlFetchApp.fetch(CONFIG.INVENTORY_SYNC_URL, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({
          key: CONFIG.INVENTORY_SYNC_KEY,
          action: 'applyStockMovements',
          location: location,
          movements: movements,
        }),
        muteHttpExceptions: true,
      });
      const data = JSON.parse(res.getContentText());
      if (data.success) {
        totalApplied += data.applied || 0;
        if (data.notFound && data.notFound.length) {
          _logSyncIssue(location, 'Code(s)-barres introuvable(s) dans l\'onglet ' + location + ' : ' + data.notFound.join(', '));
        }
      } else {
        _logSyncIssue(location, 'Échec : ' + (data.error || 'erreur inconnue'));
      }
    } catch (e) {
      _logSyncIssue(location, 'Impossible de contacter le Store Inventory Sheet : ' + e);
    }
  });

  // Mark every row this run looked at as Processed, whether it applied
  // cleanly or was logged as an issue — a bad barcode won't fix itself by
  // being retried every 15 minutes, so it goes to Sync Issues instead.
  const timestamp = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm');
  pending.forEach(p => p.sheet.getRange(p.row, 16).setValue(timestamp));

  Logger.log('syncInventoryFromSales: applied %s movement(s) across %s location(s), %s row(s) processed.',
    totalApplied, locations.length, pending.length);
}

function createInventorySyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncInventoryFromSales') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncInventoryFromSales')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Inventory sync trigger installed — runs every 15 minutes.');
}

// Optional: run manually any time to test without waiting for the timer.
function testInventorySync() {
  syncInventoryFromSales();
}

// ══════════════════════════════════════════════════════════
//  NIGHTLY REPORT — mirrors the scanner's own end-of-day report format
// ══════════════════════════════════════════════════════════
function sendNightlyReport() {
  const tz = CONFIG.TIMEZONE;
  const today = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

  const salesSh = _sheet(SHEETS.SALES,
    ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  const returnsSh = _sheet(SHEETS.RETURNS,
    ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);

  const salesRows = salesSh.getDataRange().getValues().slice(1).filter(r => _frDateStr(r[1], tz) === today);
  // "Items that left the store" / by-employee scan counts now come from
  // today's day-tab (Sales_yyyy-MM-dd) — same per-item data, just organized
  // by day instead of one continuous ScanLog. Only Type='Sale' rows count
  // here; returns are already netted separately below via the Returns tab.
  const scanRows  = _dayTab(today).getDataRange().getValues().slice(1).filter(r => r[2] === 'Sale');
  const returnRows = returnsSh.getDataRange().getValues().slice(1).filter(r => _frDateStr(r[2], tz) === today);

  const sales = salesRows.map(r => {
    let splitPayments = [];
    try { splitPayments = r[9] ? JSON.parse(r[9]) : []; } catch (e) { splitPayments = []; }
    return {
      txnId: r[0], date: r[1], time: r[2], employee: r[3], store: r[4],
      itemCount: r[5], total: Number(r[6]) || 0, payMethod: r[7], payMethodLabel: r[8],
      splitPayments, isDelivery: r[10], deliveryId: r[11], customerName: r[12], items: r[13],
    };
  });
  const scans = scanRows.map(r => ({
    date: r[0], time: r[1], barcode: r[3], name: r[4], category: r[5],
    qty: Number(r[6]) || 1, price: Number(r[7]) || 0, lineTotal: Number(r[8]) || 0,
    employee: r[11], store: r[12], txnId: r[13],
  }));
  const returns = returnRows.map(r => ({
    returnId: r[0], txnId: r[1], date: r[2], time: r[3], employee: r[4], store: r[5],
    refundTotal: Number(r[6]) || 0, refundMethod: r[7], refundMethodLabel: r[8],
    reason: r[9], items: r[10],
  }));

  const grossRevenue = sales.reduce((s, t) => s + t.total, 0);
  const totalRefunds = returns.reduce((s, r) => s + r.refundTotal, 0);
  const netRevenue = grossRevenue - totalRefunds;

  // Product summary (mirrors buildDayReport's "ITEMS THAT LEFT THE STORE")
  const products = {};
  scans.forEach(s => {
    const key = s.barcode + '||' + s.name;
    if (!products[key]) products[key] = { name: s.name, category: s.category, count: 0, revenue: 0 };
    products[key].count += s.qty;
    products[key].revenue += s.lineTotal;
  });
  const productRows = Object.values(products)
    .sort((a, b) => b.count - a.count)
    .map(p => `  • ${p.name} (${p.category || '—'}): ${p.count}x${p.revenue ? ' — ' + _fmt(p.revenue) + ' CFA' : ''}`)
    .join('\n') || '  No items scanned today';

  // By-employee scan breakdown
  const byEmp = {};
  scans.forEach(s => { byEmp[s.employee] = (byEmp[s.employee] || 0) + 1; });
  const empRows = Object.entries(byEmp).map(([e, c]) => `  ${e}: ${c} scans`).join('\n') || '  No data';

  // Payment-method breakdown. Split sales contribute their sub-amount to each
  // method they used, and returns contribute a negative line to whichever
  // method refunded them — so cash-drawer / mobile-money totals stay accurate
  // net of both split payments and refunds. Mirrors the scanner's own report.
  const payBlock = ['cash', 'orange', 'wave', 'bank', 'monthly'].map(m => {
    const contribs = [];
    sales.forEach(x => {
      if (x.payMethod === m) {
        contribs.push({ item: x, amount: x.total, kind: 'sale' });
      } else if (x.payMethod === 'split' && Array.isArray(x.splitPayments)) {
        const sp = x.splitPayments.find(p => p.method === m);
        if (sp) contribs.push({ item: x, amount: sp.amount || 0, kind: 'split' });
      }
    });
    returns.forEach(r => {
      if (r.refundMethod === m) contribs.push({ item: r, amount: -r.refundTotal, kind: 'return' });
    });
    if (!contribs.length) return '';
    const total = contribs.reduce((s, c) => s + c.amount, 0);
    const lines = contribs.map(c => {
      const x = c.item;
      if (c.kind === 'return') {
        return `    ${x.returnId} (retour de ${x.txnId})  -${_fmt(Math.abs(c.amount))} CFA — ${x.items || ''}`;
      }
      const items = x.items ? String(x.items).split(' | ').filter(Boolean) : [];
      const itemList = items.slice(0, 3).join(', ') + (items.length > 3 ? ` +${items.length - 3} autre(s)` : '');
      const dlvTag = x.isDelivery ? ` [Livraison ${x.deliveryId} — ${x.customerName}]` : '';
      const splitTag = c.kind === 'split' ? ' (paiement partagé — montant partiel)' : '';
      return `    ${x.txnId}${dlvTag}${splitTag}  ${_fmt(c.amount)} CFA — ${itemList}`;
    }).join('\n');
    return `${PAY_LABELS[m] || m} — ${contribs.length} mouvement(s) — TOTAL NET: ${_fmt(total)} CFA\n${lines}`;
  }).filter(Boolean).join('\n\n') || '  (mode de paiement non renseigné)';

  // Returns summary
  const returnRowsText = returns.map(r =>
    `  • ${r.returnId} (retour de ${r.txnId}) — ${r.employee || '—'} — -${_fmt(r.refundTotal)} CFA ` +
    `(${r.refundMethodLabel || r.refundMethod || '—'})${r.reason ? ': ' + r.reason : ''} — ${r.items || ''}`
  ).join('\n');

  // Per-store breakdown — only shown when today's data spans more than one
  // store name, so a single-store setup keeps getting the same report as
  // before. Each store gets its own totals + top items (from the day-tab),
  // useful when stores keep separate inventory.
  const storeNames = Array.from(new Set(
    [...sales.map(s => s.store), ...scans.map(s => s.store), ...returns.map(r => r.store)].filter(Boolean)
  )).sort();
  const storeBlock = storeNames.length > 1 ? storeNames.map(name => {
    const sSales = sales.filter(s => s.store === name);
    const sScans = scans.filter(s => s.store === name);
    const sReturns = returns.filter(r => r.store === name);
    const gross = sSales.reduce((s, t) => s + t.total, 0);
    const refunds = sReturns.reduce((s, r) => s + r.refundTotal, 0);
    const prod = {};
    sScans.forEach(s => {
      const key = s.barcode + '||' + s.name;
      if (!prod[key]) prod[key] = { name: s.name, count: 0 };
      prod[key].count += s.qty;
    });
    const topItems = Object.values(prod).sort((a, b) => b.count - a.count)
      .slice(0, 5).map(p => `${p.name} (${p.count}x)`).join(', ') || '—';
    return `  ${name} — ${sSales.length} vente(s), CA brut ${_fmt(gross)} CFA` +
      (sReturns.length ? `, ${sReturns.length} retour(s) -${_fmt(refunds)} CFA` : '') +
      ` — CA net ${_fmt(gross - refunds)} CFA\n    Top articles: ${topItems}`;
  }).join('\n\n') : '';

  const now = new Date();
  const body = `BOUTIQUE POS — END OF DAY REPORT (automated)
Date: ${today}
Sent at: ${Utilities.formatDate(now, tz, 'HH:mm')}
=====================================

SUMMARY
• Total items scanned: ${scans.length}
• Total transactions: ${sales.length}
• Total returns: ${returns.length}${returns.length ? ` (-${_fmt(totalRefunds)} CFA)` : ''}
• Gross revenue: ${_fmt(grossRevenue)} CFA
• Net revenue (after returns): ${_fmt(netRevenue)} CFA

─────────────────────────────────────────
VENTES PAR MODE DE PAIEMENT
─────────────────────────────────────────
${payBlock}

BY EMPLOYEE
${empRows}

ITEMS THAT LEFT THE STORE
${productRows}

${returns.length > 0 ? `RETOURS\n${returnRowsText}` : ''}

${storeBlock ? `─────────────────────────────────────────\nPAR MAGASIN\n─────────────────────────────────────────\n${storeBlock}\n` : ''}

Have a good evening!
— Boutique POS System (automated nightly report)`;

  const subject = `🏪 Boutique — End of Day Report ${today}`;
  MailApp.sendEmail(CONFIG.REPORT_EMAIL, subject, body);
}

function _fmt(n) {
  return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
}

// ══════════════════════════════════════════════════════════
//  TRIGGER SETUP — run createNightlyTrigger() ONCE from this editor
// ══════════════════════════════════════════════════════════
function createNightlyTrigger() {
  // Clear any existing triggers for this function first, so re-running
  // this doesn't create duplicates that would send the email twice.
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendNightlyReport') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendNightlyReport')
    .timeBased()
    .atHour(CONFIG.REPORT_HOUR)
    .nearMinute(CONFIG.REPORT_MINUTE)
    .everyDays(1)
    .inTimezone(CONFIG.TIMEZONE)
    .create();
  Logger.log('Nightly trigger installed for %s:%s %s daily.',
    CONFIG.REPORT_HOUR, CONFIG.REPORT_MINUTE, CONFIG.TIMEZONE);
}

// Optional: run manually any time to test the report without waiting for 9:15 PM.
function testNightlyReport() {
  sendNightlyReport();
}

// updateDashboard() is normally triggered live by _refreshDashboard() at the
// end of _logSale/_logReturn/_createInstallmentPlan/_logInstallmentPayment/
// _createOrUpdateHold — but a plain scan (_logScan) doesn't call it, and if
// nothing has closed yet today (first thing in the morning, or a day with
// no completed sale yet) the tab keeps showing whatever it last displayed —
// e.g. still stuck on yesterday's date. This trigger refreshes it on a timer
// too, so the Date/totals always reflect "now" even before the first sale
// of the day closes. Run createDashboardRefreshTrigger() ONCE from this
// editor (same way as createNightlyTrigger) to install it.
function createDashboardRefreshTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'updateDashboard') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('updateDashboard')
    .timeBased()
    .everyMinutes(15)
    .create();
  Logger.log('Dashboard auto-refresh trigger installed — runs every 15 minutes.');
}

// ══════════════════════════════════════════════════════════
//  ONE-TIME REPAIR — run fixDateTimeFormatting() ONCE, right now, from
//  this editor. Not part of the regular setup list above: this only exists
//  to clean up rows written before the Plain Text fix in _sheet()/_dayTab()
//  went in. Any sheet created from here on is protected automatically and
//  never needs this.
// ══════════════════════════════════════════════════════════
// Re-writes every already-corrupted Date/Time/Phone cell back to the exact
// text format the rest of this file expects (dd/MM/yyyy / HH:mm / digits-
// only), then locks the column to Plain Text so it can't happen again.
// Safe to re-run — cells that are already correct text pass through
// unchanged.
function fixDateTimeFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = CONFIG.TIMEZONE;

  Object.keys(TEXT_COLUMNS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) _repairSheetColumns(sh, TEXT_COLUMNS[name], tz);
  });

  // Every Sales_yyyy-MM-dd day-tab, whenever it was created.
  ss.getSheets().forEach(sh => {
    if (/^Sales_\d{4}-\d{2}-\d{2}$/.test(sh.getName())) {
      _repairSheetColumns(sh, [1, 2, 4], tz); // Date, Time, Barcode
    }
  });

  Logger.log('Date/Time/Barcode/Phone formatting repaired across all sheets.');
}

function _repairSheetColumns(sh, cols, tz) {
  const lastRow = sh.getLastRow();
  cols.forEach(col => {
    _forceTextColumns(sh, [col]); // lock the column first
    if (lastRow < 2) return;
    const range = sh.getRange(2, col, lastRow - 1, 1);
    const fixed = range.getValues().map(row => {
      const v = row[0];
      if (v instanceof Date) {
        // Sheets stores a "time only" value against its own epoch
        // (30 Dec 1899) — that year is the tell for which pattern to
        // reconstruct with; anything else was a real calendar date.
        return [v.getFullYear() === 1899
          ? Utilities.formatDate(v, tz, 'HH:mm')
          : Utilities.formatDate(v, tz, 'dd/MM/yyyy')];
      }
      if (typeof v === 'number') {
        // A barcode/phone that got auto-converted to Number. Stringifying
        // it back can't restore a leading zero that was already dropped —
        // the Plain Text lock above is what actually prevents that; this
        // just stops it from getting worse (e.g. scientific notation) on
        // the next read/write.
        return [String(v)];
      }
      return [v];
    });
    range.setValues(fixed);
  });
}

// ══════════════════════════════════════════════════════════
//  LIVE DASHBOARD — run setupDashboard() ONCE from this editor.
//  Creates/resets a "Dashboard" tab with live formulas (not a snapshot) —
//  every figure recalculates the instant Sales Sync writes a new row, so
//  opening this tab (including from the Google Sheets phone app) always
//  shows current numbers. No email, no push — check it whenever you want.
// ══════════════════════════════════════════════════════════
// Builds the Dashboard tab's static layout (labels, headers, formatting)
// and does NOT use spreadsheet formulas — updateDashboard() below fills in
// the actual numbers as plain values. This avoids a real gotcha: Apps
// Script writes formulas in US syntax (commas between arguments), but a
// Sheet whose locale uses a comma as the decimal separator (French locale,
// common when the spreadsheet's region is set to France/francophone
// countries) expects semicolons instead — every formula breaks with
// #ERROR! in that case. Plain values sidestep the problem entirely.
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Make sure every sheet referenced below exists first (a brand new
  // spreadsheet may not have Returns/InstallmentPlans/Holds rows yet).
  _sheet(SHEETS.SALES, ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  _sheet(SHEETS.RETURNS, ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);
  _instPlansSheet();
  _sheet(SHEETS.HOLDS, HOLD_HEADERS);

  let sh = ss.getSheetByName('Dashboard');
  if (sh) { sh.clear(); } else { sh = ss.insertSheet('Dashboard'); }
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1); // pin it as the first tab

  sh.getRange('A1').setValue('🏪 Tableau de bord — Aujourd\'hui').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue("Mise à jour…").setFontStyle('italic').setFontColor('#666666');

  sh.getRange('A4').setValue('Date');
  sh.getRange('A5').setValue('Ventes aujourd\'hui');
  sh.getRange('A6').setValue("Chiffre d'affaires brut");
  sh.getRange('A7').setValue('Retours aujourd\'hui');
  sh.getRange('A8').setValue('Montant remboursé');
  sh.getRange('A9').setValue("Chiffre d'affaires net").setFontWeight('bold');
  sh.getRange('B9').setFontWeight('bold');

  sh.getRange('A11').setValue('— Par mode de paiement (net, partages et retours inclus) —').setFontWeight('bold');
  const methods = [['cash','💵 Cash'],['orange','🟠 Orange Money'],['wave','🌊 Wave'],['bank','🏦 Virement'],['monthly','📅 Mensualité']];
  methods.forEach((m, i) => sh.getRange(12 + i, 1).setValue(m[1]));

  sh.getRange('A18').setValue('— Mensualités en cours —').setFontWeight('bold');
  sh.getRange('A19').setValue('Plans ouverts');
  sh.getRange('A20').setValue('Solde total dû');

  sh.getRange('A22').setValue('— Réservations —').setFontWeight('bold');
  sh.getRange('A23').setValue('Réservations actives');

  sh.getRange('A25').setValue('— 10 dernières ventes —').setFontWeight('bold');
  sh.getRange('A26').setValue('Reçu N°');
  sh.getRange('B26').setValue('Date');
  sh.getRange('C26').setValue('Heure');
  sh.getRange('D26').setValue('Employé');
  sh.getRange('E26').setValue('Total');
  sh.getRange('F26').setValue('Paiement');
  sh.getRange('G26').setValue('Magasin');
  sh.getRange('A26:G26').setFontWeight('bold');

  // Multi-store breakdown — dynamic list, one row per store name seen in
  // today's data (up to 10 reserved rows). The B4:B9 summary above stays a
  // combined total across every store; this section splits it out.
  sh.getRange('A38').setValue('— Chiffre d\'affaires par magasin (aujourd\'hui) —').setFontWeight('bold');
  sh.getRange('A39').setValue('Magasin');
  sh.getRange('B39').setValue('Ventes');
  sh.getRange('C39').setValue('CA brut');
  sh.getRange('D39').setValue('Retours');
  sh.getRange('E39').setValue('Remboursé');
  sh.getRange('F39').setValue('CA net');
  sh.getRange('A39:F39').setFontWeight('bold');

  sh.autoResizeColumns(1, 7);
  sh.setColumnWidth(1, 230);

  updateDashboard(); // fill in real numbers right away
  Logger.log('Dashboard tab created. It refreshes automatically as sales/returns/payments/holds come in.');
}

// Recomputes every figure on the Dashboard tab and writes plain values —
// called automatically after every sale, return, hold, and Mensualité
// event (see the end of each _log*/_create* handler above), so opening the
// tab always shows numbers from moments ago, not a stale snapshot.
function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Dashboard');
  if (!sh) return; // setupDashboard() hasn't been run yet — nothing to update

  const tz = CONFIG.TIMEZONE;
  const today = Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy');

  const salesSh = _sheet(SHEETS.SALES, ['TxnID','Date','Time','Employee','Store','ItemCount','Total','PayMethod','PayMethodLabel','SplitPayments','IsDelivery','DeliveryID','CustomerName','Items']);
  const returnsSh = _sheet(SHEETS.RETURNS, ['ReturnID','TxnID','Date','Time','Employee','Store','RefundTotal','RefundMethod','RefundMethodLabel','Reason','Items']);
  const plansSh = _instPlansSheet();
  const holdsSh = _sheet(SHEETS.HOLDS, HOLD_HEADERS);

  // Re-lock the Date/Time columns to Plain Text every time this runs (every
  // sale, and every 15 minutes) — cheap formatting-only calls, and a second
  // safety net in case Sheets' auto-detect ever flips the column back to
  // Automatic again (see TEXT_COLUMNS comment above for why that happens).
  _forceTextColumns(salesSh, TEXT_COLUMNS.SalesLog);
  _forceTextColumns(returnsSh, TEXT_COLUMNS.Returns);

  const allSalesRows = salesSh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const allReturnRows = returnsSh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const planRows = plansSh.getDataRange().getValues().slice(1).filter(r => r[0]);
  const holdRows = holdsSh.getDataRange().getValues().slice(1).filter(r => r[0]);

  const salesToday = allSalesRows.filter(r => _frDateStr(r[1], tz) === today).map(r => {
    let splitPayments = [];
    try { splitPayments = r[9] ? JSON.parse(r[9]) : []; } catch (e) { splitPayments = []; }
    return { total: Number(r[6]) || 0, payMethod: r[7], splitPayments, store: r[4] || '' };
  });
  const returnsToday = allReturnRows.filter(r => _frDateStr(r[2], tz) === today)
    .map(r => ({ refundTotal: Number(r[6]) || 0, refundMethod: r[7], store: r[5] || '' }));

  const grossRevenue = salesToday.reduce((s, t) => s + t.total, 0);
  const totalRefunds = returnsToday.reduce((s, r) => s + r.refundTotal, 0);
  const netRevenue = grossRevenue - totalRefunds;

  sh.getRange('B4').setValue(today);
  sh.getRange('B5').setValue(salesToday.length);
  sh.getRange('B6').setValue(grossRevenue);
  sh.getRange('B7').setValue(returnsToday.length);
  sh.getRange('B8').setValue(totalRefunds);
  sh.getRange('B9').setValue(netRevenue);

  // Payment-method breakdown, netted the same way as the nightly report:
  // split-sale sub-amounts credit their own method, refunds debit theirs.
  const methods = ['cash', 'orange', 'wave', 'bank', 'monthly'];
  methods.forEach((m, i) => {
    let total = 0;
    salesToday.forEach(x => {
      if (x.payMethod === m) total += x.total;
      else if (x.payMethod === 'split' && Array.isArray(x.splitPayments)) {
        const sp = x.splitPayments.find(p => p.method === m);
        if (sp) total += sp.amount || 0;
      }
    });
    returnsToday.forEach(r => { if (r.refundMethod === m) total -= r.refundTotal; });
    sh.getRange(12 + i, 2).setValue(total);
  });

  // Mensualité — Status is column O (index 14), Balance is column M (index 12)
  const openPlans = planRows.filter(r => r[14] === 'open');
  sh.getRange('B19').setValue(openPlans.length);
  sh.getRange('B20').setValue(openPlans.reduce((s, r) => s + (Number(r[12]) || 0), 0));

  // Holds — Status is column L (index 11)
  const activeHolds = holdRows.filter(r => r[11] === 'active');
  sh.getRange('B23').setValue(activeHolds.length);

  // 10 most recent sales overall (not just today)
  const recent = allSalesRows.slice(-10).reverse();
  sh.getRange(27, 1, 10, 7).clearContent();
  recent.forEach((r, i) => {
    sh.getRange(27 + i, 1).setValue(r[0]);          // TxnID
    sh.getRange(27 + i, 2).setValue(r[1]);          // Date
    sh.getRange(27 + i, 3).setValue(r[2]);          // Time
    sh.getRange(27 + i, 4).setValue(r[3]);          // Employee
    sh.getRange(27 + i, 5).setValue(Number(r[6]) || 0); // Total
    sh.getRange(27 + i, 6).setValue(r[8]);          // PayMethodLabel
    sh.getRange(27 + i, 7).setValue(r[4]);          // Store
  });
  if (!recent.length) sh.getRange('A27').setValue('Aucune vente pour le moment');

  // Per-store breakdown (today) — one row for EVERY store configured in
  // StoreSettings, even ones with zero sales today, up to 10 reserved rows
  // starting at row 40. (Previously this only listed stores that had at
  // least one sale/return today, so a quiet store just vanished from this
  // table instead of showing zeros — easy to mistake for a real bug.)
  sh.getRange(40, 1, 10, 6).clearContent();
  const storeSettingsSh = _sheet(SHEETS.STORE_SETTINGS, STORE_SETTINGS_HEADERS);
  const configuredStoreNames = storeSettingsSh.getDataRange().getValues().slice(1)
    .map(r => r[2]).filter(Boolean); // column C = display Name
  const seenStoreNames = [...salesToday.map(s => s.store), ...returnsToday.map(r => r.store)].filter(Boolean);
  const storeNames = Array.from(new Set([...configuredStoreNames, ...seenStoreNames])).sort();
  if (!storeNames.length) {
    sh.getRange('A40').setValue('Aucune boutique configurée pour le moment');
  } else {
    storeNames.forEach((name, i) => {
      const sSales = salesToday.filter(s => s.store === name);
      const sReturns = returnsToday.filter(r => r.store === name);
      const gross = sSales.reduce((s, t) => s + t.total, 0);
      const refunds = sReturns.reduce((s, r) => s + r.refundTotal, 0);
      sh.getRange(40 + i, 1).setValue(name);
      sh.getRange(40 + i, 2).setValue(sSales.length);
      sh.getRange(40 + i, 3).setValue(gross);
      sh.getRange(40 + i, 4).setValue(sReturns.length);
      sh.getRange(40 + i, 5).setValue(refunds);
      sh.getRange(40 + i, 6).setValue(gross - refunds);
    });
  }

  sh.getRange('A2').setValue('Mis à jour : ' + Utilities.formatDate(new Date(), tz, 'dd/MM/yyyy HH:mm:ss'))
    .setFontStyle('italic').setFontColor('#666666');
}

// ══════════════════════════════════════════════════════════
//  PRODUCT IMAGES — a photo for each barcode, found automatically
//
//  Every barcode sold in the last 90 days gets a row in the
//  "Product Images" tab. Every 15 minutes runImageLookup() takes the rows
//  still "À chercher" and looks for a photo, in this order:
//    1. Open Beauty Facts + Open Food Facts (free, no limit)
//    2. UPCitemdb (free trial: max ~95 per day, 1 call every 11 s)
//  A photo that is found is downloaded into the Drive folder
//  "Boutique 1981 – Photos produits" as <barcode>.jpg, so it keeps working
//  even if the original website removes it. The Scanner reads the list
//  through doGet ?action=getProductImages and shows the photo when the
//  item is scanned.
//
//  What the owner can do in the tab:
//    • Statut "Faux"   → wrong picture: the Scanner stops showing it.
//    • "Lien manuel"   → paste any image link (any status): it is
//                        downloaded on the next run and replaces the photo.
//  Codes made by our own Barcode Generator (0610000…) are marked
//  "Code interne" and skipped (no database knows them); a manual link
//  still works for them.
//
//  SETUP: run startImageLookup() once from the editor and accept the
//  Google Drive permission. stopImageLookup() removes the 15-min timer.
// ══════════════════════════════════════════════════════════
const IMG = {
  SHEET: 'Product Images',
  FOLDER: 'Boutique 1981 – Photos produits',
  DAYS: 90,                 // "sold in the last 3 months"
  UPC_PER_DAY: 95,          // UPCitemdb free trial allows 100/day — keep a margin
  UPC_GAP_MS: 11000,        // and max 6/minute
  RUN_MS: 270000,           // stop each run after 4.5 min (Apps Script limit is 6)
  MAX_TRIES: 3,
  INTERNAL_PREFIX: '610000', // 0610000… with the leading zero removed
};
const IMG_HEADERS = ['Barcode', 'Produit', 'Statut', 'Photo', 'Source', 'Lien photo',
  "Lien d'origine", 'Lien manuel', 'Vérifié le', 'Essais', 'Suggestions'];
const IMG_C = { BC: 0, NAME: 1, STATUS: 2, PHOTO: 3, SOURCE: 4, URL: 5, ORIG: 6, MANUAL: 7, AT: 8, TRIES: 9, SUGG: 10 };
// Brave image search (web-wide: Walmart, Sam's, Amazon, brand sites…) —
// up to 4 clean candidate photos per product, picked by the owner in the
// Scanner. The key lives in Script properties as BRAVE_API_KEY (Project
// Settings → Script properties); without it this step is simply skipped.
// $5 free credit/month ≈ 1 000 searches; we stop at BRAVE_PER_MONTH.
const IMG_BRAVE = { PER_MONTH: 900, GAP_MS: 1100, KEEP: 4, PHASE2_MS: 180000 };
const IMG_GOOD_SITES = ['samsclub.com', 'walmart.com', 'target.com', 'amazon.', 'costco.com', 'kroger.com',
  'walgreens.com', 'cvs.com', 'ulta.com', 'sephora.com', 'instacart.com', 'heb.com', 'jumia.'];
const IMG_ST = { TODO: 'À chercher', FOUND: 'Trouvé', VALID: 'Validé', NONE: 'Introuvable', WRONG: 'Faux',
  INTERNAL: 'Code interne', ERR: 'Erreur' };
const IMG_SHOWN = ['Trouvé', 'Validé']; // statuses whose photo the Scanner shows
const IMG_OFF_TRIED = 'OFF ✗'; // Source note while a row waits for its UPCitemdb turn

// Real UPC/EAN codes end with a check digit; made-up or hand-typed codes
// usually fail it. 8-digit codes are left alone (UPC-E uses another rule).
function _imgGtinValid(bc) {
  const d = String(bc || '').replace(/\D/g, '');
  if (d.length <= 8) return true;
  if (d.length > 14) return false;
  const g = ('00000000000000' + d).slice(-14);
  let sum = 0;
  for (let i = 0; i < 13; i++) sum += Number(g[i]) * (i % 2 === 0 ? 3 : 1);
  return (10 - (sum % 10)) % 10 === Number(g[13]);
}
// Our own generated codes (0610000…) or anything failing the check digit.
function _imgIsInternal(bc) {
  const k = _imgKey(bc);
  return (k.indexOf(IMG.INTERNAL_PREFIX) === 0 && k.length === 12) || !_imgGtinValid(bc);
}
function _imgKey(bc) {
  return String(bc || '').replace(/\D/g, '').replace(/^0+/, '');
}

function _imgSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(IMG.SHEET);
  if (!sh) {
    sh = ss.insertSheet(IMG.SHEET);
    sh.getRange(1, 1, 1, IMG_HEADERS.length).setValues([IMG_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@'); // keep leading zeros
    sh.setColumnWidth(2, 260); sh.setColumnWidth(4, 90); sh.setColumnWidth(6, 120);
    sh.setColumnWidth(7, 120); sh.setColumnWidth(8, 160);
  }
  // v2 (21 Oct): + "Validé" status; the Photo column formula is gone (it
  // broke in a French-locale sheet, and photos are now checked in the
  // Scanner's Admin → 📷 Photos produits).
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('IMG_SHEET_V') !== '2') {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(IMG_ST).map(k => IMG_ST[k]), true)
      .setAllowInvalid(false).build();
    sh.getRange(2, IMG_C.STATUS + 1, sh.getMaxRows() - 1, 1).setDataValidation(rule);
    sh.getRange(2, IMG_C.PHOTO + 1, sh.getMaxRows() - 1, 1).clearContent();
    sh.getRange(1, IMG_C.PHOTO + 1).setValue('Photo (voir Scanner)');
    if (sh.getLastRow() >= 2) sh.setRowHeights(2, sh.getLastRow() - 1, 21);
    props.setProperty('IMG_SHEET_V', '2');
  }
  if (props.getProperty('IMG_SHEET_V') !== '3' && props.getProperty('IMG_SHEET_V') !== '4') {
    sh.getRange(1, IMG_C.SUGG + 1).setValue('Suggestions').setFontWeight('bold');
    props.setProperty('IMG_SHEET_V', '3');
  }
  // v4: rows still being searched whose code is ours / made up → "Code interne".
  if (props.getProperty('IMG_SHEET_V') === '3') {
    const last = sh.getLastRow();
    if (last >= 2) {
      const rng = sh.getRange(2, IMG_C.STATUS + 1, last - 1, 1);
      const st = rng.getValues();
      const bcs = sh.getRange(2, 1, last - 1, 1).getValues();
      let n = 0;
      st.forEach((r, i) => {
        if ((r[0] === IMG_ST.TODO || r[0] === IMG_ST.ERR) && _imgIsInternal(bcs[i][0])) { r[0] = IMG_ST.INTERNAL; n++; }
      });
      if (n) rng.setValues(st);
      Logger.log(n + ' code(s) interne(s) reclassé(s)');
    }
    props.setProperty('IMG_SHEET_V', '4');
  }
  return sh;
}

function _imgFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('IMG_FOLDER_ID');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const it = DriveApp.getFoldersByName(IMG.FOLDER);
  const folder = it.hasNext() ? it.next() : DriveApp.createFolder(IMG.FOLDER);
  props.setProperty('IMG_FOLDER_ID', folder.getId());
  return folder;
}

// Adds every barcode sold in the last IMG.DAYS days that isn't in the tab yet.
function _imgRefreshList_(sh) {
  const tz = CONFIG.TIMEZONE;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const have = {};
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(r => { const k = _imgKey(r[0]); if (k) have[k] = true; });
  const found = {}, order = [];
  const now = new Date();
  for (let i = 0; i < IMG.DAYS; i++) {
    const tab = ss.getSheetByName('Sales_' + Utilities.formatDate(new Date(now.getTime() - i * 86400000), tz, 'yyyy-MM-dd'));
    if (!tab || tab.getLastRow() < 2) continue;
    tab.getDataRange().getValues().forEach(r => {
      if (r[2] !== 'Sale') return;
      const raw = String(r[3] || '').trim();
      if (!/^\d+$/.test(raw)) return;               // typed names, garbled scans
      const k = _imgKey(raw);
      if (k.length < 7 || have[k] || found[k]) return; // 00000 placeholder, short codes
      found[k] = { bc: raw, name: String(r[4] || '').trim() };
      order.push(k);
    });
  }
  if (!order.length) return 0;
  const rows = order.map(k => {
    const internal = _imgIsInternal(found[k].bc);
    return [found[k].bc, found[k].name, internal ? IMG_ST.INTERNAL : IMG_ST.TODO, '', '', '', '', '', '', 0, ''];
  });
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(start, 1, rows.length, IMG_HEADERS.length).setValues(rows);
  return rows.length;
}

// Open Beauty Facts + Open Food Facts, in parallel. Returns an image URL or ''.
function _imgFromOFF_(bc) {
  const codes = [bc];
  if (bc.length === 12) codes.push('0' + bc);
  const reqs = [];
  ['world.openbeautyfacts.org', 'world.openfoodfacts.org'].forEach(host => codes.forEach(c => reqs.push({
    url: 'https://' + host + '/api/v2/product/' + c + '.json?fields=image_front_url,image_url',
    muteHttpExceptions: true, headers: { 'User-Agent': 'Boutique1981-POS/1.0 (1981estvideo@gmail.com)' },
  })));
  let res;
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) { return { error: true }; }
  for (let i = 0; i < res.length; i++) {
    if (res[i].getResponseCode() !== 200) continue;
    try {
      const d = JSON.parse(res[i].getContentText());
      const p = d && d.status === 1 && d.product;
      const url = p && (p.image_front_url || p.image_url);
      if (url) return { url: url, source: i < codes.length ? 'Open Beauty Facts' : 'Open Food Facts' };
    } catch (e) {}
  }
  return {};
}

// UPCitemdb free trial. { urls:[...] } | { none:true } | { limit:true } | { error:true }
function _imgFromUPC_(bc) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
  if (props.getProperty('IMG_UPC_DAY') !== today) { props.setProperty('IMG_UPC_DAY', today); props.setProperty('IMG_UPC_COUNT', '0'); }
  const used = Number(props.getProperty('IMG_UPC_COUNT')) || 0;
  if (used >= IMG.UPC_PER_DAY) return { limit: true };
  props.setProperty('IMG_UPC_COUNT', String(used + 1));
  let r;
  try { r = UrlFetchApp.fetch('https://api.upcitemdb.com/prod/trial/lookup?upc=' + bc, { muteHttpExceptions: true }); }
  catch (e) { return { error: true }; }
  const code = r.getResponseCode();
  if (code === 429) { props.setProperty('IMG_UPC_COUNT', String(IMG.UPC_PER_DAY)); return { limit: true }; }
  if (code === 404 || code === 400) return { none: true };
  if (code !== 200) return { error: true };
  try {
    const d = JSON.parse(r.getContentText());
    const urls = [];
    (d.items || []).forEach(it => (it.images || []).forEach(u => { if (u && urls.indexOf(u) < 0) urls.push(u); }));
    return urls.length ? { urls: urls } : { none: true };
  } catch (e) { return { error: true }; }
}

// Downloads an image; returns the blob or null if it isn't a real picture.
function _imgDownload_(url) {
  try {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; Boutique1981-POS/1.0)' } });
    if (r.getResponseCode() !== 200) return null;
    const blob = r.getBlob();
    const type = String(blob.getContentType() || '').toLowerCase();
    if (type.indexOf('image/') !== 0 || blob.getBytes().length < 1500) return null;
    return blob;
  } catch (e) { return null; }
}

// Saves into the Drive folder as <barcode>.<ext>; returns the link the Scanner uses.
function _imgSave_(blob, bc, folder, oldUrl) {
  const m = /[?&]id=([\w-]+)/.exec(String(oldUrl || ''));
  if (m) { try { DriveApp.getFileById(m[1]).setTrashed(true); } catch (e) {} }
  const type = String(blob.getContentType() || '').toLowerCase();
  const ext = type.indexOf('png') >= 0 ? 'png' : type.indexOf('webp') >= 0 ? 'webp' : 'jpg';
  const file = folder.createFile(blob.setName(bc + '.' + ext));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w400';
}

function _imgWriteRow_(sh, rowNum, row) {
  row[IMG_C.PHOTO] = '';
  row[IMG_C.AT] = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  sh.getRange(rowNum, 1, 1, IMG_HEADERS.length).setValues([row]);
}

// Search words from the product name: drop sizes/counts noise, keep brand + product words.
function _imgQuery_(name) {
  return String(name || '').replace(/[\/|:;()\[\]"«»]+/g, ' ').replace(/\s+-\s+/g, ' ')
    .replace(/\s+/g, ' ').trim().slice(0, 90);
}
// { list: [{img, thumb, site, title, w, h}] } | { limit:true } | { error:true }
function _imgBraveSearch_(key, row) {
  const props = PropertiesService.getScriptProperties();
  const month = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');
  if (props.getProperty('IMG_BRAVE_MONTH') !== month) { props.setProperty('IMG_BRAVE_MONTH', month); props.setProperty('IMG_BRAVE_COUNT', '0'); }
  const used = Number(props.getProperty('IMG_BRAVE_COUNT')) || 0;
  if (used >= _braveCap_()) return { limit: true };
  const q = _imgQuery_(typeof row === 'string' ? row : row[IMG_C.NAME]);
  if (!q) return { list: [] };
  props.setProperty('IMG_BRAVE_COUNT', String(used + 1));
  let r;
  try {
    r = UrlFetchApp.fetch('https://api.search.brave.com/res/v1/images/search?count=30&safesearch=strict&country=US&q=' + encodeURIComponent(q),
      { muteHttpExceptions: true, headers: { 'X-Subscription-Token': key, 'Accept': 'application/json' } });
  } catch (e) { return { error: true }; }
  const code = r.getResponseCode();
  if (code === 429 || code === 402) return { limit: true };
  if (code !== 200) { Logger.log('Brave ' + code + ': ' + r.getContentText().slice(0, 300)); return code >= 500 ? { error: true } : { limit: true }; }
  let d; try { d = JSON.parse(r.getContentText()); } catch (e) { return { error: true }; }
  const seen = {};
  const list = (d.results || []).map(x => {
    const p = x.properties || {};
    const img = p.url || x.url || '';
    const thumb = (x.thumbnail && x.thumbnail.src) || '';
    let site = x.source || (x.meta_url && x.meta_url.hostname) || '';
    if (!site) { const m = /^https?:\/\/([^\/]+)/.exec(String(x.url || img)); site = m ? m[1] : ''; }
    site = String(site).replace(/^www\./, '');
    const w = Number(p.width || (x.thumbnail && x.thumbnail.width)) || 0;
    const h = Number(p.height || (x.thumbnail && x.thumbnail.height)) || 0;
    let score = 0;
    if (IMG_GOOD_SITES.some(g => site.indexOf(g) >= 0)) score += 3;
    if (w && h) { const ratio = w / h; if (ratio > 0.6 && ratio < 1.6) score += 1; if (Math.min(w, h) >= 400) score += 1; }
    return { img: img, thumb: thumb, site: site, title: String(x.title || '').slice(0, 80), score: score };
  }).filter(c => (c.img || c.thumb) && !seen[c.img] && (seen[c.img] = true));
  list.sort((a, b) => b.score - a.score);
  return { list: list.slice(0, IMG_BRAVE.KEEP).map(c => ({ img: c.img, thumb: c.thumb, site: c.site, title: c.title })) };
}

/** Paste your Brave key between the quotes, run once, then delete the key from this line. */
function setBraveKey() {
  const KEY = '';
  if (!KEY) throw new Error('Collez la clé Brave entre les guillemets de KEY, puis relancez.');
  PropertiesService.getScriptProperties().setProperty('BRAVE_API_KEY', KEY.trim());
  Logger.log('Clé Brave enregistrée ✔');
}

/** Timer target (every 15 min). Safe to run by hand too. */
function runImageLookup() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  const t0 = Date.now();
  try {
    const sh = _imgSheet_();
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
    if (props.getProperty('IMG_LIST_DAY') !== today) {
      const added = _imgRefreshList_(sh);
      props.setProperty('IMG_LIST_DAY', today);
      if (added) Logger.log(added + ' nouveaux codes ajoutés');
    }
    const last = sh.getLastRow();
    if (last < 2) return;
    sh.getRange(2, 1, last - 1, 1).setNumberFormat('@');
    const data = sh.getRange(2, 1, last - 1, IMG_HEADERS.length).getValues();
    const folder = _imgFolder_();
    let upcBlocked = false, lastUpc = 0;
    const braveKey = props.getProperty('BRAVE_API_KEY');
    const p2Ms = braveKey ? IMG_BRAVE.PHASE2_MS : IMG.RUN_MS; // leave time for the Brave step
    const timeLeft = () => Date.now() - t0 < IMG.RUN_MS;
    const p2Left = () => Date.now() - t0 < p2Ms;

    // 1) Manual links first — they always win.
    for (let i = 0; i < data.length && timeLeft(); i++) {
      const row = data[i], manual = String(row[IMG_C.MANUAL] || '').trim();
      if (!manual || manual === String(row[IMG_C.ORIG] || '').trim()) continue;
      const bc = String(row[IMG_C.BC]).trim();
      const blob = _imgDownload_(manual);
      if (blob) {
        row[IMG_C.URL] = _imgSave_(blob, bc, folder, row[IMG_C.URL]);
        row[IMG_C.ORIG] = manual; row[IMG_C.SOURCE] = 'Manuel'; row[IMG_C.STATUS] = IMG_ST.FOUND;
      } else {
        row[IMG_C.SOURCE] = 'Lien manuel illisible';
      }
      _imgWriteRow_(sh, i + 2, row);
    }

    // 2) Rows still to search.
    for (let i = 0; i < data.length && p2Left(); i++) {
      const row = data[i];
      if (row[IMG_C.STATUS] !== IMG_ST.TODO) continue;
      let bc = String(row[IMG_C.BC]).replace(/\D/g, '');
      if (bc.length > 8 && bc.length < 12) bc = ('000' + bc).slice(-12); // leading zeros lost by Sheets
      let found = null;

      let offJustTried = false;
      if (row[IMG_C.SOURCE] !== IMG_OFF_TRIED) {
        offJustTried = true;
        const off = _imgFromOFF_(bc);
        Utilities.sleep(2500); // stay well under Open Food Facts' 100 requests/minute
        if (off.url) {
          const blob = _imgDownload_(off.url);
          if (blob) found = { blob: blob, url: off.url, source: off.source };
        }
        if (!found && off.error) { // network trouble: count a try, move on
          row[IMG_C.TRIES] = (Number(row[IMG_C.TRIES]) || 0) + 1;
          if (row[IMG_C.TRIES] >= IMG.MAX_TRIES) row[IMG_C.STATUS] = IMG_ST.ERR;
          _imgWriteRow_(sh, i + 2, row);
          continue;
        }
        if (!found) row[IMG_C.SOURCE] = IMG_OFF_TRIED;
      }

      if (!found) {
        if (upcBlocked) {
          if (offJustTried) _imgWriteRow_(sh, i + 2, row); // remember OFF was done, wait for tomorrow's UPC quota
          continue;
        }
        const wait = IMG.UPC_GAP_MS - (Date.now() - lastUpc);
        if (wait > 0) { if (Date.now() - t0 + wait > p2Ms) break; Utilities.sleep(wait); }
        lastUpc = Date.now();
        const upc = _imgFromUPC_(bc);
        if (upc.limit) { upcBlocked = true; if (offJustTried) _imgWriteRow_(sh, i + 2, row); continue; }
        if (upc.urls) {
          for (let j = 0; j < upc.urls.length && j < 4; j++) {
            const blob = _imgDownload_(upc.urls[j]);
            if (blob) { found = { blob: blob, url: upc.urls[j], source: 'UPCitemdb' }; break; }
          }
          if (!found) { row[IMG_C.STATUS] = IMG_ST.NONE; row[IMG_C.SOURCE] = 'Photos UPCitemdb illisibles'; }
        } else if (upc.none) {
          row[IMG_C.STATUS] = IMG_ST.NONE; row[IMG_C.SOURCE] = '';
        } else {
          row[IMG_C.TRIES] = (Number(row[IMG_C.TRIES]) || 0) + 1;
          if (row[IMG_C.TRIES] >= IMG.MAX_TRIES) row[IMG_C.STATUS] = IMG_ST.ERR;
        }
      }

      if (found) {
        row[IMG_C.URL] = _imgSave_(found.blob, bc, folder, row[IMG_C.URL]);
        row[IMG_C.ORIG] = found.url; row[IMG_C.SOURCE] = found.source; row[IMG_C.STATUS] = IMG_ST.FOUND;
      }
      _imgWriteRow_(sh, i + 2, row);
    }

    // 2b) Names + photos for products found by the full-store scan
    //     ("Nouveaux codes" tab) — before the photo step, so they come first.
    try { _ncLookupBatch_(t0 + Math.min(IMG.RUN_MS, (Date.now() - t0) + 150000)); } catch (e) { Logger.log('Nouveaux codes : ' + e); }

    // 3) Brave image search: clean photo suggestions for every product not
    //    yet validated and not yet searched (Suggestions cell empty).
    if (braveKey) {
      let lastBrave = 0;
      for (let i = 0; i < data.length && timeLeft(); i++) {
        const row = data[i];
        if (row[IMG_C.STATUS] === IMG_ST.VALID || row[IMG_C.STATUS] === IMG_ST.INTERNAL || String(row[IMG_C.SUGG] || '') !== '') continue;
        if (_imgIsInternal(row[IMG_C.BC])) continue; // our own / made-up codes: no paid search
        if (row[IMG_C.STATUS] === IMG_ST.FOUND && /^(UPCitemdb|Manuel)/.test(String(row[IMG_C.SOURCE]))) continue; // already a shop/own photo
        const wait = IMG_BRAVE.GAP_MS - (Date.now() - lastBrave);
        if (wait > 0) Utilities.sleep(wait);
        lastBrave = Date.now();
        const res = _imgBraveSearch_(braveKey, row);
        if (res.limit) break;
        if (res.error) continue;
        row[IMG_C.SUGG] = JSON.stringify(res.list);
        _imgWriteRow_(sh, i + 2, row);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/** Run once: creates the tab + Drive folder, starts the 15-min timer, does a first batch. */
function startImageLookup() {
  _imgSheet_();
  _imgFolder_();
  stopImageLookup();
  ScriptApp.newTrigger('runImageLookup').timeBased().everyMinutes(15).create();
  PropertiesService.getScriptProperties().deleteProperty('IMG_LIST_DAY'); // rebuild the list now
  runImageLookup();
  imageLookupStatus();
}

function stopImageLookup() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runImageLookup') ScriptApp.deleteTrigger(t);
  });
}

/** Counts per status, in the execution log. */
function imageLookupStatus() {
  const sh = _imgSheet_();
  const last = sh.getLastRow();
  const counts = {};
  if (last >= 2) sh.getRange(2, IMG_C.STATUS + 1, last - 1, 1).getValues()
    .forEach(r => { counts[r[0]] = (counts[r[0]] || 0) + 1; });
  Logger.log(JSON.stringify(counts));
  return counts;
}

// For the Scanner: { images: { <barcode without leading zeros>: <photo link> } }
function _getProductImages() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(IMG.SHEET);
  const images = {};
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, IMG_HEADERS.length).getValues().forEach(r => {
      if (IMG_SHOWN.indexOf(r[IMG_C.STATUS]) < 0 || !r[IMG_C.URL]) return;
      const k = _imgKey(r[IMG_C.BC]);
      if (k) images[k] = String(r[IMG_C.URL]);
    });
  }
  return _json({ images: images, count: Object.keys(images).length, generatedAt: new Date().toISOString() });
}

function _imgSugg_(v) { try { const a = JSON.parse(String(v || '[]')); return Array.isArray(a) ? a : []; } catch (e) { return []; } }

// Full list for the Scanner's Admin → 📷 Photos produits review screen.
function _getProductImageList() {
  const sh = _imgSheet_();
  const items = [];
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, IMG_HEADERS.length).getValues().forEach(r => {
      if (!String(r[IMG_C.BC]).trim()) return;
      items.push({ barcode: String(r[IMG_C.BC]).trim(), name: String(r[IMG_C.NAME] || ''), status: String(r[IMG_C.STATUS] || ''),
        source: String(r[IMG_C.SOURCE] || ''), url: String(r[IMG_C.URL] || ''), at: String(r[IMG_C.AT] || ''),
        sugg: _imgSugg_(r[IMG_C.SUGG]) });
    });
  }
  return _json({ success: true, items: items });
}

// From the Scanner review screen. body: { barcode, status: 'Validé'|'Faux'|…, link: 'https://…' (optional), by }
// A link is downloaded right away into the Drive folder and marked Validé.
function _setProductImage(d) {
  const key = _imgKey(d.barcode);
  if (!key) return _json({ success: false, error: 'Code-barres manquant' });
  const sh = _imgSheet_();
  const last = sh.getLastRow();
  if (last < 2) return _json({ success: false, error: 'Liste vide' });
  const data = sh.getRange(2, 1, last - 1, IMG_HEADERS.length).getValues();
  const i = data.findIndex(r => _imgKey(r[IMG_C.BC]) === key);
  if (i < 0) return _json({ success: false, error: 'Code-barres introuvable dans la liste' });
  const row = data[i];
  const link = String(d.link || '').trim();
  if (link) {
    if (!/^https?:\/\//i.test(link)) return _json({ success: false, error: 'Lien invalide (doit commencer par http)' });
    const alt = String(d.alt || '').trim(); // Brave's copy of the same picture, used if the shop blocks the download
    const blob = _imgDownload_(link) || (/^https:\/\//i.test(alt) ? _imgDownload_(alt) : null);
    if (!blob) return _json({ success: false, error: "Ce lien ne donne pas une image (essayez « Copier l'adresse de l'image »)" });
    row[IMG_C.URL] = _imgSave_(blob, String(row[IMG_C.BC]).replace(/\D/g, ''), _imgFolder_(), row[IMG_C.URL]);
    row[IMG_C.ORIG] = link; row[IMG_C.MANUAL] = link; row[IMG_C.SOURCE] = 'Manuel';
    row[IMG_C.STATUS] = IMG_ST.VALID;
  } else {
    const st = String(d.status || '');
    if ([IMG_ST.VALID, IMG_ST.WRONG, IMG_ST.FOUND].indexOf(st) < 0) return _json({ success: false, error: 'Statut non autorisé' });
    if ((st === IMG_ST.VALID || st === IMG_ST.FOUND) && !row[IMG_C.URL]) return _json({ success: false, error: 'Pas de photo à valider' });
    row[IMG_C.STATUS] = st;
  }
  if (d.by) row[IMG_C.SOURCE] = String(row[IMG_C.SOURCE] || '').replace(/ · par .*$/, '') + ' · par ' + String(d.by).slice(0, 30);
  _imgWriteRow_(sh, i + 2, row);
  return _json({ success: true, item: { barcode: String(row[IMG_C.BC]).trim(), name: String(row[IMG_C.NAME] || ''),
    status: row[IMG_C.STATUS], source: row[IMG_C.SOURCE], url: row[IMG_C.URL], at: row[IMG_C.AT], sugg: _imgSugg_(row[IMG_C.SUGG]) } });
}

// ══════════════════════════════════════════════════════════
//  NAME SUGGESTIONS FOR UNKNOWN BARCODES
//  When the till (or the price-update session) scans a real barcode that
//  isn't in Product Master, the Scanner calls ?action=lookupBarcodeName and
//  shows up to 3 names to tap instead of typing one.
//  Order: Open Beauty/Food Facts + UPCitemdb (free, in parallel), then — only
//  if they found nothing — one Brave WEB search on the digits (good for Thai
//  products listed on Shopee/Lazada). Answers are cached 6 h per barcode.
//  Our own / made-up codes are never searched.
// ══════════════════════════════════════════════════════════
function _lookupBarcodeName(raw) {
  return _json(Object.assign({ success: true }, _nameCandidates_(raw)));
}
// Monthly Brave search cap: Script property BRAVE_PER_MONTH, else 900 (≈ the free $5 credit).
function _braveCap_() {
  return Number(PropertiesService.getScriptProperties().getProperty('BRAVE_PER_MONTH')) || IMG_BRAVE.PER_MONTH;
}
// { candidates: [{name, source, thumb}], skipped?, cached?, limit? }
function _nameCandidates_(raw) {
  const bc = String(raw || '').replace(/\D/g, '');
  if (bc.length < 8) return { candidates: [], skipped: 'court' };
  if (_imgIsInternal(bc)) return { candidates: [], skipped: 'interne' };
  const cache = CacheService.getScriptCache();
  const ck = 'nm:' + _imgKey(bc);
  const hit = cache.get(ck);
  if (hit) { try { return Object.assign({ cached: true }, JSON.parse(hit)); } catch (e) {} }

  const codes = [bc.length > 8 && bc.length < 12 ? bc.padStart(12, '0') : bc];
  if (codes[0].length === 12) codes.push('0' + codes[0]);
  if (bc.length === 13 && bc[0] === '0') codes.push(bc.slice(1));
  const reqs = [];
  ['world.openbeautyfacts.org', 'world.openfoodfacts.org'].forEach(h => codes.forEach(c => reqs.push({
    url: 'https://' + h + '/api/v2/product/' + c + '.json?fields=product_name,product_name_fr,product_name_en,brands,quantity',
    muteHttpExceptions: true, headers: { 'User-Agent': 'Boutique1981-POS/1.0 (1981estvideo@gmail.com)' } })));
  reqs.push({ url: 'https://api.upcitemdb.com/prod/trial/lookup?upc=' + bc, muteHttpExceptions: true });
  const cands = [];
  const add = (name, source, thumb) => {
    name = _nmClean_(name, bc);
    if (name.length < 3) return;
    if (cands.some(c => c.name.toLowerCase() === name.toLowerCase())) return;
    cands.push({ name: name, source: source, thumb: thumb || '' });
  };
  let res = [];
  try { res = UrlFetchApp.fetchAll(reqs); } catch (e) {}
  res.forEach((r, i) => {
    if (r.getResponseCode() !== 200) return;
    try {
      const d = JSON.parse(r.getContentText());
      if (i < reqs.length - 1) {
        const p = d && d.status === 1 && d.product;
        if (!p) return;
        const nm = p.product_name_fr || p.product_name || p.product_name_en || '';
        const brand = String(p.brands || '').split(',')[0].trim();
        if (nm) add((brand && nm.toLowerCase().indexOf(brand.toLowerCase()) < 0 ? brand + ' ' : '') + nm + (p.quantity ? ' ' + p.quantity : ''),
          i < codes.length ? 'Open Beauty Facts' : 'Open Food Facts');
      } else {
        (d.items || []).slice(0, 2).forEach(it => add(it.title, 'UPCitemdb', (it.images || [])[0]));
      }
    } catch (e) {}
  });

  const key = PropertiesService.getScriptProperties().getProperty('BRAVE_API_KEY');
  let braveLimit = false;
  if (!cands.length && key && !(braveLimit = !_braveBudgetOk_())) {
    try {
      const r = UrlFetchApp.fetch('https://api.search.brave.com/res/v1/web/search?count=8&safesearch=strict&q=' + encodeURIComponent(bc),
        { muteHttpExceptions: true, headers: { 'X-Subscription-Token': key, 'Accept': 'application/json' } });
      if (r.getResponseCode() === 200) {
        const d = JSON.parse(r.getContentText());
        ((d.web && d.web.results) || []).forEach(x => {
          if (cands.length >= 3) return;
          const site = String((x.meta_url && x.meta_url.hostname) || '').replace(/^www\./, '');
          if (/upc|barcode|ean|gtin|codigo/i.test(site)) { // barcode-lookup sites: their title is usually "<code> | <name>"
            add(String(x.title || '').replace(/^.*?\|\s*/, ''), site, x.thumbnail && x.thumbnail.src);
          } else {
            add(x.title, site, x.thumbnail && x.thumbnail.src);
          }
        });
      } else {
        Logger.log('Brave web ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
      }
    } catch (e) {}
  }
  const out = { candidates: cands.slice(0, 3) };
  if (braveLimit && !cands.length) return { candidates: [], limit: true }; // not cached: retry next month
  try { cache.put(ck, JSON.stringify(out), 21600); } catch (e) {}
  return out;
}

// Shared monthly Brave counter (same one the photo suggestions use).
function _braveBudgetOk_() {
  const props = PropertiesService.getScriptProperties();
  const month = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM');
  if (props.getProperty('IMG_BRAVE_MONTH') !== month) { props.setProperty('IMG_BRAVE_MONTH', month); props.setProperty('IMG_BRAVE_COUNT', '0'); }
  const used = Number(props.getProperty('IMG_BRAVE_COUNT')) || 0;
  if (used >= _braveCap_()) return false;
  props.setProperty('IMG_BRAVE_COUNT', String(used + 1));
  return true;
}

// Shop page title → product name: drop "Amazon.com:", "| Walmart", "- Shopee Thailand", the barcode itself, HTML entities.
function _nmClean_(t, bc) {
  let s = String(t || '').replace(/<[^>]+>/g, ' ').replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'");
  s = s.replace(/^\s*(amazon\.[a-z.]+|walmart\.com|target)\s*:\s*/i, '');
  s = s.split(/\s+[|｜]\s+/)[0];
  const tail = /\s+[-–—:]\s+(walmart(\.com)?|target|amazon(\.[a-z.]+)?|sam'?s club|shopee[^-–—]*|lazada[^-–—]*|ebay|kroger|cvs|walgreens|ulta|sephora|instacart|jumia[^-–—]*|big ?c|tops|watsons[^-–—]*|grocery[^:]*|beauty[^:]*|health[^:]*|household[^:]*|baby[^:]*|pantry[^:]*)\s*$/i;
  for (let k = 0; k < 3 && tail.test(s); k++) s = s.replace(tail, '');
  if (bc) s = s.replace(new RegExp('0*' + bc.replace(/^0+/, ''), 'g'), ' ');
  s = s.replace(/\b(UPC|EAN|GTIN|barcode|code-barres)\b[:#]?/gi, ' ').replace(/^\s*(lookup|search|database|db)\s*[-–—:|]\s*/i, '');
  return s.replace(/\s+/g, ' ').replace(/^[\s\-–—:|,.]+|[\s\-–—:|,]+$/g, '').trim().slice(0, 90);
}


// ══════════════════════════════════════════════════════════
//  FULL-STORE SCAN → NEW PRODUCTS + STOCK (Balkuy)
//  1. Paste every scanned barcode (duplicates included — each scan = 1 item)
//     into column A of the "Codes scannés" tab.
//  2. processScannedCodes() (Scanner Admin button, or run it here) counts
//     them and fills "Nouveaux codes":
//       • Connu   = already in Product Master → Balkuy stock is set to the
//                   scanned count (updateCounts + one "recount" movement,
//                   logged in the Stock Ledger) — "Stock appliqué".
//       • Nouveau = unknown real barcode → names + photos are looked up in
//                   the background (every 15 min, with runImageLookup), then
//                   validated in the Scanner (Admin → 🆕 Nouveaux produits):
//                   ✓ Ajouter puts it in Product Master with its scanned
//                   quantity as Balkuy stock, its price and its photo.
//       • Interne / Invalide = our own code or a code that fails the check
//                   digit (often a mis-scan): no search, name typed by hand.
//  Re-running after more scanning is safe: counts are recomputed, a known
//  product's stock is set again only if its count changed, and products
//  already added are left alone.
// ══════════════════════════════════════════════════════════
const NC = { PASTE: 'Codes scannés', SHEET: 'Nouveaux codes', LOC: 'Store 1', CHUNK: 120 };
// 24 Oct: columns 12-13 = store + zone, so one tool handles Balkuy, Ouaga 2000
// and the warehouse, and stock lands in the right zone (Stock1981).
const NC_HEADERS = ['Barcode', 'Qté scannée', 'Type', 'Produit', 'Stock avant', 'Statut', 'Suggestions noms', 'Suggestions photos', 'Prix', 'Mis à jour', 'Note', 'Magasin', 'Zone'];
const NC_C = { BC: 0, QTY: 1, TYPE: 2, NAME: 3, BEFORE: 4, STATUS: 5, SN: 6, SP: 7, PRICE: 8, AT: 9, NOTE: 10, LOC: 11, ZONE: 12 };
const NC_ST = { APPLY: 'À appliquer', APPLIED: 'Stock appliqué', TODO: 'À chercher', SUGG: 'Suggestions', NONE: 'Introuvable', ADDED: 'Ajouté', SKIP: 'Ignoré', MANUAL: 'À nommer' };
const NC_PASTE_HEADERS = ['Barcode (collez tous les scans ici, doublons compris)', 'Magasin (Balkuy / Ouaga 2000 / Entrepôt — vide = Balkuy)', 'Zone (ex. M2, Cashier — facultatif)', 'Nom (facultatif)', 'Prix CFA (facultatif)', 'Quantité (vide = 1 par ligne)'];
const NC_LOCS = { 'balkuy': 'Store 1', 'store 1': 'Store 1', 'ouaga 2000': 'Store 2', 'ouaga2000': 'Store 2', 'store 2': 'Store 2', 'entrepot': 'Warehouse', 'entrepôt': 'Warehouse', 'warehouse': 'Warehouse' };
const NC_LOC_NAMES = { 'Store 1': 'Balkuy', 'Store 2': 'Ouaga 2000', 'Warehouse': 'Entrepôt' };
const NC_STOCK_FIELD = { 'Store 1': 's1', 'Store 2': 's2', 'Warehouse': 'wh' };
function _ncLoc_(v) { const t = String(v || '').trim().toLowerCase(); return t ? (NC_LOCS[t] || null) : NC.LOC; }
function _ncRowLoc_(r) { return String(r[NC_C.LOC] || '') || NC.LOC; }

function _ncSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let paste = ss.getSheetByName(NC.PASTE);
  if (!paste) {
    paste = ss.insertSheet(NC.PASTE);
    paste.getRange('A:A').setNumberFormat('@');
    paste.setColumnWidth(1, 320); paste.setFrozenRows(1);
  }
  if (!String(paste.getRange('B1').getValue() || '').trim()) {
    paste.getRange(1, 1, 1, NC_PASTE_HEADERS.length).setValues([NC_PASTE_HEADERS]).setFontWeight('bold');
  }
  let sh = ss.getSheetByName(NC.SHEET);
  if (!sh) {
    sh = ss.insertSheet(NC.SHEET);
    sh.setFrozenRows(1); sh.getRange('A:A').setNumberFormat('@');
    sh.setColumnWidth(4, 280); sh.setColumnWidth(7, 160); sh.setColumnWidth(8, 160);
  }
  sh.getRange(1, 1, 1, NC_HEADERS.length).setValues([NC_HEADERS]).setFontWeight('bold');
  return { paste: paste, sh: sh };
}
function _ncNow_() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm'); }
function _invGet_(params) {
  const url = CONFIG.INVENTORY_SYNC_URL + (params ? (CONFIG.INVENTORY_SYNC_URL.indexOf('?') >= 0 ? '&' : '?') + params : '');
  const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
  return JSON.parse(r.getContentText());
}
function _invPost_(body) {
  const r = UrlFetchApp.fetch(CONFIG.INVENTORY_SYNC_URL, { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify(Object.assign({ key: CONFIG.INVENTORY_SYNC_KEY }, body)) });
  return JSON.parse(r.getContentText());
}
// key → {barcode, name} for everything in Product Master
function _ncMaster_() {
  const d = _invGet_('');
  const m = {};
  (d.products || []).forEach(p => { const k = _imgKey(p.barcode); if (k) m[k] = { barcode: String(p.barcode), name: String(p.name || '') }; });
  return m;
}
// Zone stock (Stock1981) for one store + zone; failures are only noted.
function _ncZoneStock_(loc, zone, items, by) {
  if (!zone || !items.length) return '';
  try {
    const z = _invPost_({ action: 'importZoneStock', location: loc, zone: zone, by: by || 'Nouveaux produits', items: items });
    return (z && z.success) ? '' : 'zone non mise à jour (' + ((z && z.error) || '?') + ')';
  } catch (e) { return 'zone non mise à jour'; }
}

function processScannedCodes() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, error: 'Une autre opération est en cours, réessayez dans une minute.' };
  try {
    const { paste, sh } = _ncSheets_();
    // 1) count the scans, per barcode + store + zone
    const groups = {};
    let scans = 0, ignored = 0;
    if (paste.getLastRow() >= 2) {
      paste.getRange(2, 1, paste.getLastRow() - 1, 6).getValues().forEach(r => {
        const v = String(r[0] == null ? '' : r[0]).trim();
        if (!v) return;
        const d = v.replace(/\s/g, '');
        if (!/^\d+$/.test(d)) { ignored++; return; }
        const k = _imgKey(d);
        const loc = _ncLoc_(r[1]);
        if (!k || !loc) { ignored++; return; }
        const zone = String(r[2] || '').trim();
        const gk = k + '|' + loc + '|' + zone;
        const g = groups[gk] || (groups[gk] = { k: k, loc: loc, zone: zone, qty: 0, raw: '', name: '', price: '' });
        const n = Math.max(1, Math.round(Number(r[5]) || 1));
        g.qty += n; scans += n;
        const dd = (d.length > 8 && d.length < 12) ? d.padStart(12, '0') : d; // UPC whose leading zero(s) got lost
        if (!g.raw || dd.length > g.raw.length) g.raw = dd;
        if (!g.name && String(r[3] || '').trim()) g.name = String(r[3]).replace(/\s+/g, ' ').trim();
        if (!g.price && Number(r[4]) > 0) g.price = Math.round(Number(r[4]));
      });
    }
    if (!scans) return { success: false, error: 'Aucun code dans l’onglet « ' + NC.PASTE + ' » (colonne A, à partir de la ligne 2).' };
    // 2) what the system already knows
    let master = _ncMaster_();
    const stockRes = _invGet_('action=getStockSummary');
    const stock = (stockRes && stockRes.stock) || {};
    // 3) upsert rows
    const last = sh.getLastRow();
    const data = last >= 2 ? sh.getRange(2, 1, last - 1, NC_HEADERS.length).getValues() : [];
    const rowOf = {};
    data.forEach((r, i) => { const k = _imgKey(r[NC_C.BC]); if (k) rowOf[k + '|' + _ncRowLoc_(r) + '|' + String(r[NC_C.ZONE] || '')] = i; });
    const now = _ncNow_();
    const prefilled = [];
    Object.keys(groups).forEach(gk => {
      const g = groups[gk], k = g.k, field = NC_STOCK_FIELD[g.loc];
      let i = rowOf[gk];
      if (i === undefined) {
        const m = master[k];
        const st = stock[k] || {};
        const type = m ? 'Connu' : (_imgIsInternal(g.raw) ? (k.indexOf(IMG.INTERNAL_PREFIX) === 0 ? 'Interne' : 'Invalide') : 'Nouveau');
        const row = new Array(NC_HEADERS.length).fill('');
        row[NC_C.BC] = m ? m.barcode : g.raw; row[NC_C.QTY] = g.qty; row[NC_C.TYPE] = type;
        row[NC_C.NAME] = m ? m.name : g.name; row[NC_C.BEFORE] = m ? (st[field] == null ? '' : st[field]) : '';
        row[NC_C.STATUS] = type === 'Connu' ? NC_ST.APPLY : type === 'Nouveau' ? NC_ST.TODO : NC_ST.MANUAL;
        row[NC_C.PRICE] = g.price || ''; row[NC_C.AT] = now; row[NC_C.LOC] = g.loc; row[NC_C.ZONE] = g.zone;
        data.push(row); i = rowOf[gk] = data.length - 1;
      } else {
        const r = data[i];
        if (Number(r[NC_C.QTY]) !== g.qty) {
          r[NC_C.QTY] = g.qty; r[NC_C.AT] = now;
          if (r[NC_C.TYPE] === 'Connu') r[NC_C.STATUS] = NC_ST.APPLY;
          else if (r[NC_C.STATUS] === NC_ST.ADDED) r[NC_C.NOTE] = 'Recompté après ajout : ' + g.qty + ' (stock non modifié)';
        }
        if (r[NC_C.TYPE] !== 'Connu' && master[k] && r[NC_C.STATUS] !== NC_ST.ADDED) { // added meanwhile by someone else
          r[NC_C.TYPE] = 'Connu'; r[NC_C.NAME] = master[k].name; r[NC_C.STATUS] = NC_ST.APPLY;
        }
      }
      const r = data[i];
      if (r[NC_C.TYPE] !== 'Connu' && r[NC_C.STATUS] !== NC_ST.ADDED && r[NC_C.STATUS] !== NC_ST.SKIP && g.name && g.price) prefilled.push(i);
    });
    // 4) products whose name + price came with the scans: add them straight away
    let autoAdded = 0;
    prefilled.forEach(i => {
      const r = data[i], k = _imgKey(r[NC_C.BC]);
      if (master[k]) { r[NC_C.TYPE] = 'Connu'; r[NC_C.STATUS] = NC_ST.APPLY; r[NC_C.NAME] = master[k].name; return; }
      const res = _ncAddProduct_(r, groups[k + '|' + _ncRowLoc_(r) + '|' + String(r[NC_C.ZONE] || '')].name, r[NC_C.PRICE], 'Import des scans', master);
      if (res.ok) { autoAdded++; master[k] = { barcode: String(r[NC_C.BC]), name: r[NC_C.NAME] }; }
      else r[NC_C.NOTE] = res.error;
    });
    // 5) set store (and zone) stock of known products to the scanned count
    const toApply = data.map((r, i) => i).filter(i => data[i][NC_C.TYPE] === 'Connu' && data[i][NC_C.STATUS] === NC_ST.APPLY);
    const byLoc = {};
    toApply.forEach(i => { const l = _ncRowLoc_(data[i]); (byLoc[l] = byLoc[l] || []).push(i); });
    let applied = 0, failed = 0;
    Object.keys(byLoc).forEach(loc => {
      const field = NC_STOCK_FIELD[loc], list = byLoc[loc];
      // one product listed in two zones of the same store: the store total is the sum
      const total = {};
      list.forEach(i => { const k = _imgKey(data[i][NC_C.BC]); total[k] = (total[k] || 0) + Number(data[i][NC_C.QTY]); });
      for (let c = 0; c < list.length; c += NC.CHUNK) {
        const part = list.slice(c, c + NC.CHUNK);
        try {
          const seen = {};
          const counts = [], moves = [];
          part.forEach(i => {
            const k = _imgKey(data[i][NC_C.BC]);
            const before = stock[k] && stock[k][field] != null ? Number(stock[k][field]) : 0;
            if (data[i][NC_C.BEFORE] === '') data[i][NC_C.BEFORE] = before;
            if (seen[k]) return; seen[k] = true;
            counts.push({ barcode: String(data[i][NC_C.BC]), count: total[k] });
            const delta = total[k] - before;
            if (delta) moves.push({ barcode: String(data[i][NC_C.BC]), delta: delta, type: 'recount', note: 'Scan complet ' + NC_LOC_NAMES[loc] });
          });
          const u = _invPost_({ action: 'updateCounts', location: loc, counts: counts });
          if (!u || u.success === false) throw new Error((u && u.error) || 'updateCounts');
          if (moves.length) {
            const a = _invPost_({ action: 'applyStockMovements', location: loc, movements: moves });
            if (!a || a.success === false) throw new Error((a && a.error) || 'applyStockMovements');
          }
          const byZone = {};
          part.forEach(i => { const z = String(data[i][NC_C.ZONE] || ''); if (z) (byZone[z] = byZone[z] || []).push({ barcode: String(data[i][NC_C.BC]), name: String(data[i][NC_C.NAME] || ''), qty: Number(data[i][NC_C.QTY]) }); });
          const zoneNotes = Object.keys(byZone).map(z => _ncZoneStock_(loc, z, byZone[z], 'Import des scans')).filter(Boolean);
          part.forEach(i => {
            const k = _imgKey(data[i][NC_C.BC]);
            stock[k] = Object.assign({}, stock[k]); stock[k][field] = total[k];
            data[i][NC_C.STATUS] = NC_ST.APPLIED; data[i][NC_C.AT] = now;
            if (zoneNotes.length) data[i][NC_C.NOTE] = zoneNotes[0];
          });
          applied += part.length;
        } catch (e) {
          failed += part.length;
          part.forEach(i => { data[i][NC_C.NOTE] = 'Échec stock : ' + String(e.message || e).slice(0, 80); });
        }
      }
    });
    if (data.length) {
      sh.getRange(2, 1, data.length, 1).setNumberFormat('@');
      sh.getRange(2, 1, data.length, NC_HEADERS.length).setValues(data);
    }
    const sum = _ncSummary_(data);
    sum.scans = scans; sum.ignored = ignored; sum.applied = applied; sum.failed = failed; sum.autoAdded = autoAdded;
    return Object.assign({ success: true }, sum);
  } finally { lock.releaseLock(); }
}
function _ncSummary_(data) {
  const s = { codes: data.length, known: 0, fresh: 0, toName: 0, added: 0, waiting: 0 };
  data.forEach(r => {
    if (r[NC_C.TYPE] === 'Connu') s.known++;
    else if (r[NC_C.STATUS] === NC_ST.ADDED) s.added++;
    else if (r[NC_C.STATUS] === NC_ST.TODO) s.waiting++;
    else if (r[NC_C.STATUS] !== NC_ST.SKIP) s.toName++;
    if (r[NC_C.TYPE] !== 'Connu') s.fresh++;
  });
  return s;
}

// Background step (called from runImageLookup): names, then photos, for "À chercher" rows.
function _ncLookupBatch_(until) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(NC.SHEET);
  if (!sh || sh.getLastRow() < 2) return;
  const key = PropertiesService.getScriptProperties().getProperty('BRAVE_API_KEY');
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, NC_HEADERS.length).getValues();
  let lastBrave = 0;
  for (let i = 0; i < data.length && Date.now() < until; i++) {
    const r = data[i];
    if (r[NC_C.STATUS] !== NC_ST.TODO) continue;
    const wait = IMG_BRAVE.GAP_MS - (Date.now() - lastBrave); if (wait > 0) Utilities.sleep(wait);
    const nm = _nameCandidates_(r[NC_C.BC]);
    lastBrave = Date.now();
    if (nm.limit) break; // monthly cap reached: carry on next month
    const names = nm.candidates || [];
    let photos = [];
    if (names.length && key) {
      Utilities.sleep(IMG_BRAVE.GAP_MS);
      const ph = _imgBraveSearch_(key, names[0].name);
      lastBrave = Date.now();
      if (ph.list) photos = ph.list;
    }
    r[NC_C.SN] = JSON.stringify(names);
    r[NC_C.SP] = JSON.stringify(photos);
    r[NC_C.STATUS] = names.length ? NC_ST.SUGG : NC_ST.NONE;
    r[NC_C.AT] = _ncNow_();
    sh.getRange(i + 2, 1, 1, NC_HEADERS.length).setValues([r]);
  }
}

function _getNewCodes() {
  const { sh } = _ncSheets_();
  const last = sh.getLastRow();
  const data = last >= 2 ? sh.getRange(2, 1, last - 1, NC_HEADERS.length).getValues() : [];
  const items = data.filter(r => String(r[NC_C.BC]).trim()).map(r => ({
    barcode: String(r[NC_C.BC]).trim(), qty: Number(r[NC_C.QTY]) || 0, type: String(r[NC_C.TYPE]), name: String(r[NC_C.NAME] || ''),
    before: r[NC_C.BEFORE] === '' ? null : Number(r[NC_C.BEFORE]), status: String(r[NC_C.STATUS]),
    sugg: _imgSugg_(r[NC_C.SN]), photos: _imgSugg_(r[NC_C.SP]), price: r[NC_C.PRICE] === '' ? null : Number(r[NC_C.PRICE]),
    at: String(r[NC_C.AT] || ''), note: String(r[NC_C.NOTE] || ''),
    location: _ncRowLoc_(r), store: NC_LOC_NAMES[_ncRowLoc_(r)] || _ncRowLoc_(r), zone: String(r[NC_C.ZONE] || '') }));
  const pasteSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NC.PASTE);
  return _json(Object.assign({ success: true, items: items, pasted: pasteSh ? Math.max(0, pasteSh.getLastRow() - 1) : 0 }, _ncSummary_(data)));
}

// The row for this barcode that still needs work (a code scanned in two stores has two rows).
function _ncFind_(barcode) {
  const { sh } = _ncSheets_();
  const k = _imgKey(barcode);
  const last = sh.getLastRow();
  if (!k || last < 2) return null;
  const data = sh.getRange(2, 1, last - 1, NC_HEADERS.length).getValues();
  let i = data.findIndex(r => _imgKey(r[NC_C.BC]) === k && r[NC_C.STATUS] !== NC_ST.ADDED && r[NC_C.TYPE] !== 'Connu');
  if (i < 0) i = data.findIndex(r => _imgKey(r[NC_C.BC]) === k);
  return i < 0 ? null : { sh: sh, i: i, row: data[i] };
}
// Adds one scanned product to Product Master with its store + zone stock and price.
// Updates `row` in place; the caller writes it back. Needs the script lock held.
function _ncAddProduct_(row, name, price, by, master) {
  name = String(name || '').replace(/\s+/g, ' ').trim();
  if (name.length < 3) return { ok: false, error: 'Nom trop court' };
  master = master || _ncMaster_();
  const k = _imgKey(row[NC_C.BC]);
  if (master[k]) return { ok: false, error: 'Ce code est déjà dans la liste : ' + master[k].name };
  const norm = s => String(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
  const same = Object.keys(master).find(x => norm(master[x].name) === norm(name));
  if (same) return { ok: false, error: 'Ce nom existe déjà (code ' + master[same].barcode + ') — ajoutez la teinte / couleur pour les distinguer.' };
  const bc = String(row[NC_C.BC]).trim();
  const qty = Number(row[NC_C.QTY]) || 0;
  const loc = _ncRowLoc_(row), zone = String(row[NC_C.ZONE] || '');
  const a = _invPost_({ action: 'addNewProductWithStock', name: name, barcode: bc, category: 'Cosmetics', location: loc, qty: qty });
  if (!a || a.success === false || (a.masterResult && a.masterResult.success === false))
    return { ok: false, error: 'Ajout refusé : ' + ((a && (a.error || (a.masterResult && a.masterResult.error))) || 'erreur') };
  const notes = [];
  price = Math.round(Number(price));
  if (price > 0) {
    try {
      const pr = _invPost_({ action: 'setStorePrice', mode: 'admin', barcode: bc, prices: { 'Store 1': price, 'Store 2': price }, by: by || 'Nouveaux produits' });
      if (!pr || pr.success === false) notes.push('prix non enregistré (' + ((pr && pr.error) || '?') + ')');
    } catch (e) { notes.push('prix non enregistré'); }
  }
  const zn = _ncZoneStock_(loc, zone, [{ barcode: bc, name: name, qty: qty }], by);
  if (zn) notes.push(zn);
  row[NC_C.NAME] = name; row[NC_C.STATUS] = NC_ST.ADDED; row[NC_C.TYPE] = 'Nouveau';
  row[NC_C.PRICE] = price > 0 ? price : ''; row[NC_C.AT] = _ncNow_();
  row[NC_C.NOTE] = [by ? 'par ' + String(by).slice(0, 30) : ''].concat(notes).filter(Boolean).join(' · ');
  return { ok: true, notes: notes };
}
// body: { barcode, name, price, photo: {img, thumb} | null, by }
function _addNewCode(d) {
  const f = _ncFind_(d.barcode);
  if (!f) return _json({ success: false, error: 'Code introuvable dans « Nouveaux codes »' });
  const row = f.row;
  if (row[NC_C.STATUS] === NC_ST.ADDED) return _json({ success: false, error: 'Déjà ajouté' });
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return _json({ success: false, error: 'Occupé, réessayez' });
  try {
    const res = _ncAddProduct_(row, d.name, d.price, d.by, null);
    if (!res.ok) return _json({ success: false, error: res.error });
    let photoNote = '';
    if (d.photo && (d.photo.img || d.photo.thumb)) {
      try { photoNote = _ncSavePhoto_(String(row[NC_C.BC]).trim(), row[NC_C.NAME], d.photo) ? '' : 'photo illisible'; } catch (e) { photoNote = 'photo non enregistrée'; }
    }
    if (photoNote) row[NC_C.NOTE] = [row[NC_C.NOTE], photoNote].filter(Boolean).join(' · ');
    f.sh.getRange(f.i + 2, 1, 1, NC_HEADERS.length).setValues([row]);
    return _json({ success: true, warning: res.notes.concat(photoNote ? [photoNote] : []).join(', '), store: NC_LOC_NAMES[_ncRowLoc_(row)], zone: String(row[NC_C.ZONE] || '') });
  } finally { lock.releaseLock(); }
}
// Chosen photo → Drive + "Product Images" row marked Validé (the scan card shows it right away).
function _ncSavePhoto_(bc, name, ph) {
  const blob = _imgDownload_(ph.img || '') || (ph.thumb ? _imgDownload_(ph.thumb) : null);
  if (!blob) return false;
  const sh = _imgSheet_();
  const k = _imgKey(bc);
  const last = sh.getLastRow();
  const data = last >= 2 ? sh.getRange(2, 1, last - 1, IMG_HEADERS.length).getValues() : [];
  let i = data.findIndex(r => _imgKey(r[IMG_C.BC]) === k);
  const row = i >= 0 ? data[i] : new Array(IMG_HEADERS.length).fill('');
  row[IMG_C.BC] = bc; row[IMG_C.NAME] = name; row[IMG_C.TRIES] = row[IMG_C.TRIES] || 0;
  row[IMG_C.URL] = _imgSave_(blob, bc, _imgFolder_(), row[IMG_C.URL]);
  row[IMG_C.ORIG] = ph.img || ph.thumb; row[IMG_C.STATUS] = IMG_ST.VALID; row[IMG_C.SOURCE] = 'Nouveaux produits';
  const rn = i >= 0 ? i + 2 : sh.getLastRow() + 1;
  sh.getRange(rn, 1).setNumberFormat('@');
  _imgWriteRow_(sh, rn, row);
  return true;
}
function _skipNewCode(d) {
  const f = _ncFind_(d.barcode);
  if (!f) return _json({ success: false, error: 'Code introuvable' });
  f.row[NC_C.STATUS] = d.undo ? (_imgSugg_(f.row[NC_C.SN]).length ? NC_ST.SUGG : (f.row[NC_C.TYPE] === 'Nouveau' ? NC_ST.NONE : NC_ST.MANUAL)) : NC_ST.SKIP;
  f.row[NC_C.AT] = _ncNow_();
  f.sh.getRange(f.i + 2, 1, 1, NC_HEADERS.length).setValues([f.row]);
  return _json({ success: true, status: f.row[NC_C.STATUS] });
}
