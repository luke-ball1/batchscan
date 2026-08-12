/**
 * CRATE TRACKING SYSTEM — Company Bakery (v2: physical crate tags)
 * ----------------------------------------------------------
 * Each physical crate gets ONE permanent QR sticker, generated once in a
 * batch and printed. The sticker's ID never changes for the life of the
 * crate. Route/customer are NOT tied to the crate — they're captured at
 * the moment a crate is scanned OUT, and every later event (DELIVERED, IN)
 * automatically inherits that same route/customer until the next OUT.
 *
 * TABS (created by setup()):
 *   Registry    - one row per physical crate ID ever generated
 *   Ledger      - one row per scan event (append-only log)
 *   Dashboard   - computed summary, refreshed on demand
 *   PrintQueue  - holds the most recently generated batch, read by the
 *                 print sheet
 *
 * DEPLOY AS WEB APP (Deploy > New deployment > Web app, execute as Me,
 * access Anyone with the link / your domain):
 *   Scan page:   <url>/exec?role=driver   (or packer / ops)
 *   Print sheet: <url>/exec?mode=print
 * ----------------------------------------------------------
 */

const SHEET_NAMES = {
  REGISTRY: 'Registry',
  LEDGER: 'Ledger',
  DASHBOARD: 'Dashboard',
  PRINT_QUEUE: 'PrintQueue'
};

const EVENT_TYPES = ['OUT', 'DELIVERED', 'IN', 'ISSUE'];

// Fixed route list shown as a dropdown when scanning a crate OUT.
// Edit this list as your routes change.
const FIXED_ROUTES = ['M1', 'M2', 'M3', 'M4', 'M5', 'N1', 'N2', 'N3', 'SANDWICH', 'BMB', 'INTERNAL', 'COLLECT'];

// ---------- SETUP ----------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Crate Tracking')
    .addItem('Run setup (first time only)', 'setup')
    .addItem('Generate new crate batch...', 'promptGenerateBatch')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const registry = getOrCreateSheet_(ss, SHEET_NAMES.REGISTRY);
  if (registry.getLastRow() === 0) {
    registry.appendRow(['Crate ID', 'Date Created', 'Status', 'Notes']);
    registry.setFrozenRows(1);
  }

  const ledger = getOrCreateSheet_(ss, SHEET_NAMES.LEDGER);
  if (ledger.getLastRow() === 0) {
    ledger.appendRow(['Timestamp', 'Crate ID', 'Event Type', 'Route', 'Customer', 'Scanned By', 'Notes']);
    ledger.setFrozenRows(1);
  }

  const dashboard = getOrCreateSheet_(ss, SHEET_NAMES.DASHBOARD);
  dashboard.clear();
  dashboard.appendRow(['Crate ID', 'Status', 'Last Event', 'Route', 'Customer', 'Last Event Time', 'Days Since', 'Trip Count', 'Flag']);
  dashboard.setFrozenRows(1);

  const printQueue = getOrCreateSheet_(ss, SHEET_NAMES.PRINT_QUEUE);
  if (printQueue.getLastRow() === 0) {
    printQueue.appendRow(['Crate ID']);
    printQueue.setFrozenRows(1);
  }

  SpreadsheetApp.getUi().alert('Setup complete. Registry, Ledger, Dashboard and PrintQueue tabs are ready.');
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

// ---------- CRATE BATCH GENERATION (physical stickers) ----------

function promptGenerateBatch() {
  const ui = SpreadsheetApp.getUi();
  const countResp = ui.prompt('Generate crate batch', 'How many new crate IDs do you want to generate?', ui.ButtonSet.OK_CANCEL);
  if (countResp.getSelectedButton() !== ui.Button.OK) return;
  const count = parseInt(countResp.getResponseText(), 10);
  if (!count || count <= 0) {
    ui.alert('Enter a valid number.');
    return;
  }
  const ids = generateCrateBatch(count);
  ui.alert('Generated ' + ids.length + ' crate IDs (' + ids[0] + ' to ' + ids[ids.length - 1] + ').\n\n'
    + 'Open the web app with ?mode=print to print the sticker sheet for this batch.');
}

/**
 * Creates `count` new sequential crate IDs (e.g. CB-0001, CB-0002, ...),
 * continuing from the highest existing number in the Registry. Writes them
 * to Registry (Status = Active) and overwrites PrintQueue with just this
 * batch, ready for the print sheet to pick up.
 */
function generateCrateBatch(count, prefix) {
  prefix = prefix || 'CB';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getOrCreateSheet_(ss, SHEET_NAMES.REGISTRY);

  const existingIds = registry.getLastRow() > 1
    ? registry.getRange(2, 1, registry.getLastRow() - 1, 1).getValues().flat()
    : [];
  let maxNum = 0;
  existingIds.forEach(id => {
    const match = String(id).match(new RegExp('^' + prefix + '-(\\d+)$'));
    if (match) maxNum = Math.max(maxNum, parseInt(match[1], 10));
  });

  const now = new Date();
  const newRows = [];
  const newIds = [];
  for (let i = 1; i <= count; i++) {
    const num = maxNum + i;
    const id = prefix + '-' + String(num).padStart(4, '0');
    newIds.push(id);
    newRows.push([id, now, 'Active', '']);
  }

  registry.getRange(registry.getLastRow() + 1, 1, newRows.length, 4).setValues(newRows);

  // Overwrite PrintQueue with just this batch
  const printQueue = getOrCreateSheet_(ss, SHEET_NAMES.PRINT_QUEUE);
  printQueue.clear();
  printQueue.appendRow(['Crate ID']);
  printQueue.getRange(2, 1, newIds.length, 1).setValues(newIds.map(id => [id]));

  return newIds;
}

function getPrintQueue() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const printQueue = getOrCreateSheet_(ss, SHEET_NAMES.PRINT_QUEUE);
  if (printQueue.getLastRow() <= 1) return [];
  return printQueue.getRange(2, 1, printQueue.getLastRow() - 1, 1).getValues().flat();
}

// ---------- WEB APP ----------

function doGet(e) {
  const params = (e && e.parameter) || {};

  // --- JSON API for the externally-hosted batch scanner ---
  // Using GET for writes too (not just reads) deliberately: Apps Script web
  // apps respond via a redirect, and several browsers silently downgrade a
  // POST to a GET when following it, which would make doPost never actually
  // run. GET has no such issue, so everything goes through here instead.
  if (params.action === 'crateInfo') {
    return jsonResponse_(getCrateInfo(params.crateId));
  }
  if (params.action === 'customers') {
    return jsonResponse_(getDistinctCustomers());
  }
  if (params.action === 'logScan') {
    return jsonResponse_(logScan(params.crateId, params.eventType, params.scannedBy, params.route, params.customer, params.notes));
  }

  const mode = params.mode || 'scan';
  if (mode === 'print') {
    return HtmlService.createHtmlOutputFromFile('PrintSheet')
      .setTitle('Crate Sticker Sheet');
  }

  const template = HtmlService.createTemplateFromFile('ScanApp');
  template.crateId = params.crateId || '';
  template.routesJson = JSON.stringify(FIXED_ROUTES);
  return template.evaluate()
    .setTitle('Crate Scan — Company Bakery')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------- SERVER FUNCTIONS CALLED FROM ScanApp.html ----------

/**
 * Returns whether a crate ID is a registered physical crate, plus its
 * current "assignment" (route/customer from its most recent OUT/DELIVERED
 * event) so the scan page can show context without asking the person to
 * re-enter it.
 */
function getCrateInfo(crateId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getOrCreateSheet_(ss, SHEET_NAMES.REGISTRY);
  const regData = registry.getDataRange().getValues();
  const isRegistered = regData.slice(1).some(row => row[0] === crateId);

  const ledger = getOrCreateSheet_(ss, SHEET_NAMES.LEDGER);
  const ledgerData = ledger.getDataRange().getValues().slice(1);
  const crateEvents = ledgerData.filter(row => row[1] === crateId);
  crateEvents.sort((a, b) => new Date(b[0]) - new Date(a[0]));
  const last = crateEvents[0];

  return {
    registered: isRegistered,
    crateId: crateId,
    lastEvent: last ? last[2] : null,
    route: last ? last[3] : '',
    customer: last ? last[4] : '',
    lastEventTime: last ? last[0] : null
  };
}

function getDistinctCustomers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = getOrCreateSheet_(ss, SHEET_NAMES.LEDGER);
  if (ledger.getLastRow() <= 1) return [];
  const customers = ledger.getRange(2, 5, ledger.getLastRow() - 1, 1).getValues().flat().filter(String);
  return [...new Set(customers)].sort();
}

/**
 * Logs a scan event.
 * - For OUT: route/customer are required (that's the point at which the
 *   assignment is decided).
 * - For DELIVERED/IN/ISSUE: route/customer are looked up automatically
 *   from the crate's most recent event, so the caller can pass empty
 *   strings for those.
 */
function logScan(crateId, eventType, scannedBy, route, customer, notes) {
  if (EVENT_TYPES.indexOf(eventType) === -1) {
    throw new Error('Unknown event type: ' + eventType);
  }

  let finalRoute = route;
  let finalCustomer = customer;
  if (eventType !== 'OUT') {
    const info = getCrateInfo(crateId);
    finalRoute = info.route;
    finalCustomer = info.customer;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = getOrCreateSheet_(ss, SHEET_NAMES.LEDGER);
  ledger.appendRow([new Date(), crateId, eventType, finalRoute || '', finalCustomer || '', scannedBy || '', notes || '']);

  return { success: true, crateId, route: finalRoute, customer: finalCustomer };
}

function logScanBatch(events) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = getOrCreateSheet_(ss, SHEET_NAMES.LEDGER);
  const rows = events.map(evt => {
    let route = evt.route, customer = evt.customer;
    if (evt.eventType !== 'OUT') {
      const info = getCrateInfo(evt.crateId);
      route = info.route;
      customer = info.customer;
    }
    return [new Date(evt.timestamp), evt.crateId, evt.eventType, route || '', customer || '', evt.scannedBy || '', evt.notes || ''];
  });
  if (rows.length) {
    ledger.getRange(ledger.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  }
  return { success: true, count: rows.length };
}

// ---------- DASHBOARD ----------

function refreshDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getOrCreateSheet_(ss, SHEET_NAMES.REGISTRY);
  const ledger = getOrCreateSheet_(ss, SHEET_NAMES.LEDGER);
  const dashboard = getOrCreateSheet_(ss, SHEET_NAMES.DASHBOARD);

  const regData = registry.getDataRange().getValues().slice(1);
  const ledgerData = ledger.getDataRange().getValues().slice(1);

  const eventsByCrate = {};
  ledgerData.forEach(row => {
    const crateId = row[1];
    if (!eventsByCrate[crateId]) eventsByCrate[crateId] = [];
    eventsByCrate[crateId].push(row);
  });

  const now = new Date();
  const outRows = regData.map(row => {
    const [crateId, , status] = row;
    const events = (eventsByCrate[crateId] || []).slice().sort((a, b) => new Date(b[0]) - new Date(a[0]));
    const last = events[0];
    const tripCount = events.filter(e => e[2] === 'OUT').length;

    if (!last) {
      return [crateId, status, '(no scans yet)', '', '', '', '', tripCount, ''];
    }
    const daysSince = Math.floor((now - new Date(last[0])) / (1000 * 60 * 60 * 24));
    let flag = '';
    if (last[2] === 'ISSUE') flag = 'FLAGGED';
    else if (last[2] === 'OUT' && daysSince >= 1) flag = 'STUCK IN TRANSIT';
    else if (last[2] === 'DELIVERED' && daysSince >= 3) flag = 'OVERDUE RETURN';

    return [crateId, status, last[2], last[3], last[4], last[0], daysSince, tripCount, flag];
  });

  dashboard.clear();
  dashboard.appendRow(['Crate ID', 'Status', 'Last Event', 'Route', 'Customer', 'Last Event Time', 'Days Since', 'Trip Count', 'Flag']);
  if (outRows.length) {
    dashboard.getRange(2, 1, outRows.length, 9).setValues(outRows);
  }
  dashboard.setFrozenRows(1);
}
