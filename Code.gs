// ============================================================
//  MIT ACSC – IT Portal  |  Google Apps Script Backend
//  File: Code.gs  (Fixed – all TypeError issues resolved)
//
//  ROOT CAUSE ANALYSIS & FIXES:
//
//  ERROR 1  addToSheet @ line 147
//    TypeError: Cannot convert undefined or null to object
//    CAUSE:  Object.keys(rowData) was called when rowData was
//            undefined. This happens when a row-builder like
//            staffRow(p) is triggered manually from the GAS
//            editor Run button — the editor passes no argument,
//            so p is undefined, and the row-builder returns
//            nothing useful (or crashes before returning).
//    FIX:    (a) All row-builders now start with p = p || {}
//            (b) addToSheet now validates rowData before use.
//
//  ERRORS 2-13  ticketRow/woRow/staffRow/vendorRegRow/assetRow/
//    budgetRow/challanRow/scrapRow/handoverRow/labRow/
//    equipmentRow/docRow/addEquipment/addWO/addTicket
//    TypeError: Cannot read properties of undefined (reading 'xxx')
//    CAUSE:  All these functions accept a parameter `p` that
//            comes from the parsed POST/GET body. When run
//            directly from the GAS editor (▶ Run button) no
//            event object exists, so e is undefined → p is
//            undefined → p.date / p.empId etc. all crash.
//    FIX:    Every row-builder and write handler starts with:
//              p = p || {};
//            This ensures p is always at least an empty object
//            and every `p.field || ''` fallback can operate.
//
//  ERROR 14  handleSendOTP @ line 492
//    TypeError: Cannot read properties of undefined (reading 'username')
//    CAUSE:  Same as above — p was undefined when run manually.
//    FIX:    p = p || {} at function start, plus added check
//            for empty username before proceeding.
//
//  ERROR 15  setupAllSheets @ line 654
//    Exception: Cannot call SpreadsheetApp.getUi() from this context.
//    CAUSE:  getUi() only works when called from inside the
//            Spreadsheet UI (a menu click or button). It cannot
//            be called from:
//            - The GAS editor Run button
//            - A Web App doGet/doPost execution
//            - A time-driven trigger
//    FIX:    Replaced getUi().alert() with Logger.log().
//            Logger.log() works in ALL execution contexts.
//            View results: GAS editor → View → Execution log.
//
//  RECOMMENDATION:
//    Never test doGet/doPost by pressing ▶ Run in the editor.
//    Those functions require a real event object `e`.
//    Use the testXxx() functions at the bottom of this file
//    instead — they pass a properly structured fake event.
//    Example: select testAddTicket → ▶ Run → Execution log
// ============================================================


// ── SHEET NAMES ── must match tab names in your Spreadsheet ──
var SHEETS = {
  tickets:   'Tickets',
  wos:       'VendorWOs',
  staff:     'Staff',
  vendorReg: 'VendorReg',
  assets:    'Assets',
  budget:    'Budget',
  challan:   'Challan',
  scrap:     'Scrap',
  handover:  'Handover',
  labs:      'Labs',
  equipment: 'Equipment',
  docs:      'Docs'
};

// ── Email map: username → staff email (update with real addresses) ──
var USER_EMAILS = {
  'admin':      'it-admin@mitacsc.ac.in',
  'rutuj':      'rutuj.deshmukh@mitacsc.ac.in',
  'sandeep':    'sandeep.muley@mitacsc.ac.in',
  'mangesh':    'mangesh.sonawane@mitacsc.ac.in',
  'pankaj':     'pankaj.more@mitacsc.ac.in',
  'ziyaafshan': 'ziyaafshan.pathan@mitacsc.ac.in',
  'ashwni':     'ashwni.kadam@mitacsc.ac.in',
  'bhavik':     'bhavik.shah@mitacsc.ac.in',
  'director':   'director@mitacsc.ac.in',
  'registrar':  'registrar@mitacsc.ac.in'
};


// ════════════════════════════════════════════════════════════
//  RESPONSE HELPERS
// ════════════════════════════════════════════════════════════

function makeResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok(data) { return makeResponse(data); }
function err(msg) { return makeResponse({ success: false, error: String(msg) }); }


// ════════════════════════════════════════════════════════════
//  doGet – handles ALL requests: reads AND writes
//
//  WHY GET FOR EVERYTHING (CORS fix justification):
//  Browsers send an OPTIONS preflight before any cross-origin
//  POST with Content-Type:application/json. GAS ignores OPTIONS
//  and redirects → browser blocks it → "Failed to fetch".
//  GET requests with URL params need no preflight → no CORS issue.
//  This is the standard, officially recommended GAS pattern.
// ════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    var p      = (e && e.parameter) ? e.parameter : {};
    var action = p.action || '';

    // ── PING ──────────────────────────────────────────────
    if (action === 'ping') return ok({ status: 'ok', success: true, method: 'GET', app: 'MIT-IT-Portal' });

    // ── READ actions ──────────────────────────────────────
    if (action === 'getTickets')   return ok({ rows: readSheet(SHEETS.tickets).map(normalizeTicketRow) });
    if (action === 'getWOs')       return ok({ rows: readSheet(SHEETS.wos) });
    if (action === 'getStaff')     return ok({ rows: readSheet(SHEETS.staff) });
    if (action === 'getVendorReg') return ok({ rows: readSheet(SHEETS.vendorReg) });
    if (action === 'getAssets')    return ok({ rows: readSheet(SHEETS.assets) });
    if (action === 'getBudget')    return ok({ rows: readSheet(SHEETS.budget) });
    if (action === 'getChallan')   return ok({ rows: readSheet(SHEETS.challan) });
    if (action === 'getScrap')     return ok({ rows: readSheet(SHEETS.scrap) });
    if (action === 'getHandover')  return ok({ rows: readSheet(SHEETS.handover) });
    if (action === 'getLabs')      return ok({ rows: readSheet(SHEETS.labs) });
    if (action === 'getEquipment') return ok({ rows: readSheet(SHEETS.equipment) });
    if (action === 'getDocs')      return ok({ rows: readSheet(SHEETS.docs) });
    if (action === 'getDashboard') return ok(getDashboard());
    if (action === 'sendOTP')      return handleSendOTP(p);

    // ── USER MANAGEMENT ───────────────────────────────────
    if (action === 'getUsers')    return getUsersHandler();
    if (action === 'addUser')     return addUserHandler(p);
    if (action === 'updateUser')  return updateUserHandler(p);
    if (action === 'deleteUser')  return deleteUserHandler(p);

    // ── EDIT / DELETE rows ────────────────────────────────
    if (action === 'deleteRow')   return deleteRowHandler(p);
    if (action === 'updateRow')   return updateRowHandler(p);

    // ── WRITE actions (all arrive via GET URL params) ─────
    if (action === 'updateStatus') return updateTicketStatus(p.ticketId, p.status);
    if (action === 'addTicket')    return addTicket(p);
    if (action === 'addWO')        return addWO(p);
    if (action === 'addStaff')     return addToSheet(SHEETS.staff,     staffRow(p));
    if (action === 'addVendorReg') return addToSheet(SHEETS.vendorReg, vendorRegRow(p));
    if (action === 'addAsset')     return addToSheet(SHEETS.assets,    assetRow(p));
    if (action === 'addBudget')    return addToSheet(SHEETS.budget,    budgetRow(p));
    if (action === 'addChallan')   return addToSheet(SHEETS.challan,   challanRow(p));
    if (action === 'addScrap')     return addToSheet(SHEETS.scrap,     scrapRow(p));
    if (action === 'addHandover')  return addToSheet(SHEETS.handover,  handoverRow(p));
    if (action === 'addLab')       return addToSheet(SHEETS.labs,      labRow(p));
    if (action === 'addEquipment') return addEquipment(p);
    if (action === 'addDoc')       return addToSheet(SHEETS.docs,      docRow(p));

    return ok({ status: 'ok', action: action || 'none' });
  } catch(ex) {
    return err('doGet error: ' + ex.message);
  }
}


// ════════════════════════════════════════════════════════════
//  doPost – safety fallback only
//  The portal no longer calls POST. All operations use GET.
//  doPost is kept here in case it is ever called directly.
// ════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    var p = {};
    if (e && e.postData && e.postData.contents) {
      try   { p = JSON.parse(e.postData.contents); }
      catch (_) { p = (e.parameter) ? e.parameter : {}; }
    } else if (e && e.parameter) {
      p = e.parameter;
    }
    // Delegate to doGet so POST and GET behave identically
    return doGet({ parameter: p });
  } catch(ex) {
    return err('doPost error: ' + ex.message);
  }
}


// ════════════════════════════════════════════════════════════
//  SHEET HELPERS
// ════════════════════════════════════════════════════════════

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function readSheet(name) {
  var sheet = getOrCreateSheet(name);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  return data.slice(1)
    .filter(function(row) {
      return row.some(function(cell) { return cell !== '' && cell !== null && cell !== undefined; });
    })
    .map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) {
        var v = row[i];
        obj[h] = (v instanceof Date)
          ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : (v === null || v === undefined ? '' : String(v));
      });
      return obj;
    });
}

// FIX ERROR 1: guard rowData — if undefined/null, return error immediately
function addToSheet(sheetName, rowData) {
  if (!rowData || typeof rowData !== 'object') {
    return err('addToSheet: invalid rowData for sheet "' + sheetName + '"');
  }
  var sheet   = getOrCreateSheet(sheetName);
  var headers = Object.keys(rowData);
  var values  = Object.values(rowData);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#8B1840')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setFontSize(10);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow(values);
  SpreadsheetApp.flush();
  return ok({ success: true });
}


// ════════════════════════════════════════════════════════════
//  ROW BUILDERS
//  FIX ERRORS 2-13: p = p || {} at start of EVERY function
//  prevents "Cannot read properties of undefined (reading 'x')"
//  when the function is called without an argument from the
//  GAS editor Run button or from a test with no event object.
// ════════════════════════════════════════════════════════════

function ticketRow(p, ticketId) {
  p = p || {};
  return {
    'Ticket ID':   ticketId      || '',
    'Date':        p.date        || today(),
    'Time':        p.time        || '',
    'Assigned To': p.assignedTo  || '',
    'Assigned By': p.assignedBy  || '',
    'Department':  p.dept        || '',
    'Location':    p.location    || '',
    'Category':    p.category    || '',
    'Description': p.description || '',
    'Priority':    p.priority    || 'Medium',
    'Status':      p.status      || 'Open',
    'Vendor':      p.vendor      || '',
    'Remarks':     p.remarks     || '',
    'Logged At':   new Date().toISOString()
  };
}

function woRow(p, woId) {
  p = p || {};
  return {
    'WO ID':         woId           || '',
    'Date':          p.date         || today(),
    'Vendor Name':   p.vendorName   || '',
    'Contact':       p.contact      || '',
    'Contract #':    p.contract     || '',
    'Description':   p.description  || '',
    'Location/Dept': p.location     || '',
    'Coordinator':   p.coordinator  || '',
    'Status':        p.status       || 'Pending Approval',
    'Invoice Amt':   p.invoiceAmt   || '',
    'Remarks':       p.remarks      || '',
    'Logged At':     new Date().toISOString()
  };
}

function staffRow(p) {
  p = p || {};
  return {
    'Emp ID':         p.empId   || '',
    'Name':           p.name    || '',
    'Designation':    p.desig   || '',
    'Qualification':  p.qual    || '',
    'Mobile':         p.mobile  || '',
    'Email':          p.email   || '',
    'Specialisation': p.spec    || '',
    'DOJ':            p.doj     || '',
    'Remarks':        p.remarks || '',
    'Added At':       new Date().toISOString()
  };
}

function vendorRegRow(p) {
  p = p || {};
  return {
    'Vendor ID':  p.vendorId || '',
    'Company':    p.name     || '',
    'Contact':    p.contact  || '',
    'Mobile':     p.mobile   || '',
    'Email':      p.email    || '',
    'GSTIN':      p.gstin    || '',
    'Services':   p.services || '',
    'AMC #':      p.amc      || '',
    'Start':      p.start    || '',
    'End':        p.end      || '',
    'Value (₹)':  p.value    || '',
    'Status':     p.status   || '',
    'Remarks':    p.remarks  || '',
    'Added At':   new Date().toISOString()
  };
}

function assetRow(p) {
  p = p || {};
  return {
    'Asset ID':    p.assetId    || '',
    'Type':        p.type       || '',
    'Brand/Model': p.model      || '',
    'Serial No':   p.serial     || '',
    'Location':    p.location   || '',
    'Dept':        p.dept       || '',
    'Purchase':    p.purchase   || '',
    'Warranty':    p.warranty   || '',
    'Condition':   p.condition  || '',
    'Last Svc':    '',
    'Next Svc':    '',
    'AMC':         p.amc        || '',
    'Remarks':     p.remarks    || '',
    'uploadedBy':  p.uploadedBy || '',
    'Added At':    new Date().toISOString()
  };
}

function budgetRow(p) {
  p = p || {};
  return {
    'FY':            p.year      || '',
    'Budget Head':   p.head      || '',
    'Allocated (₹)': p.allocated || 0,
    'Spent (₹)':     p.spent     || 0,
    'Balance (₹)':   p.balance   || 0,
    'Quarter':       p.quarter   || '',
    'Approved By':   p.approved  || '',
    'Description':   p.desc      || '',
    'Added At':      new Date().toISOString()
  };
}

function challanRow(p) {
  p = p || {};
  return {
    'Challan No':    p.challanNo || '',
    'Date':          p.date      || today(),
    'Type':          p.type      || '',
    'Vendor/Source': p.vendor    || '',
    'Items':         p.items     || '',
    'By':            p.by        || '',
    'WO Ref':        p.ref       || '',
    'Remarks':       p.remarks   || '',
    'Added At':      new Date().toISOString()
  };
}

function scrapRow(p) {
  p = p || {};
  return {
    'Scrap ID':    p.scrapId   || '',
    'Asset Ref':   p.assetRef  || '',
    'Description': p.desc      || '',
    'Qty':         p.qty       || '',
    'Date':        p.date      || today(),
    'Reason':      p.reason    || '',
    'Value (₹)':   p.value     || '',
    'Approved By': p.approved  || '',
    'Disposal':    p.disposal  || '',
    'Remarks':     p.remarks   || '',
    'Added At':    new Date().toISOString()
  };
}

function handoverRow(p) {
  p = p || {};
  return {
    'HO ID':       p.hoId         || '',
    'Date':        p.date         || today(),
    'Asset ID':    p.assetId      || '',
    'Description': p.desc         || '',
    'From':        p.from         || '',
    'To':          p.to           || '',
    'From Dept':   p.fromDept     || '',
    'To Dept':     p.toDept       || '',
    'Condition':   p.condition    || '',
    'Witness':     p.witness      || '',
    'Accessories': p.accessories  || '',
    'Remarks':     p.remarks      || '',
    'IT Tech':     p.tech         || '',
    'Added At':    new Date().toISOString()
  };
}

function labRow(p) {
  p = p || {};
  return {
    'Lab No.':           p.labNo    || '',
    'Lab Name':          p.name     || '',
    'School/Department': p.dept     || '',
    'Building/Block':    p.block    || '',
    'Assigned IT Tech':  p.tech     || '',
    'Capacity':          p.capacity || '',
    'Workstations':      p.ws       || '',
    'Network Switch':    p.switch_  || '',
    'Rack Location':     p.rack     || '',
    'Wi-Fi AP':          p.wifi     || '',
    'Remarks':           p.remarks  || '',
    'Added At':          new Date().toISOString()
  };
}

function equipmentRow(p, equipId) {
  p = p || {};
  return {
    'Equip ID':      equipId       || '',
    'Type':          p.type        || '',
    'Brand/Model':   p.model       || '',
    'Serial No.':    p.serial      || '',
    'Host Name':     p.hostname    || '',
    'User/Assigned': p.user_       || '',
    'OS':            p.os          || '',
    'Boot Type':     p.boot        || '',
    'Processor':     p.processor   || '',
    'RAM':           p.ram         || '',
    'Storage':       p.storage     || '',
    'Monitor':       p.monitor     || '',
    'IP Address':    p.ip          || '',
    'MAC Address':   p.mac         || '',
    'Rack':          p.rack        || '',
    'Switch No.':    p.switch_     || '',
    'Port No.':      p.port        || '',
    'I/O Ports':     p.io          || '',
    'SSID':          p.ssid        || '',
    'VLAN':          p.vlan        || '',
    'Lab/Location':  p.location    || '',
    'Department':    p.dept        || '',
    'IT Tech':       p.tech        || '',
    'Used By':       p.usedby      || '',
    'Purchase':      p.purchase    || '',
    'Warranty':      p.warranty    || '',
    'Condition':     p.condition   || '',
    'Status':        p.status      || 'Active',
    'Remarks':       p.remarks     || '',
    'Added At':      new Date().toISOString()
  };
}

function docRow(p) {
  p = p || {};
  return {
    'Doc ID':      'DOC' + String(Date.now()).slice(-6),
    'Type':        p.type        || '',
    'Title':       p.title       || '',
    'Issue Date':  p.issueDate   || '',
    'Expiry Date': p.expiryDate  || '',
    'Size':        p.size        || '',
    'Extension':   p.ext         || '',
    'Visible To':  p.visibleTo   || 'All Users',
    'Uploaded By': p.uploadedBy  || '',
    'File Name':   p.fileName    || '',
    'Added At':    new Date().toISOString()
  };
}


// ════════════════════════════════════════════════════════════
//  SPECIFIC WRITE HANDLERS
// ════════════════════════════════════════════════════════════

function addTicket(p) {
  p = p || {};
  var sheet     = getOrCreateSheet(SHEETS.tickets);
  var lastRow   = sheet.getLastRow();
  var num       = (lastRow > 0) ? lastRow : 1;
  var ticketId  = 'MIT-IT-' + String(num).padStart(3, '0');
  var result    = addToSheet(SHEETS.tickets, ticketRow(p, ticketId));
  // Auto-send critical alert email
  if ((p.priority || '').toLowerCase() === 'critical') {
    try {
      sendCriticalAlert(ticketId, p.description || '', p.dept || '', p.assignedTo || '');
    } catch(ex) {
      Logger.log('Critical alert error: ' + ex.message);
    }
  }
  return result;
}

function addWO(p) {
  p = p || {};
  var sheet   = getOrCreateSheet(SHEETS.wos);
  var lastRow = sheet.getLastRow();
  var num     = (lastRow > 0) ? lastRow : 1;
  var woId    = 'WO-' + String(num).padStart(3, '0');
  return addToSheet(SHEETS.wos, woRow(p, woId));
}

function addEquipment(p) {
  p = p || {};   // FIX: was crashing on p.equipId when p was undefined
  var sheet   = getOrCreateSheet(SHEETS.equipment);
  var lastRow = sheet.getLastRow();
  var equipId = p.equipId || ('EQ-' + String(lastRow).padStart(3, '0'));
  return addToSheet(SHEETS.equipment, equipmentRow(p, equipId));
}

function updateTicketStatus(ticketId, newStatus) {
  if (!ticketId || !newStatus) return err('ticketId and status are both required');
  var sheet = getOrCreateSheet(SHEETS.tickets);
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return err('No tickets found in sheet');

  var headers   = data[0].map(function(h) { return String(h).trim(); });
  var idCol     = headers.indexOf('Ticket ID');
  var statusCol = headers.indexOf('Status');

  if (idCol < 0)     return err('Column "Ticket ID" not found');
  if (statusCol < 0) return err('Column "Status" not found');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][idCol]).trim() === String(ticketId).trim()) {
      sheet.getRange(i + 1, statusCol + 1).setValue(newStatus);
      SpreadsheetApp.flush();
      return ok({ success: true, ticketId: ticketId, newStatus: newStatus });
    }
  }
  return err('Ticket not found: ' + ticketId);
}


// ════════════════════════════════════════════════════════════
//  DASHBOARD SUMMARY
// ════════════════════════════════════════════════════════════

function getDashboard() {
  var tickets = readSheet(SHEETS.tickets);
  var wos     = readSheet(SHEETS.wos);

  var total      = tickets.length;
  var resolved   = tickets.filter(function(t) { return ['Resolved','Closed'].indexOf(t['Status']) > -1; }).length;
  var inProgress = tickets.filter(function(t) { return t['Status'] === 'In Progress'; }).length;
  var critical   = tickets.filter(function(t) { return t['Priority'] === 'Critical'; }).length;

  var cats = {};
  tickets.forEach(function(t) { var c = t['Category'] || 'Other'; cats[c] = (cats[c] || 0) + 1; });

  var pris = {};
  tickets.forEach(function(t) { var pr = t['Priority'] || 'Medium'; pris[pr] = (pris[pr] || 0) + 1; });

  return {
    total: total,
    resolved: resolved,
    inProgress: inProgress,
    critical: critical,
    vendorWOs: wos.length,
    categories: cats,
    priorities: pris
  };
}


// ════════════════════════════════════════════════════════════
//  OTP EMAIL
//  FIX ERROR 14: p = p || {} prevents crash on undefined
// ════════════════════════════════════════════════════════════

function handleSendOTP(p) {
  p = p || {};   // FIX
  var username = (p.username || '').toLowerCase().trim();
  if (!username) return err('username parameter is required');

  var email = USER_EMAILS[username];
  if (!email) return err('No email registered for user: ' + username);

  var otp = Math.floor(100000 + Math.random() * 900000).toString();
  var exp = Date.now() + 10 * 60 * 1000;

  PropertiesService.getScriptProperties().setProperty(
    'otp_' + username,
    JSON.stringify({ otp: otp, exp: exp })
  );

  try {
    MailApp.sendEmail({
      to:      email,
      subject: 'MIT ACSC IT Portal – Password Reset OTP',
      body: [
        'Dear ' + (p.name || username) + ',',
        '',
        'Your OTP for password reset is:',
        '',
        '    ' + otp,
        '',
        'This OTP is valid for 10 minutes.',
        'If you did not request this, please ignore this email.',
        '',
        '– MIT ACSC IT Section',
        'Alandi, Pune – 412105'
      ].join('\n')
    });
    return ok({ success: true, message: 'OTP sent to registered email' });
  } catch(ex) {
    return err('Email send failed: ' + ex.message);
  }
}

function verifyOTPServer(username, enteredOTP) {
  var stored = PropertiesService.getScriptProperties().getProperty('otp_' + username);
  if (!stored) return { valid: false, reason: 'No OTP found' };
  var data = JSON.parse(stored);
  if (Date.now() > data.exp)    return { valid: false, reason: 'OTP expired' };
  if (data.otp !== enteredOTP)  return { valid: false, reason: 'Wrong OTP' };
  PropertiesService.getScriptProperties().deleteProperty('otp_' + username);
  return { valid: true };
}


// ════════════════════════════════════════════════════════════
//  UTILITY
// ════════════════════════════════════════════════════════════

function today() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}


// ════════════════════════════════════════════════════════════
//  ONE-TIME SHEET SETUP
//  Select setupAllSheets in the editor → ▶ Run
//
//  FIX ERROR 15: Replaced SpreadsheetApp.getUi().alert() with
//  Logger.log() because getUi() throws:
//    "Cannot call SpreadsheetApp.getUi() from this context"
//  when called from the Run button (not a UI menu/button).
//  Logger.log works in ALL contexts — view in Execution log.
// ════════════════════════════════════════════════════════════

function setupAllSheets() {
  var configs = [
    { name: SHEETS.tickets,
      headers: ['Ticket ID','Date','Time','Assigned To','Assigned By','Department',
                'Location','Category','Description','Priority','Status',
                'Vendor','Remarks','Logged At'] },
    { name: SHEETS.wos,
      headers: ['WO ID','Date','Vendor Name','Contact','Contract #','Description',
                'Location/Dept','Coordinator','Status','Invoice Amt','Remarks','Logged At'] },
    { name: SHEETS.staff,
      headers: ['Emp ID','Name','Designation','Qualification','Mobile','Email',
                'Specialisation','DOJ','Remarks','Added At'] },
    { name: SHEETS.vendorReg,
      headers: ['Vendor ID','Company','Contact','Mobile','Email','GSTIN',
                'Services','AMC #','Start','End','Value (₹)','Status','Remarks','Added At'] },
    { name: SHEETS.assets,
      headers: ['Asset ID','Type','Brand/Model','Serial No','Location','Dept',
                'Purchase','Warranty','Condition','Last Svc','Next Svc','AMC',
                'Remarks','uploadedBy','Added At'] },
    { name: SHEETS.budget,
      headers: ['FY','Budget Head','Allocated (₹)','Spent (₹)','Balance (₹)',
                'Quarter','Approved By','Description','Added At'] },
    { name: SHEETS.challan,
      headers: ['Challan No','Date','Type','Vendor/Source','Items','By',
                'WO Ref','Remarks','Added At'] },
    { name: SHEETS.scrap,
      headers: ['Scrap ID','Asset Ref','Description','Qty','Date','Reason',
                'Value (₹)','Approved By','Disposal','Remarks','Added At'] },
    { name: SHEETS.handover,
      headers: ['HO ID','Date','Asset ID','Description','From','To',
                'From Dept','To Dept','Condition','Witness','Accessories',
                'Remarks','IT Tech','Added At'] },
    { name: SHEETS.labs,
      headers: ['Lab No.','Lab Name','School/Department','Building/Block',
                'Assigned IT Tech','Capacity','Workstations','Network Switch',
                'Rack Location','Wi-Fi AP','Remarks','Added At'] },
    { name: SHEETS.equipment,
      headers: ['Equip ID','Type','Brand/Model','Serial No.','Host Name',
                'User/Assigned','OS','Boot Type','Processor','RAM','Storage',
                'Monitor','IP Address','MAC Address','Rack','Switch No.',
                'Port No.','I/O Ports','SSID','VLAN','Lab/Location',
                'Department','IT Tech','Used By','Purchase','Warranty',
                'Condition','Status','Remarks','Added At'] },
    { name: SHEETS.docs,
      headers: ['Doc ID','Type','Title','Issue Date','Expiry Date','Size',
                'Extension','Visible To','Uploaded By','File Name','Added At'] },
    { name: 'Users',
      headers: ['Username','Full Name','Role','Email','Deny','Status',
                'Temp Password','MustChange','Added At','Added By'] }
  ];

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var created = 0;
  var skipped = 0;

  configs.forEach(function(cfg) {
    var sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
      created++;
    }
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(cfg.headers);
      var hRange = sheet.getRange(1, 1, 1, cfg.headers.length);
      hRange.setBackground('#8B1840')
            .setFontColor('#ffffff')
            .setFontWeight('bold')
            .setFontSize(10);
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 110);
      Logger.log('✅ Set up: ' + cfg.name + ' (' + cfg.headers.length + ' columns)');
    } else {
      Logger.log('⏭  Skipped (has data): ' + cfg.name);
      skipped++;
    }
  });

  // FIX: Logger.log instead of getUi().alert()
  Logger.log('══════════════════════════════');
  Logger.log('setupAllSheets COMPLETE');
  Logger.log('Created : ' + created + ' sheet(s)');
  Logger.log('Skipped : ' + skipped + ' sheet(s)');
  Logger.log('Total   : ' + configs.length);
  Logger.log('NEXT: Deploy → New deployment → Web app');
  Logger.log('  Execute as  : Me');
  Logger.log('  Who can access: Anyone');
  Logger.log('══════════════════════════════');
}


// ════════════════════════════════════════════════════════════
//  LOCAL TEST FUNCTIONS
//  Use these to test from the GAS editor without a browser.
//  Select the function → ▶ Run → View → Execution log
//
//  WHY: doGet/doPost require a real event object `e`.
//  Pressing ▶ Run on doGet or doPost directly passes
//  undefined as `e`, which causes all the TypeErrors above.
//  These test functions pass a properly structured fake event.
// ════════════════════════════════════════════════════════════

function testPing() {
  var r = doGet({ parameter: { action: 'ping' } });
  Logger.log('GET ping: ' + r.getContent());
  var r2 = doPost({ postData: { contents: JSON.stringify({ action: 'ping' }) } });
  Logger.log('POST ping: ' + r2.getContent());
}

function testAddTicket() {
  var r = doPost({
    postData: {
      contents: JSON.stringify({
        action:      'addTicket',
        date:        today(),
        time:        '10:30',
        assignedTo:  'Rutuj Deshmukh (IT Tech)',
        assignedBy:  'System Administrator',
        dept:        'HOD Computer Application',
        location:    'CS Lab 101',
        category:    'Hardware Repair',
        description: 'Test ticket – PC not booting in Lab 101',
        priority:    'High',
        status:      'Open',
        vendor:      '',
        remarks:     'Test entry from GAS editor'
      })
    }
  });
  Logger.log('addTicket: ' + r.getContent());
}

function testAddEquipment() {
  var r = doPost({
    postData: {
      contents: JSON.stringify({
        action:    'addEquipment',
        equipId:   'EQ-TEST-001',
        type:      'Desktop',
        model:     'Dell OptiPlex 7010',
        serial:    'SN-TEST-001',
        hostname:  'MIT-TEST-PC001',
        os:        'Windows 11 Pro',
        boot:      'Single Boot',
        ip:        '192.168.1.101',
        mac:       '00:1A:2B:3C:4D:5E',
        location:  'CS Lab 101',
        dept:      'HOD Computer Application',
        tech:      'Rutuj Deshmukh (IT Tech)',
        usedby:    'Students',
        condition: 'Good',
        status:    'Active',
        remarks:   'Test equipment entry'
      })
    }
  });
  Logger.log('addEquipment: ' + r.getContent());
}

function testAddStaff() {
  var r = doPost({
    postData: {
      contents: JSON.stringify({
        action:  'addStaff',
        empId:   'EMP-001',
        name:    'Test Staff Member',
        desig:   'IT Technician',
        qual:    'B.E. Computer',
        mobile:  '9876543210',
        email:   'test@mitacsc.ac.in',
        spec:    'Networking',
        doj:     today(),
        remarks: 'Test entry'
      })
    }
  });
  Logger.log('addStaff: ' + r.getContent());
}

function testGetTickets() {
  var r    = doGet({ parameter: { action: 'getTickets' } });
  var data = JSON.parse(r.getContent());
  Logger.log('getTickets – rows: ' + (data.rows ? data.rows.length : 0));
  if (data.rows && data.rows.length > 0) Logger.log('First: ' + JSON.stringify(data.rows[0]));
}

function testGetEquipment() {
  var r    = doGet({ parameter: { action: 'getEquipment' } });
  var data = JSON.parse(r.getContent());
  Logger.log('getEquipment – rows: ' + (data.rows ? data.rows.length : 0));
}

function testDashboard() {
  var r = doGet({ parameter: { action: 'getDashboard' } });
  Logger.log('Dashboard: ' + r.getContent());
}

function testUpdateStatus() {
  // First add a ticket, then update it
  testAddTicket();
  var r    = doGet({ parameter: { action: 'getTickets' } });
  var data = JSON.parse(r.getContent());
  if (data.rows && data.rows.length > 0) {
    var id = data.rows[0]['Ticket ID'];
    var r2 = doPost({
      postData: {
        contents: JSON.stringify({ action: 'updateStatus', ticketId: id, status: 'Resolved' })
      }
    });
    Logger.log('updateStatus (' + id + '): ' + r2.getContent());
  }
}


// ════════════════════════════════════════════════════════════
//  DELETE ROW  –  called by portal Edit/Delete feature
//  action=deleteRow&sheet=Tickets&rowIndex=3
//  rowIndex is 1-based (row 1 = header, row 2 = first data row)
// ════════════════════════════════════════════════════════════

function deleteRowHandler(p) {
  p = p || {};
  var sheetName = p.sheet || '';
  var rowIndex  = parseInt(p.rowIndex || '0');
  if (!sheetName)   return err('sheet parameter required');
  if (!rowIndex || rowIndex < 2) return err('rowIndex must be >= 2 (row 1 is the header)');

  var sheet = getOrCreateSheet(sheetName);
  var lastRow = sheet.getLastRow();
  if (rowIndex > lastRow) return err('rowIndex ' + rowIndex + ' exceeds sheet rows (' + lastRow + ')');

  sheet.deleteRow(rowIndex);
  SpreadsheetApp.flush();
  return ok({ success: true, deleted: rowIndex, sheet: sheetName });
}


// ════════════════════════════════════════════════════════════
//  UPDATE ROW  –  called by portal Edit modal Save button
//  action=updateRow&sheet=Tickets&rowIndex=3&field1=val...
//  All fields except action/sheet/rowIndex are written to
//  their matching columns (matched by header name).
// ════════════════════════════════════════════════════════════

function updateRowHandler(p) {
  p = p || {};
  var sheetName = p.sheet || '';
  var rowIndex  = parseInt(p.rowIndex || '0');
  if (!sheetName)   return err('sheet parameter required');
  if (!rowIndex || rowIndex < 2) return err('rowIndex must be >= 2');

  var sheet   = getOrCreateSheet(sheetName);
  var data    = sheet.getDataRange().getValues();
  if (data.length < 1) return err('Sheet is empty');

  var headers = data[0].map(function(h) { return String(h).trim(); });

  // Build update map — exclude meta params
  var skip = { action:1, sheet:1, rowIndex:1 };
  var updates = {};
  Object.keys(p).forEach(function(k) {
    if (!skip[k]) updates[k] = p[k];
  });

  // Write each matching column
  var written = 0;
  headers.forEach(function(h, colIdx) {
    if (updates.hasOwnProperty(h)) {
      sheet.getRange(rowIndex, colIdx + 1).setValue(updates[h]);
      written++;
    }
  });

  SpreadsheetApp.flush();
  return ok({ success: true, updated: written, rowIndex: rowIndex, sheet: sheetName });
}

// Register deleteRow and updateRow in doGet
// (Added as handlers inside the existing doGet action chain)
// Note: these are called from the portal via gsGet({action:'deleteRow',...})
// They are wired into doGet at the bottom of the action chain below.


// ════════════════════════════════════════════════════════════
//  DAILY EMAIL REPORT
//
//  Sends a rich HTML summary email every morning at 7 AM to:
//    • sknadaf@mitacsc.ac.in  (IT Admin / Principal)
//    • it-admin@mitacsc.ac.in (System Administrator)
//
//  Report includes:
//    • Open / Critical / In-Progress ticket counts
//    • List of all open & critical tickets with details
//    • Vendor WOs pending / in-progress
//    • Equipment count summary
//    • Expiring warranties (next 30 days)
//
//  Setup: Run setupDailyReportTrigger() ONCE from GAS editor
//    to register the daily 7 AM time-based trigger.
// ════════════════════════════════════════════════════════════

// Email recipients — add/remove as needed
var REPORT_RECIPIENTS = [
  'sknadaf@mitacsc.ac.in',
  'it-admin@mitacsc.ac.in'
];

function sendDailyReport() {
  try {
    var tickets   = readSheet(SHEETS.tickets);
    var wos       = readSheet(SHEETS.wos);
    var equipment = readSheet(SHEETS.equipment);
    var todayStr  = today();

    // ── Ticket stats ──
    var open     = tickets.filter(function(t) { return t['Status'] === 'Open'; });
    var critical = tickets.filter(function(t) { return t['Priority'] === 'Critical' && t['Status'] !== 'Resolved' && t['Status'] !== 'Closed'; });
    var inProg   = tickets.filter(function(t) { return t['Status'] === 'In Progress'; });
    var resolved = tickets.filter(function(t) { return (t['Status'] === 'Resolved' || t['Status'] === 'Closed') && (t['Date'] || '').startsWith(todayStr.substring(0,7)); });

    // ── WO stats ──
    var woOpen   = wos.filter(function(w) { return w['Status'] === 'Pending Approval' || w['Status'] === 'In Progress'; });

    // ── Expiring warranties (next 30 days) ──
    var soon = new Date(); soon.setDate(soon.getDate() + 30);
    var soonStr = Utilities.formatDate(soon, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var expiring = equipment.filter(function(e) {
      var w = e['Warranty'] || '';
      return w && w >= todayStr && w <= soonStr;
    });

    // ── Build HTML email ──
    var html = [
      '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto">',
      '<div style="background:#8B1840;padding:20px 24px;border-radius:8px 8px 0 0">',
      '<h2 style="color:#fff;margin:0;font-size:20px">🖥️ MIT ACSC – IT Section</h2>',
      '<p style="color:#f0c4d4;margin:4px 0 0;font-size:13px">Daily IT Report — ' + todayStr + '</p>',
      '</div>',
      '<div style="background:#f9f9f9;padding:20px 24px;border:1px solid #ddd;border-top:none">',

      // KPI row
      '<table style="width:100%;border-collapse:collapse;margin-bottom:20px">',
      '<tr>',
      kpiCell('📋 Open Tickets',    open.length,     '#3b82f6'),
      kpiCell('🔴 Critical',         critical.length, '#ef4444'),
      kpiCell('🔄 In Progress',      inProg.length,   '#8b5cf6'),
      kpiCell('✅ Resolved Today',   resolved.length, '#10b981'),
      kpiCell('🏭 Active WOs',       woOpen.length,   '#f59e0b'),
      '</tr></table>',
    ].join('');

    // ── Open Tickets Table ──
    if (open.length > 0) {
      html += sectionHeader('📋 Open Tickets (' + open.length + ')');
      html += tableStart(['Ticket ID','Date','Assigned To','Dept','Category','Priority','Description']);
      open.forEach(function(t) {
        var pColor = t['Priority']==='Critical'?'#ef4444':t['Priority']==='High'?'#f97316':t['Priority']==='Medium'?'#f59e0b':'#10b981';
        html += '<tr style="border-bottom:1px solid #eee">';
        html += td(t['Ticket ID']||'—','font-weight:700;color:#1F3864') +
                td(t['Date']||'—') +
                td(t['Assigned To']||'—') +
                td((t['Department']||'').substring(0,25)) +
                td(t['Category']||'—') +
                td(t['Priority']||'—','color:'+pColor+';font-weight:700') +
                td((t['Description']||'').substring(0,50));
        html += '</tr>';
      });
      html += '</table></div>';
    }

    // ── Critical Tickets Alert ──
    if (critical.length > 0) {
      html += '<div style="background:#fff5f5;border:1px solid #fca5a5;border-radius:6px;padding:12px 16px;margin:12px 0">';
      html += '<b style="color:#ef4444">⚠️ ' + critical.length + ' CRITICAL ticket(s) require immediate attention:</b><ul style="margin:8px 0 0 16px;color:#7f1d1d">';
      critical.forEach(function(t) {
        html += '<li><b>' + (t['Ticket ID']||'—') + '</b> — ' + (t['Description']||'').substring(0,80) + ' [' + (t['Assigned To']||'').split(' (')[0] + ']</li>';
      });
      html += '</ul></div>';
    }

    // ── Pending Vendor WOs ──
    if (woOpen.length > 0) {
      html += sectionHeader('🏭 Pending Vendor Work Orders (' + woOpen.length + ')');
      html += tableStart(['WO ID','Date','Vendor','Description','Status','Coordinator']);
      woOpen.forEach(function(w) {
        html += '<tr style="border-bottom:1px solid #eee">' +
          td(w['WO ID']||'—','font-weight:700') +
          td(w['Date']||'—') +
          td(w['Vendor Name']||'—') +
          td((w['Description']||'').substring(0,40)) +
          td(w['Status']||'—','color:#f59e0b;font-weight:700') +
          td(w['Coordinator']||'—') +
          '</tr>';
      });
      html += '</table></div>';
    }

    // ── Expiring Warranties ──
    if (expiring.length > 0) {
      html += sectionHeader('⏰ Warranties Expiring in Next 30 Days (' + expiring.length + ')');
      html += tableStart(['Equip ID','Type','Brand/Model','Serial No.','Department','Warranty Date']);
      expiring.forEach(function(e) {
        html += '<tr style="border-bottom:1px solid #eee">' +
          td(e['Equip ID']||'—','font-weight:700') +
          td(e['Type']||'—') +
          td(e['Brand/Model']||'—') +
          td(e['Serial No.']||'—') +
          td(e['Department']||'—') +
          td(e['Warranty']||'—','color:#ef4444;font-weight:700') +
          '</tr>';
      });
      html += '</table></div>';
    }

    // Footer
    html += [
      '<div style="margin-top:20px;padding-top:14px;border-top:1px solid #ddd;',
      'font-size:11px;color:#888;text-align:center">',
      'MIT Arts, Commerce &amp; Science College, Alandi, Pune – 412105<br/>',
      'IT Section | This is an automated daily report generated at 7:00 AM<br/>',
      '<a href="mailto:it-admin@mitacsc.ac.in" style="color:#8B1840">it-admin@mitacsc.ac.in</a>',
      '</div></div></div>'
    ].join('');

    // ── Send email ──
    var subject = '📋 MIT ACSC IT Report – ' + todayStr +
      (critical.length > 0 ? ' 🔴 ' + critical.length + ' CRITICAL' : '') +
      ' | Open: ' + open.length;

    REPORT_RECIPIENTS.forEach(function(recipient) {
      MailApp.sendEmail({
        to:       recipient,
        subject:  subject,
        htmlBody: html,
        body:     buildPlainTextReport(open, critical, inProg, resolved, woOpen, expiring, todayStr)
      });
    });

    Logger.log('✅ Daily report sent to: ' + REPORT_RECIPIENTS.join(', '));
    return ok({ success: true, recipients: REPORT_RECIPIENTS.length, tickets: open.length, critical: critical.length });

  } catch(ex) {
    Logger.log('❌ sendDailyReport error: ' + ex.message);
    return err('sendDailyReport error: ' + ex.message);
  }
}

// ── Plain-text fallback for email clients that block HTML ──
function buildPlainTextReport(open, critical, inProg, resolved, woOpen, expiring, todayStr) {
  var lines = [
    'MIT ACSC – IT Section Daily Report',
    'Date: ' + todayStr,
    '═══════════════════════════════════',
    'SUMMARY',
    '  Open Tickets    : ' + open.length,
    '  Critical        : ' + critical.length,
    '  In Progress     : ' + inProg.length,
    '  Resolved Today  : ' + resolved.length,
    '  Active WOs      : ' + woOpen.length,
    ''
  ];
  if (critical.length > 0) {
    lines.push('⚠️ CRITICAL TICKETS:');
    critical.forEach(function(t) {
      lines.push('  • ' + (t['Ticket ID']||'—') + ' – ' + (t['Description']||'').substring(0,70));
      lines.push('    Assigned: ' + (t['Assigned To']||'—') + ' | Dept: ' + (t['Department']||'—'));
    });
    lines.push('');
  }
  if (open.length > 0) {
    lines.push('OPEN TICKETS:');
    open.forEach(function(t) {
      lines.push('  [' + (t['Priority']||'?') + '] ' + (t['Ticket ID']||'—') + ' – ' + (t['Description']||'').substring(0,60));
    });
    lines.push('');
  }
  if (expiring.length > 0) {
    lines.push('⏰ EXPIRING WARRANTIES (30 days):');
    expiring.forEach(function(e) {
      lines.push('  • ' + (e['Equip ID']||'—') + ' ' + (e['Brand/Model']||'—') + ' – expires ' + (e['Warranty']||'—'));
    });
    lines.push('');
  }
  lines.push('MIT ACSC IT Section | Alandi, Pune 412105');
  lines.push('it-admin@mitacsc.ac.in');
  return lines.join('\n');
}

// ── HTML helper functions ──
function kpiCell(label, val, color) {
  return '<td style="text-align:center;padding:10px;background:#fff;border:1px solid #e5e7eb;border-radius:6px;margin:4px">' +
    '<div style="font-size:22px;font-weight:700;color:' + color + '">' + val + '</div>' +
    '<div style="font-size:11px;color:#6b7280;margin-top:2px">' + label + '</div>' +
    '</td>';
}
function sectionHeader(title) {
  return '<div style="margin:16px 0 8px"><b style="font-size:14px;color:#1F3864">' + title + '</b></div>' +
    '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">' +
    '<thead style="background:#8B1840;color:#fff">';
}
function tableStart(headers) {
  var ths = headers.map(function(h) { return '<th style="padding:7px 8px;text-align:left;font-size:11px">' + h + '</th>'; }).join('');
  return '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px"><thead style="background:#8B1840;color:#fff"><tr>' + ths + '</tr></thead><tbody>';
}
function td(val, style) {
  return '<td style="padding:6px 8px;' + (style||'') + '">' + (val||'—') + '</td>';
}


// ════════════════════════════════════════════════════════════
//  TRIGGER SETUP
//  Run setupDailyReportTrigger() ONCE from GAS editor.
//  This creates a time-based trigger that fires daily at 7 AM.
//  Running it again deletes the old trigger first (no duplicates).
// ════════════════════════════════════════════════════════════

function setupDailyReportTrigger() {
  // Delete any existing daily report triggers to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted existing trigger: ' + trigger.getUniqueId());
    }
  });

  // Create new daily trigger at 7 AM
  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .create();

  Logger.log('✅ Daily report trigger created — fires every day at 7 AM');
  Logger.log('Recipients: ' + REPORT_RECIPIENTS.join(', '));
  Logger.log('To test immediately, run: sendDailyReport()');
}

// ── Test daily report (run from editor) ──
function testSendDailyReport() {
  var result = sendDailyReport();
  Logger.log('testSendDailyReport: ' + JSON.stringify(result));
}


// ════════════════════════════════════════════════════════════
//  USER MANAGEMENT  (Sheet: "Users")
//  Columns: Username | Full Name | Role | Email | Deny |
//           Status | Temp Password | MustChange | Added At | Added By
// ════════════════════════════════════════════════════════════

// Add Users sheet to SHEETS map
SHEETS.users = 'Users';

function getUsersSheetHeaders() {
  return ['Username','Full Name','Role','Email','Deny','Status',
          'Temp Password','MustChange','Added At','Added By'];
}

function ensureUsersSheet() {
  var sheet = getOrCreateSheet(SHEETS.users);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(getUsersSheetHeaders());
    var hRange = sheet.getRange(1, 1, 1, getUsersSheetHeaders().length);
    hRange.setBackground('#8B1840').setFontColor('#ffffff')
          .setFontWeight('bold').setFontSize(10);
    sheet.setFrozenRows(1);
    // Seed with default admin row so sheet is never empty
    sheet.appendRow(['admin','System Administrator','admin','it-admin@mitacsc.ac.in',
                     '','active','','false',today(),'System']);
  }
  return sheet;
}

// Register in doGet action chain (patched at bottom of file)
function getUsersHandler() {
  ensureUsersSheet();
  return ok({ rows: readSheet(SHEETS.users) });
}

function addUserHandler(p) {
  p = p || {};
  ensureUsersSheet();
  var row = {
    'Username':      (p.username     || '').toLowerCase().trim(),
    'Full Name':     p.fullName      || '',
    'Role':          p.role          || 'tech',
    'Email':         p.email         || '',
    'Deny':          p.deny          || '',
    'Status':        p.status        || 'active',
    'Temp Password': p.tempPassword  || '',
    'MustChange':    p.mustChange    || 'true',
    'Added At':      p.addedAt       || today(),
    'Added By':      p.addedBy       || 'admin'
  };
  if (!row['Username']) return err('username is required');
  // Check duplicate
  var existing = readSheet(SHEETS.users);
  var dup = existing.some(function(r) { return (r['Username']||'').toLowerCase() === row['Username']; });
  if (dup) return err('Username "' + row['Username'] + '" already exists');
  return addToSheet(SHEETS.users, row);
}

function updateUserHandler(p) {
  p = p || {};
  var username = (p.username || '').toLowerCase().trim();
  if (!username) return err('username required');

  var sheet = ensureUsersSheet();
  var data  = sheet.getDataRange().getValues();
  if (data.length < 2) return err('Users sheet is empty');

  var headers  = data[0].map(function(h) { return String(h).trim(); });
  var unameCol = headers.indexOf('Username');
  if (unameCol < 0) return err('"Username" column not found');

  // Field → column name mapping
  var fieldMap = {
    fullName:    'Full Name',
    role:        'Role',
    email:       'Email',
    deny:        'Deny',
    status:      'Status',
    tempPassword:'Temp Password',
    mustChange:  'MustChange'
  };

  for (var i = 1; i < data.length; i++) {
    if ((String(data[i][unameCol])||'').toLowerCase().trim() === username) {
      Object.keys(fieldMap).forEach(function(pKey) {
        if (p.hasOwnProperty(pKey)) {
          var colName = fieldMap[pKey];
          var colIdx  = headers.indexOf(colName);
          if (colIdx >= 0) sheet.getRange(i+1, colIdx+1).setValue(p[pKey]);
        }
      });
      SpreadsheetApp.flush();
      return ok({ success: true, username: username });
    }
  }
  return err('User not found: ' + username);
}

function deleteUserHandler(p) {
  p = p || {};
  var username = (p.username || '').toLowerCase().trim();
  if (!username) return err('username required');
  if (username === 'admin') return err('Cannot delete the admin account');

  var sheet = ensureUsersSheet();
  var data  = sheet.getDataRange().getValues();
  var headers  = data[0].map(function(h) { return String(h).trim(); });
  var unameCol = headers.indexOf('Username');
  if (unameCol < 0) return err('"Username" column not found');

  for (var i = 1; i < data.length; i++) {
    if ((String(data[i][unameCol])||'').toLowerCase().trim() === username) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return ok({ success: true, deleted: username });
    }
  }
  return err('User not found: ' + username);
}

// ── Test user management ──
function testUserManagement() {
  Logger.log('--- testUserManagement ---');
  // Add a test user
  var r1 = doGet({ parameter: {
    action:'addUser', username:'testuser99', fullName:'Test User 99',
    role:'tech', email:'test99@mitacsc.ac.in', deny:'invoice,budget',
    status:'active', tempPassword:'test@123', mustChange:'true',
    addedBy:'admin', addedAt:today()
  }});
  Logger.log('addUser: ' + r1.getContent());

  // Get all users
  var r2 = doGet({ parameter: { action:'getUsers' }});
  var data = JSON.parse(r2.getContent());
  Logger.log('getUsers count: ' + (data.rows ? data.rows.length : 0));

  // Update user
  var r3 = doGet({ parameter: { action:'updateUser', username:'testuser99', status:'inactive' }});
  Logger.log('updateUser: ' + r3.getContent());

  // Delete test user
  var r4 = doGet({ parameter: { action:'deleteUser', username:'testuser99' }});
  Logger.log('deleteUser: ' + r4.getContent());
  Logger.log('--- done ---');
}


// ════════════════════════════════════════════════════════════
//  NORMALIZE SHEET ROW → camelCase for portal compatibility
//  The portal uses camelCase (ticketId, assignedTo, etc.)
//  but Sheets stores Title Case ('Ticket ID', 'Assigned To')
//  This function makes both work seamlessly.
// ════════════════════════════════════════════════════════════
function normalizeTicketRow(r) {
  return {
    ticketId:    r['Ticket ID']   || r.ticketId   || '',
    date:        r['Date']        || r.date        || '',
    time:        r['Time']        || r.time        || '',
    assignedTo:  r['Assigned To'] || r.assignedTo  || '',
    assignedBy:  r['Assigned By'] || r.assignedBy  || '',
    dept:        r['Department']  || r.dept        || '',
    location:    r['Location']    || r.location    || '',
    category:    r['Category']    || r.category    || '',
    description: r['Description'] || r.description || '',
    priority:    r['Priority']    || r.priority    || 'Medium',
    status:      r['Status']      || r.status      || 'Open',
    vendor:      r['Vendor']      || r.vendor      || '',
    remarks:     r['Remarks']     || r.remarks     || '',
  };
}

// ════════════════════════════════════════════════════════════
//  CRITICAL ALERT EMAIL
//  Sends an immediate alert when a Critical ticket is logged.
//  Called automatically by addTicket() when priority=Critical.
//  Also callable manually: sendCriticalAlert(ticketId, desc, dept, assignedTo)
// ════════════════════════════════════════════════════════════

function sendCriticalAlert(ticketId, description, dept, assignedTo) {
  try {
    var subject = '🔴 CRITICAL IT Ticket: ' + ticketId + ' | ' + dept;
    var html = [
      '<div style="font-family:Arial,sans-serif;max-width:600px">',
      '<div style="background:#ef4444;padding:16px 20px;border-radius:8px 8px 0 0">',
      '<h2 style="color:#fff;margin:0;font-size:18px">🔴 Critical IT Ticket Logged</h2>',
      '<p style="color:#fee2e2;margin:4px 0 0;font-size:12px">Immediate attention required</p>',
      '</div>',
      '<div style="background:#fff5f5;border:1px solid #fca5a5;border-radius:0 0 8px 8px;padding:20px">',
      '<table style="width:100%;border-collapse:collapse;font-size:13px">',
      '<tr><td style="padding:6px 0;color:#6b7280;width:130px">Ticket ID</td><td style="padding:6px 0;font-weight:700;color:#1F3864">' + ticketId + '</td></tr>',
      '<tr><td style="padding:6px 0;color:#6b7280">Department</td><td style="padding:6px 0">' + dept + '</td></tr>',
      '<tr><td style="padding:6px 0;color:#6b7280">Assigned To</td><td style="padding:6px 0;font-weight:700">' + assignedTo + '</td></tr>',
      '<tr><td style="padding:6px 0;color:#6b7280">Description</td><td style="padding:6px 0">' + description + '</td></tr>',
      '<tr><td style="padding:6px 0;color:#6b7280">Logged At</td><td style="padding:6px 0">' + new Date().toLocaleString('en-IN') + '</td></tr>',
      '</table>',
      '<div style="margin-top:14px;padding:10px;background:#fee2e2;border-radius:6px;font-size:12px;color:#991b1b">',
      '⚠️ This ticket is marked <b>Critical</b>. Please take immediate action.',
      '</div>',
      '</div>',
      '<div style="margin-top:10px;font-size:10px;color:#9ca3af;text-align:center">',
      'MIT ACSC IT Section | Alandi, Pune – 412105',
      '</div></div>'
    ].join('');

    REPORT_RECIPIENTS.forEach(function(recipient) {
      MailApp.sendEmail({ to: recipient, subject: subject, htmlBody: html,
        body: 'CRITICAL Ticket: ' + ticketId + '\nDept: ' + dept + '\nAssigned: ' + assignedTo + '\n' + description });
    });
    Logger.log('✅ Critical alert sent for ' + ticketId);
  } catch(ex) {
    Logger.log('❌ Critical alert email failed: ' + ex.message);
  }
}

// ════════════════════════════════════════════════════════════
//  WEEKLY SUMMARY REPORT  (every Monday 8 AM)
//  Gives a full week-in-review: tickets opened/resolved,
//  avg resolution time, top categories, vendor WO summary.
//  Setup: run setupWeeklyReportTrigger() once from editor.
// ════════════════════════════════════════════════════════════

function sendWeeklyReport() {
  try {
    var tickets = readSheet(SHEETS.tickets).map(normalizeTicketRow);
    var wos     = readSheet(SHEETS.wos);

    // Last 7 days
    var since = new Date();
    since.setDate(since.getDate() - 7);
    var sinceStr = Utilities.formatDate(since, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    var todayStr = today();

    var thisWeek  = tickets.filter(function(t) { return (t.date || '') >= sinceStr; });
    var opened    = thisWeek.length;
    var resolved  = thisWeek.filter(function(t) { return t.status === 'Resolved' || t.status === 'Closed'; }).length;
    var critical  = thisWeek.filter(function(t) { return t.priority === 'Critical'; }).length;
    var stillOpen = tickets.filter(function(t) { return t.status === 'Open' || t.status === 'In Progress'; }).length;

    // Category breakdown this week
    var cats = {};
    thisWeek.forEach(function(t) { var c = t.category||'Other'; cats[c]=(cats[c]||0)+1; });
    var topCats = Object.entries(cats).sort(function(a,b){return b[1]-a[1];}).slice(0,5);

    // Tech workload
    var techLoad = {};
    tickets.filter(function(t){return t.status==='Open'||t.status==='In Progress';})
      .forEach(function(t){ var a=t.assignedTo||'Unassigned'; techLoad[a]=(techLoad[a]||0)+1; });
    var topTech = Object.entries(techLoad).sort(function(a,b){return b[1]-a[1];}).slice(0,5);

    var subject = '📊 MIT ACSC IT Weekly Report | W/E ' + todayStr +
      ' | Opened: ' + opened + ' | Resolved: ' + resolved;

    var html = [
      '<div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto">',
      '<div style="background:#1F3864;padding:20px 24px;border-radius:8px 8px 0 0">',
      '<h2 style="color:#fff;margin:0;font-size:20px">📊 Weekly IT Report</h2>',
      '<p style="color:#93c5fd;margin:4px 0 0;font-size:13px">Week ending ' + todayStr + ' | MIT ACSC IT Section</p>',
      '</div>',
      '<div style="background:#f9f9f9;padding:20px 24px;border:1px solid #ddd;border-top:none">',
      '<table style="width:100%;border-collapse:collapse;margin-bottom:20px"><tr>',
      kpiCell('📋 Opened This Week', opened,   '#3b82f6'),
      kpiCell('✅ Resolved',          resolved, '#10b981'),
      kpiCell('🔴 Critical',          critical, '#ef4444'),
      kpiCell('⏳ Still Open',        stillOpen,'#f59e0b'),
      '</tr></table>',
    ].join('');

    // Top categories table
    if (topCats.length) {
      html += '<div style="margin-bottom:16px"><b style="font-size:14px;color:#1F3864">📁 Top Categories This Week</b></div>';
      html += '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px">';
      html += '<thead style="background:#8B1840;color:#fff"><tr><th style="padding:7px 8px;text-align:left">Category</th><th style="padding:7px 8px;text-align:center">Count</th><th style="padding:7px 8px;text-align:left">Bar</th></tr></thead><tbody>';
      var maxCat = topCats[0][1];
      topCats.forEach(function(c) {
        var pct = Math.round((c[1]/maxCat)*100);
        html += '<tr style="border-bottom:1px solid #eee">' +
          '<td style="padding:6px 8px">' + c[0] + '</td>' +
          '<td style="padding:6px 8px;text-align:center;font-weight:700">' + c[1] + '</td>' +
          '<td style="padding:6px 8px"><div style="background:#e5e7eb;border-radius:3px;height:8px"><div style="background:#3b82f6;width:'+pct+'%;height:8px;border-radius:3px"></div></div></td>' +
          '</tr>';
      });
      html += '</tbody></table>';
    }

    // Tech workload
    if (topTech.length) {
      html += '<div style="margin-bottom:8px"><b style="font-size:14px;color:#1F3864">👤 Current Open Tasks by Tech</b></div>';
      html += '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px">';
      html += '<thead style="background:#8B1840;color:#fff"><tr><th style="padding:7px 8px;text-align:left">IT Tech</th><th style="padding:7px 8px;text-align:center">Open Tasks</th></tr></thead><tbody>';
      topTech.forEach(function(t) {
        html += '<tr style="border-bottom:1px solid #eee"><td style="padding:6px 8px">' + t[0].split('(')[0].trim() + '</td><td style="padding:6px 8px;text-align:center;font-weight:700;color:' + (t[1]>5?'#ef4444':t[1]>3?'#f59e0b':'#10b981') + '">' + t[1] + '</td></tr>';
      });
      html += '</tbody></table>';
    }

    html += '<div style="margin-top:16px;padding-top:14px;border-top:1px solid #ddd;font-size:11px;color:#888;text-align:center">' +
      'MIT ACSC IT Section | Alandi, Pune – 412105 | Automated weekly report every Monday 8 AM' +
      '</div></div></div>';

    REPORT_RECIPIENTS.forEach(function(r) {
      MailApp.sendEmail({ to: r, subject: subject, htmlBody: html,
        body: 'Weekly IT Report\nOpened: ' + opened + ' | Resolved: ' + resolved + ' | Critical: ' + critical + ' | Still Open: ' + stillOpen });
    });

    Logger.log('✅ Weekly report sent');
  } catch(ex) {
    Logger.log('❌ sendWeeklyReport error: ' + ex.message);
  }
}

function setupWeeklyReportTrigger() {
  // Remove existing weekly triggers
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendWeeklyReport') ScriptApp.deleteTrigger(t);
  });
  // Every Monday at 8 AM
  ScriptApp.newTrigger('sendWeeklyReport')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8)
    .create();
  Logger.log('✅ Weekly report trigger set: every Monday 8 AM');
  Logger.log('Recipients: ' + REPORT_RECIPIENTS.join(', '));
}

// ── Convenience: set up ALL triggers at once ──────────────────────────
function setupAllTriggers() {
  setupDailyReportTrigger();
  setupWeeklyReportTrigger();
  setupPendingReminderTrigger();
  Logger.log('✅ All triggers configured:');
  Logger.log('  Daily report      → every day 7 AM');
  Logger.log('  Pending reminders → every day 8 AM (per-user personalised)');
  Logger.log('  Weekly report     → every Monday 8 AM');
  Logger.log('  Critical alert    → auto on ticket submit (Priority=Critical)');
  Logger.log('Recipients: ' + REPORT_RECIPIENTS.join(', '));
}

// ── Test all email functions ──────────────────────────────────────────
function testAllEmails() {
  Logger.log('=== Testing all email functions ===');
  // Test critical alert
  sendCriticalAlert('MIT-IT-TEST', 'Test critical alert', 'HOD Computer Application', 'System Administrator');
  Logger.log('✅ Critical alert sent');
  // Test daily report
  sendDailyReport();
  Logger.log('✅ Daily report sent');
  // Test weekly report
  sendWeeklyReport();
  Logger.log('✅ Weekly report sent');
  Logger.log('=== Check inbox of: ' + REPORT_RECIPIENTS.join(', ') + ' ===');
}


// ════════════════════════════════════════════════════════════
//  DAILY PENDING TASK REMINDER
//
//  Sends a personalised reminder email to EACH non-admin user
//  showing ONLY their own pending/open/in-progress tasks.
//
//  From:    sknadaf@mitacsc.ac.in  (IT Admin / Principal)
//  To:      Each tech/view user's registered email
//  When:    Daily at 8 AM (set via setupPendingReminderTrigger)
//  Content: • Their open & in-progress tickets
//           • Their active vendor WOs (as coordinator)
//           • Overdue tickets (date older than today)
//           • A motivational closing note
//
//  Email sources:
//    1. Users Sheet "Email" column  (added via User Management)
//    2. USER_EMAILS map (fallback for hardcoded users)
//
//  HOW USER MATCHING WORKS:
//    Tickets sheet "Assigned To" column contains the user's
//    Full Name (e.g. "Rutuj Deshmukh (IT Tech)").
//    We match by first name (first word of Full Name) —
//    same logic as the portal's getVisibleRows() filter.
// ════════════════════════════════════════════════════════════

var FROM_EMAIL = 'sknadaf@mitacsc.ac.in';   // Sender — must be a Google account that runs this script

function sendDailyPendingTaskReminder() {
  try {
    var allTickets = readSheet(SHEETS.tickets).map(normalizeTicketRow);
    var allWOs     = readSheet(SHEETS.wos);
    var usersSheet = readSheet(SHEETS.users || 'Users');
    var todayStr   = today();
    var sentCount  = 0;
    var skippedCount = 0;
    var errors     = [];

    // ── Build full user list from Users sheet + USER_EMAILS fallback ──
    // Users sheet row: { Username, Full Name, Role, Email, Status, ... }
    var userList = [];

    // From Users sheet
    usersSheet.forEach(function(row) {
      var uname  = (row['Username'] || '').toLowerCase().trim();
      var role   = (row['Role']     || 'tech').toLowerCase().trim();
      var status = (row['Status']   || 'active').toLowerCase().trim();
      var email  = (row['Email']    || '').trim();
      var name   = (row['Full Name']|| '').trim();
      if (!uname || role === 'admin' || status === 'inactive') return;
      if (!email) email = USER_EMAILS[uname] || '';
      if (!email) return; // skip if no email
      userList.push({ username: uname, name: name, role: role, email: email });
    });

    // Fill in any hardcoded non-admin users not yet in sheet
    Object.keys(USER_EMAILS).forEach(function(uname) {
      if (uname === 'admin') return;
      var alreadyAdded = userList.some(function(u) { return u.username === uname; });
      if (!alreadyAdded) {
        // Determine name from email prefix or username
        var name = uname.charAt(0).toUpperCase() + uname.slice(1);
        userList.push({ username: uname, name: name, role: 'tech', email: USER_EMAILS[uname] });
      }
    });

    if (userList.length === 0) {
      Logger.log('⚠️ No non-admin users with email addresses found. Add emails in User Management.');
      return ok({ success: false, message: 'No recipients found' });
    }

    Logger.log('📧 Sending pending task reminders to ' + userList.length + ' user(s)…');

    // ── Send personalised email to each user ──
    userList.forEach(function(user) {
      try {
        var firstName = user.name.split(' ')[0] || user.username;
        var firstNameLower = firstName.toLowerCase();

        // Filter tickets assigned to this user (match by first name in Assigned To)
        var myTickets = allTickets.filter(function(t) {
          var assignedTo = (t.assignedTo || '').toLowerCase();
          return assignedTo.includes(firstNameLower);
        });

        var myOpen     = myTickets.filter(function(t) { return t.status === 'Open'; });
        var myInProg   = myTickets.filter(function(t) { return t.status === 'In Progress'; });
        var myOverdue  = myTickets.filter(function(t) {
          return (t.status === 'Open' || t.status === 'In Progress') && t.date && t.date < todayStr;
        });
        var myPending  = myTickets.filter(function(t) {
          return t.status === 'Open' || t.status === 'In Progress' || t.status === 'Pending Vendor' || t.status === 'On Hold';
        });

        // Filter WOs coordinated by this user
        var myWOs = allWOs.filter(function(w) {
          var coord = (w['Coordinator'] || w.coordinator || '').toLowerCase();
          return coord.includes(firstNameLower);
        }).filter(function(w) {
          var st = (w['Status'] || w.status || '').toLowerCase();
          return st.includes('pending') || st.includes('in progress') || st.includes('progress');
        });

        // Skip users with no pending tasks
        if (myPending.length === 0 && myWOs.length === 0) {
          Logger.log('  ⏭ ' + user.email + ' — no pending tasks, skipping');
          skippedCount++;
          return;
        }

        // ── Build personalised HTML email ──
        var overdueWarning = '';
        if (myOverdue.length > 0) {
          overdueWarning = '<div style="background:#fff5f5;border-left:4px solid #ef4444;' +
            'border-radius:0 6px 6px 0;padding:10px 14px;margin-bottom:16px">' +
            '<b style="color:#ef4444">⚠️ ' + myOverdue.length + ' Overdue Task(s)</b>' +
            '<p style="color:#7f1d1d;font-size:12px;margin:4px 0 0">These tickets are past their log date and still open. Please update or resolve them today.</p>' +
            '</div>';
        }

        var html = buildReminderHTML(user, firstName, myPending, myWOs, myOverdue, overdueWarning, todayStr);

        var subject = '🔔 Your Pending IT Tasks – ' + todayStr +
          (myOverdue.length > 0 ? ' ⚠️ ' + myOverdue.length + ' Overdue' : '') +
          ' | ' + myPending.length + ' Task(s) Pending';

        MailApp.sendEmail({
          to:       user.email,
          replyTo:  FROM_EMAIL,
          from:     FROM_EMAIL,
          name:     'MIT ACSC IT Section',
          subject:  subject,
          htmlBody: html,
          body:     buildPlainReminderText(firstName, myPending, myWOs, myOverdue, todayStr)
        });

        Logger.log('  ✅ Sent to ' + user.email + ' (' + firstName + ') — ' + myPending.length + ' tasks, ' + myOverdue.length + ' overdue');
        sentCount++;

      } catch(userErr) {
        Logger.log('  ❌ Failed for ' + user.email + ': ' + userErr.message);
        errors.push(user.email + ': ' + userErr.message);
      }
    });

    Logger.log('════════════════════════════════════');
    Logger.log('✅ Reminder emails sent: ' + sentCount);
    Logger.log('⏭  Skipped (no tasks): '  + skippedCount);
    if (errors.length) Logger.log('❌ Errors: ' + errors.join(' | '));
    Logger.log('════════════════════════════════════');

    return ok({ success: true, sent: sentCount, skipped: skippedCount, errors: errors });

  } catch(ex) {
    Logger.log('❌ sendDailyPendingTaskReminder error: ' + ex.message);
    return err('sendDailyPendingTaskReminder error: ' + ex.message);
  }
}


// ════════════════════════════════════════════════════════════
//  HTML EMAIL BUILDER — personalised per user
// ════════════════════════════════════════════════════════════

function buildReminderHTML(user, firstName, myPending, myWOs, myOverdue, overdueWarning, todayStr) {
  var totalTasks = myPending.length + myWOs.length;
  var criticalCount = myPending.filter(function(t) { return (t.priority||'').toLowerCase() === 'critical'; }).length;

  var html = [
    '<div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;background:#f4f6f9">',

    // Header
    '<div style="background:linear-gradient(135deg,#8B1840,#B82355);padding:22px 24px;border-radius:10px 10px 0 0">',
    '<table style="width:100%"><tr>',
    '<td><h2 style="color:#fff;margin:0;font-size:19px">🔔 Daily Task Reminder</h2>',
    '<p style="color:#f0c4d4;margin:4px 0 0;font-size:12px">MIT ACSC IT Section | ' + todayStr + '</p></td>',
    '<td style="text-align:right">',
    '<div style="background:rgba(255,255,255,.15);border-radius:8px;padding:8px 14px;display:inline-block">',
    '<div style="font-size:22px;font-weight:700;color:#fff">' + totalTasks + '</div>',
    '<div style="font-size:10px;color:#f0c4d4">PENDING</div>',
    '</div></td></tr></table>',
    '</div>',

    // Body
    '<div style="background:#fff;padding:22px 24px;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 10px 10px">',

    // Greeting
    '<p style="font-size:15px;color:#1f2937;margin:0 0 6px">Dear <b>' + firstName + '</b>,</p>',
    '<p style="font-size:13px;color:#6b7280;margin:0 0 18px">',
    'This is your daily reminder from MIT ACSC IT Section. ',
    'You have <b style="color:#1F3864">' + myPending.length + ' ticket(s)</b>',
    (myWOs.length > 0 ? ' and <b style="color:#1F3864">' + myWOs.length + ' vendor WO(s)</b>' : ''),
    ' pending today. Please update or resolve them at the earliest.',
    '</p>',

    // Overdue warning
    overdueWarning,

    // KPI summary row
    '<table style="width:100%;border-collapse:collapse;margin-bottom:20px;text-align:center"><tr>',
    reminderKPI('📋 Open',        myPending.filter(function(t){return t.status==='Open';}).length,         '#3b82f6'),
    reminderKPI('🔄 In Progress', myPending.filter(function(t){return t.status==='In Progress';}).length,  '#8b5cf6'),
    reminderKPI('🔴 Critical',    criticalCount,                                                            '#ef4444'),
    reminderKPI('⏰ Overdue',     myOverdue.length,                                                         '#f97316'),
    '</tr></table>',
  ].join('');

  // Pending tickets table
  if (myPending.length > 0) {
    html += '<div style="margin-bottom:8px"><b style="font-size:14px;color:#1F3864">📋 Your Pending Tickets</b></div>';
    html += '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px">';
    html += '<thead style="background:#8B1840;color:#fff"><tr>';
    ['Ticket ID','Date','Department','Category','Priority','Status','Description'].forEach(function(h) {
      html += '<th style="padding:7px 8px;text-align:left;font-size:11px;white-space:nowrap">' + h + '</th>';
    });
    html += '</tr></thead><tbody>';

    myPending.forEach(function(t) {
      var isOverdue   = t.date && t.date < todayStr && (t.status==='Open'||t.status==='In Progress');
      var rowBg       = isOverdue ? '#fff5f5' : (myPending.indexOf(t)%2===0?'#fff':'#f9fafb');
      var priColor    = t.priority==='Critical'?'#ef4444':t.priority==='High'?'#f97316':t.priority==='Medium'?'#f59e0b':'#10b981';
      var stColor     = t.status==='Open'?'#3b82f6':t.status==='In Progress'?'#8b5cf6':'#f59e0b';

      html += '<tr style="background:' + rowBg + ';border-bottom:1px solid #e5e7eb">';
      html += '<td style="padding:7px 8px;font-weight:700;color:#1F3864;white-space:nowrap">' + (t.ticketId||'—') + (isOverdue?' <span style="color:#ef4444;font-size:9px">OVERDUE</span>':'') + '</td>';
      html += '<td style="padding:7px 8px;white-space:nowrap">' + (t.date||'—') + '</td>';
      html += '<td style="padding:7px 8px">' + (t.dept||'—').substring(0,22) + '</td>';
      html += '<td style="padding:7px 8px">' + (t.category||'—') + '</td>';
      html += '<td style="padding:7px 8px;font-weight:700;color:' + priColor + '">' + (t.priority||'—') + '</td>';
      html += '<td style="padding:7px 8px;font-weight:700;color:' + stColor + '">' + (t.status||'—') + '</td>';
      html += '<td style="padding:7px 8px">' + (t.description||'—').substring(0,55) + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
  }

  // Vendor WOs table
  if (myWOs.length > 0) {
    html += '<div style="margin-bottom:8px"><b style="font-size:14px;color:#1F3864">🏭 Your Pending Vendor Work Orders</b></div>';
    html += '<div style="overflow-x:auto"><table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px">';
    html += '<thead style="background:#1F3864;color:#fff"><tr>';
    ['WO ID','Date','Vendor','Description','Status'].forEach(function(h) {
      html += '<th style="padding:7px 8px;text-align:left;font-size:11px">' + h + '</th>';
    });
    html += '</tr></thead><tbody>';
    myWOs.forEach(function(w, i) {
      html += '<tr style="background:' + (i%2===0?'#fff':'#f9fafb') + ';border-bottom:1px solid #e5e7eb">';
      html += '<td style="padding:7px 8px;font-weight:700">' + (w['WO ID']||w.woId||'—') + '</td>';
      html += '<td style="padding:7px 8px">' + (w['Date']||w.date||'—') + '</td>';
      html += '<td style="padding:7px 8px">' + (w['Vendor Name']||w.vendorName||'—') + '</td>';
      html += '<td style="padding:7px 8px">' + (w['Description']||w.description||'—').substring(0,50) + '</td>';
      html += '<td style="padding:7px 8px;color:#f59e0b;font-weight:700">' + (w['Status']||w.status||'—') + '</td>';
      html += '</tr>';
    });
    html += '</tbody></table></div>';
  }

  // Action guidance
  html += [
    '<div style="background:#f0f9ff;border:1px solid #bae6fd;border-radius:8px;padding:14px 16px;margin-bottom:16px">',
    '<b style="color:#0369a1;font-size:13px">💡 Action Required</b>',
    '<ul style="margin:8px 0 0 16px;color:#0c4a6e;font-size:12px;line-height:1.8">',
    '<li>Log in to the <b>MIT IT Portal</b> to update ticket statuses</li>',
    '<li>Mark resolved tasks as <b>Resolved</b> or <b>Closed</b></li>',
    '<li>Add remarks for any tasks that need vendor support</li>',
    myOverdue.length > 0 ? '<li style="color:#dc2626"><b>Address overdue tickets immediately</b></li>' : '',
    '</ul></div>',

    // Footer
    '<div style="padding-top:14px;border-top:1px solid #e5e7eb;font-size:11px;color:#9ca3af">',
    '<table style="width:100%"><tr>',
    '<td><b style="color:#8B1840">MIT Arts, Commerce &amp; Science College</b><br/>',
    'Alandi, Pune – 412105 | IT Section<br/>',
    'This is an automated daily reminder sent at 8:00 AM</td>',
    '<td style="text-align:right;vertical-align:top">',
    '<a href="mailto:' + FROM_EMAIL + '" style="color:#8B1840;text-decoration:none">' + FROM_EMAIL + '</a><br/>',
    '<span style="color:#d1d5db">Reply to report issues</span>',
    '</td></tr></table></div>',
    '</div></div>'
  ].join('');

  return html;
}

function reminderKPI(label, val, color) {
  return '<td style="padding:8px;background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;margin:2px">' +
    '<div style="font-size:20px;font-weight:700;color:' + color + '">' + val + '</div>' +
    '<div style="font-size:10px;color:#6b7280;margin-top:2px">' + label + '</div></td>';
}

function buildPlainReminderText(firstName, myPending, myWOs, myOverdue, todayStr) {
  var lines = [
    'MIT ACSC IT Section – Daily Pending Task Reminder',
    'Date: ' + todayStr,
    '────────────────────────────────',
    'Dear ' + firstName + ',',
    '',
    'You have ' + myPending.length + ' pending ticket(s)' + (myWOs.length ? ' and ' + myWOs.length + ' vendor WO(s)' : '') + '.',
    ''
  ];
  if (myOverdue.length > 0) {
    lines.push('⚠️ OVERDUE TASKS (' + myOverdue.length + '):');
    myOverdue.forEach(function(t) {
      lines.push('  [OVERDUE] ' + (t.ticketId||'—') + ' – ' + (t.description||'').substring(0,60) + ' [' + (t.priority||'') + ']');
    });
    lines.push('');
  }
  if (myPending.length > 0) {
    lines.push('PENDING TICKETS:');
    myPending.forEach(function(t) {
      lines.push('  [' + (t.priority||'?') + '/' + (t.status||'?') + '] ' + (t.ticketId||'—') + ' – ' + (t.description||'').substring(0,60));
      lines.push('    Dept: ' + (t.dept||'—') + ' | Category: ' + (t.category||'—') + ' | Date: ' + (t.date||'—'));
    });
    lines.push('');
  }
  if (myWOs.length > 0) {
    lines.push('VENDOR WORK ORDERS:');
    myWOs.forEach(function(w) {
      lines.push('  ' + (w['WO ID']||'—') + ' – ' + (w['Vendor Name']||'—') + ' | ' + (w['Status']||'—'));
    });
    lines.push('');
  }
  lines.push('Please log in to the MIT IT Portal to update your task statuses.');
  lines.push('────────────────────────────────');
  lines.push('MIT ACSC IT Section | Alandi, Pune – 412105');
  lines.push('Reply to: ' + FROM_EMAIL);
  return lines.join('\n');
}


// ════════════════════════════════════════════════════════════
//  TRIGGER SETUP FOR PENDING TASK REMINDER
//  Run setupPendingReminderTrigger() ONCE from GAS editor.
//  Fires every day at 8 AM (after the 7 AM admin report).
// ════════════════════════════════════════════════════════════

function setupPendingReminderTrigger() {
  // Remove any existing pending reminder triggers
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'sendDailyPendingTaskReminder') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted existing pending reminder trigger');
    }
  });

  // Create new trigger at 8 AM daily
  ScriptApp.newTrigger('sendDailyPendingTaskReminder')
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();

  Logger.log('✅ Pending task reminder trigger created — fires every day at 8 AM');
  Logger.log('From: ' + FROM_EMAIL);
  Logger.log('To: Each non-admin user with a registered email address');
  Logger.log('To test immediately, run: testSendPendingReminder()');
}

// ════════════════════════════════════════════════════════════
//  TEST FUNCTIONS
// ════════════════════════════════════════════════════════════

// Test: sends reminder to ALL non-admin users right now
function testSendPendingReminder() {
  Logger.log('=== testSendPendingReminder ===');
  var result = sendDailyPendingTaskReminder();
  try {
    var r = JSON.parse(result.getContent());
    Logger.log('Sent: ' + r.sent + ' | Skipped: ' + r.skipped);
    if (r.errors && r.errors.length) Logger.log('Errors: ' + r.errors.join(', '));
  } catch(e) {
    Logger.log('Result: ' + result.getContent());
  }
}

// Test: sends reminder to a SINGLE user by username (for debugging)
function testReminderForUser() {
  var targetUsername = 'rutuj'; // ← change to the username you want to test

  // Temporarily patch userList to only include this one user
  var allTickets = readSheet(SHEETS.tickets).map(normalizeTicketRow);
  var allWOs     = readSheet(SHEETS.wos);
  var usersSheet = readSheet(SHEETS.users || 'Users');
  var todayStr   = today();

  var targetRow = usersSheet.filter(function(r) {
    return (r['Username']||'').toLowerCase() === targetUsername.toLowerCase();
  })[0];

  if (!targetRow && !USER_EMAILS[targetUsername]) {
    Logger.log('❌ User not found: ' + targetUsername);
    return;
  }

  var user = {
    username: targetUsername,
    name:     (targetRow && targetRow['Full Name']) || (targetUsername.charAt(0).toUpperCase() + targetUsername.slice(1)),
    email:    (targetRow && targetRow['Email']) || USER_EMAILS[targetUsername] || '',
    role:     'tech'
  };

  if (!user.email) { Logger.log('❌ No email for user: ' + targetUsername); return; }

  var firstName = user.name.split(' ')[0];
  var firstNameLower = firstName.toLowerCase();

  var myPending = allTickets.filter(function(t) {
    var a = (t.assignedTo||'').toLowerCase();
    return a.includes(firstNameLower) && (t.status==='Open'||t.status==='In Progress'||t.status==='Pending Vendor'||t.status==='On Hold');
  });
  var myWOs = allWOs.filter(function(w) {
    var c = (w['Coordinator']||'').toLowerCase();
    return c.includes(firstNameLower);
  });
  var myOverdue = myPending.filter(function(t) { return t.date && t.date < todayStr; });

  Logger.log('User: ' + user.name + ' <' + user.email + '>');
  Logger.log('Pending tickets: ' + myPending.length + ' | WOs: ' + myWOs.length + ' | Overdue: ' + myOverdue.length);

  if (myPending.length === 0 && myWOs.length === 0) {
    Logger.log('No pending tasks — test email will not be sent (as per logic: skip users with no tasks)');
    Logger.log('Add a test ticket assigned to "' + user.name + '" first, then re-run.');
    return;
  }

  var overdueWarning = myOverdue.length > 0 ? '<div style="color:red">⚠️ ' + myOverdue.length + ' overdue</div>' : '';
  var html = buildReminderHTML(user, firstName, myPending, myWOs, myOverdue, overdueWarning, todayStr);

  MailApp.sendEmail({
    to:       user.email,
    replyTo:  FROM_EMAIL,
    name:     'MIT ACSC IT Section',
    subject:  '[TEST] Daily Task Reminder for ' + firstName + ' — ' + todayStr,
    htmlBody: html,
    body:     buildPlainReminderText(firstName, myPending, myWOs, myOverdue, todayStr)
  });

  Logger.log('✅ Test reminder sent to: ' + user.email);
}
