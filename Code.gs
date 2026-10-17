/**
 * ============================================================
 * DJ-CE AIRCON & ELECTRICAL SERVICES
 * Google Apps Script Backend — v4.0
 * New: Schedule Slots, Client Confirmation Email, Form Validation
 * ============================================================
 */

/* ─── CONFIG ──────────────────────────────────────────────── */
var SHEET_ID      = '1dbMhoFuINB-boqITMExPw9cH_MIrJ4wpqQjwjGcX8S4';
var NOTIF_EMAIL   = 'djceservicesforyou@gmail.com';
var BUSINESS_NAME = 'DJ-CE Aircon & Electrical Services';
var BUSINESS_PHONE = '0976 259 2458';
var BUSINESS_ADDRESS = 'Caniogan, Calumpit, Bulacan';
var TIMEZONE      = 'Asia/Manila';

/* ─── SHEET HEADERS ───────────────────────────────────────── */
var HEADERS = {
  Bookings: [
    'Booking ID','Ref No','Created At','Name','Phone','Email',
    'House No','Street','Barangay','City','Full Address','Landmark','Lat','Lng',
    'Service','Service Details','Date','Time Slot',
    'Technician','Status','Amount','Fee Service','Fee Parts','Fee Transpo','Fee Other','Fee Discount',
    'Usage Frequency','Next Cleaning Date','Notes','Source','Followed Up','Completed At',
    'Deleted At','Deleted By','Archived At'
  ],
  Accounts: [
    'Username','Password','Name','Role','Avatar',
    'Last Pw Change','Force Change','Created At','Last Login','PIN'
  ],
  Clients: [
    'Client ID','Name','Phone','Email','Address',
    'Last Service','Last Service Date','Usage Frequency',
    'Next Cleaning Date','Notes','Source','Created At'
  ],
  Complaints: [
    'Complaint ID','Title','Client Name','Booking ID','Type',
    'Priority','Description','Status','Resolution',
    'Created By','Created At','Resolved By','Resolved At'
  ],
  ActivityLog: [
    'Timestamp','User','Action','Type','Details'
  ],
  Expenses: [
    'Expense ID','Date','Category','Description','Amount','Ref No','Added By','Created At'
  ],
  RefNumbers: [
    'Key','Monthly Counter','Total Counter','Month','Year','Last Updated'
  ],
  ScheduleSlots: [
    'Slot ID','Date','Time Slot','Max Bookings','Current Bookings','Status','Created By','Created At','Notes'
  ],
  Employees: [
    'Employee ID','Name','Role','Type','Daily Rate','Status','Phone','Start Date','Notes','Created At'
  ],
  DailyTracker: [
    'Record ID','Date','Employee ID','Employee Name','Booking IDs','Jobs Count','Classification','Audited By','Audit Notes','Created At'
  ],
  Payroll: [
    'Payroll ID','Period From','Period To','Employee ID','Employee Name','Days Worked','Half Days','Full Days','OT Days','Daily Rate','Gross Pay','Deductions','Net Pay','13th Month Accrual','Status','Generated At','Approved By','Notes'
  ],
  InvestorShares: [
    'Record ID','Investor Name','Share Percent','Year','Month','Monthly Ipon','Notes','Created At'
  ]
};

var SHEET_STYLES = {
  Bookings:      { bg: '#0d1f3c', fg: '#ffffff' },
  Accounts:      { bg: '#0d1f3c', fg: '#f59e0b' },
  Clients:       { bg: '#065f46', fg: '#ffffff' },
  Complaints:    { bg: '#7f1d1d', fg: '#ffffff' },
  ActivityLog:   { bg: '#1e3a5f', fg: '#ffffff' },
  Expenses:      { bg: '#7f1d1d', fg: '#ffffff' },
  RefNumbers:    { bg: '#064e3b', fg: '#ffffff' },
  ScheduleSlots: { bg: '#1e3a5f', fg: '#f59e0b' },
  Employees:     { bg: '#065f46', fg: '#f59e0b' },
  DailyTracker:  { bg: '#1e3a5f', fg: '#ffffff' },
  Payroll:       { bg: '#7f1d1d', fg: '#f59e0b' },
  InvestorShares:{ bg: '#064e3b', fg: '#f59e0b' }
};

/* ─── TIME SLOTS ──────────────────────────────────────────── */
var TIME_SLOTS = {
  'morning':   { label: 'Morning',   time: '9:00 AM – 12:00 PM' },
  'afternoon': { label: 'Afternoon', time: '1:00 PM – 4:00 PM'  },
  'evening':   { label: 'Evening',   time: '5:00 PM – 7:00 PM'  }
};
/* ─── DRIVE PHOTO FOLDER ─────────────────────────────────── */
var PHOTO_FOLDER_NAME = 'DJCE_Job_Photos';

function getOrCreatePhotoFolder() {
  var folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}
/* ─── ENTRY POINT ─────────────────────────────────────────── */
function doGet(e) {
  var result;
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || '';
    var cb     = params.cb     || '';

    if (!action) {
      result = { ok: true, name: 'DJ-CE Backend v4.1', time: now() };
    } else {
      switch (action) {
        case 'submitBooking':       result = submitBooking(params);       break;
        case 'getBookings':         result = getBookings(params);         break;
        case 'updateBooking':       result = updateBooking(params);       break;
        case 'getAccounts':         result = getAccounts();               break;
        case 'saveAccount':         result = saveAccount(params);         break;
        case 'deleteAccount':       result = deleteAccount(params);       break;
        case 'getClients':          result = getClients();                break;
        case 'saveClient':          result = saveClient(params);          break;
        case 'deleteClient':        result = deleteClient(params);        break;
        case 'getComplaints':       result = getComplaints();             break;
        case 'saveComplaint':       result = saveComplaint(params);       break;
        case 'logActivity':         result = logActivity(params);         break;
        case 'saveExpense':         result = saveExpense(params);         break;
        case 'getExpenses':         result = getExpenses();               break;
        case 'deleteExpense':       result = deleteExpense(params);      break;
        case 'getBookingHistory':  result = getBookingHistory(params); break;
        case 'getNextRefNumber':    result = getNextRefNumber(params);    break;
        case 'getRefCounters':      result = getRefCounters();            break;
        case 'getAvailableSlots':   result = getAvailableSlots(params);   break;
        case 'openSlot':            result = openSlot(params);            break;
        case 'closeSlot':           result = closeSlot(params);           break;
        case 'deleteSlot':          result = deleteSlot(params);          break;
        case 'getAllSlots':          result = getAllSlots(params);         break;
        case 'getEmployees':        result = getEmployees();              break;
        case 'saveEmployee':        result = saveEmployee(params);        break;
        case 'deleteEmployee':      result = deleteEmployee(params);      break;
        case 'getDailyTracker':     result = getDailyTracker(params);     break;
        case 'saveDailyRecord':     result = saveDailyRecord(params);     break;
        case 'updateDailyRecord':   result = updateDailyRecord(params);   break;
        case 'getPayroll':          result = getPayroll(params);          break;
        case 'savePayroll':         result = savePayroll(params);         break;
        case 'getInvestors':        result = getInvestors();              break;
        case 'saveInvestor':        result = saveInvestor(params);        break;
        case 'deleteInvestor':      result = deleteInvestor(params);      break;
        case 'rescheduleBooking':    result = rescheduleBooking(params);   break;
        case 'sendReminder':        result = sendManualReminder(params);  break;
        case 'getSheetHeaders':     result = getSheetHeaders(params);     break;
        case 'fixBookingHeaders':   result = fixBookingHeaders();         break;
        default: result = { ok: false, error: 'Unknown action: ' + action };
      }
    }
  } catch (err) {
    result = { ok: false, error: err.message || String(err) };
  }

  var json = JSON.stringify(result);
  var output;
  if (cb) {
    output = ContentService.createTextOutput(cb + '(' + json + ')');
    output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  } else {
    output = ContentService.createTextOutput(json);
    output.setMimeType(ContentService.MimeType.JSON);
  }
  return output;
}

/* ─── BOOKINGS ────────────────────────────────────────────── */
function submitBooking(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Bookings');
  var id = data.id || ('BK-' + new Date().getTime());
  var ts = now();

  // Prevent duplicates
  var rows = sheet.getDataRange().getValues();
  for (var r = 1; r < rows.length; r++) {
    if (String(rows[r][0]) === String(id)) {
      return { ok: true, id: id, timestamp: ts, note: 'already exists' };
    }
  }

  var svcSummary = '';
  var svcDetails = '';
  if (data.services && data.services.length) {
    svcSummary = data.services.map(function(s) {
      var d = (s.details || []).filter(Boolean);
      return s.name + (d.length ? ' (' + d.join(', ') + ')' : '');
    }).join(' | ');
    svcDetails = data.services.map(function(s){ return (s.details||[]).join(', '); }).join(' | ');
  } else if (data.service) {
    svcSummary = data.service;
    svcDetails = data.serviceDetailsSummary || data.serviceDetails || '';
  }

  // Build full address from parts
  var fullAddress = [data.houseNo, data.street, data.barangay, data.city]
    .filter(Boolean).join(', ');
  if (!fullAddress) fullAddress = data.address || '';

  var row = [
    id,
    data.refNo        || id,
    ts,
    data.name         || '',
    data.phone        || '',
    data.email        || '',
    data.houseNo      || '',
    data.street       || '',
    data.barangay     || '',
    data.city         || '',
    fullAddress,
    data.landmark     || '',
    data.lat          || '',
    data.lng          || '',
    svcSummary,
    svcDetails,
    data.date         || '',
    data.timeSlot     || data.time || '',
    data.technician   || '',
    'Pending',
    data.amount       || '',
    data.feeService   || '',
    data.feeParts     || '',
    data.feeTranspo   || '',
    data.feeOther     || '',
    data.feeDiscount  || '',
    data.usage        || '',
    data.nextCleaningDate || '',
    data.notes        || '',
    data.source       || 'website',
    'false',
    '',
    '', '', ''
  ];

  sheet.appendRow(row);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 20).setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('bold');

  // Increment slot booking count
  if (data.date && data.timeSlot) {
    try { incrementSlotCount(data.date, data.timeSlot); } catch(e) {}
  }

  // Send emails
  try { sendAdminNotifEmail(id, data, svcSummary, fullAddress); } catch(e) {}
  try {
    if (data.email) sendClientConfirmationEmail(id, data, svcSummary, fullAddress);
  } catch(e) {}

  return { ok: true, id: id, timestamp: ts };
}

function getBookings(p) {
  var sheet = getOrCreateSheet('Bookings');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, bookings: [] };
  var headers = rows[0];
  var bookings = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = cellToString(row[i]); });
    return obj;
  });
  if (p && p.status) {
    bookings = bookings.filter(function(b) { return b.status === p.status; });
  }
  return { ok: true, bookings: bookings };
}

function updateBooking(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Bookings');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });

      var oldStatus = String(rows[i][headers.indexOf('Status')] || '');
      var newStatus = data.status || oldStatus;

      if (data.status)      setCellByKey(sheet, i+1, colMap, 'status', data.status);
      if (data.technician)  setCellByKey(sheet, i+1, colMap, 'technician', data.technician);
      if (data.amount)      setCellByKey(sheet, i+1, colMap, 'amount', data.amount);
      if (data.feeService)  setCellByKey(sheet, i+1, colMap, 'feeService', data.feeService);
      if (data.feeParts)    setCellByKey(sheet, i+1, colMap, 'feeParts', data.feeParts);
      if (data.feeTranspo)  setCellByKey(sheet, i+1, colMap, 'feeTranspo', data.feeTranspo);
      if (data.feeOther)    setCellByKey(sheet, i+1, colMap, 'feeOther', data.feeOther);
      if (data.feeDiscount) setCellByKey(sheet, i+1, colMap, 'feeDiscount', data.feeDiscount);
      if (data.notes)       setCellByKey(sheet, i+1, colMap, 'notes', data.notes);
      if (data.usage)       setCellByKey(sheet, i+1, colMap, 'usageFrequency', data.usage);
      if (data.nextCleaningDate) setCellByKey(sheet, i+1, colMap, 'nextCleaningDate', data.nextCleaningDate);
      if (data.completedAt) setCellByKey(sheet, i+1, colMap, 'completedAt', data.completedAt);
      if (data.deletedAt)   setCellByKey(sheet, i+1, colMap, 'deletedAt', data.deletedAt);
      if (data.deletedBy)   setCellByKey(sheet, i+1, colMap, 'deletedBy', data.deletedBy);
      if (data.archivedAt)  setCellByKey(sheet, i+1, colMap, 'archivedAt', data.archivedAt);
      if (data.followedUp)  setCellByKey(sheet, i+1, colMap, 'followedUp', data.followedUp);

      // If status changed to Cancelled → free up the slot
      if (oldStatus !== 'Cancelled' && newStatus === 'Cancelled') {
        var bDate = String(rows[i][headers.indexOf('Date')] || '');
        var bSlot = String(rows[i][headers.indexOf('Time Slot')] || '');
        if (bDate && bSlot) {
          try { decrementSlotCount(bDate, bSlot); } catch(e) {}
        }
      }

      // If status changed to Confirmed → send confirmation email to client
      if (oldStatus !== 'Confirmed' && newStatus === 'Confirmed') {
        try {
          var bEmail = data.email || cellToString(rows[i][headers.indexOf('Email')]);
          var bName  = data.name  || cellToString(rows[i][headers.indexOf('Name')]);
          var bPhone = data.phone || cellToString(rows[i][headers.indexOf('Phone')]);
          var bSvc   = data.service || cellToString(rows[i][headers.indexOf('Service')]);
          var bDate  = data.date  || cellToString(rows[i][headers.indexOf('Date')]);
          var bSlot  = data.time  || cellToString(rows[i][headers.indexOf('Time Slot')]);
          var bAddr  = data.address || cellToString(rows[i][headers.indexOf('Full Address')]);
          var bTech  = data.technician || cellToString(rows[i][headers.indexOf('Technician')]);
          var bRef   = data.refNo || cellToString(rows[i][headers.indexOf('Ref No')]) || data.id;
          if (bEmail) {
            sendBookingConfirmedEmail(bRef, bName, bEmail, bSvc, bDate, bSlot, bAddr, bTech);
          }
        } catch(emailErr) { /* don't fail the update if email fails */ }
      }

      // Color code status
      var statusColors = {
        'Pending':     { bg: '#fef3c7', fg: '#92400e' },
        'Confirmed':   { bg: '#dbeafe', fg: '#1e40af' },
        'Ongoing':     { bg: '#d1fae5', fg: '#065f46' },
        'Done':        { bg: '#ecfdf5', fg: '#064e3b' },
        'For Backjob': { bg: '#fce7f3', fg: '#9d174d' },
        'Cancelled':   { bg: '#fee2e2', fg: '#991b1b' }
      };
      if (data.status && statusColors[data.status]) {
        var sc = statusColors[data.status];
        var statusCol = colMap['status'];
        if (statusCol) sheet.getRange(i+1, statusCol).setBackground(sc.bg).setFontColor(sc.fg).setFontWeight('bold');
      }

      return { ok: true, id: data.id };
    }
  }
  return { ok: false, error: 'Booking not found: ' + data.id };
}

/* ─── SCHEDULE SLOTS ──────────────────────────────────────── */

/**
 * Get all available slots for public booking form
 * Returns only Open slots with remaining capacity
 * Optional: filter by date range
 */
function getAvailableSlots(p) {
  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, slots: [] };
  var headers = rows[0];

  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var fromDate = p.from || today;
  var toDate   = p.to   || '';

  var slots = [];
  for (var i = 1; i < rows.length; i++) {
    var obj = {};
    headers.forEach(function(h, j) { obj[headerToKey(h)] = cellToString(rows[i][j]); });

    // Only return open slots from today onwards
    if (obj.status !== 'Open') continue;
    if (obj.date < fromDate) continue;
    if (toDate && obj.date > toDate) continue;

    var max = parseInt(obj.maxBookings) || 0;
    var cur = parseInt(obj.currentBookings) || 0;
    var remaining = max - cur;
    if (remaining <= 0) continue; // fully booked

    var slotInfo = TIME_SLOTS[obj.timeSlot] || { label: obj.timeSlot, time: obj.timeSlot };
    slots.push({
      slotId:     obj.slotId,
      date:       obj.date,
      timeSlot:   obj.timeSlot,
      label:      slotInfo.label,
      time:       slotInfo.time,
      remaining:  remaining,
      maxBookings: max,
      currentBookings: cur
    });
  }

  // Sort by date then time slot
  var slotOrder = { morning: 0, afternoon: 1, evening: 2 };
  slots.sort(function(a, b) {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return (slotOrder[a.timeSlot] || 0) - (slotOrder[b.timeSlot] || 0);
  });

  return { ok: true, slots: slots };
}

/**
 * Get all slots (for portal admin view)
 */
function getAllSlots(p) {
  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, slots: [] };
  var headers = rows[0];

  var today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  var fromDate = p && p.from ? p.from : '';

  var slots = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, j) { obj[headerToKey(h)] = cellToString(row[j]); });
    var slotInfo = TIME_SLOTS[obj.timeSlot] || { label: obj.timeSlot, time: obj.timeSlot };
    obj.label = slotInfo.label;
    obj.time  = slotInfo.time;
    obj.remaining = Math.max(0, (parseInt(obj.maxBookings)||0) - (parseInt(obj.currentBookings)||0));
    return obj;
  }).filter(function(s) {
    if (fromDate && s.date < fromDate) return false;
    return true;
  });

  slots.sort(function(a, b) {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    var slotOrder = { morning: 0, afternoon: 1, evening: 2 };
    return (slotOrder[a.timeSlot] || 0) - (slotOrder[b.timeSlot] || 0);
  });

  return { ok: true, slots: slots };
}

/**
 * Open a new slot (Admin/Owner only)
 */
function openSlot(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  if (!data.date || !data.timeSlot) return { ok: false, error: 'Date and timeSlot required' };

  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();

  // Check if slot already exists for this date+timeSlot
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(data.date) && String(rows[i][2]) === String(data.timeSlot)) {
      // Re-open if closed
      sheet.getRange(i+1, 6).setValue('Open');
      if (data.maxBookings) sheet.getRange(i+1, 4).setValue(parseInt(data.maxBookings) || 2);
      return { ok: true, action: 'reopened', slotId: String(rows[i][0]) };
    }
  }

  // Create new slot
  var slotId = 'SL-' + new Date().getTime();
  sheet.appendRow([
    slotId,
    data.date,
    data.timeSlot,
    parseInt(data.maxBookings) || 2,
    0,
    'Open',
    data.createdBy || 'Admin',
    now(),
    data.notes || ''
  ]);

  return { ok: true, action: 'created', slotId: slotId };
}

/**
 * Close a slot (Admin/Owner)
 */
function closeSlot(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.slotId) ||
        (String(rows[i][1]) === String(data.date) && String(rows[i][2]) === String(data.timeSlot))) {
      sheet.getRange(i+1, 6).setValue('Closed');
      return { ok: true, action: 'closed' };
    }
  }
  return { ok: false, error: 'Slot not found' };
}

/**
 * Delete a slot permanently
 */
function deleteSlot(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.slotId)) {
      sheet.deleteRow(i+1);
      return { ok: true, action: 'deleted' };
    }
  }
  return { ok: false, error: 'Slot not found' };
}

function incrementSlotCount(date, timeSlot) {
  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(date) && String(rows[i][2]) === String(timeSlot)) {
      var cur = parseInt(rows[i][4]) || 0;
      sheet.getRange(i+1, 5).setValue(cur + 1);
      // Auto-close if full
      var max = parseInt(rows[i][3]) || 0;
      if (max > 0 && cur + 1 >= max) {
        sheet.getRange(i+1, 6).setValue('Full');
      }
      return;
    }
  }
}

function decrementSlotCount(date, timeSlot) {
  var sheet = getOrCreateSheet('ScheduleSlots');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(date) && String(rows[i][2]) === String(timeSlot)) {
      var cur = parseInt(rows[i][4]) || 0;
      var newCur = Math.max(0, cur - 1);
      sheet.getRange(i+1, 5).setValue(newCur);
      // Re-open if was Full
      var status = String(rows[i][5] || '');
      var max = parseInt(rows[i][3]) || 0;
      if (status === 'Full' && newCur < max) {
        sheet.getRange(i+1, 6).setValue('Open');
      }
      return;
    }
  }
}

/* ─── ACCOUNTS ────────────────────────────────────────────── */
function getAccounts() {
  var sheet = getOrCreateSheet('Accounts');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, accounts: [] };
  var headers = rows[0];
  var accounts = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = cellToString(row[i]); });
    return obj;
  });
  return { ok: true, accounts: accounts };
}

function saveAccount(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Accounts');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.username)) {
      if (data.password)     sheet.getRange(i+1,2).setValue(data.password);
      if (data.name)         sheet.getRange(i+1,3).setValue(data.name);
      if (data.role)         sheet.getRange(i+1,4).setValue(data.role);
      if (data.avatar)       sheet.getRange(i+1,5).setValue(data.avatar);
      if (data.lastPwChange) sheet.getRange(i+1,6).setValue(data.lastPwChange);
      if (data.forceChange !== undefined) sheet.getRange(i+1,7).setValue(data.forceChange);
      if (data.lastLogin)    sheet.getRange(i+1,9).setValue(data.lastLogin);
      if (data.pin !== undefined) sheet.getRange(i+1,10).setValue(data.pin);
      return { ok: true, action: 'updated', username: data.username };
    }
  }
  sheet.appendRow([
    data.username, data.password||'', data.name||'', data.role||'',
    data.avatar||'?', data.lastPwChange||'', data.forceChange||'false',
    now(), '', data.pin||''
  ]);
  return { ok: true, action: 'created', username: data.username };
}

function deleteAccount(p) {
  var sheet = getOrCreateSheet('Accounts');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(p.username)) {
      sheet.deleteRow(i+1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Account not found' };
}

/* ─── CLIENTS ─────────────────────────────────────────────── */
function getClients() {
  var sheet = getOrCreateSheet('Clients');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, clients: [] };
  var headers = rows[0];
  var clients = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = cellToString(row[i]); });
    return obj;
  });
  return { ok: true, clients: clients };
}

function saveClient(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Clients');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var headers = rows[0];
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });
      ['name','phone','email','address','lastService','lastServiceDate',
       'usageFrequency','nextCleaningDate','notes','source'].forEach(function(k) {
        if (data[k] !== undefined) setCellByKey(sheet, i+1, colMap, k, data[k]);
      });
      return { ok: true, action: 'updated' };
    }
  }
  sheet.appendRow([
    data.id||('CL-'+new Date().getTime()),
    data.name||'', data.phone||'', data.email||'', data.address||'',
    data.lastService||'', data.lastServiceDate||'', data.usageFrequency||'',
    data.nextCleaningDate||'', data.notes||'', data.source||'manual', now()
  ]);
  return { ok: true, action: 'created' };
}

function deleteClient(p) {
  var id = p.id || '';
  if (!id) return { ok: false, error: 'No client ID' };
  var sheet = getOrCreateSheet('Clients');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true, action: 'deleted', id: id };
    }
  }
  return { ok: true, action: 'not_found', id: id };
}

/* ─── COMPLAINTS ──────────────────────────────────────────── */
function getComplaints() {
  var sheet = getOrCreateSheet('Complaints');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, complaints: [] };
  var headers = rows[0];
  return { ok: true, complaints: rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = cellToString(row[i]); });
    return obj;
  })};
}

function saveComplaint(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Complaints');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var headers = rows[0];
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });
      ['status','resolution','resolvedBy','resolvedAt','priority'].forEach(function(k) {
        if (data[k]) setCellByKey(sheet, i+1, colMap, k, data[k]);
      });
      return { ok: true, action: 'updated' };
    }
  }
  sheet.appendRow([
    data.id||('CMP-'+new Date().getTime()),
    data.title||'', data.clientName||'', data.bookingId||'',
    data.type||'Complaint', data.priority||'Normal',
    data.description||'', data.status||'Open', '',
    data.createdBy||'', now(), '', ''
  ]);
  return { ok: true, action: 'created' };
}

/* ─── EXPENSES ────────────────────────────────────────────── */
function getExpenses() {
  var sheet = getOrCreateSheet('Expenses');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, expenses: [] };
  var headers = rows[0];
  return { ok: true, expenses: rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = cellToString(row[i]); });
    return obj;
  })};
}

function saveExpense(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Expenses');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      return { ok: true, action: 'already_exists', id: data.id };
    }
  }
  var id = data.id || ('EXP-' + new Date().getTime());
  sheet.appendRow([
    id, data.date||'', data.category||'', data.description||'',
    data.amount||'', data.ref||'', data.addedBy||'', data.createdAt||now()
  ]);
  return { ok: true, action: 'created', id: id };
}
function deleteExpense(p) {
  var id = p.id;
  if (!id) return { ok: false, error: 'Missing expense ID' };
  var sheet = getOrCreateSheet('Expenses');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true, action: 'deleted', id: id };
    }
  }
  return { ok: false, error: 'Expense not found: ' + id };
}
/* ─── BOOKING HISTORY (QR Sticker Lookup) ─────────────── */
function getBookingHistory(p) {
  var ref = p.ref;
  if (!ref) return { ok: false, error: 'Missing booking ref' };

  var sheet = getOrCreateSheet('Bookings');
  var data  = sheet.getDataRange().getValues();
  var headers = data[0];
  var refCol     = headers.indexOf('Ref No');
  var nameCol    = headers.indexOf('Name');
  var phoneCol   = headers.indexOf('Phone');
  var addressCol = headers.indexOf('Full Address');
  var serviceCol = headers.indexOf('Service');
  var dateCol    = headers.indexOf('Date');
  var slotCol    = headers.indexOf('Time Slot');
  var techCol    = headers.indexOf('Technician');
  var statusCol  = headers.indexOf('Status');
  var notesCol   = headers.indexOf('Notes');

  // Find the target booking first
  var target = null;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][refCol]) === String(ref)) {
      target = {
        ref:          data[i][refCol],
        customerName: data[i][nameCol],
        customerPhone:data[i][phoneCol],
        address:      data[i][addressCol],
        serviceType:  data[i][serviceCol],
        date:         data[i][dateCol] instanceof Date
                        ? Utilities.formatDate(data[i][dateCol], TIMEZONE, 'yyyy-MM-dd')
                        : String(data[i][dateCol]),
        timeSlot:     data[i][slotCol],
        technician:   data[i][techCol],
        status:       data[i][statusCol],
        notes:        notesCol > -1 ? data[i][notesCol] : ''
      };
      break;
    }
  }

  if (!target) return { ok: false, error: 'Booking not found: ' + ref };

  // Find all bookings for same customer + address (service history)
  var history = [];
  for (var j = 1; j < data.length; j++) {
    if (String(data[j][nameCol]).toLowerCase() === String(target.customerName).toLowerCase()
        && String(data[j][addressCol]).toLowerCase() === String(target.address).toLowerCase()) {
      var d = data[j][dateCol];
      var dateStr = (d instanceof Date) ? Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd') : String(d);
      history.push({
        ref:         data[j][refCol],
        serviceType: data[j][serviceCol],
        date:        dateStr,
        timeSlot:    data[j][slotCol],
        technician:  data[j][techCol],
        status:      data[j][statusCol]
      });
    }
  }

  // Sort history by date descending
  history.sort(function(a, b) {
    return new Date(b.date) - new Date(a.date);
  });

  return {
    ok: true,
    booking: target,
    history: history,
    totalServices: history.length
  };
}
/* ─── REFERENCE NUMBERS ───────────────────────────────────── */
function getNextRefNumber(p) {
  var docType     = p.docType     || 'BK';
  var serviceCode = p.serviceCode || 'GN';
  var nowDate     = new Date();
  var month = Utilities.formatDate(nowDate, TIMEZONE, 'MM');
  var year  = Utilities.formatDate(nowDate, TIMEZONE, 'yy');
  var mmyy  = month + year;

  var sheet = getOrCreateSheet('RefNumbers');
  var rows  = sheet.getDataRange().getValues();
  var rowIdx = -1, mthCount = 0, totCount = 0;

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === mmyy) {
      rowIdx = i + 1;
      mthCount = parseInt(rows[i][1]) || 0;
      totCount = parseInt(rows[i][2]) || 0;
      break;
    }
  }
  if (totCount === 0) {
    var maxTotal = 0;
    for (var j = 1; j < rows.length; j++) {
      var t = parseInt(rows[j][2]) || 0;
      if (t > maxTotal) maxTotal = t;
    }
    totCount = maxTotal;
  }

  mthCount++; totCount++;
  if (rowIdx === -1) {
    sheet.appendRow([mmyy, mthCount, totCount, month, year, now()]);
  } else {
    sheet.getRange(rowIdx, 2).setValue(mthCount);
    sheet.getRange(rowIdx, 3).setValue(totCount);
    sheet.getRange(rowIdx, 6).setValue(now());
  }

  var mPad = String(mthCount).padStart(3, '0');
  var tPad = String(totCount).padStart(3, '0');
  var refNo = 'DJCE-' + mmyy + '-M' + mPad + '-T' + tPad + '-' + serviceCode;
  return { ok: true, refNo: docType + '-' + refNo, monthly: mthCount, total: totCount, mmyy: mmyy };
}

function getRefCounters() {
  var sheet = getOrCreateSheet('RefNumbers');
  var rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, counters: [] };
  var headers = rows[0];
  return { ok: true, counters: rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = cellToString(row[i]); });
    return obj;
  })};
}

/* ─── ACTIVITY LOG ────────────────────────────────────────── */
function logActivity(p) {
  var sheet = getOrCreateSheet('ActivityLog');
  sheet.appendRow([now(), p.user||'System', p.message||'', p.type||'general', p.details||'']);
  var maxRows = 500;
  var total = sheet.getLastRow();
  if (total > maxRows + 1) sheet.deleteRows(2, total - maxRows - 1);
  return { ok: true };
}

/* ─── EMAILS ──────────────────────────────────────────────── */

/**
 * Admin notification email (existing — sent to NOTIF_EMAIL)
 */
function sendAdminNotifEmail(id, data, svcSummary, fullAddress) {
  var maps = (data.lat && data.lng)
    ? 'https://www.google.com/maps?q=' + data.lat + ',' + data.lng
    : 'https://maps.google.com/?q=' + encodeURIComponent(fullAddress || data.address || '');

  var slotInfo = data.timeSlot ? (TIME_SLOTS[data.timeSlot] || { label: data.timeSlot, time: data.timeSlot }) : null;
  var scheduleText = data.date ? (data.date + (slotInfo ? ' · ' + slotInfo.label + ' (' + slotInfo.time + ')' : '')) : '—';

  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">',
    '<div style="background:#0d1f3c;color:#fff;padding:20px 24px;border-radius:10px 10px 0 0">',
    '<h2 style="margin:0">🔧 ' + BUSINESS_NAME + '</h2>',
    '<p style="margin:6px 0 0;opacity:.8;font-size:13px">New booking received from website</p></div>',
    '<div style="background:#f8faff;padding:20px 24px"><table style="width:100%;border-collapse:collapse;font-size:14px">',
    row_('Booking ID', '<span style="font-family:monospace;color:#0d1f3c;font-weight:bold">' + id + '</span>'),
    row_('Client', '<strong>' + (data.name||'—') + '</strong>'),
    row_('Phone', data.phone||'—'),
    row_('Email', data.email||'—'),
    row_('Address', fullAddress||'—'),
    row_('Landmark', data.landmark||'—'),
    row_('Service', '<strong>' + (svcSummary||'—') + '</strong>'),
    row_('Schedule', scheduleText),
    row_('Notes', data.notes||'—'),
    '</table>',
    '<div style="margin-top:16px"><a href="' + maps + '" style="background:#0d1f3c;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:bold">📍 View Location</a></div>',
    '</div>',
    '<div style="background:#0d1f3c;color:rgba(255,255,255,.6);padding:12px 24px;border-radius:0 0 10px 10px;font-size:11px;text-align:center">',
    BUSINESS_NAME + ' · ' + BUSINESS_ADDRESS + ' · ' + BUSINESS_PHONE + '</div></div>'
  ].join('');

  MailApp.sendEmail({
    to: NOTIF_EMAIL,
    subject: '📋 New Booking: ' + (data.name||'Client') + ' — ' + (svcSummary||'Service'),
    htmlBody: body
  });
}

/**
 * Client confirmation email — professional invoice-style
 * Like PLDT/Converge confirmation
 */
function sendClientConfirmationEmail(id, data, svcSummary, fullAddress) {
  if (!data.email) return;

  var slotInfo = data.timeSlot ? (TIME_SLOTS[data.timeSlot] || { label: data.timeSlot, time: data.timeSlot }) : null;
  var scheduleDate = data.date || '—';
  var scheduleTime = slotInfo ? (slotInfo.label + ' · ' + slotInfo.time) : (data.time || '—');

  var refNo = data.refNo || id;

  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">',

    // ── Header ──
    '<div style="background:#0d1f3c;padding:28px 30px;text-align:center">',
    '<div style="font-size:28px;font-weight:900;color:#fff;letter-spacing:-0.5px">DJ-<span style="color:#f59e0b">CE</span></div>',
    '<div style="font-size:11px;color:rgba(255,255,255,.7);text-transform:uppercase;letter-spacing:2px;margin-top:4px">Aircon & Electrical Services</div>',
    '<div style="display:inline-block;background:rgba(245,158,11,.15);border:1px solid rgba(245,158,11,.4);border-radius:20px;padding:6px 18px;margin-top:14px">',
    '<span style="color:#f59e0b;font-size:13px;font-weight:700">✅ Booking Received!</span></div>',
    '</div>',

    // ── Booking reference banner ──
    '<div style="background:#f59e0b;padding:14px 30px;text-align:center">',
    '<div style="font-size:11px;color:#0d1f3c;text-transform:uppercase;font-weight:700;letter-spacing:1px">Your Booking Reference</div>',
    '<div style="font-size:20px;font-weight:900;color:#0d1f3c;font-family:monospace;margin-top:4px">' + refNo + '</div>',
    '</div>',

    // ── Greeting ──
    '<div style="padding:24px 30px 0">',
    '<p style="font-size:15px;color:#1a1a2e;font-weight:600;margin:0">Hi ' + (data.name ? data.name.split(' ')[0] : 'there') + '! 👋</p>',
    '<p style="font-size:13.5px;color:#475569;margin:8px 0 20px;line-height:1.6">',
    'We\'ve received your service booking request. Our team will review it and get in touch with you shortly to confirm your appointment. Here\'s a summary of your booking:',
    '</p></div>',

    // ── Booking details card ──
    '<div style="margin:0 30px 20px;background:#f8faff;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden">',
    '<div style="background:#0d1f3c;padding:10px 16px"><span style="color:#f59e0b;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px">📋 Booking Details</span></div>',
    '<table style="width:100%;border-collapse:collapse;font-size:13.5px">',
    rowCard_('Service', '<strong style="color:#0d1f3c">' + (svcSummary||'—') + '</strong>'),
    rowCard_('Date', scheduleDate),
    rowCard_('Time Slot', scheduleTime),
    rowCard_('Status', '<span style="background:#fef3c7;color:#92400e;padding:3px 10px;border-radius:10px;font-size:12px;font-weight:700">⏳ Pending Confirmation</span>'),
    '</table></div>',

    // ── Your info card ──
    '<div style="margin:0 30px 20px;background:#f8faff;border:1px solid #e2e8f0;border-radius:10px;overflow:hidden">',
    '<div style="background:#0d1f3c;padding:10px 16px"><span style="color:#f59e0b;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px">👤 Your Information</span></div>',
    '<table style="width:100%;border-collapse:collapse;font-size:13.5px">',
    rowCard_('Name', data.name||'—'),
    rowCard_('Phone', data.phone||'—'),
    rowCard_('Address', fullAddress||'—'),
    (data.landmark ? rowCard_('Landmark', data.landmark) : ''),
    '</table></div>',

    // ── What happens next ──
    '<div style="margin:0 30px 20px;padding:16px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px">',
    '<div style="font-size:12px;font-weight:700;color:#065f46;text-transform:uppercase;margin-bottom:10px">📌 What Happens Next?</div>',
    '<div style="font-size:13px;color:#374151;line-height:2">',
    '1️⃣ &nbsp;Our team will review your booking<br>',
    '2️⃣ &nbsp;We\'ll call or text you to confirm the schedule<br>',
    '3️⃣ &nbsp;Our technician will arrive at your location on time<br>',
    '4️⃣ &nbsp;Service will be completed and receipt will be provided',
    '</div></div>',

    // ── Contact ──
    '<div style="margin:0 30px 20px;padding:14px 16px;background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;text-align:center">',
    '<div style="font-size:12px;color:#9a3412;font-weight:700;margin-bottom:6px">Questions? Contact Us</div>',
    '<a href="tel:+63' + BUSINESS_PHONE.replace(/\D/g,'').replace(/^0/,'') + '" style="color:#0d1f3c;font-weight:700;font-size:14px;text-decoration:none">📞 ' + BUSINESS_PHONE + '</a>',
    ' &nbsp;·&nbsp; ',
    '<a href="mailto:' + NOTIF_EMAIL + '" style="color:#0d1f3c;font-size:13px;text-decoration:none">' + NOTIF_EMAIL + '</a>',
    '</div>',

    // ── Footer ──
    '<div style="background:#0d1f3c;padding:16px 30px;text-align:center">',
    '<div style="color:rgba(255,255,255,.8);font-size:12px;line-height:1.8">',
    BUSINESS_NAME + '<br>',
    BUSINESS_ADDRESS + '<br>',
    '<span style="color:rgba(255,255,255,.5);font-size:11px">Reliable · Safe · Long-Lasting · We serve with quality and care!</span>',
    '</div></div>',

    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: data.email,
    subject: '✅ Booking Confirmed — ' + refNo + ' | ' + BUSINESS_NAME,
    htmlBody: body,
    replyTo: NOTIF_EMAIL
  });
}

/**
 * Confirmation email sent to client when booking status is changed to "Confirmed"
 */
function sendBookingConfirmedEmail(refNo, name, email, service, date, timeSlot, address, technician) {
  if (!email) return;

  var slotInfo = TIME_SLOTS[timeSlot];
  var scheduleTime = slotInfo ? (slotInfo.label + ' · ' + slotInfo.time) : (timeSlot || '—');
  var firstName = name ? name.split(' ')[0] : 'there';

  var row_ = function(label, val) {
    return '<tr><td style="padding:8px 12px;border:1px solid #e2e8f0;font-weight:600;color:#475569;width:35%;background:#f8fafc;font-size:12px">'
      + label + '</td><td style="padding:8px 12px;border:1px solid #e2e8f0;font-size:12.5px">' + (val||'—') + '</td></tr>';
  };

  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">',

    '<div style="background:#0d1f3c;padding:28px 30px;text-align:center">',
    '<div style="font-size:28px;font-weight:900;color:#fff;letter-spacing:-0.5px">DJ-<span style="color:#f59e0b">CE</span></div>',
    '<div style="font-size:11px;color:rgba(255,255,255,.7);text-transform:uppercase;letter-spacing:2px;margin-top:4px">Aircon & Electrical Services</div>',
    '<div style="display:inline-block;background:rgba(16,185,129,.15);border:1px solid rgba(16,185,129,.4);border-radius:20px;padding:6px 18px;margin-top:14px">',
    '<span style="color:#10b981;font-size:13px;font-weight:700">✅ Booking Confirmed!</span></div>',
    '</div>',

    '<div style="background:#10b981;padding:14px 30px;text-align:center">',
    '<div style="font-size:11px;color:#fff;text-transform:uppercase;font-weight:700;letter-spacing:1px">Booking Reference</div>',
    '<div style="font-size:20px;font-weight:900;color:#fff;font-family:monospace;margin-top:4px">' + refNo + '</div>',
    '</div>',

    '<div style="padding:24px 30px 0">',
    '<p style="font-size:15px;color:#1a1a2e;font-weight:600;margin:0">Hi ' + firstName + '! 👋</p>',
    '<p style="font-size:13.5px;color:#475569;margin:8px 0 20px;line-height:1.6">',
    'Great news! Your booking has been <strong>confirmed</strong>. Our technician will be at your location on the scheduled date. Here are your booking details:</p></div>',

    '<div style="padding:0 30px 20px">',
    '<table style="width:100%;border-collapse:collapse">',
    row_('Reference', '<strong>' + refNo + '</strong>'),
    row_('Service', '<strong>' + (service||'—') + '</strong>'),
    row_('Date', date || '—'),
    row_('Time', scheduleTime),
    row_('Address', address || '—'),
    technician ? row_('Technician', technician) : '',
    '</table></div>',

    '<div style="padding:0 30px 24px">',
    '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:14px 16px">',
    '<div style="font-size:13px;font-weight:700;color:#166534;margin-bottom:4px">📋 Reminders:</div>',
    '<ul style="font-size:12px;color:#166534;margin:0;padding-left:18px;line-height:1.8">',
    '<li>Please make sure someone is at home on the scheduled date.</li>',
    '<li>Our technician will call before arrival.</li>',
    '<li>For changes or cancellation, contact us at <strong>' + BUSINESS_PHONE + '</strong>.</li>',
    '</ul></div></div>',

    '<div style="background:#0d1f3c;padding:16px 30px;text-align:center">',
    '<div style="color:rgba(255,255,255,.8);font-size:12px;line-height:1.8">',
    BUSINESS_NAME + '<br>' + BUSINESS_ADDRESS + '<br>📞 ' + BUSINESS_PHONE + '<br>',
    '<span style="color:rgba(255,255,255,.5);font-size:11px">Reliable · Safe · Long-Lasting</span>',
    '</div></div>',

    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject: '✅ Booking Confirmed — ' + refNo + ' | ' + BUSINESS_NAME,
    htmlBody: body,
    replyTo: NOTIF_EMAIL
  });
}

/* ─── EMPLOYEES ──────────────────────────────────────────── */
function getEmployees() {
  var sheet = getOrCreateSheet('Employees');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, employees: [] };
  var headers = data[0];
  var employees = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headerToKey(headers[j])] = cellToString(data[i][j]);
    }
    employees.push(obj);
  }
  return { ok: true, employees: employees };
}

function saveEmployee(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Employees');
  var rows = sheet.getDataRange().getValues();
  // Update existing
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var r = i + 1;
      sheet.getRange(r, 2).setValue(data.name || '');
      sheet.getRange(r, 3).setValue(data.role || '');
      sheet.getRange(r, 4).setValue(data.type || '');
      sheet.getRange(r, 5).setValue(data.dailyRate || '');
      sheet.getRange(r, 6).setValue(data.status || 'Active');
      sheet.getRange(r, 7).setValue(data.phone || '');
      sheet.getRange(r, 8).setValue(data.startDate || '');
      sheet.getRange(r, 9).setValue(data.notes || '');
      return { ok: true, action: 'updated', id: data.id };
    }
  }
  // Create new
  var id = data.id || ('EMP-' + new Date().getTime());
  sheet.appendRow([
    id, data.name||'', data.role||'', data.type||'', data.dailyRate||'',
    data.status||'Active', data.phone||'', data.startDate||'', data.notes||'', now()
  ]);
  return { ok: true, action: 'created', id: id };
}

function deleteEmployee(p) {
  var id = p.id;
  if (!id) return { ok: false, error: 'Missing employee ID' };
  var sheet = getOrCreateSheet('Employees');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true, action: 'deleted', id: id };
    }
  }
  return { ok: false, error: 'Employee not found: ' + id };
}

/* ─── DAILY TRACKER ──────────────────────────────────────── */
function getDailyTracker(p) {
  var sheet = getOrCreateSheet('DailyTracker');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, records: [] };
  var headers = data[0];
  var records = [];
  // Optional filter by month (ym = '2026-10')
  var ym = p.ym || '';
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headerToKey(headers[j])] = cellToString(data[i][j]);
    }
    if (ym && obj.date && obj.date.substring(0,7) !== ym) continue;
    records.push(obj);
  }
  return { ok: true, records: records };
}

function saveDailyRecord(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('DailyTracker');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      return { ok: true, action: 'already_exists', id: data.id };
    }
  }
  var id = data.id || ('DTR-' + new Date().getTime());
  sheet.appendRow([
    id, data.date||'', data.employeeId||'', data.employeeName||'',
    data.bookingIds||'', data.jobsCount||0, data.classification||'',
    data.auditedBy||'', data.auditNotes||'', now()
  ]);
  return { ok: true, action: 'created', id: id };
}

function updateDailyRecord(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var id = data.id || p.id;
  if (!id) return { ok: false, error: 'Missing record ID' };
  var sheet = getOrCreateSheet('DailyTracker');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      var r = i + 1;
      if (data.classification !== undefined) sheet.getRange(r, 7).setValue(data.classification);
      if (data.auditedBy !== undefined) sheet.getRange(r, 8).setValue(data.auditedBy);
      if (data.auditNotes !== undefined) sheet.getRange(r, 9).setValue(data.auditNotes);
      return { ok: true, action: 'updated', id: id };
    }
  }
  return { ok: false, error: 'Record not found: ' + id };
}

/* ─── PAYROLL ────────────────────────────────────────────── */
function getPayroll(p) {
  var sheet = getOrCreateSheet('Payroll');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, records: [] };
  var headers = data[0];
  var records = [];
  var ym = p.ym || '';
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headerToKey(headers[j])] = cellToString(data[i][j]);
    }
    if (ym && obj.periodFrom && obj.periodFrom.substring(0,7) !== ym) continue;
    records.push(obj);
  }
  return { ok: true, records: records };
}

function savePayroll(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Payroll');
  var rows = sheet.getDataRange().getValues();
  // Update if exists
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var r = i + 1;
      sheet.getRange(r, 6).setValue(data.daysWorked || 0);
      sheet.getRange(r, 7).setValue(data.halfDays || 0);
      sheet.getRange(r, 8).setValue(data.fullDays || 0);
      sheet.getRange(r, 9).setValue(data.otDays || 0);
      sheet.getRange(r, 10).setValue(data.dailyRate || 0);
      sheet.getRange(r, 11).setValue(data.grossPay || 0);
      sheet.getRange(r, 12).setValue(data.deductions || 0);
      sheet.getRange(r, 13).setValue(data.netPay || 0);
      sheet.getRange(r, 14).setValue(data.thirteenthMonthAccrual || 0);
      sheet.getRange(r, 15).setValue(data.status || 'Draft');
      sheet.getRange(r, 17).setValue(data.approvedBy || '');
      sheet.getRange(r, 18).setValue(data.notes || '');
      return { ok: true, action: 'updated', id: data.id };
    }
  }
  // Create new
  var id = data.id || ('PAY-' + new Date().getTime());
  sheet.appendRow([
    id, data.periodFrom||'', data.periodTo||'', data.employeeId||'', data.employeeName||'',
    data.daysWorked||0, data.halfDays||0, data.fullDays||0, data.otDays||0,
    data.dailyRate||0, data.grossPay||0, data.deductions||0, data.netPay||0,
    data.thirteenthMonthAccrual||0, data.status||'Draft', now(), data.approvedBy||'', data.notes||''
  ]);
  return { ok: true, action: 'created', id: id };
}

/* ─── INVESTOR SHARES ────────────────────────────────────── */
function getInvestors() {
  var sheet = getOrCreateSheet('InvestorShares');
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return { ok: true, investors: [] };
  var headers = data[0];
  var investors = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headerToKey(headers[j])] = cellToString(data[i][j]);
    }
    investors.push(obj);
  }
  return { ok: true, investors: investors };
}

function saveInvestor(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('InvestorShares');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var r = i + 1;
      sheet.getRange(r, 2).setValue(data.investorName || '');
      sheet.getRange(r, 3).setValue(data.sharePercent || 0);
      sheet.getRange(r, 4).setValue(data.year || '');
      sheet.getRange(r, 5).setValue(data.month || '');
      sheet.getRange(r, 6).setValue(data.monthlyIpon || 0);
      sheet.getRange(r, 7).setValue(data.notes || '');
      return { ok: true, action: 'updated', id: data.id };
    }
  }
  var id = data.id || ('INV-' + new Date().getTime());
  sheet.appendRow([
    id, data.investorName||'', data.sharePercent||0, data.year||'',
    data.month||'', data.monthlyIpon||0, data.notes||'', now()
  ]);
  return { ok: true, action: 'created', id: id };
}

function deleteInvestor(p) {
  var id = p.id;
  if (!id) return { ok: false, error: 'Missing investor record ID' };
  var sheet = getOrCreateSheet('InvestorShares');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true, action: 'deleted', id: id };
    }
  }
  return { ok: false, error: 'Investor record not found: ' + id };
}

/* ─── EMAIL HELPERS ───────────────────────────────────────── */
function row_(label, value) {
  return '<tr><td style="padding:7px 0;color:#64748b;width:120px;vertical-align:top">' + label + '</td>' +
         '<td style="padding:7px 0;color:#1e293b">' + value + '</td></tr>';
}
function rowCard_(label, value) {
  return '<tr style="border-bottom:1px solid #e2e8f0">' +
         '<td style="padding:10px 16px;color:#64748b;font-size:12px;text-transform:uppercase;letter-spacing:0.3px;white-space:nowrap;width:110px">' + label + '</td>' +
         '<td style="padding:10px 16px;color:#1e293b;font-weight:500">' + value + '</td></tr>';
}

/* ─── HELPERS ─────────────────────────────────────────────── */
function cellToString(val) {
  if (val === null || val === undefined || val === '') return '';
  if (val instanceof Date) {
    // If time is midnight, return date-only (yyyy-MM-dd) — prevents "Invalid Date" in frontends
    var h = val.getHours(), m = val.getMinutes(), s = val.getSeconds();
    if (h === 0 && m === 0 && s === 0) return Utilities.formatDate(val, TIMEZONE, 'yyyy-MM-dd');
    return Utilities.formatDate(val, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  }
  return String(val);
}

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  var headers = HEADERS[name] || [];
  if (headers.length && sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    var style = SHEET_STYLES[name] || { bg: '#1e3a5f', fg: '#ffffff' };
    var hr = sheet.getRange(1, 1, 1, headers.length);
    hr.setBackground(style.bg).setFontColor(style.fg).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function setCellByKey(sheet, rowNum, colMap, key, value) {
  var col = colMap[key];
  if (col) sheet.getRange(rowNum, col).setValue(value);
}

/* ─── SHEET HEADER DIAGNOSTICS & FIX ─────────────────────── */
function getSheetHeaders(p) {
  var name = p && p.sheet ? p.sheet : 'Bookings';
  var sheet = getOrCreateSheet(name);
  if (sheet.getLastRow() === 0) return { ok: true, headers: [], note: 'Sheet is empty' };
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var expected = HEADERS[name] || [];
  var match = true;
  if (headers.length !== expected.length) {
    match = false;
  } else {
    for (var i = 0; i < expected.length; i++) {
      if (String(headers[i]).trim() !== String(expected[i]).trim()) { match = false; break; }
    }
  }
  return { ok: true, sheet: name, actual: headers, expected: expected, match: match };
}

function fixBookingHeaders() {
  var sheet = getOrCreateSheet('Bookings');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS.Bookings);
    return { ok: true, note: 'Empty sheet — headers written', headers: HEADERS.Bookings };
  }
  var lastCol = sheet.getLastColumn();
  var actualHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  var expected = HEADERS.Bookings;

  // Check if already matching
  var match = true;
  if (actualHeaders.length === expected.length) {
    for (var i = 0; i < expected.length; i++) {
      if (actualHeaders[i] !== expected[i]) { match = false; break; }
    }
  } else {
    match = false;
  }
  if (match) return { ok: true, note: 'Headers already correct', headers: expected };

  // Build mapping: for each expected header, find its actual column index
  var colMap = {}; // expectedIndex → actualIndex
  var unmapped = [];
  for (var ei = 0; ei < expected.length; ei++) {
    var found = -1;
    for (var ai = 0; ai < actualHeaders.length; ai++) {
      if (actualHeaders[ai] === expected[ei]) { found = ai; break; }
    }
    if (found >= 0) {
      colMap[ei] = found;
    } else {
      unmapped.push(expected[ei]);
    }
  }

  // If too many unmapped, don't risk reorder — just fix headers row
  if (unmapped.length > 10) {
    // Too many missing headers — just overwrite row 1 and warn
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    return { ok: true, note: 'Headers overwritten (many missing — data may still be misaligned)', fixed: expected, unmapped: unmapped };
  }

  // Reorder all data rows to match expected header order
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var allData = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var reordered = allData.map(function(row) {
      var newRow = [];
      for (var ei = 0; ei < expected.length; ei++) {
        if (colMap[ei] !== undefined && colMap[ei] < row.length) {
          newRow.push(row[colMap[ei]]);
        } else {
          newRow.push(''); // unmapped column gets empty
        }
      }
      return newRow;
    });

    // Clear old data area and write reordered
    sheet.getRange(1, 1, lastRow, lastCol).clear();
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
    if (reordered.length > 0) {
      sheet.getRange(2, 1, reordered.length, expected.length).setValues(reordered);
    }
    // Re-style header row
    var style = SHEET_STYLES['Bookings'] || { bg: '#1e3a5f', fg: '#ffffff' };
    var hr = sheet.getRange(1, 1, 1, expected.length);
    hr.setBackground(style.bg).setFontColor(style.fg).setFontWeight('bold');
    sheet.setFrozenRows(1);
  } else {
    sheet.getRange(1, 1, 1, expected.length).setValues([expected]);
  }

  return { ok: true, note: 'Headers fixed and data reordered', fixed: expected, unmapped: unmapped, rowsReordered: lastRow - 1 };
}

function headerToKey(header) {
  return header.toLowerCase()
    .replace(/[^a-z0-9 ]/g, '')
    .replace(/\s+(.)/g, function(_, c) { return c.toUpperCase(); })
    .replace(/\s/g, '');
}

function now() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
/* ─── POST HANDLER (Photo Upload) ────────────────────────── */
function doPost(e) {
  var result;
  try {
    var params = e.parameter || {};
    var action = params.action || '';

    if (action === 'uploadPhoto') {
      var postData = e.postData ? e.postData.contents : '';
      var body = {};
      // Parse form-encoded POST body
      if (postData) {
        postData.split('&').forEach(function(pair) {
          var kv = pair.split('=');
          body[decodeURIComponent(kv[0])] = decodeURIComponent(kv[1] || '');
        });
      }
      // Also check e.parameter for form fields
      if (!body.base64 && params.base64) body.base64 = params.base64;
      if (!body.bookingId) body.bookingId = params.bookingId || '';
      if (!body.photoType) body.photoType = params.photoType || '';
      if (!body.fileName) body.fileName = params.fileName || '';
      if (!body.addedBy) body.addedBy = params.addedBy || '';

      result = uploadPhotoToDrive(body);
    } else {
      result = { ok: false, error: 'Unknown POST action: ' + action };
    }
  } catch (err) {
    result = { ok: false, error: err.message || String(err) };
  }

  // Return as HTML with postMessage for iframe communication
  var json = JSON.stringify(result);
  var html = '<html><body><script>'
    + 'try { window.parent.postMessage(' + json + ', "*"); }'
    + 'catch(e) { console.error("postMessage failed", e); }'
    + '</script></body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function uploadPhotoToDrive(body) {
  var base64 = body.base64 || '';
  var bookingId = body.bookingId || 'UNKNOWN';
  var photoType = body.photoType || 'photo';
  var fileName  = body.fileName || 'photo.jpg';
  var addedBy   = body.addedBy || '';

  if (!base64) return { ok: false, error: 'No image data provided' };

  // Strip data URI prefix if present
  var cleanBase64 = base64.replace(/^data:image\/\w+;base64,/, '');

  // Decode base64 to blob
  var decoded = Utilities.base64Decode(cleanBase64);
  var blob = Utilities.newBlob(decoded, 'image/jpeg', fileName);

  // Get or create the photo folder
  var parentFolder = getOrCreatePhotoFolder();

  // Get or create booking subfolder
  var subFolders = parentFolder.getFoldersByName(bookingId);
  var bookingFolder = subFolders.hasNext() ? subFolders.next() : parentFolder.createFolder(bookingId);

  // Save file with descriptive name
  var ts = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd-HHmmss');
  var safeName = photoType + '_' + ts + '_' + fileName.replace(/[^a-zA-Z0-9._-]/g, '');
  var file = bookingFolder.createFile(blob.setName(safeName));

  // Make publicly viewable
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var fileId = file.getId();
  var viewUrl = 'https://drive.google.com/uc?id=' + fileId;
  var thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w400';

  return {
    ok: true,
    action: 'uploaded',
    fileId: fileId,
    url: viewUrl,
    thumb: thumbUrl,
    bookingId: bookingId,
    photoType: photoType,
    fileName: safeName
  };
}

/* ─── SERVICE DUE REMINDERS (Time-Driven Trigger) ────────── */
/**
 * Run daily via Apps Script trigger (Edit → Triggers → Add → checkServiceReminders → Day timer → 8-9 AM).
 * Scans bookings with nextCleaningDate ≤ 14 days from now, sends email reminders to clients,
 * and sends a summary digest to admin.
 */
function checkServiceReminders() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Bookings');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return;
  var headers = data[0];
  var colMap = {};
  for (var j = 0; j < headers.length; j++) colMap[headerToKey(headers[j])] = j;

  var today = new Date();
  today.setHours(0,0,0,0);
  var in14 = new Date(today.getTime() + 14 * 86400000);
  var reminded = [];
  var overdue = [];

  for (var i = 1; i < data.length; i++) {
    var status = cellToString(data[i][colMap['status']]);
    if (status === 'Cancelled' || status === 'Deleted') continue;
    var ncd = cellToString(data[i][colMap['nextCleaningDate']]);
    if (!ncd) continue;
    var ncdDate = new Date(ncd + 'T00:00:00');
    if (isNaN(ncdDate.getTime())) continue;
    ncdDate.setHours(0,0,0,0);
    var diffDays = Math.round((ncdDate - today) / 86400000);
    var followedUp = cellToString(data[i][colMap['followedUp']]);
    var name = cellToString(data[i][colMap['name']]);
    var email = cellToString(data[i][colMap['email']]);
    var phone = cellToString(data[i][colMap['phone']]);
    var service = cellToString(data[i][colMap['service']]);
    var bookingId = cellToString(data[i][colMap['bookingId']]);
    var refNo = cellToString(data[i][colMap['refNo']]) || bookingId;

    // Overdue: past nextCleaningDate
    if (diffDays < 0) {
      overdue.push({ name: name, phone: phone, service: service, ncd: ncd, refNo: refNo, days: Math.abs(diffDays) });
    }

    // Send reminder: within 14 days AND not already followed up with email
    if (diffDays >= 0 && diffDays <= 14 && email && followedUp !== 'email-sent') {
      try {
        sendServiceReminderEmail(email, name, service, ncd, refNo, diffDays);
        // Mark as followed up
        sheet.getRange(i + 1, colMap['followedUp'] + 1).setValue('email-sent');
        reminded.push({ name: name, service: service, ncd: ncd, refNo: refNo, days: diffDays });
      } catch(e) {}
    }
  }

  // Send admin digest
  if (reminded.length || overdue.length) {
    sendAdminReminderDigest(reminded, overdue);
  }
}

/**
 * Reminder email to client — "Your aircon service is due soon!"
 */
function sendServiceReminderEmail(email, name, service, ncd, refNo, daysUntil) {
  var dateStr = Utilities.formatDate(new Date(ncd), TIMEZONE, 'MMMM d, yyyy');
  var urgency = daysUntil <= 3 ? '🔴' : daysUntil <= 7 ? '🟡' : '🟢';
  var dueText = daysUntil === 0 ? 'today!' : 'in ' + daysUntil + ' day' + (daysUntil>1?'s':'') + '!';

  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">',
    '<div style="background:#0d1f3c;padding:24px 30px;text-align:center">',
    '<div style="font-size:26px;font-weight:900;color:#fff">DJ-<span style="color:#f59e0b">CE</span></div>',
    '<div style="font-size:11px;color:rgba(255,255,255,.7);text-transform:uppercase;letter-spacing:2px;margin-top:4px">Aircon & Electrical Services</div></div>',

    '<div style="background:#fef3c7;padding:16px 30px;text-align:center">',
    '<div style="font-size:24px">' + urgency + '</div>',
    '<div style="font-size:14px;font-weight:700;color:#92400e">Your next service is due ' + dueText + '</div></div>',

    '<div style="padding:24px 30px">',
    '<p style="font-size:15px;color:#1a1a2e;margin:0">Hi <strong>' + (name ? name.split(' ')[0] : 'there') + '</strong>! 👋</p>',
    '<p style="font-size:13.5px;color:#475569;margin:10px 0 20px;line-height:1.6">',
    'This is a friendly reminder that your <strong>' + (service||'aircon service') + '</strong> maintenance is coming up. Regular cleaning keeps your unit running efficiently and extends its lifespan!</p>',

    '<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;padding:16px;margin-bottom:20px">',
    '<table style="width:100%;font-size:13.5px;border-collapse:collapse">',
    '<tr><td style="padding:6px 0;color:#6b7280;width:120px">Reference</td><td style="font-weight:700;color:#0d1f3c;font-family:monospace">' + refNo + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#6b7280">Service</td><td style="font-weight:600">' + (service||'—') + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#6b7280">Recommended Date</td><td style="font-weight:700;color:#065f46">' + dateStr + '</td></tr>',
    '</table></div>',

    '<div style="text-align:center;margin:20px 0">',
    '<a href="tel:+63' + BUSINESS_PHONE.replace(/\D/g,'').replace(/^0/,'') + '" style="display:inline-block;background:#0d1f3c;color:#fff;padding:12px 28px;border-radius:8px;text-decoration:none;font-weight:700;font-size:14px">📞 Book Now — ' + BUSINESS_PHONE + '</a>',
    '</div>',
    '<p style="font-size:12px;color:#94a3b8;text-align:center;margin-top:10px">You can also reply to this email or message us on Facebook!</p>',
    '</div>',

    '<div style="background:#0d1f3c;padding:14px 30px;text-align:center">',
    '<div style="color:rgba(255,255,255,.6);font-size:11px">' + BUSINESS_NAME + ' · ' + BUSINESS_ADDRESS + '</div></div>',
    '</div>'
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject: urgency + ' Service Reminder — ' + (service||'Maintenance') + ' due ' + dueText + ' | ' + BUSINESS_NAME,
    htmlBody: body,
    replyTo: NOTIF_EMAIL
  });
}

/**
 * Admin digest — summary of reminders sent + overdue bookings
 */
function sendAdminReminderDigest(reminded, overdue) {
  var rows = '';
  if (reminded.length) {
    rows += '<div style="margin-bottom:20px"><h3 style="color:#065f46;font-size:14px;margin:0 0 10px">✅ Reminders Sent Today (' + reminded.length + ')</h3>';
    rows += '<table style="width:100%;border-collapse:collapse;font-size:13px">';
    rows += '<tr style="background:#f0fdf4"><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Client</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Service</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Due Date</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Days Left</th></tr>';
    reminded.forEach(function(r) {
      rows += '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee">' + r.name + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + r.service + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + r.ncd + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + r.days + ' days</td></tr>';
    });
    rows += '</table></div>';
  }
  if (overdue.length) {
    rows += '<div><h3 style="color:#dc2626;font-size:14px;margin:0 0 10px">🔴 Overdue — Needs Follow-up (' + overdue.length + ')</h3>';
    rows += '<table style="width:100%;border-collapse:collapse;font-size:13px">';
    rows += '<tr style="background:#fef2f2"><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Client</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Phone</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Service</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Due Date</th><th style="padding:8px;text-align:left;border-bottom:1px solid #ddd">Overdue</th></tr>';
    overdue.forEach(function(o) {
      rows += '<tr><td style="padding:6px 8px;border-bottom:1px solid #eee">' + o.name + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + o.phone + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + o.service + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee">' + o.ncd + '</td><td style="padding:6px 8px;border-bottom:1px solid #eee;color:#dc2626;font-weight:700">' + o.days + ' days</td></tr>';
    });
    rows += '</table></div>';
  }

  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:650px;margin:0 auto">',
    '<div style="background:#0d1f3c;color:#fff;padding:18px 24px;border-radius:10px 10px 0 0">',
    '<h2 style="margin:0;font-size:16px">📋 DJ-CE Daily Service Reminder Digest</h2>',
    '<p style="margin:4px 0 0;opacity:.7;font-size:12px">' + Utilities.formatDate(new Date(), TIMEZONE, 'EEEE, MMMM d, yyyy') + '</p></div>',
    '<div style="padding:20px 24px;background:#fff;border:1px solid #e5e7eb;border-top:none;border-radius:0 0 10px 10px">',
    rows,
    '</div></div>'
  ].join('');

  MailApp.sendEmail({
    to: NOTIF_EMAIL,
    subject: '📋 Service Reminders: ' + reminded.length + ' sent, ' + overdue.length + ' overdue | ' + Utilities.formatDate(new Date(), TIMEZONE, 'MMM d'),
    htmlBody: body
  });
}

/**
 * Manual reminder — called from portal via JSONP: action=sendReminder&id=BK-xxx
 */
/* ─── RESCHEDULE BOOKING ──────────────────────────────────── */
function rescheduleBooking(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Bookings');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var colMap = {};
  headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });

  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var email   = cellToString(rows[i][colMap['email'] - 1]);
      var name    = cellToString(rows[i][colMap['name'] - 1]);
      var service = cellToString(rows[i][colMap['service'] - 1]);
      var refNo   = cellToString(rows[i][colMap['refNo'] - 1]) || data.id;
      var phone   = cellToString(rows[i][colMap['phone'] - 1]);

      // Update date and time slot on sheet
      if (colMap['date'])     sheet.getRange(i + 1, colMap['date']).setValue(data.newDate);
      if (colMap['timeSlot']) sheet.getRange(i + 1, colMap['timeSlot']).setValue(data.newTime);

      // Send reschedule email to client
      var emailSent = false;
      if (email) {
        try {
          sendRescheduleEmail(email, name, service, refNo, data.oldDate, data.oldTime, data.newDate, data.newTime, data.reason || '');
          emailSent = true;
        } catch(e) { Logger.log('Reschedule email failed: ' + e); }
      }

      // Send notification to admin
      try {
        MailApp.sendEmail({
          to: NOTIF_EMAIL,
          subject: '📅 Rescheduled: ' + (name || 'Client') + ' — ' + refNo,
          htmlBody: '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:20px">'
            + '<h2 style="color:#f59e0b">📅 Booking Rescheduled</h2>'
            + '<p><strong>' + refNo + '</strong> — ' + (name || 'Client') + '</p>'
            + '<p>Service: ' + (service || '—') + '</p>'
            + '<p>Old: ' + (data.oldDate || '—') + ' · ' + (data.oldTime || '—') + '</p>'
            + '<p>New: <strong>' + (data.newDate || '—') + ' · ' + (data.newTime || '—') + '</strong></p>'
            + (data.reason ? '<p>Reason: ' + data.reason + '</p>' : '')
            + (emailSent ? '<p style="color:green">✅ Client notified via email</p>' : '<p style="color:#999">⚠️ No client email — not notified</p>')
            + '</div>'
        });
      } catch(e) { Logger.log('Admin reschedule notif failed: ' + e); }

      return { ok: true, id: data.id, emailSent: emailSent };
    }
  }
  return { ok: false, error: 'Booking not found: ' + data.id };
}

function sendRescheduleEmail(email, name, service, refNo, oldDate, oldTime, newDate, newTime, reason) {
  var firstName = name ? name.split(' ')[0] : 'there';

  var newDateFormatted = newDate;
  try {
    var d = new Date(newDate + 'T00:00:00');
    newDateFormatted = d.toLocaleDateString('en-PH', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
  } catch(e) {}

  var oldDateFormatted = oldDate;
  try {
    var d2 = new Date(oldDate + 'T00:00:00');
    oldDateFormatted = d2.toLocaleDateString('en-PH', { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' });
  } catch(e) {}

  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden">',

    // Header
    '<div style="background:#0d1f3c;padding:28px 30px;text-align:center">',
    '<div style="font-size:28px;font-weight:900;color:#fff;letter-spacing:-0.5px">DJ-<span style="color:#f59e0b">CE</span></div>',
    '<div style="font-size:11px;color:rgba(255,255,255,.7);text-transform:uppercase;letter-spacing:2px;margin-top:4px">Aircon & Electrical Services</div>',
    '<div style="display:inline-block;background:rgba(245,158,11,.15);border:1px solid rgba(245,158,11,.4);border-radius:20px;padding:6px 18px;margin-top:14px">',
    '<span style="color:#f59e0b;font-size:13px;font-weight:700">📅 Schedule Update</span></div>',
    '</div>',

    // Ref banner
    '<div style="background:#f59e0b;padding:14px 30px;text-align:center">',
    '<div style="font-size:11px;color:#0d1f3c;text-transform:uppercase;font-weight:700;letter-spacing:1px">Booking Reference</div>',
    '<div style="font-size:20px;font-weight:900;color:#0d1f3c;font-family:monospace;margin-top:4px">' + refNo + '</div>',
    '</div>',

    // Greeting
    '<div style="padding:24px 30px 0">',
    '<p style="font-size:15px;color:#1a1a2e;font-weight:600;margin:0">Hi ' + firstName + '! 👋</p>',
    '<p style="font-size:13.5px;color:#475569;margin:8px 0 20px;line-height:1.6">',
    'We\'re writing to let you know that your service appointment has been rescheduled. Please see the updated details below:',
    '</p></div>',

    // Old schedule (crossed out)
    '<div style="margin:0 30px 12px;background:#fef2f2;border:1px solid #fecaca;border-radius:10px;overflow:hidden">',
    '<div style="background:#991b1b;padding:8px 16px"><span style="color:#fff;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px">❌ Previous Schedule</span></div>',
    '<table style="width:100%;border-collapse:collapse;font-size:13.5px">',
    rowCard_('Date', '<span style="text-decoration:line-through;color:#991b1b">' + oldDateFormatted + '</span>'),
    rowCard_('Time', '<span style="text-decoration:line-through;color:#991b1b">' + (oldTime || '—') + '</span>'),
    '</table></div>',

    // New schedule (highlighted)
    '<div style="margin:0 30px 20px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:10px;overflow:hidden">',
    '<div style="background:#065f46;padding:8px 16px"><span style="color:#fff;font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px">✅ New Schedule</span></div>',
    '<table style="width:100%;border-collapse:collapse;font-size:13.5px">',
    rowCard_('Date', '<strong style="color:#065f46">' + newDateFormatted + '</strong>'),
    rowCard_('Time', '<strong style="color:#065f46">' + (newTime || '—') + '</strong>'),
    rowCard_('Service', service || '—'),
    '</table></div>',

    // Reason (if provided)
    (reason ? '<div style="margin:0 30px 20px;padding:14px 16px;background:#fffbeb;border:1px solid #fde68a;border-radius:10px;font-size:13px;color:#92400e"><strong>📝 Reason:</strong> ' + reason + '</div>' : ''),

    // Contact info
    '<div style="margin:0 30px 20px;padding:16px;background:#f8faff;border:1px solid #e2e8f0;border-radius:10px">',
    '<div style="font-size:12px;font-weight:700;color:#0d1f3c;margin-bottom:8px">📞 Questions?</div>',
    '<p style="font-size:13px;color:#475569;margin:0;line-height:1.6">',
    'If you need to make changes or have any concerns, please don\'t hesitate to contact us at <strong>' + BUSINESS_PHONE + '</strong>.',
    '</p></div>',

    // Footer
    '<div style="background:#0d1f3c;color:rgba(255,255,255,.6);padding:12px 24px;border-radius:0 0 10px 10px;font-size:11px;text-align:center">',
    BUSINESS_NAME + ' · ' + BUSINESS_ADDRESS + ' · ' + BUSINESS_PHONE + '</div></div>'
  ].join('');

  MailApp.sendEmail({
    to: email,
    subject: '📅 Schedule Update — ' + BUSINESS_NAME + ' [' + refNo + ']',
    htmlBody: body
  });
}

function sendManualReminder(p) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Bookings');
  if (!sheet) return { ok: false, error: 'No Bookings sheet' };
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colMap = {};
  for (var j = 0; j < headers.length; j++) colMap[headerToKey(headers[j])] = j;
  var idCol = colMap['bookingId'];
  for (var i = 1; i < data.length; i++) {
    if (cellToString(data[i][idCol]) === p.id) {
      var email = cellToString(data[i][colMap['email']]);
      if (!email) return { ok: false, error: 'No email for this client' };
      var name = cellToString(data[i][colMap['name']]);
      var service = cellToString(data[i][colMap['service']]);
      var ncd = cellToString(data[i][colMap['nextCleaningDate']]);
      var refNo = cellToString(data[i][colMap['refNo']]) || p.id;
      var daysUntil = 0;
      if (ncd) {
        var ncdD = new Date(ncd + 'T00:00:00');
        var now = new Date(); now.setHours(0,0,0,0);
        daysUntil = Math.max(0, Math.round((ncdD - now) / 86400000));
      }
      sendServiceReminderEmail(email, name, service, ncd || '', refNo, daysUntil);
      sheet.getRange(i + 1, colMap['followedUp'] + 1).setValue('email-sent');
      return { ok: true, action: 'reminderSent', name: name };
    }
  }
  return { ok: false, error: 'Booking not found' };
}

/* ─── DATA BACKUP (Phase 4 #27a) ─────────────────────────── */
/**
 * Creates a timestamped copy of the entire spreadsheet in a "DJ-CE Backups" folder.
 * Set up as a weekly trigger: Triggers > Add Trigger > backupSpreadsheet > Time-driven > Week timer
 */
function backupSpreadsheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var name = BUSINESS_NAME + ' Backup — ' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm');

  // Find or create backup folder
  var folders = DriveApp.getFoldersByName('DJ-CE Backups');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('DJ-CE Backups');

  // Copy spreadsheet to backup folder
  var file = DriveApp.getFileById(SHEET_ID);
  var copy = file.makeCopy(name, folder);

  // Clean up old backups — keep last 8
  var backups = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    backups.push({ file: f, date: f.getDateCreated() });
  }
  backups.sort(function(a, b) { return b.date - a.date; });
  if (backups.length > 8) {
    for (var i = 8; i < backups.length; i++) {
      backups[i].file.setTrashed(true);
    }
  }

  // Send confirmation email
  try {
    MailApp.sendEmail({
      to: NOTIF_EMAIL,
      subject: '✅ DJ-CE Data Backup Complete — ' + Utilities.formatDate(new Date(), TIMEZONE, 'MMM dd, yyyy'),
      htmlBody: '<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;padding:20px">'
        + '<h2 style="color:#b8860b">✅ Backup Complete</h2>'
        + '<p>A backup of your DJ-CE database was created:</p>'
        + '<p><strong>' + name + '</strong></p>'
        + '<p>📁 Location: DJ-CE Backups folder in Google Drive</p>'
        + '<p>📊 ' + backups.length + ' total backup(s) stored (max 8 kept)</p>'
        + '<p style="color:#888;font-size:12px">This is an automated backup notification.</p>'
        + '</div>'
    });
  } catch(e) { Logger.log('Backup email failed: ' + e); }

  return { ok: true, action: 'backup', file: copy.getName(), id: copy.getId() };
}