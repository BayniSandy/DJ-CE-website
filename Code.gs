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
    'Last Pw Change','Force Change','Created At','Last Login'
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
  ScheduleSlots: { bg: '#1e3a5f', fg: '#f59e0b' }
};

/* ─── TIME SLOTS ──────────────────────────────────────────── */
var TIME_SLOTS = {
  'morning':   { label: 'Morning',   time: '9:00 AM – 12:00 PM' },
  'afternoon': { label: 'Afternoon', time: '1:00 PM – 4:00 PM'  },
  'evening':   { label: 'Evening',   time: '5:00 PM – 7:00 PM'  }
};

/* ─── ENTRY POINT ─────────────────────────────────────────── */
function doGet(e) {
  var result;
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || '';
    var cb     = params.cb     || '';

    if (!action) {
      result = { ok: true, name: 'DJ-CE Backend v4.0', time: now() };
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
        case 'getComplaints':       result = getComplaints();             break;
        case 'saveComplaint':       result = saveComplaint(params);       break;
        case 'logActivity':         result = logActivity(params);         break;
        case 'saveExpense':         result = saveExpense(params);         break;
        case 'getExpenses':         result = getExpenses();               break;
        case 'getNextRefNumber':    result = getNextRefNumber(params);    break;
        case 'getRefCounters':      result = getRefCounters();            break;
        case 'getAvailableSlots':   result = getAvailableSlots(params);   break;
        case 'openSlot':            result = openSlot(params);            break;
        case 'closeSlot':           result = closeSlot(params);           break;
        case 'deleteSlot':          result = deleteSlot(params);          break;
        case 'getAllSlots':          result = getAllSlots(params);         break;
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
      return { ok: true, action: 'updated', username: data.username };
    }
  }
  sheet.appendRow([
    data.username, data.password||'', data.name||'', data.role||'',
    data.avatar||'?', data.lastPwChange||'', data.forceChange||'false',
    now(), ''
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
  if (val instanceof Date) return Utilities.formatDate(val, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
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

function headerToKey(header) {
  return header.toLowerCase()
    .replace(/[^a-z0-9 ]/g, '')
    .replace(/\s+(.)/g, function(_, c) { return c.toUpperCase(); })
    .replace(/\s/g, '');
}

function now() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
