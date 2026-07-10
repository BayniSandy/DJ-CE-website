/**
 * ============================================================
 * DJ-CE AIRCON & ELECTRICAL SERVICES
 * Google Apps Script Backend — v3.0 (API Only)
 * ============================================================
 *
 * SETUP:
 * 1. Google Sheet → Extensions → Apps Script
 * 2. Delete all existing code in Code.gs
 * 3. Paste this entire script → Save (Ctrl+S)
 * 4. Deploy → New Deployment → Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Authorize → Copy the Web App URL
 * 6. Paste the URL into both HTML files (APPS_SCRIPT_URL)
 *
 * SHEETS CREATED AUTOMATICALLY:
 * Bookings, Accounts, Clients, Complaints, ActivityLog
 * ============================================================
 */

/* ─── CONFIG ──────────────────────────────────────────────── */
var SHEET_ID      = '1qdPMEx4_lzyW9WH8ADiixzeM15ILGBZiijJGfhaAf34';
var NOTIF_EMAIL   = 'djceservicesforyou@gmail.com';
var BUSINESS_NAME = 'DJ-CE Aircon & Electrical Services';
var TIMEZONE      = 'Asia/Manila';

/* ─── SHEET HEADERS ───────────────────────────────────────── */
var HEADERS = {
  Bookings: [
    'Booking ID','Created At','Name','Phone','Email',
    'Address','Landmark','Lat','Lng',
    'Service','Service Details','Date','Time',
    'Technician','Status','Amount','Usage Frequency',
    'Next Cleaning Date','Notes','Source','Followed Up','Completed At'
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
  ]
};

var SHEET_STYLES = {
  Bookings:    { bg: '#0d1f3c', fg: '#ffffff' },
  Accounts:    { bg: '#0d1f3c', fg: '#f59e0b' },
  Clients:     { bg: '#065f46', fg: '#ffffff' },
  Complaints:  { bg: '#7f1d1d', fg: '#ffffff' },
  ActivityLog: { bg: '#1e3a5f', fg: '#ffffff' }
};

/* ─── ENTRY POINT ─────────────────────────────────────────── */

function doGet(e) {
  var result;
  try {
    var params = (e && e.parameter) ? e.parameter : {};
    var action = params.action || '';
    var cb     = params.cb     || '';

    if (!action) {
      result = { ok: true, name: 'DJ-CE Backend v3', time: now() };
    } else {
      switch (action) {
        case 'submitBooking':  result = submitBooking(params); break;
        case 'getBookings':    result = getBookings(params);   break;
        case 'updateBooking':  result = updateBooking(params); break;
        case 'getAccounts':    result = getAccounts();         break;
        case 'saveAccount':    result = saveAccount(params);   break;
        case 'deleteAccount':  result = deleteAccount(params); break;
        case 'getClients':     result = getClients();          break;
        case 'saveClient':     result = saveClient(params);    break;
        case 'getComplaints':  result = getComplaints();       break;
        case 'saveComplaint':  result = saveComplaint(params); break;
        case 'logActivity':    result = logActivity(params);   break;
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
  var id = 'BK-' + new Date().getTime();
  var ts = now();

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

  var row = [
    id, ts,
    data.name    || '',
    data.phone   || '',
    data.email   || '',
    data.address || '',
    data.landmark|| '',
    data.lat     || '',
    data.lng     || '',
    svcSummary,
    svcDetails,
    data.date    || '',
    data.time    || '',
    data.technician || data.assigned || '',
    'Pending',
    data.amount  || '',
    data.usage   || '',
    data.nextCleaningDate || '',
    data.notes   || '',
    data.source  || 'website',
    'false',
    ''
  ];

  sheet.appendRow(row);
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 15).setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('bold');
  try { sendBookingEmail(id, data, svcSummary); } catch(e) {}
  return { ok: true, id: id, timestamp: ts };
}

function getBookings(p) {
  var sheet = getOrCreateSheet('Bookings');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, bookings: [] };
  var headers = rows[0];
  var bookings = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = String(row[i] || ''); });
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
      if (data.status)           setCellByKey(sheet, i+1, colMap, 'status', data.status);
      if (data.technician)       setCellByKey(sheet, i+1, colMap, 'technician', data.technician);
      if (data.amount)           setCellByKey(sheet, i+1, colMap, 'amount', data.amount);
      if (data.notes)            setCellByKey(sheet, i+1, colMap, 'notes', data.notes);
      if (data.usage)            setCellByKey(sheet, i+1, colMap, 'usageFrequency', data.usage);
      if (data.nextCleaningDate) setCellByKey(sheet, i+1, colMap, 'nextCleaningDate', data.nextCleaningDate);
      if (data.completedAt)      setCellByKey(sheet, i+1, colMap, 'completedAt', data.completedAt);
      if (data.followedUp)       setCellByKey(sheet, i+1, colMap, 'followedUp', data.followedUp);

      var statusCol = colMap['status'];
      if (statusCol) {
        var statusColors = {
          'Pending':   { bg: '#fef3c7', fg: '#92400e' },
          'Confirmed': { bg: '#dbeafe', fg: '#1e40af' },
          'Ongoing':   { bg: '#fde68a', fg: '#92400e' },
          'Done':      { bg: '#d1fae5', fg: '#065f46' },
          'Cancelled': { bg: '#fee2e2', fg: '#991b1b' }
        };
        var sc = statusColors[data.status];
        if (sc) sheet.getRange(i+1, statusCol).setBackground(sc.bg).setFontColor(sc.fg).setFontWeight('bold');
      }
      return { ok: true, updated: data.id };
    }
  }
  return { ok: false, error: 'Booking not found: ' + data.id };
}

/* ─── ACCOUNTS ────────────────────────────────────────────── */

function getAccounts() {
  var sheet = getOrCreateSheet('Accounts');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    seedDefaultAccounts(sheet);
    return { ok: true, accounts: getDefaultAccounts() };
  }
  var headers = rows[0];
  var accounts = {};
  rows.slice(1).forEach(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = String(row[i] || ''); });
    if (obj.username) accounts[obj.username] = obj;
  });
  if (Object.keys(accounts).length === 0) {
    seedDefaultAccounts(sheet);
    return { ok: true, accounts: getDefaultAccounts() };
  }
  return { ok: true, accounts: accounts };
}

function saveAccount(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Accounts');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.username)) {
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });
      if (data.password)     setCellByKey(sheet, i+1, colMap, 'password', data.password);
      if (data.name)         setCellByKey(sheet, i+1, colMap, 'name', data.name);
      if (data.role)         setCellByKey(sheet, i+1, colMap, 'role', data.role);
      if (data.avatar)       setCellByKey(sheet, i+1, colMap, 'avatar', data.avatar);
      if (data.lastPwChange) setCellByKey(sheet, i+1, colMap, 'lastPwChange', data.lastPwChange);
      if (data.forceChange !== undefined) setCellByKey(sheet, i+1, colMap, 'forceChange', String(data.forceChange));
      if (data.lastLogin)    setCellByKey(sheet, i+1, colMap, 'lastLogin', data.lastLogin);
      return { ok: true, action: 'updated', username: data.username };
    }
  }
  sheet.appendRow([
    data.username    || '',
    data.password    || '',
    data.name        || '',
    data.role        || 'helper',
    data.avatar      || '??',
    data.lastPwChange|| '',
    data.forceChange !== undefined ? String(data.forceChange) : 'true',
    now(),
    ''
  ]);
  return { ok: true, action: 'created', username: data.username };
}

function deleteAccount(p) {
  var username = p.username || '';
  if (username === 'owner') return { ok: false, error: 'Cannot delete owner account' };
  var sheet = getOrCreateSheet('Accounts');
  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === username) {
      sheet.deleteRow(i + 1);
      return { ok: true, deleted: username };
    }
  }
  return { ok: false, error: 'Account not found: ' + username };
}

function getDefaultAccounts() {
  return {
    owner: { username:'owner', name:'Owner', role:'owner', avatar:'OW', forceChange:'true', lastPwChange:'' }
  };
}

function seedDefaultAccounts(sheet) {
  var defaults = [
    ['owner',  'djce2025',  'Owner',             'owner',   'OW', '', 'true', now(), ''],
    ['admin',  'admin2025', 'Admin/Dispatcher',  'admin',   'AD', '', 'true', now(), ''],
    ['tech1',  'tech2025',  'Papa (Senior Tech)', 'sr_tech','PT', '', 'true', now(), ''],
    ['tech2',  'tech2025',  'Helper 1',          'helper',  'H1', '', 'true', now(), ''],
    ['tech3',  'tech2025',  'Helper 2',          'helper',  'H2', '', 'true', now(), '']
  ];
  defaults.forEach(function(row) { sheet.appendRow(row); });
}

/* ─── CLIENTS ─────────────────────────────────────────────── */

function getClients() {
  var sheet = getOrCreateSheet('Clients');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, clients: [] };
  var headers = rows[0];
  var clients = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = String(row[i] || ''); });
    return obj;
  });
  return { ok: true, clients: clients };
}

function saveClient(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Clients');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });
      if (data.name)             setCellByKey(sheet, i+1, colMap, 'name', data.name);
      if (data.phone)            setCellByKey(sheet, i+1, colMap, 'phone', data.phone);
      if (data.email)            setCellByKey(sheet, i+1, colMap, 'email', data.email);
      if (data.address)          setCellByKey(sheet, i+1, colMap, 'address', data.address);
      if (data.lastService)      setCellByKey(sheet, i+1, colMap, 'lastService', data.lastService);
      if (data.lastServiceDate)  setCellByKey(sheet, i+1, colMap, 'lastServiceDate', data.lastServiceDate);
      if (data.usage)            setCellByKey(sheet, i+1, colMap, 'usageFrequency', data.usage);
      if (data.nextCleaningDate) setCellByKey(sheet, i+1, colMap, 'nextCleaningDate', data.nextCleaningDate);
      if (data.notes)            setCellByKey(sheet, i+1, colMap, 'notes', data.notes);
      return { ok: true, action: 'updated', id: data.id };
    }
  }
  var id = data.id || ('CL-' + new Date().getTime());
  sheet.appendRow([
    id,
    data.name             || '',
    data.phone            || '',
    data.email            || '',
    data.address          || '',
    data.lastService      || '',
    data.lastServiceDate  || '',
    data.usage            || '',
    data.nextCleaningDate || '',
    data.notes            || '',
    data.source           || 'manual',
    now()
  ]);
  return { ok: true, action: 'created', id: id };
}

/* ─── COMPLAINTS ──────────────────────────────────────────── */

function getComplaints() {
  var sheet = getOrCreateSheet('Complaints');
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { ok: true, complaints: [] };
  var headers = rows[0];
  var complaints = rows.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[headerToKey(h)] = String(row[i] || ''); });
    return obj;
  });
  return { ok: true, complaints: complaints };
}

function saveComplaint(p) {
  var data = p.data ? JSON.parse(decodeURIComponent(p.data)) : p;
  var sheet = getOrCreateSheet('Complaints');
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(data.id)) {
      var colMap = {};
      headers.forEach(function(h, idx) { colMap[headerToKey(h)] = idx + 1; });
      if (data.status)     setCellByKey(sheet, i+1, colMap, 'status', data.status);
      if (data.resolution) setCellByKey(sheet, i+1, colMap, 'resolution', data.resolution);
      if (data.resolvedBy) setCellByKey(sheet, i+1, colMap, 'resolvedBy', data.resolvedBy);
      if (data.resolvedAt) setCellByKey(sheet, i+1, colMap, 'resolvedAt', data.resolvedAt);
      return { ok: true, action: 'updated', id: data.id };
    }
  }
  var id = data.id || ('CMP-' + new Date().getTime());
  sheet.appendRow([
    id,
    data.title       || '',
    data.clientName  || '',
    data.bookingId   || '',
    data.type        || 'Complaint',
    data.priority    || 'Normal',
    data.description || '',
    data.status      || 'Open',
    '',
    data.createdBy   || '',
    now(),
    '',
    ''
  ]);
  return { ok: true, action: 'created', id: id };
}

/* ─── ACTIVITY LOG ────────────────────────────────────────── */

function logActivity(p) {
  var sheet = getOrCreateSheet('ActivityLog');
  sheet.appendRow([
    now(),
    p.user    || 'System',
    p.message || '',
    p.type    || 'general',
    p.details || ''
  ]);
  var maxRows = 500;
  var total = sheet.getLastRow();
  if (total > maxRows + 1) {
    sheet.deleteRows(2, total - maxRows - 1);
  }
  return { ok: true };
}

/* ─── EMAIL ───────────────────────────────────────────────── */

function sendBookingEmail(id, data, svcSummary) {
  var subject = '📋 New Booking: ' + (data.name || 'Client') + ' — ' + (svcSummary || data.service || 'Service');
  var maps = (data.lat && data.lng)
    ? 'https://www.google.com/maps?q=' + data.lat + ',' + data.lng
    : 'https://maps.google.com/?q=' + encodeURIComponent(data.address || '');
  var body = [
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto">',
    '<div style="background:#0d1f3c;color:#fff;padding:20px 24px;border-radius:10px 10px 0 0">',
    '<h2 style="margin:0">🔧 ' + BUSINESS_NAME + '</h2>',
    '<p style="margin:6px 0 0;opacity:.8;font-size:13px">New booking received</p></div>',
    '<div style="background:#f8f6f1;padding:20px 24px"><table style="width:100%;border-collapse:collapse;font-size:14px">',
    '<tr><td style="padding:6px 0;color:#666;width:130px">Booking ID</td><td style="font-family:monospace;color:#0d1f3c;font-weight:bold">' + id + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Client</td><td><strong>' + (data.name||'—') + '</strong></td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Phone</td><td>' + (data.phone||'—') + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Address</td><td>' + (data.address||'—') + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Service</td><td><strong>' + (svcSummary||'—') + '</strong></td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Date</td><td>' + (data.date||'—') + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Time</td><td>' + (data.time||'—') + '</td></tr>',
    '<tr><td style="padding:6px 0;color:#666">Notes</td><td>' + (data.notes||'—') + '</td></tr>',
    '</table>',
    '<div style="margin-top:16px"><a href="' + maps + '" style="background:#0d1f3c;color:#fff;padding:10px 18px;border-radius:8px;text-decoration:none;font-size:13px;font-weight:bold">📍 View on Google Maps</a></div>',
    '</div>',
    '<div style="background:#0d1f3c;color:rgba(255,255,255,.6);padding:12px 24px;border-radius:0 0 10px 10px;font-size:11px;text-align:center">',
    BUSINESS_NAME + ' · Calumpit, Bulacan · 09762592458</div></div>'
  ].join('');
  MailApp.sendEmail({ to: NOTIF_EMAIL, subject: subject, htmlBody: body });
}

/* ─── HELPERS ─────────────────────────────────────────────── */

function getOrCreateSheet(name) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Ensure headers exist (handles manually created empty sheets)
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
  return header
    .toLowerCase()
    .replace(/[^a-z0-9 ]/g, '')
    .replace(/\s+(.)/g, function(_, c) { return c.toUpperCase(); })
    .replace(/\s/g, '');
}

function now() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
