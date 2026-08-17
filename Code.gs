/**
 * 8888 Complaint Monitoring and Reporting System
 * Google Apps Script + 2 Google Sheets (Users, Complaints) + Google Drive storage
 *
 * Setup: run setupSystem() once, then Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 */

var TZ = 'Asia/Manila';
var SESSION_TTL = 21600;
var MAX_UPLOAD = 8 * 1024 * 1024;

var OFFICE = {
  republic: 'Republic of the Philippines',
  department: 'Department of Labor and Employment',
  officeName: 'Regional Office',
  section: 'Human Resource Development Section (HRDS)',
  systemName: '8888 Citizens’ Complaint Hotline',
  subtitle: 'Monitoring and Reporting System'
};

var ROLES = ['Administrator', 'Encoder', 'Viewer'];
var CATEGORIES = [
  'General Labor Standards (GLS Concern)',
  'Follow-up',
  'Complaint Against DOLE',
  'Request for Assistance',
  'Other Concerns'
];
var FIELD_OFFICES = [
  'Albay Field Office',
  'Camarines Norte Field Office',
  'Camarines Sur Field Office',
  'Catanduanes Field Office',
  'Masbate Field Office',
  'Sorsogon Field Office',
  'Technical Support and Services Division (TSSD)',
  'Internal Management Services Division (IMSD)',
  'Mediation-Arbitration and Legal Service Unit (MALSU)',
  'Human Resource Development Section (HRDS)'
];
var MODES = ['Email', 'Personal Receipt', 'Internal Routing'];
var DOC_TYPES = [
  'Original 8888 Complaint Letter',
  'Signed Action Slip',
  'Reply Document',
  'HRDS Action Taken Document',
  'Other Supporting Document'
];
var FINAL_STATUSES = ['Completed', 'Pending', 'Requires Further Action'];

var USER_HEADERS = ['id', 'name', 'username', 'password', 'role'];
var COMPLAINT_HEADERS = [
  'id', 'ticket_number', 'date_received', 'subject', 'category', 'complainant_location', 'status',
  'nature_of_call', 'complainant_name', 'suggested_category', 'category_confidence',
  'original_file_id', 'original_file_name', 'original_file_url',
  'date_signed_rd', 'action_slip_file_id', 'action_slip_file_name', 'action_slip_file_url',
  'endorsed_to', 'mode', 'date_endorsed', 'date_received_fo', 'received_by',
  'reply_date', 'reply_file_id', 'reply_file_name', 'reply_file_url', 'deadline', 'compliance_status',
  'date_forwarded', 'hrds_remarks', 'hrds_file_id', 'hrds_file_name', 'hrds_file_url', 'final_status',
  'attachments_json', 'timeline_json', 'created_by', 'created_by_name', 'created_at', 'updated_at'
];

function doGet() {
  try {
    var output = htmlFile_();
    return output
      .setTitle('8888 Complaint Monitoring System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,sans-serif;background:#0b1f3a;color:#f6f1e8;min-height:100vh;padding:40px">' +
      '<h1>8888 Complaint Monitoring System</h1>' +
      '<p>The web app could not load the HTML file.</p>' +
      '<p>In Apps Script, add an HTML file named <b>Index</b> (not Index.html) and paste the Index.html contents into it.</p>' +
      '<pre style="white-space:pre-wrap;color:#f4e3a1">' + (err && err.message ? err.message : err) + '</pre></div>'
    )
      .setTitle('8888 CMS')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function htmlFile_() {
  var names = ['Index', 'index', 'Index.html', 'index.html'];
  var last = '';
  for (var i = 0; i < names.length; i++) {
    try {
      return HtmlService.createHtmlOutputFromFile(names[i]);
    } catch (e) {
      last = e.message || String(e);
    }
  }
  throw new Error(last || 'HTML file named Index was not found.');
}

function api(request) {
  request = request || {};
  var action = request.action;
  var payload = request.payload || {};
  try {
    if (action === 'login') return { ok: true, data: login_(payload) };
    if (action === 'logout') {
      if (request.token) CacheService.getScriptCache().remove('sess_' + request.token);
      return { ok: true, data: { ok: true } };
    }
    var user = requireUser_(request.token);
    switch (action) {
      case 'me': return { ok: true, data: { user: publicUser_(user), lookups: lookups_() } };
      case 'dashboard': return { ok: true, data: dashboard_(user) };
      case 'listComplaints': return { ok: true, data: listComplaints_(user, payload) };
      case 'getComplaint': return { ok: true, data: getComplaint_(user, payload) };
      case 'createComplaint': return { ok: true, data: createComplaint_(user, payload) };
      case 'updateComplaint': return { ok: true, data: updateComplaint_(user, payload) };
      case 'deleteComplaint': return { ok: true, data: deleteComplaint_(user, payload) };
      case 'extractDocument': return { ok: true, data: extractDocument_(user, payload) };
      case 'suggestCategory': return { ok: true, data: categorize_('', payload.subject, payload.nature_of_call) };
      case 'saveActionSlip': return { ok: true, data: saveActionSlip_(user, payload) };
      case 'saveEndorsement': return { ok: true, data: saveEndorsement_(user, payload) };
      case 'markReceived': return { ok: true, data: markReceived_(user, payload) };
      case 'saveReply': return { ok: true, data: saveReply_(user, payload) };
      case 'saveHrds': return { ok: true, data: saveHrds_(user, payload) };
      case 'uploadAttachment': return { ok: true, data: uploadAttachment_(user, payload) };
      case 'deleteAttachment': return { ok: true, data: deleteAttachment_(user, payload) };
      case 'reports': return { ok: true, data: reports_(user, payload) };
      case 'listUsers': return { ok: true, data: listUsers_(user) };
      case 'saveUser': return { ok: true, data: saveUser_(user, payload) };
      case 'deleteUser': return { ok: true, data: deleteUser_(user, payload) };
      default: throw new Error('Unknown action.');
    }
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

function setupSystem() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PASSWORD_SALT')) props.setProperty('PASSWORD_SALT', Utilities.getUuid());

  var ss;
  var ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('8888 Complaint Monitoring Database');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  ensureSheet_(ss, 'Users', USER_HEADERS);
  ensureSheet_(ss, 'Complaints', COMPLAINT_HEADERS);
  ss.getSheets().forEach(function (sh) {
    if (sh.getName() !== 'Users' && sh.getName() !== 'Complaints') ss.deleteSheet(sh);
  });

  var folderId = props.getProperty('DRIVE_FOLDER_ID');
  var folderOk = false;
  if (folderId) {
    try { DriveApp.getFolderById(folderId); folderOk = true; } catch (e) {}
  }
  if (!folderOk) {
    var folder = DriveApp.createFolder('8888 Complaint Documents');
    props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  }

  if (!readSheet_('Users').length) {
    var stamp = now_();
    [
      ['System Administrator', 'admin', 'Admin@8888', 'Administrator'],
      ['HRDS Encoder', 'encoder', 'Encoder@8888', 'Encoder'],
      ['Report Viewer', 'viewer', 'Viewer@8888', 'Viewer']
    ].forEach(function (u, i) {
      appendRow_('Users', { id: i + 1, name: u[0], username: u[1], password: hash_(u[2]), role: u[3] });
    });
  }
  if (!readSheet_('Complaints').length) seedComplaints_();

  return { ok: true, spreadsheetUrl: ss.getUrl(), message: 'Ready. Default login: admin / Admin@8888' };
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (!current.join('')) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setBackground('#0b1f3a').setFontColor('#f4e3a1').setFontWeight('bold').setHorizontalAlignment('center');
}

function lookups_() {
  return {
    categories: CATEGORIES,
    fieldOffices: FIELD_OFFICES,
    endorsementModes: MODES,
    docTypes: DOC_TYPES,
    finalStatuses: FINAL_STATUSES,
    roles: ROLES,
    office: OFFICE
  };
}

function now_() {
  return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
}

function hash_(password) {
  var salt = PropertiesService.getScriptProperties().getProperty('PASSWORD_SALT') || '8888';
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '::' + password)
    .map(function (b) {
      var v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? '0' + v : v;
    }).join('');
}

function db_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Run setupSystem() first.');
  return SpreadsheetApp.openById(id);
}

function sheet_(name) {
  var sh = db_().getSheetByName(name);
  if (!sh) throw new Error('Missing sheet: ' + name + '. Run setupSystem().');
  return sh;
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var out = fn();
    SpreadsheetApp.flush();
    return out;
  } finally { lock.releaseLock(); }
}

function cell_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
  }
  return v;
}

function readSheet_(name) {
  var sh = sheet_(name);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0].map(String);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    if (!values[i].join('')) continue;
    var obj = { _row: i + 1 };
    headers.forEach(function (h, idx) { obj[h] = cell_(values[i][idx]); });
    rows.push(obj);
  }
  return rows;
}

function nextId_(name) {
  var max = 0;
  readSheet_(name).forEach(function (r) { if (Number(r.id) > max) max = Number(r.id); });
  return max + 1;
}

function appendRow_(name, rec) {
  return withLock_(function () {
    var headers = name === 'Users' ? USER_HEADERS : COMPLAINT_HEADERS;
    rec.id = rec.id || nextId_(name);
    sheet_(name).appendRow(headers.map(function (h) { return rec[h] === undefined || rec[h] === null ? '' : rec[h]; }));
    return rec;
  });
}

function updateRow_(name, id, rec) {
  return withLock_(function () {
    var headers = name === 'Users' ? USER_HEADERS : COMPLAINT_HEADERS;
    var sh = sheet_(name);
    var data = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { rowIndex = i + 1; break; }
    }
    if (rowIndex < 0) throw new Error('Record not found.');
    var row = data[rowIndex - 1];
    var out = {};
    headers.forEach(function (h, idx) {
      out[h] = Object.prototype.hasOwnProperty.call(rec, h) ? (rec[h] === undefined || rec[h] === null ? '' : rec[h]) : cell_(row[idx]);
      sh.getRange(rowIndex, idx + 1).setValue(out[h]);
    });
    return out;
  });
}

function deleteRow_(name, id) {
  return withLock_(function () {
    var sh = sheet_(name);
    var data = sh.getDataRange().getValues();
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]) === String(id)) { sh.deleteRow(i + 1); return true; }
    }
    return false;
  });
}

function findComplaint_(idOrTicket) {
  var rows = readSheet_('Complaints');
  var key = String(idOrTicket || '').toUpperCase();
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(idOrTicket) || String(rows[i].ticket_number).toUpperCase() === key) return rows[i];
  }
  throw new Error('Complaint record not found.');
}

function parseJson_(raw, fallback) {
  if (!raw) return fallback;
  if (Array.isArray(raw) || typeof raw === 'object') return raw;
  try { return JSON.parse(raw); } catch (e) { return fallback; }
}

function requireUser_(token) {
  if (!token) throw new Error('Your session has expired. Please sign in again.');
  var raw = CacheService.getScriptCache().get('sess_' + token);
  if (!raw) throw new Error('Your session has expired. Please sign in again.');
  CacheService.getScriptCache().put('sess_' + token, raw, SESSION_TTL);
  var user = JSON.parse(raw);
  if (user.role === 'Administrator' || user.role === 'Encoder' || user.role === 'Viewer') return user;
  throw new Error('Invalid account.');
}

function requireWrite_(user) {
  if (user.role === 'Viewer') throw new Error('Viewers can only view records, dashboards, and reports.');
  return user;
}

function requireAdmin_(user) {
  if (user.role !== 'Administrator') throw new Error('Administrator access is required.');
  return user;
}

function publicUser_(u) {
  return { id: u.id, name: u.name, username: u.username, role: u.role };
}

function login_(payload) {
  var username = String(payload.username || '').trim().toLowerCase();
  var password = String(payload.password || '');
  if (!username || !password) throw new Error('Username and password are required.');
  var users = readSheet_('Users');
  var user = null;
  for (var i = 0; i < users.length; i++) {
    if (String(users[i].username).toLowerCase() === username) { user = users[i]; break; }
  }
  if (!user || hash_(password) !== String(user.password)) throw new Error('Invalid username or password.');
  var token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  var session = publicUser_(user);
  session.token = token;
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify(session), SESSION_TTL);
  return { token: token, user: publicUser_(user), lookups: lookups_() };
}

function ticketNorm_(value) {
  var t = String(value || '').trim().toUpperCase();
  var m = t.match(/8888[\s\-–—]*((?:20)\d{2})[\s\-–—]*(\d{3,8})/);
  if (m) return '8888-' + m[1] + '-' + ('0000' + m[2]).slice(-Math.max(4, m[2].length));
  return t;
}

function ticketTaken_(ticket, exceptId) {
  var rows = readSheet_('Complaints');
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].ticket_number).toUpperCase() === ticket.toUpperCase() && String(rows[i].id) !== String(exceptId || '')) return true;
  }
  return false;
}

function liveStatus_(c) {
  var status = c.status;
  var compliance = c.compliance_status || '';
  if (c.date_received_fo) {
    var deadline = c.deadline ? new Date(c.deadline) : new Date(new Date(c.date_received_fo).getTime() + 86400000);
    if (c.reply_date) {
      compliance = new Date(c.reply_date).getTime() <= deadline.getTime() ? 'Within 24 Hours' : 'Delayed';
    } else {
      compliance = new Date().getTime() <= deadline.getTime() ? 'Pending Reply' : 'Delayed';
    }
    if (['Completed', 'Requires Further Action', 'Forwarded to HRDS'].indexOf(status) === -1) status = compliance;
  }
  return { status: status, compliance: compliance, deadline: c.deadline || '' };
}

function decorate_(c) {
  var live = liveStatus_(c);
  var copy = {};
  COMPLAINT_HEADERS.forEach(function (h) { copy[h] = c[h]; });
  copy.attachments = parseJson_(c.attachments_json, []);
  copy.timeline = parseJson_(c.timeline_json, []);
  copy.status = live.status;
  copy.compliance_status = live.compliance || c.compliance_status;
  copy.canEndorse = !!(c.date_signed_rd && (c.action_slip_file_id || c.action_slip_file_url));
  return copy;
}

function addEvent_(c, type, text, user) {
  var timeline = parseJson_(c.timeline_json, []);
  timeline.push({ at: now_(), type: type, description: text, user: user.name });
  c.timeline_json = JSON.stringify(timeline);
  return c;
}

function addAttachment_(c, stored, docType, user) {
  var atts = parseJson_(c.attachments_json, []);
  atts.push({
    id: Utilities.getUuid(),
    document_type: docType,
    file_id: stored.file_id,
    file_name: stored.file_name,
    file_url: stored.file_url,
    preview_url: stored.preview_url,
    download_url: stored.download_url,
    uploaded_by: user.name,
    created_at: now_()
  });
  c.attachments_json = JSON.stringify(atts);
  return addEvent_(c, 'Attachment', docType + ' uploaded: ' + stored.file_name, user);
}

function rootFolder_() {
  var id = PropertiesService.getScriptProperties().getProperty('DRIVE_FOLDER_ID');
  if (!id) throw new Error('Run setupSystem() to create the Drive folder.');
  return DriveApp.getFolderById(id);
}

function ticketFolder_(ticket) {
  var name = String(ticket).replace(/[^\w\-]/g, '_');
  var root = rootFolder_();
  var it = root.getFoldersByName(name);
  return it.hasNext() ? it.next() : root.createFolder(name);
}

function saveFile_(ticket, fileObject) {
  if (!fileObject || !fileObject.data) throw new Error('Please upload a document.');
  var name = fileObject.name || 'upload.bin';
  if (!/\.(pdf|docx|doc|png|jpe?g|gif|webp)$/i.test(name)) throw new Error('Only PDF, DOCX, and image files are allowed.');
  var bytes = Utilities.base64Decode(fileObject.data);
  if (bytes.length > MAX_UPLOAD) throw new Error('File exceeds the 8 MB limit.');
  var blob = Utilities.newBlob(bytes, fileObject.mimeType || 'application/octet-stream', name);
  var file = ticketFolder_(ticket).createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  var id = file.getId();
  return {
    file_id: id,
    file_name: file.getName(),
    file_url: 'https://drive.google.com/file/d/' + id + '/view',
    preview_url: 'https://drive.google.com/file/d/' + id + '/preview',
    download_url: 'https://drive.google.com/uc?export=download&id=' + id
  };
}

function listComplaints_(user, filters) {
  filters = filters || {};
  var q = String(filters.q || filters.ticket_number || '').trim().toUpperCase();
  var month = String(filters.month || '').trim();
  var category = String(filters.category || '').trim();
  var status = String(filters.status || '').trim();
  var office = String(filters.field_office || '').trim();
  var items = readSheet_('Complaints').map(decorate_).filter(function (c) {
    if (q && String(c.ticket_number).toUpperCase().indexOf(q) === -1 && String(c.subject).toUpperCase().indexOf(q) === -1) return false;
    if (month && String(c.date_received).substring(0, 7) !== month) return false;
    if (category && c.category !== category) return false;
    if (status && c.status !== status) return false;
    if (office && c.endorsed_to !== office) return false;
    return true;
  });
  items.sort(function (a, b) {
    return String(b.date_received).localeCompare(String(a.date_received)) || String(b.ticket_number).localeCompare(String(a.ticket_number));
  });
  return { items: items, total: items.length };
}

function getComplaint_(user, payload) {
  return { complaint: decorate_(findComplaint_(payload.id || payload.ticket_number)) };
}

function validateComplaint_(payload, exceptId) {
  var ticket = ticketNorm_(payload.ticket_number);
  if (!ticket) throw new Error('8888 Reference/Ticket Number is required.');
  if (!payload.date_received) throw new Error('Date Received from Central Office is required.');
  if (!String(payload.subject || '').trim()) throw new Error('Complaint Subject is required.');
  if (CATEGORIES.indexOf(payload.category) === -1) throw new Error('Complaint Category is required.');
  if (!String(payload.complainant_location || '').trim()) throw new Error('Complainant Location is required.');
  if (ticketTaken_(ticket, exceptId)) throw new Error('Ticket number ' + ticket + ' already exists.');
  return ticket;
}

function createComplaint_(user, payload) {
  requireWrite_(user);
  var ticket = validateComplaint_(payload);
  var rec = {
    ticket_number: ticket,
    date_received: payload.date_received,
    subject: String(payload.subject).trim(),
    category: payload.category,
    complainant_location: String(payload.complainant_location).trim(),
    status: 'Pending RD Approval',
    nature_of_call: String(payload.nature_of_call || '').trim(),
    complainant_name: String(payload.complainant_name || '').trim(),
    suggested_category: payload.suggested_category || payload.category,
    category_confidence: payload.category_confidence || '',
    attachments_json: '[]',
    timeline_json: '[]',
    created_by: user.id,
    created_by_name: user.name,
    created_at: now_(),
    updated_at: now_()
  };
  addEvent_(rec, 'Created', 'Complaint ' + ticket + ' encoded from Central Office referral.', user);
  if (payload.file) {
    var stored = saveFile_(ticket, payload.file);
    rec.original_file_id = stored.file_id;
    rec.original_file_name = stored.file_name;
    rec.original_file_url = stored.file_url;
    addAttachment_(rec, stored, 'Original 8888 Complaint Letter', user);
  }
  rec = appendRow_('Complaints', rec);
  return { complaint: decorate_(rec) };
}

function updateComplaint_(user, payload) {
  requireWrite_(user);
  var existing = findComplaint_(payload.id);
  var ticket = validateComplaint_(payload, existing.id);
  var rec = Object.assign({}, existing, {
    ticket_number: ticket,
    date_received: payload.date_received,
    subject: String(payload.subject).trim(),
    category: payload.category,
    complainant_location: String(payload.complainant_location).trim(),
    nature_of_call: String(payload.nature_of_call || '').trim(),
    complainant_name: String(payload.complainant_name || '').trim(),
    suggested_category: payload.suggested_category || existing.suggested_category,
    category_confidence: payload.category_confidence || existing.category_confidence,
    updated_at: now_()
  });
  addEvent_(rec, 'Updated', 'Complaint details were updated.', user);
  if (payload.file) {
    var stored = saveFile_(ticket, payload.file);
    rec.original_file_id = stored.file_id;
    rec.original_file_name = stored.file_name;
    rec.original_file_url = stored.file_url;
    addAttachment_(rec, stored, payload.document_type || 'Original 8888 Complaint Letter', user);
  }
  return { complaint: decorate_(updateRow_('Complaints', existing.id, rec)) };
}

function deleteComplaint_(user, payload) {
  requireAdmin_(user);
  var c = findComplaint_(payload.id || payload.ticket_number);
  parseJson_(c.attachments_json, []).forEach(function (a) {
    if (a.file_id) try { DriveApp.getFileById(a.file_id).setTrashed(true); } catch (e) {}
  });
  try { ticketFolder_(c.ticket_number).setTrashed(true); } catch (e) {}
  deleteRow_('Complaints', c.id);
  return { ok: true };
}

function saveActionSlip_(user, payload) {
  requireWrite_(user);
  var c = findComplaint_(payload.complaint_id || payload.ticket_number);
  if (!payload.date_signed_rd) throw new Error('Date Signed by RD is required.');
  if (!payload.file && !c.action_slip_file_id) throw new Error('The signed Action Slip document is mandatory.');
  if (payload.file) {
    var stored = saveFile_(c.ticket_number, payload.file);
    c.action_slip_file_id = stored.file_id;
    c.action_slip_file_name = stored.file_name;
    c.action_slip_file_url = stored.file_url;
    addAttachment_(c, stored, 'Signed Action Slip', user);
  }
  c.date_signed_rd = payload.date_signed_rd;
  if (['Received', 'Pending RD Approval', ''].indexOf(String(c.status)) !== -1) c.status = 'RD Approved';
  c.updated_at = now_();
  addEvent_(c, 'RD Approval', 'Signed Action Slip recorded. Date signed by RD: ' + payload.date_signed_rd + '.', user);
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function saveEndorsement_(user, payload) {
  requireWrite_(user);
  var c = findComplaint_(payload.complaint_id || payload.ticket_number);
  if (!c.date_signed_rd || !c.action_slip_file_id) {
    throw new Error('Endorsement is blocked until the RD signature date and signed Action Slip are recorded.');
  }
  if (!payload.endorsed_to) throw new Error('Endorsed To Field Office / Concerned Unit is required.');
  if (MODES.indexOf(payload.mode) === -1) throw new Error('Mode of Endorsement is required.');
  if (!payload.date_endorsed) throw new Error('Date and Time Endorsed is required.');
  c.endorsed_to = payload.endorsed_to;
  c.mode = payload.mode;
  c.date_endorsed = payload.date_endorsed;
  c.status = c.date_received_fo ? c.status : 'Endorsed';
  c.updated_at = now_();
  addEvent_(c, 'Endorsed', 'Endorsed to ' + payload.endorsed_to + ' via ' + payload.mode + '.', user);
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function markReceived_(user, payload) {
  requireWrite_(user);
  var c = findComplaint_(payload.complaint_id || payload.ticket_number);
  if (!c.date_endorsed) throw new Error('Record the endorsement first.');
  if (!payload.date_received) throw new Error('Date and Time Received by Endorsed Office is required.');
  if (!payload.received_by) throw new Error('Received By is required.');
  c.date_received_fo = payload.date_received;
  c.received_by = payload.received_by;
  var d = new Date(payload.date_received);
  c.deadline = Utilities.formatDate(new Date(d.getTime() + 86400000), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  c.compliance_status = 'Pending Reply';
  c.status = 'Pending Reply';
  c.updated_at = now_();
  addEvent_(c, 'Received by Endorsed Office', 'Received by ' + payload.received_by + '. 24-hour monitoring started.', user);
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function saveReply_(user, payload) {
  requireWrite_(user);
  var c = findComplaint_(payload.complaint_id || payload.ticket_number);
  if (!c.date_received_fo) throw new Error('The endorsed office must receive the complaint before a reply can be recorded.');
  if (!payload.reply_date) throw new Error('Reply Received Date and Time is required.');
  if (!payload.file && !c.reply_file_id) throw new Error('Upload the reply document.');
  if (payload.file) {
    var stored = saveFile_(c.ticket_number, payload.file);
    c.reply_file_id = stored.file_id;
    c.reply_file_name = stored.file_name;
    c.reply_file_url = stored.file_url;
    addAttachment_(c, stored, 'Reply Document', user);
  }
  if (!c.deadline) {
    c.deadline = Utilities.formatDate(new Date(new Date(c.date_received_fo).getTime() + 86400000), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  }
  c.reply_date = payload.reply_date;
  c.compliance_status = new Date(payload.reply_date).getTime() <= new Date(c.deadline).getTime() ? 'Within 24 Hours' : 'Delayed';
  c.status = c.compliance_status;
  c.updated_at = now_();
  addEvent_(c, 'Reply', 'Reply received. Compliance status: ' + c.compliance_status + '.', user);
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function saveHrds_(user, payload) {
  requireWrite_(user);
  var c = findComplaint_(payload.complaint_id || payload.ticket_number);
  if (!c.reply_date) throw new Error('Record the field office reply before forwarding action taken to HRDS.');
  if (!payload.date_forwarded) throw new Error('Date Forwarded to HRDS is required.');
  if (!payload.remarks) throw new Error('HRDS Action Taken / Remarks is required.');
  if (FINAL_STATUSES.indexOf(payload.final_status) === -1) throw new Error('Final Status is required.');
  if (payload.file) {
    var stored = saveFile_(c.ticket_number, payload.file);
    c.hrds_file_id = stored.file_id;
    c.hrds_file_name = stored.file_name;
    c.hrds_file_url = stored.file_url;
    addAttachment_(c, stored, 'HRDS Action Taken Document', user);
  }
  c.date_forwarded = payload.date_forwarded;
  c.hrds_remarks = payload.remarks;
  c.final_status = payload.final_status;
  c.status = payload.final_status === 'Pending' ? 'Forwarded to HRDS' : payload.final_status;
  c.updated_at = now_();
  addEvent_(c, 'HRDS', 'Action taken forwarded to HRDS. Final status: ' + payload.final_status + '.', user);
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function uploadAttachment_(user, payload) {
  requireWrite_(user);
  var c = findComplaint_(payload.complaint_id || payload.ticket_number);
  var stored = saveFile_(c.ticket_number, payload.file);
  addAttachment_(c, stored, payload.document_type || 'Other Supporting Document', user);
  c.updated_at = now_();
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function deleteAttachment_(user, payload) {
  requireAdmin_(user);
  var c = findComplaint_(payload.complaint_id);
  var atts = parseJson_(c.attachments_json, []).filter(function (a) {
    if (a.id === payload.attachment_id || a.file_id === payload.file_id) {
      if (a.file_id) try { DriveApp.getFileById(a.file_id).setTrashed(true); } catch (e) {}
      return false;
    }
    return true;
  });
  c.attachments_json = JSON.stringify(atts);
  addEvent_(c, 'Attachment deleted', 'Administrator deleted an attachment.', user);
  c.updated_at = now_();
  return { complaint: decorate_(updateRow_('Complaints', c.id, c)) };
}

function dashboard_(user) {
  var items = listComplaints_(user, {}).items;
  var byCategory = {};
  CATEGORIES.forEach(function (cat) { byCategory[cat] = 0; });
  var byMonth = {};
  var byOffice = {};
  var within = 0, delayedReply = 0, pendingReply = 0;
  items.forEach(function (c) {
    byCategory[c.category] = (byCategory[c.category] || 0) + 1;
    var m = String(c.date_received || '').substring(0, 7);
    if (m) byMonth[m] = (byMonth[m] || 0) + 1;
    if (c.endorsed_to) byOffice[c.endorsed_to] = (byOffice[c.endorsed_to] || 0) + 1;
    if (c.compliance_status === 'Within 24 Hours') within++;
    if (c.compliance_status === 'Delayed') delayedReply++;
    if (c.compliance_status === 'Pending Reply') pendingReply++;
  });
  var replied = within + delayedReply;
  return {
    cards: {
      total: items.length,
      pending: items.filter(function (c) { return c.status !== 'Completed'; }).length,
      completed: items.filter(function (c) { return c.status === 'Completed'; }).length,
      delayed: items.filter(function (c) { return c.status === 'Delayed' || c.compliance_status === 'Delayed'; }).length
    },
    categoryDistribution: CATEGORIES.map(function (cat) { return { category: cat, count: byCategory[cat] || 0 }; }),
    monthlyTrend: Object.keys(byMonth).sort().map(function (m) { return { month: m, count: byMonth[m] }; }),
    fieldOfficeDistribution: Object.keys(byOffice).sort().map(function (o) { return { office: o, count: byOffice[o] }; }),
    compliance: { within: within, delayed: delayedReply, pending: pendingReply, rate: replied ? Math.round((within / replied) * 1000) / 10 : 0 }
  };
}

function reports_(user, payload) {
  var items = listComplaints_(user, payload || {}).items;
  var total = items.length;
  function pct(n) { return total ? Math.round((n / total) * 10000) / 100 : 0; }
  var statusSummary = {};
  var officeSummary = {};
  items.forEach(function (c) {
    statusSummary[c.status] = (statusSummary[c.status] || 0) + 1;
    var o = c.endorsed_to || 'Not yet endorsed';
    officeSummary[o] = (officeSummary[o] || 0) + 1;
  });
  return {
    period: (payload && payload.month) || 'All records',
    generatedAt: now_(),
    total: total,
    categorySummary: CATEGORIES.map(function (cat) {
      var count = items.filter(function (c) { return c.category === cat; }).length;
      return { category: cat, count: count, percentage: pct(count) };
    }),
    statusSummary: Object.keys(statusSummary).sort().map(function (s) {
      return { status: s, count: statusSummary[s], percentage: pct(statusSummary[s]) };
    }),
    officeSummary: Object.keys(officeSummary).sort().map(function (o) {
      return { office: o, count: officeSummary[o], percentage: pct(officeSummary[o]) };
    }),
    items: items
  };
}

function listUsers_(user) {
  requireAdmin_(user);
  return { items: readSheet_('Users').map(publicUser_) };
}

function saveUser_(user, payload) {
  requireAdmin_(user);
  var name = String(payload.name || '').trim();
  var username = String(payload.username || '').trim();
  if (!name || !username) throw new Error('Name and username are required.');
  if (ROLES.indexOf(payload.role) === -1) throw new Error('Invalid role.');
  var users = readSheet_('Users');
  var dup = users.filter(function (u) {
    return String(u.username).toLowerCase() === username.toLowerCase() && String(u.id) !== String(payload.id || '');
  })[0];
  if (dup) throw new Error('Username is already taken.');
  if (payload.id) {
    var existing = users.filter(function (u) { return String(u.id) === String(payload.id); })[0];
    if (!existing) throw new Error('User not found.');
    existing.name = name;
    existing.username = username;
    existing.role = payload.role;
    if (payload.password) {
      if (String(payload.password).length < 6) throw new Error('Password must be at least 6 characters.');
      existing.password = hash_(payload.password);
    }
    return { user: publicUser_(updateRow_('Users', existing.id, existing)) };
  }
  if (!payload.password || String(payload.password).length < 6) throw new Error('Password must be at least 6 characters.');
  var rec = appendRow_('Users', { name: name, username: username, password: hash_(payload.password), role: payload.role });
  return { user: publicUser_(rec) };
}

function deleteUser_(user, payload) {
  requireAdmin_(user);
  if (String(payload.id) === String(user.id)) throw new Error('You cannot delete your own account.');
  deleteRow_('Users', payload.id);
  return { ok: true };
}

function extractDocument_(user, payload) {
  requireWrite_(user);
  if (!payload.file) throw new Error('Upload a PDF, DOCX, or image of the original 8888 complaint.');
  var bytes = Utilities.base64Decode(payload.file.data);
  var blob = Utilities.newBlob(bytes, payload.file.mimeType || 'application/octet-stream', payload.file.name || 'complaint');
  var text = '';
  var method = 'none';
  try {
    text = ocrText_(blob);
    method = 'drive-ocr';
  } catch (e) {
    text = '';
  }
  var fields = parseFields_(text);
  var cat = categorize_(text, fields.subject, fields.nature_of_call);
  return {
    ticket_number: fields.ticket_number,
    subject: fields.subject,
    nature_of_call: fields.nature_of_call,
    complainant_location: fields.complainant_location,
    suggested_category: cat.category,
    category_confidence: cat.confidence,
    matched_keywords: cat.matched,
    extraction_ok: !!text,
    note: text
      ? 'Tracking fields extracted for verification. The complaint narrative was not stored.'
      : 'Automatic reading was limited. Please encode the tracking fields and confirm the category.'
  };
}

function ocrText_(blob) {
  if (typeof Drive === 'undefined' || !Drive.Files) throw new Error('Enable Drive API Advanced Service.');
  var resource = { title: '8888-extract-temp-' + new Date().getTime(), mimeType: 'application/vnd.google-apps.document' };
  var file = Drive.Files.insert(resource, blob, { convert: true, ocr: true, ocrLanguage: 'en' });
  try {
    return DocumentApp.openById(file.id).getBody().getText() || '';
  } finally {
    try { Drive.Files.trash(file.id); } catch (e) {
      try { DriveApp.getFileById(file.id).setTrashed(true); } catch (e2) {}
    }
  }
}

function parseFields_(raw) {
  var text = String(raw || '');
  var compact = text.replace(/\s+/g, ' ');
  var ticket = '';
  var m = compact.match(/8888[\s\-–—]*((?:20)\d{2})[\s\-–—]*(\d{3,8})/i);
  if (m) ticket = '8888-' + m[1] + '-' + ('0000' + m[2]).slice(-Math.max(4, m[2].length));
  function grab(re) {
    var x = compact.match(re);
    return x ? String(x[1]).trim().slice(0, 250) : '';
  }
  var subject = grab(/(?:complaint\s*subject|subject|re|concern)\s*[:\-]\s*(.+?)(?:\s{2,}|$)/i);
  if (!subject) {
    var lines = text.split('\n').map(function (l) { return l.trim(); }).filter(Boolean);
    for (var i = 0; i < lines.length; i++) {
      if (/8888|department of labor|republic of the|hotline/i.test(lines[i])) continue;
      if (lines[i].length > 12 && lines[i].length < 180) { subject = lines[i]; break; }
    }
  }
  return {
    ticket_number: ticket,
    subject: subject,
    nature_of_call: grab(/(?:nature of call|nature of concern|nature)\s*[:\-]\s*(.+?)(?:\s{2,}|$)/i),
    complainant_location: grab(/(?:complainant\s*location|location|address|municipality|city|province)\s*[:\-]\s*(.+?)(?:\s{2,}|$)/i)
  };
}

function score_(hay, list) {
  var score = 0, matched = [];
  list.forEach(function (item) {
    if (hay.indexOf(item.k) !== -1) { score += item.w; matched.push(item.k); }
  });
  return { score: score, matched: matched };
}

function categorize_(text, subject, nature) {
  var hay = (String(text || '') + ' ' + String(subject || '') + ' ' + String(nature || '')).toLowerCase();
  var ranked = [
    { category: 'General Labor Standards (GLS Concern)', r: score_(hay, [
      { k: 'unpaid wages', w: 4 }, { k: 'unpaid salary', w: 4 }, { k: '13th month', w: 4 }, { k: 'overtime', w: 3 },
      { k: 'holiday pay', w: 4 }, { k: 'employee benefits', w: 3 }, { k: 'minimum wage', w: 4 }, { k: 'labor standards', w: 4 },
      { k: 'underpayment', w: 3 }, { k: 'illegal deduction', w: 3 }, { k: 'night shift', w: 3 }, { k: 'service incentive', w: 3 }, { k: 'gls', w: 3 }
    ]) },
    { category: 'Follow-up', r: score_(hay, [
      { k: 'follow-up', w: 4 }, { k: 'follow up', w: 4 }, { k: 'status update', w: 3 }, { k: 'previous complaint', w: 4 },
      { k: 'previous ticket', w: 4 }, { k: 'any update', w: 3 }, { k: 'requesting update', w: 3 }
    ]) },
    { category: 'Complaint Against DOLE', r: score_(hay, [
      { k: 'complaint against dole', w: 5 }, { k: 'dole personnel', w: 4 }, { k: 'dole employee', w: 4 }, { k: 'dole staff', w: 4 },
      { k: 'dole inspector', w: 4 }, { k: 'delayed service', w: 4 }, { k: 'delayed processing', w: 3 }, { k: 'discourtesy', w: 3 }, { k: 'red tape', w: 3 }
    ]) },
    { category: 'Request for Assistance', r: score_(hay, [
      { k: 'request for assistance', w: 5 }, { k: 'requesting assistance', w: 4 }, { k: 'intervention', w: 3 },
      { k: 'guidance', w: 3 }, { k: 'kindly assist', w: 3 }, { k: 'mediation', w: 3 }, { k: 'support', w: 2 }
    ]) }
  ].sort(function (a, b) { return b.r.score - a.r.score; });
  var top = ranked[0];
  if (!top || top.r.score <= 0) return { category: 'Other Concerns', confidence: 0.45, matched: [] };
  var confidence = Math.min(0.97, 0.55 + top.r.score * 0.06);
  if (ranked[1] && ranked[1].r.score && top.r.score - ranked[1].r.score <= 1) confidence = Math.min(confidence, 0.62);
  return { category: top.category, confidence: Math.round(confidence * 100) / 100, matched: top.r.matched.slice(0, 8) };
}

function seedComplaints_() {
  var encoder = readSheet_('Users').filter(function (u) { return u.username === 'encoder'; })[0] || { id: 2, name: 'HRDS Encoder' };
  var samples = [
    ['8888-2026-0001', '2026-01-08', 'Unpaid wages and 13th month pay', 'General Labor Standards (GLS Concern)', 'Legazpi City, Albay', 'Completed', 'Albay Field Office', 'Juan Dela Cruz'],
    ['8888-2026-0002', '2026-01-15', 'Follow-up on previous overtime complaint', 'Follow-up', 'Naga City, Camarines Sur', 'Forwarded to HRDS', 'Camarines Sur Field Office', 'Maria Santos'],
    ['8888-2026-0003', '2026-02-03', 'Delayed processing of request at DOLE window', 'Complaint Against DOLE', 'Sorsogon City, Sorsogon', 'Delayed', 'Sorsogon Field Office', 'Pedro Reyes'],
    ['8888-2026-0004', '2026-02-18', 'Request for assistance on illegal deduction', 'Request for Assistance', 'Daet, Camarines Norte', 'Pending Reply', 'Camarines Norte Field Office', 'Ana Lopez'],
    ['8888-2026-0005', '2026-03-05', 'Holiday pay and rest day concerns', 'General Labor Standards (GLS Concern)', 'Virac, Catanduanes', 'Within 24 Hours', 'Catanduanes Field Office', 'Roberto Cruz'],
    ['8888-2026-0006', '2026-03-21', 'Inquiry on documentary requirements', 'Other Concerns', 'Masbate City, Masbate', 'RD Approved', '', ''],
    ['8888-2026-0007', '2026-04-02', 'Minimum wage underpayment', 'General Labor Standards (GLS Concern)', 'Tabaco City, Albay', 'Endorsed', 'Albay Field Office', 'Liza Fernandez'],
    ['8888-2026-0008', '2026-04-19', 'Request for guidance on employee benefits', 'Request for Assistance', 'Iriga City, Camarines Sur', 'Pending RD Approval', '', 'Carlos Mendoza'],
    ['8888-2026-0009', '2026-05-06', 'Complaint against DOLE inspector conduct', 'Complaint Against DOLE', 'Ligao City, Albay', 'Completed', 'Albay Field Office', 'Grace Bautista'],
    ['8888-2026-0010', '2026-06-11', 'Follow-up on unpaid overtime case', 'Follow-up', 'Pili, Camarines Sur', 'Requires Further Action', 'Camarines Sur Field Office', 'Mark Villanueva'],
    ['8888-2026-0011', '2026-07-09', 'Overtime pay and night shift differential', 'General Labor Standards (GLS Concern)', 'Bulan, Sorsogon', 'Pending RD Approval', '', 'Elena Ramos'],
    ['8888-2026-0012', '2026-08-04', 'Assistance for mediation with employer', 'Request for Assistance', 'Mobo, Masbate', 'Pending RD Approval', '', 'Noel Garcia']
  ];
  samples.forEach(function (s) {
    var rec = {
      ticket_number: s[0], date_received: s[1], subject: s[2], category: s[3], complainant_location: s[4], status: s[5],
      nature_of_call: s[3], complainant_name: s[7], suggested_category: s[3], category_confidence: 0.9,
      endorsed_to: s[6], attachments_json: '[]', timeline_json: '[]',
      created_by: encoder.id, created_by_name: encoder.name, created_at: s[1] + 'T08:30:00', updated_at: now_()
    };
    addEvent_(rec, 'Created', 'Complaint record created from Central Office referral.', encoder);
    var advanced = ['RD Approved', 'Endorsed', 'Pending Reply', 'Within 24 Hours', 'Delayed', 'Forwarded to HRDS', 'Completed', 'Requires Further Action'];
    if (advanced.indexOf(s[5]) !== -1) {
      rec.date_signed_rd = s[1];
      rec.action_slip_file_name = 'Signed Action Slip';
      rec.action_slip_file_id = 'seed';
      addEvent_(rec, 'RD Approval', 'Signed Action Slip recorded.', encoder);
    }
    if (['Endorsed', 'Pending Reply', 'Within 24 Hours', 'Delayed', 'Forwarded to HRDS', 'Completed', 'Requires Further Action'].indexOf(s[5]) !== -1 && s[6]) {
      rec.mode = 'Internal Routing';
      rec.date_endorsed = s[1] + 'T11:15:00';
      addEvent_(rec, 'Endorsed', 'Endorsed to ' + s[6] + '.', encoder);
      if (s[5] !== 'Endorsed') {
        rec.date_received_fo = s[1] + 'T13:00:00';
        rec.received_by = 'Field Office Desk Officer';
        rec.deadline = s[1] + 'T13:00:00';
        rec.deadline = Utilities.formatDate(new Date(new Date(rec.date_received_fo).getTime() + 86400000), TZ, "yyyy-MM-dd'T'HH:mm:ss");
        rec.compliance_status = 'Pending Reply';
        addEvent_(rec, 'Received by Endorsed Office', 'Received by Field Office Desk Officer.', encoder);
      }
    }
    if (['Within 24 Hours', 'Delayed', 'Forwarded to HRDS', 'Completed', 'Requires Further Action'].indexOf(s[5]) !== -1) {
      rec.reply_date = s[5] === 'Delayed' ? s[1] + 'T20:00:00' : s[1] + 'T16:00:00';
      rec.reply_file_name = 'Reply Document';
      rec.compliance_status = s[5] === 'Delayed' ? 'Delayed' : 'Within 24 Hours';
      addEvent_(rec, 'Reply', 'Reply recorded.', encoder);
    }
    if (['Forwarded to HRDS', 'Completed', 'Requires Further Action'].indexOf(s[5]) !== -1) {
      rec.date_forwarded = s[1];
      rec.hrds_remarks = 'Action taken forwarded to HRDS.';
      rec.final_status = s[5] === 'Forwarded to HRDS' ? 'Pending' : s[5];
      rec.hrds_file_name = 'HRDS Action Taken Document';
      addEvent_(rec, 'HRDS', 'Forwarded to HRDS.', encoder);
    }
    appendRow_('Complaints', rec);
  });
}
