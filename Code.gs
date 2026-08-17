/*******************************************************
 * EVENT TIME FIX
 *
 * Paste these functions over the same-named functions in the
 * live Apps Script Code.gs. Add the new formatEventTime12* helpers
 * next to them. Do not add this as a second .gs file if Code.gs
 * already defines these names.
 *
 * Also replace formatTimeInputValue and formatTime12 in Index.html
 * with the versions in event-time-index-functions.js.
 *
 * Event Time stays in the existing Bookings "Event Time" column.
 * Calendar stays date-only. No new storage field.
 *******************************************************/

/**
 * Canonical Event Time: 12-hour text, e.g. 11:30 PM.
 */
function normalizeEventTime_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return formatEventTime12FromDate_(value);
  }
  if (typeof value === 'number' && isFinite(value)) {
    return formatEventTimeFromDayFraction_(value);
  }

  let s = String(value).trim();
  if (!s) return '';
  s = s.replace(/^'/, '').trim();
  if (!s) return '';

  // 12-hour, optional seconds: 11:30 PM, 11:30:00 PM
  let m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/i);
  if (m) {
    let h = Number(m[1]);
    const min = Number(m[2]);
    const ap = m[3].toUpperCase();
    if (ap === 'AM' && h === 12) h = 0;
    if (ap === 'PM' && h !== 12) h += 12;
    return formatEventTime12FromHoursMinutes_(h, min);
  }

  // HTML <input type="time"> / 24-hour: 23:30 or 23:30:00
  m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
  if (m) {
    return formatEventTime12FromHoursMinutes_(Number(m[1]), Number(m[2]));
  }

  // Sheets time serial as a day fraction
  if (/^0(?:\.\d+)?$/.test(s) || /^1(?:\.0+)?$/.test(s) || /^\d*\.\d+$/.test(s)) {
    const n = Number(s);
    if (isFinite(n)) return formatEventTimeFromDayFraction_(n);
  }

  // ISO/date-time, including 1899 Sheets serials from safeValue_().
  // Read local hours/minutes. Do not use Utilities.formatDate (1899 bug).
  // Do not lift HH:mm from a trailing-Z ISO string (those digits are UTC).
  if (/\d{4}-\d{2}-\d{2}[T ]/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return formatEventTime12FromDate_(d);
  }

  return s;
}

function formatEventTime12FromHoursMinutes_(hours, minutes) {
  hours = Number(hours);
  minutes = Number(minutes);
  if (!isFinite(hours) || !isFinite(minutes)) return '';
  hours = ((Math.floor(hours) % 24) + 24) % 24;
  minutes = ((Math.floor(minutes) % 60) + 60) % 60;
  const suffix = hours < 12 ? 'AM' : 'PM';
  const display = (hours % 12) === 0 ? 12 : (hours % 12);
  const padded = minutes < 10 ? '0' + minutes : String(minutes);
  return display + ':' + padded + ' ' + suffix;
}

function formatEventTime12FromDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return formatEventTime12FromHoursMinutes_(date.getHours(), date.getMinutes());
}

function formatEventTimeFromDayFraction_(value) {
  if (!isFinite(value)) return '';
  const fraction = value - Math.floor(value);
  if (fraction < 0 || fraction >= 1) return '';
  const totalMinutes = Math.round(fraction * 24 * 60);
  return formatEventTime12FromHoursMinutes_(Math.floor(totalMinutes / 60) % 24, totalMinutes % 60);
}

function getPersistedBookingEventTime_(bookingId) {
  const sheet = getSpreadsheet_().getSheetByName('Bookings');
  if (!sheet || sheet.getLastRow() < 2) return '';
  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf('ID') + 1;
  const timeCol = headers.indexOf('Event Time') + 1;
  if (!idCol || !timeCol) return '';

  const rowCount = sheet.getLastRow() - 1;
  const ids = sheet.getRange(2, idCol, rowCount, 1).getDisplayValues().flat();
  const idx = ids.findIndex(v => String(v).trim() === String(bookingId).trim());
  if (idx < 0) return '';

  // Read displayed text because getValues() can coerce a time cell into Date.
  const display = sheet.getRange(idx + 2, timeCol).getDisplayValue();
  return normalizeEventTime_(display);
}

function getBookingEventTime(token, bookingId) {
  const auth = requireAuth_(token);
  if (!auth.ok) return auth;
  const id = String(bookingId || '').trim();
  if (!id) return {ok:false, message:'Booking / show ID is required.'};

  const sheet = getSpreadsheet_().getSheetByName('Bookings');
  if (!sheet || sheet.getLastRow() < 2) return {ok:true, eventTime:''};

  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf('ID') + 1;
  const timeCol = headers.indexOf('Event Time') + 1;
  const altTimeCol = headers.indexOf('EventTime') + 1;
  const legacyTimeCol = headers.indexOf('Time') + 1;
  if (!idCol) return {ok:true, eventTime:''};

  const rowCount = sheet.getLastRow() - 1;
  const ids = sheet.getRange(2, idCol, rowCount, 1).getDisplayValues().flat();
  const idx = ids.findIndex(v => String(v).trim() === id);
  if (idx < 0) return {ok:true, eventTime:''};

  const rowNumber = idx + 2;
  let display = '';
  const cols = [timeCol, altTimeCol, legacyTimeCol].filter(Boolean);
  for (const col of cols) {
    display = String(sheet.getRange(rowNumber, col).getDisplayValue() || '').trim();
    if (display) break;
  }

  return {ok:true, eventTime:normalizeEventTime_(display)};
}

function setBookingEventTimeText_(bookingId, eventTime) {
  const sheet = getSpreadsheet_().getSheetByName('Bookings');
  if (!sheet || sheet.getLastRow() < 2) return;
  const headers = getHeaders_(sheet);
  const idCol = headers.indexOf('ID') + 1;
  const timeCol = headers.indexOf('Event Time') + 1;
  if (!idCol || !timeCol) return;

  const ids = sheet.getRange(2, idCol, sheet.getLastRow() - 1, 1).getDisplayValues().flat();
  const idx = ids.findIndex(function(v){ return String(v).trim() === String(bookingId).trim(); });
  if (idx < 0) return;

  const cell = sheet.getRange(idx + 2, timeCol);
  const normalized = normalizeEventTime_(eventTime);

  // Event Time is stored as 12-hour text (e.g. 11:30 PM) in the existing
  // Event Time column. Do not prepend an apostrophe. Do not add another column.
  cell.setNumberFormat('@');
  cell.setValue(normalized);
}
