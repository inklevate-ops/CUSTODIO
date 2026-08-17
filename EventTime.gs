/*******************************************************
 * CUSTODIO SUPLICO FIREWORKS — EVENT TIME
 *
 * These functions belong in Code.gs. They are kept in a
 * separate file only because the rest of Code.gs has not
 * been pushed to this repository yet.
 *
 * When Code.gs is assembled, move these definitions into
 * it and delete this file. Do NOT deploy both, or the
 * Apps Script project will hold two definitions of
 * normalizeEventTime_.
 *******************************************************/

/**
 * Normalizes any stored Event Time representation to 12-hour text.
 *
 * Accepts Date objects, 'HH:mm', 'HH:mm:ss', 12-hour text with optional
 * seconds, a Sheets day-fraction number, and ISO/1899 datetime strings.
 * Returns '' when the value is empty or cannot be understood.
 */
function normalizeEventTime_(value) {
  if (value === null || value === undefined) return '';

  if (Object.prototype.toString.call(value) === '[object Date]') {
    return formatEventTime12FromDate_(value);
  }

  if (typeof value === 'number' && isFinite(value)) {
    return formatEventTimeFromDayFraction_(value);
  }

  const text = String(value).trim();
  if (!text) return '';

  // 12-hour text: '11:30 PM', '1:05:30 am', '11:30 p.m.'
  let parts = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp])\.?[Mm]\.?$/);
  if (parts) {
    let hours = parseInt(parts[1], 10);
    const minutes = parseInt(parts[2], 10);
    if (hours < 1 || hours > 12 || minutes > 59) return '';
    if (hours === 12) hours = 0;
    if (parts[4].toLowerCase() === 'p') hours += 12;
    return formatEventTime12FromHoursMinutes_(hours, minutes);
  }

  // 24-hour text: '23:30' or '23:30:00'
  parts = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (parts) {
    const hours = parseInt(parts[1], 10);
    const minutes = parseInt(parts[2], 10);
    if (hours > 23 || minutes > 59) return '';
    return formatEventTime12FromHoursMinutes_(hours, minutes);
  }

  // ISO / 1899 datetime text, e.g. '1899-12-30T23:30:00.000Z'. The wall-clock
  // digits are read straight from the string so the zone offset that Apps
  // Script applies to 1899 dates never enters the result.
  parts = text.match(/^\d{4}-\d{2}-\d{2}[T ](\d{1,2}):(\d{2})(?::(\d{2}))?/);
  if (parts) {
    return formatEventTime12FromHoursMinutes_(
      parseInt(parts[1], 10),
      parseInt(parts[2], 10)
    );
  }

  // Day fraction that arrived as text, e.g. '0.9791666666666666'.
  if (/^\d*\.\d+$/.test(text)) {
    return formatEventTimeFromDayFraction_(parseFloat(text));
  }

  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return formatEventTime12FromDate_(parsed);

  return '';
}

/**
 * Formats an hour/minute pair as 12-hour text, e.g. 23 and 30 -> '11:30 PM'.
 */
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

/**
 * Formats a Date as 12-hour text.
 *
 * Reads the clock fields directly. Utilities.formatDate shifts dates on the
 * 1899 epoch that Sheets uses for time-only cells by the timezone's historical
 * offset, which turns 11:30 PM into a different time.
 */
function formatEventTime12FromDate_(date) {
  if (Object.prototype.toString.call(date) !== '[object Date]') return '';
  if (isNaN(date.getTime())) return '';
  return formatEventTime12FromHoursMinutes_(date.getHours(), date.getMinutes());
}

/**
 * Formats the fraction-of-a-day number Sheets uses for time-only cells.
 */
function formatEventTimeFromDayFraction_(value) {
  if (!isFinite(value)) return '';

  const fraction = value - Math.floor(value);
  const totalMinutes = Math.round(fraction * 24 * 60);

  return formatEventTime12FromHoursMinutes_(
    Math.floor(totalMinutes / 60),
    totalMinutes % 60
  );
}
