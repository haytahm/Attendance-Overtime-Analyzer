'use strict';

/**
 * Date / time helpers.
 *
 * All functions are pure (no I/O, no globals) so they are easy to unit-test.
 * They are deliberately tolerant of the many shapes Excel cells arrive in:
 * strings, JS Date objects and raw Excel serial numbers.
 */

// Excel's day 0 is 1899-12-30 (the well-known 1900 leap-year bug offset).
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const MS_PER_DAY = 24 * 60 * 60 * 1000;

/** Pad a number to 2 digits. */
function pad2(n) {
  return String(n).padStart(2, '0');
}

/**
 * Normalize a date-ish value to an ISO date string `YYYY-MM-DD`.
 * Returns null when the value cannot be understood.
 */
function normalizeDate(value) {
  if (value === null || value === undefined || value === '') return null;

  // Excel serial number (e.g. 45561) -> calendar date
  if (typeof value === 'number' && isFinite(value)) {
    const d = new Date(EXCEL_EPOCH_MS + Math.floor(value) * MS_PER_DAY);
    return isValidDate(d) ? toISODate(d) : null;
  }

  if (value instanceof Date) {
    return isValidDate(value) ? toISODate(value) : null;
  }

  const str = String(value).trim();
  if (!str) return null;

  // Already ISO-ish: 2024-08-27 or 2024/08/27 (optionally with time)
  let m = str.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (m) {
    const [, y, mo, da] = m;
    return buildISO(y, mo, da);
  }

  // Day-first: 27-08-2024, 27/08/2024, 27.08.2024
  m = str.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);
  if (m) {
    const [, da, mo, y] = m;
    return buildISO(y, mo, da);
  }

  // Fallback: let the engine try (handles "Aug 27 2024" etc.)
  const d = new Date(str);
  return isValidDate(d) ? toISODate(d) : null;
}

function buildISO(y, mo, da) {
  const month = Number(mo);
  const day = Number(da);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return `${y}-${pad2(month)}-${pad2(day)}`;
}

/**
 * Normalize a time-ish value to `HH:mm:ss` (24h). Returns null if unparseable.
 */
function normalizeTime(value) {
  if (value === null || value === undefined || value === '') return null;

  // Excel stores time-of-day as a fraction of a day (0.5 == 12:00).
  if (typeof value === 'number' && isFinite(value)) {
    let frac = value;
    if (frac >= 1) frac = frac - Math.floor(frac); // strip date part of a datetime serial
    const totalSeconds = Math.round(frac * 86400);
    return secondsToHMS(totalSeconds);
  }

  if (value instanceof Date && isValidDate(value)) {
    return `${pad2(value.getHours())}:${pad2(value.getMinutes())}:${pad2(value.getSeconds())}`;
  }

  let str = String(value).trim();
  if (!str) return null;

  // 12-hour clock with AM/PM
  const ampm = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp][Mm])$/);
  if (ampm) {
    let h = Number(ampm[1]) % 12;
    if (/[Pp]/.test(ampm[4])) h += 12;
    return `${pad2(h)}:${ampm[2]}:${ampm[3] || '00'}`;
  }

  // 24-hour clock HH:mm or HH:mm:ss
  const h24 = str.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (h24) {
    const h = Number(h24[1]);
    const mi = Number(h24[2]);
    const se = h24[3] ? Number(h24[3]) : 0;
    if (h > 23 || mi > 59 || se > 59) return null;
    return `${pad2(h)}:${pad2(mi)}:${pad2(se)}`;
  }

  return null;
}

function secondsToHMS(totalSeconds) {
  const s = ((totalSeconds % 86400) + 86400) % 86400;
  const h = Math.floor(s / 3600);
  const mi = Math.floor((s % 3600) / 60);
  const se = s % 60;
  return `${pad2(h)}:${pad2(mi)}:${pad2(se)}`;
}

/**
 * Combine a normalized date + time into a Date (local). Returns null if either
 * part is invalid.
 */
function combineDateTime(dateValue, timeValue) {
  const isoDate = normalizeDate(dateValue);
  const isoTime = normalizeTime(timeValue);
  if (!isoDate || !isoTime) return null;
  const d = new Date(`${isoDate}T${isoTime}`);
  return isValidDate(d) ? d : null;
}

function isValidDate(d) {
  return d instanceof Date && !isNaN(d.getTime());
}

function toISODate(d) {
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

/** Difference in decimal hours between two Dates (out - inn). */
function diffHours(inn, out) {
  return (out.getTime() - inn.getTime()) / (1000 * 60 * 60);
}

/** Format decimal hours as a human string, e.g. 10.5 -> "10h 30m". */
function decimalToHM(hours) {
  if (hours === null || hours === undefined || isNaN(hours)) return '0h 0m';
  const sign = hours < 0 ? '-' : '';
  const totalMinutes = Math.round(Math.abs(hours) * 60);
  const h = Math.floor(totalMinutes / 60);
  const m = totalMinutes % 60;
  return `${sign}${h}h ${m}m`;
}

/** Round half away from zero to N decimals. */
function round(value, decimals = 2) {
  const f = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * f) / f;
}

/**
 * Round a minutes value per the chosen overtime-rounding mode.
 *   'none'    -> unchanged
 *   'nearest15' -> nearest 15 minutes
 *   'nearest30' -> nearest 30 minutes
 */
function roundMinutes(minutes, mode) {
  switch (mode) {
    case 'nearest15':
      return Math.round(minutes / 15) * 15;
    case 'nearest30':
      return Math.round(minutes / 30) * 30;
    case 'none':
    default:
      return minutes;
  }
}

/** Convert "HH:mm" / "HH:mm:ss" to minutes since midnight, or null. */
function timeToMinutes(value) {
  const t = normalizeTime(value);
  if (!t) return null;
  const [h, m] = t.split(':').map(Number);
  return h * 60 + m;
}

/** 0 = Sunday ... 6 = Saturday, from an ISO date string. */
function weekdayOf(isoDate) {
  const d = new Date(`${isoDate}T00:00:00`);
  return isValidDate(d) ? d.getDay() : null;
}

/**
 * Every calendar date from startISO to endISO inclusive, as `YYYY-MM-DD` strings.
 * Iterates at noon to stay clear of DST boundaries. Returns [] on bad input or
 * if the range is reversed.
 */
function eachDateInclusive(startISO, endISO) {
  const start = new Date(`${startISO}T12:00:00`);
  const end = new Date(`${endISO}T12:00:00`);
  if (!isValidDate(start) || !isValidDate(end) || start > end) return [];
  const out = [];
  const cur = new Date(start);
  while (cur <= end) {
    out.push(toISODate(cur));
    cur.setDate(cur.getDate() + 1);
  }
  return out;
}

module.exports = {
  normalizeDate,
  normalizeTime,
  combineDateTime,
  isValidDate,
  diffHours,
  decimalToHM,
  round,
  roundMinutes,
  timeToMinutes,
  weekdayOf,
  eachDateInclusive,
  secondsToHMS,
};
