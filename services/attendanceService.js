'use strict';

const {
  combineDateTime,
  normalizeDate,
  normalizeTime,
  diffHours,
  decimalToHM,
  round,
  roundMinutes,
  timeToMinutes,
  weekdayOf,
  eachDateInclusive,
} = require('../utils/dateUtils');
const { isCheckIn, punchKind } = require('../utils/columnMatcher');

/**
 * Attendance / overtime analysis.
 *
 * Pure functions only — given normalized records + settings, produce daily rows,
 * per-employee summaries, issues and totals. No I/O, no Excel, no Express.
 *
 * Preserves the original idea (first check-in, last check-out, overtime above a
 * standard day) but makes the standard configurable and adds punch-type support,
 * incomplete-day detection, late / early-leave and rounding.
 */

// How in/out punches are chosen for each day:
//   'firstLast'  -> ignore status labels; earliest punch = in, latest = out (default)
//   'punchType'  -> use recognized statuses (Check In/Out, دخول/خروج); a missing
//                   check-out is reported as an issue instead of guessing
const CALCULATION_MODES = ['firstLast', 'punchType'];

const DEFAULT_SETTINGS = {
  calculationMode: 'firstLast', // see CALCULATION_MODES
  standardHours: 10, // standard work day in hours (original behavior)
  startTime: '', // official start "HH:mm" (enables late detection)
  endTime: '', // official end "HH:mm" (enables early-leave detection)
  breakMinutes: 0, // unpaid break subtracted from work hours
  overtimeRounding: 'none', // 'none' | 'nearest15' | 'nearest30'
  minOvertimeMinutes: 0, // ignore overtime below this many minutes
  weekendDays: [], // array of weekday numbers 0..6 (0 = Sunday) to exclude
};

// Issue codes. The KEY is stable (used by the UI to localize); the value is the
// English label used in the exported report and as a fallback.
const ISSUE = {
  MISSING_IN: 'Missing check-in',
  MISSING_OUT: 'Missing check-out',
  MISSING_BOTH: 'No recognized check-in/out',
  SINGLE_PUNCH: 'Only one punch',
  INVALID: 'Invalid date/time',
  ABSENT: 'Absent',
};

/** Record an issue with both a stable key and the English label. */
function pushIssue(issues, base, key, details) {
  issues.push({
    employeeId: base.employeeId,
    employeeName: base.employeeName,
    date: base.date,
    issueType: ISSUE[key],
    issueKey: key,
    details,
  });
}

function mergeSettings(input = {}) {
  const s = { ...DEFAULT_SETTINGS, ...sanitizeSettings(input) };
  return s;
}

function sanitizeSettings(input) {
  const out = {};
  if (CALCULATION_MODES.includes(input.calculationMode)) {
    out.calculationMode = input.calculationMode;
  }
  if (input.standardHours !== undefined && input.standardHours !== '') {
    const n = Number(input.standardHours);
    if (!isNaN(n) && n >= 0) out.standardHours = n;
  }
  if (typeof input.startTime === 'string') out.startTime = normalizeTime(input.startTime) || '';
  if (typeof input.endTime === 'string') out.endTime = normalizeTime(input.endTime) || '';
  if (input.breakMinutes !== undefined && input.breakMinutes !== '') {
    const n = Number(input.breakMinutes);
    if (!isNaN(n) && n >= 0) out.breakMinutes = n;
  }
  if (['none', 'nearest15', 'nearest30'].includes(input.overtimeRounding)) {
    out.overtimeRounding = input.overtimeRounding;
  }
  if (input.minOvertimeMinutes !== undefined && input.minOvertimeMinutes !== '') {
    const n = Number(input.minOvertimeMinutes);
    if (!isNaN(n) && n >= 0) out.minOvertimeMinutes = n;
  }
  if (Array.isArray(input.weekendDays)) {
    out.weekendDays = input.weekendDays
      .map((d) => Number(d))
      .filter((d) => Number.isInteger(d) && d >= 0 && d <= 6);
  }
  return out;
}

/**
 * Main entry point.
 * @param {Array} records normalized records from excelService.toRecords
 * @param {Object} settingsInput partial settings
 * @returns {Object} { daily, employees, issues, totals, settings }
 */
function analyze(records, settingsInput) {
  const settings = mergeSettings(settingsInput);
  const hasPunchType = records.some((r) => r.punchType && r.punchType !== '');

  // Group records by employee + day.
  const groups = new Map(); // key -> { employeeId, employeeName, date, department, punches:[] }
  const issues = [];

  for (const rec of records) {
    const isoDate = normalizeDate(rec.rawDate);
    const isoTime = normalizeTime(rec.rawTime);

    if (!rec.employeeId) continue; // skip rows with no employee at all

    if (!isoDate || !isoTime) {
      pushIssue(
        issues,
        { employeeId: rec.employeeId, employeeName: rec.employeeName, date: isoDate || String(rec.rawDate || '') },
        'INVALID',
        `Row ${rec.rowNumber}: could not read ${!isoDate ? 'date' : 'time'} ("${!isoDate ? rec.rawDate : rec.rawTime}").`
      );
      continue;
    }

    const key = `${rec.employeeId}__${isoDate}`;
    if (!groups.has(key)) {
      groups.set(key, {
        employeeId: rec.employeeId,
        employeeName: rec.employeeName || '',
        date: isoDate,
        department: rec.department || '',
        punches: [],
      });
    }
    const group = groups.get(key);
    if (!group.employeeName && rec.employeeName) group.employeeName = rec.employeeName;
    if (!group.department && rec.department) group.department = rec.department;
    group.punches.push({
      time: isoTime,
      dateTime: combineDateTime(isoDate, isoTime),
      punchType: rec.punchType || '',
    });
  }

  // Compute one daily row per group.
  const daily = [];
  for (const group of groups.values()) {
    daily.push(computeDay(group, settings, issues));
  }

  // Stable sort: employee, then date.
  daily.sort((a, b) =>
    a.employeeId.localeCompare(b.employeeId, undefined, { numeric: true }) ||
    a.date.localeCompare(b.date)
  );

  // Detect the report period (dominant calendar month) so we can flag dates that
  // fall outside it — the usual reason an employee shows an "impossible" day count.
  const period = detectPeriod(daily);

  // Mark working days with no attendance record as absences.
  const absentByEmployee = detectAbsences(daily, settings, issues);

  const employees = buildEmployeeSummaries(daily, period, absentByEmployee);
  const totals = buildTotals(daily, employees, issues, period);

  return {
    daily,
    employees,
    issues,
    totals,
    settings,
    hasPunchType, // whether the file actually carries a punch-type column
    calculationMode: settings.calculationMode,
    period,
  };
}

const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
  'July', 'August', 'September', 'October', 'November', 'December'];

/**
 * Work out which month the file is "for" (the month with the most attendance
 * days) and report any months/days outside it. Pure.
 */
function detectPeriod(daily) {
  const monthCounts = new Map(); // 'YYYY-MM' -> day count
  for (const d of daily) {
    const m = d.date.slice(0, 7);
    monthCounts.set(m, (monthCounts.get(m) || 0) + 1);
  }
  const months = Array.from(monthCounts.keys()).sort();
  if (months.length === 0) {
    return { month: null, label: '', daysInMonth: 31, months: [], hasMultipleMonths: false, outOfPeriodDayCount: 0 };
  }
  // Dominant month = most days; ties broken by earliest month.
  let dominant = months[0];
  for (const m of months) {
    if (monthCounts.get(m) > monthCounts.get(dominant)) dominant = m;
  }
  const [year, mon] = dominant.split('-').map(Number);
  const daysInMonth = new Date(year, mon, 0).getDate(); // day 0 of next month
  const outOfPeriodDayCount = daily.filter((d) => d.date.slice(0, 7) !== dominant).length;
  return {
    month: dominant,
    label: `${MONTH_NAMES[mon - 1]} ${year}`,
    daysInMonth,
    months,
    hasMultipleMonths: months.length > 1,
    outOfPeriodDayCount,
  };
}

/** True if an ISO date is outside the detected report month. */
function isOutOfPeriod(isoDate, period) {
  return !!period.month && isoDate.slice(0, 7) !== period.month;
}

/**
 * Detect absences: within the date range covered by the file, any working day
 * (i.e. not a configured weekend) on which an employee has no attendance record
 * is counted as an absence and added to the issues list.
 *
 * Returns a Map of employeeId -> absent day count. Note: when no weekend days are
 * configured, every non-attended day in the range counts as absent — configure
 * weekend days for accurate absence reporting.
 */
function detectAbsences(daily, settings, issues) {
  const absentByEmployee = new Map();
  if (daily.length === 0) return absentByEmployee;

  const dates = daily.map((d) => d.date).sort();
  const range = eachDateInclusive(dates[0], dates[dates.length - 1]);

  // Per-employee: the dates they have a record on, plus their display name.
  const attended = new Map(); // empId -> Set(dates)
  const names = new Map();
  for (const d of daily) {
    if (!attended.has(d.employeeId)) {
      attended.set(d.employeeId, new Set());
      absentByEmployee.set(d.employeeId, 0);
    }
    attended.get(d.employeeId).add(d.date);
    if (!names.has(d.employeeId) && d.employeeName) names.set(d.employeeId, d.employeeName);
  }

  for (const [empId, attendedSet] of attended) {
    for (const date of range) {
      if (settings.weekendDays.includes(weekdayOf(date))) continue; // off day
      if (attendedSet.has(date)) continue; // present
      absentByEmployee.set(empId, absentByEmployee.get(empId) + 1);
      pushIssue(
        issues,
        { employeeId: empId, employeeName: names.get(empId) || '', date },
        'ABSENT',
        'No attendance record on a working day.'
      );
    }
  }
  return absentByEmployee;
}

/**
 * Determine first-in / last-out and metrics for a single employee-day.
 *
 * The selected calculation mode decides how in/out are chosen:
 *   'firstLast' (default): status labels are ignored — earliest punch is the
 *                          check-in, latest punch is the check-out.
 *   'punchType':           only recognized statuses are used (Check In/Out,
 *                          دخول/خروج). A missing check-out is reported as an
 *                          issue rather than guessing from the latest punch.
 */
function computeDay(group, settings, issues) {
  const usePunchType = settings.calculationMode === 'punchType';
  const punches = group.punches.filter((p) => p.dateTime);
  const base = {
    employeeId: group.employeeId,
    employeeName: group.employeeName,
    department: group.department,
    date: group.date,
    weekday: weekdayOf(group.date),
  };

  // Pick in / out punches according to the mode.
  let inPunch = null;
  let outPunch = null;
  let breakMinutes = settings.breakMinutes; // manual break unless punches say otherwise

  if (usePunchType) {
    // Status-driven: NO fallback to unrecognized punches.
    const kinds = punches.map((p) => ({ ...p, kind: punchKind(p.punchType) }));
    const ins = kinds.filter((p) => p.kind === 'in');
    const outs = kinds.filter((p) => p.kind === 'out');
    if (ins.length) inPunch = minBy(ins, (p) => p.dateTime.getTime());
    if (outs.length) outPunch = maxBy(outs, (p) => p.dateTime.getTime());

    // Compute the actual break taken from Break Out -> Break In pairs.
    const punchBreak = computeBreakMinutes(kinds);
    if (punchBreak !== null) breakMinutes = punchBreak; // overrides the manual setting
  } else {
    // First/last: ignore status labels entirely.
    if (punches.length) {
      inPunch = minBy(punches, (p) => p.dateTime.getTime());
      outPunch = maxBy(punches, (p) => p.dateTime.getTime());
    }
  }

  const isWeekend = settings.weekendDays.includes(base.weekday);

  // Incomplete-day detection.
  let statusKey = null; // null = complete
  if (punches.length === 0) {
    statusKey = 'INVALID';
  } else if (usePunchType) {
    if (!inPunch && !outPunch) statusKey = 'MISSING_BOTH';
    else if (!inPunch) statusKey = 'MISSING_IN';
    else if (!outPunch) statusKey = 'MISSING_OUT';
    else if (inPunch.dateTime.getTime() === outPunch.dateTime.getTime()) statusKey = 'SINGLE_PUNCH';
  } else {
    if (punches.length === 1 || inPunch.dateTime.getTime() === outPunch.dateTime.getTime()) {
      statusKey = 'SINGLE_PUNCH';
    }
  }

  if (statusKey) {
    const details = usePunchType
      ? `${punches.length} punch(es); missing a recognized ` +
        (statusKey === 'MISSING_IN' ? 'check-in.' : statusKey === 'MISSING_OUT' ? 'check-out.'
          : statusKey === 'MISSING_BOTH' ? 'check-in and check-out.' : 'pair (single timestamp).')
      : `Only ${punches.length} punch(es) recorded.`;
    pushIssue(issues, base, statusKey, details);
    return {
      ...base,
      firstIn: inPunch ? inPunch.time : '',
      lastOut: outPunch && outPunch !== inPunch ? outPunch.time : '',
      spanHours: 0,
      workHours: 0,
      overtimeHours: 0,
      breakMinutes: 0,
      lateMinutes: 0,
      earlyLeaveMinutes: 0,
      isWeekend,
      complete: false,
      statusKey,
      status: ISSUE[statusKey],
    };
  }

  // Complete day: compute hours (subtract the break, computed or manual).
  const spanHours = Math.max(0, diffHours(inPunch.dateTime, outPunch.dateTime));
  let workHours = spanHours - breakMinutes / 60;
  if (workHours < 0) workHours = 0;

  // Overtime above the standard day (weekends, if excluded, count fully as OT).
  let overtimeHours;
  if (isWeekend) {
    overtimeHours = workHours;
  } else {
    overtimeHours = Math.max(0, workHours - settings.standardHours);
  }
  overtimeHours = applyOvertimeRules(overtimeHours, settings);

  const lateMinutes = computeLate(inPunch.time, settings);
  const earlyLeaveMinutes = computeEarlyLeave(outPunch.time, settings);

  const roundedBreak = Math.round(breakMinutes);
  const notes = [];
  if (isWeekend) notes.push('Weekend');
  if (lateMinutes > 0) notes.push(`Late ${lateMinutes}m`);
  if (earlyLeaveMinutes > 0) notes.push(`Early ${earlyLeaveMinutes}m`);
  if (roundedBreak > 0) notes.push(`Break ${roundedBreak}m`);

  return {
    ...base,
    firstIn: inPunch.time,
    lastOut: outPunch.time,
    spanHours: round(spanHours, 2),
    workHours: round(workHours, 2),
    overtimeHours: round(overtimeHours, 2),
    breakMinutes: roundedBreak,
    lateMinutes,
    earlyLeaveMinutes,
    isWeekend,
    complete: true,
    statusKey: 'OK', // client composes a localized status from the flags above
    status: notes.length ? notes.join(', ') : 'OK',
  };
}

/**
 * Total break minutes from Break Out -> Break In pairs for a day.
 * Returns null when there are no break punches at all (so callers can fall back
 * to the manual break setting). Unmatched break punches are ignored.
 */
function computeBreakMinutes(kinds) {
  const outs = kinds.filter((p) => p.kind === 'breakOut').sort((a, b) => a.dateTime - b.dateTime);
  const ins = kinds.filter((p) => p.kind === 'breakIn').sort((a, b) => a.dateTime - b.dateTime);
  if (!outs.length && !ins.length) return null;

  let total = 0;
  const used = new Set();
  for (const o of outs) {
    // Pair with the earliest not-yet-used Break In at/after this Break Out.
    const idx = ins.findIndex((i, k) => !used.has(k) && i.dateTime >= o.dateTime);
    if (idx >= 0) {
      used.add(idx);
      total += (ins[idx].dateTime - o.dateTime) / 60000; // ms -> minutes
    }
  }
  return total;
}

/** Apply rounding + minimum-threshold to an overtime value (in hours). */
function applyOvertimeRules(overtimeHours, settings) {
  let minutes = overtimeHours * 60;
  minutes = roundMinutes(minutes, settings.overtimeRounding);
  if (minutes < settings.minOvertimeMinutes) minutes = 0;
  return minutes / 60;
}

function computeLate(firstInTime, settings) {
  if (!settings.startTime) return 0;
  const start = timeToMinutes(settings.startTime);
  const inMin = timeToMinutes(firstInTime);
  if (start === null || inMin === null) return 0;
  return Math.max(0, inMin - start);
}

function computeEarlyLeave(lastOutTime, settings) {
  if (!settings.endTime) return 0;
  const end = timeToMinutes(settings.endTime);
  const outMin = timeToMinutes(lastOutTime);
  if (end === null || outMin === null) return 0;
  return Math.max(0, end - outMin);
}

/**
 * Aggregate daily rows into per-employee summaries.
 *
 * Days Count is the number of UNIQUE attendance dates for the employee — never a
 * punch/row count — computed explicitly from a Set of normalized dates. Multiple
 * punches on the same date are already collapsed into one daily row upstream, but
 * the Set makes the "one date = one day" rule explicit and tamper-proof.
 */
function buildEmployeeSummaries(daily, period, absentByEmployee = new Map()) {
  const map = new Map();
  for (const d of daily) {
    if (!map.has(d.employeeId)) {
      map.set(d.employeeId, {
        employeeId: d.employeeId,
        employeeName: d.employeeName,
        department: d.department,
        dates: new Set(),
        outOfPeriodDates: new Set(),
        totalWorkHours: 0,
        totalOvertimeHours: 0,
        totalBreakMinutes: 0,
        totalLateMinutes: 0,
        totalEarlyLeaveMinutes: 0,
        incompleteDays: 0,
      });
    }
    const e = map.get(d.employeeId);
    if (!e.employeeName && d.employeeName) e.employeeName = d.employeeName;
    e.dates.add(d.date); // unique calendar dates only
    if (isOutOfPeriod(d.date, period)) e.outOfPeriodDates.add(d.date);
    e.totalWorkHours += d.workHours;
    e.totalOvertimeHours += d.overtimeHours;
    e.totalBreakMinutes += d.breakMinutes || 0;
    e.totalLateMinutes += d.lateMinutes;
    e.totalEarlyLeaveMinutes += d.earlyLeaveMinutes;
    if (!d.complete) e.incompleteDays += 1;
  }

  return Array.from(map.values())
    .map((e) => {
      const daysCount = e.dates.size;
      const outOfPeriodDays = e.outOfPeriodDates.size;
      const { dates, outOfPeriodDates, ...rest } = e;
      return {
        ...rest,
        daysCount,
        outOfPeriodDays,
        absentDays: absentByEmployee.get(e.employeeId) || 0,
        // Flag impossible counts: more attendance days than the month has.
        exceedsMonth: !!period.month && daysCount > period.daysInMonth,
        totalWorkHours: round(e.totalWorkHours, 2),
        totalOvertimeHours: round(e.totalOvertimeHours, 2),
        totalBreakMinutes: Math.round(e.totalBreakMinutes),
        totalWorkHoursText: decimalToHM(e.totalWorkHours),
        totalOvertimeHoursText: decimalToHM(e.totalOvertimeHours),
      };
    })
    .sort((a, b) => a.employeeId.localeCompare(b.employeeId, undefined, { numeric: true }));
}

function buildTotals(daily, employees, issues, period) {
  const totalWorkHours = round(daily.reduce((s, d) => s + d.workHours, 0), 2);
  const totalOvertimeHours = round(daily.reduce((s, d) => s + d.overtimeHours, 0), 2);
  const incompleteDays = daily.filter((d) => !d.complete).length;
  const lateRecords = daily.filter((d) => d.lateMinutes > 0).length;
  const earlyLeaveRecords = daily.filter((d) => d.earlyLeaveMinutes > 0).length;
  const absentDays = employees.reduce((s, e) => s + (e.absentDays || 0), 0);
  return {
    totalEmployees: employees.length,
    totalDays: daily.length, // unique employee-day rows
    totalWorkHours,
    totalWorkHoursText: decimalToHM(totalWorkHours),
    totalOvertimeHours,
    totalOvertimeHoursText: decimalToHM(totalOvertimeHours),
    incompleteDays,
    absentDays,
    lateRecords,
    earlyLeaveRecords,
    invalidRecords: issues.filter((i) => i.issueType === ISSUE.INVALID).length,
    reportMonth: period.label,
    daysInMonth: period.daysInMonth,
    hasMultipleMonths: period.hasMultipleMonths,
    months: period.months,
    outOfPeriodDays: period.outOfPeriodDayCount,
  };
}

// Small functional helpers (avoid spreading huge arrays into Math.min/max).
function minBy(arr, fn) {
  return arr.reduce((best, x) => (fn(x) < fn(best) ? x : best), arr[0]);
}
function maxBy(arr, fn) {
  return arr.reduce((best, x) => (fn(x) > fn(best) ? x : best), arr[0]);
}

module.exports = {
  DEFAULT_SETTINGS,
  CALCULATION_MODES,
  ISSUE,
  mergeSettings,
  analyze,
  computeDay,
  applyOvertimeRules,
};
