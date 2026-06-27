'use strict';

const test = require('node:test');
const assert = require('node:assert');
const { analyze } = require('../services/attendanceService');

function rec(employeeId, date, time, extra = {}) {
  return { rowNumber: 1, employeeId, employeeName: '', rawDate: date, rawTime: time, punchType: '', department: '', ...extra };
}

test('computes work hours and overtime above the standard day', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'),
    rec('E1', '2024-08-27', '19:00:00'), // 11h span -> 1h overtime over 10
  ];
  const { daily, totals } = analyze(records, { standardHours: 10 });
  assert.strictEqual(daily.length, 1);
  assert.strictEqual(daily[0].firstIn, '08:00:00');
  assert.strictEqual(daily[0].lastOut, '19:00:00');
  assert.strictEqual(daily[0].workHours, 11);
  assert.strictEqual(daily[0].overtimeHours, 1);
  assert.strictEqual(totals.totalOvertimeHours, 1);
});

test('standard hours is configurable (8h day)', () => {
  const records = [
    rec('E1', '2024-08-27', '09:00:00'),
    rec('E1', '2024-08-27', '18:00:00'), // 9h span
  ];
  const { daily } = analyze(records, { standardHours: 8 });
  assert.strictEqual(daily[0].overtimeHours, 1);
});

test('punchType mode uses status to pick in/out even when out is logged first', () => {
  const records = [
    rec('E1', '2024-08-27', '17:00:00', { punchType: 'Check Out' }),
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Check In' }),
  ];
  const { daily } = analyze(records, { standardHours: 10, calculationMode: 'punchType' });
  assert.strictEqual(daily[0].firstIn, '08:00:00');
  assert.strictEqual(daily[0].lastOut, '17:00:00');
  assert.strictEqual(daily[0].workHours, 9);
});

/* ---- Calculation modes ---- */

test('default mode is firstLast (status labels ignored)', () => {
  // Break Out is the latest punch; in firstLast mode it becomes the check-out.
  const records = [
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Check In' }),
    rec('E1', '2024-08-27', '12:00:00', { punchType: 'Break Out' }),
    rec('E1', '2024-08-27', '18:00:00', { punchType: 'Break Out' }),
  ];
  const { daily, calculationMode } = analyze(records, { standardHours: 10 });
  assert.strictEqual(calculationMode, 'firstLast');
  assert.strictEqual(daily[0].firstIn, '08:00:00');
  assert.strictEqual(daily[0].lastOut, '18:00:00'); // latest punch, regardless of label
  assert.strictEqual(daily[0].complete, true);
  assert.strictEqual(daily[0].workHours, 10);
});

test('punchType mode flags missing check-out (no silent fallback)', () => {
  // A later punch exists, but it is not a recognized Check Out.
  const records = [
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Check In' }),
    rec('E1', '2024-08-27', '16:00:00', { punchType: 'Break Out' }),
  ];
  const { daily, issues } = analyze(records, { calculationMode: 'punchType' });
  assert.strictEqual(daily[0].complete, false);
  assert.strictEqual(daily[0].statusKey, 'MISSING_OUT');
  assert.strictEqual(daily[0].workHours, 0); // no fabricated hours
  assert.strictEqual(daily[0].firstIn, '08:00:00');
  assert.strictEqual(daily[0].lastOut, ''); // not guessed from the later punch
  assert.ok(issues.some((i) => i.issueKey === 'MISSING_OUT'));
});

test('punchType mode deducts break time from Break Out/In', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Check In' }),
    rec('E1', '2024-08-27', '12:00:00', { punchType: 'Break Out' }),
    rec('E1', '2024-08-27', '13:00:00', { punchType: 'Break In' }), // 1h break
    rec('E1', '2024-08-27', '19:00:00', { punchType: 'Check Out' }),
  ];
  const { daily } = analyze(records, { standardHours: 10, calculationMode: 'punchType' });
  assert.strictEqual(daily[0].firstIn, '08:00:00');
  assert.strictEqual(daily[0].lastOut, '19:00:00');
  assert.strictEqual(daily[0].breakMinutes, 60);
  assert.strictEqual(daily[0].workHours, 10); // 11h span - 1h break
  assert.strictEqual(daily[0].overtimeHours, 0);
});

test('multiple break pairs sum, and firstLast mode ignores breaks', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Check In' }),
    rec('E1', '2024-08-27', '10:00:00', { punchType: 'Break Out' }),
    rec('E1', '2024-08-27', '10:30:00', { punchType: 'Break In' }), // 30m
    rec('E1', '2024-08-27', '13:00:00', { punchType: 'Break Out' }),
    rec('E1', '2024-08-27', '14:00:00', { punchType: 'Break In' }), // 60m
    rec('E1', '2024-08-27', '18:00:00', { punchType: 'Check Out' }),
  ];
  const punch = analyze(records, { calculationMode: 'punchType' }).daily[0];
  assert.strictEqual(punch.breakMinutes, 90); // 30 + 60
  assert.strictEqual(punch.workHours, 8.5); // 10h span - 1.5h break

  // firstLast ignores break punches entirely (no manual break set).
  const first = analyze(records, { calculationMode: 'firstLast' }).daily[0];
  assert.strictEqual(first.breakMinutes, 0);
  assert.strictEqual(first.workHours, 10);
});

test('absence: working day with no record is marked absent', () => {
  const records = [
    rec('E1', '2024-09-02', '08:00:00'), rec('E1', '2024-09-02', '18:00:00'), // Mon
    // 2024-09-03 (Tue) missing
    rec('E1', '2024-09-04', '08:00:00'), rec('E1', '2024-09-04', '18:00:00'), // Wed
  ];
  const { employees, totals, issues } = analyze(records, { weekendDays: [5, 6] });
  assert.strictEqual(employees[0].absentDays, 1);
  assert.strictEqual(totals.absentDays, 1);
  const absent = issues.filter((i) => i.issueKey === 'ABSENT');
  assert.strictEqual(absent.length, 1);
  assert.strictEqual(absent[0].date, '2024-09-03');
});

test('absence: configured weekend days are not counted absent', () => {
  // Range 2024-09-05 (Thu) .. 2024-09-09 (Mon); Fri 6 + Sat 7 are weekend.
  const records = [
    rec('E1', '2024-09-05', '08:00:00'), rec('E1', '2024-09-05', '18:00:00'),
    rec('E1', '2024-09-09', '08:00:00'), rec('E1', '2024-09-09', '18:00:00'),
  ];
  const { employees } = analyze(records, { weekendDays: [5, 6] }); // Fri, Sat
  // Missing working days in range: Sun 2024-09-08 only (Fri/Sat excluded).
  assert.strictEqual(employees[0].absentDays, 1);
});

test('punchType mode flags missing check-in', () => {
  const records = [
    rec('E1', '2024-08-27', '17:00:00', { punchType: 'Check Out' }),
    rec('E1', '2024-08-27', '12:00:00', { punchType: 'Break Out' }),
  ];
  const { daily } = analyze(records, { calculationMode: 'punchType' });
  assert.strictEqual(daily[0].statusKey, 'MISSING_IN');
  assert.strictEqual(daily[0].complete, false);
});

test('punchType mode flags days with no recognized statuses', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Break In' }),
    rec('E1', '2024-08-27', '17:00:00', { punchType: 'Break Out' }),
  ];
  const { daily } = analyze(records, { calculationMode: 'punchType' });
  assert.strictEqual(daily[0].statusKey, 'MISSING_BOTH');
});

test('same punches, different modes, different results', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00', { punchType: 'Check In' }),
    rec('E1', '2024-08-27', '18:30:00', { punchType: 'Break Out' }),
  ];
  const firstLast = analyze(records, { standardHours: 10 }).daily[0];
  const punchType = analyze(records, { standardHours: 10, calculationMode: 'punchType' }).daily[0];
  assert.strictEqual(firstLast.complete, true);   // last punch used as out
  assert.strictEqual(firstLast.lastOut, '18:30:00');
  assert.strictEqual(punchType.complete, false);  // no recognized check-out
  assert.strictEqual(punchType.statusKey, 'MISSING_OUT');
});

test('detects incomplete day (single punch)', () => {
  const records = [rec('E1', '2024-08-27', '08:00:00')];
  const { daily, issues, totals } = analyze(records, {});
  assert.strictEqual(daily[0].complete, false);
  assert.strictEqual(daily[0].workHours, 0);
  assert.strictEqual(totals.incompleteDays, 1);
  assert.ok(issues.some((i) => i.issueType === 'Only one punch'));
});

test('handles invalid date/time without crashing', () => {
  const records = [
    rec('E1', 'bad-date', '08:00:00'),
    rec('E1', '2024-08-27', 'bad-time'),
    rec('E1', '2024-08-27', '08:00:00'),
    rec('E1', '2024-08-27', '18:00:00'),
  ];
  const { issues, daily } = analyze(records, {});
  assert.strictEqual(issues.filter((i) => i.issueType === 'Invalid date/time').length, 2);
  assert.strictEqual(daily.length, 1); // one valid day
});

test('late and early-leave minutes when schedule configured', () => {
  const records = [
    rec('E1', '2024-08-27', '08:15:00'),
    rec('E1', '2024-08-27', '16:45:00'),
  ];
  const { daily } = analyze(records, { startTime: '08:00', endTime: '17:00', standardHours: 10 });
  assert.strictEqual(daily[0].lateMinutes, 15);
  assert.strictEqual(daily[0].earlyLeaveMinutes, 15);
});

test('break minutes reduce work hours', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'),
    rec('E1', '2024-08-27', '18:00:00'), // 10h span
  ];
  const { daily } = analyze(records, { standardHours: 10, breakMinutes: 60 });
  assert.strictEqual(daily[0].workHours, 9);
});

test('overtime rounding and minimum threshold', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'),
    rec('E1', '2024-08-27', '18:20:00'), // 10h20m -> 20m overtime
  ];
  const rounded = analyze(records, { standardHours: 10, overtimeRounding: 'nearest30' });
  assert.strictEqual(rounded.daily[0].overtimeHours, 0.5); // 20m -> 30m

  const thresholded = analyze(records, { standardHours: 10, minOvertimeMinutes: 30 });
  assert.strictEqual(thresholded.daily[0].overtimeHours, 0); // 20m < 30m threshold
});

test('weekend days count fully as overtime', () => {
  // 2024-08-25 is a Sunday (weekday 0)
  const records = [
    rec('E1', '2024-08-25', '09:00:00'),
    rec('E1', '2024-08-25', '14:00:00'), // 5h
  ];
  const { daily } = analyze(records, { standardHours: 10, weekendDays: [0] });
  assert.strictEqual(daily[0].isWeekend, true);
  assert.strictEqual(daily[0].overtimeHours, 5);
});

test('employee summary aggregates across days', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'), rec('E1', '2024-08-27', '19:00:00'),
    rec('E1', '2024-08-28', '08:00:00'), rec('E1', '2024-08-28', '18:00:00'),
  ];
  const { employees } = analyze(records, { standardHours: 10 });
  assert.strictEqual(employees.length, 1);
  assert.strictEqual(employees[0].daysCount, 2);
  assert.strictEqual(employees[0].totalWorkHours, 21);
  assert.strictEqual(employees[0].totalOvertimeHours, 1);
});

/* ---- Days Count = unique attendance dates (not punch/row count) ---- */

test('4 punches on the same date count as 1 attendance day', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'),
    rec('E1', '2024-08-27', '12:00:00'),
    rec('E1', '2024-08-27', '13:00:00'),
    rec('E1', '2024-08-27', '18:00:00'),
  ];
  const { employees, daily } = analyze(records, {});
  assert.strictEqual(employees[0].daysCount, 1);
  assert.strictEqual(daily.length, 1); // one row per employee per date
});

test('2 punches on the same date count as 1 attendance day', () => {
  const records = [rec('E1', '2024-08-27', '08:00:00'), rec('E1', '2024-08-27', '17:00:00')];
  assert.strictEqual(analyze(records, {}).employees[0].daysCount, 1);
});

test('1 punch on a date counts as 1 day and is incomplete', () => {
  const records = [rec('E1', '2024-08-27', '08:00:00')];
  const { employees, daily } = analyze(records, {});
  assert.strictEqual(employees[0].daysCount, 1);
  assert.strictEqual(employees[0].incompleteDays, 1);
  assert.strictEqual(daily[0].complete, false);
});

test('duplicate punch rows (same time) do not inflate Days Count', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'),
    rec('E1', '2024-08-27', '08:00:00'), // exact duplicate
    rec('E1', '2024-08-27', '18:00:00'),
    rec('E1', '2024-08-27', '18:00:00'), // exact duplicate
  ];
  const { employees, daily } = analyze(records, {});
  assert.strictEqual(employees[0].daysCount, 1);
  assert.strictEqual(daily.length, 1);
});

test('30 different dates produce Days Count of 30', () => {
  const records = [];
  for (let d = 1; d <= 30; d++) {
    const day = String(d).padStart(2, '0');
    records.push(rec('E1', `2024-09-${day}`, '08:00:00'));
    records.push(rec('E1', `2024-09-${day}`, '18:00:00'));
  }
  const { employees } = analyze(records, {});
  assert.strictEqual(employees[0].daysCount, 30);
});

test('mixed date formats for the same day group into one day', () => {
  const records = [
    rec('E1', '2024-08-27', '08:00:00'), // ISO
    rec('E1', '27/08/2024', '18:00:00'), // day-first -> same normalized date
  ];
  const { employees, daily } = analyze(records, {});
  assert.strictEqual(employees[0].daysCount, 1);
  assert.strictEqual(daily[0].date, '2024-08-27');
});

test('detects report month and flags out-of-month days', () => {
  const records = [
    // 28 days in September (dominant month)
    ...Array.from({ length: 28 }, (_, i) => {
      const day = String(i + 1).padStart(2, '0');
      return [rec('E1', `2024-09-${day}`, '08:00:00'), rec('E1', `2024-09-${day}`, '18:00:00')];
    }).flat(),
    // 2 days bleeding into October
    rec('E1', '2024-10-01', '08:00:00'), rec('E1', '2024-10-01', '18:00:00'),
    rec('E1', '2024-10-02', '08:00:00'), rec('E1', '2024-10-02', '18:00:00'),
  ];
  const { employees, totals, period } = analyze(records, {});
  assert.strictEqual(period.month, '2024-09');
  assert.strictEqual(period.daysInMonth, 30);
  assert.strictEqual(totals.hasMultipleMonths, true);
  assert.strictEqual(employees[0].daysCount, 30);
  assert.strictEqual(employees[0].outOfPeriodDays, 2);
  assert.strictEqual(employees[0].exceedsMonth, false); // 30 == 30, not over
});

test('exceedsMonth flags impossible counts (>days in month)', () => {
  // 30 days in February (dominant) + extras -> exceeds Feb's 29 days (2024 leap)
  const records = [];
  for (let d = 1; d <= 29; d++) {
    const day = String(d).padStart(2, '0');
    records.push(rec('E1', `2024-02-${day}`, '08:00:00'), rec('E1', `2024-02-${day}`, '18:00:00'));
  }
  records.push(rec('E1', '2024-03-01', '08:00:00'), rec('E1', '2024-03-01', '18:00:00'));
  const { employees, period } = analyze(records, {});
  assert.strictEqual(period.daysInMonth, 29);
  assert.strictEqual(employees[0].daysCount, 30);
  assert.strictEqual(employees[0].exceedsMonth, true);
});
