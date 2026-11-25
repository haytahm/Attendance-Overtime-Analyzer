'use strict';

const ExcelJS = require('exceljs');
const { decimalToHM } = require('../utils/dateUtils');

/**
 * Polished Excel report generation (ExcelJS).
 *
 * Supports English ('en', LTR) and Arabic ('ar', RTL) sheet names, headers,
 * statuses, issue labels and settings summary.
 */

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
const HEADER_FONT = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
const THIN_BORDER = { style: 'thin', color: { argb: 'FFD9D9D9' } };

const LABELS = {
  en: {
    sheets: {
      daily: 'Daily Attendance',
      employees: 'Employee Summary',
      issues: 'Issues',
      settings: 'Settings Used',
    },
    cols: {
      employeeId: 'Employee ID',
      employeeName: 'Employee Name',
      date: 'Date',
      firstIn: 'First In',
      lastOut: 'Last Out',
      workHours: 'Work Hours',
      overtimeHours: 'Overtime Hours',
      breakMinutes: 'Break',
      lateMinutes: 'Late',
      earlyLeaveMinutes: 'Early Leave',
      status: 'Status / Notes',
      daysCount: 'Attendance Days',
      totalWorkHours: 'Total Work Hours',
      totalOvertimeHours: 'Total Overtime Hours',
      totalBreakMinutes: 'Total Break',
      totalLateMinutes: 'Total Late',
      totalEarlyLeaveMinutes: 'Total Early Leave',
      incompleteDays: 'Incomplete Days',
      absentDays: 'Absent Days',
      issueType: 'Issue Type',
      details: 'Details',
      setting: 'Setting',
      value: 'Value',
    },
    issues: {
      MISSING_IN: 'Missing check-in',
      MISSING_OUT: 'Missing check-out',
      MISSING_BOTH: 'No recognized check-in/out',
      SINGLE_PUNCH: 'Only one punch',
      INVALID: 'Invalid date/time',
      ABSENT: 'Absent',
      BREAK_EXCEEDED: 'Break exceeded allowed limit',
    },
    status: {
      ok: 'OK',
      weekend: 'Weekend',
      late: (n) => `Late ${n}m`,
      early: (n) => `Early ${n}m`,
      brk: (n) => `Break ${n}m`,
      brkExceeded: (n) => `Break ${n}m (exceeded)`,
    },
    noIssuesType: 'None',
    noIssuesDetails: 'No issues detected.',
    weekdays: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    rounding: { none: 'No rounding', nearest15: 'Nearest 15 minutes', nearest30: 'Nearest 30 minutes' },
    mode: {
      punchType: 'Use punch status/type (Check In/Out)',
      firstLast: 'First & last punch of the day',
    },
    dateFormat: { auto: 'Auto-detect', DMY: 'Day / Month / Year', MDY: 'Month / Day / Year' },
    off: '— (off)',
    none: 'None',
    auto: '— (auto)',
    settings: {
      mode: 'Calculation mode',
      standardHours: 'Standard work hours / day',
      startTime: 'Official start time',
      endTime: 'Official end time',
      breakMinutes: 'Break duration (minutes)',
      maxBreakMinutes: 'Max allowed break / day (minutes)',
      rounding: 'Overtime rounding',
      minOvertime: 'Minimum overtime threshold (minutes)',
      weekends: 'Excluded weekend days',
      nightShift: 'Night shift cutoff',
      periodStart: 'Period start',
      periodEnd: 'Period end',
      dateFormat: 'Date format',
      secPeriod: '— Report Period —',
      reportMonth: 'Detected report month',
      daysInMonth: 'Days in month',
      monthsPresent: 'Months present in file',
      outOfPeriod: 'Attendance days outside report month',
      secTotals: '— Report Totals —',
      totalEmployees: 'Total employees',
      totalDays: 'Total attendance days (employee-days)',
      totalWork: 'Total work hours',
      totalOT: 'Total overtime hours',
      incompleteDays: 'Incomplete days',
      absentDays: 'Absent days',
      lateRecords: 'Late records',
      earlyRecords: 'Early leave records',
      generatedAt: 'Generated at',
    },
  },
  ar: {
    sheets: {
      daily: 'الحضور اليومي',
      employees: 'ملخص الموظفين',
      issues: 'المشاكل',
      settings: 'الإعدادات المستخدمة',
    },
    cols: {
      employeeId: 'رقم الموظف',
      employeeName: 'اسم الموظف',
      date: 'التاريخ',
      firstIn: 'أول دخول',
      lastOut: 'آخر خروج',
      workHours: 'ساعات العمل',
      overtimeHours: 'الساعات الإضافية',
      breakMinutes: 'الاستراحة',
      lateMinutes: 'التأخير',
      earlyLeaveMinutes: 'الانصراف المبكر',
      status: 'الحالة / ملاحظات',
      daysCount: 'أيام الحضور',
      totalWorkHours: 'إجمالي ساعات العمل',
      totalOvertimeHours: 'إجمالي الإضافي',
      totalBreakMinutes: 'إجمالي الاستراحة',
      totalLateMinutes: 'إجمالي التأخير',
      totalEarlyLeaveMinutes: 'إجمالي الانصراف المبكر',
      incompleteDays: 'الأيام غير المكتملة',
      absentDays: 'أيام الغياب',
      issueType: 'نوع المشكلة',
      details: 'التفاصيل',
      setting: 'الإعداد',
      value: 'القيمة',
    },
    issues: {
      MISSING_IN: 'دخول مفقود',
      MISSING_OUT: 'خروج مفقود',
      MISSING_BOTH: 'لا يوجد دخول/خروج معروف',
      SINGLE_PUNCH: 'بصمة واحدة فقط',
      INVALID: 'تاريخ/وقت غير صالح',
      ABSENT: 'غياب',
      BREAK_EXCEEDED: 'تجاوز حد الاستراحة',
    },
    status: {
      ok: 'سليم',
      weekend: 'نهاية أسبوع',
      late: (n) => `تأخير ${n} د`,
      early: (n) => `انصراف مبكر ${n} د`,
      brk: (n) => `استراحة ${n} د`,
      brkExceeded: (n) => `تجاوز الاستراحة (${n} د)`,
    },
    noIssuesType: 'لا يوجد',
    noIssuesDetails: 'لم يتم اكتشاف أي مشاكل.',
    weekdays: ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'],
    rounding: { none: 'بدون تقريب', nearest15: 'لأقرب ١٥ دقيقة', nearest30: 'لأقرب ٣٠ دقيقة' },
    mode: {
      punchType: 'الاعتماد على حالة/نوع البصمة',
      firstLast: 'أول وآخر بصمة في اليوم',
    },
    dateFormat: { auto: 'اكتشاف تلقائي', DMY: 'يوم / شهر / سنة', MDY: 'شهر / يوم / سنة' },
    off: '— (معطّل)',
    none: 'لا يوجد',
    auto: '— (تلقائي)',
    settings: {
      mode: 'طريقة الحساب',
      standardHours: 'ساعات العمل القياسية / اليوم',
      startTime: 'وقت بداية الدوام الرسمي',
      endTime: 'وقت نهاية الدوام الرسمي',
      breakMinutes: 'مدة الاستراحة (دقائق)',
      maxBreakMinutes: 'الحد الأقصى للاستراحة / اليوم (دقائق)',
      rounding: 'تقريب الساعات الإضافية',
      minOvertime: 'الحد الأدنى للإضافي (دقائق)',
      weekends: 'أيام نهاية الأسبوع المستثناة',
      nightShift: 'وقت قطع الوردية الليلية',
      periodStart: 'تاريخ بداية الفترة',
      periodEnd: 'تاريخ نهاية الفترة',
      dateFormat: 'تنسيق التاريخ',
      secPeriod: '— فترة التقرير —',
      reportMonth: 'شهر التقرير المكتشف',
      daysInMonth: 'عدد أيام الشهر',
      monthsPresent: 'الأشهر الموجودة بالملف',
      outOfPeriod: 'أيام خارج شهر التقرير',
      secTotals: '— الإجماليات —',
      totalEmployees: 'إجمالي الموظفين',
      totalDays: 'إجمالي أيام الحضور',
      totalWork: 'إجمالي ساعات العمل',
      totalOT: 'إجمالي الساعات الإضافية',
      incompleteDays: 'الأيام غير المكتملة',
      absentDays: 'أيام الغياب',
      lateRecords: 'سجلات التأخير',
      earlyRecords: 'سجلات الانصراف المبكر',
      generatedAt: 'تاريخ الإنشاء',
    },
  },
};

function formatDuration(minutes, unit, lang) {
  const m = Math.round(Number(minutes) || 0);
  if (m === 0) return 0;
  if (unit === 'minutes') return m;
  const h = Math.floor(m / 60);
  const rem = m % 60;
  if (lang === 'ar') {
    if (h && rem) return `${h}س ${rem}د`;
    if (h) return `${h}س`;
    return `${rem}د`;
  }
  if (h && rem) return `${h}h ${rem}m`;
  if (h) return `${h}h`;
  return `${rem}m`;
}

async function buildReport(analysis, lang = 'en', timeUnit = 'hours') {
  const L = LABELS[lang] || LABELS.en;
  const isRtl = lang === 'ar';

  const wb = new ExcelJS.Workbook();
  wb.creator = 'Attendance & Overtime Analyzer';
  wb.created = new Date();

  buildDailySheet(wb, analysis.daily, L, isRtl, timeUnit, lang);
  buildEmployeeSheet(wb, analysis.employees, L, isRtl, timeUnit, lang);
  buildIssuesSheet(wb, analysis.issues, L, isRtl);
  buildSettingsSheet(wb, analysis.settings, analysis.totals, L, isRtl);

  return wb.xlsx.writeBuffer();
}

/** Apply header styling, freeze row 1, auto filter, RTL (if Arabic) and column widths. */
function styleSheet(sheet, columns, isRtl) {
  sheet.columns = columns.map((c) => ({ header: c.header, key: c.key, width: c.width }));

  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { vertical: 'middle', horizontal: isRtl ? 'right' : 'left' };
    cell.border = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
  });
  headerRow.height = 20;

  sheet.views = [{ state: 'frozen', ySplit: 1, rightToLeft: isRtl }];
  sheet.autoFilter = {
    from: { row: 1, column: 1 },
    to: { row: 1, column: columns.length },
  };
}

function zebra(sheet) {
  sheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return;
    if (rowNumber % 2 === 0) {
      row.eachCell((cell) => {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F6FB' } };
      });
    }
  });
}

function formatDailyStatus(d, L) {
  if (!d.complete) return L.issues[d.statusKey] || d.status;
  const notes = [];
  if (d.isWeekend) notes.push(L.status.weekend);
  if (d.lateMinutes > 0) notes.push(L.status.late(d.lateMinutes));
  if (d.earlyLeaveMinutes > 0) notes.push(L.status.early(d.earlyLeaveMinutes));
  if (d.breakMinutes > 0) {
    notes.push(d.breakExceeded ? L.status.brkExceeded(d.breakMinutes) : L.status.brk(d.breakMinutes));
  }
  return notes.length ? notes.join(L === LABELS.ar ? '، ' : ', ') : L.status.ok;
}

function buildDailySheet(wb, daily, L, isRtl, timeUnit, lang) {
  const C = L.cols;
  const sheet = wb.addWorksheet(L.sheets.daily);
  styleSheet(sheet, [
    { header: C.employeeId, key: 'employeeId', width: 16 },
    { header: C.employeeName, key: 'employeeName', width: 24 },
    { header: C.date, key: 'date', width: 14 },
    { header: C.firstIn, key: 'firstIn', width: 12 },
    { header: C.lastOut, key: 'lastOut', width: 12 },
    { header: C.workHours, key: 'workHours', width: 14 },
    { header: C.overtimeHours, key: 'overtimeHours', width: 16 },
    { header: C.breakMinutes, key: 'breakMinutes', width: 15 },
    { header: C.lateMinutes, key: 'lateMinutes', width: 14 },
    { header: C.earlyLeaveMinutes, key: 'earlyLeaveMinutes', width: 18 },
    { header: C.status, key: 'status', width: 28 },
  ], isRtl);

  daily.forEach((d) => {
    sheet.addRow({
      employeeId: d.employeeId,
      employeeName: d.employeeName,
      date: d.date,
      firstIn: d.firstIn,
      lastOut: d.lastOut,
      workHours: d.workHours,
      overtimeHours: d.overtimeHours,
      breakMinutes: formatDuration(d.breakMinutes, timeUnit, lang),
      lateMinutes: formatDuration(d.lateMinutes, timeUnit, lang),
      earlyLeaveMinutes: formatDuration(d.earlyLeaveMinutes, timeUnit, lang),
      status: formatDailyStatus(d, L),
    });
  });
  numberFormat(sheet, ['workHours', 'overtimeHours'], '0.00');
  zebra(sheet);
}

function buildEmployeeSheet(wb, employees, L, isRtl, timeUnit, lang) {
  const C = L.cols;
  const sheet = wb.addWorksheet(L.sheets.employees);
  styleSheet(sheet, [
    { header: C.employeeId, key: 'employeeId', width: 16 },
    { header: C.employeeName, key: 'employeeName', width: 24 },
    { header: C.daysCount, key: 'daysCount', width: 15 },
    { header: C.totalWorkHours, key: 'totalWorkHours', width: 18 },
    { header: C.totalOvertimeHours, key: 'totalOvertimeHours', width: 19 },
    { header: C.totalBreakMinutes, key: 'totalBreakMinutes', width: 18 },
    { header: C.totalLateMinutes, key: 'totalLateMinutes', width: 18 },
    { header: C.totalEarlyLeaveMinutes, key: 'totalEarlyLeaveMinutes', width: 23 },
    { header: C.incompleteDays, key: 'incompleteDays', width: 18 },
    { header: C.absentDays, key: 'absentDays', width: 14 },
  ], isRtl);

  employees.forEach((e) => {
    sheet.addRow({
      employeeId: e.employeeId,
      employeeName: e.employeeName,
      daysCount: e.daysCount,
      totalWorkHours: e.totalWorkHours,
      totalOvertimeHours: e.totalOvertimeHours,
      totalBreakMinutes: formatDuration(e.totalBreakMinutes, timeUnit, lang),
      totalLateMinutes: formatDuration(e.totalLateMinutes, timeUnit, lang),
      totalEarlyLeaveMinutes: formatDuration(e.totalEarlyLeaveMinutes, timeUnit, lang),
      incompleteDays: e.incompleteDays,
      absentDays: e.absentDays,
    });
  });
  numberFormat(sheet, ['totalWorkHours', 'totalOvertimeHours'], '0.00');
  zebra(sheet);
}

function buildIssuesSheet(wb, issues, L, isRtl) {
  const C = L.cols;
  const sheet = wb.addWorksheet(L.sheets.issues);
  styleSheet(sheet, [
    { header: C.employeeId, key: 'employeeId', width: 16 },
    { header: C.employeeName, key: 'employeeName', width: 24 },
    { header: C.date, key: 'date', width: 14 },
    { header: C.issueType, key: 'issueType', width: 24 },
    { header: C.details, key: 'details', width: 50 },
  ], isRtl);

  if (issues.length === 0) {
    sheet.addRow({ employeeId: '', employeeName: '', date: '', issueType: L.noIssuesType, details: L.noIssuesDetails });
  } else {
    issues.forEach((i) => {
      sheet.addRow({
        employeeId: i.employeeId,
        employeeName: i.employeeName,
        date: i.date,
        issueType: L.issues[i.issueKey] || i.issueType,
        details: i.details,
      });
    });
  }
  zebra(sheet);
}

function buildSettingsSheet(wb, settings, totals, L, isRtl) {
  const C = L.cols;
  const S = L.settings;
  const sheet = wb.addWorksheet(L.sheets.settings);
  styleSheet(sheet, [
    { header: C.setting, key: 'setting', width: 34 },
    { header: C.value, key: 'value', width: 40 },
  ], isRtl);

  const roundingLabel = L.rounding[settings.overtimeRounding] || settings.overtimeRounding;
  const weekendLabel = settings.weekendDays.length
    ? settings.weekendDays.map((d) => L.weekdays[d]).join(isRtl ? '، ' : ', ')
    : L.none;
  const modeLabel = L.mode[settings.calculationMode] || settings.calculationMode;
  const dfLabel = L.dateFormat[settings.dateFormat] || settings.dateFormat;

  const rows = [
    [S.mode, modeLabel],
    [S.standardHours, settings.standardHours],
    [S.startTime, settings.startTime || L.off],
    [S.endTime, settings.endTime || L.off],
    [S.breakMinutes, settings.breakMinutes],
    [S.maxBreakMinutes, settings.maxBreakMinutes || L.off],
    [S.rounding, roundingLabel],
    [S.minOvertime, settings.minOvertimeMinutes],
    [S.weekends, weekendLabel],
    [S.nightShift, settings.nightShiftCutoff || L.off],
    [S.periodStart, settings.periodStart || L.auto],
    [S.periodEnd, settings.periodEnd || L.auto],
    [S.dateFormat, dfLabel],
    ['', ''],
    [S.secPeriod, ''],
    [S.reportMonth, totals.reportMonth || '—'],
    [S.daysInMonth, totals.daysInMonth],
    [S.monthsPresent, (totals.months || []).join(', ') || '—'],
    [S.outOfPeriod, totals.outOfPeriodDays || 0],
    ['', ''],
    [S.secTotals, ''],
    [S.totalEmployees, totals.totalEmployees],
    [S.totalDays, totals.totalDays],
    [S.totalWork, `${totals.totalWorkHours} (${decimalToHM(totals.totalWorkHours)})`],
    [S.totalOT, `${totals.totalOvertimeHours} (${decimalToHM(totals.totalOvertimeHours)})`],
    [S.incompleteDays, totals.incompleteDays],
    [S.absentDays, totals.absentDays],
    [S.lateRecords, totals.lateRecords],
    [S.earlyRecords, totals.earlyLeaveRecords],
    [S.generatedAt, new Date().toLocaleString()],
  ];
  rows.forEach(([setting, value]) => sheet.addRow({ setting, value }));
}

function numberFormat(sheet, keys, format) {
  sheet.columns.forEach((col) => {
    if (keys.includes(col.key)) col.numFmt = format;
  });
}

module.exports = { buildReport };
