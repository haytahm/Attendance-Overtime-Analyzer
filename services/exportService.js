'use strict';

const ExcelJS = require('exceljs');
const { decimalToHM } = require('../utils/dateUtils');

/**
 * Polished Excel report generation (ExcelJS).
 *
 * ExcelJS is used here (rather than SheetJS) because it supports the formatting
 * the report needs: bold headers, auto filters, frozen header row and sensible
 * column widths. Produces a 4-sheet workbook as a Buffer.
 */

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
const HEADER_FONT = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
const THIN_BORDER = { style: 'thin', color: { argb: 'FFD9D9D9' } };

async function buildReport(analysis) {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Attendance & Overtime Analyzer';
  wb.created = new Date();

  buildDailySheet(wb, analysis.daily);
  buildEmployeeSheet(wb, analysis.employees);
  buildIssuesSheet(wb, analysis.issues);
  buildSettingsSheet(wb, analysis.settings, analysis.totals);

  return wb.xlsx.writeBuffer();
}

/** Apply header styling, freeze row 1, auto filter and column widths. */
function styleSheet(sheet, columns) {
  sheet.columns = columns.map((c) => ({ header: c.header, key: c.key, width: c.width }));

  const headerRow = sheet.getRow(1);
  headerRow.eachCell((cell) => {
    cell.fill = HEADER_FILL;
    cell.font = HEADER_FONT;
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
    cell.border = { top: THIN_BORDER, bottom: THIN_BORDER, left: THIN_BORDER, right: THIN_BORDER };
  });
  headerRow.height = 20;

  sheet.views = [{ state: 'frozen', ySplit: 1 }];
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

function buildDailySheet(wb, daily) {
  const sheet = wb.addWorksheet('Daily Attendance');
  styleSheet(sheet, [
    { header: 'Employee ID', key: 'employeeId', width: 16 },
    { header: 'Employee Name', key: 'employeeName', width: 24 },
    { header: 'Date', key: 'date', width: 14 },
    { header: 'First In', key: 'firstIn', width: 12 },
    { header: 'Last Out', key: 'lastOut', width: 12 },
    { header: 'Work Hours', key: 'workHours', width: 12 },
    { header: 'Overtime Hours', key: 'overtimeHours', width: 15 },
    { header: 'Break Minutes', key: 'breakMinutes', width: 14 },
    { header: 'Late Minutes', key: 'lateMinutes', width: 13 },
    { header: 'Early Leave Minutes', key: 'earlyLeaveMinutes', width: 18 },
    { header: 'Status / Notes', key: 'status', width: 28 },
  ]);

  daily.forEach((d) => {
    sheet.addRow({
      employeeId: d.employeeId,
      employeeName: d.employeeName,
      date: d.date,
      firstIn: d.firstIn,
      lastOut: d.lastOut,
      workHours: d.workHours,
      overtimeHours: d.overtimeHours,
      breakMinutes: d.breakMinutes || 0,
      lateMinutes: d.lateMinutes,
      earlyLeaveMinutes: d.earlyLeaveMinutes,
      status: d.status,
    });
  });
  numberFormat(sheet, ['workHours', 'overtimeHours'], '0.00');
  zebra(sheet);
}

function buildEmployeeSheet(wb, employees) {
  const sheet = wb.addWorksheet('Employee Summary');
  styleSheet(sheet, [
    { header: 'Employee ID', key: 'employeeId', width: 16 },
    { header: 'Employee Name', key: 'employeeName', width: 24 },
    { header: 'Attendance Days', key: 'daysCount', width: 15 },
    { header: 'Total Work Hours', key: 'totalWorkHours', width: 16 },
    { header: 'Total Overtime Hours', key: 'totalOvertimeHours', width: 19 },
    { header: 'Total Break Minutes', key: 'totalBreakMinutes', width: 18 },
    { header: 'Total Late Minutes', key: 'totalLateMinutes', width: 17 },
    { header: 'Total Early Leave Minutes', key: 'totalEarlyLeaveMinutes', width: 23 },
    { header: 'Incomplete Days', key: 'incompleteDays', width: 16 },
    { header: 'Absent Days', key: 'absentDays', width: 13 },
  ]);

  employees.forEach((e) => {
    sheet.addRow({
      employeeId: e.employeeId,
      employeeName: e.employeeName,
      daysCount: e.daysCount,
      totalWorkHours: e.totalWorkHours,
      totalOvertimeHours: e.totalOvertimeHours,
      totalBreakMinutes: e.totalBreakMinutes,
      totalLateMinutes: e.totalLateMinutes,
      totalEarlyLeaveMinutes: e.totalEarlyLeaveMinutes,
      incompleteDays: e.incompleteDays,
      absentDays: e.absentDays,
    });
  });
  numberFormat(sheet, ['totalWorkHours', 'totalOvertimeHours'], '0.00');
  zebra(sheet);
}

function buildIssuesSheet(wb, issues) {
  const sheet = wb.addWorksheet('Issues');
  styleSheet(sheet, [
    { header: 'Employee ID', key: 'employeeId', width: 16 },
    { header: 'Employee Name', key: 'employeeName', width: 24 },
    { header: 'Date', key: 'date', width: 14 },
    { header: 'Issue Type', key: 'issueType', width: 20 },
    { header: 'Details', key: 'details', width: 50 },
  ]);

  if (issues.length === 0) {
    sheet.addRow({ employeeId: '', employeeName: '', date: '', issueType: 'None', details: 'No issues detected.' });
  } else {
    issues.forEach((i) => sheet.addRow(i));
  }
  zebra(sheet);
}

function buildSettingsSheet(wb, settings, totals) {
  const sheet = wb.addWorksheet('Settings Used');
  styleSheet(sheet, [
    { header: 'Setting', key: 'setting', width: 32 },
    { header: 'Value', key: 'value', width: 40 },
  ]);

  const roundingLabel = {
    none: 'No rounding',
    nearest15: 'Nearest 15 minutes',
    nearest30: 'Nearest 30 minutes',
  }[settings.overtimeRounding] || settings.overtimeRounding;

  const weekendLabel = settings.weekendDays.length
    ? settings.weekendDays.map((d) => WEEKDAY_NAMES[d]).join(', ')
    : 'None';

  const modeLabel = settings.calculationMode === 'punchType'
    ? 'Use punch status/type (Check In/Out)'
    : 'First & last punch of the day';

  const rows = [
    ['Calculation mode', modeLabel],
    ['Standard work hours / day', settings.standardHours],
    ['Official start time', settings.startTime || '— (late detection off)'],
    ['Official end time', settings.endTime || '— (early-leave detection off)'],
    ['Break duration (minutes)', settings.breakMinutes],
    ['Overtime rounding', roundingLabel],
    ['Minimum overtime threshold (minutes)', settings.minOvertimeMinutes],
    ['Excluded weekend days', weekendLabel],
    ['', ''],
    ['— Report Period —', ''],
    ['Detected report month', totals.reportMonth || '—'],
    ['Days in month', totals.daysInMonth],
    ['Months present in file', (totals.months || []).join(', ') || '—'],
    ['Attendance days outside report month', totals.outOfPeriodDays || 0],
    ['', ''],
    ['— Report Totals —', ''],
    ['Total employees', totals.totalEmployees],
    ['Total attendance days (employee-days)', totals.totalDays],
    ['Total work hours', `${totals.totalWorkHours} (${decimalToHM(totals.totalWorkHours)})`],
    ['Total overtime hours', `${totals.totalOvertimeHours} (${decimalToHM(totals.totalOvertimeHours)})`],
    ['Incomplete days', totals.incompleteDays],
    ['Absent days', totals.absentDays],
    ['Late records', totals.lateRecords],
    ['Early leave records', totals.earlyLeaveRecords],
    ['Generated at', new Date().toLocaleString()],
  ];
  rows.forEach(([setting, value]) => sheet.addRow({ setting, value }));
}

function numberFormat(sheet, keys, format) {
  sheet.columns.forEach((col) => {
    if (keys.includes(col.key)) col.numFmt = format;
  });
}

const WEEKDAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

module.exports = { buildReport };
