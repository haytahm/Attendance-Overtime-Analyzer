'use strict';

const test = require('node:test');
const assert = require('node:assert');
const cm = require('../utils/columnMatcher');

test('detects simple English headers', () => {
  const { mapping, needsManualMapping } = cm.matchColumns(['empId', 'date', 'time']);
  assert.strictEqual(needsManualMapping, false);
  assert.strictEqual(mapping.employeeId.index, 0);
  assert.strictEqual(mapping.date.index, 1);
  assert.strictEqual(mapping.time.index, 2);
});

test('detects ZKTeco-style headers including punch type', () => {
  const headers = ['Employee ID', 'First Name', 'Department', 'Date', 'Time', 'Punch State'];
  const { mapping, needsManualMapping } = cm.matchColumns(headers);
  assert.strictEqual(needsManualMapping, false);
  assert.strictEqual(mapping.employeeId.index, 0);
  assert.strictEqual(mapping.employeeName.index, 1);
  assert.strictEqual(mapping.department.index, 2);
  assert.strictEqual(mapping.date.index, 3);
  assert.strictEqual(mapping.time.index, 4);
  assert.strictEqual(mapping.punchType.index, 5);
});

test('detects Arabic headers', () => {
  const headers = ['رقم الموظف', 'اسم الموظف', 'التاريخ', 'الوقت', 'نوع الحركة'];
  const { mapping } = cm.matchColumns(headers);
  assert.strictEqual(mapping.employeeId.index, 0);
  assert.strictEqual(mapping.employeeName.index, 1);
  assert.strictEqual(mapping.date.index, 2);
  assert.strictEqual(mapping.time.index, 3);
  assert.strictEqual(mapping.punchType.index, 4);
});

test('flags missing required columns for manual mapping', () => {
  const { needsManualMapping, missingRequired } = cm.matchColumns(['foo', 'bar', 'baz']);
  assert.strictEqual(needsManualMapping, true);
  assert.deepStrictEqual(missingRequired.sort(), ['date', 'employeeId', 'time']);
});

test('scoreHeaderRow ranks a real header above a title row', () => {
  const titleRow = cm.scoreHeaderRow(['Transaction', '', '', '']);
  const headerRow = cm.scoreHeaderRow(['Employee ID', 'First Name', 'Date', 'Time', 'Punch State']);
  assert.ok(headerRow > titleRow);
});

test('isCheckIn interprets common punch labels', () => {
  assert.strictEqual(cm.isCheckIn('Check In'), true);
  assert.strictEqual(cm.isCheckIn('Check Out'), false);
  assert.strictEqual(cm.isCheckIn('دخول'), true);
  assert.strictEqual(cm.isCheckIn('خروج'), false);
  assert.strictEqual(cm.isCheckIn('something'), null);
});
