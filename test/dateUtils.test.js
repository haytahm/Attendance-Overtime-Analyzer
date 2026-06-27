'use strict';

const test = require('node:test');
const assert = require('node:assert');
const d = require('../utils/dateUtils');

test('normalizeDate handles ISO, day-first and Excel serials', () => {
  assert.strictEqual(d.normalizeDate('2024-08-27'), '2024-08-27');
  assert.strictEqual(d.normalizeDate('27/08/2024'), '2024-08-27');
  assert.strictEqual(d.normalizeDate('2024/8/7'), '2024-08-07');
  assert.strictEqual(d.normalizeDate(45536), '2024-09-01'); // Excel serial
});

test('normalizeDate rejects garbage', () => {
  assert.strictEqual(d.normalizeDate('not-a-date'), null);
  assert.strictEqual(d.normalizeDate(''), null);
  assert.strictEqual(d.normalizeDate('2024-13-40'), null);
});

test('normalizeTime handles 24h, 12h and Excel fractions', () => {
  assert.strictEqual(d.normalizeTime('05:13:21'), '05:13:21');
  assert.strictEqual(d.normalizeTime('5:01'), '05:01:00');
  assert.strictEqual(d.normalizeTime('1:30 PM'), '13:30:00');
  assert.strictEqual(d.normalizeTime('12:00 AM'), '00:00:00');
  assert.strictEqual(d.normalizeTime(0.5), '12:00:00'); // half a day
});

test('normalizeTime rejects invalid times', () => {
  assert.strictEqual(d.normalizeTime('25:00'), null);
  assert.strictEqual(d.normalizeTime('aa:bb'), null);
  assert.strictEqual(d.normalizeTime(''), null);
});

test('diffHours and decimalToHM', () => {
  const a = new Date('2024-08-27T08:00:00');
  const b = new Date('2024-08-27T18:30:00');
  assert.strictEqual(d.diffHours(a, b), 10.5);
  assert.strictEqual(d.decimalToHM(10.5), '10h 30m');
  assert.strictEqual(d.decimalToHM(0), '0h 0m');
});

test('roundMinutes modes', () => {
  assert.strictEqual(d.roundMinutes(37, 'none'), 37);
  assert.strictEqual(d.roundMinutes(37, 'nearest15'), 30);
  assert.strictEqual(d.roundMinutes(38, 'nearest15'), 45);
  assert.strictEqual(d.roundMinutes(44, 'nearest30'), 30);
});

test('timeToMinutes', () => {
  assert.strictEqual(d.timeToMinutes('08:30'), 510);
  assert.strictEqual(d.timeToMinutes('00:00:00'), 0);
});
