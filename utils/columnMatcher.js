'use strict';

/**
 * Flexible column detection.
 *
 * Attendance exports come with wildly different headers (English, Arabic,
 * vendor-specific). This module matches the *logical* fields the analyzer needs
 * against whatever headers a sheet actually has, using a list of aliases plus a
 * normalized fuzzy compare. It is pure and testable.
 */

// Logical fields and their known aliases (English + Arabic).
// Matching is normalized (case/space/punctuation-insensitive), so entries only
// need to capture meaningfully different spellings.
const FIELD_ALIASES = {
  employeeId: [
    'empId', 'employeeId', 'employee id', 'employee no', 'employee number',
    'emp no', 'emp code', 'employee code', 'staff id', 'id', 'code', 'ac no',
    'رقم الموظف', 'كود الموظف', 'الرقم الوظيفي', 'رقم',
  ],
  employeeName: [
    'name', 'employeeName', 'employee name', 'full name', 'first name',
    'staff name', 'الاسم', 'اسم الموظف', 'الاسم بالكامل',
  ],
  date: [
    'date', 'attendanceDate', 'attendance date', 'punchDate', 'punch date',
    'work date', 'day', 'التاريخ', 'تاريخ الحركة', 'تاريخ',
  ],
  time: [
    'time', 'attendanceTime', 'attendance time', 'punchTime', 'punch time',
    'clock', 'الوقت', 'وقت الحركة', 'وقت',
  ],
  punchType: [
    'type', 'status', 'punchType', 'punch type', 'punch state', 'state',
    'in/out', 'inout', 'check type', 'direction', 'mode',
    'الحركة', 'نوع الحركة', 'دخول/خروج', 'الحالة',
  ],
  department: [
    'department', 'dept', 'site', 'branch', 'location', 'area', 'area name',
    'القسم', 'الموقع', 'الفرع', 'المنطقة',
  ],
};

const REQUIRED_FIELDS = ['employeeId', 'date', 'time'];
const OPTIONAL_FIELDS = ['employeeName', 'punchType', 'department'];
const ALL_FIELDS = [...REQUIRED_FIELDS, ...OPTIONAL_FIELDS];

/** Lowercase, trim, drop spaces/underscores/most punctuation. Keeps Arabic letters. */
function normalize(value) {
  if (value === null || value === undefined) return '';
  return String(value)
    .toLowerCase()
    .replace(/[\s_\-./\\]+/g, '')
    .replace(/[()[\]{}:#,]/g, '')
    .trim();
}

// Pre-normalize the alias table once.
const NORMALIZED_ALIASES = Object.fromEntries(
  Object.entries(FIELD_ALIASES).map(([field, aliases]) => [
    field,
    aliases.map(normalize),
  ])
);

/**
 * Score how well a single header matches a field's aliases.
 * 3 = exact alias match, 2 = header contains alias (or vice-versa), 0 = none.
 */
function scoreHeaderForField(header, field) {
  const h = normalize(header);
  if (!h) return 0;
  let best = 0;
  for (const alias of NORMALIZED_ALIASES[field]) {
    if (!alias) continue;
    if (h === alias) return 3;
    if (h.includes(alias) || alias.includes(h)) {
      // Avoid silly partials like a 1-char header matching everything.
      if (alias.length >= 2 && h.length >= 2) best = Math.max(best, 2);
    }
  }
  return best;
}

/**
 * Given an array of header strings, return the best mapping of logical field ->
 * { index, header, score }. Each header is used at most once; higher-confidence
 * matches win first.
 */
function matchColumns(headers) {
  const candidates = [];
  ALL_FIELDS.forEach((field) => {
    headers.forEach((header, index) => {
      const score = scoreHeaderForField(header, field);
      if (score > 0) candidates.push({ field, index, header, score });
    });
  });

  // Strongest matches first; settle ties by earlier column.
  candidates.sort((a, b) => b.score - a.score || a.index - b.index);

  const mapping = {};
  const usedFields = new Set();
  const usedIndexes = new Set();
  for (const c of candidates) {
    if (usedFields.has(c.field) || usedIndexes.has(c.index)) continue;
    mapping[c.field] = { index: c.index, header: c.header, score: c.score };
    usedFields.add(c.field);
    usedIndexes.add(c.index);
  }

  const missingRequired = REQUIRED_FIELDS.filter((f) => !(f in mapping));
  return {
    mapping,
    missingRequired,
    detected: Object.keys(mapping),
    needsManualMapping: missingRequired.length > 0,
  };
}

/**
 * Score a candidate header row (array of cells) by how many distinct logical
 * fields it can satisfy. Used to locate the real header row in messy exports
 * where the first row is a title (e.g. ZKTeco "Transaction" reports).
 */
function scoreHeaderRow(cells) {
  const headers = cells.map((c) => (c === null || c === undefined ? '' : String(c)));
  const { mapping } = matchColumns(headers);
  let score = 0;
  for (const field of Object.keys(mapping)) {
    score += REQUIRED_FIELDS.includes(field) ? mapping[field].score * 2 : mapping[field].score;
  }
  return score;
}

/**
 * Heuristic: does this string look like a "Check In" / "دخول" punch?
 * Returns true (in), false (out) or null (unknown -> caller falls back to time
 * order). Uses exact tokens for bare "in"/"out" so words like "something" that
 * merely contain "in" are not misread.
 */
function isCheckIn(value) {
  const v = normalize(value);
  if (!v) return null;
  // Out first: "checkout" also contains "in" as a substring, so check out wins.
  if (/(checkout|signout|clockout|punchout)/.test(v) || /(خروج|انصراف)/.test(v)) return false;
  if (/(checkin|signin|clockin|punchin)/.test(v) || /(دخول|حضور)/.test(v)) return true;
  if (['out', 'o', '1'].includes(v)) return false;
  if (['in', 'i', '0'].includes(v)) return true;
  return null;
}

/**
 * Classify a punch label into one of:
 *   'in' | 'out' | 'breakOut' | 'breakIn' | 'unknown'
 * Breaks are detected first (they are more specific). A "Break Out" is leaving
 * for a break; a "Break In" is returning from it.
 */
function punchKind(value) {
  const v = normalize(value);
  if (!v) return 'unknown';

  const isBreak = /(break|lunch|meal|استراحة|راحة|غداء)/.test(v);
  if (isBreak) {
    if (/(out|خروج|بداية|start)/.test(v)) return 'breakOut';
    if (/(in|دخول|نهاية|عودة|end|return)/.test(v)) return 'breakIn';
    return 'unknown';
  }

  const check = isCheckIn(v);
  if (check === true) return 'in';
  if (check === false) return 'out';
  return 'unknown';
}

module.exports = {
  FIELD_ALIASES,
  REQUIRED_FIELDS,
  OPTIONAL_FIELDS,
  ALL_FIELDS,
  normalize,
  matchColumns,
  scoreHeaderRow,
  scoreHeaderForField,
  isCheckIn,
  punchKind,
};
