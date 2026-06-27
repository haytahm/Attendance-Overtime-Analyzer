'use strict';

const xlsx = require('xlsx');
const { scoreHeaderRow, matchColumns } = require('../utils/columnMatcher');

/**
 * Reading and normalizing Excel workbooks.
 *
 * SheetJS (`xlsx`) is used for reading because it tolerates a wide range of
 * real-world files (.xlsx / .xls, odd encodings, vendor exports). Writing the
 * polished report is handled separately by exportService (ExcelJS).
 */

const MAX_HEADER_SCAN_ROWS = 15; // how deep to look for the real header row

/** A friendly error type so routes can return clean messages instead of stacks. */
class ExcelError extends Error {
  constructor(message) {
    super(message);
    this.name = 'ExcelError';
    this.isFriendly = true;
  }
}

/** Read a workbook from disk, returning the SheetJS workbook + sheet names. */
function readWorkbook(filePath) {
  let workbook;
  try {
    workbook = xlsx.readFile(filePath, { cellDates: false, raw: true });
  } catch (err) {
    throw new ExcelError('The file could not be read. Make sure it is a valid .xlsx or .xls file.');
  }
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
    throw new ExcelError('The workbook does not contain any sheets.');
  }
  return { workbook, sheetNames: workbook.SheetNames };
}

/**
 * Turn a single sheet into a structured table:
 *   { headers, dataRows, headerRowIndex, totalRows }
 * Finds the header row even when the first row is a title.
 */
function extractSheet(workbook, sheetName) {
  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    throw new ExcelError(`Sheet "${sheetName}" was not found in the workbook.`);
  }

  // Read as a raw 2D array so we can locate the header ourselves.
  const aoa = xlsx.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: '',
    blankrows: false,
    raw: true,
  });

  if (!aoa || aoa.length === 0) {
    throw new ExcelError(`Sheet "${sheetName}" is empty.`);
  }

  const headerRowIndex = findHeaderRow(aoa);
  const rawHeaders = aoa[headerRowIndex] || [];
  const headers = normalizeHeaders(rawHeaders);
  const dataRows = aoa
    .slice(headerRowIndex + 1)
    .filter((row) => row.some((cell) => cell !== '' && cell !== null && cell !== undefined));

  if (dataRows.length === 0) {
    throw new ExcelError(`Sheet "${sheetName}" has headers but no data rows.`);
  }

  return {
    headers,
    dataRows,
    headerRowIndex,
    totalRows: dataRows.length,
  };
}

/**
 * Scan the first rows and pick the one that best matches known attendance
 * fields. Falls back to row 0. Handles exports where row 0 is a report title.
 */
function findHeaderRow(aoa) {
  const limit = Math.min(MAX_HEADER_SCAN_ROWS, aoa.length);
  let bestIndex = 0;
  let bestScore = -1;
  for (let i = 0; i < limit; i++) {
    const score = scoreHeaderRow(aoa[i]);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestIndex;
}

/** Ensure headers are non-empty strings; auto-name blanks as Column N. */
function normalizeHeaders(rawHeaders) {
  return rawHeaders.map((h, i) => {
    const str = h === null || h === undefined ? '' : String(h).trim();
    return str || `Column ${i + 1}`;
  });
}

/**
 * Build a summary used by the UI after upload: sheet list, chosen sheet, row
 * count, headers, auto-detected mapping and whether manual mapping is required.
 */
function buildSheetSummary(workbook, sheetName) {
  const { headers, dataRows, headerRowIndex, totalRows } = extractSheet(workbook, sheetName);
  const detection = matchColumns(headers);
  return {
    sheetName,
    headers,
    headerRowIndex,
    totalRows,
    detection,
    preview: dataRows.slice(0, 5).map((row) => rowToObject(headers, row)),
  };
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    obj[h] = row[i] === undefined ? '' : row[i];
  });
  return obj;
}

/**
 * Convert a sheet's data rows into normalized attendance records using a column
 * mapping ({ field: columnIndex }). Returns:
 *   { records, invalidRows }
 * where invalidRows captures rows whose date/time could not be parsed.
 */
function toRecords(dataRows, mapping) {
  const idx = (field) => (mapping[field] !== undefined && mapping[field] !== null ? mapping[field] : -1);
  const records = [];

  dataRows.forEach((row, i) => {
    const get = (field) => {
      const j = idx(field);
      return j >= 0 ? row[j] : undefined;
    };
    records.push({
      rowNumber: i + 1,
      employeeId: cleanId(get('employeeId')),
      employeeName: toText(get('employeeName')),
      rawDate: get('date'),
      rawTime: get('time'),
      punchType: toText(get('punchType')),
      department: toText(get('department')),
    });
  });

  return records;
}

function cleanId(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function toText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

/**
 * Accept either an auto-detected mapping (field -> {index}) or a manual mapping
 * (field -> index) and return a plain { field: index } object.
 */
function resolveMapping(detectionMapping, manualMapping) {
  const result = {};
  if (manualMapping && typeof manualMapping === 'object') {
    for (const [field, index] of Object.entries(manualMapping)) {
      const n = Number(index);
      if (Number.isInteger(n) && n >= 0) result[field] = n;
    }
  }
  if (detectionMapping) {
    for (const [field, info] of Object.entries(detectionMapping)) {
      if (result[field] === undefined && info && Number.isInteger(info.index)) {
        result[field] = info.index;
      }
    }
  }
  return result;
}

module.exports = {
  ExcelError,
  readWorkbook,
  extractSheet,
  findHeaderRow,
  buildSheetSummary,
  toRecords,
  resolveMapping,
};
