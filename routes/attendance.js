'use strict';

const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const excelService = require('../services/excelService');
const attendanceService = require('../services/attendanceService');
const exportService = require('../services/exportService');

const router = express.Router();

const UPLOAD_DIR = path.join(__dirname, '..', 'uploads');
const ALLOWED_EXT = ['.xlsx', '.xls'];
const MAX_FILE_BYTES = 15 * 1024 * 1024; // 15 MB

// In-memory registry of uploaded files for this process. Keeps the upload flow
// stateless from the client's perspective without needing a session store.
// fileId -> { path, originalName, expires }
const fileRegistry = new Map();
const FILE_TTL_MS = 60 * 60 * 1000; // 1 hour

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    cb(null, `${crypto.randomUUID()}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: MAX_FILE_BYTES },
  fileFilter: (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (!ALLOWED_EXT.includes(ext)) {
      return cb(new FriendlyError('Unsupported file type. Please upload an .xlsx or .xls file.'));
    }
    cb(null, true);
  },
}).single('attendanceFile');

/** Friendly error whose message is safe to show the user. */
class FriendlyError extends Error {
  constructor(message) {
    super(message);
    this.isFriendly = true;
  }
}

function registerFile(filePath, originalName) {
  const id = crypto.randomUUID();
  fileRegistry.set(id, { path: filePath, originalName, expires: Date.now() + FILE_TTL_MS });
  return id;
}

function getFile(id) {
  const entry = fileRegistry.get(id);
  if (!entry) return null;
  if (entry.expires < Date.now()) {
    safeUnlink(entry.path);
    fileRegistry.delete(id);
    return null;
  }
  return entry;
}

function safeUnlink(p) {
  fs.promises.unlink(p).catch(() => {});
}

// Periodically clean up expired uploads.
setInterval(() => {
  const now = Date.now();
  for (const [id, entry] of fileRegistry.entries()) {
    if (entry.expires < now) {
      safeUnlink(entry.path);
      fileRegistry.delete(id);
    }
  }
}, 10 * 60 * 1000).unref();

/**
 * POST /api/upload
 * Accepts the file, reads the workbook, returns sheet list + summary of the
 * first sheet (headers, detected mapping, preview, row count).
 */
router.post('/upload', (req, res) => {
  upload(req, res, (err) => {
    if (err) {
      const msg = err instanceof FriendlyError ? err.message
        : err.code === 'LIMIT_FILE_SIZE' ? 'File is too large (max 15 MB).'
        : 'Could not upload the file. Please try again.';
      return res.status(400).json({ error: msg });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'No file received. Please choose an Excel file.' });
    }

    try {
      const { workbook, sheetNames } = excelService.readWorkbook(req.file.path);
      const fileId = registerFile(req.file.path, req.file.originalname);
      const summary = excelService.buildSheetSummary(workbook, sheetNames[0]);
      res.json({
        fileId,
        fileName: req.file.originalname,
        sheetNames,
        selectedSheet: sheetNames[0],
        ...summary,
      });
    } catch (e) {
      safeUnlink(req.file.path);
      res.status(400).json({ error: friendlyMessage(e) });
    }
  });
});

/**
 * POST /api/sheet  { fileId, sheetName }
 * Re-reads a different sheet of an already-uploaded workbook.
 */
router.post('/sheet', (req, res) => {
  try {
    const entry = requireFile(req.body.fileId);
    const { workbook } = excelService.readWorkbook(entry.path);
    const summary = excelService.buildSheetSummary(workbook, req.body.sheetName);
    res.json({ fileId: req.body.fileId, selectedSheet: req.body.sheetName, ...summary });
  } catch (e) {
    res.status(400).json({ error: friendlyMessage(e) });
  }
});

/**
 * POST /api/analyze  { fileId, sheetName, mapping?, settings? }
 * Runs the full analysis and returns the dashboard payload as JSON.
 */
router.post('/analyze', (req, res) => {
  try {
    const analysis = runAnalysis(req.body);
    res.json(analysis);
  } catch (e) {
    res.status(400).json({ error: friendlyMessage(e) });
  }
});

/**
 * POST /api/export  { fileId, sheetName, mapping?, settings? }
 * Re-runs analysis and streams a polished .xlsx report.
 */
router.post('/export', async (req, res) => {
  try {
    const analysis = runAnalysis(req.body);
    const buffer = await exportService.buildReport(analysis);
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="attendance-report-${stamp}.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (e) {
    res.status(400).json({ error: friendlyMessage(e) });
  }
});

/** Shared: read sheet -> resolve mapping -> records -> analyze. */
function runAnalysis(body) {
  const entry = requireFile(body.fileId);
  const sheetName = body.sheetName;
  const { workbook } = excelService.readWorkbook(entry.path);
  const { headers, dataRows } = excelService.extractSheet(workbook, sheetName);
  const detection = require('../utils/columnMatcher').matchColumns(headers);

  const mapping = excelService.resolveMapping(detection.mapping, body.mapping);
  if (mapping.employeeId === undefined || mapping.date === undefined || mapping.time === undefined) {
    throw new FriendlyError(
      'Required columns (Employee ID, Date, Time) are not mapped. Please map them and try again.'
    );
  }

  const records = excelService.toRecords(dataRows, mapping);
  return attendanceService.analyze(records, body.settings || {});
}

function requireFile(fileId) {
  const entry = getFile(fileId);
  if (!entry) {
    throw new FriendlyError('Your uploaded file has expired or was not found. Please upload it again.');
  }
  return entry;
}

function friendlyMessage(e) {
  // Log every handled error server-side; only friendly text reaches the client.
  if (e && (e.isFriendly || e.name === 'ExcelError')) {
    console.warn('[attendance] handled:', e.message);
    return e.message;
  }
  console.error('[attendance] unexpected error:', e);
  return 'Something went wrong while processing the file. Please check the file and try again.';
}

module.exports = router;
