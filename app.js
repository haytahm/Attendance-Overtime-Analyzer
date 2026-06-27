'use strict';

const express = require('express');
const path = require('path');
const fs = require('fs');

const attendanceRoutes = require('./routes/attendance');

const app = express();
const PORT = process.env.PORT || 3000;

// Ensure runtime folders exist (they are git-ignored but must exist at runtime).
for (const dir of ['uploads', 'exports', 'logs']) {
  fs.mkdirSync(path.join(__dirname, dir), { recursive: true });
}

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

app.use(express.json({ limit: '1mb' }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Single-page tool.
app.get('/', (req, res) => res.render('index'));

// Lightweight health check.
app.get('/health', (req, res) => res.json({ status: 'ok' }));

// API.
app.use('/api', attendanceRoutes);

// 404 for unknown API routes.
app.use('/api', (req, res) => res.status(404).json({ error: 'Not found.' }));

// Central error handler — never leak stack traces to the client.
// eslint-disable-next-line no-unused-vars
app.use((err, req, res, next) => {
  console.error('[error]', err && err.message ? err.message : err);
  if (res.headersSent) return next(err);
  const friendly = err && err.isFriendly ? err.message : 'Unexpected server error.';
  res.status(err && err.status ? err.status : 500).json({ error: friendly });
});

app.listen(PORT, () => {
  console.log(`Attendance & Overtime Analyzer running at http://localhost:${PORT}`);
});

module.exports = app;
