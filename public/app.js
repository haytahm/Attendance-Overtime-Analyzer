'use strict';

/**
 * Client logic for the Attendance & Overtime Analyzer.
 * Flow: upload -> review workbook/mapping -> settings -> analyze -> dashboard -> export.
 */

const MAPPING_FIELDS = ['employeeId', 'date', 'time', 'employeeName', 'punchType', 'department'];
const FIELD_KEY = {
  employeeId: 'field.employeeId', date: 'field.date', time: 'field.time',
  employeeName: 'field.employeeName', punchType: 'field.punchType', department: 'field.department',
};
const REQUIRED_FIELDS = ['employeeId', 'date', 'time'];
// Weekday i18n keys (0 = Sunday). Names rendered via the dictionary's day list.
const WEEKDAY_KEYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const WEEKDAY_NAMES = {
  en: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
  ar: ['الأحد', 'الإثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'],
};

// App state for the current upload flow.
const state = {
  lang: 'en',
  fileId: null,
  fileName: null,
  sheetNames: [],
  selectedSheet: null,
  headers: [],
  detection: null,
  analysis: null,
  activeTab: 'daily',
  timeUnit: 'hours',
  search: '',
  issueFilter: '',
  sort: { key: null, dir: 1 },
};

const $ = (id) => document.getElementById(id);

/* ---------- i18n ---------- */

/** Translate a key with optional {var} interpolation. Falls back to English, then the key. */
function t(key, vars) {
  const dict = window.I18N[state.lang] || window.I18N.en;
  let s = dict[key];
  if (s === undefined) s = (window.I18N.en[key] !== undefined ? window.I18N.en[key] : key);
  if (vars) s = s.replace(/\{(\w+)\}/g, (m, k) => (vars[k] !== undefined ? vars[k] : m));
  return s;
}

/** Apply a language: set <html> lang/dir, translate static nodes, re-render dynamic UI. */
function applyLanguage(lang) {
  if (!window.I18N[lang]) lang = 'en';
  state.lang = lang;
  try { localStorage.setItem('analyzerLang', lang); } catch (e) { /* ignore */ }

  const dir = window.I18N[lang]._dir;
  document.documentElement.lang = lang;
  document.documentElement.dir = dir;

  // Static text nodes.
  document.querySelectorAll('[data-i18n]').forEach((el) => {
    el.textContent = t(el.getAttribute('data-i18n'));
  });
  document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
    el.placeholder = t(el.getAttribute('data-i18n-placeholder'));
  });
  document.title = t('app.title');

  // Active-language button.
  document.querySelectorAll('.lang-btn').forEach((b) =>
    b.classList.toggle('is-active', b.dataset.lang === lang)
  );

  // Rebuild language-dependent dynamic parts.
  rebuildWeekendChips();
  if (state.detection) buildMappingForm();
  if (state.analysis) renderDashboard();
}

function setupLanguageSwitch() {
  document.querySelectorAll('.lang-btn').forEach((b) =>
    b.addEventListener('click', () => applyLanguage(b.dataset.lang))
  );
  let saved = 'en';
  try { saved = localStorage.getItem('analyzerLang') || 'en'; } catch (e) { /* ignore */ }
  applyLanguage(saved);
}

document.addEventListener('DOMContentLoaded', () => {
  // Belt-and-suspenders: guarantee the overlay starts hidden regardless of CSS.
  hideOverlay();
  console.info('[analyzer] ready — overlay hidden, awaiting user action.');
  setupLanguageSwitch();
  setupDropzone();
  setupWeekendChips();
  setupTabs();
  setupToolbar();
  $('analyzeBtn').addEventListener('click', runAnalyze);
  $('exportBtn').addEventListener('click', runExport);
  $('resetBtn').addEventListener('click', resetAll);
  $('sheetSelect').addEventListener('change', onSheetChange);
  document.querySelectorAll('input[name="calcMode"]').forEach((r) =>
    r.addEventListener('change', () => { /* selection captured at analyze time */ })
  );
});

/* ---------- Upload ---------- */

function setupDropzone() {
  const dz = $('dropzone');
  const input = $('fileInput');
  dz.addEventListener('click', () => input.click());
  dz.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });
  input.addEventListener('change', () => { if (input.files[0]) uploadFile(input.files[0]); });

  ['dragenter', 'dragover'].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('is-dragover'); })
  );
  ['dragleave', 'drop'].forEach((ev) =>
    dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('is-dragover'); })
  );
  dz.addEventListener('drop', (e) => {
    const file = e.dataTransfer.files[0];
    if (file) uploadFile(file);
  });
}

async function uploadFile(file) {
  hideBanner();
  const ext = file.name.toLowerCase().slice(file.name.lastIndexOf('.'));
  if (!['.xlsx', '.xls'].includes(ext)) {
    return showBanner(t('err.fileType'), 'error');
  }
  setUploadStatus(t('upload.uploading', { name: `<strong>${escapeHtml(file.name)}</strong>` }));

  const form = new FormData();
  form.append('attendanceFile', file);

  const data = await apiSend('/api/upload', { method: 'POST', body: form }, t('overlay.reading'));
  if (!data) return;

  state.fileId = data.fileId;
  state.fileName = data.fileName;
  state.sheetNames = data.sheetNames;
  applySheetSummary(data);
  setUploadStatus(t('upload.loaded', { name: `<strong>${escapeHtml(data.fileName)}</strong>`, rows: data.totalRows }));
  $('resetBtn').hidden = false;
  show('workbookSection');
  show('settingsSection');
  hide('resultsSection');
  $('workbookSection').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

async function onSheetChange() {
  const sheetName = $('sheetSelect').value;
  const data = await apiSend('/api/sheet', jsonReq({ fileId: state.fileId, sheetName }), t('overlay.readingSheet'));
  if (!data) return;
  applySheetSummary(data);
}

function applySheetSummary(data) {
  state.selectedSheet = data.selectedSheet;
  state.headers = data.headers;
  state.detection = data.detection;

  $('metaFile').textContent = state.fileName;
  $('metaRows').textContent = data.totalRows;

  const sel = $('sheetSelect');
  sel.innerHTML = '';
  state.sheetNames.forEach((name) => {
    const opt = document.createElement('option');
    opt.value = name; opt.textContent = name;
    if (name === data.selectedSheet) opt.selected = true;
    sel.appendChild(opt);
  });

  const detectedFields = data.detection.detected
    .map((f) => t(FIELD_KEY[f]).replace(' *', '').replace(' *', ''))
    .join('، ');
  $('metaDetected').textContent = detectedFields || t('meta.none');

  $('mappingNotice').hidden = !data.detection.needsManualMapping;
  $('mappingDetails').open = data.detection.needsManualMapping;
  buildMappingForm();
  updateModeAvailability();
}

/** Enable "use punch status" only when the file actually has a punch-type column. */
function updateModeAvailability() {
  const hasPunch = !!(state.detection && state.detection.mapping && state.detection.mapping.punchType);
  const punchRadio = document.querySelector('input[name="calcMode"][value="punchType"]');
  const hint = $('modePunchHint');
  if (punchRadio) {
    punchRadio.disabled = !hasPunch;
    if (!hasPunch && punchRadio.checked) {
      document.querySelector('input[name="calcMode"][value="firstLast"]').checked = true;
    }
  }
  if (hint) hint.hidden = hasPunch;
}

function buildMappingForm() {
  const container = $('mappingForm');
  if (!state.detection) return;
  container.innerHTML = '';
  MAPPING_FIELDS.forEach((field) => {
    const detected = state.detection.mapping[field];
    const wrap = document.createElement('label');
    wrap.className = 'field';

    const label = document.createElement('span');
    label.className = 'field__label';
    label.textContent = t(FIELD_KEY[field]);

    const select = document.createElement('select');
    select.className = 'input';
    select.dataset.field = field;

    const noneOpt = document.createElement('option');
    noneOpt.value = '';
    noneOpt.textContent = REQUIRED_FIELDS.includes(field) ? t('mapping.selectColumn') : t('mapping.none');
    select.appendChild(noneOpt);

    state.headers.forEach((header, index) => {
      const opt = document.createElement('option');
      opt.value = String(index);
      opt.textContent = header;
      if (detected && detected.index === index) opt.selected = true;
      select.appendChild(opt);
    });

    wrap.append(label, select);
    container.appendChild(wrap);
  });
}

function collectMapping() {
  const mapping = {};
  document.querySelectorAll('#mappingForm select').forEach((sel) => {
    if (sel.value !== '') mapping[sel.dataset.field] = Number(sel.value);
  });
  return mapping;
}

/* ---------- Settings ---------- */

function setupWeekendChips() {
  rebuildWeekendChips();
}

/** (Re)build the weekend-day chips in the active language, preserving checks. */
function rebuildWeekendChips() {
  const container = $('weekendDays');
  if (!container) return;
  const checked = new Set(
    Array.from(container.querySelectorAll('input:checked')).map((c) => c.value)
  );
  container.innerHTML = '';
  (WEEKDAY_NAMES[state.lang] || WEEKDAY_NAMES.en).forEach((name, idx) => {
    const chip = document.createElement('label');
    chip.className = 'weekend-chip';
    const isChecked = checked.has(String(idx)) ? 'checked' : '';
    chip.innerHTML = `<input type="checkbox" value="${idx}" ${isChecked}/> ${escapeHtml(name)}`;
    container.appendChild(chip);
  });
}

function collectSettings() {
  const weekendDays = Array.from(document.querySelectorAll('#weekendDays input:checked'))
    .map((c) => Number(c.value));
  const calcModeEl = document.querySelector('input[name="calcMode"]:checked');
  return {
    calculationMode: calcModeEl ? calcModeEl.value : 'firstLast',
    standardHours: $('setStandardHours').value,
    startTime: $('setStartTime').value,
    endTime: $('setEndTime').value,
    breakMinutes: $('setBreakMinutes').value,
    maxBreakMinutes: $('setMaxBreakMinutes').value,
    overtimeRounding: $('setOvertimeRounding').value,
    minOvertimeMinutes: $('setMinOvertime').value,
    nightShiftCutoff: $('setNightShiftCutoff').value,
    dateFormat: $('setDateFormat').value,
    periodStart: $('setPeriodStart').value,
    periodEnd: $('setPeriodEnd').value,
    weekendDays,
  };
}

/* ---------- Analyze ---------- */

async function runAnalyze() {
  hideBanner();
  const mapping = collectMapping();
  for (const f of REQUIRED_FIELDS) {
    if (mapping[f] === undefined) {
      $('mappingDetails').open = true;
      return showBanner(t('err.mapRequired', { field: t(FIELD_KEY[f]).replace(' *', '') }), 'error');
    }
  }

  const payload = {
    fileId: state.fileId,
    sheetName: state.selectedSheet,
    mapping,
    settings: collectSettings(),
  };
  const data = await apiSend('/api/analyze', jsonReq(payload), t('overlay.analyzing'));
  if (!data) return;

  state.analysis = data;
  state.sort = { key: null, dir: 1 };
  renderDashboard();
  show('resultsSection');
  $('resultsSection').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ---------- Dashboard rendering ---------- */

function renderDashboard() {
  const totals = state.analysis.totals;
  $('statEmployees').textContent = totals.totalEmployees;
  $('statWork').textContent = `${totals.totalWorkHours}`;
  $('statOvertime').textContent = `${totals.totalOvertimeHours}`;
  $('statIncomplete').textContent = totals.incompleteDays;
  $('statAbsent').textContent = totals.absentDays || 0;
  $('statLate').textContent = totals.lateRecords;
  $('statEarly').textContent = totals.earlyLeaveRecords;
  $('statBreakExceeded').textContent = (state.analysis.issues || []).filter((i) => i.issueKey === 'BREAK_EXCEEDED').length;

  renderPeriodNote(totals);
  populateIssueFilter();
  renderTable();
}

/** Show the detected report month and warn when the file spans multiple months. */
function renderPeriodNote(totals) {
  const el = $('periodNote');
  if (!el) return;
  if (totals.hasMultipleMonths || totals.outOfPeriodDays > 0) {
    el.className = 'notice notice--warn';
    el.textContent = t('period.warn', {
      months: (totals.months || []).join('، '),
      out: totals.outOfPeriodDays,
      month: totals.reportMonth,
      days: totals.daysInMonth,
    });
  } else {
    el.className = 'notice notice--info';
    el.textContent = t('period.note', { month: totals.reportMonth, days: totals.daysInMonth });
  }
  el.hidden = false;
}

function populateIssueFilter() {
  const sel = $('issueFilter');
  // Use stable issueKey for value, localized label for display.
  const keys = Array.from(new Set(state.analysis.issues.map((i) => i.issueKey)));
  sel.innerHTML = `<option value="">${escapeHtml(t('filter.allIssues'))}</option>`;
  keys.forEach((key) => {
    const opt = document.createElement('option');
    opt.value = key; opt.textContent = t('issue.' + key);
    sel.appendChild(opt);
  });
}

const COLUMNS = {
  daily: [
    { key: 'employeeId', labelKey: 'col.employeeId' },
    { key: 'employeeName', labelKey: 'col.name' },
    { key: 'date', labelKey: 'col.date' },
    { key: 'firstIn', labelKey: 'col.firstIn' },
    { key: 'lastOut', labelKey: 'col.lastOut' },
    { key: 'workHours', labelKey: 'col.workHrs', num: true },
    { key: 'overtimeHours', labelKey: 'col.otHrs', num: true },
    { key: 'breakMinutes', labelKey: 'col.breakM', num: true, duration: true, punchDetails: true },
    { key: 'breakExcessMinutes', labelKey: 'col.breakExcess', num: true, duration: true },
    { key: 'lateMinutes', labelKey: 'col.lateM', num: true, duration: true },
    { key: 'earlyLeaveMinutes', labelKey: 'col.earlyM', num: true, duration: true },
    { key: 'status', labelKey: 'col.status', dailyStatus: true },
  ],
  employees: [
    { key: 'employeeId', labelKey: 'col.employeeId' },
    { key: 'employeeName', labelKey: 'col.name' },
    { key: 'daysCount', labelKey: 'col.attendanceDays', num: true, daysBadge: true },
    { key: 'totalWorkHours', labelKey: 'col.workHrs', num: true },
    { key: 'totalOvertimeHours', labelKey: 'col.otHrs', num: true },
    { key: 'totalBreakMinutes', labelKey: 'col.breakM', num: true, duration: true },
    { key: 'totalBreakExcessMinutes', labelKey: 'col.breakExcess', num: true, duration: true },
    { key: 'totalLateMinutes', labelKey: 'col.lateM', num: true, duration: true },
    { key: 'totalEarlyLeaveMinutes', labelKey: 'col.earlyM', num: true, duration: true },
    { key: 'incompleteDays', labelKey: 'col.incomplete', num: true },
    { key: 'absentDays', labelKey: 'col.absentDays', num: true },
  ],
  issues: [
    { key: 'employeeId', labelKey: 'col.employeeId' },
    { key: 'employeeName', labelKey: 'col.name' },
    { key: 'date', labelKey: 'col.date' },
    { key: 'issueType', labelKey: 'col.issueType', issueBadge: true },
    { key: 'details', labelKey: 'col.details' },
  ],
};

function currentRows() {
  const a = state.analysis;
  let rows = a[state.activeTab === 'employees' ? 'employees' : state.activeTab];
  rows = rows.slice();

  // Search by employee id/name.
  const q = state.search.trim().toLowerCase();
  if (q) {
    rows = rows.filter((r) =>
      String(r.employeeId || '').toLowerCase().includes(q) ||
      String(r.employeeName || '').toLowerCase().includes(q)
    );
  }
  // Issue-type filter (issues tab only), matched by stable key.
  if (state.activeTab === 'issues' && state.issueFilter) {
    rows = rows.filter((r) => r.issueKey === state.issueFilter);
  }
  // Sort.
  if (state.sort.key) {
    const k = state.sort.key, dir = state.sort.dir;
    rows.sort((a, b) => compareValues(a[k], b[k]) * dir);
  }
  return rows;
}

function compareValues(a, b) {
  const na = Number(a), nb = Number(b);
  if (!isNaN(na) && !isNaN(nb)) return na - nb;
  return String(a ?? '').localeCompare(String(b ?? ''), undefined, { numeric: true });
}

function renderTable() {
  const cols = COLUMNS[state.activeTab];
  const head = $('tableHead');
  const body = $('tableBody');

  head.innerHTML = '<tr>' + cols.map((c) => {
    let ind = '';
    if (state.sort.key === c.key) ind = state.sort.dir === 1 ? ' ▲' : ' ▼';
    return `<th data-key="${c.key}">${escapeHtml(t(c.labelKey))}<span class="sort-ind">${ind}</span></th>`;
  }).join('') + '</tr>';

  head.querySelectorAll('th').forEach((th) =>
    th.addEventListener('click', () => toggleSort(th.dataset.key))
  );

  const rows = currentRows();
  $('emptyState').hidden = rows.length > 0;

  body.innerHTML = rows.map((r, idx) =>
    '<tr>' + cols.map((c) => `<td>${renderCell(r, c, idx)}</td>`).join('') + '</tr>'
  ).join('');

  // Click any duration span to toggle that single cell between hours and minutes.
  body.querySelectorAll('.dur-toggle').forEach((el) =>
    el.addEventListener('click', () => {
      const mins = Number(el.dataset.mins);
      const cur = el.dataset.unit;
      const next = cur === 'hours' ? 'minutes' : 'hours';
      el.dataset.unit = next;
      el.textContent = formatDuration(mins, next);
    })
  );

  // Click "Details" button to open the day's punch breakdown modal.
  body.querySelectorAll('.punch-details-btn').forEach((btn) =>
    btn.addEventListener('click', () => {
      const row = rows[Number(btn.dataset.rowIdx)];
      if (row) openPunchModal(row);
    })
  );
}

function formatDuration(minutes, unit) {
  const m = Math.round(Number(minutes) || 0);
  if (m === 0) return '0';
  if (unit === 'minutes') {
    return state.lang === 'ar' ? `${m} د` : `${m}m`;
  }
  const h = Math.floor(m / 60);
  const rem = m % 60;
  if (state.lang === 'ar') {
    if (h && rem) return `${h}س ${rem}د`;
    if (h) return `${h}س`;
    return `${rem}د`;
  }
  if (h && rem) return `${h}h ${rem}m`;
  if (h) return `${h}h`;
  return `${rem}m`;
}

function renderCell(row, col, rowIdx) {
  const val = row[col.key];
  if (col.dailyStatus) return dailyStatusBadge(row);
  if (col.issueBadge) return issueBadge(row);
  if (col.duration) {
    const mins = Number(val) || 0;
    const altUnit = state.timeUnit === 'hours' ? 'minutes' : 'hours';
    const durHtml = mins > 0
      ? `<span class="dur-toggle" data-mins="${mins}" data-unit="${state.timeUnit}" title="${escapeHtml(formatDuration(mins, altUnit))}">${escapeHtml(formatDuration(mins, state.timeUnit))}</span>`
      : '0';
    if (col.punchDetails && row.punches && row.punches.length > 0) {
      return `${durHtml}<button type="button" class="btn-xs punch-details-btn" data-row-idx="${rowIdx}">${escapeHtml(t('btn.details'))}</button>`;
    }
    return durHtml;
  }
  if (col.daysBadge && row.exceedsMonth) {
    // Impossible count for a single month → flag it (caused by out-of-month dates).
    const title = t('days.badgeTitle', {
      days: state.analysis.totals.daysInMonth,
      out: row.outOfPeriodDays,
      month: state.analysis.totals.reportMonth,
    });
    return `<span class="badge badge--warn" title="${escapeHtml(title)}">⚠ ${escapeHtml(String(val))}</span>`;
  }
  if (val === null || val === undefined || val === '') return '—';
  return escapeHtml(String(val));
}

/** Localized status badge for a Daily Attendance row, composed from flags. */
function dailyStatusBadge(row) {
  if (!row.complete) {
    const label = t('issue.' + row.statusKey);
    const cls = /MISSING|INVALID|SINGLE/.test(row.statusKey) ? 'badge--danger' : 'badge--warn';
    return `<span class="badge ${cls}">${escapeHtml(label)}</span>`;
  }
  const notes = [];
  if (row.isWeekend) notes.push(t('status.weekend'));
  if (row.lateMinutes > 0) notes.push(t('status.late', { n: row.lateMinutes }));
  if (row.earlyLeaveMinutes > 0) notes.push(t('status.early', { n: row.earlyLeaveMinutes }));
  if (row.breakMinutes > 0) {
    if (row.breakExceeded) notes.push(t('status.breakExceeded', { n: row.breakMinutes }));
    else notes.push(t('status.break', { n: row.breakMinutes }));
  }
  if (!notes.length) return `<span class="badge badge--ok">${escapeHtml(t('status.OK'))}</span>`;
  return `<span class="badge badge--warn">${escapeHtml(notes.join('، '))}</span>`;
}

/** Localized issue-type badge for the Issues table. */
function issueBadge(row) {
  const label = t('issue.' + row.issueKey);
  const cls = /MISSING|INVALID|SINGLE/.test(row.issueKey) ? 'badge--danger' : 'badge--warn';
  return `<span class="badge ${cls}">${escapeHtml(label)}</span>`;
}

function toggleSort(key) {
  if (state.sort.key === key) state.sort.dir *= -1;
  else state.sort = { key, dir: 1 };
  renderTable();
}

/* ---------- Tabs / toolbar ---------- */

function setupTabs() {
  document.querySelectorAll('.tab').forEach((tab) =>
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach((el) => el.classList.remove('is-active'));
      tab.classList.add('is-active');
      state.activeTab = tab.dataset.tab;
      state.sort = { key: null, dir: 1 };
      $('issueFilter').hidden = state.activeTab !== 'issues';
      renderTable();
    })
  );
}

function setupToolbar() {
  $('searchInput').addEventListener('input', (e) => { state.search = e.target.value; renderTable(); });
  $('issueFilter').addEventListener('change', (e) => { state.issueFilter = e.target.value; renderTable(); });
  document.querySelectorAll('#unitToggle .unit-btn').forEach((btn) =>
    btn.addEventListener('click', () => {
      document.querySelectorAll('#unitToggle .unit-btn').forEach((b) => b.classList.remove('is-active'));
      btn.classList.add('is-active');
      state.timeUnit = btn.dataset.unit;
      if (state.analysis) renderTable();
    })
  );
  $('punchModalClose').addEventListener('click', closePunchModal);
  $('punchModalBackdrop').addEventListener('click', closePunchModal);
}

function openPunchModal(row) {
  const body = $('punchModalBody');
  const name = row.employeeName ? `${row.employeeName} (${row.employeeId})` : row.employeeId;
  const breakText = formatDuration(row.breakMinutes, state.timeUnit);
  const segs = row.breakSegments || [];
  const punches = row.punches || [];

  const breaksHtml = segs.length
    ? `<ul class="modal__list">${segs.map((s) =>
        `<li><span>${escapeHtml(s.from)} → ${escapeHtml(s.to)}</span><strong>${escapeHtml(formatDuration(s.minutes, state.timeUnit))}</strong></li>`
      ).join('')}</ul>`
    : `<p class="modal__meta">${escapeHtml(t('modal.noBreaks'))}</p>`;

  const punchesHtml = `<ul class="modal__list">${punches.map((p) =>
    `<li><span>${escapeHtml(p.time)}</span><span>${escapeHtml(p.punchType || '—')}</span></li>`
  ).join('')}</ul>`;

  body.innerHTML = `
    <div class="modal__meta"><strong>${escapeHtml(name)}</strong> · ${escapeHtml(row.date)} · ${escapeHtml(t('col.breakM'))}: <strong>${escapeHtml(breakText)}</strong></div>
    <div><h4 class="modal__section-title">${escapeHtml(t('modal.breaks'))}</h4>${breaksHtml}</div>
    <div><h4 class="modal__section-title">${escapeHtml(t('modal.punches'))} (${punches.length})</h4>${punchesHtml}</div>
  `;
  $('punchModal').hidden = false;
}

function closePunchModal() {
  $('punchModal').hidden = true;
}

/* ---------- Export ---------- */

async function runExport() {
  hideBanner();
  const payload = {
    fileId: state.fileId,
    sheetName: state.selectedSheet,
    mapping: collectMapping(),
    settings: collectSettings(),
    lang: state.lang,
    timeUnit: state.timeUnit,
  };
  console.info('[analyzer] export start → /api/export');
  showOverlay(t('overlay.exporting'));
  try {
    const res = await fetchWithTimeout('/api/export', jsonReq(payload));
    if (!res.ok) {
      const data = await res.json().catch(() => ({}));
      throw new Error(data.error || 'Export failed.');
    }
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `attendance-report-${new Date().toISOString().slice(0, 10)}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
    console.info('[analyzer] export done');
    showBanner(t('export.success'), 'success');
  } catch (e) {
    const msg = e.name === 'AbortError' ? t('err.timeout') : (e.message || 'Export failed.');
    console.error('[analyzer] export error:', e);
    showBanner(msg, 'error');
  } finally {
    hideOverlay();
  }
}

/* ---------- Reset ---------- */

function resetAll() {
  Object.assign(state, {
    fileId: null, fileName: null, sheetNames: [], selectedSheet: null,
    headers: [], detection: null, analysis: null, activeTab: 'daily',
    search: '', issueFilter: '', sort: { key: null, dir: 1 },
  });
  $('fileInput').value = '';
  $('searchInput').value = '';
  hide('workbookSection');
  hide('settingsSection');
  hide('resultsSection');
  $('uploadStatus').hidden = true;
  $('resetBtn').hidden = true;
  hideBanner();
  $('periodNote').hidden = true;
  document.querySelectorAll('.tab').forEach((el, i) => el.classList.toggle('is-active', i === 0));
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

/* ---------- Helpers ---------- */

function jsonReq(payload) {
  return { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) };
}

// Max time any single request may run before we give up and release the overlay,
// so the UI can never spin forever even if the backend hangs.
const REQUEST_TIMEOUT_MS = 60000;

async function fetchWithTimeout(url, options, timeoutMs = REQUEST_TIMEOUT_MS) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, { ...options, signal: controller.signal });
  } finally {
    clearTimeout(timer);
  }
}

async function apiSend(url, options, overlayText) {
  console.info(`[analyzer] request start → ${url}`);
  showOverlay(overlayText || t('overlay.working'));
  try {
    const res = await fetchWithTimeout(url, options);
    const data = await res.json().catch(() => ({ error: 'Unexpected server response.' }));
    if (!res.ok) {
      console.warn(`[analyzer] request failed (${res.status}) ← ${url}:`, data.error);
      showBanner(data.error || 'Request failed.', 'error');
      return null;
    }
    console.info(`[analyzer] request done ← ${url}`);
    return data;
  } catch (e) {
    const msg = e.name === 'AbortError' ? t('err.timeout') : t('err.network');
    console.error(`[analyzer] request error ← ${url}:`, e);
    showBanner(msg, 'error');
    return null;
  } finally {
    hideOverlay(); // always release the overlay, success or failure
  }
}

function show(id) { $(id).hidden = false; }
function hide(id) { $(id).hidden = true; }

function setUploadStatus(html) {
  const el = $('uploadStatus');
  el.innerHTML = html;
  el.hidden = false;
}

function showBanner(message, type) {
  const b = $('banner');
  b.textContent = message;
  b.className = `banner banner--${type === 'success' ? 'success' : 'error'}`;
  b.hidden = false;
  b.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}
function hideBanner() { $('banner').hidden = true; }

function showOverlay(text) {
  console.debug('[analyzer] overlay show:', text);
  $('overlayText').textContent = text;
  $('overlay').hidden = false;
}
function hideOverlay() {
  console.debug('[analyzer] overlay hide');
  $('overlay').hidden = true;
}

function escapeHtml(str) {
  return str.replace(/[&<>"']/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])
  );
}
