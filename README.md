# Attendance & Overtime Analyzer

**English** | [العربية](README.ar.md)

> A bilingual attendance and overtime analyzer for Excel-based biometric
> attendance exports, built with Node.js, Express, and Excel reporting.

A simple but professional internal tool for HR / IT / Operations. Upload an Excel
attendance export, let the app **auto-detect the relevant columns** (English or
Arabic), choose a **calculation mode**, analyze **work hours, overtime, late and
early-leave**, review a clean **bilingual (EN/AR, RTL/LTR)** dashboard, and download
a **polished multi-sheet Excel report**.

The first screen is the tool itself — no marketing, no sign-up. Files are processed
**locally** and never leave the machine.

---

## Screenshots

> Screenshots are not committed to keep the repo light. To capture your own, run the
> app (`npm start`), open <http://localhost:3000>, and add images under a `docs/`
> folder, then reference them here. Suggested shots:
>
> - `docs/upload.png` — the upload screen (English, LTR)
> - `docs/dashboard.png` — the results dashboard with the daily table
> - `docs/arabic-rtl.png` — the dashboard in Arabic (RTL)
> - `docs/report.png` — the exported Excel report

---

## Features

- **Drag-and-drop upload** of `.xlsx` / `.xls` files (max 15 MB).
- **Smart column detection** with English + Arabic aliases and fuzzy matching.
- **Header-row detection** — handles vendor exports (e.g. ZKTeco) where the first
  row is a title and the real header is further down.
- **Manual column mapping** screen when a required column can't be detected.
- **Two calculation modes** chosen before analysis (see below).
- **Break tracking** (punch-status mode): computes real break time from `Break Out`
  / `Break In` punches and deducts it from work hours.
- **Absence detection**: working days (non-weekend) with no record are flagged absent.
- **Bilingual UI (English / العربية)** with full **RTL/LTR** support, remembered
  across refreshes.
- **Multi-sheet workbooks** — choose which sheet to analyze.
- **Configurable settings**: standard work hours, official start/end time, break
  duration, overtime rounding, minimum overtime threshold, weekend exclusion.
- **Incomplete-day detection**: missing in, missing out, single punch, invalid
  date/time.
- **Report-period detection** that warns when a file spans more than one month.
- **Results dashboard**: totals (incl. absent days) + daily, employee-summary and
  issues tables with search, issue-type filter and sortable columns.
- **Polished Excel export**: 4 sheets, bold headers, auto filters, frozen header
  row and sensible column widths.
- **Pure, unit-tested calculation logic.**

---

## Tech stack

| Area | Technology |
|------|-----------|
| Runtime | **Node.js 18+** |
| Web server | **Express** |
| Views | **EJS** (server-rendered shell) + vanilla JS (no build step) |
| File upload | **Multer** |
| Excel reading | **SheetJS (`xlsx`)** — tolerant parsing of messy vendor exports |
| Excel writing | **ExcelJS** — formatted reports (bold headers, filters, frozen rows) |
| i18n | Custom lightweight EN/AR dictionary with RTL/LTR switching |
| Tests | **`node:test`** (built-in, zero extra deps) |
| Dev | **nodemon** |

No front-end framework or bundler — the UI is intentionally zero-build so it runs
anywhere with a single `npm start`.

---

## Calculation modes

You choose how check-in / check-out are determined **before** running the analysis.
The app never decides automatically just because a status column exists.

### 1. First & last punch of the day (default)

- Status labels are **ignored**.
- The **earliest** punch of the day = check-in; the **latest** = check-out.
- Labels like `Check In`, `Check Out`, `Break In`, `Break Out` are not consulted.
- Best when the punch-status column is missing, unreliable, or inconsistent.

### 2. Use punch status / type

- Uses **recognized statuses**: `Check In` / `Check Out`, `Sign/Clock/Punch In·Out`,
  and Arabic `دخول` / `خروج` / `حضور` / `انصراف`.
- The earliest recognized check-in and the latest recognized check-out are used.
- **No silent fallback**: if a recognized check-out is missing, the day is reported
  as an issue (`Missing check-out`) instead of guessing from the latest punch.
- Days with no recognized in/out are flagged (`No recognized check-in/out`).
- **Break time is measured** from `Break Out` → `Break In` pairs (also `Lunch`,
  `Meal`, Arabic `استراحة`) and subtracted from work hours. Multiple breaks in a day
  are summed. If a day has no break punches, the manual *break duration* setting is
  used instead.
- This mode is only offered when the file actually has a punch-status column.

### Punch-status behavior summary

| Situation | First/Last mode | Punch-status mode |
|---|---|---|
| `Check In` + `Check Out` | earliest / latest | uses the recognized pair |
| `Check In` only (no out) | last punch becomes out → **complete** | **issue: Missing check-out** |
| `Break In` / `Break Out` | treated as ordinary punches | **break duration measured & deducted** |
| No status column | works normally | mode is disabled in the UI |

---

## How attendance calculation works

1. Rows are grouped by **employee + normalized date** (`employeeId__YYYY-MM-DD`). The
   date key is stable regardless of time, timezone, locale or Excel serial quirks, so
   **multiple punches on the same date collapse into one attendance day**.
2. In/out are chosen per the selected **calculation mode** (above).
3. **Span hours** = Last Out − First In. **Work hours** = span − break.
4. **Overtime** = work hours above the configured standard day (default **10h**). On
   excluded weekend days, *all* work hours count as overtime.
5. Overtime is **rounded** (none / nearest 15 / nearest 30 min) and dropped if below
   the **minimum overtime threshold**.
6. **Late minutes** = First In − official start; **early-leave** = official end − Last
   Out (each enabled only when its time is set).
7. **Incomplete days** and **invalid date/time** rows are recorded as issues and
   excluded from hour totals.
8. **Attendance Days** = count of *unique* dates (`new Set(dates).size`), never a punch
   count. The app detects the **report month** and warns about out-of-month dates.
9. **Absences**: within the date range covered by the file, any **working day** (not a
   configured weekend) on which an employee has no record is counted as an absence and
   listed in the Issues table. Configure **weekend days** for accurate absence numbers —
   otherwise every non-attended day in the range (weekends included) is counted absent.

All calculation functions in `services/attendanceService.js` and `utils/` are pure
(no I/O), so they are easy to test and reason about.

### Break calculation

In **punch-status mode**, the app measures real break time from `Break Out` → `Break
In` pairs (also `Lunch`, `Meal`, Arabic `استراحة`):

- Each `Break Out` is paired with the earliest following `Break In`; durations are
  **summed** across all breaks in the day.
- The total break is **subtracted from work hours** (`work = span − break`), so overtime
  reflects actual worked time.
- If a day has **no** break punches, the manual *break duration* setting is used instead.
- In **first/last mode**, status labels are ignored, so break punches don't affect the
  span — only the manual break setting applies.

Break minutes appear per day (Daily Attendance) and as a per-employee total (Employee
Summary), in the UI and the Excel report.

### Absence calculation

Within the date range covered by the file, any **working day** (a day whose weekday is
**not** in the configured weekend set) on which an employee has **no record** is counted
as an **absence**:

- Each absence becomes an `Absent` row in the Issues table.
- Each employee gets an **Absent Days** count; the dashboard shows a total.
- **Configure your weekend days** for accurate numbers — otherwise every non-attended
  day in the range (weekends included) is treated as an absence.

---

## Supported Excel columns and aliases

Detection is case / space / punctuation-insensitive. Required fields are marked **\***.

| Logical field      | Example aliases (EN / AR) |
|--------------------|---------------------------|
| **employeeId \***  | `empId`, `Employee ID`, `Employee No`, `ID`, `Code`, `رقم الموظف`, `كود الموظف` |
| **date \***        | `date`, `Attendance Date`, `Punch Date`, `التاريخ`, `تاريخ الحركة` |
| **time \***        | `time`, `Attendance Time`, `Punch Time`, `الوقت`, `وقت الحركة` |
| employeeName       | `name`, `Employee Name`, `First Name`, `الاسم`, `اسم الموظف` |
| punchType          | `type`, `status`, `Punch State`, `In/Out`, `Check Type`, `الحركة`, `نوع الحركة` |
| department / site  | `department`, `site`, `branch`, `location`, `area`, `القسم`, `الموقع`, `الفرع` |

If a required column isn't detected confidently, the app shows a **mapping screen**
instead of crashing.

---

## Project structure

```
app.js                     # Express bootstrap, middleware, error handler
routes/attendance.js       # Upload / sheet / analyze / export API
services/
  excelService.js          # Read workbook, find header row, build records (SheetJS)
  attendanceService.js     # Pure calculation: modes, hours, overtime, issues
  exportService.js          # Polished Excel report (ExcelJS)
utils/
  dateUtils.js             # Date/time parsing & formatting helpers (pure)
  columnMatcher.js         # Column alias detection (EN/AR), punch-type heuristic
views/index.ejs            # Single-page UI shell
public/
  styles.css               # Styling incl. RTL rules
  i18n.js                  # Bilingual dictionary (EN/AR)
  app.js                   # Client logic (upload, mapping, modes, i18n, dashboard)
test/                      # node:test unit tests
samples/                   # Anonymized example files (no real data)
uploads/ exports/ logs/    # Runtime folders (git-ignored)
```

**Why two Excel libraries?** `xlsx` (SheetJS) reads messy real-world files robustly;
`exceljs` writes the report with formatting (bold headers, auto filters, frozen panes).

---

## Installation

Requires **Node.js 18+** (tested on Node 20).

```bash
npm install
```

## Run commands

```bash
npm start      # production: node app.js
npm run dev    # development: nodemon (auto-reload)
npm test       # run the unit tests (node:test)
```

Then open **http://localhost:3000** (set `PORT` to change it).

---

## How to use

1. **Choose your language** (English / العربية) from the header — it's remembered.
2. **Upload** an Excel file (drag-and-drop or browse).
3. Review the **workbook & columns** panel; pick a different sheet or fix the
   **column mapping** if needed.
4. Pick a **calculation mode** and set your **options** (standard hours, start/end
   times, break, rounding, weekend days).
5. Click **Analyze attendance** to see the dashboard. Use search, the issue filter
   and sortable headers to explore.
6. Click **Download Excel report**, or **Upload another file** to start over.

Two anonymized files are provided in `samples/` to try immediately:
`sample-simple.xlsx` (time/date/empId) and `sample-zkteco.xlsx` (ZKTeco-style with
Check In/Out).

---

## Exported report structure

A single `.xlsx` with four sheets:

1. **Daily Attendance** — Employee ID, Name, Date, First In, Last Out, Work Hours,
   Overtime Hours, **Break Minutes**, Late Minutes, Early Leave Minutes, Status/Notes.
2. **Employee Summary** — ID, Name, **Attendance Days** (unique dates), Total Work
   Hours, Total Overtime Hours, **Total Break Minutes**, Total Late Minutes, Total
   Early Leave Minutes, Incomplete Days, **Absent Days**.
3. **Issues** — ID, Name, Date, Issue Type, Details.
4. **Settings Used** — the **calculation mode**, all settings, the detected report
   period, report totals and a timestamp.

Every data sheet has bold headers, an auto filter and a frozen header row.

---

## Tests

```bash
npm test
```

Covers column detection (EN/AR + header-row finding), date/time normalization and
rounding, overtime calculation, **both calculation modes**, incomplete-day and
invalid-data handling, unique-date counting, and report-period detection.

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| *"Unsupported file type"* | Upload `.xlsx` or `.xls` only. |
| *"Required columns … are not mapped"* | Open **Column mapping** and pick Employee ID, Date and Time. |
| Many *"Missing check-out"* issues | You're in punch-status mode but the file lacks clean Check Out punches — switch to **first/last** mode. |
| Wrong in/out times | Choose the calculation mode that matches your data. |
| Day counts look too high | The file spans more than one month — see the report-period warning. |
| *"Your uploaded file has expired"* | Uploads are kept for 1 hour; just upload again. |
| Corrupted install errors | `rm -rf node_modules package-lock.json && npm install`. |

---

## Portfolio value

This project demonstrates:

- **Excel data processing** — robust reading of messy biometric exports (SheetJS) and
  formatted report generation (ExcelJS).
- **Attendance analytics & business logic** — configurable overtime, late/early-leave,
  incomplete-day detection, and unique-day counting with report-period awareness.
- **Bilingual product thinking** — full English/Arabic UI with correct **RTL/LTR**
  layout and localized dynamic content (tables, badges, messages).
- **Clean architecture** — pure, unit-tested calculation core separated from I/O,
  Express routing, and the view layer.
- **Practical HR/Operations automation** — a real internal-tool workflow that turns a
  raw attendance export into a shareable report in a few clicks.

---

## Assumptions & limitations

- One in/out pair per employee-day (first in, last out / recognized pair). Multiple
  in/out sessions in a day are summarized as a single span, not per-session.
- Overnight shifts crossing midnight are treated per calendar date.
- Uploaded files are tracked in-process (no database); restarting clears them.
- No authentication — intended for trusted internal / local use.

## Future improvement ideas

- Multi-session (split-shift) handling beyond a single in/out span per day.
- Overnight-shift linking across midnight.
- Per-employee or per-department schedules instead of one global schedule.
- Persisted history and saved setting presets.
- Authentication and role-based access for shared deployments.

---

## License

[MIT](LICENSE).
