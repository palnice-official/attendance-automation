# Attendance Automation — RAMA Group

A browser-based tool that turns a raw biometric punch report into a payroll-ready
`.xlsx` workbook in the company's standard format. It runs entirely client-side
(no server, no data leaves the machine), built with **React + Vite** and
**ExcelJS**.

The payroll output template and the FY2026 holiday calendar are **built in**, so
each month you upload only two files: the biometric report and the list of 5-day
working employees.

---

## Prerequisites

- **Node.js 18+** and npm (https://nodejs.org)

## Quick start

```bash
npm install        # install dependencies
npm run dev        # start the dev server (prints a localhost URL)
```

Open the URL it prints (usually http://localhost:5173). To create a static
production build:

```bash
npm run build      # outputs to dist/
npm run preview    # serve the production build locally
```

The contents of `dist/` are static files and can be hosted anywhere (GitHub
Pages, Netlify, an internal server, or opened via any static file server).

## How to use the app

1. Upload the **biometric report** (Daily Attendance Detail Report) and the
   **5-day working employee list**. The payroll month auto-detects from the
   biometric report header.
2. If you are processing a month other than the embedded January baseline,
   expand **Opening leave balances** and update each person's EL / CL / CO
   (this month's closing becomes next month's opening).
3. Click **Run processing**, review the mapping table and the preview tabs
   (Daily Register, Employee Summary, Missing Punch, Comp-Off Review, Audit Log).
4. Click **Export** to download `Attendance_Final_<Month>_<Year>.xlsx`.

All absences are written as `L` (LOP) first; apply approved EL/CL/ML/OD and any
comp-off manually afterwards. Comp-off is only ever *flagged*, never auto-applied.

## Rules baked in

- Cycle: 25th of the previous month -> 24th of the payroll month.
- Sundays = `WO` for everyone; Saturdays = `WO` only for 5-day staff.
- Hours: >= 9h -> Present; 8.5-9h -> Present (flagged for review); < 8.5h -> half day.
- Holidays applied per work schedule from the built-in 2026 list.

## Project structure

```
attendance-automation/
  index.html              Vite entry
  package.json
  vite.config.js
  src/
    main.jsx              React bootstrap
    App.jsx               UI layer (upload, settings, dashboard, preview, export)
    styles.css
    data/
      template.js         built-in payroll template (.xlsx, base64)
      holidays.js         built-in FY2026 holiday list + coverage end date
      logo.js             RAMA logo (PNG data URI)
    lib/                  pure logic (framework-free, unit-testable)
      codeDictionary.js   attendance-code dictionary
      excelHelpers.js     date/time/serial normalisation helpers
      payrollCycle.js     cycle + report-period parsing
      fiveDayParser.js    5-day list parsing + schedule classification
      holidayParser.js    holiday map construction
      templateParser.js   output-template parsing
      biometricParser.js  biometric block parsing (merged-cell aware)
      matcher.js          employee matching (ID-then-name scoring)
      attendanceEngine.js per-day status, LOP/leave, balances
      runEngine.js        orchestration
      exportEngine.js     writes the final workbook (re-dates headers, fixes formulas)
```

## Maintenance notes

- **Holiday list (FY2026 completion):** the built-in calendar covers Jan–Dec 2026,
  which is correct for every payroll cycle ending through December 2026. Add the
  Jan/Feb/Mar 2027 holidays to `src/data/holidays.js` (both `6-Day` and `5-Day`
  sections) when published; cycles running past the coverage date raise a warning
  in the Audit Log.
- **Roster / opening balances:** the employee roster and the January opening
  balances are embedded in `src/data/template.js`. Update opening balances in the
  app each month; to change the roster permanently, replace the embedded template.

© Copyrights owned by RAMA Group of Companies 2026.
