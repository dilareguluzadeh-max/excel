# SDF 1 Online Exam Portal — Introduction to Data Analytics (Excel & Google Sheets)

Baku Engineering University · Year 1 · Group 116i

A browser-based exam system with anti-cheat protection. Students solve the tasks in a **built-in Excel-style spreadsheet**; the **Check** button grades the sheet (a partially correct solution earns partial points). No extra software or external grading service is needed.

**35 students · 5 variants (7 students each) · 6 tasks per variant = 100 points.**

---

## Files

| File | Purpose |
|---|---|
| `index.html` | **Student portal** — login with full name + seat number (1–35); full screen, 3 warnings, copy/paste/DevTools blocking, “Check” button |
| `admin.html` | **Supervisor panel** — password protected, live results, violation log, Excel (CSV) export |
| `sheet-ui.js` | Spreadsheet editor: cells, formula bar, fill handle, number formats, Data Validation, Conditional Formatting, PivotTable (slicer), Chart |
| `sheetengine.js` | Excel formula engine (~70 functions) + pivot/chart computation |
| `grader.js` | Grader: weighted checks, partial credit |
| `exam_data.js` | Tasks and grading checks of the 5 variants (**auto-generated**) |
| `sual_banki_xeritesi.md` | Which tasks are in which variant |
| `telebeler_35.py` | Split of 35 students into 5 variants (7 per variant) |
| `regrade.js` | Independent re-grading of the answers stored in the database (Node.js 18+) |
| `tools/bank_excel.js` | The 42 tasks: prefilled data, grading checks, reference solution |
| `tools/build_bank.js` | Builds the variants (`exam_data.js` + map) |
| `tools/verify.js` | Checks that every reference solution scores 100% and an untouched sheet scores 0 |
| `supabase_setup.sql` | Supabase tables, RLS rules and password-protected read functions |

---

## Question bank and variants
- Source: *Excel_GoogleSheets_Exam_Questions_Bank.docx* — **42 tasks, 8 modules**.
- Each variant has **6 tasks from 6 different modules**. Eight modules do not fit into six slots, so every variant covers 6 of the 8; **all 8 modules appear across the variants** (each module in 3–4 variants). See `sual_banki_xeritesi.md`.
- Task order is shuffled per variant. Points: 4 tasks × 17 + 2 tasks × 16 = 100 (harder tasks get 17).
- Data is prefilled in the sheet — the student only writes formulas / formats / rules.
- Excel-only features (Goal Seek, Data Table, Scenario Manager, Analysis ToolPak, `TEXTSPLIT`) are replaced by formula-based equivalents that test the same skill and work in any tool (stated in the task text).

---

## Setup

### 1. Supabase
- **Same project as before:** nothing to do. Results are stored under group `116i`, so they do not clash with older exams.
- **New project:** Dashboard → **SQL Editor** → paste `supabase_setup.sql` → **Run**.

### 2. Vercel
Push the folder to a **new GitHub repo** and import it as a **new Vercel project** (the old project stays untouched). Environment Variables: `SUPABASE_URL`, `SUPABASE_ANON_KEY`, `ADMIN_PASSWORD` (the build runs `node generate-env.js` to create `env.js`).

### 3. Variant split
```bash
python telebeler_35.py                 # template list
python telebeler_35.py list.txt        # one name per line (35 lines)
```
Give each student a **seat number (1–35)**. Variant = `(seat − 1) % 5`:
`1,6,11,… → A · 2,7,12,… → B · 3,8,13,… → C · 4,9,14,… → D · 5,10,15,…,35 → E`.

### 4. Check before the exam
Open the site, log in with seat 1 and a test name, type a formula, press **Check**; then open `/admin`, see the result and annul the test record.

---

## Anti-cheat
Mandatory full screen; 3 warnings (tab/window switch, leaving full screen, refresh) → on the 3rd the exam stops with 0 points; right-click, drag-and-drop, `F12`, `Ctrl+Shift+I/J/C`, `Ctrl+U` are blocked; copy/paste works only inside the exam (the spreadsheet's own clipboard). Time and answers survive a refresh.

## Re-grading
```bash
node regrade.js --password INSTRUCTOR_PASSWORD --csv new_scores.csv
```

## Changing the questions
Edit `tools/bank_excel.js`, then:
```bash
node tools/verify.js        # reference solutions must score 100%, an untouched sheet 0
node tools/build_bank.js    # regenerates exam_data.js and the map
```
