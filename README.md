# MLCIS Gradebook Converter

React + Django application that batch-imports Grade 6 subject E-gradebooks into a fresh copy of the MLCIS scoring format. The original `.xlsm` template is never overwritten.

## Dynamic import flow

1. Select Q1, Q2, Q3, or Q4.
2. Add one or more subject E-gradebook `.xlsx` files. You may add files repeatedly before clicking Import.
3. The application detects each file's Subject from `SETTINGS!F23` and sends its `Scoring System!AM` cached Total into the correct selected-quarter subject column.

After a successful import, the application also creates one PDF report card per
student from `1- F- Quarter Report Card Grade G4-6 - MLCIS (2).docx`. Download
the PDF ZIP to receive separate files, one for each student. The report uses
only the imported scores and student identity. Unavailable subjects and
attendance fields are left blank rather than retaining template sample data.

For example, Q1 plus English and Khmer Dictation files imports their Totals into `Q1!P` and `Q1!G` respectively. The first upload supplies the roster; each additional upload must have the identical roster, matched by student ID when present or exact English name otherwise.

Every subject's Teacher's Name is read from `SETTINGS!F20` and placed in that subject's Teacher's Name column in the selected `Comment Q#` sheet. English narrative comments are mapped from `Scoring System!AP:AR` to `Comment Q#!G:I`; other narrative comment destinations are deliberately not guessed. The template VBA project, formatting, formulas, merges, broken references, and external links are preserved.

## Print packages and continued scoring workbooks

The converter follows the school's four-print workflow:

1. **Print 1** imports Q1 and Teacher Comment into a clean scoring workbook, using `1- F- Quarter Report Card Grade G4-6 - MLCIS (3).docx` for its individual student reports.
2. **Print 2** requires the scoring `.xlsm` downloaded from Print 1, then imports Q2 without clearing Q1 or comments. Enter Exam 1 manually in Excel, recalculate, save, and print the template's `FinalSem1 reportcard` sheet.
3. **Print 3** continues the same scoring workbook and imports Q3 and its Teacher Comment without removing Q1, Q2, or Exam 1. It also uses `1- F- Quarter Report Card Grade G4-6 - MLCIS (3).docx`.
4. **Print 4** requires the continued scoring `.xlsm`, imports Q4, then the teacher enters Exam 2 manually, recalculates, saves, and prints the `FinalSem2` sheet.

When a continued workbook is uploaded, the converter preserves every existing quarter, comment, formula, macro, and manually entered exam cell. It only writes the quarter selected by the print package. A source E-gradebook whose `SETTINGS` period conflicts with the selected package is rejected instead of being placed into the wrong quarter.

`2- F- SEM1 and Q2 Report Card G4-6.doc` is configured as the semester report template for FS1 and FS2. `4- F- End of Year Report Card G4-6.doc` is configured as the end-of-year template. Both are legacy Word `.doc` files; the current implementation intentionally does not populate them until their merge fields are verified after safe conversion to `.docx`.

## Supported subject detection

Khmer Dictation, Khmer Essays, Khmer Reading, Khmer Public Speaking, Khmer Math, Khmer Social Studies, Khmer Science, Sport, Health Education, English, English Grammar, English Math, English Science, Chinese, RoboThink, and Computer.

## Template location

The default scoring-format reference is:

- `D:\New folder\Telegram Desktop\G4-G6 Scoring System Final 2024-2025-MLCIS -6A.xlsm`

Set `MLCIS_SCORING_TEMPLATE` if you need another approved format file.
Set `MLCIS_REPORT_TEMPLATE` to change the report-card Word template.

Each download begins as a clean format: it removes the sample values from all
Q1–Q4 and Comment Q1–Q4 sheets, then writes data only in the Quarter selected
for that import. Importing Q1 therefore leaves Q2, Q3, and Q4 empty.

## Run backend

```powershell
cd backend
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python manage.py test
python manage.py runserver
```

## Run frontend

```powershell
cd frontend
npm install
npm run dev
```

Open the Vite address shown in the terminal, normally `http://localhost:5173`. The frontend sends `/api` requests to Django at `http://127.0.0.1:8000`.
