# MLCIS Gradebook Converter

React + Django application that batch-imports Grade 1-9 subject E-gradebooks into a fresh copy of the matching MLCIS scoring format. The original `.xlsm` template is never overwritten.

## Supported grade profiles

- Grades 1-3 use the G1-G3 scoring format and its 17 verified score columns.
- Grades 4-6 use the G4-G6 scoring format and its 16 verified score columns.
- Grades 7-9 use the G7-G9 scoring format and its 19 verified score columns.

The grade selector chooses the profile; it does not assume that all grades have
the same subjects. G4-G6 uses the approved quarter report-card template. G1-G3
and G7-G9 currently generate the scoring workbook only until their own Word
report-card templates are supplied and their merge fields are verified.

The supplied E-Gradebook G1-G8 format is also supported as an all-subject
source: it has a blank `SETTINGS!F23`, subject data on separate subject tabs,
and final Total values in the tab's `Total` column. Subject-tab aliases such as
`Reading`, `Callingraphy`, and `Kh-Social Studies` are matched to the relevant
grade profile. Subjects not present in the selected scoring template are not
written into a different column.

## Early Years K1-K3 (separate feature)

The **Early Years K1-K3** button opens a separate workflow for legacy
E-Progress files. It does not read, modify, or share templates with the G1-G9
E-Gradebook workflow.

1. Select K1, K2, or K3 and Term 1-4.
2. Upload one or more legacy `.xls` E-Progress workbooks. You can add a
   separate English, Khmer, Math, Science, or CP file for the same class and term.
3. Download a new `.xlsx` copy of the matching K1/K2/K3 Score System.

The service reads and verifies each `Student List` roster (or a populated
subject roster when the list has not been filled), then transfers only outcome values whose
three-row headers exactly match the selected Term sheet. K1 expects English,
Khmer, and CP; K2/K3 also expect Math and Science. Missing sheets and unmatched
outcomes are reported and left blank rather than guessed.

The supplied K1-K3 Score System templates are legacy `.xls` files. They are
copied and converted through **LibreOffice Calc** to `.xlsx`; the originals are
never changed. Local use therefore requires LibreOffice with Calc installed.
The Docker deployment installs `libreoffice-calc` automatically. K1/K3 legacy
Word report-card templates are not automated yet; their `.doc` layouts and
merge fields must be converted and verified before generating reports.

## Dynamic import flow

1. Select Q1, Q2, Q3, or Q4.
2. Add one or more subject E-gradebook `.xlsx` files. You may add files repeatedly before clicking Import.
3. The application detects each file's Subject from `SETTINGS!F23` and sends its `Scoring System!AM` cached Total into the correct selected-quarter subject column.

After a successful import, the application creates one PDF report card per
student from `1- F- Quarter Report Card Grade G4-6 - MLCIS (2).docx`. The
included Render Docker configuration installs LibreOffice Writer, so deployed
Render jobs produce a PDF ZIP. The report uses only the imported scores and
student identity. Unavailable subjects and attendance fields are left blank
rather than retaining template sample data.

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

The approved Grade 6 scoring and report templates are included in
`backend/templates/grade6`. This makes the app work both locally and after a
network deployment; it does not depend on `D:\New folder` existing on the
server. Set `MLCIS_SCORING_TEMPLATE` or `MLCIS_REPORT_TEMPLATE` only when you
intentionally need another approved format.

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

## Local network use

For another device on the same network, start Django with a reachable address,
for example `python manage.py runserver 0.0.0.0:8000`. Set the backend
environment values to include the server IP and the frontend origin:

```powershell
$env:ALLOWED_HOSTS="192.168.1.10,localhost,127.0.0.1"
$env:CORS_ALLOWED_ORIGINS="http://192.168.1.10:5173,http://localhost:5173"
```

In `frontend/.env.local`, set `VITE_API_BASE_URL=http://192.168.1.10:8000`,
then restart Vite. Do not use `localhost` on a second device: there it refers
to that device, not to the computer hosting Django.
