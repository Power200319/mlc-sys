"""Standalone K1--K3 (Early Years) import service.

This module deliberately does not import, share, or modify the E-Gradebook
pipeline.  Early Years teachers use legacy ``.xls`` E-Progress workbooks and
the result is a separate converted Score System ``.xlsx`` workbook.
"""

from __future__ import annotations

import re
import shutil
import subprocess
import uuid
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlrd
from django.conf import settings
from openpyxl import load_workbook

from .services import ProcessingError


@dataclass(frozen=True)
class EarlyYearsProfile:
    grade: str
    subjects: tuple[str, ...]
    template: str


PROFILES = {
    "K1": EarlyYearsProfile("K1", ("English", "Khmer", "CP"), "templates/early-years/k1/score-system.xls"),
    "K2": EarlyYearsProfile("K2", ("English", "Math", "Science", "Khmer", "CP"), "templates/early-years/k2/score-system.xls"),
    "K3": EarlyYearsProfile("K3", ("English", "Math", "Science", "Khmer", "CP"), "templates/early-years/k3/score-system.xls"),
}
TERMS = {"T1": "Term 1", "T2": "Term 2", "T3": "Term 3", "T4": "Term 4"}
OLE_SIGNATURE = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
GRADE_PATTERN = re.compile(r"\b(K[1-3])(?:\s*[-–]\s*[A-Z0-9]+|\b)", re.IGNORECASE)


class XlsxSheetAdapter:
    """Present an openpyxl sheet through the small legacy-reader interface."""
    def __init__(self, sheet):
        self.sheet = sheet
        self.nrows = sheet.max_row
        self.ncols = sheet.max_column

    def cell_value(self, row: int, column: int):
        return self.sheet.cell(row + 1, column + 1).value


class XlsxBookAdapter:
    def __init__(self, workbook):
        self.workbook = workbook

    def sheet_names(self):
        return self.workbook.sheetnames

    def sheet_by_name(self, name: str):
        return XlsxSheetAdapter(self.workbook[name])

    def close(self):
        self.workbook.close()


def early_years_profiles() -> list[dict[str, Any]]:
    return [
        {"value": profile.grade, "label": f"{profile.grade} Early Years", "subjects": list(profile.subjects)}
        for profile in PROFILES.values()
    ]


def normalize(value: object) -> str:
    """Compare header labels without relying on layout column positions."""
    return "".join(char for char in str(value or "").casefold() if char.isalnum())


def display_value(value: object) -> object:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def is_legacy_xls(upload) -> bool:
    if upload is None or not upload.name.lower().endswith(".xls"):
        return False
    upload.seek(0)
    valid = upload.read(8) == OLE_SIGNATURE
    upload.seek(0)
    return valid


def _libreoffice() -> str | None:
    executable = shutil.which("libreoffice") or shutil.which("soffice")
    if executable:
        return executable
    # LibreOffice's Windows installer does not always add soffice to PATH.
    # Check its documented/default install locations so a backend restart is
    # sufficient after installation.
    if os.name == "nt":
        program_roots = (os.environ.get("ProgramW6432"), os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)"))
        for root in program_roots:
            if not root:
                continue
            candidate = Path(root) / "LibreOffice" / "program" / "soffice.exe"
            if candidate.is_file():
                return str(candidate)
    return None


def convert_xls_template(template: Path, output_dir: Path) -> Path:
    """Convert a *copy* of the supplied legacy template; never edit it."""
    executable = _libreoffice()
    if not executable:
        raise ProcessingError(
            "Early Years needs LibreOffice Calc to convert the approved .xls Score System. "
            "Install LibreOffice locally, or use the deployed server."
        )
    profile_dir = output_dir / "lo-profile"
    profile_dir.mkdir(exist_ok=True)
    command = [
        executable,
        f"-env:UserInstallation={profile_dir.as_uri()}",
        "--headless",
        "--convert-to", "xlsx:Calc MS Excel 2007 XML",
        "--outdir", str(output_dir),
        str(template),
    ]
    completed = subprocess.run(command, capture_output=True, text=True, timeout=120, check=False)
    converted = output_dir / f"{template.stem}.xlsx"
    if completed.returncode or not converted.exists():
        detail = (completed.stderr or completed.stdout or "conversion produced no workbook").strip()
        raise ProcessingError(f"Could not prepare the Early Years Score System: {detail[:240]}")
    return converted


def open_recalculated_source(source_path: Path, output_dir: Path):
    """Use Calc to calculate legacy E-Progress formulas before reading them.

    Teacher .xls files commonly display scores in Excel but have stale or empty
    cached formula values.  Converting only a temporary copy makes Calc refresh
    those values; the original uploaded file is never altered.
    """
    converted = convert_xls_template(source_path, output_dir)
    workbook = load_workbook(converted, read_only=True, data_only=True)
    return XlsxBookAdapter(workbook), converted


def _header_rows_xls(sheet, column: int) -> tuple[str, str, str]:
    # E-Progress subject headers occupy Excel rows 9--11.
    return tuple(normalize(sheet.cell_value(row, column)) for row in (8, 9, 10))


def _header_rows_xlsx(sheet, column: int) -> tuple[str, str, str]:
    # Score System uses the same three visual header rows.
    return tuple(normalize(sheet.cell(row + 1, column + 1).value) for row in (8, 9, 10))


def _first_data_row_xls(sheet) -> int:
    for row in range(11, sheet.nrows):
        if str(sheet.cell_value(row, 1)).strip():
            return row
    return sheet.nrows


def _first_data_row_xlsx(sheet) -> int:
    for row in range(12, sheet.max_row + 1):
        if str(sheet.cell(row, 2).value or "").strip():
            return row
    # Legacy Score System templates merge the identity headers through row 12.
    # When sample students have just been cleared, row 12 is a MergedCell, so
    # it must never become a data row.
    return 13


def source_students(book: xlrd.book.Book) -> list[dict[str, object]]:
    if "Student List" not in book.sheet_names():
        raise ProcessingError("E-Progress workbook is missing the 'Student List' sheet.")
    sheet = book.sheet_by_name("Student List")

    def rows_from_sheet(candidate) -> list[dict[str, object]]:
        rows: list[dict[str, object]] = []
        for row in range(_first_data_row_xls(candidate), candidate.nrows):
            name = str(candidate.cell_value(row, 1)).strip()
            if not name:
                continue
            rows.append({
                "No": display_value(candidate.cell_value(row, 0)),
                "Full Name": name,
                "Khmer Name": display_value(candidate.cell_value(row, 2)),
                "ID": display_value(candidate.cell_value(row, 3)),
                "Sex": display_value(candidate.cell_value(row, 4)),
                "Group": display_value(candidate.cell_value(row, 5)),
            })
        return rows

    students = rows_from_sheet(sheet)
    # Teachers sometimes fill roster data only on their subject sheet.  This
    # is still an E-Progress layout, so use it as a verified roster fallback.
    if not students:
        for name in book.sheet_names():
            if name == "Student List":
                continue
            students = rows_from_sheet(book.sheet_by_name(name))
            if students:
                break
    if not students:
        raise ProcessingError("No students were found in E-Progress 'Student List'.")
    return students


def source_grade(book: xlrd.book.Book) -> str | None:
    """Read K1/K2/K3 from the visible E-Progress class title when present."""
    for sheet_name in book.sheet_names():
        sheet = book.sheet_by_name(sheet_name)
        for row in range(min(sheet.nrows, 12)):
            for column in range(min(sheet.ncols, 10)):
                match = GRADE_PATTERN.search(str(sheet.cell_value(row, column) or ""))
                if match:
                    return match.group(1).upper()
    return None


def subjects_named_by_file(filename: str, profile: EarlyYearsProfile) -> list[str]:
    """Respect a teacher's explicit `_English` / `_Khmer` file naming.

    E-Progress copies contain every subject tab, often with old values in tabs
    the teacher did not intend to submit. A recognized filename limits import
    to that one subject; generic filenames retain all populated subjects.
    """
    normalized_name = normalize(Path(filename).stem)
    matches = [subject for subject in profile.subjects if normalize(subject) in normalized_name]
    return matches[:1]


def student_key(name: object, student_id: object) -> str:
    """Use ID when supplied; name is the cautious fallback for legacy files."""
    normalized_id = str(student_id or "").strip().casefold()
    return f"id:{normalized_id}" if normalized_id else f"name:{str(name or '').strip().casefold()}"


def rosters_match(reference: list[dict[str, object]], candidate: list[dict[str, object]]) -> bool:
    if len(reference) != len(candidate):
        return False
    for left, right in zip(reference, candidate):
        left_id = str(left.get("ID") or "").strip().casefold()
        right_id = str(right.get("ID") or "").strip().casefold()
        if left_id and right_id:
            if left_id != right_id:
                return False
        elif str(left.get("Full Name") or "").strip().casefold() != str(right.get("Full Name") or "").strip().casefold():
            return False
    return True


def source_subject_values(book: xlrd.book.Book, subject: str) -> tuple[dict[str, dict[tuple[str, str, str], object]], list[str]]:
    if subject not in book.sheet_names():
        return {}, [f"{subject}: sheet not present in this E-Progress file; it was skipped."]
    sheet = book.sheet_by_name(subject)
    headers = {column: _header_rows_xls(sheet, column) for column in range(6, sheet.ncols)}

    def is_metadata_column(column: int) -> bool:
        raw_header = " ".join(str(sheet.cell_value(row, column) or "").casefold() for row in (8, 9, 10))
        return "student" in raw_header or "remark" in raw_header or "ចំនួនសិស្ស" in raw_header

    values: dict[str, dict[tuple[str, str, str], object]] = {}
    for row in range(_first_data_row_xls(sheet), sheet.nrows):
        name = str(sheet.cell_value(row, 1)).strip()
        if not name:
            continue
        values[student_key(name, sheet.cell_value(row, 3))] = {
            header: display_value(sheet.cell_value(row, column))
            for column, header in headers.items()
            if any(header) and not is_metadata_column(column) and sheet.cell_value(row, column) not in ("", None)
        }
    return values, []


def input_score_last_column(sheet) -> int:
    """Locate the end of score-entry columns before duplicated result blocks."""
    english_columns = [
        column for column in range(7, sheet.max_column + 1)
        if _header_rows_xlsx(sheet, column - 1)[0] == "english"
    ]
    # K2/K3 result areas repeat the subject headers after the input area.
    return english_columns[1] - 1 if len(english_columns) > 1 else sheet.max_column


def target_columns(sheet) -> dict[tuple[str, str, str], list[int]]:
    """Return only unambiguous outcome columns discovered from target headers."""
    grouped: dict[tuple[str, str, str], list[int]] = {}
    for column in range(7, input_score_last_column(sheet) + 1):
        header = _header_rows_xlsx(sheet, column - 1)
        if any(header):
            grouped.setdefault(header, []).append(column)
    return grouped


def target_outcome_columns(sheet) -> dict[str, list[int]]:
    """Index the final (learning-outcome) header for a cautious fallback.

    Different teacher E-Progress copies sometimes differ in the group title
    while retaining the actual outcome wording.  Only a uniquely occurring
    outcome may use this fallback; repeated labels such as ``Overall Grade``
    remain unresolved rather than being guessed.
    """
    grouped: dict[str, list[int]] = {}
    for column in range(7, input_score_last_column(sheet) + 1):
        # English uses row 11 for the outcome; Khmer and some older layouts
        # use row 10 and leave row 11 blank.  The right-most non-empty level
        # is therefore the most specific available label.
        outcome = next((label for label in reversed(_header_rows_xlsx(sheet, column - 1)) if label), "")
        if outcome:
            grouped.setdefault(outcome, []).append(column)
    return grouped


def target_identity_columns(sheet) -> dict[str, int]:
    columns: dict[str, int] = {"No": 1}
    aliases = {
        "no": "No", "n": "No", "fullname": "Full Name", "khmername": "Khmer Name",
        "fullnameinkhmer": "Khmer Name", "id": "ID", "sex": "Sex", "group": "Group",
    }
    for column in range(1, min(sheet.max_column, 12) + 1):
        label = normalize(sheet.cell(9, column).value)
        if label in aliases:
            columns[aliases[label]] = column
    return columns


def clear_input_constants(sheet) -> None:
    """Remove sample typed values while preserving formulas and formatting."""
    start = _first_data_row_xlsx(sheet)
    for row in range(start, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            cell = sheet.cell(row, column)
            if cell.value is not None and not (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.value = None


def target_term_sheet_name(workbook, term: str) -> str:
    """Resolve legacy term-tab spelling without assuming one template version."""
    expected = normalize(TERMS[term])
    for sheet_name in workbook.sheetnames:
        normalized = normalize(sheet_name)
        if normalized in {expected, f"{expected}score"}:
            return sheet_name
    raise ProcessingError(
        f"Score System is missing the {TERMS[term]} score sheet. "
        f"Available sheets: {', '.join(workbook.sheetnames)}"
    )


def populate_term(
    workbook,
    term: str,
    students: list[dict[str, object]],
    subject_values: dict[str, dict[str, dict[tuple[str, str, str], object]]],
    profile: EarlyYearsProfile,
) -> tuple[int, list[str]]:
    term_sheet = target_term_sheet_name(workbook, term)
    target = workbook[term_sheet]
    clear_input_constants(target)
    identities = target_identity_columns(target)
    required = {"Full Name", "ID"}
    if not required.issubset(identities):
        raise ProcessingError(f"{term_sheet}: could not verify student identity headers in the Score System.")
    start = _first_data_row_xlsx(target)
    # Keep both identifiers. A subject file may omit IDs while another file in
    # the same batch has them, in which case an exact name remains the allowed
    # fallback.
    student_rows: dict[str, int] = {}
    for index, student in enumerate(students):
        row = start + index
        name_key = student_key(student["Full Name"], "")
        student_rows[name_key] = row
        if str(student.get("ID") or "").strip():
            student_rows[student_key(student["Full Name"], student.get("ID"))] = row
    for student in students:
        row = student_rows[student_key(student["Full Name"], student.get("ID"))]
        for key, column in identities.items():
            target.cell(row, column).value = student.get(key, "")

    targets = target_columns(target)
    target_outcomes = target_outcome_columns(target)
    transferred = 0
    warnings: list[str] = []
    for subject in profile.subjects:
        source_values = subject_values.get(subject)
        if source_values is None:
            continue
        matched_headers = 0
        fallback_headers = 0
        for key, outcomes in source_values.items():
            target_row = student_rows.get(key)
            if target_row is None:
                warnings.append(f"{subject}: one student is not in the verified roster and was skipped.")
                continue
            for header, value in outcomes.items():
                destination = targets.get(header, [])
                used_fallback = False
                source_outcome = next((label for label in reversed(header) if label), "")
                if len(destination) != 1 and source_outcome:
                    candidates = target_outcomes.get(source_outcome, [])
                    if len(candidates) == 1:
                        destination = candidates
                        used_fallback = True
                if len(destination) == 1:
                    target.cell(target_row, destination[0]).value = value
                    transferred += 1
                    matched_headers += 1
                    fallback_headers += int(used_fallback)
        if source_values and not matched_headers:
            warnings.append(f"{subject}: no verified outcome headers matched {term_sheet}; no {subject} scores were written.")
        elif fallback_headers:
            warnings.append(f"{subject}: {fallback_headers} outcome value(s) matched by a unique learning-outcome header.")
    return transferred, list(dict.fromkeys(warnings))


def process_early_years(uploads, grade: str, term: str) -> dict[str, Any]:
    grade = (grade or "").upper().strip()
    term = (term or "").upper().strip()
    profile = PROFILES.get(grade)
    if profile is None:
        raise ProcessingError("Choose Early Years grade K1, K2, or K3.")
    if term not in TERMS:
        raise ProcessingError("Choose Term 1, Term 2, Term 3, or Term 4.")
    if not uploads:
        raise ProcessingError("Choose one or more legacy .xls E-Progress workbooks.")
    if any(not is_legacy_xls(upload) for upload in uploads):
        raise ProcessingError("Every selected Early Years E-Progress file must be a valid legacy .xls workbook.")

    job_id = uuid.uuid4().hex
    job_dir = Path(settings.MLCIS_JOB_DIR) / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    source_paths: list[Path] = []
    for index, upload in enumerate(uploads, start=1):
        source_path = job_dir / f"e-progress-{index}.xls"
        with source_path.open("wb+") as destination:
            for chunk in upload.chunks():
                destination.write(chunk)
        source_paths.append(source_path)
    try:
        students: list[dict[str, object]] | None = None
        subject_data: dict[str, dict[str, dict[tuple[str, str, str], object]]] = {}
        source_names: dict[str, str] = {}
        warnings: list[str] = []
        for source_path, upload in zip(source_paths, uploads):
            # Read teacher source files directly. Recalculating every source
            # through LibreOffice can hang on Windows and block a class import
            # for several minutes. The Score System template itself is still
            # converted once below because it is a legacy .xls workbook.
            source_book = xlrd.open_workbook(source_path, formatting_info=False)
            try:
                detected_grade = source_grade(source_book)
                if detected_grade and detected_grade != grade:
                    raise ProcessingError(
                        f"{upload.name} is an {detected_grade} E-Progress file, but {grade} is selected. "
                        "Choose the matching Early Years grade before importing."
                    )
                roster = source_students(source_book)
                if students is None:
                    students = roster
                elif not rosters_match(students, roster):
                    raise ProcessingError(
                        f"{upload.name}: student roster does not match the first selected E-Progress file. "
                        "Use files for the same class; matching uses ID when present, otherwise name and order."
                    )
                named_subjects = subjects_named_by_file(upload.name, profile)
                detected = [subject for subject in named_subjects if subject in source_book.sheet_names()] if named_subjects else [
                    subject for subject in profile.subjects if subject in source_book.sheet_names()
                ]
                if not detected:
                    warnings.append(f"{upload.name}: no expected {grade} subject sheet was found; file was skipped.")
                    continue
                for subject in detected:
                    values, subject_warnings = source_subject_values(source_book, subject)
                    warnings.extend(subject_warnings)
                    if not any(outcomes for outcomes in values.values()):
                        continue
                    if subject in subject_data:
                        warnings.append(f"{subject}: {upload.name} replaced subject values from {source_names[subject]}.")
                    subject_data[subject] = values
                    source_names[subject] = upload.name
            finally:
                # xlrd workbooks hold no OS file handle that needs closing.
                pass
        if students is None:
            raise ProcessingError("No students were found in the selected E-Progress workbooks.")
        if not subject_data:
            raise ProcessingError(f"No expected {grade} subject sheets were found in the selected files.")
        template = Path(settings.BASE_DIR) / profile.template
        if not template.exists():
            raise ProcessingError(f"Approved {grade} Score System template is not configured.")
        converted = convert_xls_template(template, job_dir)
        output = job_dir / f"{grade}_{term}_Score_System.xlsx"
        workbook = load_workbook(converted, data_only=False)
        transferred, populate_warnings = populate_term(workbook, term, students, subject_data, profile)
        warnings.extend(populate_warnings)
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(output)
        workbook.close()
        converted.unlink(missing_ok=True)
    finally:
        for source_path in source_paths:
            source_path.unlink(missing_ok=True)

    return {
        "job_id": job_id,
        "grade": grade,
        "term": term,
        "student_count": len(students),
        "fields_transferred": transferred,
        "subjects": [subject for subject in profile.subjects if subject in subject_data],
        "warnings": [
            "Early Years is a separate K1-K3 workflow. No E-Gradebook data or templates were used.",
            "The approved legacy .xls Score System was copied and converted to .xlsx; the original template remains unchanged.",
            "Only outcome columns whose three-row header matched exactly were imported. Unmatched layout fields were left blank.",
            f"Imported subjects: {', '.join(subject for subject in profile.subjects if subject in subject_data)}.",
            *warnings,
        ],
        "output": {"available": True, "filename": output.name, "download": f"/api/early-years/download/{job_id}/"},
    }


def consume_early_years_output(job_id: str) -> tuple[bytes, str]:
    if not re.fullmatch(r"[0-9a-f]{32}", job_id or ""):
        raise ProcessingError("Requested output was not found.")
    job_dir = Path(settings.MLCIS_JOB_DIR) / job_id
    matches = list(job_dir.glob("K[1-3]_T[1-4]_Score_System.xlsx"))
    if len(matches) != 1:
        raise ProcessingError("Requested output was not found.")
    path = matches[0]
    content = path.read_bytes()
    shutil.rmtree(job_dir, ignore_errors=True)
    return content, path.name
