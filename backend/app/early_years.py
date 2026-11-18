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
    return shutil.which("libreoffice") or shutil.which("soffice")


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
    return 12


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


def source_subject_values(book: xlrd.book.Book, subject: str) -> tuple[dict[str, dict[tuple[str, str, str], object]], list[str]]:
    if subject not in book.sheet_names():
        return {}, [f"{subject}: sheet not present in this E-Progress file; it was skipped."]
    sheet = book.sheet_by_name(subject)
    headers = {column: _header_rows_xls(sheet, column) for column in range(6, sheet.ncols)}
    values: dict[str, dict[tuple[str, str, str], object]] = {}
    for row in range(_first_data_row_xls(sheet), sheet.nrows):
        name = str(sheet.cell_value(row, 1)).strip()
        if not name:
            continue
        values[name.casefold()] = {
            header: display_value(sheet.cell_value(row, column))
            for column, header in headers.items()
            if any(header) and sheet.cell_value(row, column) not in ("", None)
        }
    return values, []


def target_columns(sheet) -> dict[tuple[str, str, str], list[int]]:
    """Return only unambiguous outcome columns discovered from target headers."""
    grouped: dict[tuple[str, str, str], list[int]] = {}
    for column in range(7, sheet.max_column + 1):
        header = _header_rows_xlsx(sheet, column - 1)
        if any(header):
            grouped.setdefault(header, []).append(column)
    return grouped


def target_identity_columns(sheet) -> dict[str, int]:
    columns: dict[str, int] = {}
    aliases = {
        "no": "No", "fullname": "Full Name", "khmername": "Khmer Name",
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


def populate_term(workbook, term_sheet: str, students: list[dict[str, object]], source_book: xlrd.book.Book, profile: EarlyYearsProfile) -> tuple[int, list[str]]:
    if term_sheet not in workbook.sheetnames:
        raise ProcessingError(f"Score System is missing '{term_sheet}'.")
    target = workbook[term_sheet]
    clear_input_constants(target)
    identities = target_identity_columns(target)
    required = {"Full Name", "ID"}
    if not required.issubset(identities):
        raise ProcessingError(f"{term_sheet}: could not verify student identity headers in the Score System.")
    start = _first_data_row_xlsx(target)
    student_rows = {str(student["Full Name"]).casefold(): start + index for index, student in enumerate(students)}
    for student in students:
        row = student_rows[str(student["Full Name"]).casefold()]
        for key, column in identities.items():
            target.cell(row, column).value = student.get(key, "")

    targets = target_columns(target)
    transferred = 0
    warnings: list[str] = []
    for subject in profile.subjects:
        source_values, source_warnings = source_subject_values(source_book, subject)
        warnings.extend(source_warnings)
        matched_headers = 0
        for name, outcomes in source_values.items():
            target_row = student_rows.get(name)
            if target_row is None:
                warnings.append(f"{subject}: '{name}' is not in Student List and was skipped.")
                continue
            for header, value in outcomes.items():
                destination = targets.get(header, [])
                if len(destination) == 1:
                    target.cell(target_row, destination[0]).value = value
                    transferred += 1
                    matched_headers += 1
        if source_values and not matched_headers:
            warnings.append(f"{subject}: no verified outcome headers matched {term_sheet}; no {subject} scores were written.")
    return transferred, list(dict.fromkeys(warnings))


def process_early_years(upload, grade: str, term: str) -> dict[str, Any]:
    grade = (grade or "").upper().strip()
    term = (term or "").upper().strip()
    profile = PROFILES.get(grade)
    if profile is None:
        raise ProcessingError("Choose Early Years grade K1, K2, or K3.")
    if term not in TERMS:
        raise ProcessingError("Choose Term 1, Term 2, Term 3, or Term 4.")
    if not is_legacy_xls(upload):
        raise ProcessingError("Upload one valid legacy .xls E-Progress workbook for Early Years.")

    job_id = uuid.uuid4().hex
    job_dir = Path(settings.MLCIS_JOB_DIR) / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    source_path = job_dir / "e-progress.xls"
    with source_path.open("wb+") as destination:
        for chunk in upload.chunks():
            destination.write(chunk)
    try:
        source_book = xlrd.open_workbook(source_path, formatting_info=False)
        students = source_students(source_book)
        template = Path(settings.BASE_DIR) / profile.template
        if not template.exists():
            raise ProcessingError(f"Approved {grade} Score System template is not configured.")
        converted = convert_xls_template(template, job_dir)
        output = job_dir / f"{grade}_{term}_Score_System.xlsx"
        workbook = load_workbook(converted, data_only=False)
        transferred, warnings = populate_term(workbook, TERMS[term], students, source_book, profile)
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(output)
        workbook.close()
        converted.unlink(missing_ok=True)
    finally:
        source_path.unlink(missing_ok=True)

    return {
        "job_id": job_id,
        "grade": grade,
        "term": term,
        "student_count": len(students),
        "fields_transferred": transferred,
        "subjects": list(profile.subjects),
        "warnings": [
            "Early Years is a separate K1-K3 workflow. No E-Gradebook data or templates were used.",
            "The approved legacy .xls Score System was copied and converted to .xlsx; the original template remains unchanged.",
            "Only outcome columns whose three-row header matched exactly were imported. Unmatched layout fields were left blank.",
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
