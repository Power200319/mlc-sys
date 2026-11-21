"""Batch-import subject E-gradebooks into a fresh scoring-format copy.

Supports multiple grades via grade configuration files in backend/config/grades/.
Pass `grade` (1-12) to functions that need grade-specific subject mappings,
column mappings, credits, report cell maps, or template paths.
"""

from __future__ import annotations

import uuid
import zipfile
import json
import time
import os
import gc
from copy import copy
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any
import re
import shutil
import subprocess
import tempfile
from xml.sax.saxutils import escape

from django.conf import settings
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string

from .grade_config import GradeConfig, SubjectMapping, load_grade_config


class ProcessingError(Exception):
    pass


@dataclass
class Student:
    source_row: int
    number: int | None
    english_name: str
    khmer_name: str
    student_id: str
    sex: str
    group: str
    total: float | None
    comments: tuple[str | None, str | None, str | None]


@dataclass
class SubjectUpload:
    filename: str
    subject_name: str
    teacher_name: str
    mapping: SubjectMapping
    source_period: str | None
    students: list[Student]


def is_xlsx_upload(upload) -> bool:
    upload.seek(0)
    valid = upload.name.lower().endswith(".xlsx") and zipfile.is_zipfile(upload)
    upload.seek(0)
    return valid


def is_scoring_workbook_upload(upload) -> bool:
    """A continued scoring workbook must remain an Excel macro workbook."""
    if upload is None:
        return False
    upload.seek(0)
    valid = upload.name.lower().endswith(".xlsm") and zipfile.is_zipfile(upload)
    upload.seek(0)
    return valid


def configured_templates(grade: int = 6) -> dict[str, Path]:
    """Return template paths for a given grade."""
    config = load_grade_config(grade)
    base = Path(settings.BASE_DIR)
    def template_path(relative_path: str | None) -> Path | None:
        return base / relative_path if relative_path else None

    return {
        "scoring": template_path(config.scoring_template),
        "quarter_report": template_path(config.report_template),
        "semester_report": template_path(config.semester_template),
        "end_year_report": template_path(config.end_year_template),
    }


def purge_expired_jobs() -> None:
    """Remove unclaimed output files after the short privacy retention window."""
    jobs_dir = Path(settings.MLCIS_JOB_DIR)
    if not jobs_dir.exists():
        return
    cutoff = time.time() - settings.MLCIS_JOB_RETENTION_SECONDS
    for job_dir in jobs_dir.iterdir():
        if job_dir.is_dir() and job_dir.stat().st_mtime < cutoff:
            shutil.rmtree(job_dir, ignore_errors=True)


def _source_setting(workbook_path: Path, cell: str) -> str:
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if "SETTINGS" not in workbook.sheetnames:
            raise ProcessingError("An uploaded file does not contain a SETTINGS sheet.")
        return str(workbook["SETTINGS"][cell].value or "").strip()
    finally:
        workbook.close()


def source_period(workbook_path: Path, grade: int = 6) -> str | None:
    value = _source_setting(workbook_path, "F14")
    normalised = value.upper().replace("QUARTER", "Q").replace(" ", "")
    supported = load_grade_config(grade).periods
    return normalised if normalised in supported else None


def source_subject(workbook_path: Path) -> str:
    return _source_setting(workbook_path, "F23")


def source_teacher_name(workbook_path: Path) -> str:
    """Teacher's Name is stored in SETTINGS!F20 in the supplied E-gradebook."""
    return _source_setting(workbook_path, "F20")


def _read_subject_students_sheet(sheet) -> list[Student]:
    """Read a loaded common E-gradebook Scoring System sheet once."""
    if str(sheet["AM9"].value or "").strip().casefold() != "total":
        raise ProcessingError("An uploaded file has no verified Scoring System Total column (AM9).")
    students: list[Student] = []
    blank_streak = 0
    for row in range(11, sheet.max_row + 1):
        name = sheet[f"C{row}"].value
        if not name:
            blank_streak += 1
            if students and blank_streak >= 3:
                break
            continue
        blank_streak = 0
        total = sheet[f"AM{row}"].value
        students.append(Student(
            source_row=row,
            number=sheet[f"B{row}"].value if isinstance(sheet[f"B{row}"].value, int) else None,
            english_name=str(name).strip(),
            khmer_name=str(sheet[f"D{row}"].value or "").strip(),
            student_id=str(sheet[f"E{row}"].value or "").strip(),
            sex=str(sheet[f"F{row}"].value or "").strip(),
            group=str(sheet[f"G{row}"].value or "").strip(),
            total=float(total) if isinstance(total, (int, float)) else None,
            comments=tuple(str(sheet[f"{column}{row}"].value or "").strip() or None for column in ("AP", "AQ", "AR")),
        ))
    if not students:
        raise ProcessingError("No student rows with displayed names were found in Scoring System!C11:C.")
    return students


def read_subject_students(workbook_path: Path) -> list[Student]:
    """Compatibility wrapper for callers that only need one standard subject."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        if "Scoring System" not in workbook.sheetnames:
            raise ProcessingError("An uploaded file does not contain a Scoring System sheet.")
        return _read_subject_students_sheet(workbook["Scoring System"])
    finally:
        workbook.close()


def _first_nonempty(values) -> str:
    return next((str(value).strip() for value in values if values and value not in (None, "")), "")


def _read_multi_subject_sheet(sheet, sheet_name: str) -> tuple[list[Student], str]:
    """Read one already-loaded actual subject worksheet."""
    limit = min(sheet.max_column, 100)
    total_column = next((column for column in range(1, limit) if str(sheet.cell(9, column).value or "").strip().casefold() == "total" and "letter" in str(sheet.cell(9, column + 1).value or "").casefold()), None)
    if total_column is None and sheet_name == "Health Education ":
        total_column = 14
    if total_column is None:
        raise ProcessingError(f"{sheet_name}: could not locate final Total beside Letter Grade.")
    teacher_column = next((column for column in range(1, limit + 1) if "teacher" in str(sheet.cell(9, column).value or "").casefold()), None)
    comment_columns = (teacher_column - 3, teacher_column - 2, teacher_column - 1) if teacher_column and teacher_column > 3 else None
    students: list[Student] = []
    blank_streak = 0
    teacher_names: list[str] = []
    for row in range(11, sheet.max_row + 1):
        name = sheet[f"C{row}"].value
        if not name:
            blank_streak += 1
            if students and blank_streak >= 3:
                break
            continue
        blank_streak = 0
        total = sheet.cell(row, total_column).value
        if teacher_column:
            teacher_names.append(str(sheet.cell(row, teacher_column).value or "").strip())
        comments = tuple(str(sheet.cell(row, column).value or "").strip() or None for column in comment_columns) if comment_columns else (None, None, None)
        students.append(Student(
            source_row=row, number=sheet[f"B{row}"].value if isinstance(sheet[f"B{row}"].value, int) else None,
            english_name=str(name).strip(), khmer_name=str(sheet[f"D{row}"].value or "").strip(),
            student_id=str(sheet[f"E{row}"].value or "").strip(), sex=str(sheet[f"F{row}"].value or "").strip(),
            group=str(sheet[f"G{row}"].value or "").strip(), total=float(total) if isinstance(total, (int, float)) else None, comments=comments,
        ))
    if not students:
        raise ProcessingError(f"{sheet_name}: no student rows were found.")
    return students, _first_nonempty(teacher_names)


def read_multi_subject_sheet(workbook_path: Path, sheet_name: str) -> tuple[list[Student], str]:
    """Compatibility wrapper for callers that need a single subject sheet."""
    workbook = load_workbook(workbook_path, read_only=True, data_only=True)
    try:
        return _read_multi_subject_sheet(workbook[sheet_name], sheet_name)
    finally:
        workbook.close()


def ensure_scoring_template_is_usable(path: Path, grade: int = 6) -> None:
    if not path.exists():
        raise ProcessingError("The scoring template file is not configured or cannot be found.")
    config = load_grade_config(grade)
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        required = set(config.period_sheets.values()) | set(config.comment_sheets.values())
        if not required.issubset(workbook.sheetnames):
            raise ProcessingError("The scoring template is missing a required Quarter or Comment Quarter sheet.")
    finally:
        workbook.close()
    with zipfile.ZipFile(path) as archive:
        if "xl/vbaProject.bin" not in archive.namelist():
            raise ProcessingError("The scoring template VBA project could not be read.")


def roster_matches(reference: list[Student], candidate: list[Student]) -> bool:
    if len(reference) != len(candidate):
        return False
    for left, right in zip(reference, candidate):
        if left.student_id and right.student_id:
            if left.student_id != right.student_id:
                return False
        elif left.english_name.casefold() != right.english_name.casefold():
            return False
    return True


def clear_sample_constants(sheet, first_data_row: int, last_column: int) -> None:
    """Remove sample data without deleting the format's formulas or styling."""
    for row in range(first_data_row, sheet.max_row + 1):
        for column in range(1, last_column + 1):
            cell = sheet.cell(row, column)
            if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                cell.value = None


def clear_all_quarters(workbook, grade: int = 6) -> None:
    """An output contains data only for quarters imported in this job."""
    config = load_grade_config(grade)
    last_score_column = max(
        column_index_from_string(subject.target_score_column)
        for subject in config.subjects
    )
    for quarter in config.period_sheets.values():
        clear_sample_constants(workbook[quarter], 15, last_score_column)
    for comment_quarter in config.comment_sheets.values():
        clear_sample_constants(workbook[comment_quarter], 9, workbook[comment_quarter].max_column)


def extend_quarter_rows(quarter_sheet, required_students: int) -> None:
    """Extend the final student row when the supplied format is too short."""
    first_row = 15
    existing_students = quarter_sheet.max_row - first_row + 1
    while existing_students < required_students:
        source_row = first_row + existing_students - 1
        destination_row = source_row + 1
        quarter_sheet.row_dimensions[destination_row].height = quarter_sheet.row_dimensions[source_row].height
        for column in range(1, quarter_sheet.max_column + 1):
            source = quarter_sheet.cell(source_row, column)
            target = quarter_sheet.cell(destination_row, column)
            if isinstance(source.value, str) and source.value.startswith("="):
                target.value = Translator(source.value, origin=source.coordinate).translate_formula(target.coordinate)
            else:
                target.value = source.value
            if source.has_style:
                target._style = copy(source._style)
            if source.number_format:
                target.number_format = source.number_format
            if source.alignment:
                target.alignment = copy(source.alignment)
            if source.font:
                target.font = copy(source.font)
            if source.fill:
                target.fill = copy(source.fill)
            if source.border:
                target.border = copy(source.border)
            if source.protection:
                target.protection = copy(source.protection)
        existing_students += 1


def target_identity_columns(quarter_sheet) -> dict[str, str]:
    """Locate roster fields by the format's visible row-13 headers.

    Q1-Q3 place Sex before ID, while the supplied Q4 formats place ID before
    Sex. Reading labels prevents a Q4 import from silently swapping them.
    """
    columns = {
        "number": "A",
        "english_name": "B",
        "khmer_name": "C",
        "sex": "D",
        "student_id": "E",
        "group": "F",
    }
    header_fields = {
        "fullname": "english_name",
        "fullnameinkhmer": "khmer_name",
        "sex": "sex",
        "id": "student_id",
        "group": "group",
    }
    for column in range(1, min(quarter_sheet.max_column, 20) + 1):
        value = quarter_sheet.cell(13, column).value
        label = "".join(character for character in str(value or "").casefold() if character.isalnum())
        field = header_fields.get(label)
        if field:
            columns[field] = quarter_sheet.cell(13, column).column_letter
    return columns


def comment_identity_columns(comment_sheet) -> dict[str, str]:
    """Locate Comment Q roster headers and do not rely on broken template links.

    Several approved formats have stale formulas in Comment Q that point at a
    different quarter or swap ID and Sex.  Writing by the visible row-6 header
    guarantees comments carry the same identity as the imported quarter.
    """
    columns = {
        "number": "A",
        "english_name": "B",
        "khmer_name": "C",
        "student_id": "D",
        "sex": "E",
        "group": "F",
    }
    header_fields = {
        "nº": "number",
        "no": "number",
        "number": "number",
        "fullname": "english_name",
        "fullnameinkhmer": "khmer_name",
        "id": "student_id",
        "sex": "sex",
        "group": "group",
    }
    for column in range(1, min(comment_sheet.max_column, 20) + 1):
        value = comment_sheet.cell(6, column).value
        label = "".join(character for character in str(value or "").casefold() if character.isalnum())
        field = header_fields.get(label)
        if field:
            columns[field] = comment_sheet.cell(6, column).column_letter
    return columns


def populate_scoring_template(
    template_path: Path,
    output_path: Path,
    period: str,
    uploads: list[SubjectUpload],
    grade: int = 6,
    *,
    clear_samples: bool = True,
) -> tuple[int, int]:
    config = load_grade_config(grade)
    workbook = load_workbook(template_path, keep_vba=template_path.suffix.lower() == ".xlsm", keep_links=True)
    try:
        quarter_sheet = workbook[config.period_sheets[period]]
        comment_sheet = workbook[config.comment_sheets[period]]
        roster = uploads[0].students
        first_target_row, first_comment_row = 15, 9
        comment_capacity = comment_sheet.max_row - first_comment_row + 1
        if len(roster) > comment_capacity:
            raise ProcessingError(f"The {period} Comment format has room for {comment_capacity} students, but an upload contains {len(roster)}.")
        extend_quarter_rows(quarter_sheet, len(roster))

        if clear_samples:
            clear_all_quarters(workbook, grade)
        identity_columns = target_identity_columns(quarter_sheet)
        comment_columns = comment_identity_columns(comment_sheet)
        for index, student in enumerate(roster):
            row = first_target_row + index
            quarter_sheet[f"{identity_columns['number']}{row}"] = student.number or index + 1
            quarter_sheet[f"{identity_columns['english_name']}{row}"] = student.english_name
            quarter_sheet[f"{identity_columns['khmer_name']}{row}"] = student.khmer_name
            quarter_sheet[f"{identity_columns['sex']}{row}"] = student.sex
            quarter_sheet[f"{identity_columns['student_id']}{row}"] = student.student_id
            quarter_sheet[f"{identity_columns['group']}{row}"] = student.group
            comment_row = first_comment_row + index
            comment_sheet[f"{comment_columns['number']}{comment_row}"] = student.number or index + 1
            comment_sheet[f"{comment_columns['english_name']}{comment_row}"] = student.english_name
            comment_sheet[f"{comment_columns['khmer_name']}{comment_row}"] = student.khmer_name
            comment_sheet[f"{comment_columns['student_id']}{comment_row}"] = student.student_id
            comment_sheet[f"{comment_columns['sex']}{comment_row}"] = student.sex
            comment_sheet[f"{comment_columns['group']}{comment_row}"] = student.group

        transferred_scores = transferred_comments = 0
        for upload in uploads:
            for index, student in enumerate(upload.students):
                row = first_target_row + index
                if student.total is not None:
                    quarter_sheet[f"{upload.mapping.target_score_column}{row}"] = student.total
                    transferred_scores += 1
                if upload.mapping.comment_columns:
                    for column, value in zip(upload.mapping.comment_columns, student.comments):
                        if value:
                            comment_sheet[f"{column}{first_comment_row + index}"] = value
                            transferred_comments += 1
                if upload.mapping.teacher_name_column and upload.teacher_name:
                    comment_sheet[f"{upload.mapping.teacher_name_column}{first_comment_row + index}"] = upload.teacher_name

        # openpyxl preserves formulas but cannot calculate their displayed
        # values. Ask Excel to recalculate all dependent FinalSem formulas
        # when the generated workbook is opened.
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcOnSave = True
        workbook.save(output_path)
        return transferred_scores, transferred_comments
    finally:
        workbook.close()


def report_values(
    student_index: int,
    uploads: list[SubjectUpload],
    academic_year: str,
    class_size: int,
    grade: int = 6,
) -> dict[str, str]:
    config = load_grade_config(grade)
    roster_student = uploads[0].students[student_index]
    uploads_by_key = {item.mapping.key: item for item in uploads}
    scores = [
        uploads_by_key[key].students[student_index].total if key in uploads_by_key else None
        for key in config.subject_score_keys
    ]
    grades = [config.letter_grade(score) for score in scores]
    present_scores = [score for score in scores if score is not None]
    overall = sum(present_scores) / len(present_scores) if present_scores else None
    ranking = []
    for row in range(class_size):
        row_scores = [item.students[row].total for item in uploads if item.students[row].total is not None]
        if row_scores:
            ranking.append((sum(row_scores) / len(row_scores), row))
    ranking.sort(key=lambda value: value[0], reverse=True)
    rank = next((position + 1 for position, (_, row) in enumerate(ranking) if row == student_index), "")
    earned_credits = sum(
        s.credit for s in config.subjects
        if scores[config.subject_score_keys.index(s.key)] is not None and s.key in uploads_by_key
    )
    gpa_values = [config.grade_point(g) for g in grades if g]
    gpa = sum(gpa_values) / len(gpa_values) if gpa_values else None

    cell_map = config.report_cell_map
    values: dict[str, str] = {
        cell_map.get("english_name", "F2"): roster_student.english_name,
        cell_map.get("sex", "F4"): roster_student.sex,
        cell_map.get("student_id", "F5"): roster_student.student_id,
        cell_map.get("khmer_name", "M_50"): roster_student.khmer_name,
        cell_map.get("academic_year", "ACADEMIC_YEAR"): academic_year,
    }

    # Per-subject score/grade/credit cells
    for idx, (key, score, grade_letter, subject) in enumerate(zip(
        config.subject_score_keys, scores, grades, config.subjects
    )):
        if score is not None:
            values[f"F{7 + idx}"] = f"{score:.2f}"
            values[f"F{7 + idx + 16}"] = grade_letter
            values[f"F{7 + idx + 32}"] = str(subject.credit)
        else:
            values[f"F{7 + idx}"] = ""
            values[f"F{7 + idx + 16}"] = ""
            values[f"F{7 + idx + 32}"] = ""

    # Khmer aggregate
    khmer_keys = config.subject_score_keys[:9] if len(config.subject_score_keys) >= 9 else config.subject_score_keys
    khmer_present = [s for s, k in zip(scores, config.subject_score_keys) if s is not None and k in khmer_keys]
    khmer_avg = sum(khmer_present) / len(khmer_present) if khmer_present else None
    values[cell_map.get("khmer_avg", "F90")] = f"{khmer_avg:.2f}" if khmer_avg is not None else ""
    values[cell_map.get("khmer_grade", "F91")] = config.letter_grade(khmer_avg) if khmer_avg is not None else ""
    khmer_credits = sum(s.credit for s in config.subjects if s.key in khmer_keys and s.key in uploads_by_key and uploads_by_key[s.key].students[student_index].total is not None)
    values[cell_map.get("khmer_credits", "F87")] = str(khmer_credits) if khmer_present else ""

    # English aggregate
    english_keys = config.subject_score_keys[9:18] if len(config.subject_score_keys) >= 18 else config.subject_score_keys[9:]
    english_present = [s for s, k in zip(scores, config.subject_score_keys) if s is not None and k in english_keys]
    english_avg = sum(english_present) / len(english_present) if english_present else None
    values[cell_map.get("english_avg", "F96")] = f"{english_avg:.2f}" if english_avg is not None else ""
    values[cell_map.get("english_grade", "F97")] = config.letter_grade(english_avg) if english_avg is not None else ""
    english_credits = sum(s.credit for s in config.subjects if s.key in english_keys and s.key in uploads_by_key and uploads_by_key[s.key].students[student_index].total is not None)
    values[cell_map.get("english_credits", "F93")] = str(english_credits) if english_present else ""

    # Totals
    values[cell_map.get("total_credits", "F99")] = str(earned_credits) if present_scores else ""
    values[cell_map.get("overall_avg", "F102")] = f"{overall:.2f}" if overall is not None else ""
    values[cell_map.get("overall_grade", "F103")] = config.letter_grade(overall) if overall is not None else ""
    values[cell_map.get("gpa", "F104")] = f"{gpa:.2f}" if gpa is not None else ""
    values[cell_map.get("rank", "F105")] = str(rank)
    values[cell_map.get("class_size", "F107")] = str(class_size)

    # Attendance
    for cell in cell_map.get("attendance", []):
        values[cell] = ""
    values[cell_map.get("total_school_day", "Total_School_Day")] = ""

    return values


def read_docx_template_parts(template_path: Path) -> list[tuple[zipfile.ZipInfo, bytes]]:
    """Read the report template once for an entire student batch."""
    with zipfile.ZipFile(template_path) as source:
        return [(info, source.read(info.filename)) for info in source.infolist()]


def fill_merge_fields(
    template_path: Path,
    output_path: Path,
    values: dict[str, str],
    period: str,
    template_parts: list[tuple[zipfile.ZipInfo, bytes]] | None = None,
) -> None:
    """Fill cached merge-field results without serializing/rebuilding Word XML."""
    quarter_names = {"Q1": "First", "Q2": "Second", "Q3": "Third", "Q4": "Fourth"}
    field_pattern = re.compile(
        rb'(<w:fldChar\b[^>]*w:fldCharType="begin"[^>]*/>.*?<w:instrText\b[^>]*>(?P<instruction>.*?)</w:instrText>.*?<w:fldChar\b[^>]*w:fldCharType="separate"[^>]*/>)(?P<result>.*?)(<w:fldChar\b[^>]*w:fldCharType="end"[^>]*/>)',
        re.DOTALL,
    )
    text_pattern = re.compile(rb'(<w:t\b[^>]*>)(.*?)(</w:t>)', re.DOTALL)

    def replace_field(match: re.Match[bytes]) -> bytes:
        instruction = match.group("instruction").decode("utf-8", errors="ignore")
        field = re.search(r'MERGEFIELD\s+"?([^"\\ ]+)', instruction)
        if not field:
            return match.group(0)
        replacement = escape(values.get(field.group(1), "")).encode("utf-8")
        result, changed = text_pattern.subn(lambda text: text.group(1) + replacement + text.group(3), match.group("result"), count=1)
        return match.group(1) + (result if changed else match.group("result")) + match.group(4)

    parts = template_parts or read_docx_template_parts(template_path)
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as destination:
        for info, original_content in parts:
            content = original_content
            if info.filename == "word/document.xml":
                content = field_pattern.sub(replace_field, content)
                content = content.replace(b">Fourth</w:t>", f">{quarter_names[period]}</w:t>".encode("utf-8"))
                content = content.replace(b">2025-2026</w:t>", f">{escape(values.get('ACADEMIC_YEAR', ''))}</w:t>".encode("utf-8"))
            destination.writestr(info, content)


def convert_word_reports_to_pdf(reports_dir: Path) -> None:
    """Convert one report at a time to keep the worker's peak memory bounded."""
    libreoffice = shutil.which("libreoffice") or shutil.which("soffice")
    if libreoffice:
        documents = sorted(reports_dir.glob("*.docx"))
        if not documents:
            raise ProcessingError("No student report documents were created.")
        # LibreOffice can retain every document supplied on one command line.
        # A separate temporary profile and process for each student has a much
        # lower peak RAM requirement on Render's small instances.
        for document in documents:
            profile_dir = Path(tempfile.mkdtemp(prefix="mlcis-libreoffice-"))
            command = [
                libreoffice,
                f"-env:UserInstallation=file://{profile_dir.as_uri()}",
                "--headless",
                "--nologo",
                "--nofirststartwizard",
                "--norestore",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", str(reports_dir),
                str(document),
            ]
            try:
                completed = subprocess.run(command, capture_output=True, text=True, timeout=60)
            finally:
                shutil.rmtree(profile_dir, ignore_errors=True)
            if completed.returncode != 0 or not document.with_suffix(".pdf").exists():
                detail = (completed.stderr or completed.stdout or "LibreOffice could not export the PDF.").strip()
                raise ProcessingError(f"PDF export failed for {document.name}: {detail[:260]}")
    elif os.name == "nt":
        script = Path(__file__).with_name("word_to_pdf.ps1")
        command = ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", str(script), str(reports_dir)]
        completed = subprocess.run(command, capture_output=True, text=True, timeout=180)
    else:
        raise ProcessingError(
            "PDF conversion is unavailable on this server. Student report cards are provided as a DOCX ZIP instead."
        )
    expected = list(reports_dir.glob("*.docx"))
    if not libreoffice and (completed.returncode != 0 or any(not document.with_suffix(".pdf").exists() for document in expected)):
        detail = (completed.stderr or completed.stdout or "Microsoft Word could not export the PDF.").strip()
        raise ProcessingError(f"PDF export failed: {detail[:300]}")


def safe_filename(value: str) -> str:
    return re.sub(r"[<>:\\/?*\"|]", "_", value).strip(" .") or "student"


def generate_report_cards(
    template_path: Path,
    job_dir: Path,
    uploads: list[SubjectUpload],
    academic_year: str,
    period: str,
    grade: int = 6,
) -> tuple[Path, int, str, str | None]:
    if not template_path.exists():
        raise ProcessingError("The report-card Word template file could not be found.")
    reports_dir = job_dir / "student-report-cards"
    reports_dir.mkdir(exist_ok=True)
    # The previous implementation repeatedly opened and decompressed the same
    # Word template once per student. Keep the original archive parts in memory
    # for this short-lived job instead.
    template_parts = read_docx_template_parts(template_path)
    for index, student in enumerate(uploads[0].students, start=1):
        stem = f"{index:02d}-{safe_filename(student.english_name)}"
        docx_path = reports_dir / f"{stem}.docx"
        values = report_values(index - 1, uploads, academic_year, len(uploads[0].students), grade)
        fill_merge_fields(template_path, docx_path, values, period, template_parts)
    # Generated DOCX files are now on disk. Releasing the cached template
    # parts before starting LibreOffice lowers concurrent Python + Office RAM.
    template_parts.clear()
    del template_parts
    gc.collect()
    report_format = "PDF"
    conversion_warning = None
    try:
        convert_word_reports_to_pdf(reports_dir)
    except ProcessingError as error:
        report_format = "DOCX"
        conversion_warning = str(error)
    else:
        for docx_path in reports_dir.glob("*.docx"):
            docx_path.unlink()
    config = load_grade_config(grade)
    archive_base = job_dir / f"Grade-{config.label.replace(' ', '')}-Student-Report-Cards"
    archive = Path(shutil.make_archive(str(archive_base), "zip", reports_dir))
    return archive, len(uploads[0].students), report_format, conversion_warning


def build_subject_uploads(path: Path, filename: str, grade: int = 6) -> list[SubjectUpload]:
    config = load_grade_config(grade)
    mapping_by_source = config.mapping_by_source_subject
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        if "SETTINGS" not in workbook.sheetnames:
            raise ProcessingError(f"{filename}: an uploaded file does not contain a SETTINGS sheet.")
        settings_sheet = workbook["SETTINGS"]
        subject_name = str(settings_sheet["F23"].value or "").strip()
        setting_teacher = str(settings_sheet["F20"].value or "").strip()
        period_value = str(settings_sheet["F14"].value or "").strip()
        normalised_period = period_value.upper().replace("QUARTER", "Q").replace(" ", "")
        period = normalised_period if normalised_period in config.periods else None
        sheet_names = set(workbook.sheetnames)
        mapping = mapping_by_source.get("".join(c for c in subject_name.casefold() if c.isalnum()))
        if mapping is not None:
            if "Scoring System" not in workbook.sheetnames:
                raise ProcessingError("An uploaded file does not contain a Scoring System sheet.")
            return [SubjectUpload(
                filename, subject_name, setting_teacher, mapping, period,
                _read_subject_students_sheet(workbook["Scoring System"]),
            )]

        # Multi-subject workbooks: read every detected sheet from this one
        # already-open ZIP archive. Before this, a 19-subject file was opened
        # and parsed up to 20 times, which is the main avoidable delay.
        detected_by_key = {
            mapping_by_source["".join(c for c in sheet_name.casefold() if c.isalnum())].key: sheet_name
            for sheet_name in sheet_names
            if "".join(c for c in sheet_name.casefold() if c.isalnum()) in mapping_by_source
        }
        detected = [
            (detected_by_key[mapping.key], mapping)
            for mapping in config.subjects
            if mapping.key in detected_by_key
        ]
        if not detected:
            raise ProcessingError(f"{filename}: Subject is blank in SETTINGS!F23 and no supported subject sheets were found.")
        results: list[SubjectUpload] = []
        for sheet_name, subject_mapping in detected:
            students, sheet_teacher = _read_multi_subject_sheet(workbook[sheet_name], sheet_name)
            results.append(SubjectUpload(filename, sheet_name, sheet_teacher or setting_teacher, subject_mapping, period, students))
        return results
    finally:
        workbook.close()


def print_stages(grade: int = 6) -> dict[str, dict[str, Any]]:
    config = load_grade_config(grade)
    return {
        "P1": {"period": "Q1", "label": "Print 1: Q1 and Teacher Comment", "report": True, "template": "quarter_report"},
        "P2": {"period": "Q2", "label": "Print 2: Q2, Exam 1, and Final SEM1", "report": False},
        "P3": {"period": "Q3", "label": "Print 3: Q3 and Teacher Comment", "report": True, "template": "quarter_report"},
        "P4": {"period": "Q4", "label": "Print 4: Q4, Exam 2, and Final SEM2", "report": False},
    }


def output_filename(grade: int = 6) -> str:
    config = load_grade_config(grade)
    extension = Path(config.scoring_template).suffix or ".xlsm"
    return f"G{config.grade} Scoring System Final 2024-2025-MLCIS{extension}"


def reports_filename(grade: int = 6) -> str:
    config = load_grade_config(grade)
    return f"Grade-{config.label.replace(' ', '')}-Student-Report-Cards.zip"


def save_report_context(job_dir: Path, uploads: list[SubjectUpload], academic_year: str, period: str, grade: int) -> None:
    """Keep only the short-lived information required after scoring review."""
    payload = {
        "academic_year": academic_year,
        "period": period,
        "grade": grade,
        "subjects": [
            {"key": upload.mapping.key, "label": upload.mapping.label, "source_subject": upload.subject_name,
             "students": [asdict(student) for student in upload.students]}
            for upload in uploads
        ],
    }
    (job_dir / "report-context.json").write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


def generate_reports_for_job(job_id: str, grade: int) -> dict[str, Any]:
    """Generate report cards only after the teacher has reviewed scoring."""
    if not re.fullmatch(r"[0-9a-f]{32}", job_id or ""):
        raise ProcessingError("Scoring review session was not found.")
    job_dir = Path(settings.MLCIS_JOB_DIR) / job_id
    context_path = job_dir / "report-context.json"
    if not context_path.exists() or not output_path(job_id, "scoring", grade).exists():
        raise ProcessingError("Scoring review session has expired. Import the gradebooks again.")
    context = json.loads(context_path.read_text(encoding="utf-8"))
    if context.get("grade") != grade:
        raise ProcessingError("The selected grade does not match this scoring review session.")
    config = load_grade_config(grade)
    uploads = [
        SubjectUpload(
            "", item["source_subject"], "", config.mapping_by_key[item["key"]], context["period"],
            [Student(**{**student, "comments": tuple(student["comments"])}) for student in item["students"]],
        )
        for item in context["subjects"]
    ]
    stage = print_stages(grade).get(next((key for key, value in print_stages(grade).items() if value["period"] == context["period"]), ""))
    if not stage or not stage["report"]:
        raise ProcessingError("This print package uses the Final Semester sheet after the manual exam; student cards are not generated here.")
    template = configured_templates(grade)[stage["template"]]
    if template is None:
        raise ProcessingError(f"{config.label} has no verified report-card template yet.")
    archive, count, report_format, warning = generate_report_cards(template, job_dir, uploads, context["academic_year"], context["period"], grade)
    return {
        "report_cards": {"available": True, "filename": archive.name, "count": count, "format": report_format,
                         "download": f"/api/download/{job_id}/report_cards/"},
        "warnings": [warning] if warning else [],
    }


def process_gradebooks(
    uploads,
    period: str,
    scoring_workbook=None,
    print_stage: str = "",
    grade: int = 6,
) -> dict[str, Any]:
    started_at = time.monotonic()
    purge_expired_jobs()
    config = load_grade_config(grade)
    if period not in config.periods:
        raise ProcessingError(f"Choose Quarter {', '.join(str(i) for i in range(1, len(config.periods) + 1))}.")
    if not uploads:
        raise ProcessingError("Choose one or more .xlsx E-gradebook files.")
    if any(not is_xlsx_upload(upload) for upload in uploads):
        raise ProcessingError("Every selected file must be a valid .xlsx E-gradebook.")
    stages = print_stages(grade)
    stage = stages.get(print_stage)
    if print_stage and stage is None:
        raise ProcessingError("Choose a valid print package.")
    if stage and stage["period"] != period:
        raise ProcessingError(f"{stage['label']} must import into {stage['period']}.")
    if stage and stage["period"] in {"Q2", "Q3", "Q4"} and scoring_workbook is None:
        raise ProcessingError(f"{stage['label']} requires the scoring .xlsm downloaded from the previous print package.")
    if scoring_workbook is not None and not is_scoring_workbook_upload(scoring_workbook):
        raise ProcessingError("The continued scoring workbook must be the downloaded .xlsm file.")
    templates = configured_templates(grade)
    ensure_scoring_template_is_usable(templates["scoring"], grade)

    job_id = uuid.uuid4().hex
    job_dir = Path(settings.MLCIS_JOB_DIR) / job_id
    job_dir.mkdir(parents=True, exist_ok=False)
    source_template = templates["scoring"]
    if scoring_workbook is not None:
        source_template = job_dir / "continued-scoring.xlsm"
        with source_template.open("wb+") as destination:
            for chunk in scoring_workbook.chunks():
                destination.write(chunk)
        try:
            ensure_scoring_template_is_usable(source_template, grade)
        except Exception:
            source_template.unlink(missing_ok=True)
            raise
    subject_uploads: list[SubjectUpload] = []
    uploads_by_subject: dict[str, SubjectUpload] = {}
    replacements: list[str] = []
    academic_year = ""
    for number, upload in enumerate(uploads, start=1):
        source_path = job_dir / f"source-{number}.xlsx"
        with source_path.open("wb+") as destination:
            for chunk in upload.chunks():
                destination.write(chunk)
        try:
            if not academic_year:
                academic_year = source_academic_year(source_path)
            for subject_upload in build_subject_uploads(source_path, upload.name, grade):
                if subject_upload.source_period and subject_upload.source_period != period:
                    raise ProcessingError(
                        f"{upload.name} / {subject_upload.mapping.label} is set to {subject_upload.source_period}; "
                        f"it cannot be imported into {period}."
                    )
                if subject_uploads and not roster_matches(subject_uploads[0].students, subject_upload.students):
                    raise ProcessingError(
                        f"{upload.name} / {subject_upload.mapping.label}: student roster does not match "
                        "the first detected subject. Use files for the same class only."
                    )
                previous = uploads_by_subject.get(subject_upload.mapping.key)
                if previous is not None:
                    subject_uploads[subject_uploads.index(previous)] = subject_upload
                    replacements.append(
                        f"{subject_upload.mapping.label}: {upload.name} replaced the earlier file {previous.filename}."
                    )
                else:
                    subject_uploads.append(subject_upload)
                uploads_by_subject[subject_upload.mapping.key] = subject_upload
        finally:
            source_path.unlink(missing_ok=True)

    scoring_output = job_dir / output_filename(grade)
    try:
        transferred_scores, transferred_comments = populate_scoring_template(
            source_template,
            scoring_output,
            period,
            subject_uploads,
            grade,
            clear_samples=scoring_workbook is None,
        )
    finally:
        if scoring_workbook is not None:
            source_template.unlink(missing_ok=True)
    with zipfile.ZipFile(scoring_output) as archive:
        vba_preserved = "xl/vbaProject.bin" in archive.namelist()
    if not vba_preserved:
        raise ProcessingError("Output validation failed: the copied scoring workbook has no VBA project.")

    first = subject_uploads[0]
    subjects_without_cached_totals = [
        item.mapping.label
        for item in subject_uploads
        if not any(student.total is not None for student in item.students)
    ]
    warnings = [
        ("The format workbook is used only as an output template. Sample data in all Q sheets is cleared."
         if scoring_workbook is None else
         "The uploaded scoring workbook is continued. Existing quarters, comments, and manual exam values are preserved."),
        f"Selected destination: {config.period_sheets[period]}; imported subjects: {', '.join(item.mapping.label for item in subject_uploads)}.",
        "All uploaded subject files must have the same student roster. Matching uses student ID when present, otherwise exact English name and order.",
        "Only cached numeric Total values are imported. Missing cached totals remain blank.",
        "Achievement, Attitude, Areas to improve, and Teacher's Name are imported into the matching Comment Q subject group when source comments are present.",
        "Existing #REF! formulas and external links in the format are preserved and not repaired.",
    ]
    if subjects_without_cached_totals:
        warnings.append(
            "No cached Total values were found for: "
            + ", ".join(subjects_without_cached_totals)
            + ". Open those source sheets in Excel, recalculate, save, then upload again."
        )
    warnings.extend(replacements)
    save_report_context(job_dir, subject_uploads, academic_year, period, grade)
    if stage and not stage["report"]:
        exam_label = "Exam 1" if period == "Q2" else "Exam 2"
        report_reason = f"Enter {exam_label} manually in the downloaded scoring workbook, then recalculate and print the Final Semester sheet."
    else:
        report_reason = "Review and download the scoring workbook first, then click Generate Student Cards."
    report_output = {"available": False, "reason": report_reason}
    return {
        "job_id": job_id,
        "print_stage": print_stage or None,
        "print_label": stage["label"] if stage else None,
        "period": period,
        "grade": grade,
        "grade_label": config.label,
        "student_count": len(first.students),
        "populated_students": len(first.students),
        "processing_seconds": round(time.monotonic() - started_at, 1),
        "subjects": [{"name": item.mapping.label, "source_subject": item.subject_name, "source_period": item.source_period} for item in subject_uploads],
        "warnings": warnings,
        "outputs": {
            "scoring": {"filename": scoring_output.name, "available": True, "download": f"/api/download/{job_id}/scoring/", "vba_preserved": vba_preserved, "fields_transferred": transferred_scores, "comments_transferred": transferred_comments},
            "report_cards": report_output,
        },
    }


def source_academic_year(path: Path) -> str:
    return _source_setting(path, "F11")


def output_path(job_id: str, kind: str, grade: int = 6) -> Path:
    if kind == "scoring":
        filename = output_filename(grade)
    elif kind == "report_cards":
        filename = reports_filename(grade)
    else:
        raise ProcessingError("Requested output was not found.")
    if not job_id.isalnum():
        raise ProcessingError("Requested output was not found.")
    path = Path(settings.MLCIS_JOB_DIR) / job_id / filename
    if not path.exists():
        raise ProcessingError("Requested output was not found.")
    return path


def consume_output(job_id: str, kind: str, grade: int = 6) -> tuple[bytes, str]:
    """Keep scoring temporarily for review; erase cards and job after delivery."""
    purge_expired_jobs()
    path = output_path(job_id, kind, grade)
    content = path.read_bytes()
    filename = path.name
    if kind == "report_cards":
        path.unlink(missing_ok=True)
        shutil.rmtree(path.parent / "student-report-cards", ignore_errors=True)
        shutil.rmtree(path.parent, ignore_errors=True)
    return content, filename
