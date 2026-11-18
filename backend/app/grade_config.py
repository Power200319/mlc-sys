"""Grade configuration loader for MLCIS Gradebook Converter.

Loads grade-specific subject mappings, column mappings, credits, and report
cell maps from JSON config files in backend/config/grades/.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

CONFIG_DIR = Path(__file__).resolve().parent.parent / "config" / "grades"
GROUP_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config" / "grade_groups"


@dataclass(frozen=True)
class SubjectMapping:
    key: str
    label: str
    source_subject: str
    target_score_column: str
    target_letter_column: str
    source_aliases: tuple[str, ...] = ()
    comment_columns: tuple[str, str, str] | None = None
    teacher_name_column: str | None = None
    credit: int = 1
    status: str = "VERIFIED"


@dataclass(frozen=True)
class GradeConfig:
    grade: int
    label: str
    scoring_template: str
    report_template: str | None
    semester_template: str | None = None
    end_year_template: str | None = None
    periods: tuple[str, ...] = ("Q1", "Q2", "Q3", "Q4")
    period_sheets: dict[str, str] = field(default_factory=dict)
    comment_sheets: dict[str, str] = field(default_factory=dict)
    subjects: tuple[SubjectMapping, ...] = ()
    subject_score_keys: tuple[str, ...] = ()
    grading_scale: tuple[dict[str, Any], ...] = ()
    report_cell_map: dict[str, Any] = field(default_factory=dict)

    @property
    def mapping_by_key(self) -> dict[str, SubjectMapping]:
        return {m.key: m for m in self.subjects}

    @property
    def mapping_by_source_subject(self) -> dict[str, SubjectMapping]:
        result: dict[str, SubjectMapping] = {}
        for m in self.subjects:
            for source_name in (m.source_subject, *m.source_aliases):
                key = "".join(c for c in source_name.casefold() if c.isalnum())
                result[key] = m
        return result

    def get_subject_by_key(self, key: str) -> SubjectMapping | None:
        return self.mapping_by_key.get(key)

    def get_subject_by_source(self, source_subject: str) -> SubjectMapping | None:
        normalised = "".join(c for c in source_subject.casefold() if c.isalnum())
        return self.mapping_by_source_subject.get(normalised)

    def letter_grade(self, score: float | None) -> str:
        if score is None:
            return ""
        for entry in sorted(self.grading_scale, key=lambda e: e["min"], reverse=True):
            if score >= entry["min"]:
                return entry["grade"]
        return "F"

    def grade_point(self, grade: str) -> float:
        for entry in self.grading_scale:
            if entry["grade"] == grade:
                return entry["point"]
        return 0.0


def load_grade_config(grade: int) -> GradeConfig:
    """Load grade configuration from JSON file."""
    if 1 <= grade <= 3:
        config_path = GROUP_CONFIG_DIR / "g1-g3.json"
    elif 4 <= grade <= 6:
        config_path = GROUP_CONFIG_DIR / "g4-g6.json"
    elif 7 <= grade <= 9:
        config_path = GROUP_CONFIG_DIR / "g7-g9.json"
    else:
        raise ValueError("Only Grades 1 through 9 are currently supported")
    if not config_path.exists():
        raise ValueError(f"No configuration found for grade {grade}")

    with config_path.open("r", encoding="utf-8") as f:
        data = json.load(f)

    subjects = tuple(
        SubjectMapping(
            key=s["key"],
            label=s["label"],
            source_subject=s["source_subject"],
            source_aliases=tuple(s.get("source_aliases", [])),
            target_score_column=s["target_score_column"],
            target_letter_column=s["target_letter_column"],
            comment_columns=tuple(s.get("comment_columns", ())) or None,
            teacher_name_column=s.get("teacher_name_column"),
            credit=s.get("credit", 1),
        )
        for s in data.get("subjects", [])
    )

    return GradeConfig(
        grade=grade,
        label=f"Grade {grade} ({data['label']})",
        scoring_template=data["scoring_template"],
        report_template=data["report_template"],
        semester_template=data.get("semester_template"),
        end_year_template=data.get("end_year_template"),
        periods=tuple(data.get("periods", ["Q1", "Q2", "Q3", "Q4"])),
        period_sheets=data.get("period_sheets", {}),
        comment_sheets=data.get("comment_sheets", {}),
        subjects=subjects,
        subject_score_keys=tuple(data.get("subject_score_keys", [])),
        grading_scale=tuple(data.get("grading_scale", [])),
        report_cell_map=data.get("report_cell_map", {}),
    )


def get_available_grades() -> list[int]:
    """Return grades backed by an approved group template."""
    return list(range(1, 10)) if GROUP_CONFIG_DIR.exists() else []


def get_grade_label(grade: int) -> str:
    """Return human-readable label for a grade."""
    try:
        config = load_grade_config(grade)
        return config.label
    except (ValueError, FileNotFoundError):
        return f"Grade {grade}"
