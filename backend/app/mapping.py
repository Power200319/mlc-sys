"""Verified transfer map for subject E-gradebooks.

Supports multiple grades via grade configuration files in backend/config/grades/.
Each grade defines its own subjects, column mappings, credits, and report cell maps.
"""

from __future__ import annotations

from dataclasses import dataclass

from .grade_config import SubjectMapping, get_available_grades, get_grade_label, load_grade_config

# Backward-compatible default: Grade 6 mappings
DEFAULT_GRADE = 6


def get_subject_mappings(grade: int = DEFAULT_GRADE) -> tuple[SubjectMapping, ...]:
    """Return subject mappings for a given grade."""
    config = load_grade_config(grade)
    return config.subjects


def get_mapping_by_key(grade: int = DEFAULT_GRADE) -> dict[str, SubjectMapping]:
    """Return subject mappings keyed by subject key for a given grade."""
    config = load_grade_config(grade)
    return config.mapping_by_key


def get_mapping_by_source_subject(grade: int = DEFAULT_GRADE) -> dict[str, SubjectMapping]:
    """Return subject mappings keyed by normalised source subject name."""
    config = load_grade_config(grade)
    return config.mapping_by_source_subject


def get_period_sheets(grade: int = DEFAULT_GRADE) -> dict[str, str]:
    """Return period-to-sheet mapping for a given grade."""
    config = load_grade_config(grade)
    return config.period_sheets


def get_comment_sheets(grade: int = DEFAULT_GRADE) -> dict[str, str]:
    """Return comment period-to-sheet mapping for a given grade."""
    config = load_grade_config(grade)
    return config.comment_sheets


def get_supported_periods(grade: int = DEFAULT_GRADE) -> tuple[str, ...]:
    """Return supported periods for a given grade."""
    config = load_grade_config(grade)
    return config.periods


def normalise_subject(value: str) -> str:
    """Match harmless spacing, hyphen and plural variations from SETTINGS."""
    return "".join(character for character in value.casefold() if character.isalnum())


# Export commonly used names for backward compatibility
SUBJECT_MAPPINGS = get_subject_mappings(DEFAULT_GRADE)
MAPPING_BY_KEY = get_mapping_by_key(DEFAULT_GRADE)
MAPPING_BY_SOURCE_SUBJECT = get_mapping_by_source_subject(DEFAULT_GRADE)
PERIOD_SHEETS = get_period_sheets(DEFAULT_GRADE)
COMMENT_SHEETS = get_comment_sheets(DEFAULT_GRADE)
SUPPORTED_PERIODS = get_supported_periods(DEFAULT_GRADE)