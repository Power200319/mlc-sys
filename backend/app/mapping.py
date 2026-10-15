"""Verified transfer map for subject-specific Grade 6 E-gradebooks.

The scoring-system workbook is a format. An E-gradebook upload contains one
subject, and the selected quarter plus subject determines the destination.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class SubjectMapping:
    key: str
    label: str
    source_subject: str
    target_score_column: str
    target_letter_column: str
    comment_columns: tuple[str, str, str] | None = None
    teacher_name_column: str | None = None
    status: str = "VERIFIED"


# Every subject E-gradebook has the same inspected Scoring System layout:
# AM is Total. The target score columns are read from the supplied scoring
# format. English comments are additionally verified as AP:AR -> G:I.
SUBJECT_MAPPINGS = (
    SubjectMapping("khmer_dictation", "Khmer Dictation", "Khmer Dictation", "G", "W", ("AQ", "AR", "AS"), "AT"),
    SubjectMapping("khmer_essays", "Khmer Essays", "Khmer Essays", "H", "X", ("AU", "AV", "AW"), "AX"),
    SubjectMapping("khmer_reading", "Khmer Reading", "Khmer Reading", "I", "Y", ("AY", "AZ", "BA"), "BB"),
    # The format labels this fourth Khmer comment group "Khmer Calligraphy".
    SubjectMapping("khmer_public_speaking", "Khmer Public Speaking", "Khmer Public Speaking", "J", "Z", ("BC", "BD", "BE"), "BF"),
    SubjectMapping("khmer_math", "Khmer Math", "Khmer Math", "K", "AA", ("BG", "BH", "BI"), "BJ"),
    SubjectMapping("khmer_social_studies", "Khmer Social Studies", "Khmer Social Studies", "L", "AB", ("BK", "BL", "BM"), "BN"),
    SubjectMapping("khmer_science", "Khmer Science", "Khmer Science", "M", "AC", ("BO", "BP", "BQ"), "BR"),
    SubjectMapping("sport", "Sport", "Sport", "N", "AD", ("BS", "BT", "BU"), "BV"),
    SubjectMapping("health_education", "Health Education", "Health Education", "O", "AE", ("BW", "BX", "BY"), "BZ"),
    SubjectMapping("english", "English", "English", "P", "AF", ("G", "H", "I"), "J"),
    SubjectMapping("english_grammar", "English Grammar", "English Grammar", "Q", "AG", ("AI", "AJ", "AK"), "AL"),
    SubjectMapping("english_math", "English Math", "English Math", "R", "AH", ("K", "L", "M"), "N"),
    SubjectMapping("english_science", "English Science", "English Science", "S", "AI", ("O", "P", "Q"), "R"),
    SubjectMapping("chinese", "Chinese", "Chinese", "T", "AJ", ("AA", "AB", "AC"), "AD"),
    SubjectMapping("robothink", "RoboThink", "RoboThink", "U", "AK", ("S", "T", "U"), "V"),
    SubjectMapping("computer", "Computer", "Computer", "V", "AL", ("AE", "AF", "AG"), "AH"),
)

MAPPING_BY_KEY = {mapping.key: mapping for mapping in SUBJECT_MAPPINGS}


def normalise_subject(value: str) -> str:
    """Match harmless spacing, hyphen and plural variations from SETTINGS."""
    return "".join(character for character in value.casefold() if character.isalnum())


MAPPING_BY_SOURCE_SUBJECT = {
    normalise_subject(mapping.source_subject): mapping for mapping in SUBJECT_MAPPINGS
}
# Known spelling variants used by subject E-gradebook Settings sheets.
MAPPING_BY_SOURCE_SUBJECT.update({
    "khmerdictations": MAPPING_BY_KEY["khmer_dictation"],
    "khmermathematics": MAPPING_BY_KEY["khmer_math"],
    "englishmathematics": MAPPING_BY_KEY["english_math"],
})

# Multi-subject workbooks do not set SETTINGS!F23. Their subject identity is
# carried by worksheet names instead. These are the exact supplied tab names.
MULTI_SUBJECT_SHEETS = {
    "Khmer Dictations": MAPPING_BY_KEY["khmer_dictation"],
    "Khmer Essays": MAPPING_BY_KEY["khmer_essays"],
    "Reading": MAPPING_BY_KEY["khmer_reading"],
    "Public Speaking": MAPPING_BY_KEY["khmer_public_speaking"],
    "KH-Math": MAPPING_BY_KEY["khmer_math"],
    "Kh-Social Studies": MAPPING_BY_KEY["khmer_social_studies"],
    "KH-Science": MAPPING_BY_KEY["khmer_science"],
    "SPORT": MAPPING_BY_KEY["sport"],
    "Health Education ": MAPPING_BY_KEY["health_education"],
}
PERIOD_SHEETS = {"Q1": "Q1", "Q2": "Q2", "Q3": "Q3", "Q4": "Q4"}
COMMENT_SHEETS = {"Q1": "Comment Q1", "Q2": "Comment Q2", "Q3": "Comment Q3", "Q4": "Comment Q4"}
SUPPORTED_PERIODS = tuple(PERIOD_SHEETS)
