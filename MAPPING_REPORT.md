# Batch Subject Import Mapping

The scoring workbook is an output format. Its sample roster and scores are never used as import data.

| Source E-gradebook value | Dynamic condition | Output destination |
| --- | --- | --- |
| `SETTINGS!F23` | Detect subject | Select subject column below |
| `Scoring System!AM[r]` cached Total | Source header must be `AM9 = Total` | Selected `Q#` subject score column |
| `Scoring System!AP:AR[r]` | Subject = English only | Selected `Comment Q#!G:I` |

## Score columns in every selected Quarter

| Detected subject | Column |
| --- | --- |
| Khmer Dictation, Khmer Essays, Khmer Reading, Khmer Public Speaking | G, H, I, J |
| Khmer Math, Khmer Social Studies, Khmer Science, Sport, Health Education | K, L, M, N, O |
| English, English Grammar, English Math, English Science | P, Q, R, S |
| Chinese, RoboThink, Computer | T, U, V |

The batch is rejected if an uploaded file has an unsupported subject, a missing `Scoring System` / Total header, a duplicate subject, or a different student roster. A selected Quarter controls output even if an uploaded Settings sheet reports a different quarter.

VBA, formulas, styles, merged cells, existing broken references, and external links are preserved. No broken formulas are repaired.
