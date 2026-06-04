# Data Dictionary

This dictionary documents the expected restricted workbook schema used by `SVH.py`. Definitions are inferred from the legacy code unless marked reviewed. The source workbook is restricted row-level clinical data and must not be committed.

## Source Workbook

| Item | Value |
| --- | --- |
| Expected local path | `data/private/WorkingDb.xls` |
| Sheet | first worksheet |
| Header row | first row |
| Index column | first column |
| Unit of observation | one hospitalization or cohort record |
| Data status | restricted clinical/goals-of-care data |

## Key Domains

| Domain | Variables |
| --- | --- |
| Demographics | `Age`, `Gender`, `BMI`, `Ethnicity` |
| Hospital course | `Setting`, `Oxygen Delivery`, `New Discharge O2`, `LOS`, `Death`, `Palliative Consult` |
| Advance care planning | `Prior ACP type`, `Prior Decision Maker`, `Prior Code status`, `Current Decision Maker`, `Code Status At Hospitalization`, `Comfort care`, `ICU transfer acceptable to patient?` |
| Symptoms and signs | `Symptoms prior to admit`, `Fever`, `SOB`, `Cough`, `Temp`, `SBP`, `DBP`, `Pulse`, `RR`, `O2`, `WBC`, `SIRS criteria met` |
| Comorbidities | `CCI`, `MI`, `CHF`, `PVD`, `CVA or TIA`, `Dementia`, `COPD`, `DM`, `Mod-Sev CKD`, `AIDS` |
| Derived fields | `Days after admit test positive` |
| Generated outputs | `output.xlsx`, `tables.xlsx`, `statistical_tests.txt`, `code_status_alluvial.html`, `figures/*.png` |

## Review Notes

- Most definitions are inferred from variable names and recoding logic in `SVH.py`.
- The alluvial diagram uses manually embedded historical counts and is not dynamically recalculated from the supplied workbook.
- Output files may contain row-level or small-cell derived information and should remain ignored unless intentionally reviewed for public release.
- The CSV version contains one row per source variable, derived variable, and generated artifact.
