# AGENTS

## Project Purpose

This public repository preserves a legacy Python analysis for COVID-19 goals-of-care decisions in a state veterans home cohort. It is currently framed as an abstract-only repository with no verified DOI or indexed abstract record in the public repo materials.

## Data And Publication Safety

- Treat the source workbook as restricted row-level clinical data.
- Do not commit PHI, private workbooks, derived row-level exports, collaborator drafts, credentials, or local paths.
- Do not add publisher-formatted text or third-party PDFs. Link verified public records instead.
- If an abstract DOI or official abstract record is recovered, update `README.md`, `llms.txt`, and `CITATION.cff`; do not invent metadata.

## Workflow

Run from the repository root:

```bash
python SVH.py --input data/private/WorkingDb.xls --output-dir outputs/legacy-python
```

Generated tables, logs, figures, and alluvial HTML belong under ignored `outputs/`.

## Change Discipline

- Keep scientific logic changes narrow.
- Preserve the legacy patient-versus-surrogate grouping and manually specified alluvial counts unless a separate scientific review updates them.
- Keep `data_dictionary.csv` synchronized with workbook columns and derived/output artifacts.
- Use synthetic or de-identified fixtures only for tests.

## Verification

Before publishing changes, run:

```bash
python -m pytest
python SVH.py --input tests/fixtures/synthetic_svh_workbook.xlsx --output-dir /tmp/state-vets-home-smoke
git diff --check
```

Also validate `CITATION.cff` after citation edits and scan for hard-coded local paths or generated artifacts in the tracked tree.
