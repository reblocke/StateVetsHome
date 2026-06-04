# StateVetsHome

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Citation metadata](https://img.shields.io/badge/CITATION.cff-validatable-blue.svg)](CITATION.cff)

Legacy Python analysis for a COVID-19 goals-of-care cohort from a state veterans home outbreak. The workflow summarizes admission characteristics, advance-care-planning variables, patient versus surrogate decision-maker groups, code-status transitions, and hospitalization outcomes.

## Project Status

This is an abstract-only legacy analysis repository. A verifiable abstract DOI or indexed abstract record was not recovered in the current public repository, CV, or website materials during this cleanup pass. Until that record is supplied, cite this repository as research software using `CITATION.cff`.

The patient-level source workbook is restricted clinical data and is not part of this public repository.

## Authors, Funding, And COI

Maintainer: Brian W. Locke, MD, MSCI (`@reblocke`, ORCID `0000-0002-3588-5238`).

Additional abstract authors, affiliations, funding, and conflicts of interest were not verified from current public repository materials. Add them here and in `CITATION.cff` if a citable abstract record is recovered.

## Data Access

Expected local-only input:

```text
data/private/WorkingDb.xls
```

The workbook appears to contain row-level clinical and goals-of-care data. Do not commit source data, derived row-level exports, PHI, private workbooks, or local collaborator files. The public repository documents the expected schema and legacy workflow only.

## Quick Start

```bash
python -m pip install -r requirements.txt
python SVH.py --input data/private/WorkingDb.xls --output-dir outputs/legacy-python
```

Generated outputs are written under the selected output directory:

- `output.xlsx`: cleaned analysis table
- `tables.xlsx`: legacy aggregate summary tables
- `statistical_tests.txt`: printed statistical test output
- `code_status_alluvial.html`: manually specified legacy alluvial/Sankey diagram
- `figures/*.png`: generated distribution plots

## Repository Layout

| Path | Purpose |
| --- | --- |
| `SVH.py` | Legacy Python workflow for loading, cleaning, summarizing, testing, and plotting the cohort. |
| `data_dictionary.md`, `data_dictionary.csv` | Human-readable and machine-readable documentation of expected workbook columns, derived fields, and outputs. |
| `REPRODUCING.md` | Reproduction guide for running the workflow with a compatible local workbook. |
| `llms.txt` | Concise machine-readable index for LLM and search-agent use. |
| `CITATION.cff` | Structured repository-software citation metadata. |
| `tests/` | No-PHI synthetic-workbook smoke tests. |

Historical aggregate outputs that were previously tracked in the repository root and `dist figs/` are archived in the GitHub release `legacy-covid-goc-outputs-2026-06-04`.

## Dependencies

| Dependency | Use |
| --- | --- |
| Python 3.10+ | Runtime used for the cleaned legacy workflow. |
| pandas, openpyxl, xlrd | Workbook loading and Excel export. |
| matplotlib, seaborn | Static distribution plots. |
| scipy | Legacy t tests and Fisher exact tests. |
| plotly | Code-status alluvial/Sankey HTML output. |
| pytest | No-PHI smoke tests. |

## Citation

Use `CITATION.cff` for the repository citation. If a citable abstract record is later supplied, add it as `preferred-citation` and update this README and `llms.txt`.

## License

Repository code is released under the MIT License. Restricted clinical data, third-party content, private drafts, and publisher-formatted text are excluded.

## Contact

Use GitHub issues or pull requests for repository-specific questions.
