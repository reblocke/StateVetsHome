# Changelog

## 2026-06-04

- Reframed the repository as an abstract-only legacy COVID goals-of-care analysis with no verified abstract DOI currently available.
- Added machine-readable documentation, citation metadata, data dictionaries, dependency files, and no-PHI smoke tests.
- Replaced the hard-coded local workbook path with `--input` and `--output-dir` arguments.
- Routed generated tables, figures, statistical-test logs, and alluvial HTML under the selected output directory.
- Replaced interactive plotting with file-based output suitable for local and CI smoke tests.
- Archived historical aggregate generated outputs in release `legacy-covid-goc-outputs-2026-06-04` before removing them from the branch tip.
- Fixed the Python identity comparison `num_patients is not 0` to `num_patients != 0`.
