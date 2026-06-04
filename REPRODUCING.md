# Reproducing The Legacy Analysis

This repository does not include the restricted source workbook. To rerun the workflow, place a compatible local workbook at:

```text
data/private/WorkingDb.xls
```

Install dependencies and run:

```bash
python -m pip install -r requirements.txt
python SVH.py --input data/private/WorkingDb.xls --output-dir outputs/legacy-python
```

The workflow expects the workbook schema documented in `data_dictionary.md` and `data_dictionary.csv`. Generated outputs are local artifacts and should not be committed unless intentionally reviewed and released as aggregate historical outputs.

For a no-PHI smoke test:

```bash
python -m pip install -r requirements.txt -r requirements-dev.txt
python -m pytest
python SVH.py --input tests/fixtures/synthetic_svh_workbook.xlsx --output-dir /tmp/state-vets-home-smoke
```

The alluvial diagram is a legacy manually specified figure. It is saved as HTML for portability, but its embedded counts should not be interpreted as dynamically recalculated from the supplied workbook.
