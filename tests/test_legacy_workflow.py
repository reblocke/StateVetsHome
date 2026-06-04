from pathlib import Path

import pandas as pd

import SVH


FIXTURE = Path(__file__).parent / "fixtures" / "synthetic_svh_workbook.xlsx"


def test_file_to_df_loads_expected_columns_and_categories():
    db = SVH.file_to_df(FIXTURE)

    assert len(db) == 4
    assert "Days after admit test positive" in db.columns
    assert db["Days after admit test positive"].tolist() == [1, 2, 3, 4]
    assert str(db["Current Decision Maker"].dtype) == "category"
    assert list(db["ICU transfer acceptable to patient?"].cat.categories) == ["No", "Yes"]
    assert set(db["Fever"].astype(str)) == {"No", "Yes"}


def test_primary_grouping_has_patient_and_surrogate_records():
    db = SVH.file_to_df(FIXTURE)

    patient_dec = db.loc[db["Current Decision Maker"] == "Patient"]
    surrogate_dec = db.loc[db["Current Decision Maker"] != "Patient"]

    assert len(patient_dec) == 2
    assert len(surrogate_dec) == 2


def test_run_analysis_writes_expected_outputs(tmp_path):
    SVH.run_analysis(FIXTURE, tmp_path)

    assert (tmp_path / "output.xlsx").exists()
    assert (tmp_path / "tables.xlsx").exists()
    assert (tmp_path / "statistical_tests.txt").exists()
    assert (tmp_path / "code_status_alluvial.html").exists()
    assert (tmp_path / "figures" / "Display Dist Age.png").exists()
    assert (tmp_path / "figures" / "Display Cat Current Decision Maker.png").exists()

    exported = pd.read_excel(tmp_path / "output.xlsx", index_col=0)
    assert "Days after admit test positive" in exported.columns
