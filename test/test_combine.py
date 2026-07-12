"""Regression tests for restart-safe workbook combination."""
import pandas as pd

from capiq_excel.combine import combine_all_capiq_xlsx


def test_restart_replaces_output_instead_of_appending(tmp_path):
    workbook_dir = tmp_path / "workbooks"
    workbook_dir.mkdir()
    pd.DataFrame({"Date": ["2026-01-01"], "Revenue": [10]}).to_excel(
        workbook_dir / "IQ1.xlsx",
        index=False,
    )
    output = tmp_path / "combined.csv"
    pd.DataFrame({"Date": ["old"], "Revenue": [-1]}).to_csv(output, index=False)

    combine_all_capiq_xlsx(str(workbook_dir), str(output), restart=True, num_parts=1)
    first = pd.read_csv(output)
    assert len(first) == 1
    assert first.loc[0, "Revenue"] == 10
    assert first.loc[0, "CQID"] == "IQ1"

    combine_all_capiq_xlsx(str(workbook_dir), str(output), restart=True, num_parts=1)
    second = pd.read_csv(output)
    assert len(second) == 1
