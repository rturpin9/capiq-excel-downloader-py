"""Tests for identifier result filtering."""
import pandas as pd
import pytest

from capiq_excel.ids import _get_ids_from_csv_path


def test_get_ids_filters_failed_rows(tmp_path):
    path = tmp_path / "ids.csv"
    pd.DataFrame({"IQID": ["IQ1", None, "#INVALID COMPANY ID", "IQ2"]}).to_csv(
        path,
        index=False,
    )
    assert _get_ids_from_csv_path(str(path)) == ["IQ1", "IQ2"]


def test_get_ids_raises_when_none_resolve(tmp_path):
    path = tmp_path / "ids.csv"
    pd.DataFrame({"IQID": [None, "#ERROR"]}).to_csv(path, index=False)
    with pytest.raises(ValueError, match="No identifiers resolved"):
        _get_ids_from_csv_path(str(path))
