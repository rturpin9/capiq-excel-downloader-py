"""Pure regression tests for chart formulas and dated series output."""
import csv

import pandas as pd

from capiq_excel.engines.chart import (
    _build_market_formula,
    _build_multiple_formula,
    _date_from_expanded_formula,
    _metric_series,
    _write_raw_chart_csv,
)


def test_chart_formulas_include_currency():
    market = _build_market_formula(
        '"DSGX"', "IQ_CLOSEPRICE", "1/1/2026", "2/1/2026", "Price", "CAD"
    )
    multiple = _build_multiple_formula(
        '"DSGX"', "IQ_TEV_EBITDA", "IQ_LTM",
        "1/1/2026", "2/1/2026", "EV/EBITDA", "USD",
    )
    assert "Options: Curr=CAD" in market
    assert "Options: Curr=USD" in multiple
    assert multiple.startswith("=CIQRANGE(")


def test_expanded_formula_date_drives_series_and_csv(tmp_path):
    formula = '=CIQ("DSGX","IQ_CLOSEPRICE","1/15/2026")'
    assert _date_from_expanded_formula(formula) == pd.Timestamp("2026-01-15")

    data = {
        "DSGX": {
            "_num_points": 2,
            "_metric_dates": {
                "IQ_CLOSEPRICE": [
                    pd.Timestamp("2026-01-15"),
                    pd.Timestamp("2026-01-20"),
                ],
            },
            "dates": [pd.Timestamp("2026-01-15"), pd.Timestamp("2026-01-20")],
            "IQ_CLOSEPRICE": [100.0, 101.0],
        }
    }
    series = _metric_series(data["DSGX"], "IQ_CLOSEPRICE")
    assert list(series.index) == [pd.Timestamp("2026-01-15"), pd.Timestamp("2026-01-20")]

    path = tmp_path / "raw.csv"
    _write_raw_chart_csv(data, ["IQ_CLOSEPRICE"], str(path))
    with path.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.reader(handle))
    assert rows[1][1] == "2026-01-15"
    assert rows[2][1] == "2026-01-20"
