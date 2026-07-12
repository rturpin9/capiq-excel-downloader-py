"""Tests for safe restart cleanup of generated workbooks."""
from capiq_excel.fileops import clear_generated_xlsx_files


def test_restart_cleanup_only_removes_top_level_xlsx(tmp_path):
    workbook = tmp_path / "old.xlsx"
    keep = tmp_path / "notes.txt"
    nested = tmp_path / "failed" / "failed.xlsx"
    nested.parent.mkdir()
    workbook.write_bytes(b"old")
    keep.write_text("keep", encoding="utf-8")
    nested.write_bytes(b"failed")

    clear_generated_xlsx_files(str(tmp_path))

    assert not workbook.exists()
    assert keep.exists()
    assert nested.exists()
