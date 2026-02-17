import json
import os
import re

import pytest

from parsers.normalize_source import normalize_source, save_downloaded_file
from parsers.raw_models import RawAttendance
from parsers.raw_to_local import raw_to_local_df


def test_save_downloaded_file_naming(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"dummy")

    saved_path = save_downloaded_file(str(source), "1.26", "xlsx")

    assert os.path.exists(saved_path)
    assert os.path.isdir(tmp_path / "downloads" / "1.26")
    assert re.match(
        r".*downloads/1\.26/\d{2}\.\d{2}\.\d{2}(?:_\d+)?\.xlsx$",
        saved_path,
    )


def test_raw_to_local_ignores_service_columns():
    raw = RawAttendance(
        meta={"month": 1, "year": 2026, "warnings": []},
        rows=[
            {
                "תאריך": "02",
                "כניסה": "07:00",
                "יציאה": "15:00",
                "ת.כניסה": "00:01",
                "ת.יציאה": "00:01",
            }
        ],
    )
    df = raw_to_local_df(raw)
    assert df.loc[0, "כניסה"] == "07:00"
    assert df.loc[0, "יציאה"] == "15:00"


def test_write_local_xlsx_atomic_no_overwrite_on_failure(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    out_path = tmp_path / "local_data.xlsx"
    out_path.write_bytes(b"original")
    before_stat = out_path.stat()

    bad_pdf = tmp_path / "bad.pdf"
    bad_pdf.write_bytes(b"%PDF-1.4\n%")

    def fake_parse_pdf(_path: str):
        return RawAttendance(meta={"warnings": []}, rows=[{"תאריך": "", "כניסה": "", "יציאה": ""}])

    monkeypatch.setattr("parsers.normalize_source.parse_pdf_ylm", fake_parse_pdf)

    with pytest.raises(RuntimeError):
        normalize_source(str(bad_pdf), out_path=str(out_path))

    after_stat = out_path.stat()
    assert before_stat.st_size == after_stat.st_size
    assert before_stat.st_mtime == after_stat.st_mtime


def test_raw_saved_in_debug(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    source = tmp_path / "source.xlsx"
    source.write_bytes(b"dummy")

    raw = RawAttendance(
        meta={"month": 1, "year": 2026, "warnings": []},
        rows=[{"תאריך": "01", "כניסה": "08:00", "יציאה": "16:00"}],
    )

    def fake_parse_excel(_path: str):
        return raw

    monkeypatch.setattr("parsers.normalize_source.parse_excel_ylm", fake_parse_excel)
    monkeypatch.setenv("DEBUG_SAVE_RAW", "1")

    normalize_source(str(source), out_path=str(tmp_path / "local_data.xlsx"))

    raw_path = tmp_path / "source.xlsx.raw.json"
    assert raw_path.exists()
    payload = json.loads(raw_path.read_text(encoding="utf-8"))
    assert "meta" in payload and "rows" in payload
