import sys


def test_source_file_priority_over_skip_download(tmp_path, monkeypatch):
    import run

    source_file = tmp_path / "source.pdf"
    source_file.write_bytes(b"%PDF-1.4\n%")

    calls = {}

    def fake_normalize(source_path, out_path="local_data.xlsx"):
        calls["source"] = source_path
        return out_path

    def fake_load_config():
        return {
            "SOURCE_FILE": str(source_file),
            "SKIP_DOWNLOAD": True,
            "EXCEL_PATH": "local_data.xlsx",
            "SITE_USERNAME": "user",
            "SITE_PASSWORD": "pass",
            "GSHEET_ID": "gsheet",
            "GOOGLE_JSON_FILE": "service_key.json",
            "HEADLESS": True,
            "MANUAL_PORTAL": False,
            "MANUAL_DOWNLOAD_TIMEOUT_MS": 0,
        }

    monkeypatch.setattr(run, "normalize_source", fake_normalize)
    monkeypatch.setattr(run, "load_config", fake_load_config)
    monkeypatch.setattr(run, "open_spreadsheet", lambda gsheet_id, google_json_file: object())
    monkeypatch.setattr(run, "get_worksheet", lambda spreadsheet, sheet_name: object())

    def fake_build_changes_sheet(spreadsheet, base_ws, sheet_name, excel_path):
        calls["excel_path"] = excel_path
        return False

    monkeypatch.setattr(run, "build_changes_sheet", fake_build_changes_sheet)
    monkeypatch.setattr(run, "download_excel", lambda **kwargs: (_ for _ in ()).throw(AssertionError("download called")))

    monkeypatch.setattr(sys, "argv", ["run.py"])
    run.main()

    assert calls["source"] == str(source_file)
    assert calls["excel_path"] == "local_data.xlsx"
