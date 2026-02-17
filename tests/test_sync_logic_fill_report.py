import pandas as pd

import sync_logic


class _FakeBaseWorksheet:
    def __init__(self):
        self.batch_updates = []

    def get_values(self, _range: str):
        return [["01.01.2026", "", "", "", "", "", "", "", "", "", ""]]

    def batch_update(self, updates, value_input_option="RAW"):
        self.batch_updates.append((updates, value_input_option))


class _FakeSpreadsheet:
    def worksheet(self, _title: str):
        raise Exception("worksheet not found")


def test_build_changes_sheet_prints_fill_report(monkeypatch, capsys):
    df = pd.DataFrame(
        [
            {
                "תאריך": "01.01.2026",
                "כניסה": "08:00",
                "יציאה": "17:00",
            }
        ]
    )
    monkeypatch.setattr(sync_logic.pd, "read_excel", lambda _path: df)

    base_ws = _FakeBaseWorksheet()
    spreadsheet = _FakeSpreadsheet()

    result = sync_logic.build_changes_sheet(
        spreadsheet=spreadsheet,
        base_ws=base_ws,
        sheet_name="1.26",
        excel_path="local_data.xlsx",
    )
    out = capsys.readouterr().out

    assert result is False
    assert len(base_ws.batch_updates) == 1

    updates, value_input_option = base_ws.batch_updates[0]
    assert value_input_option == "USER_ENTERED"
    assert updates == [
        {"range": "C1", "values": [["08:00"]]},
        {"range": "D1", "values": [["17:00"]]},
    ]

    assert "Дозаполнение: добавлено ячеек=2, дат=1." in out
    assert "01.01.2026: C=08:00, D=17:00" in out
    assert "Расхождений нет" in out
