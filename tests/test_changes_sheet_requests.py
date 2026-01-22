from datetime import datetime

from sync_logic import _build_changes_sheet_requests


def test_build_changes_sheet_requests_structure():
    changes_rows = [
        ["01.01.2024", "08:00", "17:00", "08:30", "17:00", "", 1, 0],
        ["", "", "", "09:00", "12:00", "", -1, 1],
    ]

    requests = _build_changes_sheet_requests(
        "Изменения 1.24",
        12345,
        changes_rows,
        datetime(2024, 1, 2),
    )

    assert any("addSheet" in req for req in requests)
    assert sum(1 for req in requests if "mergeCells" in req) == 2
    assert sum(1 for req in requests if "addConditionalFormatRule" in req) == 9

    update_requests = [req for req in requests if "updateCells" in req]
    assert len(update_requests) >= 3
    assert any("Дата изменений" in cell["userEnteredValue"]["stringValue"] for cell in update_requests[0]["updateCells"]["rows"][0]["values"])

    data_update = next(
        req
        for req in update_requests
        if req["updateCells"]["start"]["rowIndex"] == 4
        and req["updateCells"]["start"]["columnIndex"] == 0
    )
    data_rows = data_update["updateCells"]["rows"]
    start_row_index = data_update["updateCells"]["start"]["rowIndex"]
    first_row_number = start_row_index + 1
    second_row_number = start_row_index + 2

    main_formula = data_rows[0]["values"][5]["userEnteredValue"]["formulaValue"]
    bonus_formula = data_rows[1]["values"][5]["userEnteredValue"]["formulaValue"]

    assert "N(" not in main_formula
    assert f'B{first_row_number}<>""' in main_formula
    assert f'C{first_row_number}<>""' in main_formula
    assert f'D{first_row_number}<>""' in main_formula
    assert f'E{first_row_number}<>""' in main_formula
    assert "N(" in bonus_formula
    assert f'B{second_row_number}=""' in bonus_formula
    assert f'C{second_row_number}=""' in bonus_formula
    assert f'D{second_row_number}=""' in bonus_formula
    assert f'E{second_row_number}=""' in bonus_formula
