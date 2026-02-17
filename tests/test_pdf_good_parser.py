import pdfplumber

from parsers.parse_pdf_ylm import parse_pdf_ylm


class _FakePage:
    def __init__(self, tables, text):
        self._tables = tables
        self._text = text

    def extract_tables(self):
        return self._tables

    def extract_text(self):
        return self._text


class _FakePDF:
    def __init__(self, pages):
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


def test_parse_good_pdf_canonical(monkeypatch):
    tables = [
        [
            ["יום", "כניסה", "יציאה", "אתר", "הערות"],
            ["28", "7:00", "15:00", "Site A", ""],
            ["29", "08:00", "16:00", "Site B", "תשלום מובטח"],
        ]
    ]
    text = "\n".join(
        [
            "ינואר 2026",
            "תאריך דוח: 29/01/2026",
            "שעות עבודה בפועל: 251:00",
            "ימי עבודה בפועל: 23",
            "נסיעות: 0.00",
        ]
    )
    pages = [_FakePage(tables, text)]

    def fake_open(_path):
        return _FakePDF(pages)

    monkeypatch.setattr(pdfplumber, "open", fake_open)

    raw = parse_pdf_ylm("dummy.pdf")

    assert set(raw.meta.keys()) == {
        "report_date",
        "month_total_hours",
        "month_total_days",
        "trips",
        "total_row",
    }
    assert raw.meta["report_date"] == "29.01.2026"
    assert raw.meta["month_total_hours"] == "251:00"
    assert raw.meta["month_total_days"] == "23"
    assert raw.meta["trips"] == "0.00"
    assert raw.meta["total_row"] == ""

    assert len(raw.rows) == 2
    assert list(raw.rows[0].keys()) == ["date", "time_in", "time_out", "site", "notes"]
    assert raw.rows[0]["date"] == "28.01.2026"
    assert raw.rows[0]["time_in"] == "07:00"
    assert raw.rows[0]["time_out"] == "15:00"
    assert raw.rows[0]["site"] == "Site A"
    assert raw.rows[0]["notes"] == ""

    assert raw.rows[1]["notes"] != ""
