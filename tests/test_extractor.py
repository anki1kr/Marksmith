"""Self-checks that need no real student data: parser, roll loader, Excel sheets, PDF render + merge.

Run:  .venv\\Scripts\\python -m pytest tests      (PDF test needs: python -m playwright install chromium)
"""
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from pypdf import PdfReader

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import automation as auto  # noqa: E402

FIXTURE = Path(__file__).parent / "fixtures" / "report_card_sample.html"
ROLL = "99999900001"


@pytest.fixture
def parsed():
    return auto.parse_result(FIXTURE.read_text(encoding="utf-8"), ROLL)


def test_parse_report_card(parsed):
    assert parsed["name"] == "TEST STUDENT ONE"
    assert parsed["reg_no"] == "00-TEST-000001"
    assert (parsed["sgpa"], parsed["cgpa"]) == ("5.25", "5.80")
    assert [s["Code"] for s in parsed["subjects"]] == ["AEC-109-V", "BCD-202-V", "BCD-204-V"]
    aec = parsed["subjects"][0]
    assert (aec["Theory"], aec["Sessional"], aec["Practical"], aec["Total"], aec["Grade"]) == \
        ("43", "10", "", 53, "B")
    absent = parsed["subjects"][1]
    assert (absent["Theory"], absent["Total"], absent["Grade"]) == ("AB", "", "AB")  # no total when a part is missing
    lab = parsed["subjects"][2]
    assert (lab["Theory"], lab["Sessional"], lab["Practical"], lab["Total"]) == ("", "20", "60", 80)
    assert parsed["re_codes"] == ["BCD-202-V", "BCG-212-V"]
    assert parsed["result"] == "RE (BCD-202-V, BCG-212-V)"
    assert parsed["total_marks"] == 133


def test_parse_without_grid_returns_no_subjects():
    d = auto.parse_result("<html><body><table><tr><td>Name:</td><td>X</td></tr></table></body></html>", "1")
    assert d["subjects"] == [] and d["result"] == "" and d["total_marks"] == ""


def test_load_rolls_prefers_roll_column_over_phone_numbers(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Class list BCA 4th sem"])
    ws.append(["Sr", "Roll No", "Student Name", "Mobile"])
    ws.append([1, 99999900001, "TEST STUDENT ONE", 9876543210])
    ws.append([2, "99999900002", "TEST STUDENT TWO", 9123456789])
    ws.append([3, None, "", None])
    f = tmp_path / "class.xlsx"
    wb.save(f)
    rolls, names = auto.load_rolls(f)
    assert rolls == ["99999900001", "99999900002"]
    assert names == {"99999900001": "TEST STUDENT ONE", "99999900002": "TEST STUDENT TWO"}


def test_load_rolls_falls_back_to_any_roll_like_cell(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append([99999900001, 99999900002])
    ws.append([99999900001])  # duplicate kept once
    f = tmp_path / "plain.xlsx"
    wb.save(f)
    assert auto.load_rolls(f) == (["99999900001", "99999900002"], {})


def test_load_rolls_reads_csv_with_roll_and_name(tmp_path):
    f = tmp_path / "class.csv"
    f.write_text("﻿Roll No,Student Name,Mobile\n"  # BOM, like Excel-exported CSV
                 "99999900001,TEST STUDENT ONE,9876543210\n"
                 "99999900002,TEST STUDENT TWO,9123456789\n", encoding="utf-8")
    rolls, names = auto.load_rolls(f)
    assert rolls == ["99999900001", "99999900002"]
    assert names == {"99999900001": "TEST STUDENT ONE", "99999900002": "TEST STUDENT TWO"}


def test_load_rolls_csv_no_header_falls_back_and_quotes(tmp_path):
    # no "roll" header -> any roll-looking cell; quoted name with a comma stays intact
    f = tmp_path / "plain.csv"
    f.write_text('99999900001,"STUDENT, TEST"\n'
                 "9876543210,ten digits counts with no header present\n"
                 "99999900001,dupe\n", encoding="utf-8")
    rolls, names = auto.load_rolls(f)
    assert rolls == ["99999900001", "9876543210"] and names == {}  # dedup, order kept, no name col


def test_excel_has_marks_grid(tmp_path, parsed, monkeypatch):
    monkeypatch.setattr(auto, "RESULTS_XLSX", tmp_path / "final.xlsx")
    auto.set_students([ROLL, "99999900002"], {"99999900002": "TEST STUDENT TWO"})
    auto.state["students"][0].update(status="COMPLETED", **{
        k: parsed[k] for k in ("name", "reg_no", "sgpa", "cgpa", "result", "total_marks", "re_codes", "subjects")})
    auto.state["students"][1].update(status="FAILED", error="Skipped by user")
    auto.save_excel()

    wb = load_workbook(tmp_path / "final.xlsx")
    assert wb.sheetnames == ["Student Summary", "Marks Grid", "Subject-wise Results", "Failed Students"]
    grid = list(wb["Marks Grid"].iter_rows(values_only=True))
    name_band, head = grid[0], grid[1]  # row 1 = subject names, row 2 = code + component
    assert "AEC-109-V Theory" in head and "AEC-109-V Practical" not in head  # NA component dropped
    assert "BCD-204-V Practical" in head and "BCD-204-V Theory" not in head
    # subject name sits over the first cell of its code block (name- and code-based header)
    assert name_band[head.index("AEC-109-V Theory")] == "Environmental Studies"
    assert name_band[head.index("BCD-204-V Sessional")] == "Data Structures Lab"
    assert not name_band[head.index("Roll No")] and not name_band[head.index("SGPA")]
    row = dict(zip(head, grid[2]))
    assert (row["Roll No"], row["Student Name"], row["AEC-109-V Theory"], row["AEC-109-V Total"]) == \
        (ROLL, "TEST STUDENT ONE", 43, 53)
    assert (row["BCD-202-V Grade"], row["No. of RE"], row["Codes of RE"], row["SGPA"]) == \
        ("AB", 2, "BCD-202-V, BCG-212-V", 5.25)
    failed = dict(zip(head, grid[3]))
    assert (failed["Roll No"], failed["Student Name"], failed["Status"]) == ("99999900002", "TEST STUDENT TWO", "FAILED")
    assert failed["AEC-109-V Theory"] == "AB" and failed["AEC-109-V Total"] == "AB"  # no result -> AB
    assert failed["Total Marks"] == "AB" and failed["Result"] == "AB"
    assert wb["Failed Students"]["A2"].value == "99999900002"
    assert wb["Student Summary"]["E2"].value == 133
    assert wb["Student Summary"]["E3"].value == "AB" and wb["Student Summary"]["H3"].value == "AB"  # failed -> AB


def test_grid_rows_untitled_subject_and_union_of_codes():
    # missing Title -> blank name band, not a crash; a 2nd student's new subject joins the header
    students = [
        {"roll": "1", "name": "A", "reg_no": "", "status": "COMPLETED", "sgpa": "", "cgpa": "",
         "result": "", "total_marks": "", "re_codes": [],
         "subjects": [{"Code": "X-1", "Title": "", "Max Theory": "75", "Theory": "40",
                       "Total": 40, "Grade": "C"}]},
        {"roll": "2", "name": "B", "reg_no": "", "status": "COMPLETED", "sgpa": "", "cgpa": "",
         "result": "", "total_marks": "", "re_codes": [],
         "subjects": [{"Code": "Y-2", "Title": "Maths", "Max Theory": "75", "Theory": "60",
                       "Total": 60, "Grade": "A"}]},
    ]
    name_band, head, *data = auto._grid_rows(students)
    assert "X-1 Theory" in head and "Y-2 Theory" in head  # union across students
    assert name_band[head.index("X-1 Theory")] == ""      # no title -> blank, no crash
    assert name_band[head.index("Y-2 Theory")] == "Maths"
    assert dict(zip(head, data[0]))["Y-2 Theory"] == ""    # student 1 has no Y-2 -> blank, not AB (COMPLETED)


def test_fill_master_writes_marks_and_ab(tmp_path, parsed, monkeypatch):
    master = tmp_path / "master_result-1.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Sr", "Roll No", "Name", "Environmental Studies", None, None, "SGPA", "CGPA"])
    ws.append([None, None, None, "AEC-109-V", None, None])          # row 2 = codes
    ws.append([None, None, None, "External", "Internal", "Total"])  # row 3 = labels
    ws.append([1, "99999900001", "TEST STUDENT ONE"])
    ws.append([2, "99999900002", "TEST STUDENT TWO"])
    wb.save(master)
    monkeypatch.setattr(auto, "MASTER_XLSX", master)
    monkeypatch.setattr(auto, "MASTER_OUT", tmp_path / "filled.xlsx")

    auto.set_students(["99999900001", "99999900002"])
    auto.state["students"][0].update(status="COMPLETED", **{
        k: parsed[k] for k in ("name", "reg_no", "sgpa", "cgpa", "result", "total_marks",
                               "re_codes", "subjects")})
    auto.state["students"][1].update(status="FAILED", error="Skipped by user")
    auto.fill_master()

    out = load_workbook(tmp_path / "filled.xlsx").active
    assert out.cell(row=4, column=4).value == 43     # External from the report card
    assert out.cell(row=4, column=5).value == 10     # Internal
    assert out.cell(row=4, column=6).value == 53     # Total
    assert out.cell(row=4, column=7).value == 5.25   # SGPA
    # no result -> AB in every mark cell, never left blank
    assert [out.cell(row=5, column=c).value for c in (4, 5, 6)] == ["AB", "AB", "AB"]
    assert load_workbook(master).active.cell(row=4, column=4).value is None  # original untouched


def test_printable_html_strips_scripts_and_adds_base():
    out = auto._printable_html(FIXTURE.read_text(encoding="utf-8"))
    assert "<script" not in out and "window.print" not in out
    assert out.count("<base href=") == 1 and out.index("<base") > out.index("<head")
    assert auto._printable_html("<p>no head</p>").startswith("<base href=")


def test_pdf_render_and_merge(tmp_path, monkeypatch):
    from playwright.sync_api import sync_playwright
    monkeypatch.setattr(auto, "PDF_DIR", tmp_path)
    monkeypatch.setattr(auto, "ALL_PDF", tmp_path / "all.pdf")
    auto.set_students([ROLL, "99999900002"])
    auto.state["students"][0].update(status="COMPLETED", name="TEST STUDENT ONE")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        out = auto.render_pdf(browser.new_page(), ROLL, FIXTURE.read_text(encoding="utf-8"))
        browser.close()
    assert out == tmp_path / f"{ROLL}.pdf" and out.read_bytes().startswith(b"%PDF")
    assert auto.merge_pdfs() == 1  # the FAILED student has no PDF and is skipped
    merged = PdfReader(tmp_path / "all.pdf")
    assert len(merged.pages) == len(PdfReader(out).pages)
    assert merged.outline[0]["/Title"] == f"{ROLL} TEST STUDENT ONE"
