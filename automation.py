"""Playwright worker, result parser, PDF + Excel writers for Marksmith.

The worker fills roll no + semester, shows the portal's captcha image in the
dashboard and waits for the teacher to type it, then opens the report card.

Outputs (written after every student, so a crash never loses progress):
  output/raw/<roll>.html + .png   raw report card (audit trail / re-parse source)
  output/pdf/<roll>.pdf           report card as PDF; output/all_results.pdf = whole class
  output/final_results.xlsx       Student Summary, Marks Grid, Subject-wise Results, Failed Students
  master_result-filled.xlsx       teacher's master sheet with marks filled in (if master_result-1.xlsx exists)
"""
import base64
import csv
import logging
import os
import re
import threading
from pathlib import Path

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from pypdf import PdfWriter

FORM_URL = "https://jcboseustymca.co.in/Forms/Student/ResultStudents.aspx"
BASE_DIR = Path(__file__).parent
MASTER_XLSX = BASE_DIR / "master_result-1.xlsx"
MASTER_OUT = BASE_DIR / "master_result-filled.xlsx"
OUT_DIR = BASE_DIR / "output"
RAW_DIR = OUT_DIR / "raw"
PDF_DIR = OUT_DIR / "pdf"
RESULTS_XLSX = OUT_DIR / "final_results.xlsx"
ALL_PDF = OUT_DIR / "all_results.pdf"
LOG_FILE = OUT_DIR / "processing.log"
POPUP_TIMEOUT_MS = 20000
MAX_CAPTCHA_ATTEMPTS = 5
ROLL_RE = re.compile(r"^\d{5,15}$")
# SHOW_BROWSER=1 opens a visible Chrome window (debugging). Default: hidden
# browser, the captcha image is shown in the dashboard instead.
SHOW_BROWSER = os.environ.get("SHOW_BROWSER", "") == "1"

for _d in (OUT_DIR, RAW_DIR, PDF_DIR):
    _d.mkdir(exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, encoding="utf-8"), logging.StreamHandler()],
)
log = logging.getLogger("extractor")

# ---------------------------------------------------------------- shared state
lock = threading.Lock()
state = {
    "phase": "idle",  # idle | running | paused | done | error
    "semester": "04",
    "current": None,          # roll currently at the form
    "message": "",            # last info/error line for the dashboard
    "captcha_png": "",        # base64 PNG of the captcha the teacher must type
    "captcha_text": "",       # what the teacher typed in the dashboard
    "students": [],           # [{roll,status,name,sgpa,cgpa,result,error,subjects:[{...}]}]
    "output": str(RESULTS_XLSX),
}
captcha_event = threading.Event()   # dashboard captcha submitted
skip_event = threading.Event()      # dashboard "Skip student" clicked
pause_event = threading.Event()     # set = running, cleared = paused
pause_event.set()
_worker = None
_last_dialog = {"msg": ""}          # text of the last alert() the portal raised


def snapshot():
    with lock:
        st = dict(state)
        st["students"] = [
            {k: v for k, v in s.items() if k != "subjects"} for s in state["students"]
        ]
        st["counts"] = _counts()
        st["paused"] = not pause_event.is_set()
    st["files"] = {  # which downloads exist right now
        "results.xlsx": RESULTS_XLSX.exists(),
        "all-results.pdf": ALL_PDF.exists(),
        "pdfs.zip": any(PDF_DIR.glob("*.pdf")),
        "master-filled.xlsx": MASTER_OUT.exists(),
    }
    return st


def _counts():
    c = {"total": len(state["students"]), "completed": 0, "failed": 0}
    for s in state["students"]:
        if s["status"] == "COMPLETED":
            c["completed"] += 1
        elif s["status"] == "FAILED":
            c["failed"] += 1
    c["processed"] = c["completed"] + c["failed"]
    c["remaining"] = c["total"] - c["processed"]
    return c


def _set(student, **kw):
    with lock:
        student.update(kw)


# ---------------------------------------------------------------- excel input
def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def _sheets(path):
    """Yield one sheet of normalized string rows. .csv = one sheet (utf-8-sig
    strips Excel's BOM); otherwise every worksheet of the workbook."""
    if str(path).lower().endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig") as fh:
            yield [[_cell(v) for v in r] for r in csv.reader(fh)]
        return
    wb = load_workbook(path, read_only=True, data_only=True)
    for ws in wb.worksheets:
        yield [[_cell(v) for v in r] for r in ws.iter_rows(values_only=True)]
    wb.close()


def load_rolls(path):
    """Roll numbers (+ names when a Name column exists) from an .xlsx/.csv.

    Prefers the column whose header says 'roll', so phone numbers in other
    columns are never mistaken for rolls; falls back to any roll-looking cell.
    Returns (rolls, {roll: name}).
    """
    rolls, names = [], {}
    for rows in _sheets(path):
        roll_col = name_col = None
        for r in rows[:10]:  # header row is near the top
            hits = [i for i, v in enumerate(r) if re.search(r"roll", v, re.I)]
            if hits:
                roll_col = hits[0]
                name_hits = [i for i, v in enumerate(r) if re.search(r"name", v, re.I)]
                name_col = name_hits[0] if name_hits else None
                break
        for r in rows:
            if roll_col is None:
                cells = r
            elif roll_col < len(r):
                cells = [r[roll_col]]
            else:
                cells = []
            for v in cells:
                if ROLL_RE.match(v) and v not in rolls:  # dedupe, keep order
                    rolls.append(v)
                    if name_col is not None and name_col < len(r) and r[name_col]:
                        names[v] = r[name_col]
    return rolls, names


def set_students(rolls, names=None):
    names = names or {}
    with lock:
        state["students"] = [
            {"roll": r, "status": "PENDING", "name": names.get(r, ""), "sgpa": "",
             "cgpa": "", "result": "", "error": "", "reg_no": "", "total_marks": "",
             "re_codes": [], "subjects": []}
            for r in rolls
        ]
        state["message"] = f"Loaded {len(rolls)} roll numbers"


def load_master():
    """Pre-load rolls + names from the teacher's master_result sheet in cwd."""
    if not MASTER_XLSX.exists():
        return False
    wb = load_workbook(MASTER_XLSX, read_only=True, data_only=True)
    ws = wb.active
    rolls, names = [], {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        v = _cell(row[1] if len(row) > 1 else None)
        if v and ROLL_RE.match(v) and v not in rolls:
            rolls.append(v)
            if len(row) > 2 and row[2]:
                names[v] = str(row[2]).strip()
    wb.close()
    if rolls:
        set_students(rolls, names)
        log.info("Pre-loaded %d rolls from %s", len(rolls), MASTER_XLSX.name)
    return bool(rolls)


def reparse_raw():
    """Resume support: rebuild COMPLETED students from saved output/raw HTML.

    Missing PDFs for restored students are rendered by the worker on its next
    run (needs a browser)."""
    n = 0
    for s in state["students"]:
        f = RAW_DIR / f"{s['roll']}.html"
        if s["status"] != "PENDING" or not f.exists():
            continue
        d = parse_result(f.read_text(encoding="utf-8"), s["roll"])
        if d["subjects"]:
            s.update(status="COMPLETED", name=d["name"] or s["name"],
                     sgpa=d["sgpa"], cgpa=d["cgpa"], result=d["result"],
                     reg_no=d["reg_no"], total_marks=d["total_marks"],
                     re_codes=d["re_codes"], subjects=d["subjects"])
            n += 1
    if n:
        log.info("Restored %d students from saved raw HTML", n)
        _save_outputs()
    return n


# ---------------------------------------------------------------- result parse
_NUM_RE = re.compile(r"^\d+(\.\d+)?$")


def _txt(el):
    return el.get_text(" ", strip=True)


def parse_result(html, roll):
    """Parse PrintReportCardNew.aspx (structure verified from a live report).

    Marks grid = the table whose header row contains 'Sr.No'; each subject
    row has 6 tds: sr, nested [code,title] table, credits, nested max-marks
    table (Theory/Sessional/Practical), nested obtained-marks table (rows
    align positionally with max-marks), grade. Footer carries SGPA/CGPA and
    an explicit reappear list like 'R-BCG-212-V BCD-202-V ...'.
    """
    soup = BeautifulSoup(html, "html.parser")
    data = {"roll": roll, "subjects": [], "name": "", "reg_no": "", "sgpa": "",
            "cgpa": "", "result": "", "total_marks": "", "re_codes": []}

    # label:value pairs from leaf tables ('Name:', 'Regn Number:', 'SGPA:', ...)
    labels = {}
    for t in soup.find_all("table"):
        if t.find("table"):
            continue
        for tr in t.find_all("tr"):
            cs = [_txt(td) for td in tr.find_all(["td", "th"])]
            for i in range(len(cs) - 1):
                if cs[i].endswith(":") and cs[i + 1] and not cs[i + 1].endswith(":"):
                    labels.setdefault(cs[i].rstrip(":").strip().lower(), cs[i + 1])
    data["name"] = labels.get("name", "")
    data["reg_no"] = labels.get("regn number", "") or labels.get("registration no", "")
    for key in ("sgpa", "cgpa"):
        v = labels.get(key, "")
        data[key] = v if _NUM_RE.match(v) else ""
    m = re.search(r"\bR-((?:[A-Z]{2,4}-\d{2,3}(?:-[A-Z0-9]+)?\s*)+)",
                  soup.get_text("  ", strip=True))
    if m:
        data["re_codes"] = m.group(1).split()

    grid = None
    for t in soup.find_all("table"):
        first = t.find("tr")
        if first and "Sr.No" in first.get_text():
            grid = t
    if grid is None:
        log.warning("No marks grid found for %s", roll)
        return data
    body = grid.find("tbody") or grid
    for tr in body.find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if len(tds) < 6 or tds[1].find("table") is None:
            continue
        code_cells = [_txt(c) for c in tds[1].find("table").find_all("td")]
        maxes = [[_txt(c) for c in r.find_all("td")]
                 for r in tds[3].find("table").find_all("tr")]
        obtained = [[_txt(c) for c in r.find_all("td")]
                    for r in tds[4].find("table").find_all("tr")]
        sub = {"Code": code_cells[0] if code_cells else "",
               "Title": code_cells[1] if len(code_cells) > 1 else ""}
        total, complete = 0, True
        for i, mrow in enumerate(maxes):  # rows align by position
            if len(mrow) < 2:
                continue
            comp, mx = mrow[0], mrow[1]          # e.g. 'Theory', '75'
            got = obtained[i][-1] if i < len(obtained) and obtained[i] else ""
            sub[comp] = "" if got.upper() == "NA" else got
            sub[f"Max {comp}"] = mx
            if mx.upper() != "NA":               # component counts toward total
                if _NUM_RE.match(sub[comp] or ""):
                    total += float(sub[comp])
                else:
                    complete = False
        sub["Total"] = _num(total) if complete else ""
        sub["Grade"] = _txt(tds[5])
        sub["Credits"] = _txt(tds[2])
        data["subjects"].append(sub)

    totals = [s["Total"] for s in data["subjects"] if s["Total"] != ""]
    data["total_marks"] = _num(sum(totals)) if totals else ""
    if data["re_codes"]:
        data["result"] = "RE (" + ", ".join(data["re_codes"]) + ")"
    elif data["subjects"]:
        data["result"] = "PASS"
    return data


# ---------------------------------------------------------------- excel output
_COMPONENTS = ("Theory", "Sessional", "Practical")


def _grid_rows(students):
    """One row per student, one column block per subject (class marks register).

    Two header rows so each subject is both name- and code-based: row 1 carries
    the subject NAME over its block, row 2 the '<code> <component>' cells.
    """
    codes, comps, titles = [], {}, {}   # order of first appearance; code -> components / subject name
    for s in students:
        for sub in s["subjects"]:
            code = sub.get("Code", "")
            if not code:
                continue
            if code not in codes:
                codes.append(code)
                comps[code] = []
            if not titles.get(code) and sub.get("Title"):
                titles[code] = sub["Title"]
            for comp in _COMPONENTS:
                if sub.get(f"Max {comp}", "NA").upper() != "NA" and comp not in comps[code]:
                    comps[code].append(comp)
    lead = ["Roll No", "Student Name", "Reg No"]
    tail = ["Total Marks", "SGPA", "CGPA", "Result", "No. of RE", "Codes of RE", "Status"]
    name_row, head = [""] * len(lead), list(lead)
    for c in codes:
        block = comps[c] + ["Total", "Grade"]
        name_row += [titles.get(c, "")] + [""] * (len(block) - 1)  # name over the block
        head += [f"{c} {k}" for k in block]
    name_row += [""] * len(tail)
    head += tail
    rows = [name_row, head]
    for s in students:
        # A student whose result never came through (not declared / absent / skipped)
        # reads "AB" in every mark cell, so no input roll is left blank in the register.
        done = s["status"] == "COMPLETED"
        ab = "" if done else "AB"
        by_code = {sub.get("Code", ""): sub for sub in s["subjects"]}
        r = [s["roll"], s["name"], s.get("reg_no", "")]
        for c in codes:
            sub = by_code.get(c, {})
            for k in comps[c] + ["Total", "Grade"]:
                v = sub.get(k, "")
                r.append(_num(v) if v != "" else ab)
        re_codes = s.get("re_codes") or []
        r += [s.get("total_marks", "") or ab, _num(s["sgpa"]), _num(s["cgpa"]),
              s["result"] or ab, len(re_codes) if s["subjects"] else "",
              ", ".join(re_codes), s["status"]]
        rows.append(r)
    return rows


def save_excel():
    with lock:
        students = [dict(s, subjects=list(s["subjects"])) for s in state["students"]]
        semester = state["semester"]
    wb = Workbook()

    ws = wb.active
    ws.title = "Student Summary"
    ws.append(["Roll No", "Student Name", "Reg No", "Semester", "Total Marks",
               "SGPA", "CGPA", "Result", "Status", "Error"])
    for s in students:
        ab = "" if s["status"] == "COMPLETED" else "AB"  # no result -> AB, not blank
        ws.append([s["roll"], s["name"], s.get("reg_no", ""), semester,
                   s.get("total_marks", "") or ab, s["sgpa"], s["cgpa"], s["result"] or ab,
                   s["status"], s["error"]])

    ws_grid = wb.create_sheet("Marks Grid")
    for row in _grid_rows(students):
        ws_grid.append(row)
    ws_grid.freeze_panes = "D3"  # two header rows: subject name band + code/component

    ws2 = wb.create_sheet("Subject-wise Results")
    cols = []
    for s in students:
        for row in s["subjects"]:
            for k in row:
                if k not in cols:
                    cols.append(k)
    ws2.append(["Roll No"] + cols)
    for s in students:
        for row in s["subjects"]:
            ws2.append([s["roll"]] + [row.get(c, "") for c in cols])

    ws3 = wb.create_sheet("Failed Students")
    ws3.append(["Roll No", "Error"])
    for s in students:
        if s["status"] == "FAILED":
            ws3.append([s["roll"], s["error"]])

    tmp = str(RESULTS_XLSX) + ".tmp"
    wb.save(tmp)
    os.replace(tmp, RESULTS_XLSX)  # never leave a half-written workbook


# ---------------------------------------------------------------- pdf output
_SCRIPT_RE = re.compile(r"<script\b.*?</script>", re.S | re.I)
_HEAD_RE = re.compile(r"<head\b[^>]*>", re.I)


def _printable_html(html):
    """Saved report card -> printable page: scripts dropped (a print page may call
    window.print()), <base> added so the portal's relative CSS/images still resolve."""
    html = _SCRIPT_RE.sub("", html)
    base = f'<base href="{FORM_URL}">'
    if _HEAD_RE.search(html):
        return _HEAD_RE.sub(lambda m: m.group(0) + base, html, count=1)
    return base + html


def render_pdf(pdf_page, roll, html):
    """Render one saved report card to output/pdf/<roll>.pdf on a headless page."""
    pdf_page.set_content(_printable_html(html), wait_until="load")
    try:
        pdf_page.wait_for_load_state("networkidle", timeout=10000)
    except Exception:  # offline or slow portal assets: print what we have
        pass
    out = PDF_DIR / f"{roll}.pdf"
    pdf_page.pdf(path=str(out), format="A4", print_background=True,
                 margin={"top": "10mm", "bottom": "10mm", "left": "8mm", "right": "8mm"})
    return out


def merge_pdfs():
    """Every completed student's PDF -> output/all_results.pdf, class order, one bookmark each."""
    with lock:
        order = [(s["roll"], s["name"]) for s in state["students"] if s["status"] == "COMPLETED"]
    n = 0
    tmp = str(ALL_PDF) + ".tmp"
    with PdfWriter() as writer:
        for roll, name in order:
            f = PDF_DIR / f"{roll}.pdf"
            if not f.exists():
                continue
            first = len(writer.pages)
            writer.append(str(f))
            writer.add_outline_item(f"{roll} {name}".strip(), first)
            n += 1
        if n:
            writer.write(tmp)
    if n:
        os.replace(tmp, ALL_PDF)
    return n


def _render_missing_pdfs(pdf_page):
    """Students restored from raw HTML (older run / older version) get their PDF now."""
    with lock:
        rolls = [s["roll"] for s in state["students"] if s["status"] == "COMPLETED"]
    n = 0
    for roll in rolls:
        raw = RAW_DIR / f"{roll}.html"
        if (PDF_DIR / f"{roll}.pdf").exists() or not raw.exists():
            continue
        try:
            render_pdf(pdf_page, roll, raw.read_text(encoding="utf-8"))
            n += 1
        except Exception:
            log.exception("PDF render failed for %s", roll)
    if n:
        log.info("Rendered %d missing PDFs", n)
        merge_pdfs()


# ---------------------------------------------------------------- master fill
_CODE_RE = re.compile(r"[A-Z]{2,4}[- ]?\d{3}", re.I)
_FIELD_HINTS = [
    ("code", re.compile(r"code", re.I)),
    ("internal", re.compile(r"intern|sessional|minor", re.I)),
    ("external", re.compile(r"extern|theory|major", re.I)),
    ("practical", re.compile(r"pract|viva", re.I)),
    ("total", re.compile(r"total|obtain", re.I)),
    ("grade", re.compile(r"grade|result", re.I)),
]
# the row-3 labels that name a mark cell (same set the fill loop below handles)
_MARK_LABEL_RE = re.compile(r"extern|intern|sessional|pract|total")


def _norm(s):
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def _canon_subject(row):
    """Map a result-page subject row (unknown headers) to canonical fields."""
    out = {}
    for k, v in row.items():
        for field, pat in _FIELD_HINTS:
            if pat.search(k) and field not in out:
                out[field] = str(v).strip()
                break
    if "code" not in out:  # fall back: any cell that looks like a course code
        for v in row.values():
            if _CODE_RE.fullmatch(str(v).strip().replace(" ", "")):
                out["code"] = str(v).strip()
                break
    return out


def _num(v):
    try:
        f = float(v)
        return int(f) if f.is_integer() else f
    except (TypeError, ValueError):
        return v or ""


def fill_master():
    """Write the class's marks into a copy of the teacher's master sheet.

    Row 1 = subject names + CGPA/SGPA/RE headers, row 2 = subject codes,
    row 3 = External/Internal/Total (or Sessional/Practical/Total) labels.
    A student whose result never came through gets "AB" in every subject cell,
    matching the Marks Grid, so no roll is left silently blank.
    Original file is never modified; output goes to master_result-filled.xlsx.
    """
    if not MASTER_XLSX.exists():
        return
    with lock:
        students = [dict(s, subjects=list(s["subjects"])) for s in state["students"]]
    if not any(s["status"] == "COMPLETED" for s in students):
        return  # nothing fetched yet: leave the master copy alone
    wb = load_workbook(MASTER_XLSX)
    ws = wb.active
    head = list(ws.iter_rows(min_row=1, max_row=3))
    special = {}
    for cell in head[0]:
        key = _norm(cell.value) if cell.value else ""
        if key in ("CGPA", "SGPA", "NOOFRE", "CODESOFRE"):
            special[key] = cell.column
    code_cols = {_norm(c.value): c.column for c in head[1]
                 if c.value and _CODE_RE.search(str(c.value))}
    sub_labels = {c.column: str(c.value).strip().lower()
                  for c in head[2] if c.value}
    roll_rows = {}
    for r in ws.iter_rows(min_row=2):
        v = r[1].value if len(r) > 1 else None
        if v is not None and ROLL_RE.match(str(v).strip()):
            roll_rows[str(v).strip()] = r[0].row

    for s in students:
        xr = roll_rows.get(s["roll"])
        if xr is None:
            log.warning("Roll %s not found in master sheet", s["roll"])
            continue
        if s["status"] != "COMPLETED":  # not declared / absent / skipped -> AB
            for start in code_cols.values():
                for col in range(start, start + 3):
                    if _MARK_LABEL_RE.search(sub_labels.get(col, "")):
                        ws.cell(row=xr, column=col, value="AB")
            continue
        re_codes = list(s.get("re_codes") or [])
        for row in s["subjects"]:
            canon = _canon_subject(row)
            code = _norm(canon.get("code", ""))
            if not s.get("re_codes") and re.match(  # fallback: derive from grade
                    r"^(F\b|F$|RE|RL|AB)", canon.get("grade", "").strip(), re.I):
                re_codes.append(canon.get("code", code))
            start = code_cols.get(code)
            if not start:
                continue
            for col in range(start, start + 3):
                label = sub_labels.get(col, "")
                if "extern" in label:
                    val = canon.get("external", "")
                elif "intern" in label or "sessional" in label:
                    val = canon.get("internal", "")
                elif "pract" in label:
                    val = canon.get("practical", canon.get("external", ""))
                elif "total" in label:
                    val = canon.get("total", "")
                else:
                    continue
                if val != "":
                    ws.cell(row=xr, column=col, value=_num(val))
        if s["sgpa"] and "SGPA" in special:
            ws.cell(row=xr, column=special["SGPA"], value=_num(s["sgpa"]))
        if s["cgpa"] and "CGPA" in special:
            ws.cell(row=xr, column=special["CGPA"], value=_num(s["cgpa"]))
        if s["subjects"] and "NOOFRE" in special:
            ws.cell(row=xr, column=special["NOOFRE"], value=len(re_codes))
        if re_codes and "CODESOFRE" in special:
            ws.cell(row=xr, column=special["CODESOFRE"], value=", ".join(re_codes))

    tmp = str(MASTER_OUT) + ".tmp"
    wb.save(tmp)
    os.replace(tmp, MASTER_OUT)


def _save_outputs():
    """Excel + master + merged PDF; one failing (e.g. file open in Excel) never stops the batch."""
    for fn in (save_excel, fill_master, merge_pdfs):
        try:
            fn()
        except Exception:
            log.exception("%s failed (will retry after the next student)", fn.__name__)


# ---------------------------------------------------------------- worker
def start_worker(semester):
    global _worker
    with lock:
        state["semester"] = semester
        state["phase"] = "running"
        state["message"] = ""
    pause_event.set()
    if _worker and _worker.is_alive():
        return
    _worker = threading.Thread(target=_run, daemon=True)
    _worker.start()


def _next_pending():
    with lock:
        for s in state["students"]:
            if s["status"] == "PENDING":
                return s
    return None


def _on_dialog(d):
    _last_dialog["msg"] = d.message
    d.accept()


def _run():
    from playwright.sync_api import sync_playwright
    log.info("Worker started, launching browser (visible=%s)", SHOW_BROWSER)
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=not SHOW_BROWSER)
            page = browser.new_page()
            page.on("dialog", _on_dialog)
            # page.pdf() needs headless Chromium: a page in the same context when
            # hidden (shares the portal session), a second hidden browser otherwise.
            pdf_browser = p.chromium.launch(headless=True) if SHOW_BROWSER else browser
            pdf_page = pdf_browser.new_page() if SHOW_BROWSER else page.context.new_page()
            _render_missing_pdfs(pdf_page)
            while True:
                pause_event.wait()
                s = _next_pending()
                if s is None:
                    break
                try:
                    _process_student(page, pdf_page, s)
                except Exception as e:  # keep the batch alive whatever breaks
                    log.exception("Unhandled error for %s", s["roll"])
                    _set(s, status="FAILED", error=str(e)[:300])
                _save_outputs()
            if pdf_browser is not browser:
                pdf_browser.close()
            browser.close()
        with lock:
            state["phase"] = "done"
            state["current"] = None
            state["captcha_png"] = ""
            state["message"] = f"Finished. Output: {RESULTS_XLSX}"
        log.info("Worker finished")
    except Exception as e:
        log.exception("Worker crashed")
        with lock:
            state["phase"] = "error"
            state["captcha_png"] = ""
            state["message"] = f"Worker crashed: {e}"


def _captcha_b64(page):
    try:
        img = page.locator("img[src*='Captcha'], img[alt='CAPTCHA']").first
        return base64.b64encode(img.screenshot(timeout=10000)).decode("ascii")
    except Exception:
        log.warning("Could not capture the captcha image")
        return ""


def _process_student(page, pdf_page, s):
    from playwright.sync_api import TimeoutError as PWTimeout
    roll = s["roll"]
    log.info("Processing %s", roll)
    with lock:
        state["current"] = roll
        semester = state["semester"]

    for attempt in range(1, MAX_CAPTCHA_ATTEMPTS + 1):
        _last_dialog["msg"] = ""
        page.goto(FORM_URL, wait_until="load", timeout=60000)
        page.fill("#txtRollNo", roll)
        page.select_option("#ddlSem", semester)
        page.focus("#txtCaptcha")

        with lock:
            state["captcha_png"] = _captcha_b64(page)
            state["captcha_text"] = ""
        _set(s, status="CAPTCHA_REQUIRED")
        captcha_event.clear()
        skip_event.clear()
        while not captcha_event.wait(0.5):
            if skip_event.is_set():
                with lock:
                    state["captcha_png"] = ""
                _set(s, status="FAILED", error="Skipped by user")
                return
        with lock:
            text = state["captcha_text"].strip()
            state["captcha_png"] = ""
        if text:  # typed in the dashboard (empty = typed in the visible browser window)
            page.fill("#txtCaptcha", text)
        _set(s, status="PROCESSING")

        try:
            with page.expect_popup(timeout=POPUP_TIMEOUT_MS) as pinfo:
                page.click("#btnResult")
            popup = pinfo.value
            popup.wait_for_load_state("load", timeout=60000)
            popup.wait_for_timeout(1500)  # let late scripts/grids settle
            html = popup.content()
            (RAW_DIR / f"{roll}.html").write_text(html, encoding="utf-8")
            try:
                popup.screenshot(path=str(RAW_DIR / f"{roll}.png"), full_page=True)
            except Exception:
                log.warning("Screenshot failed for %s", roll)
            popup.close()
        except PWTimeout:
            if _has_inline_result(page):  # result rendered on same page, no popup
                html = page.content()
                (RAW_DIR / f"{roll}.html").write_text(html, encoding="utf-8")
            else:
                msg = _page_error(page)
                log.warning("Attempt %d for %s: %s", attempt, roll, msg)
                _set(s, error=msg)
                if "captcha" not in msg.lower():  # invalid roll / not declared: re-typing won't help
                    _set(s, status="FAILED")
                    return
                with lock:
                    state["message"] = f"{roll}: {msg} — type the new captcha or skip"
                continue  # fresh page + fresh captcha, teacher retries

        data = parse_result(html, roll)
        if not data["subjects"]:  # popup opened but no marks grid = no result
            _set(s, status="FAILED", error="Result page opened but no marks found "
                 "(result not declared?)")
            return
        try:
            render_pdf(pdf_page, roll, html)
        except Exception:
            log.exception("PDF failed for %s (raw HTML and PNG are saved)", roll)
        _set(s, status="COMPLETED", error="", name=data["name"] or s["name"],
             sgpa=data["sgpa"], cgpa=data["cgpa"], result=data["result"],
             reg_no=data["reg_no"], total_marks=data["total_marks"],
             re_codes=data["re_codes"], subjects=data["subjects"])
        with lock:
            state["message"] = f"{roll}: done ({len(data['subjects'])} subject rows)"
        log.info("Completed %s: name=%r sgpa=%r subjects=%d",
                 roll, data["name"], data["sgpa"], len(data["subjects"]))
        return

    _set(s, status="FAILED",
         error=s["error"] or f"No result after {MAX_CAPTCHA_ATTEMPTS} attempts")


def _page_error(page):
    if _last_dialog["msg"]:
        return _last_dialog["msg"].strip()[:300]
    try:
        msg = page.inner_text("#lblMessage", timeout=3000).strip()
    except Exception:
        msg = ""
    return msg or "No result window opened (wrong captcha / invalid roll?)"


def _has_inline_result(page):
    try:
        if page.locator("#gvResult tr").count() > 1:
            return True
        return bool(page.inner_text("#lblSGPA", timeout=1000).strip())
    except Exception:
        return False


# ---------------------------------------------------------------- self-check
def demo():
    """Self-check against whatever report card is sitting in output/raw.

    Checks the SHAPE of a parsed card, never one student's identity, so it keeps
    working on the next capture and no real name is ever written into this file.
    Synthetic checks that need no real data: `python -m pytest tests`.
    """
    raws = sorted(RAW_DIR.glob("*.html"))
    assert raws, "run the app once first to capture a raw report"
    raw = raws[0]
    roll = raw.stem
    d = parse_result(raw.read_text(encoding="utf-8"), roll)
    assert d["name"], "no student name parsed"
    assert d["reg_no"], "no registration number parsed"
    assert d["subjects"], "no subject rows parsed"
    for sub in d["subjects"]:
        assert sub["Code"], sub
        assert sub["Grade"] or sub["Total"] != "", sub
    assert d["result"] == "PASS" or d["result"].startswith("RE"), d["result"]
    print(f"parsed {raw.name}: {len(d['subjects'])} subjects, "
          f"sgpa={d['sgpa']!r} cgpa={d['cgpa']!r} result={d['result']!r}")

    if not load_master():
        print("no master_result-1.xlsx present, skipping the master-fill check")
        return
    reparse_raw()
    done = [s for s in state["students"] if s["status"] == "COMPLETED"]
    assert done, "no student could be rebuilt from output/raw"
    wb = load_workbook(MASTER_OUT)
    ws = wb.active
    codes = {_norm(c.value): c.column for c in list(ws.iter_rows(
        min_row=2, max_row=2))[0] if c.value and _CODE_RE.search(str(c.value))}
    assert codes, "no subject codes found in row 2 of the master sheet"
    filled = sum(
        1 for s in done
        for r in ws.iter_rows(min_row=2)
        if str(r[1].value).strip() == s["roll"]
        and any(ws.cell(row=r[0].row, column=c).value not in (None, "")
                for c in codes.values()))
    assert filled, "master sheet has no marks written into any subject column"
    print(f"self-check OK — {filled} student(s) filled in {MASTER_OUT.name}")


if __name__ == "__main__":
    demo()
