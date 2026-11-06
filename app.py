"""Marksmith — local dashboard. Run: run.bat (or python app.py)"""
import io
import webbrowser
import zipfile

from flask import Flask, jsonify, render_template, request, send_file

import automation as auto

app = Flask(__name__)
UPLOAD_PATH = auto.OUT_DIR / "uploaded_rolls.xlsx"
SEMESTERS = {f"{n:02d}" for n in range(1, 9)}
DOWNLOADS = {  # fixed names only: never serve an arbitrary path
    "results.xlsx": auto.RESULTS_XLSX,
    "all-results.pdf": auto.ALL_PDF,
    "master-filled.xlsx": auto.MASTER_OUT,
}


@app.get("/")
def index():
    return render_template("index.html")


@app.get("/api/state")
def api_state():
    return jsonify(auto.snapshot())


@app.post("/api/upload")
def api_upload():
    f = request.files.get("file")
    if not f or not f.filename.lower().endswith((".xlsx", ".xlsm", ".csv")):
        return jsonify(error="Upload an .xlsx or .csv file"), 400
    save_path = UPLOAD_PATH.with_suffix(".csv" if f.filename.lower().endswith(".csv") else ".xlsx")
    f.save(save_path)
    rolls, names = auto.load_rolls(save_path)
    if not rolls:
        return jsonify(error="No roll numbers found in the file"), 400
    auto.set_students(rolls, names)
    restored = auto.reparse_raw()  # students scraped in an earlier run need no captcha again
    return jsonify(ok=True, count=len(rolls), restored=restored)


@app.post("/api/start")
def api_start():
    if not auto.state["students"]:
        return jsonify(error="Load roll numbers first"), 400
    sem = (request.get_json(silent=True) or {}).get("semester", "04")
    if sem not in SEMESTERS:
        return jsonify(error="Pick a semester"), 400
    auto.start_worker(sem)
    return jsonify(ok=True)


@app.post("/api/pause")
def api_pause():
    auto.pause_event.clear()
    auto.state["phase"] = "paused"
    return jsonify(ok=True)


@app.post("/api/resume")
def api_resume():
    auto.pause_event.set()
    auto.state["phase"] = "running"
    return jsonify(ok=True)


@app.post("/api/captcha")
def api_captcha():
    """Teacher typed the captcha in the dashboard (empty text = typed in the visible browser)."""
    text = (request.get_json(silent=True) or {}).get("text", "")
    with auto.lock:
        auto.state["captcha_text"] = str(text)[:20]
    auto.captcha_event.set()
    return jsonify(ok=True)


@app.post("/api/skip")
def api_skip():
    auto.skip_event.set()
    return jsonify(ok=True)


@app.post("/api/retry-failed")
def api_retry_failed():
    with auto.lock:
        n = 0
        for s in auto.state["students"]:
            if s["status"] == "FAILED":
                s["status"] = "PENDING"
                s["error"] = ""
                n += 1
    if n:
        auto.start_worker(auto.state["semester"])
    return jsonify(ok=True, retried=n)


@app.get("/download/<name>")
def download(name):
    if name == "pdfs.zip":  # every per-student PDF, built on the fly
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(auto.PDF_DIR.glob("*.pdf")):
                z.write(f, f.name)
        buf.seek(0)
        return send_file(buf, as_attachment=True, download_name="result_pdfs.zip",
                         mimetype="application/zip")
    path = DOWNLOADS.get(name)
    if path is None or not path.exists():
        return jsonify(error="Not generated yet"), 404
    return send_file(path, as_attachment=True)


if __name__ == "__main__":
    auto.load_master()  # pre-load rolls from master_result-1.xlsx if present
    auto.reparse_raw()  # restore already-scraped students from output/raw/
    webbrowser.open("http://127.0.0.1:5001")
    app.run(host="127.0.0.1", port=5001, debug=False)
