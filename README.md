# Marksmith

A class's results, one sheet, one run.

A small tool for a teacher to collect a whole class's semester results from the
university result portal in one go, and get them back as Excel sheets and PDFs.

Supported portal: **J.C. Bose University of Science and Technology, YMCA,
Faridabad**. More universities later; each one is a new portal to teach it.

You give it a list of roll numbers. It opens the result portal for each student,
reads the report card, and saves everything neatly. The one thing it asks of you
is the CAPTCHA: the portal shows one per lookup, and you type it.

## What you need

- A Windows PC (it also runs on Mac/Linux, see below).
- A file with the roll numbers — either an Excel file (`.xlsx`) or a CSV file.
  Put the roll numbers in a column with the heading **Roll No**. A **Name**
  column is used if present. Any other columns (phone numbers, etc.) are ignored.

## How to run it

**Windows:** double-click **`run.bat`**.
The first time, it sets itself up (this takes a few minutes and downloads about
150 MB). After that it starts straight away and opens the dashboard in your
browser at `http://127.0.0.1:5001`.

**Mac/Linux (or if you prefer the command line):** you need Python 3.10 or newer.

```bash
python -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m playwright install chromium   # one-time, ~150 MB
.venv/bin/python app.py
```

(On Windows the Python path is `.venv/Scripts/python` instead of `.venv/bin/python`.)

## Using the dashboard

1. **Load the roll numbers.** Click *Load Excel / CSV* and pick your file. (If a
   file named `master_result-1.xlsx` is already in the folder, it loads by itself.)
2. **Pick the semester** and click **Start**.
3. For each student the portal's **CAPTCHA appears on the dashboard**. Read it,
   type it in the box, and press Enter.
4. If you type it wrong, a new CAPTCHA loads and you just try again (up to 5
   times). You can also click **Skip student** to move on.
5. Watch the progress line (e.g. *Processed 17/40 · Successful 16 · Failed 1*).
6. When it's done, the **download buttons** at the bottom light up.
7. If some students failed, click **Retry failed** to put them back in the queue
   and go through them again.

You can **Pause** and **Resume** any time. If the program crashes or you close
it, just start it again — the students it already finished are remembered, so you
don't type their CAPTCHAs again.

## What you get

All files land in the `output` folder:

| File | What's in it |
|---|---|
| `output/final_results.xlsx` | The main workbook. **Marks Grid** is the class register — one row per student, and for each subject a group of columns with the subject **name** on top and the subject **code** + Theory/Sessional/Practical below. Also has a **Student Summary**, a **Subject-wise Results** sheet, and a list of **Failed Students**. |
| `master_result-filled.xlsx` | If you had a master sheet (`master_result-1.xlsx`), this is a filled copy of it — marks dropped into the matching subject columns, plus SGPA, CGPA and reappear info. Your original file is never changed. |
| `output/all_results.pdf` | Every student's report card in one PDF, in class order, bookmarked by roll number and name. |
| `output/pdf/<roll>.pdf` | Each student's report card as its own PDF. |
| `output/raw/` | The raw report card pages the tool saved (kept so results can be rebuilt without re-fetching). |
| `output/processing.log` | A log of everything that happened. |

**The Excel files are yours to edit.** Open them in Excel or Google Sheets and
change anything you like — nothing is locked.

**Students whose result didn't come** (not declared, absent, or skipped) are not
dropped. Their mark cells show **`AB`** — in the Marks Grid and in the filled
master sheet — so every roll number from your input is accounted for.

## Moving to another computer

Copy the whole folder — **including the `output` folder**, which holds your
progress — set it up the same way, and start it. It carries on where it left off.

## Running the tests

```bash
.venv/Scripts/python -m pytest tests     # Windows; Mac/Linux: .venv/bin/python
```

These check the parts that don't need the live portal (reading roll numbers,
parsing a report card, building the Excel sheets, making a PDF). The PDF test
needs Chromium, which `playwright install chromium` provides.

## A note on privacy

`master_result-1.xlsx`, `master_result-filled.xlsx` and the `output` folder hold
real student data. They are kept out of version control on purpose. Don't commit
or share them.

## The CAPTCHA step

The portal asks for a CAPTCHA on every result lookup, so the tool shows you that
image and you type it. That is the one manual step; everything before and after
it is automatic.
