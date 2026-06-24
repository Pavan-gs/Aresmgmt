# Fund Ops Automation Workshop — Participant Workbook
## Your Complete Reference: Prompts · Lab Steps · Expected Outputs · Debugging Guide

---

## BEFORE YOU START — SETUP CHECKLIST

```
□ 1. VS Code is open with folder: C:\FundAutomation\
□ 2. Copilot Chat is accessible: Ctrl+Shift+I opens it
□ 3. Run: python create_workshop_data.py  (creates all 15 fund files)
□ 4. Verify: data\input\format_a_citco\ has 5 Excel files
□ 5. Verify: data\sql\funddb.sqlite exists
□ 6. Terminal works: open with Ctrl+` in VS Code
□ 7. Python works: type  python --version  → should show 3.10+
```

**If any check fails → ask your trainer before starting.**

---

## THE LEARNING ARC — RUNS FOR EVERY LAB

```
STEP 1 → DEMO      See what you're building (trainer runs it)
STEP 2 → PROMPT    Paste the prompt below into Copilot Chat
STEP 3 → UNDERSTAND Walk through the code line by line
STEP 4 → VERIFY    Run the specific tests listed. Compare to EXPECTED OUTPUT
STEP 5 → ENTERPRISE Apply the production upgrade using the checklist
```

---

## THE 5-PART PROMPT FORMULA

Every prompt you write should have all 5 parts:

```
[CONTEXT]    Who you are, what system you have, what files exist
[ACTION]     What the code should do — step by step
[DATA SHAPE] Exact column names, cell references, data formats
[OUTPUT]     Where to save results, what format, what columns
[CONSTRAINTS] What to do when things go wrong — never crash silently
```

**Test your prompt:** If Copilot generates code that works but crashes on the broken fund files (NU013, OMICRON015) — your CONSTRAINTS section needs more detail.

---

## PROMPT QUALITY LADDER

| Level | Prompt | What You Get |
|---|---|---|
| ❌ Empty | `# read excel file` | Copilot guesses. 50% wrong. |
| ⚠️ Vague | `Read NAV from Excel` | Right concept, missing everything else |
| ✓ Descriptive | Full sentence with path and field | Works on clean data |
| ✅ Expert | Uses all 5 parts of the formula | Production quality, first time |

---

## MASTER PROMPT LIBRARY — COPY AND CUSTOMISE

---

### PROMPT 1: Excel Collection — Citco Format (Format A)

**Use for:** Labs 1 and 2  
**Fill in:** folder path, output path

```
I work in fund accounting. I have Excel files in this folder:
C:\FundAutomation\data\input\format_a_citco\

Each file is a monthly NAV report from Citco Fund Services.
Each file has a sheet called "NAV Summary" with this exact layout:
  - Cell A2: Fund name (text)
  - Cell B4: Valuation date (formatted as DD-MMM-YYYY)
  - Cell B5: Net Asset Value (a float like 125432891.50)
  - Cell B6: Currency code (USD, EUR, or GBP)
  - Cell B7: Units outstanding (a float)

I need Python code using pandas and openpyxl to:
1. Use pathlib.Path to loop through every .xlsx file in the folder
2. Open each file with openpyxl in read-only mode (faster, use data_only=True)
3. Extract all 5 fields from sheet "NAV Summary"
4. Compute NAV per unit = B5 / B7 (round to 4 decimal places)
5. Add columns: FileName (just the filename, not full path) and ExtractedAt (datetime.now())
6. Append to a list of dictionaries
7. After the loop, convert the list to a pandas DataFrame
8. Save the DataFrame to: reports\daily\master_nav_[YYYYMMDD].xlsx
   where [YYYYMMDD] is today's date (use datetime.now().strftime('%Y%m%d'))
9. Also save a CSV version to the same folder

ERROR HANDLING (this is critical):
- If the sheet "NAV Summary" is NOT in the workbook: 
    log a WARNING to logs\collection_errors.txt and CONTINUE to next file
- If cell B5 is blank or None:
    log a WARNING with the fund name and CONTINUE — do not use 0 as a substitute
- If ANY other exception occurs:
    log it with full error message and CONTINUE — never crash the whole script

LOGGING:
- Use Python's logging module
- Log to: logs\collection_agent_[DATE].log
- Include: timestamp, file name, action (SUCCESS/WARNING/ERROR), NAV value extracted
- At the end, print and log: "Completed: X successful, Y failed"

Use pathlib.Path for all file paths (not os.path — pathlib works on Windows and Mac).
Add a comment above each major section explaining what it does in one plain English sentence.
```

**Expected output after running on format_a_citco/:**
- 5 rows in master_nav_YYYYMMDD.xlsx (all Citco funds)
- 0 errors (all 5 Citco files are clean)
- Log file shows SUCCESS for each file

---

### PROMPT 2: Dynamic Discovery — Any Format

**Use for:** Lab 3  
**No customisation needed — it works on all formats**

```
I work with Excel files from multiple fund administrators.
Each file has different layouts — NAV, date, and fund name are in 
different cells and sometimes different sheets in each file.

I need a Python function called discover_nav_fields(filepath: str) -> dict that:

1. Opens the Excel file using openpyxl (read_only=True, data_only=True)
2. Iterates through EVERY sheet, then EVERY row, then EVERY cell
3. Skips cells that are None, empty, or purely numeric
4. For each cell with text content, check ALL of these keyword lists:

   NAV_LABELS = ["net asset value", "nav", "total nav", "fund total", "fund nav",
                 "net assets", "total fund value"]
   DATE_LABELS = ["valuation date", "nav date", "as of", "report date", 
                  "statement date", "period end"]
   UNITS_LABELS = ["units outstanding", "units", "shares outstanding", 
                   "total units", "units issued"]
   NAME_LABELS = ["fund name", "portfolio", "fund:", "portfolio name"]

4. When a matching label cell is found:
   - Check the cell to its RIGHT (same row, column+1)  
   - Check the cell BELOW it (row+1, same column)
   - Use whichever contains a non-None, non-empty value

5. Return a dictionary with these keys:
   {
     "NAVValue": float or None,
     "NAVLabel": str (the exact label text matched),
     "NAVCellRef": str (e.g. "Sheet1!B5"),
     "NAVDate": datetime.date or None,
     "NAVDateRef": str,
     "Units": float or None,
     "FundName": str or None,
     "SheetFound": str (name of sheet where NAV was found),
     "ConfidenceScore": int (0-100, add 25 for each field found)
   }

6. Handle these edge cases:
   - Merged cells: catch AttributeError and skip
   - String numbers like "1,234,567.89": strip commas, convert to float
   - Date strings: try common formats with dateutil.parser.parse()
   - If nothing found: return dict with all values as None, ConfidenceScore=0

7. Wrap the ENTIRE function in try/except — return None on any unhandled error

Also write batch_discover(folder: str) -> pd.DataFrame that:
- Calls discover_nav_fields for every .xlsx file in the folder
- Returns a DataFrame with one row per file
- Adds columns: FileName, ProcessedAt, ErrorMessage

Use openpyxl and pandas. Install dateutil with: pip install python-dateutil
```

**Expected output on all 15 funds:**
- Funds 1-10 (Citco + SS&C): ConfidenceScore > 90%
- LAMBDA011, MU012, XI014: ConfidenceScore 75-90% (generic but valid)
- NU013: ConfidenceScore 0% — sheet "NAV_Summary" found but wrong label (no match)
- OMICRON015: ConfidenceScore ~25% — date and name found, NAV is blank

---

### PROMPT 3: Outlook Attachment Agent

**Use for:** Lab 4  
**Fill in:** tenant_id, client_id, etc. from config/settings.json

```
I need to automate downloading fund NAV email attachments from Outlook
using the Microsoft Graph API. I do NOT want to use any Microsoft SDK —
use only the standard Python requests library.

AUTHENTICATION:
  Method: client_credentials (app-only, no user login needed)
  Token URL: https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token
  All credentials are loaded from config/settings.json
  Keys: TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAILBOX_EMAIL

I need these Python functions:

1. get_access_token(cfg: dict) -> str
   - Posts to the token URL with grant_type=client_credentials
   - Returns the access_token string
   - Raises a clear RuntimeError if authentication fails (include status code)

2. list_emails_with_attachments(token: str, mailbox: str, subject_filter: str = None) -> list
   - GET https://graph.microsoft.com/v1.0/users/{mailbox}/messages
   - Filter: hasAttachments eq true AND isRead eq false
   - If subject_filter provided: also filter contains(subject, '{subject_filter}')
   - Returns list of {id, subject, from, receivedDateTime, hasAttachments}

3. download_attachments(token: str, mailbox: str, message_id: str, output_folder: str) -> list
   - GET the attachments for the given message_id
   - For each attachment:
       a. Check contentType — only process "application/vnd.openxmlformats..." (xlsx) 
          and "application/pdf" — SKIP others (log skipped filename)
       b. Decode base64 contentBytes
       c. Save to output_folder with original filename
       d. Return list of {filename, size_bytes, saved_path}

4. mark_email_read(token: str, mailbox: str, message_id: str) -> bool
   - PATCH the message to set isRead = true
   - Returns True on success, False on failure

5. run_email_agent(config_path: str = "config/settings.json")
   - Main function: load config → get token → list emails → 
     download attachments → mark read → log audit entry
   - AUDIT LOG: append to logs/email_agent_audit.csv:
     columns: timestamp, sender, subject, attachment_filename, file_size, saved_path, status
   - Print summary: X emails processed, Y attachments downloaded, Z rejected

All HTTP calls must have:
- timeout=30 (never hang indefinitely)
- response.raise_for_status() (catch HTTP errors)
- retry logic: if status code 429 (rate limit), wait 60 seconds and retry once

Use requests, json, base64, logging, csv, pathlib.
```

**Expected output:**
- Token obtained → first 20 chars printed as confirmation
- Email list printed with subject and sender
- Files saved to correct input folder
- audit.csv updated with each download

---

### PROMPT 4: PDF Extraction (Text File Simulation)

**Use for:** Lab 8  
**Works on .txt files in data/input/pdf_reports/ — identical logic for real PDFs**

```
I have text files in: C:\FundAutomation\data\input\pdf_reports\
Each file simulates what pdfplumber.extract_text() returns from a fund statement PDF.

Sample content structure (I will paste a section):
  Fund Name:          Alpha Capital Equity Fund
  Valuation Date:     22 June 2026
  Net Asset Value (Total):          USD    125,432,891.50
  Management Fee Rate:              1.500% per annum
  [section separator lines with ===]

I need a Python function extract_fund_statement(filepath: str) -> dict that:

1. Opens the text file and reads ALL content as a single string
2. Uses regex to extract these 6 fields:

   fund_name:      After "Fund Name:" — capture everything until newline
   nav_date:       After "Valuation Date:" — format "DD Month YYYY" (e.g. "22 June 2026")
                   Convert to Python date object using datetime.strptime(s, "%d %B %Y")
   total_nav:      After "Net Asset Value (Total):" — extract the number
                   Strip: currency codes (USD/EUR/GBP), spaces, commas
                   Convert to float
   nav_per_unit:   After "Net Asset Value per Unit:" — same stripping as above
   mgmt_fee_rate:  After "Management Fee Rate:" — extract the decimal before "%"
                   E.g. "1.500% per annum" → 0.015 (divide by 100)
   admin_fee_rate: After "Administration Fee Rate:" — same as mgmt_fee_rate

3. Return dict with keys exactly as above.
   Any field not found → set to None (never raise an error for missing fields)

4. Add these derived fields:
   fund_code:      Extract from filename: the part before the first "_" (e.g. "ALPHA001")
   source_file:    Just the filename (not full path)
   extracted_at:   datetime.now().isoformat()

5. Handle number formats:
   - "125,432,891.50" → strip commas → float(125432891.50)
   - "USD    125,432,891.50" → strip "USD" and spaces first, then strip commas

Also write batch_extract(folder: str) -> pd.DataFrame that:
- Processes every .txt file in folder
- Returns DataFrame with one row per file
- Catches exceptions per file (log and continue — never crash the batch)

For REAL PDFs: replace  text = open(filepath).read()
                  with   text = pdfplumber.open(filepath).pages[0].extract_text()
The regex extraction code is IDENTICAL — only the file reading changes.

Use Python re, datetime, pandas, pathlib. Add comments before each regex.
```

**Expected output on all 10 PDF text files:**

| FundCode | ExpectedNAV | Note |
|---|---|---|
| ALPHA001 | 125,432,891.50 | Clean — all 6 fields |
| DELTA004 | 450,000,000 * 1.02 | Slightly inflated vs Excel (recon exception) |
| All others | As per data generator | Check against EXPECTED_OUTPUT_verification.xlsx |

---

### PROMPT 5: PDF vs Excel Reconciliation

**Use for:** Lab 9  
**Fill in:** file paths from your output folder

```
I need to reconcile two fund NAV data sources to find differences.

SOURCE A — EXCEL (from my collection agent):
  File: reports\daily\master_nav_[DATE].xlsx  (use glob to find latest file)
  Key columns: FundCode, FundName, NAVDate, NAVValue, Currency

SOURCE B — PDF EXTRACT (from my extraction script):
  File: reports\daily\pdf_extracted_data.xlsx
  Key columns: fund_code, fund_name, nav_date, total_nav, currency

I need Python code using pandas to:

STEP 1 — LOAD AND STANDARDISE:
  - Load both files
  - Rename columns so both have: FundCode, FundName, NAVDate, NAVValue, Currency
  - Convert NAVDate to datetime.date in BOTH DataFrames before merging
    (Excel dates may be datetime, PDF dates may be strings — handle both)
  - Strip whitespace from FundCode in both (common source of join failures)
  - Uppercase FundCode in both (avoid case mismatch)

STEP 2 — MERGE:
  - Perform an OUTER join on ['FundCode', 'NAVDate']
  - Use suffixes: ('_Excel', '_PDF')

STEP 3 — CLASSIFY:
  - Variance = NAVValue_Excel - NAVValue_PDF (positive = Excel higher)
  - Variance_Pct = (Variance / NAVValue_Excel) * 100
  - Status column rules:
    "MATCHED"        if both values exist AND abs(Variance) <= 1000.0
    "EXCEPTION"      if both values exist AND abs(Variance) > 1000.0
    "MISSING_IN_PDF" if NAVValue_PDF is NaN (fund in Excel but not in PDF)
    "MISSING_IN_EXCEL" if NAVValue_Excel is NaN (fund in PDF but not in Excel)

STEP 4 — OUTPUT (4 separate Excel files in reports\daily\):
  1. matched_clean_[DATE].xlsx       — only Status == "MATCHED"
  2. exceptions_[DATE].xlsx          — only Status == "EXCEPTION", sorted by abs(Variance) DESC
  3. missing_in_pdf_[DATE].xlsx      — only Status == "MISSING_IN_PDF"
  4. summary_report_[DATE].xlsx      — one sheet, these stats:
       Category | Count | Total_Variance
       MATCHED  | X     | sum of abs(Variance)
       EXCEPTION| Y     | sum of abs(Variance)
       MISSING_IN_PDF | Z | N/A
       Grand Total | N | total_abs_variance

FORMATTING (use xlsxwriter for colour formatting):
  - exceptions.xlsx: rows with EXCEPTION status → red fill (#FFE0E0)
  - matched_clean.xlsx: all rows → light green fill (#E0FFE8)
  - summary_report.xlsx: header row → dark navy fill (#1A2E4A), white font

ALSO generate a plain text email summary:
  - Save to: reports\daily\recon_email_draft_[DATE].txt
  - Content: "Reconciliation complete. X matched, Y exceptions, Z missing."
    List each exception: FundCode | Excel NAV | PDF NAV | Variance

Tolerance is 1000.0 — this should come from config\settings.json (key: "NAV_TOLERANCE")
Load config at the start of the script, never hardcode tolerance.

Use pandas, openpyxl, xlsxwriter, pathlib, json.
```

**Expected outputs:**
- DELTA004 in exceptions.xlsx with ~$4,500 variance (by design in sample data)
- 9 funds in matched_clean.xlsx
- 4 funds in missing_in_pdf.xlsx (generic format funds have no PDF)
- 0 funds in missing_in_excel.xlsx

---

### PROMPT 6: SQLite / SQL Server Query

**Use for:** Lab 10

```
I need to query a fund NAV database and reconcile against Excel output.

DATABASE: SQLite at data\sql\funddb.sqlite
  (For SQL Server: replace connection line only — all other code is identical)
  
TABLE: daily_nav
  Columns: fund_code, nav_date (TEXT, ISO format "YYYY-MM-DD"), 
           total_nav (REAL), units_outstanding (REAL), 
           nav_per_unit (REAL), currency (TEXT), source (TEXT)

TABLE: nav_history  
  Columns: fund_code, nav_date (TEXT), total_nav (REAL), 
           nav_per_unit (REAL), units_outstanding (REAL)

I need Python code to:

PART 1 — DAILY NAV QUERY:
1. Connect to SQLite: conn = sqlite3.connect(config["SQL_DB_PATH"])
   (For SQL Server: conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=...'))
2. Read the valuation date from config\settings.json (key: "VALUATION_DATE")
3. Run this PARAMETERISED query (use ? placeholders — NEVER string formatting):
   SELECT fund_code, nav_date, total_nav, units_outstanding, 
          nav_per_unit, currency, source
   FROM daily_nav 
   WHERE nav_date = ?
   ORDER BY fund_code
4. Load into df_sql using pd.read_sql(query, conn, params=[valuation_date])
5. Print: row count, column names, first 3 rows

PART 2 — RECONCILE SQL vs EXCEL:
6. Load the latest master_nav_*.xlsx from reports\daily\ into df_excel
7. Apply the same reconciliation logic as Lab 9:
   - Outer join on FundCode + NAVDate
   - Variance = Excel NAV - SQL NAV
   - Classify: MATCHED / EXCEPTION / MISSING
   - Save 4-file output to reports\daily\ with "sql_" prefix
8. Expected: DELTA004 should appear as EXCEPTION (SQL has slightly different NAV by design)

PART 3 — TREND ANALYSIS:
9. Query nav_history for all funds, all 12 months
10. For each fund, calculate:
    - Latest NAV vs 30-day prior NAV
    - NAV change %
    - Min NAV in period, Max NAV in period
11. Save to reports\daily\nav_trend_analysis.xlsx

Always close the connection: use conn.close() in a finally block.
For SQL Server: use pyodbc. Install: conda install pyodbc  OR  pip install pyodbc

Use sqlite3, pandas, pathlib, json.
```

**Expected output:**
- 13 rows in df_sql (OMICRON015 has no NAV in DB — was None in generator)
- DELTA004 in sql_exceptions.xlsx with a small variance (SQL NAV is Excel NAV * 0.999)
- 12-row trend table per fund in nav_trend_analysis.xlsx

---

### PROMPT 7: Exception Engine

**Use for:** Lab 12

```
I need a rule-based exception detection engine for fund NAV data.
Rules must be CONFIGURABLE from config\exception_rules.json — 
adding a new rule requires ONLY a JSON edit, NO Python code change.

The JSON structure is:
{
  "rules": [
    {"name": "NAV_MISSING", "field": "NAVValue", "condition": "is_null", 
     "severity": "CRITICAL", "action": "Contact fund admin immediately"},
    {"name": "NAV_ZERO", "field": "NAVValue", "condition": "eq_zero", ...},
    ...
  ]
}

I need a Python class ExceptionEngine with:

__init__(self, rules_path: str):
  - Load rules from the JSON file
  - Print: "Loaded X rules from [path]"

run(self, df: pd.DataFrame) -> pd.DataFrame:
  - For each row in df, evaluate every rule
  - Return exceptions DataFrame with columns:
    FundCode, FundName, NAVDate, RuleName, Severity, FieldName, 
    ActualValue, ActionRequired, DetectedAt

  Implement these condition evaluators:
  "is_null"      → pd.isna(value) or value is None
  "eq_zero"      → value == 0
  "negative"     → isinstance(value, (int,float)) and value < 0
  "future_date"  → isinstance(value, date) and value > date.today()
  "stale_3d"     → isinstance(value, date) and (date.today() - value).days > 3
  "gt_5_pct"     → isinstance(value, (int,float)) and abs(value) > 0.05
  "gt_0025"      → isinstance(value, (int,float)) and value > 0.025
  "gt_tolerance" → isinstance(value, (int,float)) and abs(value) > 1000.0

save_outputs(self, exceptions_df: pd.DataFrame, output_folder: str):
  - Save ALL exceptions → exceptions\all_exceptions_[DATE].xlsx
  - Save CRITICAL only → exceptions\CRITICAL_ALERTS_[DATE].xlsx
  - Save summary → exceptions\exception_summary_[DATE].xlsx
    (count and % of total per rule name)
  - Apply cell colour: CRITICAL=red(FFE0E0), HIGH=orange(FFE8C0), MEDIUM=yellow(FFFAC0)
  - Print: "X total exceptions. Y CRITICAL alerts saved."

Also write a standalone run_exception_check(config_path) function that:
- Loads master_nav from reports/daily/
- Loads reconciliation exceptions from reports/daily/ 
- Merges both into one DataFrame for the engine to check
- Calls engine.run() and engine.save_outputs()

Use pandas, openpyxl, json, pathlib, datetime.
```

**Expected exceptions from sample data:**
- BETA002: DATE_FUTURE (NAVDate = tomorrow)
- DELTA004: NAV_LARGE_MOVE (>8% change vs prior day) + RECON_EXCEPTION
- OMICRON015: NAV_MISSING (CRITICAL)

**Total: 3 exceptions across 3 rules across 2 funds + 1 in CRITICAL_ALERTS**

---

### PROMPT 8: Productionise Any Script (Universal Template)

**Use for:** Lab 5 — apply to ANY script you've written

```
I have a working Python script. I need to make it production-ready 
for daily automated use. The script is:

[PASTE YOUR ENTIRE SCRIPT HERE]

Please refactor it to add ALL of the following. Do not change the 
core business logic — only add the infrastructure around it:

1. CONFIG FILE READING:
   - Load ALL paths, filenames, tolerances and settings from config\settings.json
   - Use: cfg = json.load(open("config/settings.json"))
   - Replace every hardcoded path with: cfg["KEY_NAME"]
   - Replace every hardcoded threshold with: cfg["THRESHOLD_NAME"]

2. STRUCTURED LOGGING:
   - Use Python logging module
   - Handler 1: StreamHandler (prints to console with colours if colorlog installed)
   - Handler 2: RotatingFileHandler writing to logs\[scriptname]_[DATE].log
   - Format: "%(asctime)s | %(levelname)s | %(message)s"
   - Log at these points: script start, each file processed, each error, final summary

3. ERROR HANDLING:
   - Wrap ALL file/network/database operations in try/except
   - Catch specific exceptions first (FileNotFoundError, PermissionError, ValueError)
   - Catch generic Exception last
   - Log the FULL traceback on unexpected exceptions: logging.exception("Unexpected error")
   - Script must NEVER exit with an unhandled exception

4. TIMESTAMP IN OUTPUT FILENAMES:
   - Replace fixed output names (e.g. output.xlsx) with dated names
   - Pattern: output_[YYYYMMDD].xlsx

5. COMMAND LINE ARGUMENTS:
   - Add argparse with:
     --date [YYYY-MM-DD]  (default: today)
     --dry-run            (process but don't save outputs — useful for testing)
     --verbose            (set logging level to DEBUG)

6. SUMMARY AT THE END:
   - Print and log: files_processed, records_extracted, errors_count, run_time_seconds
   - Format: "SUMMARY: 13 processed | 0 errors | 4.2 seconds"

7. GRACEFUL CRASH HANDLING:
   - Wrap entire script in if __name__ == "__main__": try/except
   - On unexpected crash: write full traceback to logs\crash_[TIMESTAMP].log
   - Print: "Script crashed. See logs\crash_*.log for details."
```

---

### PROMPT 9: Debug Any Error (Universal Template)

**Use when:** Any code gives you an error

```
I ran Python code and got an error. Please help me understand and fix it.

PYTHON VERSION: [type: python --version in terminal and paste result]
ENVIRONMENT: Anaconda on Windows

MY CODE:
[PASTE YOUR COMPLETE SCRIPT — do not cut it down, include all imports]

THE EXACT ERROR MESSAGE:
[PASTE THE COMPLETE ERROR — starting from "Traceback (most recent call last)"]

WHAT I WAS TRYING TO DO:
[One sentence: what should this script do?]

WHAT I'VE ALREADY TRIED:
[List any changes you already made]

Please:
1. Explain what caused this error in plain English — no Python jargon
2. Show me the corrected code with ONLY the fix (don't rewrite the whole script)
3. Explain what you changed and why in one sentence each
4. Tell me how to prevent this type of error in future

If there are multiple possible causes, explain the most likely one first.
```

---

## LAB-BY-LAB EXPECTED OUTPUT VERIFICATION

Use this table to verify your scripts are working correctly.

### LAB 2: Excel Collection Agent

| Fund | Expected NAV | Expected Date | Should Succeed? |
|---|---|---|---|
| ALPHA001 | 125,432,891.50 | 22-Jun-2026 | YES |
| BETA002 | 88,210,445.00 | 23-Jun-2026 (tomorrow!) | YES — but flag DATE_FUTURE |
| GAMMA003 | 210,750,000.00 | 22-Jun-2026 | YES |
| DELTA004 | 450,000,000.00 | 22-Jun-2026 | YES — but flag LARGE_MOVE |
| EPSILON005 | 175,600,000.00 | 22-Jun-2026 | YES |
| NU013 | — | — | NO — wrong sheet name |
| OMICRON015 | — | — | NO — NAV cell is blank |

**Validation command:**
```python
# Run this after your collection agent to verify:
import pandas as pd
df = pd.read_excel("reports/daily/master_nav_*.xlsx")  # use glob
assert len(df) == 5, f"Expected 5 Citco funds, got {len(df)}"
assert df[df.FundCode=='ALPHA001']['NAVValue'].values[0] == 125432891.50
print("✓ Citco collection verified")
```

---

### LAB 3: Dynamic Discovery

**Check confidence scores:**
- All 5 Citco files: score ≥ 90
- All 5 SS&C files: score ≥ 85
- LAMBDA011, MU012, XI014: score ≥ 75
- NU013: score = 0 or 25 (no NAV found — sheet has wrong label)
- OMICRON015: score = 25 (date found, NAV blank)

---

### LAB 8: PDF Extraction

Run and verify DELTA004:
```python
df = pd.read_excel("reports/daily/pdf_extracted_data.xlsx")
delta_pdf = df[df.fund_code=='DELTA004']['total_nav'].values[0]
delta_excel = 450_000_000.00
variance = abs(delta_excel - delta_pdf)
print(f"DELTA004 variance: {variance:,.2f}")
# Should be around 4,500 — this creates the reconciliation exception in Lab 9
```

---

### LAB 9: Reconciliation

**Your exceptions.xlsx should contain:**
```
FundCode   | NAVValue_Excel    | NAVValue_PDF      | Variance  | Status
DELTA004   | 450,000,000.00    | 450,004,500.00    | -4,500    | EXCEPTION
```

**Your missing_in_pdf.xlsx should contain:**
```
LAMBDA011, MU012, NU013, XI014, OMICRON015 (generic format — no PDF generated for these)
```

---

### LAB 12: Exception Engine

**CRITICAL_ALERTS file must contain exactly these funds:**
```
OMICRON015 — NAV_MISSING (CRITICAL)
```

**all_exceptions file must contain:**
```
BETA002    — DATE_FUTURE    (HIGH)
DELTA004   — NAV_LARGE_MOVE (HIGH) + RECON_EXCEPTION (HIGH)
OMICRON015 — NAV_MISSING    (CRITICAL)
```

---

## BREAK IT — TEST SCRIPTS

For every lab, run these tests before declaring it done:

### Test Set A: File System Tests
```python
# Test 1: Empty folder
import shutil, os
os.makedirs("test_empty", exist_ok=True)
# run your script with folder="test_empty"
# Expected: "0 files found, exiting cleanly" — NOT a crash

# Test 2: Read-only file
import stat
os.chmod("data/input/format_a_citco/ALPHA001_..._NAV.xlsx", stat.S_IREAD)
# run your script
# Expected: PermissionError caught, file logged, others continue
os.chmod("data/input/format_a_citco/ALPHA001_..._NAV.xlsx", stat.S_IWRITE)  # restore

# Test 3: Wrong column name
# Open ALPHA001, rename sheet "NAV Summary" to "NAV_Summary" (add underscore)
# run your script
# Expected: sheet not found error, file in errors log, others succeed
```

### Test Set B: Data Quality Tests
```python
# Test 4: Text in numeric cell
# Open ALPHA001, type "PENDING" in cell B5 (the NAV cell)
# run your script
# Expected: ValueError caught, fund in errors.xlsx, others succeed

# Test 5: Run twice
# run your script once, then run it again without moving files
# Expected: output is identical — not doubled rows

# Test 6: Missing output folder
import shutil
shutil.rmtree("reports/daily/")  # delete the output folder
# run your script
# Expected: folder created automatically OR clear FileNotFoundError with helpful message
os.makedirs("reports/daily/", exist_ok=True)  # restore
```

---

## COPILOT DEBUGGING WORKFLOW

When your code doesn't work, follow this exact sequence:

```
STEP 1: Read the error message from BOTTOM to TOP
        The last line tells you WHAT went wrong
        The lines above tell you WHERE

STEP 2: Open Copilot Chat (Ctrl+Shift+I)

STEP 3: Use Prompt 9 (Debug Template) — paste:
        - Your complete code
        - The complete error message (Ctrl+A in terminal, Ctrl+C)
        - What you were trying to do

STEP 4: Read Copilot's explanation BEFORE looking at the fix
        Can you understand what went wrong?

STEP 5: Apply the fix. Run again.

STEP 6: If it fails again: add this at the TOP of the script:
        import traceback
        try:
            [your code here]
        except Exception as e:
            traceback.print_exc()
        This gives you the FULL error chain.
```

### 10 Most Common Errors and How to Fix Them

| Error | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named 'X'` | Package not installed | `pip install X --break-system-packages` or `conda install X` |
| `FileNotFoundError: [Errno 2]` | File path is wrong or file doesn't exist | Print the path before using it: `print(filepath)`. Use `pathlib.Path.exists()` |
| `KeyError: 'NAV Summary'` | Sheet name doesn't match (case-sensitive!) | Print `wb.sheetnames` to see actual names |
| `AttributeError: 'NoneType' object has no attribute` | A variable is None when you expect a value | Add a None check: `if value is not None:` |
| `ValueError: could not convert string to float` | Cell has text like "1,234" or "N/A" | Strip: `float(str(val).replace(',','').strip())` |
| `PermissionError: [Errno 13]` | File is open in Excel | Close the file in Excel first |
| `pandas.errors.MergeError` | Merging on columns with different dtypes | Convert both to same type before merge |
| `AADSTS700016` | Azure AD App ID wrong | Copy ID fresh from Azure Portal |
| `sqlite3.OperationalError: no such table` | Wrong database path | `print(sqlite3.connect(path).execute("SELECT name FROM sqlite_master").fetchall())` |
| `RecursionError` | Infinite loop | Check your loop exit conditions |

---

## THE ENTERPRISE UPGRADE CHECKLIST

Run this against EVERY script before calling it production-ready:

```
REQUIRED (must have):
□ All paths read from config/settings.json — NO hardcoded paths
□ Python logging module used — NOT just print()
□ Log file in logs/ folder with daily rotation
□ try/except around every file, network, and database call
□ Script continues after individual file failures (does not crash)
□ Output filenames include date: output_20240622.xlsx
□ Summary logged at the end: X processed, Y failed, Z seconds

PRODUCTION (should have):
□ Duplicate detection before processing (MD5 hash check)
□ Audit CSV: every action logged with timestamp and status
□ Email notification on completion and on CRITICAL failure
□ Command-line arguments for date, dry-run mode, verbosity
□ Can be scheduled: run.bat or cron job created and tested
□ README.md explaining what the script does and how to run it

ENTERPRISE (eventually have):
□ Configuration validated at startup (check all paths exist)
□ Automated tests for at least the happy path
□ Version number tracked in script header
□ Code reviewed by a second person before deploying to production
□ Monitoring: alert if script does NOT run (not just if it fails)
```

---

## YOUR 30-DAY CHALLENGE

By the end of the workshop, write down ONE process you will automate:

```
My automation target: ___________________________________

What it does today (manual): ____________________________

What it will do automated: ______________________________

Estimated time saved per week: __________________________

Deadline to have it running: ____________________________
```

**Week 1:** Build v1 using today's prompts. Get it working on sample data.  
**Week 2:** Run it in parallel with your manual process. Verify outputs match.  
**Week 3:** Apply the enterprise upgrade checklist. Schedule it.  
**Week 4:** Hand off to your team. Document it. Share the prompt templates.

---

## QUICK REFERENCE: KEY FILE PATHS

```
Workshop root:          C:\FundAutomation\
Fund data (Citco):      data\input\format_a_citco\     ← 5 funds, Format A
Fund data (SS&C):       data\input\format_b_ssc\       ← 5 funds, Format B  
Fund data (Generic):    data\input\format_c_generic\   ← 5 funds, Format C (2 broken)
PDF simulations:        data\input\pdf_reports\        ← 10 text files
Reference master:       data\reference\fund_master.xlsx
Expected answers:       data\reference\EXPECTED_OUTPUT_verification.xlsx
SQLite database:        data\sql\funddb.sqlite
Config:                 config\settings.json
Exception rules:        config\exception_rules.json
Pipeline config:        config\pipeline_config.json
Your scripts Day 1:     scripts\day1\
Your scripts Day 2:     scripts\day2\
Log files:              logs\
Reports:                reports\daily\ and reports\exceptions\
```

---

*Keep this workbook open throughout the workshop. Every prompt you need is here.*
