# Workshop Prompts — Day 2
## Fund Operations Automation · Alternative Investment Management

> **Before starting:** All your files are in `C:\FundAutomation\`  
> Open VS Code, open that folder, open a new terminal (`Ctrl+`` `), and keep Copilot Chat open (`Ctrl+Shift+I`).

---

# PART 1 — PDF EXTRACTION

---

## Prompt 1 · Open One PDF and Read It

We start with one file. Paste this, run the code, and see exactly what pdfplumber reads from the PDF.

```
I want to open one PDF file and print its text content so I can see what's inside.

The file is at:
C:\FundAutomation\data\input\pdf_reports\ALPHA001_Statement_20260622.pdf

Write Python code that:
1. Installs pdfplumber if needed (pip install pdfplumber)
2. Opens that one PDF file
3. Extracts and prints all the text from page 1

That's it. Just print the raw text so I can see what we're working with.
```

**What you will see** — the raw text of the fund statement, including the NAV, fees, and date.

---

## Prompt 2 · Pull Out the NAV Number

Now we extract the one number we care about most. The NAV appears in the text as:
`Net Asset Value (Total) USD 125,432,891.50`

```
I have text extracted from a fund statement PDF.
The text contains a line like this:
  "Net Asset Value (Total) USD 125,432,891.50"

Write a Python function called get_nav(text) that:
1. Searches the text for the pattern "Net Asset Value (Total)"
2. Extracts the number that follows — stripping "USD" and any commas
3. Returns it as a plain number (float), e.g. 125432891.50
4. Returns None if it cannot find it — never crashes

Test it by running it on the text from ALPHA001_Statement_20260622.pdf
and printing the result.
```

**Expected result:** `125432891.5`

---

## Prompt 3 · Add the Date and the Fee Rate

```
I have the get_nav() function from the previous step.

Now add two more functions to the same file:

get_date(text)
  The date appears in the text after the words "period ended"
  For example: "...constitutes the official NAV for the period ended\n22 June 2026."
  Extract "22 June 2026" and convert it to a Python date object.
  Return None if not found.

get_fee_rate(text)
  The management fee appears as: "Management Fee Rate 1.500% per annum"
  Extract the number 1.500, divide by 100, and return 0.015 (a decimal, not a percentage).
  Return None if not found.

Test all three functions on ALPHA001_Statement_20260622.pdf and print:
  NAV: 125,432,891.50
  Date: 2026-06-22
  Fee Rate: 0.015
```

---

## Prompt 4 · Get the Fund Code from the Filename

```
The PDF filename tells us the fund code.
For example: ALPHA001_Statement_20260622.pdf → fund code is ALPHA001

Add a function get_fund_code(filename) that:
1. Takes just the filename (not the full path), e.g. "ALPHA001_Statement_20260622.pdf"
2. Returns the part before the first underscore: "ALPHA001"

Now combine all four functions into one function called extract_one_pdf(filepath):
  - Opens the PDF at filepath
  - Calls get_fund_code, get_nav, get_date, get_fee_rate
  - Returns a dictionary with keys: FundCode, NAV, Date, FeeRate, FileName
  - If any field is missing, set it to None — never crash

Test it on ALPHA001 and print the dictionary.
```

---

## Prompt 5 · Process All 10 PDFs and Save to Excel

```
I have the extract_one_pdf() function working on one file.

Now scale it up:

Write a function called extract_all_pdfs(folder_path) that:
1. Finds every .pdf file in the folder:
   C:\FundAutomation\data\input\pdf_reports\
2. Calls extract_one_pdf() on each one
3. Collects the results into a list
4. Converts the list to a pandas DataFrame
5. Saves it to: C:\FundAutomation\reports\daily\pdf_extract.xlsx
6. Prints one line per file: "✓ ALPHA001  NAV: 125,432,891.50" or "✗ FILENAME  not found"
7. At the end prints: "Done. 10 extracted. 0 failed."

Run it now.
```

**After running:** Open `pdf_extract.xlsx`. You should see 10 rows — one per fund.

---

# PART 2 — PDF vs EXCEL RECONCILIATION

---

## Prompt 6 · Load Both Sources

We now have two sources of NAV data. Let's load both and look at them side by side.

```
I have two Excel files:

FILE 1 — PDF extract (what the administrator sent):
  C:\FundAutomation\reports\daily\pdf_extract.xlsx
  Columns: FundCode, NAV, Date, FeeRate, FileName

FILE 2 — Our internal Excel records (what our collection agent found):
  C:\FundAutomation\data\input\format_a_citco\
  (there are 5 xlsx files here — we need to read NAV from cell B5 of sheet "NAV Summary"
   and fund code from the filename in each file)

Write Python code that:
1. Loads FILE 1 into a variable called df_pdf
2. Reads all 5 Excel files from the Citco folder, extracts FundCode and NAV from each,
   and loads them into a variable called df_excel
   (FundCode = first part of filename before underscore, NAV = cell B5)
3. Prints both DataFrames so I can see them

No comparison yet — just load and display.
```

---

## Prompt 7 · Match Them by Fund Code

```
I have df_pdf and df_excel from the previous step.

Now match them:

1. Join df_pdf to df_excel using FundCode as the key
   Use an outer join so we don't lose any fund that appears in only one source
2. After joining, rename the NAV columns:
   NAV from the PDF → NAV_PDF
   NAV from the Excel → NAV_Excel
3. Print the joined table

How many rows does the result have? Are all fund codes present in both sources?
```

---

## Prompt 8 · Calculate the Differences

```
I have the joined table from the previous step.

Now calculate:

1. Difference = NAV_Excel minus NAV_PDF  (positive = Excel is higher)
2. Difference_Pct = (Difference / NAV_Excel) × 100  (percentage)
3. Round both to 2 decimal places

Add a Status column:
  "MATCHED"   — if the absolute difference is less than or equal to 1,000
  "EXCEPTION" — if the absolute difference is greater than 1,000
  "MISSING"   — if either NAV_PDF or NAV_Excel is blank

Print the full table sorted so EXCEPTION rows appear first.

Which fund has the largest difference?
```

**You will see DELTA004 with a $9,000,000 difference — the PDF says $459M, Excel says $450M.**

---

## Prompt 9 · Save with Colour Coding

```
I have the reconciliation table from the previous step.

Save it to Excel with colour coding:
  File: C:\FundAutomation\reports\daily\recon_result.xlsx

Formatting rules:
  EXCEPTION rows → red background on the entire row  (colour: #FFD0D0)
  MATCHED rows   → green background on the entire row (colour: #D0FFD8)
  MISSING rows   → yellow background on the entire row (colour: #FFFACC)

Use openpyxl for the formatting.
Make the header row bold with a dark navy background (#1A2E4A) and white text.
Set column widths so everything is readable without scrolling.

Open the file after saving to confirm the colours are correct.
```

---

## Prompt 10 · Add a Summary Tab

```
Add a second sheet to recon_result.xlsx called "Summary".

On the Summary sheet, show:

  Reconciliation Date:   [today's date]
  Total Funds:           [count of all rows]
  Matched:               [count of MATCHED rows]
  Exceptions:            [count of EXCEPTION rows]
  Missing:               [count of MISSING rows]
  Largest Difference:    [the fund code and amount of the biggest exception]
  Total Variance:        [sum of all absolute differences]

Make it look like a clean management summary — bold labels, formatted numbers.
This is what you would show your manager.
```

---

# PART 3 — POWER BI INTEGRATION

---

> **Setup required before these prompts.** See `Setup_Guide.md` — Power BI section.  
> Estimated setup time: 15 minutes.

---

## Prompt 11 · Prepare Clean Data for Power BI

Power BI needs data in a specific format. This prompt creates a file Power BI can connect to directly.

```
I want to feed my reconciliation data into Power BI Desktop.

Power BI works best with clean, flat data — no merged cells, no formatting,
just plain columns and rows.

Take the reconciliation table we built and create a new file:
  C:\FundAutomation\reports\powerbi\fund_nav_data.xlsx

This file should have ONE sheet called "Data" with these columns:
  FundCode, FundName, NAV_PDF, NAV_Excel, Difference, Difference_Pct,
  Status, Date, FeeRate, ReportDate (today's date)

Rules:
  - No merged cells
  - No colour formatting — Power BI applies its own
  - All dates in YYYY-MM-DD format
  - All numbers as plain floats — no currency symbols or commas in the cells
  - Column names use underscores not spaces

Print: "Power BI data file ready: X rows written"
```

---

## Prompt 12 · Add a Python Visual Script for Power BI

This script runs inside Power BI Desktop as a Python Visual — it draws a chart directly in your report.

```
I want to add a Python-powered chart inside Power BI Desktop.

Write a Python script that Power BI Desktop can use as a "Python visual".
In Power BI Desktop, Power BI passes data to the script as a variable called "dataset".

The script should:
1. Take the dataset variable (it will have columns: FundCode, NAV_Excel, Status)
2. Draw a horizontal bar chart showing NAV_Excel per fund
3. Colour each bar:
   MATCHED   → green  (#27AE60)
   EXCEPTION → red    (#C0392B)
   MISSING   → grey   (#95A5A6)
4. Add the fund code as labels on each bar
5. Title: "Fund NAV — Reconciliation Status"
6. Use matplotlib — Power BI Desktop has this built in

Format the numbers on the x-axis as millions (e.g. 125M not 125000000).
```

**In Power BI Desktop:** Insert → Python Visual → paste this script.

---

## Prompt 13 · Generate a Filtered Report Extract Per Fund

```
Some LPs or trustees only need to see their own fund's data.

Write a Python script that:
1. Reads fund_nav_data.xlsx from C:\FundAutomation\reports\powerbi\
2. For each unique FundCode in the file:
   a. Filters the data to just that fund's rows
   b. Creates a formatted Excel file for that fund:
      File name: C:\FundAutomation\reports\fund_extracts\[FundCode]_Report_[Date].xlsx
   c. The Excel should look like a clean one-page summary:
      Fund name at top (large, bold)
      NAV from PDF, NAV from Excel, Difference — each on its own labelled row
      Status highlighted in the appropriate colour
      Date at the bottom
3. Prints: "Created extract for ALPHA001" for each fund

Create the folder C:\FundAutomation\reports\fund_extracts\ if it does not exist.
```

---

## Prompt 14 · Email the Report to the Right Person

```
I want to email each fund's extract to the right person automatically.

I have a distribution list in Excel at:
  C:\FundAutomation\data\reference\fund_master.xlsx
Use the FundCode column to match funds.
For this exercise, use a single test email address for all funds: [YOUR EMAIL HERE]

Write Python code that:
1. Reads the list of fund extract files from C:\FundAutomation\reports\fund_extracts\
2. For each file, sends an email using my Windows Outlook (already open on this machine)
3. Subject: "[FundCode] — NAV Reconciliation — [Date]"
4. Body: "Please find attached the NAV reconciliation for [FundCode] as at [Date].
          Status: [MATCHED / EXCEPTION].
          For queries contact fundops@company.com"
5. Attaches the Excel file for that fund
6. Logs each email sent to: C:\FundAutomation\logs\email_log.txt

Use the win32com.client library which works with Outlook on Windows.
Do not put any passwords in the code — use my existing Outlook session.

Before sending, print what it is about to send and ask me to type YES to confirm.
```

---

# PART 4 — YOUR OWN PDF · GENERIC WALKTHROUGH

---

> From here, use your own PDF and your own Excel file.  
> These prompts work for any document — trustee reports, LP statements,  
> compliance certificates, rent rolls, valuation reports.

---

## Prompt 15 · Show Copilot Your PDF

Open your PDF. Copy any 30 lines from the section that contains the data you need. Paste below this prompt.

```
I am going to paste text from a PDF I receive regularly at work.
Please read it carefully and then answer these questions:

[PASTE YOUR PDF TEXT HERE — 20 to 40 lines from the relevant section]

Based on what you can see:
1. What is the overall structure of this document?
2. Where is the main value I would want to extract — what label appears before it?
3. What is the date format used?
4. Are there any tables or is it all label/value pairs?
5. What might make extraction difficult?

Do not write any code yet. Just tell me what you observe.
```

---

## Prompt 16 · Extract the Fields You Need

```
Based on your analysis of my PDF:

Write a Python function called extract_my_pdf(filepath) that extracts these fields:
  [List exactly what you need — e.g. "Total NAV", "Valuation Date", "Fund Name", "Fee Rate"]

For each field, use the pattern you identified from the text I showed you.

Rules:
  - If a field is not found, return None for it — do not crash
  - Strip any currency symbols, commas, and extra spaces from numbers
  - Convert dates to YYYY-MM-DD format
  - Return a dictionary with one key per field

Test it on the file I give you:
  [YOUR PDF FILE PATH HERE]

Print the extracted dictionary so I can verify it looks correct.
```

---

## Prompt 17 · Load Your Internal Excel

```
I also have an internal Excel file that contains the data I want to compare against the PDF.

The file is at: [YOUR EXCEL FILE PATH]
The sheet I need is: [YOUR SHEET NAME]
The columns I care about are: [LIST THE COLUMNS]

Write Python code that:
1. Opens this Excel file
2. Reads the relevant sheet
3. Loads it into a table (one row per fund or per record)
4. Prints the first few rows so I can confirm it loaded correctly

If there are blank rows or merged header rows at the top, skip them.
```

---

## Prompt 18 · Match and Compare

```
I now have:
  df_pdf    — data extracted from my PDFs (built in Prompt 16)
  df_excel  — data from my internal Excel (built in Prompt 17)

Match them using: [TELL COPILOT WHICH COLUMN TO USE AS THE KEY — e.g. Fund Code, Property Name, LP Name]

After matching:
1. Calculate the difference between the PDF value and the Excel value
2. Flag as EXCEPTION if the difference is greater than [YOUR TOLERANCE — e.g. 1000 or 0.1%]
3. Flag as MISSING if a record appears in one source but not the other
4. Sort the results: EXCEPTION first, then MISSING, then MATCHED

Print the first 10 rows of the result.
```

---

## Prompt 19 · Output Your Exception Report

```
Save the reconciliation results to Excel:
  File: [YOUR OUTPUT PATH]\recon_[today's date].xlsx

Sheet 1 — "Exceptions": only the rows where Status is EXCEPTION or MISSING
Sheet 2 — "All Results": every row, colour coded (red/green/yellow)
Sheet 3 — "Summary": counts and totals — same format as Prompt 10

Also generate a plain text summary I can paste into an email or Teams message:
  "[DATE] Reconciliation
   Matched: X | Exceptions: Y | Missing: Z
   Largest exception: [description]
   Action required: [list each exception with the difference amount]"
```

---

# PART 5 — PRODUCTIONISE IT

---

## Prompt 20 · Move Paths to a Config File

One change, anywhere, updates everything.

```
I have a working script. I want to make it easier to maintain.

Right now the file paths are written directly in the code.
I want to move all paths and settings to a separate config file so I only
have to change one file — not hunt through the code.

Do two things:

1. Create a file called config.json in the same folder as my script.
   It should contain all the paths and settings currently hardcoded in the script:
   folder paths, file names, tolerance values, email addresses.

2. Update my script to read from config.json at the start using:
   import json
   cfg = json.load(open("config.json"))
   And replace every hardcoded value with cfg["KEY_NAME"]

Show me the config.json content and show me only the changed lines in the script.
```

---

## Prompt 21 · Add a Log File

```
Add logging to my script so I always know what happened when I run it.

Use Python's built-in logging module.
Write log messages to: C:\FundAutomation\logs\recon_[today's date].log

Log these events:
  - Script started (with timestamp)
  - Each file processed: SUCCESS or FAILED, with the value extracted
  - Each exception found: fund code and difference amount
  - Script finished: total processed, total exceptions, time taken

Also keep printing to the screen — the log is a backup record, not a replacement.

At the very end, print:
  "Log saved to: [log file path]"
```

---

## Prompt 22 · Make It Run on a Schedule

```
I want this script to run automatically every morning at 8:00 AM
without me having to open it or click anything.

Write a file called run_recon.bat that:
1. Opens the Anaconda Python environment
2. Runs my script from the correct folder
3. Saves any errors to a crash log

Then tell me step by step how to set this up in Windows Task Scheduler
so it runs every weekday at 8:00 AM.
Write the instructions as a numbered list — no technical jargon.
```

---

# PART 6 — SQL INTEGRATION

---

> **Setup required before these prompts.** See `Setup_Guide.md` — SQL section.  
> Estimated setup time: 5 minutes (SQLite needs no installation).

---

## Prompt 23 · Query Today's NAVs from the Database

```
I have a database file at:
  C:\FundAutomation\data\sql\funddb.sqlite

It contains a table called daily_nav with these columns:
  fund_code, nav_date, total_nav, units_outstanding, nav_per_unit, currency

Write Python code that:
1. Connects to this database file using Python's built-in sqlite3 library
   (no installation needed — sqlite3 is already in Python)
2. Queries all rows where nav_date = '2026-06-22'
3. Loads the result into a table
4. Prints it — one line per fund, showing fund_code and total_nav

This is our third source of NAV data alongside the PDFs and the Excel files.
```

**Expected:** 13 rows. ALPHA001 NAV = 125,432,891.50

---

## Prompt 24 · Three-Way Reconciliation: SQL vs PDF vs Excel

```
I now have NAV data from three sources for the same date:
  df_sql   — from the database query (Prompt 23)
  df_pdf   — from the PDF extraction (Prompt 5)
  df_excel — from the Excel collection (Prompt 6)

Join all three on fund_code.

For each fund, calculate:
  SQL_vs_PDF_Diff   = total_nav (SQL) minus NAV_PDF
  SQL_vs_Excel_Diff = total_nav (SQL) minus NAV_Excel
  PDF_vs_Excel_Diff = NAV_PDF minus NAV_Excel

Flag as THREE_WAY_EXCEPTION if any of the three differences exceeds 1,000.
Flag as CONSISTENT if all three differences are within 1,000.

Print the results. Which fund shows a difference across all three sources?
Save to: C:\FundAutomation\reports\daily\three_way_recon.xlsx
```

**You will see DELTA004: Excel=$450M, PDF=$459M, SQL=$449.5M — three different values.**

---

## Prompt 25 · 30-Day NAV Trend from the Database

```
The database also has a table called nav_history with monthly NAV data
going back 12 months for each fund.

Columns: fund_code, nav_date, total_nav, nav_per_unit, units_outstanding

Write Python code that:
1. Queries all rows from nav_history for fund_code = 'ALPHA001'
2. Sorts by nav_date oldest to newest
3. Calculates month-on-month change: (this_nav / previous_nav - 1) × 100
4. Prints a simple table:
   Date         | Total NAV       | Month Change
   2025-08-22   | 106,250,957.80  | —
   2025-09-21   | 107,313,467.38  | +1.00%
   ...

Then run the same query for DELTA004 — which month had the largest movement?
```

---

## Prompt 26 · Database + Exception History

```
I want to save every exception we find to the database so I can track patterns over time.

Add a table to funddb.sqlite called exception_history:
  fund_code TEXT
  exception_date TEXT
  exception_type TEXT  (PDF_VS_EXCEL, SQL_VS_PDF, SQL_VS_EXCEL, MISSING)
  difference_amount REAL
  difference_pct REAL
  detected_at TEXT  (timestamp when we ran the script)
  status TEXT       (OPEN, RESOLVED)

Write Python code that:
1. Creates this table if it does not already exist
2. After every reconciliation run, inserts each exception found as a new row
3. At the start of the next run, queries: "Has this fund had exceptions in the last 5 runs?"
   If yes, prints a warning: "⚠ DELTA004 has had exceptions in 3 of the last 5 runs"

This turns your one-off reconciliation into an ongoing monitoring tool.
```

---

# PART 7 — REMAINING USE CASES

---

## Prompt 27 · Outlook — Automatically Download Attachments

```
Every morning I receive Excel files as email attachments from our fund administrators.
I want to download them automatically without opening Outlook manually.

Write Python code using the Microsoft Graph API that:
1. Reads my credentials from config.json:
   Keys needed: TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAILBOX_EMAIL
   (See Setup_Guide.md for how to get these — takes 10 minutes in Azure Portal)
2. Connects to my mailbox
3. Finds all unread emails with Excel attachments received today
4. Downloads each attachment to:
   C:\FundAutomation\data\input\downloads\[today's date]\
5. Marks the email as read after downloading
6. Logs each download: sender, subject, filename, size, time

Run this first thing in the morning before the reconciliation script.
Use only the requests library — no Microsoft SDK needed.
```

---

## Prompt 28 · SharePoint — Upload Your Reports Automatically

```
After my reconciliation runs, I want to upload the output files to SharePoint
so my team can access them without me sending emails.

Write Python code that:
1. Reads SharePoint credentials from config.json:
   Keys needed: TENANT_ID, CLIENT_ID, CLIENT_SECRET, SHAREPOINT_SITE_URL
2. Connects to the SharePoint site
3. Uploads these files to the SharePoint document library "Fund Reports":
   - recon_result.xlsx
   - three_way_recon.xlsx
   - Any file in reports\fund_extracts\
4. Creates a subfolder for today's date if it does not exist:
   Fund Reports / 2026-06-22 /
5. Prints: "Uploaded: [filename] → [SharePoint path]"

Use the Microsoft Graph API with the requests library.
```

---

## Prompt 29 · The Full Pipeline — One Command Does Everything

```
I want one script that runs the entire workflow in the right order:

Step 1: Download today's attachments from Outlook
Step 2: Extract NAV from all PDFs in the downloads folder
Step 3: Read Excel NAV files from the Citco and SS&C folders
Step 4: Query today's NAVs from the database
Step 5: Run the three-way reconciliation (PDF vs Excel vs SQL)
Step 6: Save exceptions to the exception_history table
Step 7: Export fund extracts (one Excel per fund)
Step 8: Email exceptions to: [address from config.json]
Step 9: Upload reports to SharePoint
Step 10: Print a final summary

Write a script called run_pipeline.py that runs all ten steps in order.
If any step fails, log the error and continue to the next step — never stop completely.
At the end, print:
  "Pipeline complete. Steps passed: X/10. Exceptions found: Y. Reports sent: Z."

Load all settings from config.json — no hardcoded values anywhere in the script.
```

---

# PROMPT CARDS — QUICK REFERENCE

---

## When Something Goes Wrong

```
My script produced an error. Here is the complete error message:

[PASTE THE FULL ERROR — everything from "Traceback" to the last line]

Explain what went wrong in plain language.
Show me only the lines that need to change to fix it.
Do not rewrite the whole script.
```

---

## When the Output Is Wrong But There Is No Error

```
The script ran without errors but the output is not right.

What I expected: [describe]
What I got: [paste a few rows from the output]
What the source file looks like: [paste a few rows from the input]

Find the problem and fix only that specific issue.
```

---

## When You Want to Add a Feature

```
I have a working script. I want to add one new thing:

[Describe the new feature in one sentence — e.g.
"Add a column showing what percentage of the NAV the fee represents"
OR "Also send a Teams message when there is a CRITICAL exception"]

Add only this feature. Do not change anything else.
Show me only the new lines.
```

---

## When You Want to Understand What the Code Is Doing

```
I want to understand what this part of my code does.
Please explain it in plain English — no technical terms:

[PASTE THE SECTION OF CODE YOU WANT EXPLAINED]

Explain it as if I am a fund accountant who has never seen Python before.
Tell me what it does and why it is written that way.
```
