# Regal Extract

Regal Extract is a local Windows web app for extracting key values from ICEGATE shipping-bill PDFs and appending them into a master Excel sheet.

It is designed for private offline-style use:

- no OpenAI API
- no cloud upload
- no paid subscription
- no data sent out by the app itself

The app runs in your browser on:

`http://127.0.0.1:8765`

## What the app does

- reads one or more shipping-bill PDFs
- extracts common export fields using fixed PDF rules
- shows the extracted rows in an editable review screen
- appends new rows into a master Excel file
- updates existing rows when the same Bill No. already exists
- preserves user-created Excel columns and formulas

## Main extracted fields

The app fills these standard columns:

- Particulars
- Bill No.
- Date
- SB NO.
- SB DATE
- Port Code
- $ / E
- Exch Rate
- INR
- Value@5%
- IGST@5%
- FOB
- DBK
- RoDTEP

### Simple meaning of the tax columns

- `INR` = foreign amount multiplied by exchange rate
- `Value@5%` = IGST taxable value taken from the customs PDF
- `IGST@5%` = 5% IGST amount from the PDF

Sometimes `INR` and `Value@5%` are the same. Sometimes they are different. That depends on the figures present in the PDF.

## What is inside the zip

- `app.py` - local web server
- `extractor.py` - PDF extraction logic
- `workbook_service.py` - Excel append/update logic
- `static/` - front-end files
- `setup_and_start.bat` - first-time setup and launch
- `start_app.bat` - normal launch after setup
- `requirements.txt` - Python packages needed

## System requirements

- Windows
- Python 3.11 or newer installed and available as `python`
- Internet only for the first package installation step

Python libraries used:

- `openpyxl`
- `pdfplumber`
- `pypdf`

## First-time setup

1. Extract the zip to any folder.
2. Open the extracted `Regal Extract` folder.
3. Double-click `setup_and_start.bat`.
4. Let it install the required Python packages.
5. The app will start automatically after installation.

If setup succeeds, a browser window/tab should open on:

`http://127.0.0.1:8765`

## Normal daily use

After the first setup:

1. Open the `Regal Extract` folder.
2. Double-click `start_app.bat`.
3. Use the app in your browser.

To stop the app, close the command window or press `Ctrl + C`.

## How to use the app

### 1) Choose the master workbook

You can either:

- create a new master workbook, or
- upload an existing master workbook

The active master is stored locally as:

`data/master.xlsx`

### 2) Upload PDFs

Upload one or more supported ICEGATE shipping-bill PDFs.

The app extracts values and shows them in a review table before anything is written to Excel.

### 3) Review the extracted rows

Before append, you can edit any extracted value manually in the table.

This is useful when:

- a PDF layout is slightly different
- a name or address needs correction
- a number should be verified before writing to the master

### 4) Append to the master workbook

When you click append:

- new Bill No. values are added as new rows
- existing Bill No. values are updated in recognized columns only
- user-created columns are preserved
- formulas in manual columns are carried into genuinely new rows

## Excel safety behavior

This app is built to reduce accidental damage to the master workbook.

- Columns are matched by header names, not hard-coded letters.
- Existing rows are detected by `Bill No.` / invoice number.
- Only recognized extraction columns are updated.
- Extra custom columns are not wiped.
- Saves are atomic to reduce corruption risk.
- When a stored master is replaced, the previous master is backed up.

## Recognized headers

The app can work with these headers directly:

- `Particulars`
- `Bill No.`
- `Date`
- `SB NO.`
- `SB DATE`
- `Port Code`
- `$ / E`
- `Exch Rate`
- `INR`
- `Value@5%`
- `IGST@5%`
- `FOB`
- `DBK`
- `RoDTEP`

It also supports common header aliases such as:

- `Invoice No.`
- `Shipping Bill No.`
- `Shipping Bill Date`
- `Exchange Rate`
- `Drawback`

## Notes about extraction quality

The extractor is rule-based, not AI-based. That means:

- it is fast and private
- it works well for the targeted PDF layout
- it can need review if the PDF format changes

The app shows field warnings when some values could not be confidently read.

The current build also handles placeholder consignee text more safely. For example, if a PDF contains a line like `FOR THE ORDER OF : -`, the app avoids treating `:-` as a real party name and instead tries to read the actual buyer/consignee block.

## Local data folders

The app creates and uses these local folders under `data/`:

- `data/master.xlsx` - active working master file
- `data/backups/` - backup copies of previous masters
- `data/pending/` - temporary extracted batch data
- `data/audit.jsonl` - append/update audit log

## Troubleshooting

### Python is not recognized

If double-clicking the batch file shows that `python` is not recognized:

- install Python 3.11 or newer
- during installation, enable the option to add Python to PATH
- run `setup_and_start.bat` again

### Required Python packages are missing

Run:

- `setup_and_start.bat`

That installs packages from `requirements.txt`.

### The app does not open in the browser

Open this address manually:

`http://127.0.0.1:8765`

If it still does not load, check whether the command window shows an error.

### Port 8765 is already in use

Close the other running copy of the app, then start again.

### A field looks wrong or blank

- review the row before append
- correct the value manually in the UI
- if a new PDF layout is consistently giving trouble, the extraction rules may need adjustment

## Privacy

This app is intended for local use only.

- PDFs are processed locally
- Excel files are stored locally
- no external API is required for extraction
- no cloud service is needed for normal use

The only time internet may be needed is during first-time Python package installation through `setup_and_start.bat`.

## Quick summary

If you only need the short version:

1. Extract the zip
2. Run `setup_and_start.bat` once
3. Later run `start_app.bat`
4. Create or upload a master Excel
5. Upload PDFs
6. Review rows
7. Append to master

