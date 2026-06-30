# Regal Extract

Regal Extract is a private, local web app for ICEGATE shipping-bill PDFs. It uses fixed layout and text rules - no AI API, subscription, cloud upload, or paid service.

## Start

1. Install Python 3.11 or newer.
2. Double-click `setup_and_start.bat` the first time. It installs the open-source packages in `requirements.txt` and opens the app.
3. Later, double-click `start_app.bat`.
4. In the web page, choose an existing master workbook or create a new one, upload PDFs, review the extracted row, and append.

The app runs at `http://127.0.0.1:8765`. Close the command window or press `Ctrl+C` to stop it.

## Workbook safety rules

- Columns are located by header name, not a hard-coded column letter.
- Existing invoice numbers are updated only in recognized extraction columns.
- New invoice numbers are appended.
- User-created columns are not cleared or overwritten.
- Formulas in user-created columns are carried into genuinely new rows.
- The `$ / E` cell remains numeric but receives the correct `$`, `€`, `£`, or `¥` number format.
- INR is a live Excel formula: foreign amount multiplied by exchange rate.
- Saves are atomic; when a new existing master is uploaded, the previous stored master is backed up.

## Recognized headers

`Particulars`, `Bill No.`, `Date`, `SB NO.`, `SB DATE`, `Port Code`, `$ / E`, `Exch Rate`, `INR`, `Value@5%`, `IGST@5%`, `FOB`, `DBK`, `RoDTEP`.

Reasonable aliases such as `Invoice No.`, `Shipping Bill No.`, `Exchange Rate`, and `Drawback` are also supported.

## Files and privacy

The active master is stored in `data/master.xlsx`. Backups and a small append audit log are stored under `data/`. Nothing is sent over the internet.
