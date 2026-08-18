# Lookup Studio

Lookup Studio is a local Streamlit application for combining related data from
multiple Excel workbooks. It is designed for the kind of task normally handled
with VLOOKUP or XLOOKUP, but it lets the user choose the files and columns
through a browser interface instead of writing formulas.

The first uploaded workbook is the base table. Every additional workbook is a
lookup source. The app keeps the selected rows from the base table, finds
matching records in each lookup source, adds the requested columns, and creates
a new Excel workbook for download.

## A practical example

Suppose you have these two files:

`product_catalog.xlsx`

| SKU | PRODUCT NAME |
| --- | --- |
| SKU-1001 | Desk Lamp |
| SKU-1002 | Office Chair |

`warehouse_stock.xlsx`

| SKU | UNITS IN STOCK |
| --- | ---: |
| SKU-1001 | 24 |
| SKU-1002 | 11 |

To combine them, use the product catalog as the base table and select:

- Base lookup column: `SKU`
- Base column to return: `PRODUCT NAME`
- Lookup column in the stock file: `SKU`
- Lookup column to return: `UNITS IN STOCK`

The resulting data will contain the SKU, product name, available quantity, and
a match-status column. The status makes it easy to see which products were
found in both files.

## Main features

- Upload between 2 and 10 Excel workbooks in one session.
- Generate the correct number of upload controls from the number selected by
  the user.
- Review each uploaded workbook in its own app tab.
- Select the worksheet and header row for every workbook.
- Select which base-table columns should remain in the output.
- Select the lookup key and return columns for every lookup source.
- Match text without being affected by capitalization or surrounding spaces.
- Match numeric Excel values to equivalent numeric text when requested.
- Detect missing cells, duplicate keys, duplicate rows, blank headers, and
  columns containing inconsistent value types.
- Choose whether duplicate lookup keys should use the first result or produce
  every matching result.
- Preview the merged data before downloading it.
- Choose the output filename.
- Download an `.xlsx` workbook with merged data, a lookup audit, and a data
  quality summary.

## Supported file types

The app accepts:

- `.xlsx`
- `.xlsm`
- `.xls`

Modern Excel files are opened with openpyxl in read-only mode and are consumed
one row at a time. This reduces the amount of temporary workbook data loaded by
the Excel reader. The selected worksheets are still converted to DataFrames
because the matching operation needs access to all selected keys.

Legacy `.xls` support is provided by `xlrd`.

Macro code from `.xlsm` files is not copied to the exported workbook. The
download is always a new `.xlsx` file containing values and reports.

## Requirements

- Python 3.10 or newer
- pip
- A current web browser

The Python dependencies are listed in [requirements.txt](requirements.txt).

## Installation

Clone or download the project, then open a terminal in the project directory.

Create a virtual environment:

```bash
python3 -m venv .venv
```

Activate it on macOS or Linux:

```bash
source .venv/bin/activate
```

On Windows PowerShell, activate it with:

```powershell
.venv\Scripts\Activate.ps1
```

Install the dependencies:

```bash
python3 -m pip install -r requirements.txt
```

## Starting the application

The provided shell script is the simplest way to start the app on macOS,
Linux, Git Bash, or Windows Subsystem for Linux:

```bash
./run_streamlit.sh
```

The script performs the following checks:

1. It changes to the project directory, so it works even when called from a
   different directory.
2. It uses `.venv/bin/python` when that virtual environment exists.
3. It checks that Python and Streamlit are available.
4. It starts `app.py` with Streamlit.

Arguments placed after the script name are passed to Streamlit. For example,
to use port 8502:

```bash
./run_streamlit.sh --server.port 8502
```

You can also start the app directly:

```bash
python3 -m streamlit run app.py
```

Streamlit will print a local address, normally `http://localhost:8501`. Open
that address in a browser if it does not open automatically.

## Using the application

### 1. Choose the number of files

Set **Number of files** to the total number of workbooks you want to combine.
The app creates one upload control for each file.

File 1 is always the base table. Files 2 through 10 are lookup sources.

### 2. Upload the base table

Upload the workbook whose rows should define the result. Every base row is
kept, even when no matching row exists in a lookup source.

In the base-file tab:

1. Select the worksheet containing the data.
2. Set the header row. Use `1` when column names are in the first row.
3. Select the base lookup column.
4. Select the other base columns you want to keep.

The base lookup column is always included in the output. It does not need to be
selected again under **Columns to return from the base table**.

### 3. Configure each lookup source

For every lookup workbook:

1. Select the worksheet and header row.
2. Select the column that corresponds to the base lookup column.
3. Select one or more columns to return.
4. Review or change the output column prefix.

Prefixes show where returned values came from and prevent column-name
collisions. For example, a prefix of `Warehouse` and a source column of
`UNITS IN STOCK` produce an output column named `Warehouse.UNITS IN STOCK`.

### 4. Review the data-quality preview

The app highlights potential problems before the merge:

- Red cells are missing values.
- Amber cells contain a value type that differs from the likely type of the
  rest of the column.
- A purple edge marks repeated values in the selected lookup-key column.

The quality section also reports:

- Row and column counts
- Missing-cell counts
- Mixed-type column counts
- Duplicate rows
- Blank or repeated headers that were renamed
- Trailing blank rows and unnamed blank columns that were removed

These findings are warnings. They do not automatically prevent the merge.

### 5. Choose the matching rules

The app offers three key-normalization options.

**Ignore surrounding spaces** removes spaces before and after a lookup value.
It does not remove spaces from the middle of a value.

**Ignore capitalization** treats values such as `SKU-1001` and `sku-1001` as
the same key.

**Match numbers to numeric text** allows an Excel number such as `42` to match
the text value `"42"`. Zero-padded identifiers remain distinct, so `0042` does
not automatically match `42`.

Blank keys never match another blank key.

### 6. Choose how duplicate lookup keys are handled

**Use the first matching row** behaves like a traditional VLOOKUP. If a lookup
source contains the same key more than once, only the first matching source row
is used. The output normally has the same number of rows as the base table.

**Return every matching row** keeps every source match. If one base row matches
three rows in a lookup source, the base row appears three times in the output.
With several lookup sources, repeated keys can multiply the number of result
rows.

Use the first matching row unless repeated keys represent separate records that
must all be preserved.

### 7. Create and download the workbook

Enter the desired output filename and select **Create merged workbook**. The
app shows an output preview and a lookup audit. Select **Download Excel
workbook** to save the result.

The `.xlsx` extension is added automatically. Characters that are not valid in
filenames are replaced safely.

## Understanding the output workbook

The downloaded workbook contains three worksheets.

### Merged Data

This sheet contains the selected base columns, returned lookup columns, and a
match-status column for every lookup source.

Possible status values are:

- `Matched`: the base key was found in the lookup source.
- `Not matched`: the base key had a value, but it was not found in the lookup
  source.
- `Missing key`: the base lookup-key cell was empty.

Missing output cells are highlighted. The header row is frozen and filters are
enabled.

### Lookup Audit

This sheet records how each lookup performed. It includes:

- Source filename and worksheet
- Selected lookup key
- Number of source rows
- Number of distinct keys
- Number of rows using duplicate keys
- Matched and unmatched base-row counts
- Missing base-key counts
- Row count after each lookup
- Number of returned columns

The audit is useful when the output contains fewer matches or more rows than
expected.

### Data Quality

This sheet summarizes the quality checks for every selected worksheet. It
includes missing cells, mixed-type columns, duplicate rows, removed blank
structure, header fixes, and a concise findings column.

## How multiple lookup files are processed

Every lookup source is matched against the key selected in the original base
table. Lookup file 2 does not use a value returned by lookup file 1 as its key.

For example, if three workbooks contain a product catalog, warehouse stock, and
supplier pricing, use the product catalog as the base and select the SKU as the
common key in all three files.

## Data handling

Uploaded file bytes and processed DataFrames are held by the running Streamlit
process. The application code does not save uploaded source files to the
project directory. Restarting the Streamlit process clears its in-memory
session and cache.

The upload limit is configured as 500 MB per file. Available memory and the
size of the merged result may impose a lower practical limit. Large workbooks
take longer to profile, merge, and export.

Excel worksheets support a maximum of 1,048,576 rows and 16,384 columns. The
app stops the export with a clear message if the merged result exceeds those
limits.

## Formula cells

The reader uses the most recently saved result of a formula cell. If formula
cells appear blank, open the source workbook in Excel or another spreadsheet
application, allow it to recalculate, save it, and upload it again.

The exported workbook contains values rather than the source formulas.

## Troubleshooting

### The launcher says permission denied

Make the script executable, then run it again:

```bash
chmod +x run_streamlit.sh
./run_streamlit.sh
```

### Streamlit is not installed

Activate the intended virtual environment and install the requirements:

```bash
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

### Port 8501 is already in use

Start the app on another port:

```bash
./run_streamlit.sh --server.port 8502
```

### The wrong row is being used as the header

Change **Header row** in that file's tab. This is common when a worksheet has a
title or notes above the table.

### Very few rows match

Check the following:

- The correct lookup columns are selected in both files.
- Capitalization matching is enabled when the files use different letter case.
- Surrounding-space matching is enabled when values may contain extra spaces.
- Numeric-text matching is enabled when one workbook stores keys as numbers.
- Identifiers have the same punctuation and zero padding.

This app performs exact normalized matching. It does not perform fuzzy matching
or guess that two different identifiers belong to the same record.

### The output contains more rows than the base table

This normally means **Return every matching row** was selected and at least one
lookup source contains duplicate keys. Use the audit sheet to identify the
source, or select **Use the first matching row**.

### A workbook cannot be opened

Confirm that the file is a valid, unencrypted Excel workbook. Password-protected
files are not supported. For an old `.xls` workbook, confirm that the
dependencies from `requirements.txt` were installed.

## Running the tests

Install the development requirements:

```bash
python3 -m pip install -r requirements-dev.txt
```

Run the test suite:

```bash
python3 -m pytest
```

The tests cover Excel ingestion, header cleanup, quality reporting, key
normalization, duplicate handling, base-column selection, workbook export, and
the dynamic Streamlit upload controls.

## Project structure

```text
app.py                         Streamlit interface
run_streamlit.sh               Application launcher
excel_merger/excel_io.py       Excel reading and header cleanup
excel_merger/quality.py        Data-quality profiling and preview styling
excel_merger/merge.py          Lookup-key normalization and merge logic
excel_merger/exporter.py       Excel download generation
tests/                         Automated checks
requirements.txt               Runtime dependencies
requirements-dev.txt           Test dependencies
.streamlit/config.toml         Streamlit upload and theme settings
```
