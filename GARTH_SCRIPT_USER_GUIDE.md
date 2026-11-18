# Garth Script — Easy User Guide

## What this program does

`garth_script.py` reads data files and uploads the data into Snowflake tables.

It can read:

- `.xlsx` and `.xls` Excel files
- `.txt` files that use tabs between columns
- A local file or local folder
- Files stored in Google Drive under the `UW` or `EIS` branch

Each uploaded row also receives a `SOURCE_FILE_NAME` value. This is the original file name without the file extension.

For example:

```text
Active Portfolio - 2023-12.xlsx
```
becomes:

```text
Active Portfolio - 2023-12
```

This helps identify which file supplied each row and allows the script to refresh old rows from that file.

---

## Main commands

### Run the UW files from Google Drive

```powershell
python .\garth_script.py --drive-branch UW
```

This is the normal UW command.

### Run the EIS files from Google Drive

```powershell
python .\garth_script.py --drive-branch EIS
```

### Upload one local Excel or TXT file

```powershell
python .\garth_script.py "C:\Data\Active Portfolio - 2023-12.xlsx"
```

### Upload every supported file directly inside a local folder

```powershell
python .\garth_script.py "C:\Data\Paid Claims"
```

The folder name becomes the Snowflake table name. For example, folder `Paid Claims` becomes table `Paid_Claims`.

---

## Simple flow

```text
1. Start the script
       ↓
2. Select the local path, or UW/EIS Drive branch
       ↓
3. Find supported files
       ↓
4. Read the Excel or TXT file
       ↓
5. Find the column headings
       ↓
6. Add SOURCE_FILE_NAME to every row
       ↓
7. Prepare data types and column names
       ↓
8. Connect to Snowflake
       ↓
9. Replace data previously loaded from the same source file
       ↓
10. Insert the new data
       ↓
11. Successfully loaded Drive files move to the archive folder
       ↓
12. Write the result to the pipeline log
       ↓
13. Show the final success or failure summary
```

---

## 1. Start the program

The script reads the command entered by the user.

For example:

```powershell
python .\garth_script.py --drive-branch UW
```

means:

- Run `garth_script.py`
- Read files from the Google Drive UW branch

The script checks that the command contains either:

- A local file or folder; or
- A Google Drive branch such as `UW` or `EIS`

It does not allow both at the same time.

---

## 2. Local file names and folder names

### One local file

The file name becomes the table name.

Example:

```text
Reported Claims - 2026-09.xlsx
```

becomes:

```text
Reported_Claims_2026_09
```

The file extension is removed. Spaces and punctuation are changed to `_` where required.

### One local folder

The folder name becomes the table name.

Example:

```text
C:\Data\Paid Claims
```

becomes table:

```text
Paid_Claims
```

The script processes supported files directly inside that folder and combines them into the same table.

It does not read nested subfolders for the local-folder mode.

---

## 3. Google Drive file and table names

### Files directly inside a UW folder

The file name becomes the table name.

Example:

```text
Revenue.xlsx
```

becomes table:

```text
Revenue
```

### Files inside a UW child folder

The child folder name becomes the table name, and files inside that folder are uploaded to the same table.

Example:

```text
UW / Active Portfolio / Active Portfolio - 2023-12.xlsx
UW / Active Portfolio / Active Portfolio - 2024-12.xlsx
```

becomes table:

```text
Active_Portfolio
```

---

## 4. TXT conversion

A TXT file must normally be tab-delimited. This means that columns are separated by tabs, not commas.

The script reads the TXT data and performs these checks:

- Detects the file encoding
- Reads the first row as column names
- Removes duplicate header names
- Keeps the first copy of a duplicated column
- Converts the data for use by Snowflake

For local TXT files, the script may also create an `.xlsx` copy beside the TXT file.

Example:

```text
Transaction Dump_2020.txt
```

creates a local workbook during conversion, then the data is uploaded using the loader.

To replace an already converted Excel file in local mode, use:

```powershell
python .\garth_script.py "C:\Data\Transaction Dump_2020.txt" --overwrite
```

Drive TXT files are read directly and do not need a separate Excel copy before upload.

---

## 5. Excel header detection

Excel files sometimes contain:

- A report title
- Blank rows
- Instructions or notes
- PowerBI-style preamble rows

The script searches the first 50 rows and finds the first row that looks like column headings.

A heading row normally contains several text values, such as:

```text
Policy No | Insured Name | Current End Date | Gross Premium
```

When found, the script reports it, for example:

```text
INFO Active Portfolio - 2024-12.xlsx: using worksheet 'Sheet1', Excel row 4 as column headers
```

This means the actual column headings start on Excel row 4.

### Reported Claims special rule

For `Reported Claims`, the script uses worksheet `Data` rather than using the first worksheet automatically.

---

## 6. Column preparation

Before upload, the script prepares the file.

It can:

- Remove empty or duplicate column names
- Add missing columns
- Keep extra columns found in a newer file
- Convert blank values to NULL or empty values where required
- Add `SOURCE_FILE_NAME`
- Match the incoming columns to the Snowflake table

If one file has an extra column, the column is added and that file supplies its data. Other files without that column receive empty/NULL values for it.

---

## 7. Data type handling

Some values may be written differently in different years, for example a number in one file and text in another.

The script checks selected columns against the target Snowflake table and prepares values for that database type.

Examples:

| Source value | Required handling |
|---|---|
| `P0000700550` | Uploaded as text, not as a number |
| `M1` | Uploaded as text |
| `Drivers Date of Birth` | Uploaded as text when the source contains invalid date values |
| Current End Date | Prepared as a date when valid; invalid values are handled as NULL |
| Gross Premium | Prepared as a decimal value rather than losing precision |

Special text columns are kept as strings when identifiers or code values may lose leading zeroes.

---

## 8. Existing-file replacement

The normal append setting checks whether the file name already exists in the `SOURCE_FILE_NAME` column.

Example:

```text
SOURCE_FILE_NAME = Active Portfolio - 2023-12
```

When the same file is uploaded again, the script:

1. Starts the new upload as a replacement operation.
2. Finds the old rows with that `SOURCE_FILE_NAME`.
3. Deletes only those old rows.
4. Inserts the rows from the new file.
5. Keeps rows belonging to every other source file.

It does not delete the complete table.

File extensions are removed from the source name. Therefore this TXT file:

```text
Active Portfolio - 2023-12.txt
```

matches:

```text
Active Portfolio - 2023-12
```
in Snowflake.

---

## 9. Successful Drive files move to archive

Google Drive files are archived only after Snowflake loading succeeds.

The script:

1. Reads a file.
2. Uploads its rows.
3. Confirms the expected rows were written.
4. Moves the Drive file to the `archive` folder.
5. Prints `ARCHIVED`.

If loading fails, the source file remains available in the original location so it can be corrected and retried.

---

## 10. Pipeline logging

The script records processing steps in:

```text
DW_AGAM_DEV.UW_RAW_STG.ETL_PIPELINE_LOG
```

The two important settings are:

```text
SNOWFLAKE_DATABASE=DW_AGAM_DEV
SNOWFLAKE_AUDIT_SCHEMA=UW_RAW_STG
```

The script records these steps:

| Step | Meaning |
|---|---|
| `ETL_PIPELINE` | The complete script run |
| `CONNECT_SNOWFLAKE` | Snowflake connection |
| `LOAD_MAPPING` | Upload or mapping of one file/table |
| `CLOSE_SNOWFLAKE_CURSOR` | Snowflake cursor cleanup |
| `CLOSE_4D_CONNECTION` | 4D connection cleanup, marked skipped by this script |

The status can be:

| Status | Meaning |
|---|---|
| `SUCCESS` | Completed successfully |
| `FAILED` | Something failed |
| `SKIPPED` | This step is not used |
| `INTERRUPTED` | The user stopped the run, such as with Ctrl+C |

The log stores the pipeline run ID, step number, table, message, error details, start time, and finish time.

---

## 11. Final summary

At the end of a run, the script prints a summary similar to:

```text
Finished: 3 table group(s), 421578 row(s) loaded, 0 failed file(s).
```

This tells the user:

- How many table groups were processed
- How many rows were loaded
- How many source files failed

If a file failed, the previous lines show the file name and the reason.

---

## Common commands

### Normal UW run

```powershell
python .\garth_script.py --drive-branch UW
```

The `--if-exists append` option is optional because `append` is already the default.

### Test file validation without uploading

```powershell
python .\garth_script.py --drive-branch UW --dry-run
```

Dry-run mode reads and checks files, but does not:

- Upload rows
- Replace existing rows
- Move Drive files to archive

### Upload using a larger batch size

```powershell
python .\garth_script.py --drive-branch UW --chunk-size 20000
```

Chunk size controls how many rows are sent in each Snowflake write operation. It does not change the total row count.

### Stop if an existing destination table is found

```powershell
python .\garth_script.py "C:\Data\Revenue.xlsx" --if-exists fail
```

### Replace complete-table behavior

```powershell
python .\garth_script.py "C:\Data\Revenue.xlsx" --if-exists replace
```

Warning: this option has complete-table replacement behavior and should not be confused with the normal source-file refresh behavior used by `append`.

---

## `.env` settings

The script uses these important settings:

```text
SNOWFLAKE_USER
SNOWFLAKE_ACCOUNT
SNOWFLAKE_WAREHOUSE
SNOWFLAKE_ROLE
SNOWFLAKE_DATABASE
SNOWFLAKE_GARTH_SCHEMA
SNOWFLAKE_AUDIT_SCHEMA
GOOGLE_DRIVE_ROOT_FOLDER_ID
GOOGLE_SERVICE_ACCOUNT_FILE
```

Example:

```text
SNOWFLAKE_DATABASE=DW_AGAM_DEV
SNOWFLAKE_GARTH_SCHEMA=GARTH_CD_DEV
SNOWFLAKE_AUDIT_SCHEMA=UW_RAW_STG
GOOGLE_DRIVE_ROOT_FOLDER_ID=<shared-drive-folder-id>
GOOGLE_SERVICE_ACCOUNT_FILE=C:\ETL\service-account.json
```

The service account must have access to the configured Google Drive root folder.

---

## Export `CREATED_ON` summaries for a Snowflake schema

`created_on_schema_summary.py` runs the null count, total row count, and maximum
`CREATED_ON` value for each base table in the selected schema. Tables without a
`CREATED_ON` column are skipped. The results are written to a CSV with one row
for each metric and table.

```powershell
python .\created_on_schema_summary.py --database ODS_DEV --schema UWTR_ALL_DBO
```

By default, the output is `created_on_summary.csv` in the current folder. Use
`--output` to choose another path:

```powershell
python .\created_on_schema_summary.py --database ODS_DEV --schema UWTR_ALL_DBO --output .\reports\created_on_summary.csv
```

The script uses the Snowflake connection settings in `.env` (or the process
environment). The database and schema can also be taken from
`SNOWFLAKE_DATABASE` and `SNOWFLAKE_SCHEMA` when the command-line options are
omitted.

---

## Where the script stores or changes files

### Local mode

A TXT file may create a matching `.xlsx` file beside it, for example:

```text
Transaction Dump_2020.txt
Transaction Dump_2020.xlsx
```

Use `--overwrite` if this workbook already exists and should be regenerated.

### Drive mode

Successfully loaded files are moved into the Drive folder named `archive`.

Failed files stay in their original location for retry.

---

## If the script stops with an error

### The file or table name looks incorrect

Check the source file or folder name. Spaces and punctuation are converted into a Snowflake-compatible table name where needed.

### Snowflake says a column cannot change type

The script expects the existing target column and incoming data to be compatible. Columns such as policy numbers and account codes should remain text.

### The file is not loading into the correct worksheet

Excel headers are detected automatically. Check that the workbook contains:

- At least three text column names together in the header row
- No different text string inside the actual header row
- For Reported Claims, a worksheet named `Data`

### The Drive file remains and is not archived

This normally means its Snowflake load failed. Read the exact message printed after:

```text
FAILED UW table=... file=...
```

Only successfully loaded Drive files are archived.

### The script asks for missing environment settings

Check the `.env` file and confirm that all required settings exist.

---

## Quick checklist before running

1. Open PowerShell in the script folder.
2. Confirm Python and required packages are installed.
3. Confirm `.env` contains the Snowflake and Google Drive settings.
4. Confirm Drive files are shared with the service account.
5. Use the normal UW command:

```powershell
python .\garth_script.py --drive-branch UW
```

6. Review the `FAILED` or `LOADED` lines.
7. Check the Snowflake table `ETL_PIPELINE_LOG` for the detailed pipeline status.
