# SAS Macro Toolkit — CDM Clinical Data Management

26 production-grade SAS macros organized by category. Each macro lives in its own folder with standalone code, bilingual documentation, test scripts with mock data, and expected test results.

## Quick Navigation

### Category 1: Data Inspection & Validation

| # | Macro | Folder | Purpose |
|---|-------|--------|---------|
| 01 | `%varExist` | [01_varExist/](01_varExist/) | Test if a variable exists in a dataset (OPEN/VARNUM) |
| 02 | `%dsRecCount` | [02_dsRecCount/](02_dsRecCount/) | Record/variable count summary for every dataset in a library |
| 03 | `%dateVarScan` | [03_dateVarScan/](03_dateVarScan/) | Scan all datasets for date/time variables (DTC/DAT/TIM) |

### Category 2: Data Manipulation & Cleaning

| # | Macro | Folder | Purpose |
|---|-------|--------|---------|
| 04 | `%dateCut` | [04_dateCut/](04_dateCut/) | Filter entire library by cutoff date |
| 05 | `%textSplit` | [05_textSplit/](05_textSplit/) | Split long text (>200 chars) into multiple columns |
| 06 | `%stripBlank` | [06_stripBlank/](06_stripBlank/) | Remove blank/duplicate placeholder rows from CDM datasets |
| 07 | `%dropVar` | [07_dropVar/](07_dropVar/) | Batch-drop variable(s) from all datasets in a library |
| 08 | `%dropVarSuffix` | [08_dropVarSuffix/](08_dropVarSuffix/) | Batch-drop variables by name suffix (e.g., `_U`) |
| 09 | `%mergeVar` | [09_mergeVar/](09_mergeVar/) | Merge variables from one form into all other datasets |

### Category 3: Data Conversion & I/O

| # | Macro | Folder | Purpose |
|---|-------|--------|---------|
| 10 | `%sas2csv` | [10_sas2csv/](10_sas2csv/) | Batch-export SAS datasets to CSV files |
| 11 | `%csv2sas` | [11_csv2sas/](11_csv2sas/) | Import CSV with intelligent variable naming |
| 12 | `%xlsxImport` | [12_xlsxImport/](12_xlsxImport/) | Import all Excel sheets as SAS datasets |
| 13 | `%sas2xlsx` | [13_sas2xlsx/](13_sas2xlsx/) | Batch-export SAS datasets to one Excel workbook |
| 14 | `%sas2xls` | [14_sas2xls/](14_sas2xls/) | Export to Excel sheet via DDE (legacy) |
| 15 | `%xls2sas` | [15_xls2sas/](15_xls2sas/) | Import Excel sheet via PROC IMPORT |

### Category 4: Date & Version Management

| # | Macro | Folder | Purpose |
|---|-------|--------|---------|
| 16 | `%stampRunDate` | [16_stampRunDate/](16_stampRunDate/) | Assign fixed run-date to all datasets in a library |
| 17 | `%runDateDiff` | [17_runDateDiff/](17_runDateDiff/) | Single-dataset old-vs-new comparison with date tracking |
| 18 | `%batchRunDateDiff` | [18_batchRunDateDiff/](18_batchRunDateDiff/) | Library-wide old-vs-new comparison |

### Category 5: Data Comparison & Review

| # | Macro | Folder | Purpose |
|---|-------|--------|---------|
| 19 | `%compareLib` | [19_compareLib/](19_compareLib/) | Cell-by-cell library comparison → color-coded Excel |
| 20 | `%reviewData` | [20_reviewData/](20_reviewData/) | Interactive Excel-based data review with query coloring |
| 21 | `%versionDiff` | [21_versionDiff/](21_versionDiff/) | Version diff: New/Update/Inactive/Stable classification |

### Category 6: Quality Control & Reporting

| # | Macro | Folder | Purpose |
|---|-------|--------|---------|
| 22 | `%qcCodeCheck` | [22_qcCodeCheck/](22_qcCodeCheck/) | MedDRA/WHODD coding QC (7 checks) |
| 23 | `%qcLabCheck` | [23_qcLabCheck/](23_qcLabCheck/) | Lab data QC (4 checks) |
| 24 | `%qcRecist` | [24_qcRecist/](24_qcRecist/) | RECIST 1.1 tumor response validation |
| 25 | `%dmQueryReport` | [25_dmQueryReport/](25_dmQueryReport/) | DMR Query Summary tables T5.1-T5.3 in RTF |
| 26 | `%validateSDS` | [26_validateSDS/](26_validateSDS/) | SDS validation (25 rule-based checks) |

## Folder Contents

Each folder contains:
- **`.sas`** — Standalone macro (includes required internal helpers)
- **`README.md`** — Bilingual documentation (Chinese + English) with parameter table and usage examples
- **`test_*.sas`** — Test script with mock SAS data and verification steps
- **`test_results.txt`** — Expected output and validation criteria

## Typical Workflows

```
Receipt:    %xlsxImport → %dsRecCount → %dateVarScan
Cleaning:   %stripBlank → %dropVar → %dropVarSuffix → %textSplit
Cut:        %dateCut → %mergeVar → %stampRunDate
Version:    %batchRunDateDiff → %compareLib → %reviewData
QC:         %validateSDS → %qcCodeCheck → %qcLabCheck → %dmQueryReport
Export:     %sas2csv / %sas2xlsx
```

## Prerequisites

- SAS 9.4 or later
- SAS/ACCESS to PC Files (for xlsx/xls import/export)
- Windows + Excel with DDE support (for macros 14, 19, 20, 21)

## Source

Rebuilt and optimized from the original 26 SAS macro files in `宏-DMI/宏-DMI/`.
Master toolkit: `DM_Toolbox.sas` in the project root.
