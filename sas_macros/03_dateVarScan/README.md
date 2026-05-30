# %dateVarScan

## Category
Data Inspection & Validation

## Purpose / 目的

**English:** Scan all datasets in a SAS library and identify every character variable whose name contains `DTC`, `DAT`, or `TIM` — the standard CDISC suffixes for date/time variables. Non-missing values are collected and written to `libref.datelist`, providing a complete inventory of date/time fields across the study data.

**中文：** 扫描 SAS 库中所有数据集，识别名称中包含 `DTC`、`DAT` 或 `TIM` 的字符型变量（CDISC 标准的日期/时间变量后缀）。非缺失值将被收集并写入 `libref.datelist`，生成整个研究数据的日期/时间字段清单。

## Parameters / 参数

| Parameter | Required | Description |
|-----------|----------|-------------|
| `libref`  | Yes      | Library reference (libref) pointing to the SAS data library |

## Output / 输出

Dataset `libref.datelist` with the following columns:
- `studyid` through `line` — all original dataset variables (inherited via `studyid--line`)
- `vName`  — Character ($32): variable name
- `vLabel` — Character ($200): variable label
- `value`  — Character ($200): non-missing value of the date/time variable

## Usage Examples / 使用示例

### Example 1: Basic scan of a library
```sas
libname sdtm "D:\project\sdtm";
%dateVarScan(libref=sdtm);

/* Browse the date/time variable inventory */
proc print data=sdtm.datelist; run;
proc freq data=sdtm.datelist; tables vName; run;
```

### Example 2: Check which datasets have date variables
```sas
%dateVarScan(libref=adam);
proc sql;
    select distinct memname, vName, vLabel
    from adam.datelist
    order by memname, vName;
quit;
```

### Example 3: Cross-reference with expected variable list
```sas
%dateVarScan(libref=raw);
/* Compare with expected DTC variables */
proc sql;
    select vName, count(distinct memname) as n_datasets
    from raw.datelist
    group by vName;
quit;
```

## Notes / 注意事项

- Variable filtering pattern: any character variable whose uppercased name contains `DTC`, `DAT`, or `TIM`, AND whose first character is not an underscore.
- The `studyid--line` double-dash range in the KEEP statement preserves common study-level variables in the output; if those variables do not exist in the input dataset, SAS will issue a harmless note.
- Only non-missing values are retained in the output.
- The macro skips any dataset named `FORMATS`.
- The internal helper `_dropIfExists` is included in this standalone .sas file.

## Source Reference / 来源

`DM_Toolbox.sas` — Tool 3, Category 1: Data Inspection & Validation
