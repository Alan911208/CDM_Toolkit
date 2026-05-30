# %dsRecCount

## Category
Data Inspection & Validation

## Purpose / 目的

**English:** Scan every SAS dataset in a directory and produce a summary table (`work.jilu`) showing the dataset name, record count, and variable count in a single readable column. Uses SAS dictionary tables (`dictionary.tables`) for fast metadata access without opening each dataset individually.

**中文：** 扫描指定目录中的所有 SAS 数据集，生成汇总表（`work.jilu`），在单个可读列中显示数据集名称、记录数和变量数。通过 SAS 字典表（`dictionary.tables`）快速获取元数据，无需逐个打开数据集。

## Parameters / 参数

| Parameter | Required | Description |
|-----------|----------|-------------|
| `lib`     | Yes      | File system path to the directory containing SAS datasets |
| `label`   | No       | Label applied to the output variable `e` in `work.jilu` |

## Output / 输出

Dataset `work.jilu` with a single variable:
- `e` — Character (length 200), formatted as: `"DatasetName  Records:N, Variables:M"`

## Usage Examples / 使用示例

### Example 1: Basic usage — scan a directory
```sas
%dsRecCount(lib=D:\project\rawdata);
proc print data=work.jilu noobs; run;
/*
Sample output:
--------------------------------------------
DM  Records:120, Variables:25
AE  Records:450, Variables:18
LB  Records:3200, Variables:12
--------------------------------------------
*/
```

### Example 2: With a label for the output column
```sas
%dsRecCount(lib=D:\project\final, label=Final SDTM Datasets);
proc print data=work.jilu label noobs; run;
```

### Example 3: Use output for quality check
```sas
%dsRecCount(lib=D:\project\rawdata);
data _null_;
    set work.jilu;
    if index(e, 'Records:0') then
        put "WARNING: Empty dataset detected — " e;
run;
```

## Notes / 注意事项

- The macro creates a temporary library reference `_rc` pointing to `&lib`, reads metadata, and clears the reference afterward.
- Uses `by memname; if last.memname;` to keep only the last row per dataset (in case `dictionary.tables` returns multiple rows per dataset).
- The internal working table `_rc_alldata` is dropped after use.
- If the directory contains no SAS datasets, `work.jilu` will be empty (0 observations, 1 variable).
- This macro does NOT depend on any internal helper macros.

## Source Reference / 来源

`DM_Toolbox.sas` — Tool 2, Category 1: Data Inspection & Validation
