# Macro 21: %versionDiff (+ %versionColor)

## 分类 Category
Data Comparison / 数据比较

## 用途 Purpose

### %versionDiff
Compare two SAS library versions dataset-by-dataset, detecting new, updated, and inactive records. Export differences and change-type frequency summaries to a timestamped Excel workbook (`Compare_YYYYMMDD.xlsx`).

批量比较两个 SAS 库版本（新旧），检出新增、更新和失效记录，并将差异与变更频数导出到带时间戳的 Excel 工作簿。

### %versionColor
Apply cell colour highlighting (DDE) to the already-open Excel workbook produced by `%versionDiff`. Uses red pattern (colour index 37) to mark changed cells.

对已打开的 `%versionDiff` 生成的 Excel 工作簿通过 DDE 施加红色单元格高亮。

## 参数表 Parameter Table

### %versionDiff

| Parameter | Required | Description |
|-----------|----------|-------------|
| newLib    | YES      | Pre-assigned SAS libref pointing to the NEW library |
| oldLib    | YES      | Pre-assigned SAS libref pointing to the OLD library |
| outLib    | YES      | Pre-assigned SAS libref pointing to the OUTPUT library |
| keyvar    | -        | Key variable(s) for merge matching (space-separated), e.g. `USUBJID SITEID VISIT` |
| exclVar   | -        | Variable(s) to exclude from comparison, comma-separated and quoted SAS names, e.g. `"_STATUE","RUNDATE"` |
| dropVar   | -        | Variable(s) to drop from all datasets, space-separated, e.g. `folderid SDVTier` |

### %versionColor

| Parameter | Required | Description |
|-----------|----------|-------------|
| outLib    | YES      | Same output libref used in %versionDiff |

## 使用示例 Usage Examples

### Example 1: Basic comparison
```sas
libname new "D:\Data\20240601";
libname old "D:\Data\20240501";
libname out "D:\Output";

%versionDiff(newLib=new, oldLib=old, outLib=out,
    keyvar=USUBJID SITEID VISIT PAGE FORM LINE);
```

### Example 2: With exclusion and drop variables
```sas
libname dNew "C:\Data\v2";
libname dOld "C:\Data\v1";
libname dOut "C:\Compare";

%versionDiff(newLib=dNew, oldLib=dOld, outLib=dOut,
    keyvar=USUBJID SITEID VISIT,
    exclVar="_STATUE","RUNDATE","USUBJID",
    dropVar=SDVTier folderid);
```

### Example 3: Add colour highlighting (run in same session, open Excel first)
```sas
* Open Compare_20240601.xlsx manually, then run:;
%versionColor(outLib=out);
```

## 注意 Notes

1. **Requires Excel with DDE support** — `%versionColor` uses DDE to communicate with a running Excel instance. The workbook must be open before running.
2. **Library assignment** — All three librefs (`newLib`, `oldLib`, `outLib`) must be pre-assigned with `libname` statements before calling the macro.
3. **_STATUE** — Records activated in a previous version receive `_STATUE="Inact"` when they are only in the old library.
4. **Frequency summary** — The `SourceSummary` sheet lists counts of New, Update, Inact, and Subject Total per dataset.
5. **Large libraries** — For libraries with many datasets, the loop may take significant time. Consider subsetting datasets first.
