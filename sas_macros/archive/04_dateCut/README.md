# %dateCut

## Category
Data Manipulation & Cleaning

## Purpose / 目的

**English:** Filter a library of clinical datasets by a cutoff date. Each dataset is processed according to its role: datasets with visit-date variables are filtered (only rows with visits before the cutoff are kept); datasets specified in `otherVar` are filtered by a direct date comparison; datasets in `keepAll` are copied without filtering. A record-count comparison report (`_rpt_`) is produced in the target library showing original vs. filtered counts.

**中文：** 按截止日期筛选临床数据集库。每个数据集根据其角色处理：带访视日期变量的数据集被过滤（仅保留截止日期前的访视行）；`otherVar` 中指定的数据集按直接日期比较过滤；`keepAll` 中的数据集完整复制。目标库中生成记录数对比报告（`_rpt_`），显示原始与过滤后的记录数。

## Parameters / 参数

| Parameter  | Required | Description |
|------------|----------|-------------|
| `src`      | Yes      | Source data library path |
| `tgt`      | Yes      | Target (output) data library path |
| `cutoff`   | Yes      | Cutoff date in YYYYMMDD format (e.g., `20170228`) |
| `visitVar` | No       | Visit-date variables: `"DS.var DS.var ..."` (DS=dataset, var=date variable) |
| `otherVar` | No       | Other date-filtered datasets: `"DS.var DS.var ..."` (DS=dataset, var=date variable) |
| `keepAll`  | No       | Datasets to copy without filtering (space-separated names, default: `FORMATS CO`) |
| `byVar`    | No       | Key variables for merging (default: `Usubjid Visit`) |

## Output / 输出

- Filtered datasets in `&tgt` library
- `&tgt._rpt_` — comparison report with columns: `memname`, `nlobs0` (original), `nlobs1` (new), `pct` (percentage)

## Usage Examples / 使用示例

### Example 1: Basic cutoff with visit-date filtering
```sas
%dateCut(
    src      = D:\project\rawdata,
    tgt      = D:\project\interim,
    cutoff   = 20240131,
    visitVar = SV.SVSTDTC_YMD,
    keepAll  = FORMATS
);

proc print data=_oData._rpt_ label; run;
```

### Example 2: Multiple visit sources and other-date filtering
```sas
%dateCut(
    src      = D:\project\sdtm,
    tgt      = D:\project\cutoff,
    cutoff   = 20231231,
    visitVar = SV1.SVSTDTC_YMD SV2.SVSTDTC_YMD,
    otherVar = AE.AESTDTC_YMD DS.DSSTDTC_YMD,
    keepAll  = FORMATS CO DM
);
```

### Example 3: Custom by-variables
```sas
%dateCut(
    src      = D:\raw,
    tgt      = D:\cut,
    cutoff   = 20230630,
    visitVar = VS.VSDTC,
    byVar    = SUBJID
);
```

## Notes / 注意事项

- `visitVar` datasets are filtered by merging with a master visit list: only rows matching subjects and visits present before the cutoff are retained.
- `otherVar` datasets use a simple `WHERE variable < cutoff` row filter (no merge).
- `keepAll` datasets are copied byte-for-byte with no row filtering.
- The macro uses internal librefs `_iData` (source) and `_oData` (target) — avoid conflicts with these names.
- Date comparison assumes numeric SAS dates (days since 1960-01-01). The raw cutoff string is converted via `input("&cutoff", yymmdd8.)`.
- The internal helper `_dropIfExists` is included in this standalone .sas file.

## Source Reference / 来源

`DM_Toolbox.sas` — Tool 4, Category 2: Data Manipulation & Cleaning
