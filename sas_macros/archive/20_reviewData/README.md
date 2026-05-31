# reviewData

**Category:** Data Comparison & Review

## Purpose / 目的
**English:** Interactive Excel-based data review. Exports all datasets from a SAS library to individual Excel sheets and applies query/crosscheck-based color coding for visual data inspection. Supports finding codes from both the QUERY and CROSSCHK datasets to highlight cells needing attention.

**中文:** 交互式 Excel 数据审查。将 SAS 数据目录中的所有数据集导出到独立的 Excel 工作表，并根据查询/交叉检查的结果进行颜色编码，便于可视化数据检查。支持来自 QUERY 和 CROSSCHK 数据集的发现代码，高亮需要关注的单元格。

## Parameters / 参数

| Parameter | Description | Required | Default |
|-----------|-------------|----------|---------|
| `path` | Path to the data library / 数据目录路径 | Yes | `D:\Temp\tmp` |
| `style` | Review style: `FULL` (all cols) or `COMPACT` (exclude STUDYID, DOMAIN, SUBJID, SITEID, RANDOM, VISITNUM) | No | `FULL` |
| `showFail` | Show unmatched findings: `YES` or `NO` / 显示未匹配的发现 | No | `NO` |

## Usage Examples / 使用示例

### Example 1: Full review with default settings
```sas
%reviewData(path=P:\Study\RawData);
```

### Example 2: Compact review showing only key columns
```sas
%reviewData(path=P:\Study\RawData, style=COMPACT);
```

### Example 3: Full review with unmatched findings diagnostic
```sas
%reviewData(path=P:\Study\RawData, style=FULL, showFail=YES);
```

## Output / 输出
- An Excel workbook with one sheet per dataset in the library.
- A **TOC** (Table of Contents) sheet listing all datasets with descriptions.
- Color coding from QUERY/CROSSCHK findings (see table below).
- A legend on the TOC sheet explaining the color codes (only if QUERY dataset exists).

### Finding Color Codes / 发现颜色编码

| Code | Meaning | Excel Pattern | Appearance |
|------|---------|---------------|------------|
| 001 | Data Confirmed / 数据已确认 | pattern 07 | Green |
| 002 | Supported by AE/SAE record / AE/SAE支持 | pattern 19 | Purple |
| 003 | Supported by lab finding / 实验室支持 | pattern 24 | Orange |
| 004 | Supported by medical history / 病史支持 | pattern 34 | Blue |
| 005 | Supported by operation history / 手术史支持 | pattern 36 | Teal |
| 006 | Supported by physical exam / 体检支持 | pattern 38 | Brown |
| 007 | Supported by treatment/drugs / 治疗/药物支持 | pattern 39 | Rose |
| 008 | Supported by vital sign / 生命体征支持 | pattern 40 | Dark green |
| 999 | Unknown Type / 未知类型 | pattern 03 | Red |
| 021 | Entry Comment / 录入注释 | border 5 | Thick border |
| 022 | Answered programme query / 已回答程序核查 | border 2 | Left border |
| 023 | Answered key-in query / 已回答录入核查 | border 6 | Right border |
| 024 | Open programme query / 未答程序核查 | border 2 dashed | Left dashed |
| 025 | Open key-in query / 未答录入核查 | border 6 dashed | Right dashed |

## Supported Library Datasets / 支持的库数据集
All datasets are included **except** those in the exclusion list:
`WVISIT`, `LABREF`, `CMEDDRA`, `CUSERD`, `MISPAGE`, `CWHODD`, `PAGESTATUS`, `CROSSCHK`, `TRANSLATION`, `QUERY`, `FORMATS`, `PAT`

Optional special datasets:
- **QUERY**: If present, findings are extracted and color-coded onto corresponding data sheets
- **CROSSCHK**: If present, additional crosscheck findings are extracted
- **FORMATS**: If present, used as a user-defined format catalog (`fmtsearch` option)

## Requirements & Limitations / 系统要求与限制

### Requirements / 系统要求
- **Windows operating system** — REQUIRED
- **Microsoft Excel with DDE enabled** — REQUIRED
  - Enable: `File > Options > Advanced > General > "Ignore other applications that use Dynamic Data Exchange (DDE)"` — UNCHECK
- **32-bit SAS + 32-bit Excel** is recommended
- A logged-in Windows desktop session (DDE requires a visible Excel window)

### Limitations / 限制
- DDE may not work with 64-bit Office + 64-bit SAS
- Large datasets may cause DDE buffer overflow
- Not suitable for fully automated batch/server environments
- Color coding relies on the QUERY and CROSSCHK datasets having a specific structure

## Internal Helpers / 内部辅助宏
- **%sas2xls** — Simplified DDE export of a SAS dataset to an Excel sheet

## Notes / 注意事项
- The macro attempts to launch Excel if it is not already running (10-second timeout).
- Three blank worksheets are inserted at the start for macro storage (later hidden).
- If `showFail=YES` and unmatched findings exist, a dataset `__DataSetName` is created in WORK with the unmatched items, and a WARNING is printed to the SAS log.
- The `style=COMPACT` option drops variables STUDYID, DOMAIN, SUBJID, SITEID, RANDOM, VISITNUM to reduce visual clutter.
