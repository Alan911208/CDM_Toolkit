# compareLib

**Category:** Data Comparison & Review

## Purpose / 目的
**English:** Compare two SAS data libraries cell by cell and export a color-coded Excel workbook via DDE. The comparison identifies new records, deleted (old-only) records, and changed cell values. Changed cells display old/new values separated by "/". Color coding makes differences immediately visible.

**中文:** 逐单元格比较两个 SAS 数据目录，通过 DDE 导出颜色编码的 Excel 工作簿。比对能够识别新增记录、删除（仅旧版）记录和变更的单元格值。变更的单元格以 "/" 分隔显示旧/新值。颜色编码使差异一目了然。

## Parameters / 参数

| Parameter | Description | Required | Default |
|-----------|-------------|----------|---------|
| `pCompare` | Path to OLD (baseline) data library / 旧版数据目录路径 | Yes | — |
| `pBase` | Path to NEW (current) data library / 新版数据目录路径 | Yes | — |
| `kVar` | Key variables for row matching (space-separated) / 行匹配键变量 | Yes | — |
| `eVar` | Variables to always flag as different / 始终标记为差异的变量 | No | (none) |
| `mode` | Comparison mode: 1=ignore spaces, 2=ignore case, 3=ignore both, 9=exact | No | `9` |

## Usage Examples / 使用示例

### Example 1: Standard library comparison
```sas
%compareLib(
    pCompare=D:\temp\tmp\old,
    pBase=D:\temp\tmp\new,
    kVar=studyid usubjid visit pageid line,
    mode=9);
```

### Example 2: Exclude certain variables from comparison
```sas
%compareLib(
    pCompare=D:\oldData,
    pBase=D:\newData,
    kVar=usubjid pageid,
    eVar=datetime comment,
    mode=3);
```

### Example 3: Case-insensitive comparison
```sas
%compareLib(
    pCompare=C:\baseline,
    pBase=C:\current,
    kVar=usubjid,
    mode=2);
```

## Output / 输出
- An Excel workbook with one sheet per dataset.
- A **TOC** (Table of Contents) sheet listing all datasets with location and change counts.
- Color coding legend:
  - **Green** (pattern 4): New records (only in new library)
  - **Grey** (pattern 15): Old records (only in old library)
  - **Yellow** (pattern 7): Individual differing data points
- Changed cells display as `OLD_VALUE/NEW_VALUE`.

## Requirements & Limitations / 系统要求与限制

### Requirements / 系统要求
- **Windows operating system** — REQUIRED
- **Microsoft Excel with DDE enabled** — REQUIRED
  - Excel 2010-2013: DDE is enabled by default
  - Excel 2016+: DDE is disabled by default. Enable via:
    `File > Options > Advanced > General > "Ignore other applications that use Dynamic Data Exchange (DDE)"` — UNCHECK
- **32-bit SAS + 32-bit Excel** is recommended for reliable DDE communication
- A logged-in Windows desktop session (DDE requires a visible Excel window)

### Limitations / 限制
- DDE may not work with 64-bit Office + 64-bit SAS
- Large datasets may cause DDE buffer overflow (results may be truncated)
- Not suitable for fully automated batch/server environments
- Excel must either be already running or capable of being launched by the macro
- The comparison compares values as character strings; numeric precision issues may cause false differences

## Internal Helpers / 内部辅助宏
- **%sas2xls** — Simplified DDE export of a SAS dataset to an Excel sheet

## Notes / 注意事项
- If a dataset exists only in the new library: all records are marked "New" (green).
- If a dataset exists only in the old library: all records are marked "Old" (grey).
- For datasets in both libraries: records are matched by key variables and compared cell-by-cell.
- Intermediate Excel sheets created during processing are deleted at the end.
- The TOC sheet includes a `_recNote` column showing the number of differing data points per dataset.
