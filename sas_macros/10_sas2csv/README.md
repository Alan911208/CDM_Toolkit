# sas2csv — Batch SAS-to-CSV Conversion 批量SAS转CSV

## Category | 分类
Data Conversion & I/O | 数据转换与输入输出

## Purpose | 目的
Batch-convert every SAS dataset in a source directory to individual CSV (comma-separated values) files. Each dataset is exported using `PROC EXPORT` with the CSV DBMS engine, including variable labels as column headers.

批量将源目录中的每个SAS数据集转换为单独的CSV（逗号分隔值）文件。使用`PROC EXPORT`和CSV DBMS引擎导出每个数据集，变量标签作为列标题。

## Parameters | 参数

| Parameter | Description | Default | Required |
|-----------|-------------|---------|----------|
| `src` | Path to SAS data directory / SAS数据目录路径 | — | **Yes** |
| `tgt` | Output directory for CSV files / CSV文件输出目录 | Same as src | No |

## How It Works | 工作原理

1. Assigns a SAS library to the source directory.
2. Sets `nofmterr` and `fmtsearch` to avoid format warnings during export.
3. Uses `_getDatasetList` to enumerate all DATA-type datasets (excluding FORMATS catalog).
4. For each dataset, calls `PROC EXPORT` with `dbms=csv replace label` to export to a file named `<dataset>.csv` in the target directory.
5. Clears the library assignment.

## Usage Examples | 使用示例

### Example 1: Export to the same directory
```
%sas2csv(src=D:\Project\Derived);
```

### Example 2: Export to a separate CSV directory
```
%sas2csv(src=D:\Project\Derived, tgt=D:\Project\csv);
```

### Example 3: Export work datasets for inspection
```
%sas2csv(src=C:\Users\analyst\SAS\work, tgt=C:\Users\analyst\Desktop\csv_exports);
```

## Notes | 注意事项

- CSV files are named `<dataset_name>.csv`. Existing files with the same name will be overwritten (REPLACE option).
- The `LABEL` option in PROC EXPORT causes variable labels (if defined) to be used as column headers instead of variable names.
- If no target directory is specified, CSV files are written to the source directory alongside the `.sas7bdat` files.
- User-defined formats must be accessible in the source library's `formats` catalog, or `nofmterr` will suppress format errors.
- Supported by SAS 9.3+ (PROC EXPORT with DBMS=CSV).

## Source Reference | 源文件参考
DM_Toolbox.sas, lines 678-695
