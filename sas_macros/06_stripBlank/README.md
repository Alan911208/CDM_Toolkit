# stripBlank — Remove Blank/Duplicate Placeholder Rows 移除空白/重复占位行

## Category | 分类
Data Cleaning & CDM Processing | 数据清洗与CDM处理

## Purpose | 目的
Remove blank and duplicate placeholder rows from CDM (Clinical Data Management) page-based datasets. Automatically handles pStatus column removal, seqCD/seqTEXT cleanup, and auto-filling of missing index variable values.

从CDM（临床数据管理）分页数据集中移除空白和重复的占位行。自动处理pStatus列的移除、seqCD/seqTEXT的清理，以及自动填充缺失的索引变量值。

## Parameters | 参数

| Parameter | Description | Default | Required |
|-----------|-------------|---------|----------|
| `inlib`   | Input library name / 输入逻辑库名称 | work | No |
| `outlib`  | Output library name / 输出逻辑库名称 | Same as inlib | No |
| `insets`  | Input dataset name(s), space-separated / 输入数据集名称，空格分隔 | — | **Yes** |
| `outsets` | Output dataset name(s), space-separated / 输出数据集名称，空格分隔 | Same as insets | No |
| `idxVar`  | Index variable name / 索引变量名称 | SEQNUM | No |
| `keepPS`  | Keep the pStatus column? / 是否保留pStatus列 | NO | No |

## How It Works | 工作原理

1. Collects column metadata for all datasets in the input library.
2. For each dataset, checks whether the index variable (`idxVar`), `seqCD`, and `seqTEXT` exist, and counts non-missing values.
3. In the output DATA step:
   - Filters out rows where `line` is missing (blank placeholder rows).
   - If the index variable has zero non-missing values and is numeric-formatted: deletes rows with blank Status and non-first-line rows, plus rows where pStatus='Added'. Drops the empty index variable.
   - If the index variable has values but is character-formatted: fills missing values using the `line` variable.
   - Drops empty `seqCD` and `seqTEXT` columns.
   - Drops `pStatus` unless `keepPS=YES`.

## Usage Examples | 使用示例

### Example 1: Basic cleaning
```
%stripBlank(inlib=raw, outlib=clean, insets=lb1 lb2, outsets=lb1_clean lb2_clean);
```

### Example 2: Specify a custom index variable
```
%stripBlank(inlib=work, outlib=derived, insets=ae1 vs1, idxVar=LBTESTCD);
```

### Example 3: Keep pStatus column
```
%stripBlank(inlib=raw, outlib=clean, insets=dm, keepPS=YES);
```

## Notes | 注意事项

- The input dataset is expected to contain a `line` variable used to identify blank placeholder rows.
- If multiple datasets are processed, `insets` and `outsets` should have a one-to-one correspondence by position.
- The macro uses `dictionary.columns` for metadata inspection and `dictionary.tables` via the `_getDatasetList` helper.
- Temporary working tables `_tmp_` and `pat` are automatically dropped at the end of execution.

## Source Reference | 源文件参考
DM_Toolbox.sas, lines 452-536
