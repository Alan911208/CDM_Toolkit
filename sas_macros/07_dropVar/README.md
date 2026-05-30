# dropVar — Batch-Drop Variables from Library 批量删除库中变量

## Category | 分类
Data Cleaning & Variable Management | 数据清洗与变量管理

## Purpose | 目的
Batch-drop the same variable(s) from every SAS dataset in a source library, writing the cleaned datasets to a target library. Useful for removing system-generated variables (e.g., `__STUDYOID`) or unwanted tracking columns across an entire project.

批量删除源库中每个SAS数据集中的相同变量，将清理后的数据集写入目标库。适用于在项目中批量移除系统生成的变量（如`__STUDYOID`）或不需要的跟踪列。

## Parameters | 参数

| Parameter | Description | Default | Required |
|-----------|-------------|---------|----------|
| `var` | Variable name(s) to drop, space-separated / 要删除的变量名，空格分隔 | — | **Yes** |
| `src` | Source library path / 源库路径 | — | **Yes** |
| `tgt` | Target library path / 目标库路径 | — | **Yes** |

## How It Works | 工作原理

1. Assigns SAS libraries to the source and target directories.
2. Uses `_getDatasetList` to enumerate all DATA-type datasets in the source library (excluding FORMATS catalog).
3. For each dataset, creates a copy in the target library with the specified variable(s) dropped using a DATA step `DROP` statement.
4. Clears the library assignments.

## Usage Examples | 使用示例

### Example 1: Drop a single system variable
```
%dropVar(var=__STUDYOID, src=D:\raw, tgt=D:\clean);
```

### Example 2: Drop multiple variables at once
```
%dropVar(var=VAR1 VAR2 VAR3, src=C:\data\derived, tgt=C:\data\final);
```

## Notes | 注意事项

- The source and target paths must be different directories, or the source will be overwritten.
- If a variable does NOT exist in a given dataset, SAS will generate a WARNING in the log but the DATA step will still execute successfully.
- The macro uses `_r1` and `_r2` as internal library references.
- All dataset names are preserved exactly in the output library.

## Source Reference | 源文件参考
DM_Toolbox.sas, lines 551-568
