# mergeVar — Merge Variables Across Library 跨数据集合并变量

## Category | 分类
Data Integration & Merging | 数据整合与合并

## Purpose | 目的
Merge one or more variables from a designated "source form" dataset into every other dataset in a library, using specified key variables for the join. This is commonly used to propagate subject-level information (e.g., informed consent date, randomization date) from the DM domain across all other SDTM/ADaM domains.

使用指定的关键变量，将指定"源表单"数据集中的一个或多个变量合并到库中所有其他数据集。通常用于将受试者级别信息（如知情同意日期、随机化日期）从DM域传播到所有其他SDTM/ADaM域。

## Parameters | 参数

| Parameter | Description | Default | Required |
|-----------|-------------|---------|----------|
| `form` | Source dataset name (e.g., dm) / 源数据集名称 | — | **Yes** |
| `vars` | Variable(s) to merge, space-separated / 要合并的变量，空格分隔 | — | **Yes** |
| `key`  | Key variable(s) for merge join, space-separated / 合并连接的关键变量，空格分隔 | — | **Yes** |
| `src`  | Source library path / 源库路径 | — | **Yes** |
| `tgt`  | Target library path / 目标库路径 | — | **Yes** |

## How It Works | 工作原理

1. Assigns SAS libraries to the source and target directories.
2. Extracts the key variable(s) and requested variable(s) from the source form dataset into a temporary dataset `_a`, then sorts it by key.
3. Iterates through all other datasets in the source library.
4. For each dataset, sorts the source data by key (into `__src`), then merges with `_a` by key, writing the result to the target library.
5. Cleans up temporary datasets and clears library assignments.

## Usage Examples | 使用示例

### Example 1: Merge informed consent date from DM
```
%mergeVar(form=dm, vars=rficdtc, key=usubjid, src=D:\raw, tgt=D:\enriched);
```

### Example 2: Merge multiple date variables
```
%mergeVar(form=dm, vars=rficdtc rfendtc, key=usubjid, src=C:\sdtm, tgt=C:\enriched);
```

### Example 3: Merge using a composite key
```
%mergeVar(form=suppae, vars=qnam qval, key=usubjid idvar, src=D:\data, tgt=D:\merged);
```

## Notes | 注意事项

- The source form dataset is also merged with itself (same key values), which is harmless but produces a duplicate of the source form in the output.
- All datasets in the source library must contain the key variable(s) for the merge to succeed.
- The temporary datasets `work._a` and `work.__src` are automatically dropped.
- Source and target directories should differ unless intentional overwrite is desired.

## Source Reference | 源文件参考
DM_Toolbox.sas, lines 631-659
