# dropVarSuffix — Batch-Drop Variables with Suffix 批量删除带后缀的变量

## Category | 分类
Data Cleaning & Variable Management | 数据清洗与变量管理

## Purpose | 目的
Batch-drop variables whose names end with a given suffix from every SAS dataset in a source library. The default suffix is `_U`, which is commonly used for unit variables in SDTM/ADaM datasets. All matching variables are identified across the entire library, then dropped from every dataset.

批量删除源库中每个SAS数据集中名称以指定后缀结尾的变量。默认后缀为`_U`，通常用于SDTM/ADaM数据集中的单位变量。首先在整个库中识别所有匹配的变量，然后从每个数据集中删除。

## Parameters | 参数

| Parameter | Description | Default | Required |
|-----------|-------------|---------|----------|
| `suffix` | Variable name suffix to match (case-insensitive) / 要匹配的变量名后缀（不区分大小写） | _U | No |
| `src`    | Source library path / 源库路径 | — | **Yes** |
| `tgt`    | Target library path / 目标库路径 | — | **Yes** |

## How It Works | 工作原理

1. Assigns SAS libraries to the source and target directories.
2. Queries `dictionary.columns` to find all DISTINCT variable names in the source library that end with the specified suffix (case-insensitive LIKE match).
3. If no matching variables are found, prints a WARNING and exits.
4. Otherwise, iterates through all datasets in the library and drops the collected list of suffixed variables from each copy.
5. Clears the library assignments.

## Usage Examples | 使用示例

### Example 1: Drop all _U (unit) variables
```
%dropVarSuffix(suffix=_U, src=D:\raw, tgt=D:\clean);
```

### Example 2: Drop variables with a custom suffix
```
%dropVarSuffix(suffix=_OLD, src=C:\data\derived, tgt=C:\data\final);
```

### Example 3: Use the default suffix (_U)
```
%dropVarSuffix(src=D:\sdtm, tgt=D:\sdtm_noUnits);
```

## Notes | 注意事项

- The suffix matching is done via SQL `LIKE` with a `%` wildcard, so `_U` matches names like `AGE_U`, `WEIGHT_U`, `LBSTRESC_U`, etc.
- The variable list is collected once from the entire library — the same `DROP` list is applied to every dataset. If a variable in the list does not exist in a particular dataset, SAS issues a WARNING but continues.
- The source and target directories should differ to avoid overwriting.
- The macro variable `_keeplist` is named for historical reasons; it actually contains the list of variables to **drop**.

## Source Reference | 源文件参考
DM_Toolbox.sas, lines 584-612
