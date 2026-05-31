# batchRunDateDiff

**Category:** Date & Version Management

## Purpose / 目的
**English:** Apply `%runDateDiff` to every dataset in a library. This is the batch version of the comparison engine -- it iterates over all datasets in the old library and runs the single-dataset `%runDateDiff` on each one, producing a complete output library with version-tracked data.

**中文:** 将 `%runDateDiff` 应用于库中的每个数据集。这是比对引擎的批量版本 -- 它遍历旧库中的所有数据集，对每个数据集执行单表 `%runDateDiff`，生成一个包含版本追踪数据的完整输出库。

## Parameters / 参数

| Parameter | Description | Required | Example |
|-----------|-------------|----------|---------|
| `oldLib` | Path to old (baseline) data / 旧版（基线）数据路径 | Yes | `D:\old` |
| `newLib` | Path to new (current) data / 新版（当前）数据路径 | Yes | `D:\new` |
| `outLib` | Path to output data / 输出数据路径 | Yes | `D:\final` |
| `keyvar` | Key variables for row matching / 行匹配的关键变量 | Yes | `usubjid PAGE LINE` |
| `exvar` | Variables to exclude from comparison / 排除比对的变量 | Yes | `STUDYID SITE FORM rundate` |

## Usage Examples / 使用示例

### Example 1: Standard batch comparison
```sas
%batchRunDateDiff(newLib=D:\new, oldLib=D:\old, outLib=D:\final,
    keyvar=usubjid PAGE LINE, exvar=STUDYID SITE FORM rundate);
```

### Example 2: Batch comparison with medical history key
```sas
%batchRunDateDiff(newLib=D:\newData, oldLib=D:\oldData, outLib=D:\stampedData,
    keyvar=usubjid MHSEQ, exvar=STUDYID SITEID rundate);
```

### Example 3: Part of a delivery pipeline (after stampRunDate)
```sas
* Step 1: Stamp all new datasets with a fixed date;
%stampRunDate(lib=D:\new, date='15AUG2023'd);
* Step 2: Run batch comparison against old baseline;
%batchRunDateDiff(newLib=D:\new, oldLib=D:\baseline, outLib=D:\delivery,
    keyvar=usubjid VISIT PAGE LINE, exvar=STUDYID SITEID FORMID rundate);
```

## Internal Helpers / 内部辅助宏
- **%_getDatasetList** — Queries `dictionary.tables` to enumerate all datasets in the old library. Excludes the FORMATS catalog.
- **%runDateDiff** — The full single-dataset comparison macro (included in this file for stand-alone use).

## Logic / 逻辑
1. Assign librefs `_new`, `_old`, `_out` to the three directories.
2. Call `%_getDatasetList` to get all dataset names from the old library.
3. Loop through each dataset name:
   - Call `%runDateDiff` with the same dataset name for `oldset`, `newset`, and `outset`.
   - This means each output dataset mirrors its input name.
4. Clear all temporary librefs.

## Notes / 注意事项
- Each dataset in `oldLib` must also exist in `newLib` with the same name.
- Each dataset should already have a `rundate` field (e.g., set via `%stampRunDate` on the new library).
- The output library `outLib` directory must exist and be writeable.
- All three paths (`oldLib`, `newLib`, `outLib`) must be accessible filesystem paths.
