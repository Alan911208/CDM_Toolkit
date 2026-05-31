# runDateDiff

**Category:** Date & Version Management

## Purpose / 目的
**English:** Compare old-vs-new of a single dataset at the observation level. Only changed rows receive today's date as their `rundate`; unchanged rows keep the original rundate from the old dataset. This is the core comparison engine for version tracking.

**中文:** 在观测级别上比较单个数据集的新旧版本。只有发生变化的行会获得今天的日期作为 `rundate`；未变化的行保留旧数据集中的原始 rundate。这是版本追踪的核心比对引擎。

## Parameters / 参数

| Parameter | Description | Required | Example |
|-----------|-------------|----------|---------|
| `oldlib` | Old (baseline) library / 旧版（基线）库 | Yes | `old` |
| `oldset` | Old dataset name / 旧数据集名 | Yes | `_mh2` |
| `newlib` | New (current) library / 新版（当前）库 | Yes | `new` |
| `newset` | New dataset name / 新数据集名 | Yes | `_mh2` |
| `outlib` | Output library / 输出库 | Yes | `out` |
| `outset` | Output dataset name / 输出数据集名 | Yes | `rst` |
| `keyvar` | Key variables for row matching / 行匹配的关键变量 | Yes | `usubjid PAGE LINE` |
| `exvar` | Variables to exclude from comparison / 排除比对变量 | Yes | `STUDYID SITE FORM rundate` |

## Usage Examples / 使用示例

### Example 1: Basic single-dataset comparison
```sas
%runDateDiff(oldlib=old, oldset=_mh2, newlib=new, newset=_mh2,
    outlib=out, outset=_mh2, keyvar=usubjid PAGE LINE,
    exvar=STUDYID SITE FORM rundate);
```

### Example 2: Compare lab data with compound key
```sas
%runDateDiff(oldlib=old, oldset=lb, newlib=new, newset=lb,
    outlib=out, outset=lb, keyvar=usubjid VISIT LBTESTCD,
    exvar=STUDYID SITEID LBSEQ LBDTC rundate);
```

### Example 3: Minimal comparison with single key
```sas
%runDateDiff(oldlib=old, oldset=dm, newlib=new, newset=dm,
    outlib=out, outset=dm, keyvar=usubjid,
    exvar=STUDYID SITEID rundate);
```

## Logic / 逻辑
1. Identify all columns in the new dataset, excluding keyvar and exvar.
2. Merge old and new datasets by keyvar. Old columns are renamed with a leading underscore prefix. The old `rundate` is renamed to `_rundate`.
3. For each observation in the new dataset:
   - If a matching row exists in the old dataset AND all comparable columns are identical: keep `_rundate` as the `rundate`.
   - Otherwise (new row or changed row): assign `&sysdate9.` as the `rundate`.
4. Drop temporary underscore-prefixed columns and `_rundate`.

## Notes / 注意事项
- Both `oldset` and `newset` must exist in their respective libraries.
- The new dataset should already contain a `rundate` variable (e.g., set by `%stampRunDate`).
- `rundate` should be included in `exvar` to prevent it from being compared as a data column.
- Key variables must uniquely identify rows within each dataset for correct matching.
