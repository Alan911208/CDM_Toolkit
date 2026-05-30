# stampRunDate

**Category:** Date & Version Management

## Purpose / 目的
**English:** Assign a fixed run-date value to every dataset in a library. Each dataset receives a new variable `rundate` set to the specified date literal. This is typically used as a first step before running `%runDateDiff` to establish a baseline date stamp.

**中文:** 为库中的每个数据集分配一个固定的运行日期。每个数据集会新增变量 `rundate`，其值设为指定的 SAS 日期常量。通常在运行 `%runDateDiff` 之前使用，用于建立基线日期戳。

## Parameters / 参数

| Parameter | Description | Required | Example |
|-----------|-------------|----------|---------|
| `lib` | Path to the SAS data library / SAS 数据目录路径 | Yes | `C:\delivery` |
| `date` | SAS date literal / SAS 日期字面量 | Yes | `'07JUL2020'd` |

## Usage Examples / 使用示例

### Example 1: Stamp with a specific date
```sas
%stampRunDate(lib=C:\delivery, date='07JUL2020'd);
```

### Example 2: Stamp with today's date
```sas
%stampRunDate(lib=C:\delivery, date="&sysdate9."d);
```

### Example 3: Stamp with a computed date (7 days ago)
```sas
%let weekAgo = %sysfunc(intnx(day, %sysfunc(today()), -7));
%stampRunDate(lib=D:\rawData, date=&weekAgo);
```

## Internal Helpers / 内部辅助宏
- **%_getDatasetList** — Queries `dictionary.tables` to populate a space-delimited macro variable with dataset names from the given libref. Excludes the FORMATS catalog by default.

## Notes / 注意事项
- A temporary libref `_tmp1` is assigned and cleared within the macro.
- If a variable named `rundate` already exists, it is overwritten.
- The `date` parameter must evaluate to a numeric SAS date value (number of days since 01JAN1960).
- All datasets in the library are processed; the FORMATS catalog is excluded.
