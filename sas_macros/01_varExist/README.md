# %varExist

## Category
Data Inspection & Validation

## Purpose / 目的

**English:** Test whether a given variable exists in a SAS dataset. The result is returned via the global macro variable `&rc`: `0` means the variable was not found, `1` means it exists. This is a fast, dictionary-free check using `OPEN` and `VARNUM` functions — ideal for use inside conditional macro logic.

**中文：** 检测 SAS 数据集中是否存在指定变量。结果通过全局宏变量 `&rc` 返回：`0` 表示未找到变量，`1` 表示存在。该宏使用 `OPEN` 和 `VARNUM` 函数实现快速检测，无需查询数据字典，适用于宏条件逻辑中。

## Parameters / 参数

| Parameter | Required | Description |
|-----------|----------|-------------|
| `dsn`     | Yes      | Dataset name (1- or 2-level, e.g. `work.demo` or `sashelp.class`) |
| `var`     | Yes      | Variable name to search for |

## Return / 返回值

Global macro variable `&rc`:
- `0` — variable does NOT exist
- `1` — variable exists

## Usage Examples / 使用示例

### Example 1: Check existence and conditionally execute code
```sas
%varExist(dsn=work.ae, var=AESTDTC);
%if &rc = 1 %then %do;
    proc print data=work.ae; var AESTDTC; run;
%end;
%else %put WARNING: AESTDTC not found, skipping print;
```

### Example 2: Loop-based variable check
```sas
%let vars = A B C D;
%do i = 1 %to 4;
    %let v = %scan(&vars, &i);
    %varExist(dsn=work.dm, var=&v);
    %put NOTE: Variable &v exists? &rc;
%end;
```

### Example 3: Validate input before a merge
```sas
%varExist(dsn=work.dm, var=USUBJID);
%if &rc = 0 %then %do;
    %put ERROR: USUBJID is required in DM dataset;
    %return;
%end;
```

## Notes / 注意事项

- The macro opens the dataset in INPUT mode (`open(&dsn,i)`), which is read-only and does not lock the dataset.
- If the dataset cannot be opened (e.g., does not exist), a WARNING is printed and `&rc` remains 0.
- If either `dsn` or `var` is blank, the macro prints an ERROR and returns immediately without setting `&rc`.
- `&rc` is declared as `%global`, so it is available after the macro call regardless of scope.
- This macro does NOT depend on any internal helper macros.

## Source Reference / 来源

`DM_Toolbox.sas` — Tool 1, Category 1: Data Inspection & Validation
