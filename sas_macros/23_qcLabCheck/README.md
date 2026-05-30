# Macro 23: %qcLabCheck

## 分类 Category
Quality Control / 质量控制

## 用途 Purpose

Run 4 automated QC checks on laboratory data (LB domain). Detects missing metadata, extreme outliers (>10x normal range), inconsistent units across visits, and non-numeric characters in result fields. Flagged records are exported to `Lab_QC.xls` with one sheet per check.

对实验室数据执行 4 项自动QC检查。检测缺失元数据、极端异常值（>10倍正常范围）、跨访视单位不一致、结果字段含非数字字符。问题记录导出到 Excel 工作簿，每个检查一个 sheet。

### Checks:
| # | Sheet                     | 检查内容                              |
|---|---------------------------|--------------------------------------|
| 1 | 01_Missing_Unit_Range     | 有结果但缺失单位/正常范围             |
| 2 | 02_Range_Outlier          | 结果值 > 10倍正常上限 或 *10 < 下限   |
| 3 | 03_Inconsistent_Unit      | 同一检验项目在不同访视单位不一致      |
| 4 | 04_NonNumeric_Result      | 结果字段含非数字字符（如 "<5", ">10"）|

## 参数表 Parameter Table

| Parameter | Required | Description |
|-----------|----------|-------------|
| lib       | YES      | Path to SAS data library containing the lab dataset |
| dsn       | YES      | Lab dataset name (e.g., `lb`, `lb2`) |
| out       | YES      | Output directory for `Lab_QC.xls` |

## 使用示例 Usage Examples

### Example 1: Basic lab QC
```sas
%qcLabCheck(lib=D:\Derived, dsn=lb, out=D:\QC);
```

### Example 2: Post-processing check on derived lab data
```sas
%qcLabCheck(lib=D:\Derived, dsn=lb2, out=D:\QC);
```

### Example 3: Check with custom library
```sas
libname mydata "C:\Project\Data";
%qcLabCheck(lib=C:\Project\Data, dsn=lab_results, out=C:\Project\QC);
```

## 注意 Notes

1. **Expected variables** — The macro expects the dataset to contain: `usubjid`, `siteid`, `visit`, `form`, `LBTEST`, `LBORRES`, `LBORRESU`, `LBORNRLO`, `LBORNRHI`, `LBSIG`. Missing variables will cause a SAS error.
2. **Numeric conversion** — Non-numeric results that fail `input(..., best.)` will be treated as missing in check 2 (range outlier). They will be flagged in check 4.
3. **Unit consistency** — Check 3 flags records where the same LBTEST has different LBORRESU across rows, regardless of subject. Review these to confirm whether the unit change is legitimate (e.g., g/dL vs mg/dL).
4. **Output format** — `Lab_QC.xls` uses the SAS libname Excel engine (legacy format).
