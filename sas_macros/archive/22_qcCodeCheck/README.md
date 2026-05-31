# Macro 22: %qcCodeCheck

## 分类 Category
Quality Control / 质量控制

## 用途 Purpose

Cross-validate MedDRA (AE) and WHODD (CM) coding data by running 7 automated consistency checks. Flagged records are written to an Excel workbook (`Coding_QC.xls`), with one sheet per check.

对 MedDRA（不良事件）和 WHODD（合并用药）编码数据进行交叉验证，执行 7 项自动一致性检查。有问题的记录写入 Excel 工作簿，每个检查一个 sheet。

### AE Checks (MedDRA):
| #   | Sheet Name                   | 检查内容                       |
|-----|------------------------------|-------------------------------|
| AE-1| AE1_SameTerm_DiffLLTCode     | 同一AE名称，不同LLT编码         |
| AE-2| AE2_SameTerm_DiffPTCode      | 同一AE名称，不同PT编码          |
| AE-3| AE3_SameLLT_DiffLLTCode      | 同一LLT名称，不同LLT编码        |
| AE-4| AE4_LLT_eq_PT_DiffCode       | LLT=PT但编码不同               |

### CM Checks (WHODD):
| #   | Sheet Name                   | 检查内容                       |
|-----|------------------------------|-------------------------------|
| CM-1| CM1_SameDrug_DiffDrugCode    | 同一药名，不同药品编码          |
| CM-2| CM2_SameDrug_DiffPrefName    | 同一药名，不同首选名称          |
| CM-3| CM3_SameCodeInd_DiffATC      | 同一编码+适应症，不同ATC4       |

## 参数表 Parameter Table

| Parameter | Required | Default         | Description |
|-----------|----------|-----------------|-------------|
| project   | YES      | -               | Project ID filter |
| out       | YES      | -               | Output directory for Coding_QC.xls |
| aeFile    | -        | -               | Path to MedDRA coding Excel file |
| aeSheet   | -        | MedDRA-DM-13-01 | AE sheet name in Excel |
| cmFile    | -        | -               | Path to WHODD coding Excel file |
| cmSheet   | -        | WHODD-DM-13-02  | CM sheet name in Excel |
| aeTerm    | -        | _c7_            | AE reported term column |
| llt       | -        | _c8_            | LLT name column |
| lltCode   | -        | _c9_            | LLT code column |
| pt        | -        | _c10_           | PT name column |
| ptCode    | -        | _c11_           | PT code column |
| drug      | -        | _c8_            | Drug name column (CM) |
| ind       | -        | _c9_            | Indication column (CM) |
| drugCode  | -        | _c17_           | Drug code column (CM) |
| atcName   | -        | _c18_           | ATC name column (CM) |
| atc4      | -        | _c26_           | ATC4 code column (CM) |

## 使用示例 Usage Examples

### Example 1: Full AE + CM check with Excel files
```sas
%qcCodeCheck(
    project = ZGJAK018,
    out     = C:\QC,
    aeFile  = C:\Data\MedDRA_Coding.xlsx,
    cmFile  = C:\Data\WHODD_Coding.xlsx
);
```

### Example 2: AE-only check with custom columns
```sas
%qcCodeCheck(
    project = ZGJAK018,
    out     = C:\QC,
    aeFile  = C:\Data\MedDRA_Coding.xlsx,
    aeTerm  = AE_TERM,
    llt     = LLT_NAME,
    lltCode = LLT_CODE,
    pt      = PT_NAME,
    ptCode  = PT_CODE
);
```

### Example 3: CM-only check
```sas
%qcCodeCheck(
    project = ZGJAK018,
    out     = C:\QC,
    cmFile  = C:\Data\WHODD_Coding.xlsx,
    drug     = DRUGTERM,
    ind      = INDICATION,
    drugCode = DRUGCODE,
    atcName  = ATCNAME,
    atc4     = ATC4
);
```

## 注意 Notes

1. **Input format** — The input Excel files should have the coding data starting from row 1 with no header. Column `_c1_` (first column) is expected to contain the project ID.
2. **Pre-loaded data** — If `aeFile` and `cmFile` are omitted, the macro will look for already-loaded `_ae` and `_cm` datasets in the WORK library.
3. **Output format** — `Coding_QC.xls` (legacy Excel format, compatible with SAS libname engine).
4. **All checks** are performed via LAG function with proper sorting; duplicate rows with identical values are skipped.
