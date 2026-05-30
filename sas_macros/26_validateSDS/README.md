# Macro 26: %validateSDS

## 分类 Category
Quality Control / 质量控制

## 用途 Purpose

Comprehensive SDS (Study Design Specification) validation with 25 rule-based checks. Reads all key sheets from an SDS Excel workbook (DatadictionaryEntry, DataDictionary, Field, VisitWindowSetting) and optionally a LabKey reference file. Flagged violations are written to a timestamped Excel workbook with one sheet per rule.

对 SDS（研究设计规格书）进行 25 项规则的全面验证。从 SDS Excel 工作簿读取关键 sheet，标记违规记录并导出到带时间戳的多 sheet Excel 报告。

### Rules:

| #  | Sheet                     | Rule Category         | 检查内容                                   |
|----|---------------------------|-----------------------|--------------------------------------------|
| 01 | 01_DupItemDataString      | Duplicate             | 重复的 ItemDataString 行（仅活动条目）        |
| 02 | 02_ClinicalType_Missing   | Clinical Dictionary   | 临床字典缺少 ClinicalType/ClinicalQueryString |
| 03 | 03_FieldOID_Over8         | Naming Convention     | FieldOID 超过 8 个字符                    |
| 04 | 04_FieldName_Over40       | Naming Convention     | FieldName 超过 40 个字符                   |
| 05 | 05_DateCtrl_NoUnit        | Date Control          | 日期控件缺少 Unit                           |
| 06 | 06_Grid_NoLabel           | Grid                 | 网格字段缺少 Label                          |
| 07 | 07_Select_NoDictID        | Selection            | 单选/下拉/复选框缺少 DataDictionaryID      |
| 08 | 08_Label_HasReviewGroup   | Text Label           | Text-label 有不期望的 ReviewGroups        |
| 09 | 09_Label_RequireVerify    | Text Label           | Text-label 设置了 RequireVerification=1   |
| 10 | 10_NonSelect_HasDictID    | Data Dictionary      | 非选择/非Lab字段有不期望的 DataDictionaryID |
| 11 | 11_NoReviewGroup          | Review Groups        | 非 Text-label 字段缺少 ReviewGroups        |
| 12 | 12_NoVerify               | Verification         | 非 Text-label 字段 RequireVerification=0   |
| 13 | 13_Date_NoFutureCheck     | Date Validation      | yyyy-MM-dd 格式缺少 IsFutureDateTime=1    |
| 14 | 14_FieldOID_Ne_VariableNo | Naming Consistency   | FieldOID 不等于 VariableNo               |
| 15 | 15_FieldOID_Ne_LabKey     | LabKey               | FieldOID 不等于 labKey                    |
| 16 | 16_DupLabKey              | LabKey               | 重复的 labKey 值                          |
| 17 | 17_LabNameMismatch        | LabKey               | FieldName 与 LabKey KeyDescription 不匹配 |
| 18 | 18_FormOID_Ne_SASText      | Naming Consistency   | FormOID 不等于 SASText                    |
| 19 | 19_DefaultVal_RowAdd       | Grid                 | 网格有默认值但允许添加行                    |
| 20 | 20_Window_MissingGroup    | Visit Window         | 访视窗口缺少 inGroup 或 outGroup           |
| 21 | 21_DM_View_NoReview       | View Restriction     | DM 视图限制但无 DM Review 组              |
| 22 | 22_CRA_View_NoVerify      | View Restriction     | CRA 视图限制但 RequireVerification=1      |
| 23 | 23_PI_View_NoSign         | View Restriction     | PI 视图限制但 IsRequireSign=1             |
| 24 | 24_DupVariableNo          | Duplicate            | 重复的 VariableNo                         |
| 25 | 25_Lab_NotClinicalRequired | Lab                  | Lab 字段但 IsClinicalRequired=0           |

## 参数表 Parameter Table

| Parameter | Required | Description |
|-----------|----------|-------------|
| sds       | YES      | Full path to the SDS Excel workbook |
| labKey    | -        | Full path to the LabKey reference Excel file (optional) |
| out       | YES      | Output directory for the validation report |

## 使用示例 Usage Examples

### Example 1: Full SDS validation with LabKey
```sas
%validateSDS(
    sds    = C:\Specs\SDS_ProjectX.xlsx,
    labKey = C:\Specs\LabKey_Ref.xlsx,
    out    = C:\QC\SDS_Validation
);
```

### Example 2: SDS validation without LabKey reference
```sas
%validateSDS(
    sds = D:\Study123\SDS\SDS_v2.0.xlsx,
    out = D:\Study123\QC
);
```

### Example 3: Quick check on a draft SDS
```sas
%validateSDS(
    sds = C:\Temp\SDS_Draft.xlsx,
    out = C:\Temp\QC
);
```

## 注意 Notes

1. **Sheet names** — The macro expects exact sheet names: `DatadictionaryEntry`, `DataDictionary`, `Field`, `VisitWindowSetting`, and optionally `LabKey`. The import will fail if any sheet is missing.
2. **Character encoding** — Column names from `getnames=yes` are imported as-is from Excel. Ensure the SDS uses consistent column naming (e.g., `fieldOID`, not `FieldOID` if case mismatch causes issues).
3. **LabKey optional** — Rules 15-17 will be skipped if no LabKey file is provided. Rule 17 sheet is only created when LabKey data is available.
4. **Output format** — The output is a .xlsx (XML Excel) workbook containing one sheet per rule. Sheets with no violations also appear (empty).
5. **Performance** — For very large SDS files (>1000 fields), the macro may take 1-2 minutes to complete all 25 checks.
