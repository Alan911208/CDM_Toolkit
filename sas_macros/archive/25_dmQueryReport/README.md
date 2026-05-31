# Macro 25: %dmQueryReport

## 分类 Category
Reporting / 报表生成

## 用途 Purpose

Generate the DMR Query Summary tables (T5.1-T5.3) in RTF format. Reads query detail and form metadata from Excel sheets and produces formatted RTF output suitable for inclusion in DMR (Data Management Report) documents.

从 Excel 读取 Query 明细和 Form 元数据，生成 DMR Query Summary 三张表（T5.1-T5.3）并输出为 RTF 格式。

### Tables:
| Table | Title                                   | 说明                              |
|-------|----------------------------------------|-----------------------------------|
| T5.1  | Query Counts by Form + Top-4 Texts      | 按表单的Query数量及Top-4文本       |
| T5.2  | Response Time by Query Type             | 按Query类型的响应时间统计          |
| T5.3  | Site-Level Key Findings                 | 按中心的Query关键指标              |

## 参数表 Parameter Table

| Parameter | Required | Default | Description |
|-----------|----------|---------|-------------|
| qSheet    | YES      | -       | Excel file path and sheet name for query detail |
| form      | YES      | -       | Excel file path and sheet name for form metadata |
| outRTF    | YES      | -       | Output RTF file path |
| ver       | -        | 3       | Template version: 2 (legacy) or 3 (v3.0) |

## 使用示例 Usage Examples

### Example 1: Standard v3.0 DMR query report
```sas
%dmQueryReport(
    qSheet = D:\DMR\QueryDetail.xlsx,
    form   = D:\DMR\Form.xlsx,
    outRTF = D:\output\DMR_Query.rtf,
    ver    = 3
);
```

### Example 2: Legacy v2 format
```sas
%dmQueryReport(
    qSheet = D:\DMR\Legacy_Queries.xlsx,
    form   = D:\DMR\Legacy_Forms.xlsx,
    outRTF = D:\output\DMR_Query_v2.rtf,
    ver    = 2
);
```

### Example 3: Same file, different sheets
```sas
%dmQueryReport(
    qSheet = C:\DMR\QueryData.xlsx,
    form   = C:\DMR\FormMeta.xlsx,
    outRTF = C:\DMR\Output\T5_QuerySummary.rtf
);
```

## 注意 Notes

1. **Template versions** — `ver=2` uses legacy column mapping (`_c8_` as form key, `_c21_` as status), while `ver=3` (default) uses `_c9_` as form key and `_c22_` as status.
2. **Cancelled queries excluded** — Queries with status "Cancel" or "Cancelled" are removed from all calculations.
3. **Date parsing** — Query Open Date and Answer Date are parsed using `yymmdd10.` format. Ensure dates are consistently formatted (YYYY-MM-DD or similar).
4. **Output style** — Uses the SAS `journal` ODS style for RTF output.
5. **Field_OID** — In ver=3, the field OID column (`c`) comes from the form metadata merge, linking fields to forms.
