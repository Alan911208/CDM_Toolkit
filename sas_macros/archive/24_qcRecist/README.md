# Macro 24: %qcRecist

## 分类 Category
Quality Control / 质量控制

## 用途 Purpose

Verify RECIST 1.1 tumor target response evaluation consistency by comparing EDC-entered response (CR/PR/SD/PD) against calculated values from target-lesion measurements. Also checks lesion-level method and location consistency across visits.

通过比较EDC录入的靶病灶疗效评估（CR/PR/SD/PD）与从靶病灶测量值计算出的结果，验证RECIST 1.1疗效评价一致性。同时检查病灶层面的评估方法和位置在不同访视间的一致性。

### Two output sheets:
- **Recist** — Target response comparison (EDC vs calculated CR/PR/SD/PD)
- **Consistency** — Lesion-level method and location consistency across visits

## 参数表 Parameter Table

| Parameter      | Required | Description |
|----------------|----------|-------------|
| lib            | YES      | Path to SAS data library |
| out            | YES      | Output directory for Excel reports |
| scrDS          | YES      | Screening visit dataset name |
| fuDS           | YES      | Follow-up visit dataset name |
| respDS         | YES      | Dataset with EDC-assessed overall response |
| respVar        | YES      | EDC response variable name (e.g., TRGRESP) |
| scrDateVar     | -        | Screening assessment date variable |
| fuDateVar      | -        | Follow-up assessment date variable |
| scrLongVar     | -        | Screening longest diameter variable |
| fuLongVar      | -        | Follow-up longest diameter variable |
| scrSumVar      | -        | Screening sum of diameters variable |
| fuSumVar       | -        | Follow-up sum of diameters variable |
| scrLocVar      | -        | Screening lesion location variable |
| fuLocVar       | -        | Follow-up lesion location variable |
| scrSiteVar     | -        | Screening location-within-site variable |
| fuSiteVar      | -        | Follow-up location-within-site variable |
| scrRecVar      | -        | Screening record position variable |
| fuRecVar       | -        | Follow-up record position variable |
| scrMetVar      | -        | Screening method variable |
| fuMetVar       | -        | Follow-up method variable |
| scrMetSpecVar  | -        | Screening method-specify variable |
| fuMetSpecVar   | -        | Follow-up method-specify variable |
| scrLabel       | -        | Screening visit label (default: Screening) |

## 使用示例 Usage Examples

### Example 1: Basic RECIST QC
```sas
%qcRecist(
    lib      = D:\Derived,
    out      = D:\QC,
    scrDS    = _1recist1,
    fuDS     = _3recist1,
    respDS   = overall,
    respVar  = TRGRESP,
    scrDateVar = TRCDAT1,
    fuDateVar  = TRCDAT1,
    scrSumVar  = SLDIAM,
    fuSumVar   = SLDIAM,
    scrLocVar  = TULOC1,
    fuLocVar   = TULOC1,
    scrSiteVar = S_R1SIT,
    fuSiteVar  = S_R1SIT,
    scrLabel   = Screening
);
```

### Example 2: Full parameter specification
```sas
%qcRecist(
    lib      = C:\Project\Data,
    out      = C:\Project\QC,
    scrDS    = _1recist1,    fuDS    = _3recist1,
    respDS   = recist2,      respVar = R3IORESP,
    scrDateVar = TRCDAT1,    fuDateVar  = TRCDAT1,
    scrLongVar = LDIAM,      fuLongVar  = LDIAM,
    scrSumVar  = SLDIAM,     fuSumVar   = SLDIAM,
    scrLocVar  = TULOC1,     fuLocVar   = TULOC1,
    scrSiteVar = S_R1SIT,    fuSiteVar  = S_R1SIT,
    scrRecVar  = RecordPosition, fuRecVar  = RecordPosition,
    scrMetVar  = TRMETHD,    fuMetVar   = TRMETHD,
    scrMetSpecVar = TRMETHDO, fuMetSpecVar = TRMETHDO,
    scrLabel   = Screening
);
```

## 注意 Notes

1. **RECIST 1.1 criteria** are hard-coded: PD (>20% increase + >5mm), PR (>=30% decrease), SD (neither), CR (all lesions disappeared, lymph <10mm).
2. **Key variables** — The screening and follow-up datasets must contain `Subject` and `InstanceName` (visit label) variables.
3. **Date formats** — Dates are parsed via both `date8.` and `date9.` format to handle common SAS date representations.
4. **Multiple standards** — If a subject triggers both PD and PR (or SD and PD, etc.) simultaneously, the `q` column is set to "Multiple standards" to flag the inconsistency.
5. **Output** — `Recist_YYYYMMDD.xlsx` with two sheets: "Recist" (response comparison) and "Consistency" (lesion details).
