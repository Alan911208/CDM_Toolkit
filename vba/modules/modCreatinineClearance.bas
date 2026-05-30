Attribute VB_Name = "modCreatinineClearance"
Option Explicit

' Module: modCreatinineClearance
' Purpose: Creatinine Clearance / eGFR calculation
'   - DM_CalcCockcroftGault: Cockcroft-Gault formula (mL/min)
'   - DM_CalcCK DEPI: CKD-EPI 2021 formula (mL/min/1.73m2)
'   - DM_BatchCreatinineClearance: Batch calculate from worksheet columns
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

'============================================================================
' DM_CalcCockcroftGault
' Cockcroft-Gault creatinine clearance formula.
' CrCl (mL/min) = ((140 - age) * weight_kg) / (72 * Scr_mg_dL)
'                  * 0.85 (if female)
' Parameters:
'   age       - Age in years
'   weightKg  - Weight in kg
'   scrMgDL   - Serum creatinine in mg/dL
'   isFemale  - True for female (applies 0.85 factor)
' Returns: Double - CrCl in mL/min
'============================================================================
Public Function DM_CalcCockcroftGault(ByVal age As Double, _
    ByVal weightKg As Double, _
    ByVal scrMgDL As Double, _
    Optional ByVal isFemale As Boolean = False) As Double

    On Error GoTo ErrHandler

    If scrMgDL <= 0 Then
        DM_CalcCockcroftGault = -1: Exit Function
    End If

    Dim result As Double
    result = ((140 - age) * weightKg) / (72 * scrMgDL)

    If isFemale Then result = result * 0.85
    If result < 0 Then result = 0

    DM_CalcCockcroftGault = Round(result, 2)
    Exit Function

ErrHandler:
    DM_CalcCockcroftGault = -1
End Function

'============================================================================
' DM_CalcCK DEPI
' CKD-EPI 2021 formula for eGFR.
' eGFR = 142 * (Scr/A)^B * (0.9938)^age * (1.012 if female)
' Where:
'   A = 0.7 (female) or 0.9 (male)
'   B = -0.241 (female, Scr<=A) or -1.200 (female, Scr>A)
'     = -0.302 (male, Scr<=A)   or -1.200 (male, Scr>A)
' Parameters:
'   age       - Age in years
'   scrMgDL   - Serum creatinine in mg/dL
'   isFemale  - True for female
' Returns: Double - eGFR in mL/min/1.73m2
'============================================================================
Public Function DM_CalcCK DEPI(ByVal age As Double, _
    ByVal scrMgDL As Double, _
    Optional ByVal isFemale As Boolean = False) As Double

    On Error GoTo ErrHandler

    If scrMgDL <= 0 Then
        DM_CalcCK DEPI = -1: Exit Function
    End If

    Dim A As Double, B As Double

    If isFemale Then
        A = 0.7
        If scrMgDL <= A Then B = -0.241 Else B = -1.2
    Else
        A = 0.9
        If scrMgDL <= A Then B = -0.302 Else B = -1.2
    End If

    Dim result As Double
    result = 142 * ((scrMgDL / A) ^ B) * (0.9938 ^ age)

    If isFemale Then result = result * 1.012
    If result < 0 Then result = 0

    DM_CalcCK DEPI = Round(result, 2)
    Exit Function

ErrHandler:
    DM_CalcCK DEPI = -1
End Function

'============================================================================
' DM_BatchCreatinineClearance
' Batch calculate CrCl/eGFR from worksheet columns and write to result column.
' Parameters:
'   ws         - Worksheet containing patient data
'   method     - "CG" for Cockcroft-Gault, "CKD" for CKD-EPI 2021
'   ageCol     - Column letter for age (years)
'   weightCol  - Column letter for weight (kg) - CG only, ignored for CKD
'   scrCol     - Column letter for serum creatinine (mg/dL)
'   genderCol  - Column letter for gender ("M"/"F" or "Male"/"Female")
'   resultCol  - Column letter to write results
'   rowStart   - First data row (default 2)
' Returns: Long - number of calculations performed
'============================================================================
Public Function DM_BatchCreatinineClearance(ByVal ws As Worksheet, _
    ByVal method As String, _
    ByVal ageCol As String, _
    ByVal scrCol As String, _
    ByVal genderCol As String, _
    ByVal resultCol As String, _
    Optional ByVal weightCol As String = "", _
    Optional ByVal rowStart As Long = 2) As Long

    On Error GoTo ErrHandler

    Dim age As Double, weight As Double, scr As Double
    Dim isFemale As Boolean
    Dim lastRow As Long, i As Long, count As Long
    Dim colAge As Long, colWt As Long, colScr As Long, colGen As Long, colRes As Long
    Dim genderStr As String

    colAge = ws.Range(ageCol & "1").Column
    colScr = ws.Range(scrCol & "1").Column
    colGen = ws.Range(genderCol & "1").Column
    colRes = ws.Range(resultCol & "1").Column
    If Len(weightCol) > 0 Then colWt = ws.Range(weightCol & "1").Column

    lastRow = ws.Cells(ws.Rows.Count, colAge).End(xlUp).Row
    If lastRow < rowStart Then Exit Function

    method = UCase(Trim(method))

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = rowStart To lastRow
        ' Parse age
        If IsNumeric(ws.Cells(i, colAge).Value) Then
            age = CDbl(ws.Cells(i, colAge).Value)
        Else
            GoTo NextRow
        End If

        ' Parse serum creatinine
        If IsNumeric(ws.Cells(i, colScr).Value) Then
            scr = CDbl(ws.Cells(i, colScr).Value)
        Else
            GoTo NextRow
        End If

        ' Parse gender
        genderStr = UCase(Trim(CStr(ws.Cells(i, colGen).Value)))
        isFemale = (genderStr = "F" Or genderStr = "FEMALE")

        ' Parse weight (CG only)
        If method = "CG" And Len(weightCol) > 0 Then
            If IsNumeric(ws.Cells(i, colWt).Value) Then
                weight = CDbl(ws.Cells(i, colWt).Value)
            Else
                GoTo NextRow
            End If
        End If

        ' Calculate
        Dim result As Double
        If method = "CG" Then
            result = DM_CalcCockcroftGault(age, weight, scr, isFemale)
        ElseIf method = "CKD" Then
            result = DM_CalcCK DEPI(age, scr, isFemale)
        Else
            GoTo NextRow
        End If

        ws.Cells(i, colRes).Value = result
        ws.Cells(i, colRes).NumberFormat = "0.00"
        count = count + 1

NextRow:
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_BatchCreatinineClearance = count
    Exit Function

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DM_BatchCreatinineClearance Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function
