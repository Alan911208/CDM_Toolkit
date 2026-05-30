Attribute VB_Name = "modUnitConvert"
Option Explicit

' Module: modUnitConvert
' Purpose: CDM Lab Unit Conversion - convert lab results to a standard unit
'   - DM_ConvertLabUnits: Convert using built-in conversion table (24 factor pairs)
'   - DM_GetSupportedConversions: List all available conversions
'   - DM_AddUnitConversion: Extend conversion table at runtime
'   Covers: Mass (g/mg/ug/kg), Volume (L/mL/dL), Length (cm/mm/m),
'           Concentration (g/L, mg/mL, mg/dL, umol/L), Time (h/min/d)
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Private mConversionTable As Object

Private Sub DM_InitConversionTable()
    If Not mConversionTable Is Nothing Then Exit Sub

    Set mConversionTable = CreateObject("Scripting.Dictionary")

    ' ??
    DM_AddConv "g", "mg", 1000
    DM_AddConv "g", "ug", 1000000
    DM_AddConv "mg", "g", 0.001
    DM_AddConv "mg", "ug", 1000
    DM_AddConv "ug", "g", 0.000001
    DM_AddConv "ug", "mg", 0.001
    DM_AddConv "kg", "g", 1000
    DM_AddConv "g", "kg", 0.001

    ' ??
    DM_AddConv "L", "mL", 1000
    DM_AddConv "mL", "L", 0.001
    DM_AddConv "dL", "L", 0.1
    DM_AddConv "L", "dL", 10

    ' ??
    DM_AddConv "cm", "mm", 10
    DM_AddConv "mm", "cm", 0.1
    DM_AddConv "m", "cm", 100
    DM_AddConv "cm", "m", 0.01

    ' ??
    DM_AddConv "g/L", "mg/mL", 1
    DM_AddConv "mg/mL", "g/L", 1
    DM_AddConv "mg/dL", "g/L", 0.01
    DM_AddConv "g/L", "mg/dL", 100
    DM_AddConv "g/dL", "g/L", 10
    DM_AddConv "g/L", "g/dL", 0.1
    DM_AddConv "umol/L", "mg/dL", 0.0113
    DM_AddConv "mg/dL", "umol/L", 88.42

    ' ??
    DM_AddConv "h", "min", 60
    DM_AddConv "min", "h", 1 / 60
    DM_AddConv "d", "h", 24
    DM_AddConv "h", "d", 1 / 24
End Sub

Private Sub DM_AddConv(ByVal fromUnit As String, ByVal toUnit As String, ByVal factor As Double)
    Dim key As String
    key = LCase(fromUnit) & "?" & LCase(toUnit)
    If Not mConversionTable.Exists(key) Then
        mConversionTable.Add key, factor
    End If
End Sub

Public Function DM_ConvertLabUnits(ByVal ws As Worksheet, _
    ByVal testCol As String, _
    ByVal resultCol As String, _
    ByVal unitCol As String, _
    ByVal standardUnit As String, _
    Optional ByVal testFilter As String = "") As Long

    On Error GoTo ErrHandler

    Call DM_InitConversionTable

    Dim tc As Long, rc As Long, uc As Long
    Dim lastRow As Long
    Dim i As Long
    Dim testName As String
    Dim unitVal As String
    Dim resultVal As Variant
    Dim factor As Variant
    Dim key As String
    Dim count As Long

    tc = ws.Range(testCol & "1").Column
    rc = ws.Range(resultCol & "1").Column
    uc = ws.Range(unitCol & "1").Column
    lastRow = ws.Cells(ws.Rows.count, tc).End(xlUp).Row

    If lastRow < 2 Then Exit Function

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = 2 To lastRow
        testName = CStr(ws.Cells(i, tc).Value)
        unitVal = CStr(ws.Cells(i, uc).Value)
        resultVal = ws.Cells(i, rc).Value

        If Len(testFilter) > 0 And LCase(testName) <> LCase(testFilter) Then GoTo NextRow
        If LCase(unitVal) = LCase(standardUnit) Then GoTo NextRow
        If Not IsNumeric(resultVal) Then GoTo NextRow

        key = LCase(unitVal) & "?" & LCase(standardUnit)
        If mConversionTable.Exists(key) Then
            factor = mConversionTable(key)
            ws.Cells(i, rc).Value = Round(CDbl(resultVal) * CDbl(factor), 6)
            ws.Cells(i, uc).Value = standardUnit
            count = count + 1
        Else
            ws.Cells(i, rc).Interior.Color = RGB(255, 153, 153)
        End If

NextRow:
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_ConvertLabUnits = count
    Exit Function

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_ConvertLabUnits = -1
    MsgBox "DM_ConvertLabUnits ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_GetSupportedConversions() As String()
    Call DM_InitConversionTable

    Dim result() As String
    Dim key As Variant
    Dim i As Long

    ReDim result(0 To mConversionTable.count - 1)
    i = 0
    For Each key In mConversionTable.Keys
        result(i) = CStr(key) & " = " & mConversionTable(key)
        i = i + 1
    Next key
    DM_GetSupportedConversions = result
End Function

Public Sub DM_AddUnitConversion(ByVal fromUnit As String, _
    ByVal toUnit As String, ByVal factor As Double)

    Call DM_InitConversionTable
    DM_AddConv fromUnit, toUnit, factor
End Sub
