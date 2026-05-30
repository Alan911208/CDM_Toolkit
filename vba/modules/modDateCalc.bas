Attribute VB_Name = "modDateCalc"
Option Explicit

' Module: modDateCalc
' Purpose: CDM Date Calculator
'   - DM_AddDaysToColumn: Add/subtract days to every date in a column
'   - DM_CalcDateDiff: Compute (colEnd - colStart) in days, write to result column
'   - DM_DateAddSelection: Convenience - add days to selected cells
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Public Function DM_AddDaysToColumn(ByVal ws As Worksheet, _
    ByVal colLetter As String, _
    ByVal days As Long, _
    Optional ByVal rowStart As Long = 2, _
    Optional ByVal dateFormat As String = "YYYY-MM-DD") As Long

    On Error GoTo ErrHandler

    Dim col As Long
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim count As Long

    col = ws.Range(colLetter & "1").Column
    lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).Row

    If lastRow < rowStart Then Exit Function

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = rowStart To lastRow
        Set cell = ws.Cells(i, col)
        If IsDate(cell.Value) Then
            cell.Value = DateAdd("d", days, CDate(cell.Value))
            cell.NumberFormat = dateFormat
            count = count + 1
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_AddDaysToColumn = count
    Exit Function

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_AddDaysToColumn = -1
    MsgBox "DM_AddDaysToColumn ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_CalcDateDiff(ByVal ws As Worksheet, _
    ByVal colStart As String, _
    ByVal colEnd As String, _
    ByVal resultCol As String, _
    Optional ByVal rowStart As Long = 2) As Long

    On Error GoTo ErrHandler

    Dim c1 As Long, c2 As Long, cr As Long
    Dim lastRow As Long
    Dim i As Long
    Dim d1 As Date, d2 As Date
    Dim count As Long

    c1 = ws.Range(colStart & "1").Column
    c2 = ws.Range(colEnd & "1").Column
    cr = ws.Range(resultCol & "1").Column
    lastRow = ws.Cells(ws.Rows.count, c1).End(xlUp).Row

    If lastRow < rowStart Then Exit Function

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For i = rowStart To lastRow
        If IsDate(ws.Cells(i, c1).Value) And IsDate(ws.Cells(i, c2).Value) Then
            d1 = CDate(ws.Cells(i, c1).Value)
            d2 = CDate(ws.Cells(i, c2).Value)
            ws.Cells(i, cr).Value = d2 - d1
            count = count + 1
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_CalcDateDiff = count
    Exit Function

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_CalcDateDiff = -1
    MsgBox "DM_CalcDateDiff ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Sub DM_DateAddSelection(ByVal days As Long)

    On Error GoTo ErrHandler

    Dim cell As Range
    Dim count As Long

    If TypeName(Selection) <> "Range" Then Exit Sub

    Application.ScreenUpdating = False

    For Each cell In Selection
        If IsDate(cell.Value) Then
            cell.Value = DateAdd("d", days, CDate(cell.Value))
            count = count + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox "??? " & count & " ???????", vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_DateAddSelection ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
