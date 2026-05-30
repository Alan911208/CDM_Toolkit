Attribute VB_Name = "modStrikethrough"
Option Explicit

' Module: modStrikethrough
' Purpose: CDM Strikethrough Content Management
'   - DM_DeleteStrikethroughRows: Delete all rows containing strikethrough cells
'   - DM_ClearStrikethrough: Remove strikethrough formatting from all cells
'   - DM_StrikethroughSelected: Toggle strikethrough on selected cells
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Public Function DM_DeleteStrikethroughRows(ByVal ws As Worksheet, _
    Optional ByVal confirmFirst As Boolean = True) As Long

    On Error GoTo ErrHandler

    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim hasStrike As Boolean
    Dim deleted As Long
    Dim cell As Range

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    If lastRow <= 1 Then Exit Function

    ' Count strikethrough rows first for confirmation
    Dim strikeCount As Long
    strikeCount = 0
    For i = 1 To lastRow
        hasStrike = False
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If cell.Font.Strikethrough And Not IsEmpty(cell.Value) Then
                hasStrike = True: Exit For
            End If
        Next j
        If hasStrike Then strikeCount = strikeCount + 1
    Next i

    If strikeCount = 0 Then
        MsgBox "No strikethrough content found.", vbInformation, "CDM Toolkit"
        DM_DeleteStrikethroughRows = 0: Exit Function
    End If

    If confirmFirst Then
        If MsgBox("Found " & strikeCount & " row(s) with strikethrough content." & vbCrLf & _
            "Delete them?", vbYesNo + vbQuestion, "CDM Toolkit - Confirm Delete") <> vbYes Then
            DM_DeleteStrikethroughRows = 0: Exit Function
        End If
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Delete bottom-up to preserve row indices
    For i = lastRow To 1 Step -1
        hasStrike = False
        For j = 1 To lastCol
            Set cell = ws.Cells(i, j)
            If cell.Font.Strikethrough And Not IsEmpty(cell.Value) Then
                hasStrike = True: Exit For
            End If
        Next j
        If hasStrike Then
            ws.Rows(i).Delete Shift:=xlUp
            deleted = deleted + 1
        End If
    Next i

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    DM_DeleteStrikethroughRows = deleted
    MsgBox deleted & " row(s) deleted.", vbInformation, "CDM Toolkit"
    Exit Function

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "DM_DeleteStrikethroughRows Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Sub DM_ClearStrikethrough(ByVal ws As Worksheet)
    On Error GoTo ErrHandler

    Dim cell As Range
    Dim count As Long

    Application.ScreenUpdating = False

    For Each cell In ws.UsedRange
        If cell.Font.Strikethrough Then
            cell.Font.Strikethrough = False
            count = count + 1
        End If
    Next cell

    Application.ScreenUpdating = True
    MsgBox count & " cell(s) cleared.", vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_ClearStrikethrough Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_StrikethroughSelected()
    On Error GoTo ErrHandler

    Dim cell As Range

    If TypeName(Selection) <> "Range" Then Exit Sub

    For Each cell In Selection
        cell.Font.Strikethrough = Not cell.Font.Strikethrough
    Next cell

    Exit Sub

ErrHandler:
    MsgBox "DM_StrikethroughSelected Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
