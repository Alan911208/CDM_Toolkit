Attribute VB_Name = "modSheetCreate"
Option Explicit

' Module: modSheetCreate
' Purpose: Batch create empty sheets from a name list
'   - DM_CreateSheetsFromList: Create multiple sheets from comma/line/array input
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Public Function DM_CreateSheetsFromList(ByVal wb As Workbook, _
    ByVal sheetNames As Variant, _
    Optional ByVal clearFirst As Boolean = False) As Long

    On Error GoTo ErrHandler

    Dim names() As String, i As Long, nameStr As String
    Dim created As Long, ws As Worksheet

    If IsArray(sheetNames) Then
        ReDim names(0 To UBound(sheetNames))
        For i = 0 To UBound(sheetNames)
            names(i) = CStr(sheetNames(i))
        Next i
    Else
        nameStr = CStr(sheetNames)
        nameStr = Replace(nameStr, vbCrLf, ",")
        nameStr = Replace(nameStr, vbCr, ",")
        nameStr = Replace(nameStr, vbLf, ",")
        names = Split(nameStr, ",")
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 0 To UBound(names)
        nameStr = Trim(names(i))
        If Len(nameStr) = 0 Then GoTo NextName

        If Len(nameStr) > 31 Then nameStr = Left(nameStr, 31)

        Dim illegal As String, j As Long
        illegal = "\/*?:[]"
        For j = 1 To Len(illegal)
            nameStr = Replace(nameStr, Mid(illegal, j, 1), "_")
        Next j

        If Len(nameStr) = 0 Then GoTo NextName

        On Error Resume Next
        Set ws = wb.Worksheets(nameStr)
        If Err.Number = 0 Then
            If clearFirst Then
                ws.Delete
            Else
                Dim suffix As Long: suffix = 1
                Dim base As String: base = nameStr
                Do While Err.Number = 0
                    Err.Clear
                    nameStr = Left(base, 28) & "_" & suffix
                    suffix = suffix + 1
                    Set ws = wb.Worksheets(nameStr)
                Loop
            End If
        End If
        Err.Clear
        On Error GoTo ErrHandler

        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = nameStr
        created = created + 1

NextName:
    Next i

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    DM_CreateSheetsFromList = created
    Exit Function

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "DM_CreateSheetsFromList Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function
