Attribute VB_Name = "modSheetManager"
Option Explicit

' Module: modSheetManager
' Purpose: CDM Sheet Manager - batch rename, list names, delete hidden/specified sheets
'   - DM_ListSheetNames: Return array of sheet names
'   - DM_BatchRenameSheets: Rename multiple sheets via Dictionary {old: new}
'   - DM_DeleteHiddenSheets: Remove all hidden sheets
'   - DM_DeleteSheets: Delete specified sheets with confirmation
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Public Function DM_ListSheetNames(ByVal wb As Workbook, _
    Optional ByVal includeHidden As Boolean = False) As String()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim count As Long
    Dim result() As String
    Dim i As Long

    count = 0
    For Each ws In wb.Worksheets
        If includeHidden Or ws.Visible = xlSheetVisible Then
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        ReDim result(0 To -1)
        DM_ListSheetNames = result
        Exit Function
    End If

    ReDim result(0 To count - 1)
    i = 0
    For Each ws In wb.Worksheets
        If includeHidden Or ws.Visible = xlSheetVisible Then
            result(i) = ws.Name
            i = i + 1
        End If
    Next ws

    DM_ListSheetNames = result
    Exit Function

ErrHandler:
    MsgBox "DM_ListSheetNames ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_BatchRenameSheets(ByVal wb As Workbook, _
    ByVal nameMap As Object) As Long

    On Error GoTo ErrHandler

    Dim key As Variant
    Dim newName As String
    Dim count As Long

    If nameMap Is Nothing Then Exit Function
    If nameMap.count = 0 Then Exit Function

    Application.ScreenUpdating = False

    For Each key In nameMap.Keys
        newName = CStr(nameMap(key))
        If Len(newName) > 31 Then newName = Left(newName, 31)

        On Error Resume Next
        wb.Worksheets(CStr(key)).Name = newName
        If Err.Number = 0 Then count = count + 1
        Err.Clear
        On Error GoTo ErrHandler
    Next key

    Application.ScreenUpdating = True
    DM_BatchRenameSheets = count
    Exit Function

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_BatchRenameSheets ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Function DM_DeleteHiddenSheets(ByVal wb As Workbook) As String()

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Dim count As Long
    Dim result() As String
    Dim i As Long
    Dim j As Long

    count = 0
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then
            count = count + 1
        End If
    Next ws

    If count = 0 Then
        ReDim result(0 To -1)
        DM_DeleteHiddenSheets = result
        Exit Function
    End If

    ReDim result(0 To count - 1)

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    i = 0
    For j = wb.Worksheets.count To 1 Step -1
        Set ws = wb.Worksheets(j)
        If ws.Visible <> xlSheetVisible Then
            result(i) = ws.Name
            ws.Delete
            i = i + 1
        End If
    Next j

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    DM_DeleteHiddenSheets = result
    Exit Function

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_DeleteHiddenSheets ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Sub DM_DeleteSheets(ByVal wb As Workbook, ParamArray sheetNames() As Variant)

    On Error GoTo ErrHandler

    Dim i As Long
    Dim sheetName As String
    Dim sheetList As String
    Dim ws As Worksheet

    If UBound(sheetNames) < 0 Then Exit Sub

    For i = 0 To UBound(sheetNames)
        sheetList = sheetList & vbCrLf & "  - " & CStr(sheetNames(i))
    Next i

    If MsgBox("??????????" & sheetList, vbYesNo + vbQuestion, "CDM Toolkit - ????") <> vbYes Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    For i = 0 To UBound(sheetNames)
        sheetName = CStr(sheetNames(i))
        On Error Resume Next
        Set ws = wb.Worksheets(sheetName)
        If Err.Number = 0 Then
            ws.Delete
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next i

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "DM_DeleteSheets ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
