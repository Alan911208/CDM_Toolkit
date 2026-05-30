Attribute VB_Name = "modTrackChanges"
Option Explicit

' Module: modTrackChanges
' Purpose: CDM Track Changes - compare worksheet against original backup
'   - DM_TrackChanges: Compare & mark changes (green fill + red border + comment "Old->New")
'   - DM_ClearTrackChanges: Remove all change-tracking marks
'   - DM_ChangesReport: Generate structured Changes Report sheet
' Author: CDM Toolkit Team
' Version: 1.0 / 2026-05-30
' Dependency: None

Public Function DM_TrackChanges(ByVal ws As Worksheet, _
    ByVal originalFilePath As String, _
    Optional ByVal addComments As Boolean = True, _
    Optional ByVal highlightChanges As Boolean = True) As Long

    On Error GoTo ErrHandler

    Dim origWb As Workbook
    Dim origWs As Worksheet
    Dim cell As Range
    Dim origVal As Variant
    Dim changes As Long
    Dim cmt As Comment
    Dim sheetName As String

    sheetName = ws.Name

    If Dir(originalFilePath) = "" Then
        MsgBox "???????: " & originalFilePath, vbExclamation, "CDM Toolkit"
        DM_TrackChanges = -1
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set origWb = Workbooks.Open(originalFilePath, ReadOnly:=True)

    On Error Resume Next
    Set origWs = origWb.Worksheets(sheetName)
    If Err.Number <> 0 Then
        origWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "?????????sheet: " & sheetName, vbExclamation, "CDM Toolkit"
        DM_TrackChanges = -1
        Exit Function
    End If
    On Error GoTo ErrHandler

    Dim usedRange As Range
    Set usedRange = ws.UsedRange

    For Each cell In usedRange
        On Error Resume Next
        origVal = origWs.Cells(cell.Row, cell.Column).Value
        On Error GoTo ErrHandler

        If CStr(cell.Value) <> CStr(origVal) Then
            changes = changes + 1

            If Not cell.Comment Is Nothing Then cell.Comment.Delete

            If addComments Then
                Set cmt = cell.AddComment
                cmt.Text Text:="CDM Track Changes:" & vbCrLf & _
                    "??: " & IIf(IsEmpty(origVal), "(?)", CStr(origVal)) & vbCrLf & _
                    "??: " & IIf(IsEmpty(cell.Value), "(?)", CStr(cell.Value))
                cmt.Shape.TextFrame.AutoSize = True
            End If

            If highlightChanges Then
                cell.Interior.Color = RGB(146, 208, 80)
                cell.Borders.Color = RGB(255, 0, 0)
                cell.Borders.Weight = xlMedium
            End If
        End If
    Next cell

    origWb.Close SaveChanges:=False

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_TrackChanges = changes
    Exit Function

ErrHandler:
    On Error Resume Next
    origWb.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    DM_TrackChanges = -1
    MsgBox "DM_TrackChanges ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Sub DM_ClearTrackChanges(ByVal ws As Worksheet)
    On Error GoTo ErrHandler

    Dim cell As Range

    Application.ScreenUpdating = False

    For Each cell In ws.UsedRange
        If Not cell.Comment Is Nothing Then cell.Comment.Delete
        cell.Interior.Pattern = xlNone
        cell.Borders.LineStyle = xlNone
    Next cell

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_ClearTrackChanges ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_ChangesReport(ByVal ws As Worksheet, ByVal originalFilePath As String)

    On Error GoTo ErrHandler

    Dim origWb As Workbook
    Dim origWs As Worksheet
    Dim report As Worksheet
    Dim cell As Range
    Dim origVal As Variant
    Dim currVal As Variant
    Dim reportRow As Long
    Dim changeType As String
    Dim sheetName As String

    sheetName = ws.Name

    If Dir(originalFilePath) = "" Then
        MsgBox "???????: " & originalFilePath, vbExclamation, "CDM Toolkit"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    Set origWb = Workbooks.Open(originalFilePath, ReadOnly:=True)

    On Error Resume Next
    Set origWs = origWb.Worksheets(sheetName)
    If Err.Number <> 0 Then
        origWb.Close SaveChanges:=False
        Application.ScreenUpdating = True
        MsgBox "?????????sheet: " & sheetName, vbExclamation, "CDM Toolkit"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    On Error Resume Next
    ws.Parent.Worksheets("Changes Report").Delete
    On Error GoTo ErrHandler
    Set report = ws.Parent.Worksheets.Add(Before:=ws.Parent.Worksheets(1))
    report.Name = "Changes Report"

    report.Range("A1").Value = "????"
    report.Range("B1").Value = "????"
    report.Range("C1").Value = "??"
    report.Range("D1").Value = "??"
    report.Range("A1:D1").Font.Bold = True
    report.Range("A1:D1").Interior.Color = RGB(68, 114, 196)
    report.Range("A1:D1").Font.Color = vbWhite

    reportRow = 2

    For Each cell In ws.UsedRange
        On Error Resume Next
        origVal = origWs.Cells(cell.Row, cell.Column).Value
        On Error GoTo ErrHandler

        currVal = cell.Value

        If IsEmpty(origVal) And Not IsEmpty(currVal) Then
            changeType = "??"
        ElseIf Not IsEmpty(origVal) And IsEmpty(currVal) Then
            changeType = "??"
        ElseIf CStr(currVal) <> CStr(origVal) Then
            changeType = "??"
        Else
            GoTo SkipCell
        End If

        report.Cells(reportRow, 1).Value = cell.Address(False, False)
        report.Cells(reportRow, 2).Value = changeType
        report.Cells(reportRow, 3).Value = IIf(IsEmpty(origVal), "(?)", CStr(origVal))
        report.Cells(reportRow, 4).Value = IIf(IsEmpty(currVal), "(?)", CStr(currVal))
        reportRow = reportRow + 1

SkipCell:
    Next cell

    report.Columns("A:D").AutoFit

    origWb.Close SaveChanges:=False
    Application.ScreenUpdating = True

    MsgBox "???????????? " & (reportRow - 2) & " ????", _
        vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    On Error Resume Next
    origWb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "DM_ChangesReport ??: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
