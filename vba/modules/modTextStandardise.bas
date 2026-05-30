Attribute VB_Name = "modTextStandardise"
Option Explicit

'============================================================================
' 模块: modTextStandardise
' 描述: CDM文本术语标准化 — 批量查找替换非标准术语为标准CDM术语
'       Text/Term Standardisation — batch replace non-standard terms
' 作者: CDM Toolkit Team
' 版本: 1.0 / 2026-05-30
' 依赖: 无（使用后期绑定 Scripting.Dictionary）
'============================================================================

Private mTermMap As Object

Private Sub DM_InitTermMap()
    If Not mTermMap Is Nothing Then Exit Sub
    Set mTermMap = CreateObject("Scripting.Dictionary")
    With mTermMap
        .Add "Recovered", "Recovered / Resolved"
        .Add "Ongoing", "Not Recovered / Not Resolved"
        .Add "Unknown", "Unknown"
        .Add "Reasonable Possib", "Reasonable Possibility"
        .Add "Not Recovered/Not Resolved", "Not Recovered / Not Resolved"
        .Add "Recovered/Resolved", "Recovered / Resolved"
        .Add "Recovered/Resolved with Sequelae", "Recovered / Resolved with Sequelae"
    End With
End Sub

Public Function DM_GetDefaultTermMap() As Object
    Call DM_InitTermMap
    Set DM_GetDefaultTermMap = mTermMap
End Function

Public Function DM_TextStandardise(ByVal ws As Worksheet, _
    Optional ByVal termMap As Object = Nothing) As Long
    On Error GoTo ErrHandler
    Dim cell As Range, key As Variant, count As Long, map As Object
    If termMap Is Nothing Then
        Call DM_InitTermMap: Set map = mTermMap
    Else
        Set map = termMap
    End If
    If map.Count = 0 Then DM_TextStandardise = 0: Exit Function
    Application.ScreenUpdating = False
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            For Each key In map.Keys
                If InStr(1, cell.Value, CStr(key), vbTextCompare) > 0 Then
                    cell.Value = Replace(cell.Value, CStr(key), CStr(map(key)), 1, -1, vbTextCompare)
                    count = count + 1
                End If
            Next key
        End If
    Next cell
    Application.ScreenUpdating = True
    DM_TextStandardise = count
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    DM_TextStandardise = -1
    MsgBox "DM_TextStandardise Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Function

Public Sub DM_TextReplace(ByVal ws As Worksheet, _
    ByVal oldText As String, ByVal newText As String, _
    Optional ByVal lookAt As XlLookAt = xlPart)
    On Error GoTo ErrHandler
    If Len(oldText) = 0 Then Exit Sub
    Application.ScreenUpdating = False
    ws.Cells.Replace What:=oldText, Replacement:=newText, _
        LookAt:=lookAt, SearchOrder:=xlByRows, MatchCase:=False, _
        SearchFormat:=False, ReplaceFormat:=False
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_TextReplace Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub
