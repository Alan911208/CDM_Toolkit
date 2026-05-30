Attribute VB_Name = "CDM_Install"
Option Explicit

' Module: CDM_Install
' Purpose: CDM Toolkit One-Click Install/Uninstall
'   - DM_InstallAll: Copy all 16 CDM modules to Personal.xlsb for global use
'   - DM_UninstallAll: Remove all CDM modules from Personal.xlsb
'   - DM_CreatePersonal: Guide to create Personal.xlsb if not exists
' Prerequisites:
'   1. Import all vba/modules/*.bas + CDM_Toolkit.bas + this file into a workbook
'   2. Run DM_InstallAll
'   3. Enable "Trust access to the VBA project object model" in Excel Trust Center
' Version: 1.1 / 2026-05-30

' v3.0: modSheetSplit/modWorkbookMerge/modSpecialChars/modFileListing migrated to Web UI
Private Const MODULE_NAMES As String = _
    "modTextStandardise,modHighlightDup,modBlankRows,modStrikethrough," & _
    "modTOC,modReturnToTOC," & _
    "modSheetManager,modSheetCreate," & _
    "modTrackChanges," & _
    "modDateCalc,modUnitConvert,modCreatinineClearance"

Public Sub DM_InstallAll()
    On Error GoTo ErrHandler

    Dim personalWb As Workbook
    Dim modules() As String
    Dim i As Long
    Dim comp As Object, srcComp As Object
    Dim installed As Long
    Dim failed As String

    modules = Split(MODULE_NAMES, ",")

    On Error Resume Next
    Set personalWb = Workbooks("PERSONAL.XLSB")
    If Err.Number <> 0 Then
        MsgBox "PERSONAL.XLSB not found." & vbCrLf & _
            "Create it first: View -> Macro -> Record Macro -> Store in Personal Macro Workbook", _
            vbInformation, "CDM Toolkit"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    For i = 0 To UBound(modules)
        Dim modName As String
        modName = Trim(modules(i))

        On Error Resume Next
        Set srcComp = ThisWorkbook.VBProject.VBComponents(modName)
        If Err.Number <> 0 Then
            failed = failed & modName & " (not found in this workbook)" & vbCrLf
            Err.Clear
            GoTo NextModule
        End If
        On Error GoTo ErrHandler

        On Error Resume Next
        personalWb.VBProject.VBComponents(modName).Activate
        If Err.Number = 0 Then
            personalWb.VBProject.VBComponents.Remove _
                personalWb.VBProject.VBComponents(modName)
        End If
        Err.Clear
        On Error GoTo ErrHandler

        Dim tempPath As String
        tempPath = Environ("TEMP") & "\" & modName & ".bas"
        srcComp.Export tempPath
        personalWb.VBProject.VBComponents.Import tempPath
        Kill tempPath

        installed = installed + 1

NextModule:
    Next i

    ' Also install master modules
    Dim masterMods As Variant
    masterMods = Array("CDM_Toolkit", "CDM_Install")
    For Each modName In masterMods
        On Error Resume Next
        Set srcComp = ThisWorkbook.VBProject.VBComponents(CStr(modName))
        If Err.Number = 0 Then
            personalWb.VBProject.VBComponents(CStr(modName)).Activate
            If Err.Number = 0 Then
                personalWb.VBProject.VBComponents.Remove _
                    personalWb.VBProject.VBComponents(CStr(modName))
            End If
            Err.Clear
            tempPath = Environ("TEMP") & "\" & modName & ".bas"
            srcComp.Export tempPath
            personalWb.VBProject.VBComponents.Import tempPath
            Kill tempPath
        End If
        On Error GoTo ErrHandler
    Next modName

    personalWb.Save
    Application.ScreenUpdating = True

    Dim msg As String
    msg = "CDM Toolkit v1.1 installed!" & vbCrLf & _
        "Modules installed: " & installed & " / " & (UBound(modules) + 1) & vbCrLf
    If Len(failed) > 0 Then
        msg = msg & vbCrLf & "Failed:" & vbCrLf & failed
    End If
    MsgBox msg, vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "DM_InstallAll Error: " & Err.Description & vbCrLf & vbCrLf & _
        "Enable: Excel -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings" & vbCrLf & _
        "Check: Trust access to the VBA project object model", _
        vbCritical, "CDM Toolkit"
End Sub

Public Sub DM_CreatePersonal()
    MsgBox "To create PERSONAL.XLSB:" & vbCrLf & vbCrLf & _
        "1. View -> Macros -> Record Macro" & vbCrLf & _
        "2. Set 'Store macro in' to 'Personal Macro Workbook'" & vbCrLf & _
        "3. Click OK, then immediately click Stop Recording" & vbCrLf & vbCrLf & _
        "After creation, run DM_InstallAll again.", _
        vbInformation, "CDM Toolkit"
End Sub

Public Sub DM_UninstallAll()
    On Error GoTo ErrHandler

    Dim personalWb As Workbook
    Dim modules() As String
    Dim i As Long
    Dim removed As Long

    On Error Resume Next
    Set personalWb = Workbooks("PERSONAL.XLSB")
    If Err.Number <> 0 Then
        MsgBox "PERSONAL.XLSB not found.", vbInformation, "CDM Toolkit"
        Exit Sub
    End If
    On Error GoTo ErrHandler

    modules = Split(MODULE_NAMES, ",")

    For i = 0 To UBound(modules)
        Dim modName As String
        modName = Trim(modules(i))

        On Error Resume Next
        personalWb.VBProject.VBComponents(modName).Activate
        If Err.Number = 0 Then
            personalWb.VBProject.VBComponents.Remove _
                personalWb.VBProject.VBComponents(modName)
            removed = removed + 1
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next i

    ' Remove master modules
    Dim masterMods As Variant
    masterMods = Array("CDM_Toolkit", "CDM_Install")
    For Each modName In masterMods
        On Error Resume Next
        personalWb.VBProject.VBComponents(CStr(modName)).Activate
        If Err.Number = 0 Then
            personalWb.VBProject.VBComponents.Remove _
                personalWb.VBProject.VBComponents(CStr(modName))
        End If
        Err.Clear
        On Error GoTo ErrHandler
    Next modName

    personalWb.Save
    MsgBox removed & " CDM modules removed.", vbInformation, "CDM Toolkit"
    Exit Sub

ErrHandler:
    MsgBox "DM_UninstallAll Error: " & Err.Description, vbCritical, "CDM Toolkit"
End Sub

'============================================================================
' DM_LaunchWeb
' Launch the local CDM Web Toolkit server and open browser
' v3.0 — Web UI for batch operations (merge, split, scan, files, dates, units)
'============================================================================
Public Sub DM_LaunchWeb()
    On Error GoTo ErrHandler

    Dim pythonCmd As String
    Dim serverScript As String

    ' Try to locate server.py relative to this workbook
    serverScript = ThisWorkbook.Path & "\server.py"

    If Dir(serverScript) = "" Then
        MsgBox "server.py not found. Please ensure the CDM Toolkit project" & vbCrLf & _
               "directory contains server.py and the cdm_engine/ package." & vbCrLf & vbCrLf & _
               "Expected location: " & serverScript, _
               vbExclamation, "CDM Toolkit"
        Exit Sub
    End If

    pythonCmd = "python """ & serverScript & """"
    Shell pythonCmd, vbNormalFocus

    MsgBox "CDM Web Toolkit is starting..." & vbCrLf & vbCrLf & _
           "Browser will open at: http://localhost:8520" & vbCrLf & _
           "Keep this terminal window open while using the Web tools.", _
           vbInformation, "CDM Toolkit v3.0"
    Exit Sub

ErrHandler:
    MsgBox "DM_LaunchWeb Error: " & Err.Description & vbCrLf & vbCrLf & _
           "Make sure Python 3.10+ is installed and accessible from the command line." & vbCrLf & _
           "Install: pip install fastapi uvicorn[standard] python-multipart", _
           vbCritical, "CDM Toolkit"
End Sub
