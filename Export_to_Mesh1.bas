Attribute VB_Name = "Export_to_Mesh1"
'This macro was written by SigmaRelief
'Version 1.3.2 2021-08-05
'https://github.com/SigmaRelief/SOLIDWORKS-Export-to-Mesh
'
'--- Modified version ---
'Changes from original (marked with 'FIX or 'NEW comments):
'  1. CreateFolder rewritten with FileSystemObject - works with UNC paths (\\server\share\...)
'  2. INI read buffer increased 128 -> 512 chars (long NAS paths were being truncated)
'  3. Safe boolean parsing from INI (empty/missing keys no longer cause Type Mismatch on first run)
'  4. Output coordinate system support via new INI key "CoordinateSystemName"
'     - applies to STL/3MF/AMF exports, restores the previous setting afterward
'  5. Overwrite guard: in Export-All-Configs mode, if the name template lacks [ConfigName],
'     "_[ConfigName]" is appended automatically so configs don't overwrite each other
'  6. Fixed bug where ExportPathFormat ending in "\" silently overwrote the user-edited export path
'  7. Single-config mode: array sized correctly, removed the "i = 1" loop hack
'  8. Case-insensitive .3mf extension check for the Python rename step
'  9. SaveAs result is checked; failures are reported with the offending path
' 10. (v2) Success verified by checking the file on disk, because SaveAs
'     returns False with lErrors=0 for some export formats even on success
' 11. (v2) A log file (ExportLog.txt) is appended in the export folder each
'     run, with settings used and per-file status
' 12. (v2) Python rename step guarded - a bad Python path logs a warning
'     instead of aborting the whole batch
' 13. (v3) Orientation via Python post-processing. The API preference
'     swFileSaveAsCoordinateSystem is rejected (returns False) on this
'     SolidWorks version, so instead the macro reads the coordinate system's
'     transform with GetCoordinateSystemTransformByName and passes it to
'     reorient_3mf.py, which bakes the rotation into the exported 3MF's
'     vertices. Requires: reorient_3mf.py in the macro's folder, and a valid
'     PythonPath setting (e.g. "py" or full path to python.exe).

Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Function ReadIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal IniFileName As String) As String
Dim Worked As Long
Dim RetStr As String * 512   'FIX: was 128 - long UNC/NAS paths were truncated
Dim StrSize As Long

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Read Not Specified !!!", vbExclamation, "INI"
  Else
    sProfileString = ""
    RetStr = Space(512)      'FIX: match buffer size
    StrSize = Len(RetStr)
    Worked = GetPrivateProfileString(Sect, Keyname, "", RetStr, StrSize, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Left$(RetStr, Worked)
    End If
  End If
  ReadIniFileString = sIniString
End Function

'NEW: safe boolean read - empty or missing keys return False instead of crashing
Private Function ReadIniBool(ByVal Sect As String, ByVal Keyname As String, ByVal IniFileName As String) As Boolean
Dim s As String
  s = LCase$(Trim$(ReadIniFileString(Sect, Keyname, IniFileName)))
  ReadIniBool = (s = "true" Or s = "1" Or s = "yes")
End Function

Private Function WriteIniFileString(ByVal Sect As String, ByVal Keyname As String, ByVal Wstr As String, ByVal IniFileName As String) As String
Dim Worked As Long

  iNoOfCharInIni = 0
  sIniString = ""
  If Sect = "" Or Keyname = "" Then
    MsgBox "Section Or Key To Write Not Specified !", vbExclamation, "INI"
  Else
    Worked = WritePrivateProfileString(Sect, Keyname, Wstr, IniFileName)
    If Worked Then
      iNoOfCharInIni = Worked
      sIniString = Wstr
    End If
    WriteIniFileString = sIniString
  End If
End Function

'FIX: rewritten with FileSystemObject so UNC paths (\\DS916\home\...) work.
'The old Split-on-"\" walk tried to MkDir "\\" and threw error 52.
Private Function CreateFolder(ByVal strFolderPath As String) As Boolean
Dim fso As Object
Dim strParent As String

  On Error GoTo CreateFolderFail
  Set fso = CreateObject("Scripting.FileSystemObject")

  strFolderPath = Trim$(strFolderPath)
  If Right$(strFolderPath, 1) = "\" Then strFolderPath = Left$(strFolderPath, Len(strFolderPath) - 1)

  If Len(strFolderPath) = 0 Then
    Err.Raise vbObjectError + 1, , "Export path is empty - check macro settings."
  End If
  If InStr(strFolderPath, "[") > 0 Then
    Err.Raise vbObjectError + 2, , "Unresolved [variable] in path: " & strFolderPath
  End If

  If fso.FolderExists(strFolderPath) Then
    CreateFolder = True
    Exit Function
  End If

  'Recursively ensure the parent exists, then create this level.
  'GetParentFolderName stops at the share root for UNC paths, so we never
  'attempt to create "\\" or "\\server" - those either exist or are unreachable.
  strParent = fso.GetParentFolderName(strFolderPath)
  If Len(strParent) > 0 Then
    If Not CreateFolder(strParent) Then Exit Function
  End If
  fso.CreateFolder strFolderPath
  CreateFolder = True
  Exit Function

CreateFolderFail:
  MsgBox "Could not create folder:" & vbCrLf & strFolderPath & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Export to Mesh"
  CreateFolder = False
End Function

'NEW v2: append a line to the run log. Fails silently (logging must never
'be the thing that kills an export run).
Private Sub WriteLog(ByVal LogFilePath As String, ByVal LogText As String)
Dim iFile As Integer
  On Error Resume Next
  iFile = FreeFile
  Open LogFilePath For Append As #iFile
  Print #iFile, Format(Now, "yyyy-mm-dd hh:nn:ss") & "  " & LogText
  Close #iFile
  On Error GoTo 0
End Sub

Public Function ReplaceVariables(ProcessString As String)

Dim FindKey As String
Dim FindLeft As Long
Dim FindRight As Long
Dim FindString As String
Dim ReplacementString As String
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim valOut As String
Dim wasResolved As Boolean

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

'Replace Date
FindLeft = InStr(ProcessString, "[Date ")
If FindLeft > 0 Then
    FindRight = InStr(FindLeft, ProcessString, "]")
    If FindLeft > 0 And FindRight > 0 Then
        FindString = Mid(ProcessString, FindLeft, FindRight - FindLeft + 1)
        FindKey = Mid(FindString, 7, FindRight - FindLeft - 6)
        ReplacementString = Format(Now, FindKey)
        ProcessString = Replace(ProcessString, FindString, ReplacementString)
    End If
End If

'Replace Custom Properties
FindLeft = InStr(ProcessString, "[Prop ")
While FindLeft > 0
    If FindLeft > 0 Then
        FindRight = InStr(FindLeft, ProcessString, "]")
        If FindLeft > 0 And FindRight > 0 Then
            Set swCustProp = swModel.Extension.CustomPropertyManager("")
            FindString = Mid(ProcessString, FindLeft, FindRight - FindLeft + 1)
            FindKey = Mid(FindString, 7, FindRight - FindLeft - 6)
            swCustProp.Get5 FindKey, True, valOut, ReplacementString, wasResolved
            ProcessString = Replace(ProcessString, FindString, ReplacementString)
        End If
        FindLeft = 0
    End If
Wend

'Replace Config Name
Set swConfig = swModel.GetActiveConfiguration
ProcessString = Replace(ProcessString, "[ConfigName]", swConfig.Name)

'Replace File Name
CurrentPathAndName = swModel.GetPathName

If CurrentPathAndName = "" Then 'Check if file has been saved
    CurrentName = "New File"
    Else
    CurrentPath = Left(CurrentPathAndName, InStrRev(swModel.GetPathName, "\") - 1)
    CurrentName = Strings.Mid(CurrentPathAndName, Strings.Len(CurrentPath) + 2, Strings.Len(CurrentPathAndName) - Strings.Len(CurrentPath) - 8) 'File name no extension
End If
ProcessString = Replace(ProcessString, "[FileName]", CurrentName)
ProcessString = Replace(ProcessString, "[FilePath]", CurrentPath)

'Replace User profile directory
ProcessString = Replace(ProcessString, "[UserDir]", Environ("USERPROFILE"))

ReplaceVariables = ProcessString

End Function

Sub main()

Dim FileNameProcessed As String
Dim ConfigNameArr() As String
Dim CurrentName As String
Dim CurrentPath As String
Dim CurrentPathAndName As String
Dim ExportAllConfigs As Boolean
Dim ExportConfigNames As String
Dim ExportConfigNamesProcessed As String
Dim ExportFilePathArr() As String
Dim ExportPath As String
Dim ExportPathFormat As String
Dim ExportPathProcessed As String
Dim FileCommandFormat As String
Dim FileExtension As String
Dim FileName As String
Dim FileNameFormat As String
Dim fso As Object
Dim lErrors As Long
Dim lWarnings As Long
Dim OpenFile As Boolean
Dim OpenFileCommand As String
Dim OpenTargetFolder As Boolean
Dim PythonArgs As String
Dim PythonCommand As String
Dim PythonExe As String
Dim PythonPath As String
Dim PythonScript As String
Dim Rename3MFObjects As Boolean
Dim retval As String
Dim SaveSettings As Boolean
Dim SettingsFile As String
Dim STLDeviation As Double
Dim STLDeviationInitial As Double
Dim STLAngleTolerance As Double
Dim STLAngleToleranceInitial As Double
Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swCustProp As CustomPropertyManager
Dim Version As String
Dim CoordSystemName As String            'NEW
Dim CoordSystemInitial As String         'NEW
Dim bSaveOK As Boolean                   'NEW
Dim FailedExports As String              'NEW
Dim LogFile As String                    'NEW v2
Dim fsoCheck As Object                   'NEW v2
Dim nOK As Long, nFail As Long           'NEW v2
Dim swXform As Object                    'NEW v3 (MathTransform)
Dim vXformData As Variant                'NEW v3
Dim CSArgs As String                     'NEW v3
Dim ReorientScript As String             'NEW v3
Dim ReorientCommand As String            'NEW v3
Dim wsh As Object                        'NEW v3 (WScript.Shell)
Dim exitCode As Long                     'NEW v3
Dim j As Long                            'NEW v3

Const Pi = 3.14159265358979

Set swApp = Application.SldWorks
Set swModel = swApp.ActiveDoc

'NEW: bail out politely if nothing is open
If swModel Is Nothing Then
    MsgBox "No document is open.", vbExclamation, "Export to Mesh"
    Exit Sub
End If

Version = "1.3.2-mod"

'Import settings from ini file
SettingsFile = Left(swApp.GetCurrentMacroPathName, Len(swApp.GetCurrentMacroPathName) - 3) & "ini"
FileNameFormat = ReadIniFileString("options", "FileNameFormat", SettingsFile)
ExportPathFormat = ReadIniFileString("options", "ExportPathFormat", SettingsFile)
PythonPath = ReadIniFileString("options", "PythonPath", SettingsFile)
FileExtension = ReadIniFileString("options", "FileExtension", SettingsFile)
Rename3MFObjects = ReadIniBool("options", "Rename3MFObjects", SettingsFile)   'FIX: safe bool
OpenFile = ReadIniBool("options", "OpenFile", SettingsFile)                   'FIX: safe bool
OpenFileCommand = ReadIniFileString("options", "OpenFileCommand", SettingsFile)
OpenTargetFolder = ReadIniBool("options", "OpenTargetFolder", SettingsFile)   'FIX: safe bool
ExportAllConfigs = ReadIniBool("options", "ExportAllConfigs", SettingsFile)   'FIX: safe bool
ExportConfigNames = ReadIniFileString("options", "ExportConfigNames", SettingsFile)
CoordSystemName = ReadIniFileString("options", "CoordinateSystemName", SettingsFile)  'NEW

ExportPath = ExportPathFormat
ExportPath = ReplaceVariables(ExportPath)
CurrentName = FileNameFormat
CurrentName = ReplaceVariables(CurrentName)

STLDeviationInitial = swApp.GetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swSTLDeviation)
STLAngleToleranceInitial = swApp.GetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swSTLAngleTolerance)

'Load values to userform
Load Export_Options
Load Variable_Inputs

Export_Options.FileNameBox.Value = CurrentName
Export_Options.FileExtensionCombo.Value = FileExtension
Export_Options.ExportPathBox.Value = ExportPath
Export_Options.STLDeviationBox.Value = STLDeviationInitial * 1000
Export_Options.STLAngleToleranceBox.Value = STLAngleToleranceInitial * 180 / Pi
Export_Options.Rename3MFObjectsBox.Value = Rename3MFObjects
Export_Options.PythonPathBox.Value = PythonPath
Export_Options.OpenFileBox.Value = OpenFile
Export_Options.OpenFileCommandBox.Value = OpenFileCommand
Export_Options.OpenTargetFolderBox.Value = OpenTargetFolder
Export_Options.ExportAllConfigsBox.Value = ExportAllConfigs
Export_Options.ExportConfigNamesCombo.Value = ExportConfigNames
Export_Options.SaveSettingsBox.Value = False
Export_Options.VersionLabel.Caption = "Version " & Version

Variable_Inputs.FileNameFormatBox.Value = FileNameFormat
Variable_Inputs.ExportPathFormatBox.Value = ExportPathFormat

If LCase(FileExtension) = ".3mf" Then    'FIX: case-insensitive
    Export_Options.Rename3MFObjectsBox.Visible = True
Else
    Export_Options.Rename3MFObjectsBox.Visible = False
End If

'Ask user for information
Export_Options.Show

FileName = Export_Options.FileNameBox.Value
FileExtension = Export_Options.FileExtensionCombo.Value
ExportPath = Export_Options.ExportPathBox.Value
STLDeviation = Export_Options.STLDeviationBox.Value / 1000
STLAngleTolerance = Export_Options.STLAngleToleranceBox.Value * Pi / 180
Rename3MFObjects = Export_Options.Rename3MFObjectsBox.Value
PythonPath = Export_Options.PythonPathBox.Value
OpenFile = Export_Options.OpenFileBox.Value
OpenFileCommand = Export_Options.OpenFileCommandBox.Value
OpenTargetFolder = Export_Options.OpenTargetFolderBox.Value
ExportAllConfigs = Export_Options.ExportAllConfigsBox.Value
ExportConfigNames = Export_Options.ExportConfigNamesCombo.Value
SaveSettings = Export_Options.SaveSettingsBox.Value

FileNameFormat = Variable_Inputs.FileNameFormatBox.Value
ExportPathFormat = Variable_Inputs.ExportPathFormatBox.Value

'Set mesh quality
If STLDeviation < STLDeviationInitial * 0.99 Or STLDeviation > STLDeviationInitial * 1.01 Then
    retval = swApp.SetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swSTLDeviation, STLDeviation)
End If

If STLAngleTolerance < STLAngleToleranceInitial * 0.99 Or STLAngleTolerance > STLAngleToleranceInitial * 1.01 Then
    retval = swApp.SetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swSTLAngleTolerance, STLAngleTolerance)
End If

'v3: the swFileSaveAsCoordinateSystem preference is rejected by the API on
'this SolidWorks version (Set returns False), so orientation is applied by
'the reorient_3mf.py post-processing step inside the export loop instead.
'CoordinateSystemName in the .ini still drives it - it names the coordinate
'system feature whose transform is read per configuration.

'Evaluate Output Path
If Right(ExportPath, 1) = "\" Then
    ExportPath = Left(ExportPath, Len(ExportPath) - 1)
End If

'FIX: original overwrote the user-edited ExportPath with the raw format string here.
'Now it only trims the trailing backslash from the format itself.
If Right(ExportPathFormat, 1) = "\" Then
    ExportPathFormat = Left(ExportPathFormat, Len(ExportPathFormat) - 1)
End If

'Python prep
If PythonPath = "python" Or PythonPath = "Python" Then
    PythonExe = PythonPath
Else
    PythonExe = Chr(34) & PythonPath & Chr(34)
End If

PythonScript = Chr(34) & Left(swApp.GetCurrentMacroPathName, Len(swApp.GetCurrentMacroPathName) - 3) & "py" & Chr(34)

'NEW v3: reorientation script lives next to the macro
ReorientScript = Left(swApp.GetCurrentMacroPathName, InStrRev(swApp.GetCurrentMacroPathName, "\")) & "reorient_3mf.py"
Set wsh = CreateObject("WScript.Shell")

'Create list of configurations to export
If ExportAllConfigs = True Then
    ConfigNameArr = swModel.GetConfigurationNames
    ReDim Preserve ExportFilePathArr(0 To UBound(ConfigNameArr))
    ExportConfigNamesProcessed = ExportConfigNames

    'NEW: overwrite guard - if the template has no [ConfigName], every config
    'would export to the same file and silently overwrite. Append it.
    If InStr(ExportConfigNamesProcessed, "[ConfigName]") = 0 Then
        ExportConfigNamesProcessed = ExportConfigNamesProcessed & "_[ConfigName]"
    End If
Else
    Set swConfig = swModel.GetActiveConfiguration
    ReDim ConfigNameArr(0)                     'FIX: was (1) - created a phantom empty element
    ConfigNameArr(0) = swConfig.Name
    ExportConfigNamesProcessed = "[FileName]"
    ReDim ExportFilePathArr(0 To 0)
End If

'Export Files from Solidworks
Set fsoCheck = CreateObject("Scripting.FileSystemObject")   'NEW v2

For i = 0 To UBound(ConfigNameArr)
    swModel.ShowConfiguration2 ConfigNameArr(i)
    retval = swModel.ForceRebuild3(False)

    'Create file path and file name
    ExportPathProcessed = ReplaceVariables(ExportPath)
    FileNameProcessed = Replace(ExportConfigNamesProcessed, "[FileName]", FileName)
    FileNameProcessed = ReplaceVariables(FileNameProcessed)
    ExportFilePathArr(i) = ExportPathProcessed & "\" & FileNameProcessed & FileExtension

    If Not CreateFolder(ExportPathProcessed) Then   'FIX: abort cleanly if folder can't be made
        GoTo Cleanup
    End If

    'NEW v2: start the run log on first iteration (needs the resolved path)
    If i = 0 Then
        LogFile = ExportPathProcessed & "\ExportLog.txt"
        WriteLog LogFile, "===== Export run started ====="
        WriteLog LogFile, "Model: " & swModel.GetPathName
        WriteLog LogFile, "Extension: " & FileExtension & "   Configs: " & (UBound(ConfigNameArr) + 1) & _
                          "   CoordSys: " & IIf(Len(CoordSystemName) > 0, CoordSystemName, "(default)") & _
                          "   Deviation: " & STLDeviation * 1000 & "   AngleTol: " & Format(STLAngleTolerance * 180 / Pi, "0.###")
    End If

    'NEW v2: delete any pre-existing target so disk-check can't see a stale file
    On Error Resume Next
    If fsoCheck.FileExists(ExportFilePathArr(i)) Then fsoCheck.DeleteFile ExportFilePathArr(i), True
    On Error GoTo 0

    bSaveOK = swModel.Extension.SaveAs(ExportFilePathArr(i), swSaveAsCurrentVersion, swSaveAsOptions_Silent, Nothing, lErrors, lWarnings)

    'FIX v2: SaveAs returns False with lErrors=0 for some export formats even
    'on success. The file on disk is the ground truth - trust that instead.
    If Not bSaveOK Then
        If fsoCheck.FileExists(ExportFilePathArr(i)) Then bSaveOK = True
    End If

    'NEW v3: read the coordinate system transform for THIS configuration
    '(translation can differ between configs if the CS is attached to geometry)
    CSArgs = ""
    If Len(CoordSystemName) > 0 And LCase(FileExtension) = ".3mf" Then
        Set swXform = swModel.Extension.GetCoordinateSystemTransformByName(CoordSystemName)
        If swXform Is Nothing Then
            WriteLog LogFile, "WARN  Coordinate system '" & CoordSystemName & "' not found in [" & ConfigNameArr(i) & "] - exporting in model space"
        Else
            vXformData = swXform.ArrayData
            For j = 0 To 11   '9 rotation + 3 translation values
                CSArgs = CSArgs & " " & Trim$(Str$(vXformData(j)))   'Str$ always uses "." decimals
            Next j
        End If
    End If

    If bSaveOK Then
        nOK = nOK + 1
        WriteLog LogFile, "OK    [" & ConfigNameArr(i) & "]  " & ExportFilePathArr(i)
    Else
        nFail = nFail + 1
        FailedExports = FailedExports & vbCrLf & ExportFilePathArr(i) & " (error " & lErrors & ")"
        WriteLog LogFile, "FAIL  [" & ConfigNameArr(i) & "]  " & ExportFilePathArr(i) & "  (error " & lErrors & ", warnings " & lWarnings & ")"
    End If

    'NEW v3: bake the print orientation into the 3MF (synchronous, so the
    'rename step and the next loop iteration can't race it)
    If bSaveOK And Len(CSArgs) > 0 Then
        ReorientCommand = PythonExe & " " & Chr(34) & ReorientScript & Chr(34) & " " & Chr(34) & ExportFilePathArr(i) & Chr(34) & CSArgs
        On Error Resume Next
        exitCode = wsh.Run(ReorientCommand, 0, True)   'hidden window, wait for completion
        If Err.Number <> 0 Then
            WriteLog LogFile, "WARN  Could not run reorient script (check PythonPath): " & ExportFilePathArr(i)
            Err.Clear
        ElseIf exitCode <> 0 Then
            WriteLog LogFile, "WARN  Reorient script failed (exit " & exitCode & ", see reorient_error.txt): " & ExportFilePathArr(i)
        Else
            WriteLog LogFile, "      reoriented via [" & CoordSystemName & "]"
        End If
        On Error GoTo 0
    End If

    'Run python script
    'FIX v2: guarded so a bad Python path logs a warning instead of aborting the batch
    If Rename3MFObjects = True And LCase(FileExtension) = ".3mf" And bSaveOK Then
        On Error Resume Next
        PythonArgs = " " & Chr(34) & ExportFilePathArr(i) & Chr(34)
        PythonCommand = PythonExe & " " & PythonScript & " " & PythonArgs
        Shell PythonCommand, vbNormalFocus
        If Err.Number <> 0 Then
            WriteLog LogFile, "WARN  Python rename failed (check Python path setting): " & ExportFilePathArr(i)
            Err.Clear
        End If
        On Error GoTo 0
    End If

Next i
'FIX: removed the "i = 1" single-config loop hack; array sizing handles it now

WriteLog LogFile, "===== Export run finished: " & nOK & " ok, " & nFail & " failed ====="

'Open Explorer Folder
If OpenTargetFolder = True Then
    Shell "explorer.exe " & Chr(34) & ExportPathProcessed & Chr(34), vbNormalFocus   'FIX: quoted for paths with spaces
End If

'Launch output file
If OpenFile = True Then
    FileCommandFormat = Mid(OpenFileCommand, InStr(OpenFileCommand, "{") + 1, InStr(OpenFileCommand, "}") - InStr(OpenFileCommand, "{") - 1)

    For i = 0 To UBound(ExportFilePathArr)
        OpenFileList = OpenFileList & Replace(FileCommandFormat, "[File]", ExportFilePathArr(i))
    Next i

    OpenFileCommand = Left(OpenFileCommand, InStr(OpenFileCommand, "{") - 1) & OpenFileList & Right(OpenFileCommand, Len(OpenFileCommand) - InStr(OpenFileCommand, "}"))
    Shell OpenFileCommand, vbNormalFocus
End If

'Save Settings
If SaveSettings = True Then

    retval = WriteIniFileString("options", "FileNameFormat", FileNameFormat, SettingsFile)
    retval = WriteIniFileString("options", "ExportPathFormat", ExportPathFormat, SettingsFile)
    retval = WriteIniFileString("options", "PythonPath", PythonPath, SettingsFile)
    retval = WriteIniFileString("options", "FileExtension", FileExtension, SettingsFile)
    retval = WriteIniFileString("options", "OpenTargetFolder", OpenTargetFolder, SettingsFile)
    retval = WriteIniFileString("options", "ExportAllConfigs", ExportAllConfigs, SettingsFile)
    retval = WriteIniFileString("options", "ExportConfigNames", ExportConfigNames, SettingsFile)
    retval = WriteIniFileString("options", "Rename3MFObjects", Rename3MFObjects, SettingsFile)
    retval = WriteIniFileString("options", "CoordinateSystemName", CoordSystemName, SettingsFile)   'NEW
End If

Cleanup:
'Reset mesh quality
If SaveSettings = False Then
    retval = swApp.SetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swSTLDeviation, STLDeviationInitial)
    retval = swApp.SetUserPreferenceDoubleValue(swUserPreferenceDoubleValue_e.swSTLAngleTolerance, STLAngleToleranceInitial)
End If

'NEW: report any failed exports
If Len(FailedExports) > 0 Then
    MsgBox "Some exports failed:" & vbCrLf & FailedExports, vbExclamation, "Export to Mesh"
End If

End Sub

