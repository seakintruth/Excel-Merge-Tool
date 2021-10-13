Attribute VB_Name = "basExternalLoging"
Option Explicit
'Authored 2014-2021 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
    'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     'http://www.copyright.gov/title17/
     'In accrordance with 17 U.S.C. § 105 This work is 'noncopyright' or in the 'public domain'
         'Subject matter of copyright: United States Government works
         'protection under this title is not available for
         'any work of the United States Government, but the United States
         'Government is not precluded from receiving and holding copyrights
         'transferred to it by assignment, bequest, or otherwise.
     'as defined by 17 U.S.C § 101
         '...
         'A “work of the United States Government” is a work prepared by an
         'officer or employee of the United States Government as part of that
         'person’s official duties.
         '...
Private mfCancelLogging  As Boolean

Public Sub logStatus(strStatus As String, Optional fIgnoreLongDelay As Boolean = False)
    If Not mfCancelLogging Then
        Dim dblStartCheck As Double
        dblStartCheck = Evaluate("Now()") * (86400)
        If FolderExists(GetConfigurationFromDocumentComment("UseLogFolderPath")) Then
            If (Evaluate("Now()") * (86400)) - dblStartCheck > 1 _
            Then
                'Took longer than a second to check if the folder existed, stop all logging...
                mfCancelLogging = True
                If fIgnoreLongDelay Then
                   WriteToLog strStatus, GetLogPath
                End If
            Else
                WriteToLog strStatus, GetLogPath
            End If
        Else
            'Stop all logging attempts if we can't find the logging folder
            mfCancelLogging = True
        End If
    End If
End Sub
         
Public Sub WriteToLog( _
    Optional ByRef strContent As Variant = vbNullString, _
    Optional ByRef strFileName As String = "UserLog.txt", _
    Optional ByRef fRecordDescriptiveMachineInfo As Boolean = True, _
    Optional ByRef fOpenFile As Boolean = False, _
    Optional ByRef fOpenWithExplorerViceNotepad As Boolean = True, _
    Optional ByRef fAppendInsteadOfOverwriting As Boolean = True, _
    Optional ByRef fConvertToUnicode As Boolean = False _
    )
'Defaults are set for logging
strContent = CStr(strContent)
On Error GoTo HandleError
Dim fso As Object
Dim tf As Object
Dim objNetwork As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fConvertToUnicode Then
        strContent = StrConv(strContent, vbUnicode)
    End If
    'strFilename = mGetNamedAvailableFile(strFilename)
    If fso.FileExists(strFileName) Then
        If fAppendInsteadOfOverwriting Then
            Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
        Else
            fso.DeleteFile strFileName
            Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
            If fRecordDescriptiveMachineInfo Then
                tf.WriteLine "UserName,ComputerName,Version,Time,Path,Notes"
            Else
                tf.WriteLine "UserName,Time,Notes"
            End If
        End If
    Else
        Set tf = fso.OpenTextFile(strFileName, 8, True) ' 8 = ForAppending
        If fRecordDescriptiveMachineInfo Then
            If fRecordDescriptiveMachineInfo Then
                tf.WriteLine "UserName,ComputerName,Version,Time,Path,Notes"
            Else
                tf.WriteLine "UserName,Time,Notes"
            End If
        End If
    End If
    Dim strMsg As String
    If fRecordDescriptiveMachineInfo Then
        Set objNetwork = CreateObject("WScript.Network")
        'Precise time (millisecons) in excel using evaluate: https://stackoverflow.com/a/39019601, this is more precise than using cdbl(now())
        strMsg = objNetwork.UserName & ","
        strMsg = strMsg & objNetwork.ComputerName & ","
        strMsg = strMsg & HandleCsvColumn(GetConfigurationFromDocumentComment("VersionNumber")) & ","
        strMsg = strMsg & Evaluate("Now()") & ","
        strMsg = strMsg & HandleCsvColumn(ConvertMappedDrivePathToUNCPath(ThisWorkbook.path) & "\" & ThisWorkbook.Name) & ","
        strMsg = strMsg & HandleCsvColumn(Trim(strContent))
    Else
        strMsg = objNetwork.UserName & ","
        strMsg = strMsg & Evaluate("Now()") & ","
        strMsg = strMsg & HandleCsvColumn(Trim(strContent))
    End If
    tf.WriteLine strMsg
    #If DebugEnabled Then
        Debug.Print _
            Timer() & "," & _
            HandleCsvColumn(Trim(strContent)) & "," & _
            HandleCsvColumn(ThisWorkbook.Names("VersionNumber").RefersToRange.Value)
    #End If
    tf.Close
    'Clean up
    Set tf = Nothing
    Set fso = Nothing
    Set objNetwork = Nothing
    ' Open file
    If fOpenFile Then
        If fOpenWithExplorerViceNotepad Then
            OpenWithExplorer strFileName
        Else
            Shell "Notepad.exe " & strFileName, vbNormalFocus
        End If
    End If
    Exit Sub
HandleError:
    #If DebugEnabled Then
        Debug.Print Err.Number & ":" & Err.Description, vbExclamation, "basExternalLogging.SaveTextToFile", Err.HelpFile, Err.HelpContext
    #End If
    Resume Next
End Sub

Public Function FileAvailable(sFileName) As Boolean '
    'Modified from https://stackoverflow.com/a/22309698 and https://stackoverflow.com/a/9373914
    Dim ff As Long: ff = FreeFile()
    On Error Resume Next
    Open sFileName For Binary Access Read Lock Read As #ff
    ' or use 'Open FileName For Input Lock Read As #ff
    Close #ff
    Dim ErrNo As Long: ErrNo = Err.Number
    On Error GoTo 0
    Select Case ErrNo 'or use 'FileInUse = IIf(Err.Number > 0, True, False)
        Case 0: FileAvailable = True
        Case 52: FileAvailable = False  ' Bad File name or number
        Case 70: FileAvailable = False ' Permission Denied
        Case 75: FileAvailable = False 'Path\File access error
    Case Else: Error ErrNo
    FileAvailable = False
    End Select
End Function

Private Function mGetNamedAvailableFile(strFileName As String) As String
    Dim strFilePath As String
    Dim fAssignFromFileNameOnly As Boolean: fAssignFromFileNameOnly = False
    If InStrRev(strFileName, "\") <> 0 Then
        If FolderExists(Left(strFileName, InStrRev(strFileName, "\"))) Then
            If FileAvailable(ConvertMappedDrivePathToUNCPath(strFileName)) Then
                strFilePath = ConvertMappedDrivePathToUNCPath(strFileName)
            Else
                strFilePath = mGetThisWbOrTempPath(strFileName)
            End If
        Else
            strFileName = Right(strFileName, InStrRev(strFileName, "\"))
            fAssignFromFileNameOnly = True
        End If
    Else
        fAssignFromFileNameOnly = True
    End If
    If fAssignFromFileNameOnly Then
        If FolderExists(ThisWorkbook.Names("UsageLogFolder").RefersToRange.Value) Then
            If FileAvailable(ThisWorkbook.Names("UsageLogFolder").RefersToRange.Value & "\" & strFileName) Then
                strFilePath = ThisWorkbook.Names("UsageLogFolder").RefersToRange.Value & "\" & strFileName
            Else
                strFilePath = mGetThisWbOrTempPath(strFileName)
            End If
        Else
            strFilePath = mGetThisWbOrTempPath(strFileName)
        End If
    End If
    mGetNamedAvailableFile = strFilePath
End Function

Private Function mGetThisWbOrTempPath(strFileName As String) As String
Dim strFilePath As String
   If FolderExists(ThisWorkbook.path) Then
        strFilePath = ThisWorkbook.path & "\" & strFileName
    Else
        strFilePath = Environ("temp") & "\" & strFileName
    End If
    mGetThisWbOrTempPath = strFilePath
End Function

Private Sub OpenWithExplorer(ByRef strFilePath As String, Optional ByRef fReadOnly As Boolean = True)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    Set wshShell = Nothing
End Sub

Private Function HandleCsvColumn(ByVal strText As String)
    strText = IIf(Left(strText, 1) = "=", "`" & strText, strText)
    If Len(strText) > 0 Then
        HandleCsvColumn = """" & Replace(strText, """", """""") & """"
    End If
End Function


