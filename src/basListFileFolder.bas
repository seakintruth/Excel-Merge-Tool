Attribute VB_Name = "basListFileFolder"
Option Explicit
'Written by: Jeremy Dean Gerdes
'Norfolk Naval Shipyard
'C105 Health Physicist
'jeremy.gerdes@navy.mil
    'http://www.copyright.gov/title17/
    'In accordance with 17 U.S.C. § 105 This work is 'noncopyright'
        'Subject matter of copyright: United States Government works
        'Copyright protection under this title is not available for
        'any work of the United States Government, but the United States
        'Government is not precluded from receiving and holding copyrights
        'transferred to it by assignment, bequest, or otherwise.
        'as defined by 17 U.S.C § 101
        '...
        'A "work of the United States Government" is a work prepared by an
        'officer or employee of the United States Government as part of that
        'person's official duties.
        '...

Public Sub ListDirToLog( _
    strDirectory, _
    Optional fIncludeFiles As Boolean = False, _
    Optional nSubDirectoryDepth As Integer = 1 _
)
' convert all optional to required to move from VBA to VBS.
' some other minor function calls will need to be modified, like Environ for VBS...
' This is mostly a simplified example of subfolder recursion, there are much faster robocopy - list only methods tha this VBA to get a list of folders/files.
Const strcLogName = "ListFolderReport.txt"
Const ForWriting = 2
Const ForAppending = 8
    Dim strcSearchPath
    Dim Folder
    strcSearchPath = strDirectory
    If Len(strcSearchPath) = 0 Then
    MsgBox "ListFolders Canceled: You must enter a path"
        Exit Sub
    End If
    Dim fs
    Set fs = CreateObject("Scripting.FileSystemObject")

    If fs.FolderExists(strcSearchPath) Then
        Set Folder = fs.GetFolder(strcSearchPath)
    Else
        MsgBox "ListFolders Canceled: You must enter a valid path"
        Exit Sub
    End If

    Dim fExitLoop
    Dim strCurrentFolder

    ' Get Log Folder
    strCurrentFolder = Environ("TEMP")
    Dim strLogPath
    strLogPath = strCurrentFolder & "\" & strcLogName
    Dim filLog
    If Not fs.FileExists(strLogPath) Then
         Set filLog = fs.OpenTextFile(strLogPath, ForWriting, True)
         'Allways Overwrite
    Else
         Set filLog = fs.OpenTextFile(strLogPath, ForWriting, True)
    End If
    mCrawlFolder Folder, filLog, fIncludeFiles, nSubDirectoryDepth
    OpenFileWithExplorer strLogPath
    MsgBox "ListFolders Completed"
End Sub

Private Sub mCrawlFolder(Folder, filLog, fIncludeFiles, nSubDirectoryDepth) 'recursive routine
    DoEventsOccasionally 'prevents locking up excel
    Dim fExitLoop ' for in line error handling
    'If the File exist or we declare an overwrite (Listing Files or Folders)
    Do
        On Error Resume Next
        filLog.WriteLine Folder.path
        If Err.Number <> 0 Then
            Select Case MsgBox("The log file may be locked or you don't have permission to create the file" & vbCrLf & "If you can close the file do so and click Yes to try again, or click No to resume without logging or click Cancel to exit this script.", vbYesNoCancel)
                Case vbYes
                    'Do nothing as another attempt will be made.
                Case Else
                    fExitLoop = True
            End Select
        End If 'list folder file contents
    Loop Until (fExitLoop Or Err.Number = 0)
    On Error GoTo 0
    If fIncludeFiles And Not fExitLoop Then
        On Error Resume Next
        Do
            Dim fil As Object 'file
            For Each fil In Folder.Files
                If Err.Number = 70 Then
                    filLog.WriteLine "{File Access Denied}"
                Else
                    DoEventsOccasionally 'prevents locking up excel
                    filLog.WriteLine Folder.path & "\" & fil.Name
                    If Err.Number <> 0 Then
                        Select Case MsgBox("The log file may be locked or you don't have permission to create the file" & vbCrLf & "If you can close the file do so and click Yes to try again, or click No to resume without logging or click Cancel to exit this script.", vbYesNoCancel)
                            Case vbYes
                                'Do nothing as another attempt will be made.
                            Case Else
                                fExitLoop = True
                        End Select
                    End If 'list folder file contents
                End If
            Next
        Loop Until (fExitLoop Or Err.Number = 0)
    End If
    On Error GoTo 0
    If nSubDirectoryDepth > 0 And Not fExitLoop Then
        Dim fldr
        On Error Resume Next
        For Each fldr In Folder.SubFolders
            If Err.Number = 70 Then
                filLog.WriteLine "{SubFolder Access Denied}"
            Else
                mCrawlFolder fldr, filLog, fIncludeFiles, nSubDirectoryDepth - 1
            End If
        Next
        On Error GoTo 0
    End If
End Sub

Public Sub DoEventsOccasionally(Optional dblPercentage As Double = 0.01) ' 0.1 = 10% of the time doevents to prevent locking up this application
    If Rnd() > (1 - dblPercentage) Then
        DoEvents
    End If
End Sub

Private Sub OpenFileWithExplorer(strFilePath)
    Dim wshShell
    Set wshShell = CreateObject("WScript.Shell")
    wshShell.Exec ("Explorer.exe " & strFilePath)
    'Cleanup
    Set wshShell = Nothing
End Sub



