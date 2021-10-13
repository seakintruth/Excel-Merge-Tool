Attribute VB_Name = "basUi"
Option Explicit
'Authored 2021 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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

Public mfProcessing As Boolean

Public Sub SetCoverButtonVisibility(fDurringProcessing As Boolean)
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets("Cover")
    'Hide excecute buttons
    sht.Shapes("btnExecute").Visible = Not fDurringProcessing
    sht.Shapes("btnListSchemas").Visible = Not fDurringProcessing
    'Display Status
    sht.Shapes("rectangleStatus").Visible = fDurringProcessing
    sht.Shapes("btnCancel").Visible = fDurringProcessing
End Sub
        
Sub btnListSchemas_Click()
    On Error GoTo HandleError
    'Hide button (display's running message)
    SetEcho False
    mfProcessing = True
    SetCoverButtonVisibility mfProcessing
    'Execute
    MergeLikeExcelFiles fDryRun:=True
    
ExitHere:
    mfProcessing = False
    SetCoverButtonVisibility mfProcessing
    SetEcho True
    Exit Sub
HandleError:
    Debug.Print Err.Number, Err.Description
    GoTo ExitHere
End Sub
        
Sub btnExecute_Click()
    On Error GoTo HandleError
    'Hide button (display's running message)
    SetEcho False
    mfProcessing = True
    SetCoverButtonVisibility mfProcessing
    
    'Execute
    MergeLikeExcelFiles
    
ExitHere:
    mfProcessing = False
    SetCoverButtonVisibility mfProcessing
    SetEcho True
    Exit Sub
HandleError:
    Debug.Print Err.Number, Err.Description
    GoTo ExitHere
End Sub

Sub Picture3_Click()
    If Not mfProcessing Then
        Dim strFolder As String
        Dim strInitialPath As String
        strInitialPath = GetNamedRangeValue("SourceFolderPath")
        If Not FolderExists(strInitialPath) Then
            strInitialPath = vbNullString
        End If
        strFolder = GetFolderFromUser(strInitialPath)
        If Len(strFolder) > 0 Then
            GetNamedRange("SourceFolderPath").Value = strFolder
        End If
    End If
End Sub

Sub Picture4_Click()
    If Not mfProcessing Then
        Dim strFolder As String
        Dim strInitialPath As String
        strInitialPath = GetNamedRangeValue("ResultsFileName")
        If Not FolderExists(strInitialPath) And Not FolderExists(GetParentFolderName(strInitialPath)) Then
            strInitialPath = vbNullString
        End If
        strFolder = GetSaveAsFile(strInitialPath)
        If Len(strFolder) > 0 Then
            GetNamedRange("ResultsFileName").Value = strFolder
        End If
    End If
End Sub

Function GetFolderFromUser(Optional strDefaultFolder As String = vbNullString) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        If strDefaultFolder = vbNullString Then
            .InitialFileName = Application.DefaultFilePath
        Else
            .InitialFileName = strDefaultFolder & "\"
        End If
        If .Show <> -1 Then
            GoTo NextCode
        End If
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetFolderFromUser = sItem
    Set fldr = Nothing
End Function

Function GetSaveAsFile(Optional strDefaultFile As String = vbNullString) As String
    Dim fldr As FileDialog
    Dim sItem As String
    Set fldr = Application.FileDialog(msoFileDialogSaveAs)
    With fldr
        .Title = "Save Results As"
        
        If strDefaultFile = vbNullString Then
            .InitialFileName = Application.DefaultFilePath
        Else
            .InitialFileName = strDefaultFile
        End If
        If .Show <> -1 Then
            GoTo NextCode
        End If
        sItem = .SelectedItems(1)
    End With
NextCode:
    GetSaveAsFile = sItem
    Set fldr = Nothing
End Function

Private Sub mImportFolderFiles( _
    fDryRun As Boolean, _
    fFileNameAsCategory As Boolean, _
    fIncludeSubFolderDirectories As Boolean, _
    fFirstFile As Boolean, _
    lngSourceHeaderRowNumber As Long, _
    lngFileCount As Long, _
    lngTotalFileCount As Long, _
    lngDestinationRow As Long, _
    lngHeaderFirstColumn As Long, _
    lngHeaderLastColumn As Long, _
    lngRecordCountTotal As Long, _
    strSourceHeaderAddress As String, _
    strSourcePassword As String, _
    strCategoryColumnTitle As String, _
    arySourceFileExtensionFilter As Variant, _
    rngCategoryColumn As Range, _
    shtDestination As Worksheet, _
    wkbDestination As Workbook, _
    fldr As Object)
On Error GoTo HandleError
If Not basUi.mfProcessing Then
    GoTo ExitHere
End If
    If fIncludeSubFolderDirectories Then
        Dim fldrSub As Object
        For Each fldrSub In fldr.SubFolders
            mImportFolderFiles _
                fDryRun, _
                fFileNameAsCategory, _
                fIncludeSubFolderDirectories, _
                fFirstFile, _
                lngSourceHeaderRowNumber, _
                lngFileCount, _
                lngTotalFileCount, _
                lngDestinationRow, _
                lngHeaderFirstColumn, _
                lngHeaderLastColumn, _
                lngRecordCountTotal, _
                strSourceHeaderAddress, _
                strSourcePassword, _
                strCategoryColumnTitle, _
                arySourceFileExtensionFilter, _
                rngCategoryColumn, _
                shtDestination, _
                wkbDestination, _
                fldrSub
        Next
    End If
    Dim filSource As Object
    For Each filSource In fldr.Files
        Dim shtSource As Worksheet
        Dim lngRecordCount As Long
        Dim strCategory As String
                        
        If mIsFileExtentionTypeInFilter(filSource.Name, arySourceFileExtensionFilter) Then
            lngFileCount = lngFileCount + 1
            mSetShapeStatus "Status: Processing " & lngFileCount & " of " & lngTotalFileCount & " files"
            If fDryRun Then
                Dim strSchemaInfo As String
                strSchemaInfo = LogWorkbookSchema(filSource.path, strSourcePassword, lngFileCount, lngRecordCount)
                lngRecordCountTotal = lngRecordCountTotal + lngRecordCount
                'Debug.Print strSchemaInfo
            Else
                Dim rngDestinationColumns As Range
                Dim rngDestination As Range
                '[TODO] optionally export directly to .csv file with filesystem writes without limit, and modify ado methods
                '       to allow for paginated getrows() of 50,000 rows at a time to ensure we don't exceed the 1 MB limit
                Dim rsData As ADODB.Recordset
                Dim rngHeaderDestination As Range
                lngRecordCount = OpenAdoExcelRecordsetWithCount(filSource.path, False, strSourcePassword, rsData)
                lngRecordCountTotal = lngRecordCountTotal + lngRecordCount
                Dim fAdoQueryFailed As Boolean
                Dim strTmpCategory As String
                fAdoQueryFailed = False
                If lngRecordCount > 0 Then
                    lngDestinationRow = GetLastRow(GetNonBlankCellsFromWorksheet(shtDestination))
                    If lngDestinationRow = 0 Or fFirstFile Then
                        lngDestinationRow = 1
                        Set rngDestination = shtDestination.Rows(lngDestinationRow)
                        Set rngHeaderDestination = shtDestination.Range(strSourceHeaderAddress)
                        lngHeaderLastColumn = GetLastColumn(rngHeaderDestination)
                        lngHeaderFirstColumn = rngHeaderDestination.Columns(1).Column
                        Set rngDestinationColumns = shtDestination.Columns((lngHeaderFirstColumn)).Resize(, lngHeaderLastColumn).EntireColumn
                        Set rngDestination = Intersect(rngDestination, rngDestinationColumns)
                        Dim fldHeader As Object
                        Dim lngCurrentHeader As Long
                        'Insert the header row...
                        For Each fldHeader In rsData.Fields
                            lngCurrentHeader = lngCurrentHeader + 1
                            rngDestination.Cells(lngCurrentHeader).Value = fldHeader.Name
                        Next
                        Set rngCategoryColumn = rngDestination.Cells(rngDestination.Cells.Count).Offset(0, 1)
                        rngCategoryColumn.Value = strCategoryColumnTitle
                        fFirstFile = False
                    End If
                    lngDestinationRow = lngDestinationRow + 1
                    rsData.MoveFirst
                    Dim fCopyRecordsetFailed As Boolean
                    fCopyRecordsetFailed = False
                    Set rngDestination = shtDestination.Range(lngDestinationRow & ":" & lngDestinationRow)
                    Set rngDestinationColumns = shtDestination.Columns((lngHeaderFirstColumn)).EntireColumn
                    Set rngDestination = Intersect(rngDestination, rngDestinationColumns)
                    'rngDestination Should be a single cell
                    On Error Resume Next
                    rngDestination.CopyFromRecordset rsData
                    'Fill category values from file name
                    strTmpCategory = Right(filSource.path, Len(filSource.path) - Len(GetNamedRangeValue("SourceFolderPath")) - 1)
                    rngDestination.Offset(, lngHeaderLastColumn).Resize(lngRecordCount).Value2 = Left(strTmpCategory, InStrRev(strTmpCategory, ".") - 1)
                    If Err.Number = 0 Then
                        LogStatusToRange "Merged:CopyFromRecordset", lngFileCount, lngRecordCount, filSource.Name
                    Else
                        fCopyRecordsetFailed = True
                    End If
                    If fCopyRecordsetFailed Then
                        'reset fAdoQueryFailed
                        fAdoQueryFailed = False
                        Dim varData() As Variant
                        On Error Resume Next
                        varData = rsData.GetRows()
                        Set rngDestination = shtDestination.Range(lngDestinationRow & ":" & (lngDestinationRow + lngRecordCount - 1))
                        Set rngDestinationColumns = shtDestination.Columns((lngHeaderFirstColumn)).Resize(, UBound(varData, 1) + 1).EntireColumn
                        Set rngDestination = Intersect(rngDestination, rngDestinationColumns)
                        Dim varDataTransposed() As Variant
                        If Err.Number <> 0 Then
                            fAdoQueryFailed = True
                        End If
                        If (TransposeArray(varData, varDataTransposed)) Then
                            'Fill destination with source data
                            rngDestination = varDataTransposed
                            'Fill category values from file name
                            If fFileNameAsCategory Then
                                strTmpCategory = Right(filSource.path, Len(filSource.path) - Len(GetNamedRangeValue("SourceFolderPath")) - 1)
                                rngDestination.Columns(lngHeaderLastColumn).Offset(0, 1).Value2 = Left(strTmpCategory, InStrRev(strTmpCategory, ".") - 1)
                            Else
                                '[TODO] put the worksheet name instead of the filename here
                                strTmpCategory = Right(filSource.path, Len(filSource.path) - Len(GetNamedRangeValue("SourceFolderPath")) - 1)
                                rngDestination.Columns(lngHeaderLastColumn).Offset(0, 1).Value2 = Left(strTmpCategory, InStrRev(strTmpCategory, ".") - 1)
                            End If
                            If Err.Number = 0 Then
                                LogStatusToRange "Merged:GetRows() array", lngFileCount, lngRecordCount, filSource.Name
                            Else
                                fAdoQueryFailed = True
                            End If
                        Else
                            fAdoQueryFailed = True
                        End If
                    End If
                    On Error GoTo HandleError
                End If
                On Error GoTo HandleError
                If fAdoQueryFailed Then 'Open the work book if failed to use ADO to query the data
                    Dim wkbSource As Workbook
                    If Len(strSourcePassword) = 0 Then
                        Set wkbSource = ThisWorkbook.Application.Workbooks.Open( _
                            FileName:=filSource.path, _
                            UpdateLinks:=False, ReadOnly:=True, _
                            IgnoreReadOnlyRecommended:=True _
                        )
                    Else
                        On Error Resume Next
                        Set wkbSource = ThisWorkbook.Application.Workbooks.Open( _
                            FileName:=filSource.path, _
                            UpdateLinks:=False, _
                            Password:=strSourcePassword, ReadOnly:=True, _
                            WriteResPassword:=strSourcePassword, _
                            IgnoreReadOnlyRecommended:=True _
                        )
                        If Err.Number = 1004 Then ' password is not correct
                            On Error GoTo HandleError
                            Set wkbSource = ThisWorkbook.Application.Workbooks.Open( _
                                FileName:=filSource.path, _
                                UpdateLinks:=False, ReadOnly:=True, _
                                IgnoreReadOnlyRecommended:=True _
                            )
                        End If
                    End If
                    HideWorkbook wkbSource.Name
                    'mSetShapeStatus will call DoEvents to prevent locking up excel
                    Set shtSource = wkbSource.Sheets(1)
                    RevealSheet shtSource
                    'Excel.Application.DisplayAlerts = True
                    If fFileNameAsCategory Then
                        strTmpCategory = Right(filSource.path, Len(filSource.path) - Len(GetNamedRangeValue("SourceFolderPath")) - 1)
                        strCategory = Left(strTmpCategory, InStrRev(strTmpCategory, ".") - 1)
                    Else
                        strCategory = shtSource.Name
                    End If
                    shtSource.Cells.AutoFilter
                    Dim rngSourceUsedRange As Range
                    Set rngSourceUsedRange = shtSource.UsedRange
                    Dim rngSource As Range
                    Dim rngSourceColumns As Range
                    If fFirstFile Then
                        'Import the header row...
                        Set rngDestination = shtDestination.Rows(lngDestinationRow)
                        Set rngHeaderDestination = shtDestination.Range(strSourceHeaderAddress)
                        lngHeaderLastColumn = GetLastColumn(rngHeaderDestination)
                        lngHeaderFirstColumn = rngHeaderDestination.Columns(1).Column
                        Set rngDestinationColumns = shtDestination.Columns((lngHeaderFirstColumn)).Resize(, lngHeaderLastColumn).EntireColumn
                        Set rngDestination = Intersect(rngDestination, rngDestinationColumns)
                        Set rngSource = wkbSource.Sheets(1).Rows(lngSourceHeaderRowNumber)
                        Set rngSourceColumns = rngSource.Worksheet.Columns((lngHeaderFirstColumn)).Resize(, lngHeaderLastColumn).EntireColumn
                        Set rngSource = Intersect(rngSource, rngSourceColumns)
                        rngDestination.Value2 = rngSource.Value2
                        'rngSource.Copy Destination:=rngDestination
                        'rngDestination.PasteSpecial (XlPasteType.xlPasteFormats)
                        'Add the category title
                        Set rngCategoryColumn = rngDestination.Cells(rngDestination.Cells.Count).Offset(0, 1)
                        rngCategoryColumn.Value = strCategoryColumnTitle
                        'rngCategoryColumn.Offset(0, -1).Copy Destination:=rngCategoryColumn
                        fFirstFile = False
                    End If
                    Dim lngLastImportRow As Long
                    lngLastImportRow = GetLastRow(rngSourceUsedRange)
                    lngRecordCount = lngLastImportRow - lngSourceHeaderRowNumber
                    'lngRecordCount = lngRecordCount + lngLastImportRow - lngSourceHeaderRowNumber
                    lngRecordCountTotal = lngRecordCountTotal + lngRecordCount
                    Set rngSource = shtSource.Range(lngSourceHeaderRowNumber + 1 & ":" & lngLastImportRow)
                    Set rngSource = Intersect( _
                        rngSource, _
                        shtSource.Columns(lngHeaderFirstColumn).Resize(, lngHeaderLastColumn) _
                    )
                    lngDestinationRow = GetLastRow(GetNonBlankCellsFromWorksheet(shtDestination))
                    Set rngDestination = shtDestination.Range(lngDestinationRow + 1 & ":" & (lngDestinationRow + lngLastImportRow - 1))
                    Set rngDestinationColumns = shtDestination.Columns((lngHeaderFirstColumn)).Resize(, lngHeaderLastColumn).EntireColumn
                    Set rngDestination = Intersect(rngDestination, rngDestinationColumns)
                    rngDestination.Value2 = rngSource.Value2
                    'rngSource.Copy Destination:=rngDestination                 '.PasteSpecial XlPasteType.xlPasteFormats
                    'Add the category values (file name) to the last column
                    rngDestination.Columns(lngHeaderLastColumn).Offset(0, 1).Value2 = Left(filSource.Name, InStrRev(filSource.Name, ".") - 1)
                    'rngDestination.Columns(lngHeaderLastColumn).Copy Destination:=rngDestination.Columns(lngHeaderLastColumn).Offset(0, 1) '.PasteSpecial XlPasteType.xlPasteFormats
                    'de-select the cut copy mode
                    'wkbSource.Application.CutCopyMode = False
                    'Fake the save so we can close without prompt
                    wkbSource.Saved = True
                    wkbSource.Close False
                    'wkbDestination.Application.CutCopyMode = False
                    wkbDestination.Save 'saves the export data info
                    ThisWorkbook.Save 'saves the log info
                    LogStatusToRange "Merged:Open Workbook", lngFileCount, lngRecordCount, filSource.Name
                    'Clean as we go
                    Set wkbSource = Nothing
                End If
            End If 'End fDryRun/excecute check
            ' Check if user clicked the cancel button
            If Not basUi.mfProcessing Then
                GoTo ExitHere
            End If
        End If
    Next
ExitHere:
    
    Exit Sub
HandleError:
    Debug.Print "basUi::mImportFolderFiles", Err.Number, Err.Description
    'Cancel all execution
    basUi.mfProcessing = False
    GoTo ExitHere
    'for debugging
    Resume
End Sub

Private Function mGetTotalFileCount( _
    fIncludeSubfolders As Boolean, _
    fldr As Object, _
    arySourceFileExtensionFilter As Variant _
) As Long
Dim filSource As Object
Dim lngFileCount As Long
Dim fldrSub As Object
    If fIncludeSubfolders Then
        For Each fldrSub In fldr.SubFolders
           lngFileCount = lngFileCount + mGetTotalFileCount( _
                fIncludeSubfolders, _
                fldrSub, _
                arySourceFileExtensionFilter _
            )
        Next
    End If
    For Each filSource In fldr.Files
        If mIsFileExtentionTypeInFilter(filSource.Name, arySourceFileExtensionFilter) Then
            lngFileCount = lngFileCount + 1
        End If
    Next
    mGetTotalFileCount = lngFileCount
End Function

Public Sub MergeLikeExcelFiles( _
    Optional strFolderPath As String = vbNullString, _
    Optional fDryRun As Boolean = False _
)
On Error GoTo HandleError
    Dim rngLogHeader As Range
    'Check user input
    If Len(strFolderPath) = 0 Then
        strFolderPath = HandleFileFolderPath(GetNamedRangeValue("SourceFolderPath"))
        If Len(strFolderPath) = 0 Then
            SetEcho True, True
            MsgBox "Source folder path can not be empty", vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
            Exit Sub
        End If
    End If
    Dim strSourceHeaderAddress As String
    strSourceHeaderAddress = GetNamedRange("SourceHeaderAddress")
    If Len(strSourceHeaderAddress) = 0 Or Range(strSourceHeaderAddress) Is Nothing Then
        SetEcho True, True
        MsgBox "Source header address must be a valid address", vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
        Exit Sub
    End If
    Dim arySourceFileExtensionFilter As Variant
    arySourceFileExtensionFilter = Split(GetNamedRangeValue("SourceFileExtensionFilter"), ";")
    Dim lngSourceHeaderRowNumber As Long
    lngSourceHeaderRowNumber = CLng(Range(strSourceHeaderAddress).Rows(1).Row)
    If lngSourceHeaderRowNumber = 0 Then
        SetEcho True, True
        MsgBox "Source Header row number must be an integer > 0, and is defined by 'Source Header Address'", vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
        Exit Sub
    End If
    Dim strSourcePassword As String
    If GetNamedRangeValue("SourceFilesPassword") = True Then
        SetEcho True, True
        strSourcePassword = PasswordPrompt.GetPassword("Enter Password", "Enter Common Workbook Password")
        SetEcho False
    Else
        strSourcePassword = vbNullString
    End If
    Dim strResultsFileName As String
    strResultsFileName = HandleFileFolderPath(GetNamedRangeValue("ResultsFileName"))
    If Len(strResultsFileName) = 0 Then
        SetEcho True, True
        MsgBox "Results File Name can not be blank", vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
        Exit Sub
    ElseIf InStrRev(strResultsFileName, ".") = 0 Then
        strResultsFileName = strResultsFileName & ".xlsx"
    End If
    Dim fIncludeSubFolderDirectories As Boolean
    fIncludeSubFolderDirectories = GetNamedRangeValue("IncludeSubdirectories") = True
    Dim strCategoryColumnTitle As String
    strCategoryColumnTitle = GetNamedRangeValue("CategoryColumnTitle")
    Dim fFileNameAsCategory As Boolean
    fFileNameAsCategory = GetNamedRangeValue("UseFilenameAsCategory") = True
    'Split the results folder from the filename
    Dim strParentResultsFile As String
    strParentResultsFile = GetParentFolderName(strResultsFileName)
    Dim strDestinationPath  As String
    If Len(strParentResultsFile) > 0 Then
        If FolderExists(strParentResultsFile) Then
            strResultsFileName = Right(strResultsFileName, Len(strResultsFileName) - Len(strParentResultsFile) - 1)
            strResultsFileName = RemoveForbiddenFilenameCharacters(strResultsFileName)
        Else
            SetEcho True, True
            MsgBox "Destination path is not in a valid folder.", vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
            Exit Sub
        End If
        strDestinationPath = strParentResultsFile & "\" & strResultsFileName
    Else
        strParentResultsFile = strFolderPath
        If MkDir(strFolderPath & "\" & "Merged") Then
            strResultsFileName = RemoveForbiddenFilenameCharacters(strResultsFileName)
            strDestinationPath = strFolderPath & "\" & "Merged" & "\" & strResultsFileName
        Else
            SetEcho True, True
            MsgBox _
                "Failed to create destination folder:" & strFolderPath & "\" & "Merged" & "\" & _
                strResultsFileName, vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
            Exit Sub
        End If
    End If
    If Not fDryRun Then
        'Check to see if a workbook of the same name as the results file is allready open
        Dim wkbk As Workbook
        For Each wkbk In Excel.Workbooks
            If wkbk.Name = strResultsFileName Then
                GetNamedRange("ResultsFileName").Select
                SetEcho True, True
                Select Case _
                    MsgBox( _
                        "A work book with the same results file name is allready open. " & vbCrLf & vbCrLf & vbCrLf & _
                        "    Click 'Yes' to close that workbook (without saving),  " & vbCrLf & _
                        "    Click 'No' to change the new 'Results File Name' value " & vbCrLf & _
                        "    Click 'Cancel' to close this dialog and fix this yourself" & vbCrLf, _
                        vbYesNoCancel + vbQuestion, _
                         ThisWorkbook.Name & " -Input Error- " _
                    )
                    Case VbMsgBoxResult.vbYes
                        wkbk.Saved = True ' allows workbook to close without prompting for save
                        wkbk.Close
                    Case VbMsgBoxResult.vbNo
                        strResultsFileName = InputBox("Enter the new 'Results File Name'")
                        GetNamedRange("ResultsFileName").Value = strResultsFileName
                    Case Else
                        Exit Sub
                End Select
                SetEcho False
            End If
        Next
        'Create a new work book
        mSetShapeStatus "Status: Creating Destination Workbook"
        Dim wkbDestination As Object
        Set wkbDestination = Excel.Workbooks.Add
        If FileExists(strDestinationPath) Then
            DeleteFile strDestinationPath
        End If
        'Save the new results file
        wkbDestination.SaveAs (strDestinationPath)
        HideWorkbook wkbDestination.Name
        Dim shtDestination As Worksheet
        Set shtDestination = wkbDestination.Sheets(1)
        RevealSheet shtDestination
    End If
    Dim fs As Object: Set fs = CreateObject("Scripting.FileSystemObject")
    If fs.FolderExists(strFolderPath) Then
        Dim fldr As Object
        Set fldr = fs.GetFolder(strFolderPath)
        'Count the number of files that will be processed
        Dim lngTotalFileCount  As Long
        lngTotalFileCount = mGetTotalFileCount( _
            fIncludeSubFolderDirectories, _
            fldr, _
            arySourceFileExtensionFilter _
        )
        If lngTotalFileCount > 0 Then
            If fDryRun Then
                mSetShapeStatus "Status: Listing " & lngTotalFileCount & " files"
            Else
                mSetShapeStatus "Status: Processing " & lngTotalFileCount & " files"
            End If
            Dim fFirstFile As Boolean
            fFirstFile = True
            Dim lngFileCount As Long
            Dim lngDestinationRow As Long
            lngDestinationRow = 1
            Dim rngCategoryColumn As Range
            Dim lngHeaderFirstColumn As Long
            Dim lngHeaderLastColumn As Long
            Dim lngRecordCountTotal As Long
            
            mImportFolderFiles _
                fDryRun, _
                fFileNameAsCategory, _
                fIncludeSubFolderDirectories, _
                fFirstFile, _
                lngSourceHeaderRowNumber, _
                lngFileCount, _
                lngTotalFileCount, _
                lngDestinationRow, _
                lngHeaderFirstColumn, _
                lngHeaderLastColumn, _
                lngRecordCountTotal, _
                strSourceHeaderAddress, _
                strSourcePassword, _
                strCategoryColumnTitle, _
                arySourceFileExtensionFilter, _
                rngCategoryColumn, _
                shtDestination, _
                wkbDestination, _
                fldr
            If fDryRun Then
                RevealSheet ThisWorkbook.Worksheets("logDetails")
                If basUi.mfProcessing Then
                    LogStatusToRange "Dry Run Complete", lngFileCount, lngRecordCountTotal
                Else
                    LogStatusToRange "Dry Run Canceled", lngFileCount, lngRecordCountTotal
                End If
            Else
                RevealWorkbook wkbDestination.Name
                If basUi.mfProcessing Then
                    LogStatusToRange "Merge Complete", lngFileCount, lngRecordCountTotal
                Else
                    LogStatusToRange "Merge Canceled", lngFileCount, lngRecordCountTotal
                End If
            End If
            'Cleanup long log... (everything past 20,000 rows is deleted
            Set rngLogHeader = GetNamedRange("LogHeader")
            rngLogHeader.Offset(20000, 0).Resize(20000, 6).Clear
        Else
            LogStatusToRange "Complete: No files found to process", lngFileCount, lngRecordCountTotal
        End If
    Else
        SetEcho True, True
        MsgBox "MergeLikeExcelFiles Canceled: You must enter a valid 'Source Folder Path'", vbOKOnly + vbError, "Input Error:" & ThisWorkbook.Name
        Exit Sub
    End If
ExitHere:
    'Fix UI stuff
    SetEcho True
    ThisWorkbook.Activate
    Exit Sub
HandleError:
    Debug.Print Err.Number, Err.Description
    GoTo ExitHere
    Resume Next 'for debugging
    Resume
End Sub

Sub btnCancel_Click()
    basUi.mfProcessing = False
    SetCoverButtonVisibility basUi.mfProcessing
    SetEcho True
End Sub

Private Sub mSetShapeStatus(strMsg As String)
Dim shtExecuting As Worksheet
    Set shtExecuting = ThisWorkbook.Worksheets("Cover")
    Application.Windows(shtExecuting.Parent.Name).Visible = True
    shtExecuting.Activate
    SetEcho True, True
    DoEvents
    shtExecuting.Shapes("rectangleStatus").TextFrame2.TextRange.Characters.Text = strMsg
    SetEcho False
End Sub

Private Function mIsFileExtentionTypeInFilter(strFileName As String, aryFileExtention As Variant) As Boolean
    Dim intArrayElement As Integer
    If IsArrayAllocated(aryFileExtention) Then
        For intArrayElement = LBound(aryFileExtention) To UBound(aryFileExtention)
            If UCase(GetFileExtension(strFileName)) = UCase(Replace(aryFileExtention(intArrayElement), ".", "")) Then
                 mIsFileExtentionTypeInFilter = True
                 Exit Function
            End If
        Next
    End If
    'If we didn't find a match then set to false
    mIsFileExtentionTypeInFilter = False
End Function

Public Sub LogStatusToRange( _
    strStatus As String, _
    lngFileCount As Long, _
    lngRecordCount As Long, _
    Optional strFileName As String = vbNullString, _
    Optional strDetails As String = vbNullString _
)
Dim rngLogHeader As Range
    Set rngLogHeader = GetNamedRange("LogHeader")
    rngLogHeader.Offset(1, 0).Resize(, 7).Insert xlShiftDown
    rngLogHeader.Offset(1, 0).Value = strStatus
    rngLogHeader.Offset(1, 1).Value = strFileName
    rngLogHeader.Offset(1, 2).Value = lngFileCount
    rngLogHeader.Offset(1, 3).Value = lngRecordCount
    rngLogHeader.Offset(1, 4).Value = Now()
    rngLogHeader.Offset(1, 5).Value = Environ("USERNAME")
    rngLogHeader.Offset(1, 6).Value = strDetails
End Sub
