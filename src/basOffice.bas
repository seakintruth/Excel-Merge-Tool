Attribute VB_Name = "basOffice"
Option Explicit
'Authored 2014-2019 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
    'Public Domain in the United States of America,
     'any international rights are waived through the CC0 1.0 Universal public domain dedication
     '<https://creativecommons.org/publicdomain/zero/1.0/legalcode>
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
' 2011 By Bradley M. Gough modified for Excel by Jeremy Gerdes employees of the U.S. Govt
Public Type OfficeFileProperties
    Title As String
    Subject As String
    Author As String
    Category As String
    Keywords As String
    Comments As String
    RevisionNumber As String
    CreationDate As Date
    LastSaveTime As Date
End Type

Public Enum oatcOfficeApplictionType
    oatcWord = 1
    oatcExcel = 2
    oatcPowerPoint = 3
End Enum

Public Enum avcOfficeApplictionVersion
    oavc95 = 1
    oavc97 = 2
    oavc2000 = 3
    oavc2002 = 4
    oavc2003 = 5
    oavc2007 = 6
    oavc2010 = 7
End Enum

Public Enum oftcOfficeFileType
    oftcWord = 1
    oftcExcel = 2
    oftcPowerPoint = 3
End Enum

Public Enum ofrcOfficeFileReturn
    ofrcSuccess
    ofrcFailure
    ofrcFileTypeNotSupported
End Enum


Public Enum XlLookAt
    xlWhole
    xlPart
End Enum

Public Enum XlSearchOrder
    xlByRows
    xlByColumns
End Enum

Public Function GetOfficeApplicationClass(ByRef oatOfficeApplictionType As oatcOfficeApplictionType) As String

    Select Case oatOfficeApplictionType
        Case oatcWord
            GetOfficeApplicationClass = "Word.Application"
        Case oftcExcel
            GetOfficeApplicationClass = "Excel.Application"
        Case oftcPowerPoint
            GetOfficeApplicationClass = "PowerPoint.Application"
    End Select

End Function

Public Function GetExcelWkbOpenAsObject(strPath As String, Optional strSourcePassword As String = vbNullString, Optional fVisible As Boolean = True) As Object
          ' Get reference to Microsoft Excel application object
            Dim objApplication As Object
            Dim i As Long
            Dim fFileOpen As Boolean
            Dim wkb As Object
            Set objApplication = GetOfficeApplication(oatcExcel)
            ' Check if workbook already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Workbooks.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Workbooks(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate workbook
                    Set wkb = objApplication.Workbooks(i)
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                Set wkb = objApplication.Workbooks.Open( _
                    FileName:=strPath, _
                    UpdateLinks:=False, _
                    Password:=strSourcePassword, _
                    WriteResPassword:=strSourcePassword, _
                    IgnoreReadOnlyRecommended:=True _
                )
            End If
            On Error Resume Next
            objApplication.UserControl = fVisible
            objApplication.Visible = fVisible
            If fVisible Then
                objApplication.Activate
            End If
                        
            Set GetExcelWkbOpenAsObject = wkb
            
End Function

Public Function Nz(ByVal var As Variant, Optional ByVal strAlt As String) As Variant
    If IsNull(var) Then
        If Len(var & vbNullString) = 0 Then
            Nz = vbNullString
        Else
            Nz = strAlt
        End If
    Else
        Nz = var
    End If
End Function

Public Function GetOfficeApplication( _
    ByRef oatOfficeApplictionType As oatcOfficeApplictionType, _
    Optional ByRef fOfficeApplicationOpen As Boolean, _
    Optional fForceNewInstance _
) As Object
On Error Resume Next
    Dim strClass As String
    ' Get office application class
    strClass = GetOfficeApplicationClass(oatOfficeApplictionType)    ' Get instance of Microsoft Office Application Object (Application already open)
    ' This will only work if called from the same application
    If fForceNewInstance Then
        Select Case oatOfficeApplictionType
            Case oatcWord
                'Set GetOfficeApplication = New Word.Application
            Case oftcExcel
                Set GetOfficeApplication = New Excel.Application
            Case oftcPowerPoint
                'Set GetOfficeApplication = New PowerPoint.Application
        End Select
        If GetOfficeApplication Is Nothing Then
            fOfficeApplicationOpen = False
            Set GetOfficeApplication = CreateObject(Class:=strClass)
        End If
        If Err.Number <> 0 Then
            'Failed so get the application, so retry without attempting to force a new instance
            GetOfficeApplication oatOfficeApplictionType, fOfficeApplicationOpen, False
        End If
    Else
        Set GetOfficeApplication = GetObject(Class:=strClass)
        If GetOfficeApplication Is Nothing Then
            fOfficeApplicationOpen = False
            ' Create new instance of Microsoft Office Application Object (Application not already open)
            Set GetOfficeApplication = CreateObject(Class:=strClass)
        Else
            fOfficeApplicationOpen = True
        End If
    End If
End Function

Private Function GetOfficeFileType(ByRef strPath As String) As oftcOfficeFileType

    Select Case GetFileExtension(strPath)
        Case "doc", "docx", "docm", "dot", "dotx", "dotm"
            GetOfficeFileType = oftcWord
        Case "xls", "xlsx", "xlsm", "xlt", "xltx", "xltm", "xla", "xlam"
            GetOfficeFileType = oftcExcel
        Case "ppt", "pptx", "pptm", "pot", "potx", "potm", "pps", "ppsx", "ppsm", "ppa", "ppam", "sldx", "sldm"
            GetOfficeFileType = oftcPowerPoint
    End Select

End Function

Public Function OpenOfficeFile(ByRef strPath As String, Optional ByRef lngPage As Long, Optional ByRef strTag As String) As ofrcOfficeFileReturn

On Error Resume Next

Dim objApplication As Object
Dim odtOfficeFileType As oftcOfficeFileType
Dim i As Long
Dim fFileOpen As Boolean

    Err.Clear

    ' Return success by default
    OpenOfficeFile = ofrcSuccess

    ' Get type of office document
    odtOfficeFileType = GetOfficeFileType(strPath)

    ' Get office document properties using automation based
    ' on document type.
    Select Case odtOfficeFileType
        Case oftcWord
            ' Get reference to Microsoft Word application object
            Set objApplication = GetOfficeApplication(oatcWord)
            ' Check if document already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Documents.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Documents(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate document
                    objApplication.Documents(i).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                objApplication.Documents.Open FileName:=strPath
            End If
            ' Navigate to tag or page
            If LenB(strTag) <> 0 Then
                ' Select bookmark with name of strTag
                ' wdGoToBookmark = -1
                objApplication.Selection.Goto -1, , , strTag
            Else
                If lngPage <> 0 Then
                    ' Goto page lngPage
                    ' wdGoToPage = 1
                    ' wdGoToAbsolute = 1
                    objApplication.Selection.Goto 1, 1, lngPage
                End If
            End If
            ' Display application
            objApplication.UserControl = True
            objApplication.Visible = True
            objApplication.Activate
        Case oftcExcel
            ' Get reference to Microsoft Excel application object
            Set objApplication = GetOfficeApplication(oatcExcel)
            ' Check if workbook already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Workbooks.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Workbooks(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate workbook
                    objApplication.Workbooks(i).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                objApplication.Workbooks.Open FileName:=strPath
            End If
            ' Navigate to tag or page
            If LenB(strTag) <> 0 Then
                ' Select spreadsheet with name of strTag
                objApplication.ActiveWorkbook.Worksheets(strTag).Activate
                ' Select named range with name of strTag
                objApplication.Goto strTag, True
            Else
                If lngPage <> 0 Then
                    ' Select spreadsheet with index of lngPage
                    objApplication.ActiveWorkbook.Worksheets(lngPage).Activate
                End If
            End If
            ' Display application
            objApplication.UserControl = True
            objApplication.Visible = True
            ' Excel does not have an application activate method
            ' objApplication.Activate
        Case oftcPowerPoint
            ' Get reference to Microsoft PowerPoint application object
            Set objApplication = GetOfficeApplication(oatcPowerPoint)
            ' Display application
            ' Note: Must make PowerPoint visible or Presentations.Open will fail.  WTF
            objApplication.UserControl = True
            objApplication.Visible = True
            objApplication.Activate
            ' Check if presentation already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Presentations.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Presentations(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate presentation
                    ' PowerPoint does not have a presentation activate method
                    objApplication.Presentations(i).Windows(1).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open presentation
                objApplication.Presentations.Open FileName:=strPath
            End If
            ' Navigate to page
            If lngPage <> 0 Then
                ' Run presentation
                objApplication.ActivePresentation.SlideShowSettings.Run
                ' Select slide with slide number of lngPage
                objApplication.ActivePresentation.SlideShowWindow.View.GotoSlide lngPage
            End If
        Case Else
            OpenOfficeFile = ofrcFileTypeNotSupported
    End Select

    ' If any unexprected error occurs, indicated failure.
    If Err.Number <> 0 Then _
        OpenOfficeFile = ofrcFailure

    ' Clean up
    Set objApplication = Nothing

End Function

Public Function GetOfficeFileProperties(ByRef strPath As String, ByRef ofp As OfficeFileProperties) As ofrcOfficeFileReturn

On Error Resume Next

Dim objApplication As Object
Dim odtOfficeFileType As oftcOfficeFileType
Dim objProperties As Object
Dim fApplicationOpen As Boolean
Dim i As Long
Dim fFileOpen As Boolean

    ' Return success by default
    GetOfficeFileProperties = ofrcSuccess

    ' Get type of office document
    odtOfficeFileType = GetOfficeFileType(strPath)

    ' Get office file properties using automation based file type
    Select Case odtOfficeFileType
        Case oftcWord
            ' Get reference to Microsoft Word application object
            Set objApplication = GetOfficeApplication(oatcWord, fApplicationOpen)
            ' Check if document already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Documents.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Documents(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate document
                    objApplication.Documents(i).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                objApplication.Documents.Open FileName:=strPath, ReadOnly:=True, AddToRecentFiles:=False
            End If
            ' Get properties
            Set objProperties = objApplication.ActiveDocument.BuiltinDocumentProperties
        Case oftcExcel
            ' Get reference to Microsoft Excel application object
            Set objApplication = GetOfficeApplication(oatcExcel, fApplicationOpen)
            ' Check if workbook already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Workbooks.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Workbooks(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate workbook
                    objApplication.Workbooks(i).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                objApplication.Workbooks.Open FileName:=strPath, ReadOnly:=True, AddToMru:=False
            End If
            ' Get properties
            Set objProperties = objApplication.ActiveWorkbook.BuiltinDocumentProperties
        Case oftcPowerPoint
            ' Get reference to Microsoft PowerPoint application object
            Set objApplication = GetOfficeApplication(oatcPowerPoint, fApplicationOpen)
            ' Display application
            ' Note: Must make PowerPoint visible or Presentations.Open will fail.  WTF
            objApplication.Visible = True
            objApplication.Activate
            ' Check if presentation already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Presentations.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Presentations(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate presentation
                    ' PowerPoint does not have a presentation activate method
                    objApplication.Presentations(i).Windows(1).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open presentation
                objApplication.Presentations.Open FileName:=strPath, ReadOnly:=True
            End If
            ' Get properties
            Set objProperties = objApplication.ActivePresentation.BuiltinDocumentProperties
        Case Else
            GetOfficeFileProperties = ofrcFileTypeNotSupported
            GoTo ExitHere
    End Select

    ' If any unexprected error occurs, indicate failure
    ' Note: If we made it this far, we assume that we will be able to get the properties.
    ' We don't check after getting properties because we don't know if we will be able to
    ' get all the properties.
    If objProperties Is Nothing Then _
        GetOfficeFileProperties = ofrcFailure

    ofp.Title = objProperties("Title")
    ofp.Subject = objProperties("Subject")
    ofp.Author = objProperties("Author")
    ofp.Category = objProperties("Category")
    ofp.Keywords = objProperties("Keywords")
    ofp.Comments = objProperties("Comments")
    ofp.RevisionNumber = objProperties("Revision number")
    ofp.CreationDate = objProperties("Creation date")
    ofp.LastSaveTime = objProperties("Last save time")

    Set objProperties = Nothing

    ' Quit application (if it wasn't already open) and close file (if it wasn't already open)
    ' Note: We were only reading properties so we don't want to save changes.
    Select Case odtOfficeFileType
        Case oftcWord
            objApplication.ActiveDocument.Saved = True
            If Not fFileOpen Then _
                objApplication.ActiveDocument.Close
            If Not fApplicationOpen Then _
                objApplication.Quit
        Case oftcExcel
            objApplication.ActiveWorkbook.Saved = True
            If Not fFileOpen Then _
                objApplication.ActiveWorkbook.Close
            If Not fApplicationOpen Then _
                objApplication.Quit
        Case oftcPowerPoint
            objApplication.ActivePresentation.Saved = True
            If Not fFileOpen Then _
                objApplication.ActivePresentation.Close
            If Not fApplicationOpen Then _
                objApplication.Quit
    End Select

ExitHere:

    ' Clean up
    Set objApplication = Nothing
    Set objProperties = Nothing

End Function

Public Function SetOfficeFileProperties(ByRef strPath As String, ByRef ofp As OfficeFileProperties) As ofrcOfficeFileReturn

On Error Resume Next

Dim objApplication As Object
Dim odtOfficeFileType As oftcOfficeFileType
Dim objProperties As Object
Dim fApplicationOpen As Boolean
Dim i As Long
Dim fFileOpen As Boolean

    ' Return success by default
    SetOfficeFileProperties = ofrcSuccess

    ' Get type of office document
    odtOfficeFileType = GetOfficeFileType(strPath)

    ' Get office file properties using automation based file type
    Select Case odtOfficeFileType
        Case oftcWord
            ' Get reference to Microsoft Word application object
            Set objApplication = GetOfficeApplication(oatcWord, fApplicationOpen)
            ' Check if document already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Documents.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Documents(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate document
                    objApplication.Documents(i).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                objApplication.Documents.Open FileName:=strPath, AddToRecentFiles:=False
            End If
            ' Get properties
            Set objProperties = objApplication.ActiveDocument.BuiltinDocumentProperties
        Case oftcExcel
            ' Get reference to Microsoft Excel application object
            Set objApplication = GetOfficeApplication(oatcExcel, fApplicationOpen)
            ' Check if workbook already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Workbooks.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Workbooks(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate workbook
                    objApplication.Workbooks(i).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open document
                objApplication.Workbooks.Open FileName:=strPath, AddToMru:=False
            End If
            ' Get properties
            Set objProperties = objApplication.ActiveWorkbook.BuiltinDocumentProperties
        Case oftcPowerPoint
            ' Get reference to Microsoft PowerPoint application object
            Set objApplication = GetOfficeApplication(oatcPowerPoint, fApplicationOpen)
            ' Display application
            ' Note: Must make PowerPoint visible or Presentations.Open will fail.  WTF
            objApplication.Visible = True
            objApplication.Activate
            ' Check if presentation already open
            ' Note: FullName can return a path using mapped drives even if open path used UNC.  WTF.
            strPath = ConvertMappedDrivePathToUNCPath(strPath)
            For i = 1 To objApplication.Presentations.Count
                If StrComp(ConvertMappedDrivePathToUNCPath(UCase$(objApplication.Presentations(i).FullName)), UCase$(strPath), vbBinaryCompare) = 0 Then
                    ' Activate presentation
                    ' PowerPoint does not have a presentation activate method
                    objApplication.Presentations(i).Windows(1).Activate
                    fFileOpen = True
                    Exit For
                End If
            Next
            If Not fFileOpen Then
                ' Open presentation
                objApplication.Presentations.Open FileName:=strPath
            End If
            ' Get properties
            Set objProperties = objApplication.ActivePresentation.BuiltinDocumentProperties
        Case Else
            SetOfficeFileProperties = ofrcFileTypeNotSupported
            GoTo ExitHere
    End Select

    ' If any unexprected error occurs, indicate failure
    ' Note: If we made it this far, we assume that we will be able to set the properties.
    ' We don't check after setting properties because we don't know if we will be able to
    ' set all the properties.
    If objProperties Is Nothing Then _
        SetOfficeFileProperties = ofrcFailure

    objProperties("Title") = ofp.Title
    objProperties("Subject") = ofp.Subject
    objProperties("Author") = ofp.Author
    objProperties("Category") = ofp.Category
    objProperties("Keywords") = ofp.Keywords
    objProperties("Comments") = ofp.Comments
    objProperties("Revision number") = ofp.RevisionNumber
    objProperties("Creation date") = ofp.CreationDate
    objProperties("Last save time") = ofp.LastSaveTime

    ' Quit application (if it wasn't already open) and close file (if it wasn't already open)
    ' Note: We were updating properties so we want to save changes.
    Select Case odtOfficeFileType
        Case oftcWord
            ' Note: If Saved is not set to false, word will detect that no changes have
            ' been made and ignore the save method.  Excel and PowerPoint will save anyways
            ' but we include the same code for consistence and potential office updates.
            objApplication.ActiveDocument.Saved = False
            objApplication.ActiveDocument.Save
            If Not fFileOpen Then _
                objApplication.ActiveDocument.Close
            If Not fApplicationOpen Then _
                objApplication.Quit
        Case oftcExcel
            objApplication.ActiveWorkbook.Saved = False
            objApplication.ActiveWorkbook.Save
            If Not fFileOpen Then _
                objApplication.ActiveWorkbook.Close
            If Not fApplicationOpen Then _
                objApplication.Quit
        Case oftcPowerPoint
            objApplication.ActivePresentation.Saved = False
            objApplication.ActivePresentation.Save
            If Not fFileOpen Then _
                objApplication.ActivePresentation.Close
            If Not fApplicationOpen Then _
                objApplication.Quit
    End Select

ExitHere:

    ' Clean up
    Set objApplication = Nothing
    Set objProperties = Nothing

End Function

Public Function IsOfficeApplicationObjectVariableValid(ByRef objApplication As Object) As Boolean
On Error Resume Next
Dim strName As String
    Err.Clear
    strName = objApplication.Name
    IsOfficeApplicationObjectVariableValid = (Err.Number = 0)
End Function

Public Sub QuitWordApplicationIfNotVisible(ByRef objApplication As Object)
On Error Resume Next
Dim i As Long
    ' If office application if not visible.
    ' Used by a clean up procedure to prevent a hidden instance of an office
    ' application from being open in the backgroud following an unexpected error.
    If Not objApplication Is Nothing Then
        If Not objApplication.Visible Then
              ' Ensure display is updated before quiting application
            DoEvents
            ' Quit application
            objApplication.Quit False
            ' Wait for application to quit up to 10 times.
            ' If it take more then 10 tiems we assume word
            ' hung and we try to carry on so this application
            ' does not also hang.  Testing indicates it normally
            ' only takes one time.
            For i = 1 To 10
                If IsOfficeApplicationObjectVariableValid(objApplication) Then
                    DoEvents
                Else
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Public Sub QuitExcelApplicationIfNotVisible(ByRef objApplication As Object)
On Error Resume Next
Dim i As Long
    ' If office application if not visible.
    ' Used by a clean up procedure to prevent a hidden instance of an office
    ' application from being open in the backgroud following an unexpected error.
    If Not objApplication Is Nothing Then
        If Not objApplication.Visible Then
            ' Ensure display is updated before quiting application
            DoEvents
            ' Quit application
            objApplication.DisplayAlerts = False
            objApplication.Quit
            ' Wait for application to quit up to 10 times.
            ' If it take more then 10 tiems we assume word
            ' hung and we try to carry on so this application
            ' does not also hang.  Testing indicates it normally
            ' only takes one time.
            For i = 1 To 10
                If IsOfficeApplicationObjectVariableValid(objApplication) Then
                    DoEvents
                Else
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Public Function GetWordTemplateDocument(Optional ByRef strTemplateFilePath As String, Optional ByRef fApplicationVisible As Boolean = True) As Object ' Word.Document

On Error GoTo HandleError

Dim objApplication As Object ' Word.Application
Dim objDocument As Object ' Word.Document
Dim strFilter As String

    ' Get template file path
'    If LenB(strTemplateFilePath) = 0 Then
'        strFilter = "Word Template (*.dot; *.dotx; *.dotm; *.doc; *.docx; *.docm)" & vbNullChar & "*.dot;*.dotx;*.dotm;*.doc;*.docx;*.docm" & vbNullChar
'        strFilter = strFilter & "All Files (*.*)" & vbNullChar & "*.*" & vbNullChar
'        strFilter = strFilter & vbNullChar & vbNullChar
'        strTemplateFilePath = FileDialog(strFilter:=strFilter, strDialogTitle:="Select Word Template")
'    End If

    If LenB(strTemplateFilePath) <> 0 Then

        ' Note*: Testing indicates the various parts of the below code
        ' sometimes trigger Out of Memory Error but works anyways.  WTF.
        On Error Resume Next

        ' Open Word File based on template if strTemplateFilePath passed
        Set objApplication = GetOfficeApplication(oatcWord)
        Set objDocument = objApplication.Documents.Add(Template:=strTemplateFilePath, NewTemplate:=False, DocumentType:=0)

        ' Show Word
        If fApplicationVisible Then
            objApplication.UserControl = True
            objApplication.Visible = True
            objApplication.Activate
        End If

        On Error GoTo HandleError

        ' Verify Microsoft Word template document was opened
        If objDocument Is Nothing Then _
            Err.Raise Number:=vbObjectError + 1, Description:="Failed to open Microsoft Word template document."

        ' Return word document object
        Set GetWordTemplateDocument = objDocument

    End If

ExitHere:
    ' Clean up
    Set objApplication = Nothing
    Set objDocument = Nothing
    Exit Function

HandleError:
    Select Case Err.Number
        Case Else
        MsgBox "basOffice::GetWordTemplateDocument" & Err.Description, vbCritical + vbOKOnly, "Barcode Generator" & Err.Number
'            Select Case hecExitProcedure ' HandleError("basOffice::GetWordTemplateDocument")
'                Case hecResume
'                    Resume
'                Case hecResumeNext
'                    Resume Next
'                Case hecExitProcedure
'                    Resume ExitHere
'            End Select
    End Select

End Function

' Used to read bookmark text without triggering an error if bookmark does not exist
Public Function GetWordDocumentBookmark(ByRef objDocument As Object, ByRef strName As String, ByRef strText As String, Optional ByRef fDelete As Boolean = False) As Boolean

On Error Resume Next

Dim objRange As Object

    ' Note*: Normal bookmarks cannot be longer then 40 characters
    ' and formfield bookmarks cannot be longer then 20 characters.
    ' FormField Bookmark
    Err.Clear
    strText = objDocument.Bookmarks(strName).Range.Fields(1).Result.Text
    If Err.Number <> 0 Then
        ' Normal Bookmark
        Err.Clear
        Set objRange = objDocument.Bookmarks(strName).Range
        strText = objRange.Text
        If fDelete Then _
            objDocument.Bookmarks.Item(strName).Delete
    End If

    GetWordDocumentBookmark = Err.Number = 0

    ' Clean up
    Set objRange = Nothing

End Function

' Used to set bookmark text without triggering an error if bookmark does not exist
Public Function SetWordDocumentBookmark(ByRef objDocument As Object, ByRef strName As String, ByRef strText As String, Optional ByRef fDelete As Boolean = False) As Boolean

On Error Resume Next

Dim objRange As Object ' Word.Range

    ' Note*: Normal bookmarks cannot be longer then 40 characters
    ' and formfield bookmarks cannot be longer then 20 characters.
    ' FormField Bookmark
    Err.Clear
    objDocument.Bookmarks(strName).Range.Fields(1).Result.Text = strText
    If Err.Number <> 0 Then
        ' Normal Bookmark
        Err.Clear
        Set objRange = objDocument.Bookmarks(strName).Range
        objRange.Text = strText
        If fDelete Then
            objDocument.Bookmarks.Item(strName).Delete
        Else
            ' Update bookmark range to include added text
            objDocument.Bookmarks.Add Name:=strName, Range:=objRange
        End If
    End If

    SetWordDocumentBookmark = Err.Number = 0

    ' Clean up
    Set objRange = Nothing

End Function

Public Sub RestartWordListParagraphNumbering(ByRef objDocument As Object, ByRef lngListParagraphNumber As Long)
On Error Resume Next
    ' Note*: I attempted to use wdListApplyToThisPointForward instead of wdListApplyToWholeList
    ' but wdListApplyToWholeList must be used for numbering to be correctly restarted.
    Const wdListApplyToWholeList As Variant = 0
    Const wdWord9ListBehavior As Variant = 1
    objDocument.ListParagraphs(lngListParagraphNumber).Range.ListFormat.ApplyListTemplate objDocument.ListParagraphs(lngListParagraphNumber).Range.ListFormat.ListTemplate, False, wdListApplyToWholeList, wdWord9ListBehavior
End Sub



