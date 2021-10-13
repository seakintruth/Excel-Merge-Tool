Attribute VB_Name = "basWorkBookActions"
Option Explicit
' Authored 2014-2021 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
    ' Public Domain in the United States of America,
     ' any international rights are waived through the CC0 1.0 Universal public domain dedication <https://creativecommons.org/publicdomain/zero/1.0/legalcode>
     ' http://www.copyright.gov/title17/
     ' In accrordance with 17 U.S.C. § 105 This work is ' noncopyright' or in the ' public domain'
         ' Subject matter of copyright: United States Government works
         ' protection under this title is not available for
         ' any work of the United States Government, but the United States
         ' Government is not precluded from receiving and holding copyrights
         ' transferred to it by assignment, bequest, or otherwise.
     ' as defined by 17 U.S.C § 101
         ' ...
         ' A “work of the United States Government” is a work prepared by an
         ' officer or employee of the United States Government as part of that
         ' person’s official duties.
         ' ...

Sub SaveandExit()
    ActiveWorkbook.Save
    ActiveWorkbook.Close
End Sub

Public Function WorksheetExistsByName(strName As String)
    'Function expects that names with spaces are wrapped in single quotes.
    On Error Resume Next
    Dim sht As Worksheet
    Set sht = Range(strName & "!A1").Worksheet
    'This method was failing, by selecting sheets that didn't match by name...
    '    strName = TrimSingleQuote(strName)
    '    Set sht = ThisWorkbook.Worksheets(strName)
    WorksheetExistsByName = Err.Number = 0
    Set sht = Nothing
    Err.Clear
    
End Function

Public Sub ActivateWorksheet(sht As Worksheet)
    On Error Resume Next
    sht.Select
    sht.Activate
    'sht.Range("A1").Select
    Err.Clear
    On Error GoTo 0
End Sub

Public Sub ActivateWorksheetByName(strSheetName As String)
    On Error Resume Next
'    Range(strSheetName & "!A1").Parent.Visible = XlSheetVisibility.xlSheetVisible
'    Range(strSheetName & "!A1").Parent.Select
'    Range(strSheetName & "!A1").Parent.Activate
'    Range(strSheetName & "!A1").Select
'    Err.Clear
    'This method was failing, by selecting sheets that didn't match by name...
    Dim sht As Worksheet
    Set sht = ThisWorkbook.Worksheets(strSheetName)
    If Not sht Is Nothing Then
        sht.Activate
        sht.Select
    End If
    'Cleanup
    Set sht = Nothing
    On Error GoTo 0
End Sub

Public Function GetLogPath() As String
    Dim strUseLogPath As String
    strUseLogPath = GetConfigurationFromDocumentComment("UseLogFolderPath")
    If Len(strUseLogPath) = 0 Then
        strUseLogPath = Environ("TEMP")
    End If
    GetLogPath = strUseLogPath & "\" & Left(ThisWorkbook.Name, InStrRev(ThisWorkbook.Name, ".") - 1) & ".log"
End Function

Public Function GetConfigurationFromDocumentComment(strConfigName As String) As String

    Dim arryDocumentComments As Variant
    arryDocumentComments = Split(ThisWorkbook.BuiltinDocumentProperties("Comments"), ";")
    Dim intDocCommentElement As Integer
    'Build a collection in another module if we add more configuration elements to the Document Comments
    For intDocCommentElement = LBound(arryDocumentComments) To UBound(arryDocumentComments)
        If Left(arryDocumentComments(intDocCommentElement), Len(strConfigName)) = strConfigName Then
            GetConfigurationFromDocumentComment = Split(arryDocumentComments(intDocCommentElement), "=")(1)
            Exit For
        End If
    Next
    If Len(GetConfigurationFromDocumentComment) = 0 Then
        Debug.Print "WARNING: This document's Comment Property should contain a semicolon delimited list including element " & strConfigName & " =...; Example=Return;"
    End If
End Function

Public Function TrimSingleQuote(strValue) As String
    If Left(strValue, 1) = "'" And Right(strValue, 1) = "'" Then
        TrimSingleQuote = Mid(strValue, 2, Len(strValue) - 2)
    Else
        TrimSingleQuote = strValue
    End If
End Function

Public Sub ToolUnhideAllWorkSheets(Optional fListOnly As Boolean = True)
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        If sht.Visible <> XlSheetVisibility.xlSheetVisible Then
            If Not fListOnly Then
                sht.Visible = XlSheetVisibility.xlSheetVisible
                ActivateWorksheet sht
            End If
            Debug.Print sht.Name
        End If
    Next
End Sub


