Attribute VB_Name = "basNamedRange"
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

Public Function GetNamedRange( _
    strName As String, _
    Optional strWorkSheetName As String = vbNullString _
) As Range
    Dim sht As Worksheet
    Dim rngObject  As Range
    On Error Resume Next
    If Len(strWorkSheetName) = 0 Then
        Set rngObject = ThisWorkbook.Names(strName).RefersToRange
        Err.Clear
        If rngObject Is Nothing Then 'look in each workbook
            For Each sht In ThisWorkbook.Sheets
                Set rngObject = sht.Names(strName).RefersToRange
                If Not rngObject Is Nothing Then
                    Exit For
                End If
            Next
        End If
        On Error GoTo 0
    Else
        If WorksheetExistsByName(strWorkSheetName) Then
            Set sht = ThisWorkbook.Worksheets(TrimSingleQuote(strWorkSheetName))
            Set rngObject = sht.Names(strName).RefersToRange
        Else
            Set rngObject = Nothing
        End If
    End If
    'Return
    If Not rngObject Is Nothing Then
        Set GetNamedRange = rngObject
    Else
        Set GetNamedRange = Nothing
    End If
End Function

Public Function GetNamedRangeValue( _
    strName As String, _
    Optional strWorkSheetName As String = vbNullString _
) As Variant
    On Error Resume Next
    GetNamedRangeValue = GetNamedRange(strName, strWorkSheetName).Value
    If Err.Number <> 0 Then
        GetNamedRangeValue = vbNullString
    End If
End Function

Private Sub DoEventsOccasionally(Optional dblPercentage As Double = 0.1) ' 10% of the time doevents to prevent locking up this application
    If Rnd() > (1 - dblPercentage) Then
        DoEvents
    End If
End Sub





