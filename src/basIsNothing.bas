Attribute VB_Name = "basIsNothing"
Option Explicit
'Authored 2014-2019 by Jeremy Dean Gerdes <jeremy.gerdes@navy.mil>
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
Public Function IsNothing(ByRef obj As Variant) As Boolean
On Error Resume Next
    If obj Is Nothing Then
        IsNothing = True
    Else
        IsNothing = False
    End If
End Function

Public Function ObjectHasValue(ByRef obj As Variant) As Boolean
    On Error Resume Next
    Dim strValue As Variant
    strValue = obj.Value
    ObjectHasValue = Err.Number = 0
End Function

Public Function IsSomething(ByRef obj As Variant) As Boolean
    IsSomething = Not IsNothing(obj)
End Function

Public Function IsNullEmptyMissing(ByRef varValue As Variant) As Boolean
    IsNullEmptyMissing = (IsNull(varValue) Or IsEmpty(varValue) Or IsMissing(varValue) Or varValue = "")
End Function

Public Function IsArrayAllocated(ByRef Arr As Variant) As Boolean
    On Error Resume Next
    ' Normally we only need to check LBound to determine if an array has been allocated.
    ' Some function such as Split will set LBound and UBound even if array is not allocated.
    ' See http://www.cpearson.com/excel/isarrayallocated.aspx for more details.
    IsArrayAllocated = IsArray(Arr) And _
        Not IsError(LBound(Arr, 1)) And _
        LBound(Arr, 1) <= UBound(Arr, 1)
End Function

Public Function IsActiveSheetPresent() As Boolean
On Error Resume Next
    Dim sht As Worksheet: Set sht = ActiveSheet
    If Err.Number = 0 Then
        IsActiveSheetPresent = True
    Else
        IsActiveSheetPresent = False
    End If
    
    'Clean up
    Set sht = Nothing
End Function






