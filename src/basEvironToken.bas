Attribute VB_Name = "basEvironToken"
Option Explicit

Public Function ReplaceNamedToken(ByVal strText, strTokenDelimeter)
    Dim arryTokenizedText
    arryTokenizedText = Split(strText, strTokenDelimeter)
    Dim intTextPart
    Dim strReturnText
    strReturnText = vbNullString
    For intTextPart = LBound(arryTokenizedText) To UBound(arryTokenizedText)
        'Every Even intTextPart should be the Token Keywords
        If intTextPart Mod 2 = 1 Then
            strReturnText = strReturnText & GetExpectedToken(arryTokenizedText(intTextPart))
        Else
            strReturnText = strReturnText & arryTokenizedText(intTextPart)
        End If
    Next
    ReplaceNamedToken = strReturnText
End Function

Public Function GetExpectedToken(strTokenKeyword)
    Select Case strTokenKeyword
        Case "~"
            GetExpectedToken = Environ("USERPROFILE")
        Case "DD"
            GetExpectedToken = Right("0" & Day(Now()), 2)
        Case "MM"
            GetExpectedToken = Right("0" & Month(Now()), 2)
        Case "YYYY"
            GetExpectedToken = Year(Now())
        Case "YY"
            GetExpectedToken = Right(Year(Now()), 2)
        Case "YYYYMMDD"
            GetExpectedToken = Year(Now()) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
        Case "YYMMDD"
            GetExpectedToken = Right(Year(Now()), 2) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
        Case "YYYY-MM-DD" ' same as :: GetExpectedToken("YYYY") & "-" & GetExpectedToken("MM") & "-" & GetExpectedToken("DD")
            GetExpectedToken = GetExpectedToken("YYYY") & "-" & GetExpectedToken("MM") & "-" & GetExpectedToken("DD")
        Case Else ' assume the token keyword is an environment variable
            GetExpectedToken = Environ(strTokenKeyword)
    End Select
End Function

Public Function HandleFileFolderPath(ByVal strPath As String) As String
    If Left(strPath, 2) = "~\" Then
        strPath = Environ("userprofile") & "\" & Right(strPath, Len(strPath) - 2)
    End If
    strPath = ReplaceNamedToken(strPath, "%")
    strPath = ConvertMappedDrivePathToUNCPath(strPath)
    'Assign
    HandleFileFolderPath = strPath
End Function

Public Function Environ(ByRef strName)
    If Left(strName, 1) = "%" And Right(strName, 1) = "%" Then
        strName = Mid(strName, 2, Len(strName) - 2)
    End If
    'Replaces VBA.Envrion Public Function with wscript version for use in all VB engines
    Dim wshShell: Set wshShell = CreateObject("WScript.Shell")
    Dim strResult: strResult = wshShell.ExpandEnvironmentStrings("%" & strName & "%")
    'wshShell.ExpandEnvironmentStrings behaves differently than VBA.Environ when no environment variable is found,
    '  conforming all results to return nothing if no result was found, like VBA.Environ
    If strResult = "%" & strName & "%" Then
        Environ = vbNullString
    Else
        Environ = strResult
    End If
    'cleanup
    Set wshShell = Nothing
End Function

