Attribute VB_Name = "basReadExcelWithAdo"
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

Private Function mBuildAdoExcelConnectionString( _
    strExcelFilePath As String, _
    Optional fReturnHeader As Boolean = False, _
    Optional strPassword As String = vbNullString _
) As String
    If Len(strPassword) Then
        strPassword = "Password=" & strPassword & ";"
    End If
    Dim strTreatHeaderAsFieldNames As String
    If fReturnHeader Then
        strTreatHeaderAsFieldNames = "NO"
    Else
        strTreatHeaderAsFieldNames = "YES"
    End If
    'inspired by https://stackoverflow.com/a/1397181/1146659
    Dim strConnection As String
    'IMEX=1 treats all data as text, safest way to import data...
    Select Case UCase(GetFileExtension(strExcelFilePath))
        Case "XLSX"
            strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strExcelFilePath & ";Extended Properties=""Excel 12.0;HDR=" & strTreatHeaderAsFieldNames & ";IMEX=1"";"
        Case "XLSB"
            strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strExcelFilePath & ";Extended Properties=""Excel 12.0;HDR=" & strTreatHeaderAsFieldNames & ";IMEX=1"";" 'not sure if imex=1 works for xlsb ... ;IMEX=1"""""
        Case "XLSM"
            strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strExcelFilePath & ";Extended Properties=""Excel 12.0 Macro;HDR=" & strTreatHeaderAsFieldNames & ";IMEX=1"";" 'not sure if imex=1 works for xlsm ... ;IMEX=1"""""
        Case "XLS"
            strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strExcelFilePath & ";Extended Properties=""Excel 8.0;HDR=" & strTreatHeaderAsFieldNames & """;" 'not sure if imex=1 works for xls ... ;IMEX=1"""""
        Case "CSV", "ASC", "TAB", "TXT"
            'For a text file, Data Source is the folder, not the file. The file is the table (SELECT * FROM FileName.csv).
            strConnection = "Provider=Microsoft.ACE.OLEDB.16.0;Data Source=" & GetParentFolderName(strExcelFilePath) & ";Extended Properties='text'" & ";"   'not sure if imex=1 works for xls ... ;IMEX=1"""""
            '[TODO] optionally test these:
            'strConnection = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & GetParentFolderName(strExcelFilePath) & ";Extensions=asc,csv,tab,txt;"
            'strConnection = "Driver=Microsoft Access Text Driver (*.txt, *.csv);Dbq=" & GetParentFolderName(strExcelFilePath) & "; Extensions=asc,csv,tab,txt;"
            'strConnection = "Data Source='" & GetParentFolderName(strExcelFilePath) & "';Delimiter=',';Has Quotes=True;Skip Rows=0;Has Header=True;Comment Prefix='';Column Type=String,String,String,Int32,Boolean,String,String;Trim Spaces=False;Ignore Empty Lines=True;"
    End Select
    mBuildAdoExcelConnectionString = strConnection

End Function

Public Sub MakeOpenWorkbookConnection( _
    strExcelFilePath As String, _
    ByRef objConnection As ADODB.Connection, _
    Optional fReturnHeader As Boolean = False, _
    Optional strPassword As String = vbNullString _
)
    On Error Resume Next
    'passwords don't appear to ever help in the connection string (even though they should?)
    objConnection.Open mBuildAdoExcelConnectionString(strExcelFilePath, fReturnHeader, strPassword)
    If Err.Number <> 0 Then ' So allways try without the password
        Err.Clear
        objConnection.Open mBuildAdoExcelConnectionString(strExcelFilePath, fReturnHeader, vbNullString)
        Err.Clear
    End If
End Sub

Public Sub GetWorkbookSchemaRecordset( _
    strExcelFilePath As String, _
    ByRef objRecordset As ADODB.Recordset, _
    ByRef objConnection As ADODB.Connection _
)
    Dim rsTmp As New ADODB.Recordset
    Select Case UCase(GetFileExtension(strExcelFilePath))
        Case "XLSX", "XLSB", "XLSM", "XLS"
            On Error Resume Next
            Set objRecordset = objConnection.OpenSchema(adSchemaTables)
            On Error GoTo 0
            
        Case "CSV", "ASC", "TAB", "TXT"
            Dim strSql As String
            strSql = "SELECT * FROM [" & Right(strExcelFilePath, Len(strExcelFilePath) - (Len(GetParentFolderName(strExcelFilePath)) + 1)) & "]"
            rsTmp.Open strSql, objConnection
            Set objRecordset = rsTmp
    End Select
End Sub

Public Sub GetWorkbookRecordset( _
    strExcelFilePath As String, _
    ByRef objConnection As ADODB.Connection, _
    ByRef objRecordset As ADODB.Recordset, _
    Optional strWorkbookName As String = vbNullString _
)
    Dim rsTmp As New ADODB.Recordset
    Select Case UCase(GetFileExtension(strExcelFilePath))
        Case "XLSX", "XLSB", "XLSM", "XLS"
            rsTmp.Open "SELECT * FROM [" & strWorkbookName & "]", objConnection
            Set objRecordset = rsTmp
        Case "CSV", "ASC", "TAB", "TXT"
            Dim strSql As String
            strSql = "SELECT * FROM " & Right(strExcelFilePath, Len(strExcelFilePath) - (Len(GetParentFolderName(strExcelFilePath)) + 1))
            rsTmp.Open strSql, objConnection
            Set objRecordset = rsTmp
    End Select
End Sub

Private Function mBuildExcelCommandSql(strExcelFilePath As String) As String
Dim strSql As String
    Select Case UCase(GetFileExtension(strExcelFilePath))
        Case "XLSX", "XLSB", "XLSM", "XLS"
            Dim oRs  As ADODB.Recordset
            Dim objConnection As New ADODB.Connection
            MakeOpenWorkbookConnection strExcelFilePath, objConnection
            GetWorkbookSchemaRecordset strExcelFilePath, oRs, objConnection
            Do While Not oRs.EOF
                Dim fld As Object
                On Error Resume Next
                Set fld = oRs.Fields("table_name")
                If IsSomething(fld) Then
                    If Right(fld.Value, 1) = "$" Then ' Worksheets end with '$'
                        strSql = "SELECT * FROM [" & fld.Value & "]"
                        'Currently only using the first sheet found...
                        Exit Do
                    End If
                End If
                oRs.MoveNext
            Loop
        Case "CSV", "ASC", "TAB", "TXT"
            strSql = "SELECT * FROM [" & Right(strExcelFilePath, Len(strExcelFilePath) - (Len(GetParentFolderName(strExcelFilePath)) + 1)) & "]"
    End Select
    mBuildExcelCommandSql = strSql
End Function

Public Function LogWorkbookSchema( _
    strExcelFilePath As String, _
    strPassword As String, _
    Optional lngFileCount As Long, _
    Optional lngRecordCount As Long) _
As String
On Error GoTo HandleError
    Dim strStatus As String
    Dim strDetails As String
    Dim strResults As String
    
    Dim oRs  As ADODB.Recordset
    Dim objConnection As New ADODB.Connection
    MakeOpenWorkbookConnection strExcelFilePath, objConnection, strPassword:=strPassword
    GetWorkbookSchemaRecordset strExcelFilePath, oRs, objConnection
    strResults = vbNullString
    Do While Not oRs.EOF
        Dim fld As Object
        On Error Resume Next
        Set fld = oRs.Fields("table_name")
        On Error GoTo HandleError
        If IsSomething(fld) Then
            If Right(fld.Value, 1) = "$" Then ' Worksheets end with '$'
                strStatus = strStatus & """Worksheet Named:" & fld.Value
                strResults = strResults & """Worksheet Named:" & fld.Value & ""","
                strResults = strResults & vbCrLf
                Dim fldData As Object
                Dim rsData As ADODB.Recordset
                GetWorkbookRecordset strExcelFilePath, objConnection, rsData, fld.Value
                For Each fldData In rsData.Fields
                    strResults = strResults & """" & fldData.Name & ""","
                    strDetails = strDetails & """" & fldData.Name & ""","
                Next
                'remove trailing commas
                strDetails = Left(strDetails, Len(strDetails) - 1)
                strDetails = strDetails & vbCrLf
                strResults = Left(strResults, Len(strResults) - 1)
                strResults = strResults & vbCrLf
                
                'Get recordcount
                lngRecordCount = OpenAdoExcelRecordsetWithCount(strExcelFilePath, True, strPassword, rsData)
              Debug.Print rsData.Fields.Count
            End If
        Else
            strStatus = strStatus & "From text file type " & UCase(GetFileExtension(strExcelFilePath)) & vbCrLf
            For Each fld In oRs.Fields
                strResults = strResults & """" & fld.Name & ""","
                strDetails = strDetails & """" & fld.Name & ""","
            Next
            
            'remove trailing commas
            strResults = Left(strResults, Len(strResults) - 1)
            strDetails = Left(strDetails, Len(strDetails) - 1) & vbCrLf
                
            'Get recordcount
            lngRecordCount = OpenAdoExcelRecordsetWithCount(strExcelFilePath, True, strPassword, oRs)
            
            'For text files the recordset returned is all records, so we only need to inspect the fields once...
            Exit Do
        End If
        oRs.MoveNext
    Loop
    
    LogWorkbookSchema = Left(strResults, Len(strResults) - 1)
    
    'Remove vbCrLf
    strDetails = Left(strDetails, Len(strDetails) - 2)
    strStatus = Left(strStatus, Len(strStatus) - 2)
    
    LogStatusToRange strStatus, lngFileCount, lngRecordCount, strExcelFilePath, strDetails
    
ExitHere:
    'Cleanup
    objConnection.Close
    Set objConnection = Nothing
    Exit Function
HandleError:
    Debug.Print "basReadExcelWithADo::LogWorkbookShcema", Err.Number, Err.Description
    GoTo ExitHere
    'for debugging
    Resume
End Function

Public Function OpenAdoExcelRecordsetWithCount( _
    strExcelFilePath As String, _
    fReturnHeader As Boolean, _
    strPassword As String, _
    ByRef rsReturnData As ADODB.Recordset _
) As Long
On Error GoTo HandleError
    Dim connTemp  As ADODB.Connection
    Set connTemp = New ADODB.Connection
    connTemp.ConnectionString = mBuildAdoExcelConnectionString(strExcelFilePath, fReturnHeader, strPassword)
    connTemp.Open
    Dim rsTemp As ADODB.Recordset
    Set rsTemp = New ADODB.Recordset
    rsTemp.ActiveConnection = connTemp
    rsTemp.CursorType = adOpenDynamic
    rsTemp.LockType = adLockOptimistic
    Dim strSql As String
    strSql = mBuildExcelCommandSql(strExcelFilePath)
    rsTemp.Source = strSql
    'rsTemp.ActiveCommand.CommandText = strSql
    rsTemp.Open
    If Not rsTemp.EOF Then
        rsTemp.MoveLast
        rsTemp.MoveFirst
        'GetRows() row limit exists, may need to be handled via pagination
        On Error Resume Next
        'Get record count by upper bound of array's second dimention of results from GetRows()
        OpenAdoExcelRecordsetWithCount = UBound(rsTemp.GetRows(), 2) + 1
        If Err.Number = 7 Or Err.Number = -2174024882# Or Err.Number = -2147467259 Then
            If OpenAdoExcelRecordsetWithCount = 0 Then
                OpenAdoExcelRecordsetWithCount = -1
            End If
            Debug.Print Err.Number, Err.Description 'out of storage...
        End If
        On Error GoTo HandleError
        Set rsReturnData = Nothing
        Set rsReturnData = rsTemp
    End If
ExitHere:
    'Cleanup
    Exit Function
HandleError:
    Debug.Print "basReadExcellWithADo::OpenAdoExcelRecordsetWithCount", Err.Number, Err.Description
    GoTo ExitHere
    'for debugging
    Resume
End Function
