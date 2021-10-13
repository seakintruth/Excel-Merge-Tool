Attribute VB_Name = "funcMkDir"
Option Explicit

Public Function MkDir(ByRef strPath)
    ' Version: 1.0.3
    ' Dependancies: NONE
    ' Returns: True if no errors, i.e. folder path allready existed, or was able to be created without errors
    ' Usage Example: MkDir Environ("temp") & "\" & "opsRunner"
    ' Emulates linux 'MkDir -p' command:  creates folders without complaining if it allready exists
    ' Superceeds the the VBA.MkDir Public Function, but requires that drive be included in strPath
    ' By jeremy.gerdes@navy.mil
    Dim fso ' As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject") ' New Scripting.FileSystemObject
    If Not fso.FolderExists(strPath) Then
        On Error Resume Next
        Dim fRestore ' As Boolean
        fRestore = False
        'Handle Network Paths
        If Left(strPath, 2) = "\\" Then
            strPath = Right(strPath, Len(strPath) - 2)
            fRestore = True
        End If
        Dim arryPaths 'As Variant
        arryPaths = Split(strPath, "\")
        'Restore Server file path prefix
        If fRestore Then
            arryPaths(0) = "\\" & arryPaths(0)
        End If
        Dim intDir ' As Integer
        Dim strBuiltPath ' As String
        For intDir = LBound(arryPaths) To UBound(arryPaths)
            strBuiltPath = strBuiltPath & arryPaths(intDir) & "\"
            If Not fso.FolderExists(strBuiltPath) Then
                fso.CreateFolder strBuiltPath
            End If
        Next
    End If
    MkDir = (Err.Number = 0)
    'cleanup
    Set fso = Nothing
End Function


