Attribute VB_Name = "basAppOptions"
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
'Public Variables
Public gfOriginalAppOptionsAreCurrent As Boolean
Public gxlEnableCancelKey As XlEnableCancelKey

'Define App Option variables in this modual to restore later
Dim mfCalculateBeforeSave As Boolean
Dim mfEnableAutoRecover As Boolean
Dim mfEnableLargeOperationAlert As Boolean
Dim mxlCalculation As XlCalculation
Dim mrngLastChangedRange As Range
Dim mwshLastChangedSheet As Worksheet
'
Public Sub SetCustomAppOptions()
    #If Not DebugEnabled Then
        Application.EnableCancelKey = xlDisabled
    #End If
    GetOriginalAppOptions
    'Set Custom App options
    Application.Cursor = xlWait
    Application.CalculateBeforeSave = False
    ThisWorkbook.EnableAutoRecover = False
    Application.Calculation = xlCalculationManual
    Application.EnableLargeOperationAlert = False
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Application.EnableCancelKey = xlDisabled
    'ScreenDrawing False
    Application.EnableEvents = False
End Sub

Private Sub GetOriginalAppOptions()
    If Not gfOriginalAppOptionsAreCurrent Then
        gxlEnableCancelKey = Application.EnableCancelKey
        mfCalculateBeforeSave = Application.CalculateBeforeSave
        mfEnableAutoRecover = ThisWorkbook.EnableAutoRecover
        mxlCalculation = Application.Calculation
        mfEnableLargeOperationAlert = Application.EnableLargeOperationAlert
        gfOriginalAppOptionsAreCurrent = True
    End If
    Application.EnableCancelKey = xlDisabled
End Sub

Public Sub SetOriginalAppOptions()
On Error Resume Next
    Application.CalculateBeforeSave = mfCalculateBeforeSave
    ThisWorkbook.EnableAutoRecover = mfEnableAutoRecover
    Application.Calculation = xlCalculationAutomatic 'overridding mxlCalculation
    Application.EnableLargeOperationAlert = mfEnableLargeOperationAlert
    If Not Application.DisplayAlerts Then
        Application.DisplayAlerts = True
    End If
    Application.Cursor = xlDefault
    'Re enable Cancel Key for this application
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableEvents = True
    Application.ScreenUpdating = True ' This must be set to true
    'ScreenDrawing True
    DoEvents
ExitHere:
    Exit Sub
HandleError:
'    Select Case HandleError("basAppOptions.SetOriginalAppOptions", Err)
'        Case hecResume
'            Resume
'        Case hecResumeNext
'            Resume Next
'        Case hecExitProcedure
'            Resume ExitHere
'    End Select
End Sub

Public Sub ResetAppOptions()
    'Set Custom App options
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.CalculateBeforeSave = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableLargeOperationAlert = True
    Application.Cursor = xlDefault
    Application.EnableCancelKey = xlErrorHandler
    ThisWorkbook.EnableAutoRecover = True
   ' ScreenDrawing True
End Sub

