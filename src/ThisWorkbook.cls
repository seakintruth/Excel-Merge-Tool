VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'Private variable that holds a worksheet object
Private shtSelectedPriorToSave As Worksheet
Public fDisplayNotice As Boolean

Private Sub Workbook_Open()
    fDisplayNotice = False
    RevealSheet ThisWorkbook.Worksheets("Cover")
    HideSheet ThisWorkbook.Worksheets("Notice")
    SetCoverButtonVisibility basUi.mfProcessing
    ' on open attempt to log even if the network is slow
    logStatus "Opened", True
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    logStatus "Closed"
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If Not basUi.mfProcessing Then
       logStatus "Saved"
    End If
    SetCoverButtonVisibility basUi.mfProcessing
    ActivateWorksheet shtSelectedPriorToSave
    HideSheet ThisWorkbook.Worksheets("Notice")
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Set shtSelectedPriorToSave = ThisWorkbook.ActiveSheet
    RevealSheet ThisWorkbook.Worksheets("Notice")
    fDisplayNotice = True
    ThisWorkbook.Worksheets("Notice").Activate
    fDisplayNotice = False
End Sub



