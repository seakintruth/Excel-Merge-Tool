VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    If Not ThisWorkbook.fDisplayNotice Then
        mLeaveNotice
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    mLeaveNotice
End Sub

Private Sub mLeaveNotice()
    SetEcho False
    RevealSheet ThisWorkbook.Worksheets("Cover")
    ActiveSheet.Visible = xlSheetVeryHidden
    SetEcho True
End Sub
