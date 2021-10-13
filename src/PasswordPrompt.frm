VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PasswordPrompt 
   Caption         =   "Common Password"
   ClientHeight    =   1515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4080
   OleObjectBlob   =   "PasswordPrompt.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PasswordPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrPassword As String

Public Function GetPassword(Optional mstrPasswordCaption As String = vbNullString, Optional mstrPasswordLabel As String = vbNullString) As String
    If Len(mstrPasswordCaption) > 0 Then
        Me.Caption = mstrPasswordCaption
    Else
        Me.Caption = ThisWorkbook.Name & ": Password Prompt"
    End If
    If Len(mstrPasswordLabel) > 0 Then
        Me.lblCommonPassword = mstrPasswordLabel
    Else
        Me.lblCommonPassword = "Enter Password"
    End If
    Me.Show
    GetPassword = mstrPassword
    'Restore
    mstrPasswordCaption = vbNullString
    mstrPasswordLabel = vbNullString
    Me.txtInput = vbNullString
End Function

Private Sub btnOk_Click()
    mstrPassword = Me.txtInput
    Me.Hide
End Sub
    
