VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3285
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IsOk As Boolean
Public EnteredPassword As String

Private Sub UserForm_Initialize()
    IsOk = False
    EnteredPassword = ""

    ' ‰”Ìﬁ »”Ìÿ »«·ﬂÊœ («Œ Ì«—Ì)
    Me.caption = "› Õ «·Õ„«Ì…"
    Me.BackColor = &HF7F7F7

    txtPassword.PasswordChar = "*"
    txtPassword.Value = ""
End Sub

Private Sub UserForm_Activate()
    txtPassword.SetFocus
End Sub

Private Sub cmdOK_Click()
    If Trim(txtPassword.Value) = "" Then
        MsgBox "„‰ ›÷·ﬂ √œŒ· ﬂ·„… «·„—Ê—", vbExclamation
        Exit Sub
    End If

    EnteredPassword = txtPassword.Value
    IsOk = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    IsOk = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '·Ê ÷€ÿ X
    IsOk = False
End Sub

