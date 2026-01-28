VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ê—ﬁ…11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()
Call Open_Kashf_Hesab
End Sub

Private Sub Worksheet_Activate()
    On Error GoTo SafeExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    '????? ?????????? ????????
    Update_Monthly_Sales_To_TotalsSheet

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

