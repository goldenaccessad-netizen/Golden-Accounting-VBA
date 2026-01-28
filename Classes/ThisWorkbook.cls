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

Private Sub Workbook_Open()
    On Error Resume Next
    AdminMode = False
    Call Lock_All
    On Error GoTo 0
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error GoTo ErrHandler

    Application.EnableEvents = False

    '·Ê ﬂ«‰ ›ÌÂ ‘Ì  ⁄„Ì· „› ÊÕ „ƒﬁ « Ê› Õ‰« ‘Ì   «‰Ì
    If TempOpenedCustomerSheet <> "" Then
        If Sh.Name <> TempOpenedCustomerSheet Then
            HideTempCustomerSheet
        End If
    End If

    '·Ê ﬂ«‰ ›ÌÂ ‘Ì  ≈Ã„«·Ì«  „› ÊÕ „ƒﬁ «
    If TempOpenedSummarySheet <> "" Then
        If Sh.Name <> TempOpenedSummarySheet Then
            HideTempSummarySheet
        End If
    End If

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    Resume CleanExit
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next

    '⁄‰œ «·Œ—ÊÃ „‰ ‘Ì  ⁄„Ì· „› ÊÕ „ƒﬁ «
    If TempOpenedCustomerSheet <> "" Then
        If Sh.Name = TempOpenedCustomerSheet Then
            ThisWorkbook.Unprotect Password:=ADMIN_PWD
            Sh.Visible = xlSheetVeryHidden
            TempOpenedCustomerSheet = ""
            ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
        End If
    End If

    '⁄‰œ «·Œ—ÊÃ „‰ ‘Ì  ≈Ã„«·Ì«  „› ÊÕ „ƒﬁ «
    If TempOpenedSummarySheet <> "" Then
        If Sh.Name = TempOpenedSummarySheet Then
            ThisWorkbook.Unprotect Password:=ADMIN_PWD
            Sh.Visible = xlSheetVeryHidden
            TempOpenedSummarySheet = ""
            ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
        End If
    End If
End Sub

