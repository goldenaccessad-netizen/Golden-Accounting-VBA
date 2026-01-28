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

    'لو كان فيه شيت عميل مفتوح مؤقتًا وفتحنا شيت تاني
    If TempOpenedCustomerSheet <> "" Then
        If Sh.Name <> TempOpenedCustomerSheet Then
            HideTempCustomerSheet
        End If
    End If

    'لو كان فيه شيت إجماليات مفتوح مؤقتًا
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

    Dim wasProtected As Boolean

    'عند الخروج من شيت عميل مفتوح مؤقتًا
    If TempOpenedCustomerSheet <> "" Then
        If Sh.Name = TempOpenedCustomerSheet Then
            wasProtected = TryUnprotectWorkbook()
            Sh.Visible = xlSheetVeryHidden
            TempOpenedCustomerSheet = ""
            RestoreProtectWorkbook wasProtected
        End If
    End If

    'عند الخروج من شيت إجماليات مفتوح مؤقتًا
    If TempOpenedSummarySheet <> "" Then
        If Sh.Name = TempOpenedSummarySheet Then
            wasProtected = TryUnprotectWorkbook()
            Sh.Visible = xlSheetVeryHidden
            TempOpenedSummarySheet = ""
            RestoreProtectWorkbook wasProtected
        End If
    End If
End Sub

