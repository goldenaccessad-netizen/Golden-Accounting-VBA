Attribute VB_Name = "Module2"
Option Explicit

'=========================
' كلمة مرور الإدارة (واحدة فقط)
'=========================
' Public Const ADMIN_PWD As String = "mina2040"
    

'Flag لمنع رسائل/أحداث أثناء التفريغ
Public IsClearingInvoice As Boolean

 '=========================
 ' نطاقات السماح بالكتابة
 '=========================

'=========================
' قفل الحماية (للإدارة)
'=========================
Public Sub Lock_All()
    AdminMode = False

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ProtectInvoiceSheet True
    ProtectKashfSheet True
    ProtectCustomersSheet True
    ProtectTemplateSheet True
    ProtectCustomerSheets True

    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "تم قفل الحماية بنجاح", vbInformation
End Sub


'=========================
' فتح الحماية (للإدارة) - باستخدام UserForm1
'=========================
Public Sub Unlock_All()
    Dim frm As UserForm1
    Set frm = New UserForm1

    frm.Show vbModal
    If frm.IsOk = False Then Exit Sub

    If frm.EnteredPassword <> ADMIN_PWD Then
        MsgBox "كلمة المرور غير صحيحة", vbCritical
        Exit Sub
    End If

    AdminMode = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Call TryUnprotectWorkbook

    TryUnprotectSheet ThisWorkbook.Worksheets(SHEET_INVOICE)
    TryUnprotectSheet ThisWorkbook.Worksheets(SHEET_CUSTOMER_STATEMENT)
    TryUnprotectSheet ThisWorkbook.Worksheets(SHEET_CUSTOMERS)
    TryUnprotectSheet ThisWorkbook.Worksheets(SHEET_TEMPLATE)

    UnprotectCustomerSheets

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "تم فتح الحماية (وضع الإدارة) بنجاح", vbInformation
End Sub

'====================================================
' حماية شيت إدخال_فاتورة (السماح بالإدخال فقط)
'====================================================
Private Sub ProtectInvoiceSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_INVOICE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        On Error GoTo CleanExit
        Call TryUnprotectSheet(ws)

        ws.Cells.Locked = True
        ws.Range(RANGE_INVOICE_UNLOCK).Locked = False

        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        ws.EnableSelection = xlUnlockedCells
    End If
    Exit Sub

CleanExit:
    RestoreProtectSheet ws, True, True, True
End Sub

'====================================================
' حماية شيت كشف_حساب_العملاء
'====================================================
Private Sub ProtectKashfSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CUSTOMER_STATEMENT)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        On Error GoTo CleanExit
        Call TryUnprotectSheet(ws)

        ws.Cells.Locked = True
        ws.Range(RANGE_KASHF_UNLOCK).Locked = False

        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
        ws.EnableSelection = xlUnlockedCells
    End If
    Exit Sub

CleanExit:
    RestoreProtectSheet ws, True
End Sub

'====================================================
' حماية شيت قائمة_عملاء (السماح بالكتابة في A فقط)
'====================================================
Private Sub ProtectCustomersSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CUSTOMERS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        On Error GoTo CleanExit
        Call TryUnprotectSheet(ws)

        ws.Cells.Locked = True
        ws.Range(RANGE_CUSTOMERS_UNLOCK).Locked = False

        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
        ws.EnableSelection = xlUnlockedCells
    End If
    Exit Sub

CleanExit:
    RestoreProtectSheet ws, True
End Sub

'====================================================
' حماية شيت القالب
'====================================================
Private Sub ProtectTemplateSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_TEMPLATE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        On Error GoTo CleanExit
        Call TryUnprotectSheet(ws)
        ws.Cells.Locked = True
        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
        ws.EnableSelection = xlUnlockedCells
    End If
    Exit Sub

CleanExit:
    RestoreProtectSheet ws, True
End Sub

'====================================================
' حماية شيتات العملاء (أي شيت غير الأساسية)
'====================================================
Private Sub ProtectCustomerSheets(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_INVOICE _
           And ws.Name <> SHEET_CUSTOMER_STATEMENT _
           And ws.Name <> SHEET_CUSTOMERS _
           And ws.Name <> SHEET_TEMPLATE Then

            If doProtect Then
                On Error GoTo CleanExit
                Call TryUnprotectSheet(ws)
                ws.Cells.Locked = True
                ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next ws
    Exit Sub

CleanExit:
    RestoreProtectSheet ws, True
End Sub

Private Sub UnprotectCustomerSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_INVOICE _
           And ws.Name <> SHEET_CUSTOMER_STATEMENT _
           And ws.Name <> SHEET_CUSTOMERS _
           And ws.Name <> SHEET_TEMPLATE Then
            TryUnprotectSheet ws
        End If
    Next ws
End Sub

'=========================
' تفريغ الفاتورة بدون حفظ (مع تأكيد)
'=========================
Public Sub Clear_Invoice_Without_Save()

    Dim wsI As Worksheet
    Set wsI = ThisWorkbook.Worksheets(SHEET_INVOICE)

    If MsgBox("هل تريد تفريغ الفاتورة بدون حفظ؟", vbYesNo + vbQuestion, "تأكيد التفريغ") = vbNo Then Exit Sub

    On Error GoTo SafeExit

    IsClearingInvoice = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'تفريغ العميل + الكومبو
    wsI.Range(CELL_INVOICE_CUSTOMER).ClearContents
    On Error Resume Next
    wsI.OLEObjects("ComboBox1").Object.Value = ""
    wsI.OLEObjects("ComboBox1").Object.Text = ""
    On Error GoTo SafeExit

    'تفريغ الهيدر
    wsI.Range(CELL_INVOICE_NUMBER).ClearContents
    wsI.Range(CELL_INVOICE_DATE).ClearContents
    wsI.Range(RANGE_INVOICE_HEADER).ClearContents

    'تفريغ بنود الإدخال فقط (بدون مسح معادلات H و J)
    wsI.Range(RANGE_INVOICE_ITEMS_C).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_DE).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_F).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_G).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_I).ClearContents

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    IsClearingInvoice = False

    If Err.Number <> 0 Then
        MsgBox "حدث خطأ أثناء التفريغ: " & Err.Description, vbExclamation
        Exit Sub
    End If

    MsgBox "تم تفريغ الفاتورة بنجاح (بدون حفظ).", vbInformation, "تم التفريغ"
End Sub


Public Sub HideTempCustomerSheet()
    On Error GoTo ErrH

    If TempOpenedCustomerSheet = "" Then Exit Sub
    If SheetExists(TempOpenedCustomerSheet) = False Then
        TempOpenedCustomerSheet = ""
        Exit Sub
    End If

    'فك Structure مؤقتًا
    Dim wasProtected As Boolean
    wasProtected = TryUnprotectWorkbook()

    'مهم: لا يمكن إخفاء الشيت النشط
    If ActiveSheet.Name = TempOpenedCustomerSheet Then
        RestoreProtectWorkbook wasProtected
        Exit Sub
    End If

    ThisWorkbook.Worksheets(TempOpenedCustomerSheet).Visible = xlSheetVeryHidden
    TempOpenedCustomerSheet = ""

    RestoreProtectWorkbook wasProtected
    Exit Sub

ErrH:
    MsgBox "خطأ أثناء إخفاء شيت العميل:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    RestoreProtectWorkbook wasProtected
End Sub


Public Sub HideTempSummarySheet()
    On Error GoTo ErrH

    If TempOpenedSummarySheet = "" Then Exit Sub
    If SheetExists(TempOpenedSummarySheet) = False Then
        TempOpenedSummarySheet = ""
        Exit Sub
    End If

    Dim wasProtected As Boolean
    wasProtected = TryUnprotectWorkbook()

    If ActiveSheet.Name = TempOpenedSummarySheet Then
        RestoreProtectWorkbook wasProtected
        Exit Sub
    End If

    ThisWorkbook.Worksheets(TempOpenedSummarySheet).Visible = xlSheetVeryHidden
    TempOpenedSummarySheet = ""

    RestoreProtectWorkbook wasProtected
    Exit Sub

ErrH:
    MsgBox "خطأ أثناء إخفاء شيت الإجماليات:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    RestoreProtectWorkbook wasProtected
End Sub



