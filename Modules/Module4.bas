Attribute VB_Name = "Module4"
Option Explicit

'========================================================
' 1) تحديث ملخص حسابات العملاء (لا يخفي الشيت وهو نشط)
'========================================================
Public Sub UpdateCustomerAccountsSummary(Optional ByVal HideAfterUpdate As Boolean = True)

    Dim wsSum As Worksheet, wsList As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, i As Long, outRow As Long
    Dim custName As String, shName As String
    Dim wsBack As Worksheet

    Set wsList = ThisWorkbook.Worksheets(SHEET_CUSTOMERS)

    'احفظ الشيت الحالي للرجوع له لو احتجنا نخفي ملخص وهو نشط
    Set wsBack = ActiveSheet

    'فك Structure مؤقتًا
    Dim wasProtected As Boolean
    On Error GoTo CleanExit
    wasProtected = TryUnprotectWorkbook()

    'احصل على شيت الملخص أو أنشئه
    On Error Resume Next
    Set wsSum = ThisWorkbook.Worksheets(SHEET_ACCOUNTS_SUMMARY)
    On Error GoTo 0
    On Error GoTo CleanExit

    If wsSum Is Nothing Then
        Set wsSum = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        wsSum.Name = SHEET_ACCOUNTS_SUMMARY
        On Error GoTo 0
        On Error GoTo CleanExit
    End If

    'تجهيز الهيدر (صف 1)
    wsSum.Range("A1").Value = "اسم العميل"
    wsSum.Range("B1").Value = "إجمالي المبيعات"
    wsSum.Range("C1").Value = "إجمالي المدفوعات"
    wsSum.Range("D1").Value = "الرصيد"
    wsSum.Rows(1).Font.Bold = True

    'امسح بيانات فقط من صف 2
    wsSum.Range("A2:D100000").ClearContents

    lastRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    outRow = 2

    For i = 2 To lastRow
        custName = Trim(CStr(wsList.Cells(i, "A").Value))
        If custName <> "" Then

            shName = SafeSheetName(custName)

            wsSum.Cells(outRow, "A").Value = custName

            If SheetExists(shName) Then
                Set wsCust = ThisWorkbook.Worksheets(shName)
                wsSum.Cells(outRow, "B").Value = wsCust.Range("K2").Value
                wsSum.Cells(outRow, "C").Value = wsCust.Range("K3").Value
                wsSum.Cells(outRow, "D").Value = wsCust.Range("K4").Value
            Else
                wsSum.Cells(outRow, "B").Value = 0
                wsSum.Cells(outRow, "C").Value = 0
                wsSum.Cells(outRow, "D").Value = 0
            End If

            outRow = outRow + 1
        End If
    Next i

    'لو مطلوب نخفيه بعد التحديث: لا يمكن إخفاء الشيت النشط
    If HideAfterUpdate Then
        If ActiveSheet.Name = wsSum.Name Then
            wsBack.Activate
        End If
        wsSum.Visible = xlSheetVeryHidden
    End If

    'إعادة قفل Structure
CleanExit:
    RestoreProtectWorkbook wasProtected
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub


'========================================================
' 2) فتح شيت ملخص/إجماليات بكلمة مرور + تحديث تلقائي
'========================================================
Public Sub OpenSummarySheet_WithPassword(ByVal SheetName As String)

    Dim frm As UserForm1

    If SheetExists(SheetName) = False Then
        MsgBox "الشيت غير موجود: " & SheetName, vbExclamation
        Exit Sub
    End If

    Set frm = New UserForm1
    frm.Show vbModal
    If frm.IsOk = False Then Exit Sub

    If frm.EnteredPassword <> ADMIN_PWD Then
        MsgBox "كلمة المرور غير صحيحة", vbCritical
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo CleanExit

    'فك Structure
    Dim wasProtected As Boolean
    wasProtected = TryUnprotectWorkbook()

    'لو شيت الملخص: حدث البيانات
    If SheetName = SHEET_ACCOUNTS_SUMMARY Then
        UpdateCustomerAccountsSummary False
        '? مهم جدًا: الدالة تقفل Structure مرة أخرى، فافتحه تاني هنا
        wasProtected = TryUnprotectWorkbook()
    End If

    'تأكد أن Structure مفتوح قبل تغيير Visible
    If ThisWorkbook.ProtectStructure Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "لا يمكن فتح الشيت لأن بنية المصنف ما زالت محمية.", vbCritical
        Exit Sub
    End If

    With ThisWorkbook.Worksheets(SheetName)
        .Visible = xlSheetVisible
        .Activate
    End With

    TempOpenedSummarySheet = SheetName

    'إعادة القفل
    RestoreProtectWorkbook wasProtected

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

CleanExit:
    RestoreProtectWorkbook wasProtected
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub




'========================================================
' 3) ماكروز جاهزة للأزرار (استدعاء مباشر)
'========================================================
Public Sub Open_AccountsSummary_Sheet()
    OpenSummarySheet_WithPassword SHEET_ACCOUNTS_SUMMARY
End Sub

Public Sub Open_TotalSales_Sheet()
    OpenSummarySheet_WithPassword SHEET_TOTAL_SALES
End Sub


'========================================================
' 4) رجوع + إخفاء الشيت الحالي فورًا (للملخصات)
'========================================================
Public Sub HideCurrentAndGo(ByVal TargetSheet As String)

    Dim wsCurrent As Worksheet
    Set wsCurrent = ActiveSheet

    If SheetExists(TargetSheet) = False Then
        MsgBox "الشيت غير موجود: " & TargetSheet, vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False

    'فك Structure مؤقتًا
    Dim wasProtected As Boolean
    On Error GoTo CleanExit
    wasProtected = TryUnprotectWorkbook()

    'اذهب للهدف أولاً
    ThisWorkbook.Worksheets(TargetSheet).Activate

    'اخفِ الشيت الذي كنت فيه
    wsCurrent.Visible = xlSheetVeryHidden

    'صفّر المتغير لو كان ملخص مفتوح مؤقتاً
    If TempOpenedSummarySheet = wsCurrent.Name Then TempOpenedSummarySheet = ""

    'إعادة القفل
    RestoreProtectWorkbook wasProtected

    Application.EnableEvents = True
    Exit Sub

CleanExit:
    RestoreProtectWorkbook wasProtected
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Public Sub Back_To_CustomersList()
    HideCurrentAndGo SHEET_CUSTOMERS
End Sub

Public Sub Back_To_CustomerStatement()
    HideCurrentAndGo SHEET_CUSTOMER_STATEMENT
End Sub

Public Sub Go_To_CustomersList()
    On Error GoTo ErrH

    If SheetExists(SHEET_CUSTOMERS) = False Then
        MsgBox "شيت قائمة_عملاء غير موجود.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False

    ThisWorkbook.Worksheets(SHEET_CUSTOMERS).Activate

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrH:
    Application.EnableEvents = True
    MsgBox "حدث خطأ أثناء الانتقال إلى قائمة العملاء.", vbCritical
End Sub

