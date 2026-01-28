Attribute VB_Name = "Module1"
Option Explicit

Public TempOpenedCustomerSheet As String
Public TempOpenedSummarySheet As String
Public AdminMode As Boolean
Public IsAddingCustomerFromMenu As Boolean



'========================
' أدوات مساعدة
'========================
Public Function SafeSheetName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("/", "\", "?", "*", "[", "]", ":", "'")
    s = Trim(s)
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), " ")
    Next i
    If Len(s) > 31 Then s = Left$(s, 31)
    SafeSheetName = s
End Function

Public Function SheetExists(ByVal shName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'هل العميل موجود في قائمة_عملاء؟ (بحث عمود A)
Public Function CustomerExistsInList(ByVal customerName As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim v As String

    customerName = Trim(customerName)
    If customerName = "" Then
        CustomerExistsInList = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(SHEET_CUSTOMERS)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow
        v = Trim(CStr(ws.Cells(i, "A").Value))
        If StrComp(v, customerName, vbTextCompare) = 0 Then
            CustomerExistsInList = True
            Exit Function
        End If
    Next i

    CustomerExistsInList = False
End Function

'========================
' إنشاء شيت عميل من القالب (مخفي VeryHidden)
' - لا يفتح الشيت بعد الإنشاء
' - يمنع ترك _قالب_عميل (2) عند فشل التسمية
' - يتعامل مع القفل Structure بشكل مؤكد
'========================
Public Sub CreateCustomerSheet(ByVal SheetName As String)

    Dim nm As String
    Dim wsTemplate As Worksheet
    Dim wsNew As Worksheet
    Dim wsBack As Worksheet
    Dim lastIndex As Long
    Dim wasProtected As Boolean

    nm = SafeSheetName(SheetName)
    If nm = "" Then Exit Sub
    If SheetExists(nm) Then Exit Sub

    Set wsBack = ActiveSheet

    'تأكد من وجود القالب
    Set wsTemplate = Nothing
    On Error Resume Next
    Set wsTemplate = ThisWorkbook.Worksheets(SHEET_TEMPLATE)
    On Error GoTo 0
    If wsTemplate Is Nothing Then
        MsgBox "شيت القالب غير موجود: _قالب_عميل", vbCritical
        Exit Sub
    End If

    'هل Structure مقفولة؟
    wasProtected = TryUnprotectWorkbook()

    'فك Structure مؤقتًا وبـتأكيد
    If wasProtected Then
        If ThisWorkbook.ProtectStructure Then
            MsgBox "لا يمكن إنشاء شيت العميل لأن بنية المصنف ما زالت محمية." & vbCrLf & _
                   "اضغط (فتح الملف) أولاً ثم أعد المحاولة.", vbCritical
            GoTo CleanExit
        End If
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    lastIndex = ThisWorkbook.Sheets.Count

    'نسخ القالب
    wsTemplate.Copy After:=ThisWorkbook.Sheets(lastIndex)

    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    On Error GoTo RenameFailed
    wsNew.Name = nm
    wsNew.Visible = xlSheetVeryHidden

CleanExit:
    'ارجع كما كنت
    On Error Resume Next
    wsBack.Activate
    On Error GoTo 0

    RestoreProtectWorkbook wasProtected

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

RenameFailed:
    On Error Resume Next
    wsNew.Delete
    On Error GoTo 0

    MsgBox "فشل إنشاء شيت العميل بسبب مشكلة في الاسم.", vbCritical
    Resume CleanExit
End Sub


'========================
' فحص الفاتورة قبل الحفظ (حسب اتفاقنا النهائي)
' - لا ينشئ عميل/شيت
' - يتأكد من: العميل موجود بالقائمة + التاريخ موجود + الشيت موجود
'========================
Public Function ValidateInvoiceForSave() As Boolean
    Dim ws As Worksheet, r As Long
    Dim customer As String, invDate As Variant
    Dim targetSheetName As String

    Set ws = ThisWorkbook.Worksheets(SHEET_INVOICE)

    customer = Trim(CStr(ws.Range(CELL_INVOICE_CUSTOMER).Value))
    invDate = ws.Range(CELL_INVOICE_DATE).Value

    '1) العميل
    If customer = "" Then
        MsgBox "اختر/اكتب اسم العميل أولاً.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '2) العميل موجود في قائمة العملاء؟
    If CustomerExistsInList(customer) = False Then
        MsgBox "هذا العميل غير موجود في قائمة العملاء." & vbCrLf & _
               "من فضلك أضف العميل أولاً ثم قم بالحفظ.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '3) التاريخ I2
    If IsEmpty(invDate) Or Trim(CStr(invDate)) = "" Then
        MsgBox "أدخل تاريخ الفاتورة في الخلية I2.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    If Not IsDate(invDate) Then
        MsgBox "تاريخ الفاتورة غير صحيح. أدخل تاريخ صحيح في I2.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '4) رقم الفاتورة F2
    If Trim(CStr(ws.Range(CELL_INVOICE_NUMBER).Value)) = "" Then
        MsgBox "ادخل رقم الفاتورة.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '5) فحص بنود الفاتورة (كما كان عندك)
    For r = 7 To 31
        If Trim(CStr(ws.Cells(r, "C").Value)) <> "" Then

            If Trim(CStr(ws.Cells(r, "F").Value)) = "" Then
                MsgBox "اختر الوحدة (عدد/قياس) (السطر " & r & ")", vbExclamation
                ValidateInvoiceForSave = False
                Exit Function
            End If

            If Val(ws.Cells(r, "G").Value) <= 0 Then
                MsgBox "ادخل عدد صحيح أكبر من صفر (السطر " & r & ")", vbExclamation
                ValidateInvoiceForSave = False
                Exit Function
            End If

            If ws.Cells(r, "F").Value = "قياس" Then
                If Val(ws.Cells(r, "D").Value) <= 0 Or Val(ws.Cells(r, "E").Value) <= 0 Then
                    MsgBox "في حالة (قياس) يجب إدخال العرض والارتفاع (السطر " & r & ")", vbExclamation
                    ValidateInvoiceForSave = False
                    Exit Function
                End If
            End If

            If Val(ws.Cells(r, "I").Value) <= 0 Then
                MsgBox "السعر لازم يكون أكبر من صفر (السطر " & r & ")", vbExclamation
                ValidateInvoiceForSave = False
                Exit Function
            End If
        End If
    Next r

    '6) شيت العميل لازم يكون موجود (لأننا لن ننشئه عند الحفظ)
    targetSheetName = SafeSheetName(customer)
    If SheetExists(targetSheetName) = False Then
        MsgBox "شيت العميل غير موجود بعد." & vbCrLf & _
               "من فضلك أضف العميل (ليتم إنشاء شيت العميل) ثم أعد الحفظ.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    ValidateInvoiceForSave = True
End Function

'========================
' حفظ/ترحيل الفاتورة إلى شيت العميل
' - لا ينشئ عميل/شيت
' - يرحّل فقط بعد التحقق
'========================
Public Sub AddInvoice()

    Dim wsI As Worksheet, wsTarget As Worksheet
    Dim customer As String, invNo As String
    Dim invDate As Date
    Dim r As Long, lastTarget As Long
    Dim targetSheetName As String
    Dim oldVis As XlSheetVisibility
    Dim wasProtected As Boolean

    If ValidateInvoiceForSave = False Then Exit Sub
    On Error GoTo CleanExit

    Set wsI = ThisWorkbook.Worksheets(SHEET_INVOICE)

    customer = Trim(CStr(wsI.Range(CELL_INVOICE_CUSTOMER).Value))
    invNo = Trim(CStr(wsI.Range(CELL_INVOICE_NUMBER).Value))
    invDate = CDate(wsI.Range(CELL_INVOICE_DATE).Value)

    targetSheetName = SafeSheetName(customer)
    Set wsTarget = ThisWorkbook.Worksheets(targetSheetName)

    'فك قفل بنية المصنف مؤقتًا
    wasProtected = TryUnprotectWorkbook()

    'افتح الشيت مؤقتًا للترحيل ثم ارجعه كما كان
    oldVis = wsTarget.Visible
    wsTarget.Visible = xlSheetVisible

    For r = 7 To 31
        If Trim(CStr(wsI.Cells(r, "C").Value)) <> "" Then
            lastTarget = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1

            wsTarget.Cells(lastTarget, "A").Value = invNo
            wsTarget.Cells(lastTarget, "B").Value = invDate
            wsTarget.Cells(lastTarget, "C").Value = wsI.Cells(r, "C").Value
            wsTarget.Cells(lastTarget, "D").Value = wsI.Cells(r, "D").Value
            wsTarget.Cells(lastTarget, "E").Value = wsI.Cells(r, "E").Value
            wsTarget.Cells(lastTarget, "F").Value = wsI.Cells(r, "F").Value
            wsTarget.Cells(lastTarget, "G").Value = wsI.Cells(r, "G").Value
            wsTarget.Cells(lastTarget, "H").Value = wsI.Cells(r, "H").Value
            wsTarget.Cells(lastTarget, "I").Value = wsI.Cells(r, "I").Value
            wsTarget.Cells(lastTarget, "J").Value = wsI.Cells(r, "J").Value
        End If
    Next r

    wsTarget.Visible = oldVis

    'تفريغ بيانات الفاتورة بدون مسح المعادلات
    wsI.Range(CELL_INVOICE_NUMBER).ClearContents
    wsI.Range(CELL_INVOICE_DATE).ClearContents
    wsI.Range(RANGE_INVOICE_HEADER).ClearContents

    wsI.Range(RANGE_INVOICE_ITEMS_C).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_DE).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_F).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_G).ClearContents
    wsI.Range(RANGE_INVOICE_ITEMS_I).ClearContents

    'تفريغ العميل (B2) + ComboBox1 إن وجد
    wsI.Range(CELL_INVOICE_CUSTOMER).ClearContents
    On Error Resume Next
    wsI.OLEObjects("ComboBox1").Object.Value = ""
    wsI.OLEObjects("ComboBox1").Object.Text = ""
    On Error GoTo CleanExit

    MsgBox "? تم حفظ الفاتورة وترحيلها إلى حساب العميل: " & customer, vbInformation

CleanExit:
    RestoreProtectWorkbook wasProtected
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'========================
' فتح كشف حساب العملاء
'========================
Public Sub Open_Kashf_Hesab()
    On Error GoTo ErrH
    ThisWorkbook.Worksheets(SHEET_CUSTOMER_STATEMENT).Activate
    Exit Sub
ErrH:
    MsgBox "شيت (كشف_حساب_العملاء) غير موجود. تأكد من اسم الشيت.", vbExclamation
End Sub

'فتح صفحة إدخال الفاتورة
Public Sub Open_InvoiceEntry()
    On Error GoTo ErrH
    ThisWorkbook.Worksheets(SHEET_INVOICE).Activate
    Exit Sub
ErrH:
    MsgBox "شيت (إدخال_فاتورة) غير موجود. تأكد من الاسم.", vbExclamation
End Sub

'========================
' إضافة عميل جديد (يضيف الاسم في قائمة_عملاء + ينشئ شيت العميل مخفي)
'========================
Public Sub Add_New_Customer()

    Dim ws As Worksheet
    Dim newName As String
    Dim lastRow As Long
    Dim nm As String

    On Error GoTo ErrH
    Set ws = ThisWorkbook.Worksheets(SHEET_CUSTOMERS)

    newName = InputBox("اكتب اسم العميل الجديد:", "إضافة عميل")
    newName = Trim(newName)
    If newName = "" Then Exit Sub

    nm = SafeSheetName(newName)
    If nm = "" Then
        MsgBox "اسم العميل غير صالح.", vbExclamation
        Exit Sub
    End If

    If CustomerExistsInList(newName) Then
        MsgBox "هذا العميل موجود بالفعل في القائمة.", vbExclamation
        Exit Sub
    End If

    If SheetExists(nm) Then
        MsgBox "يوجد شيت بنفس اسم العميل بالفعل. اختر اسم مختلف.", vbExclamation
        Exit Sub
    End If

    'فك القفل مؤقتًا
    Dim wasProtected As Boolean
    wasProtected = TryUnprotectWorkbook()

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If lastRow < 2 Then lastRow = 2

    ws.Cells(lastRow, "A").Value = newName

    'إنشاء شيت العميل (مخفي)
    CreateCustomerSheet nm
    
    


    'تسجيل اسم الشيت في C (اختياري)
    ws.Cells(lastRow, "C").Value = nm

    'إعادة القفل
    RestoreProtectWorkbook wasProtected

    ws.Activate
    ws.Cells(lastRow, "A").Select
    MsgBox "? تمت إضافة العميل بنجاح: " & newName, vbInformation
    Exit Sub

ErrH:
    RestoreProtectWorkbook wasProtected
    MsgBox "حدث خطأ أثناء إضافة العميل. تأكد من وجود شيت (قائمة_عملاء) وأن كلمة المرور صحيحة.", vbExclamation
    
    
End Sub

'========================
' فتح حساب العميل مؤقتاً (يظهر الشيت ثم يُخفى عند الخروج – عبر ThisWorkbook)
'========================
Public Sub OpenCustomerSheet2()
    Dim shName As String
    Dim wsCust As Worksheet
    Dim wasProtected As Boolean

    shName = Trim(ThisWorkbook.Worksheets(SHEET_CUSTOMER_STATEMENT).Range(CELL_KASHF_CUSTOMER).Value)
    If shName = "" Then
        MsgBox "اختر اسم العميل أولاً", vbExclamation
        Exit Sub
    End If

    shName = SafeSheetName(shName)

    If SheetExists(shName) = False Then
        MsgBox "شيت العميل غير موجود: " & shName, vbExclamation
        Exit Sub
    End If

    On Error GoTo CleanExit
    wasProtected = TryUnprotectWorkbook()

    Set wsCust = ThisWorkbook.Worksheets(shName)
    wsCust.Visible = xlSheetVisible

    TempOpenedCustomerSheet = shName

    wsCust.Activate

    RestoreProtectWorkbook wasProtected
    Exit Sub

CleanExit:
    RestoreProtectWorkbook wasProtected
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub


