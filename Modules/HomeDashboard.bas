Attribute VB_Name = "HomeDashboard"
Option Explicit

'==============================
' تشغيل مرة واحدة لإنشاء الرئيسية + الأزرار + إخفاء باقي الشيتات
'==============================
Public Sub Setup_Home_Dashboard()
    Dim wsHome As Worksheet

    'إنشاء/تجهيز شيت الرئيسية
    Set wsHome = GetOrCreateHomeSheet

    'امسح أزرار قديمة بالشيت الرئيسية (اختياري)
    DeleteAllShapes wsHome

    'إنشاء الأزرار (Shapes) في الشيت الرئيسية
    AddHomeButton wsHome, "فاتورة مبيعات", "Go_Invoice", 30, 40, 220, 40
    AddHomeButton wsHome, "قائمة العملاء", "Go_CustomersList", 30, 90, 220, 40
    AddHomeButton wsHome, "إضافة عميل", "Go_AddCustomer", 30, 140, 220, 40
    AddHomeButton wsHome, "كشف حساب عميل", "Go_Kashf", 30, 190, 220, 40
    AddHomeButton wsHome, "إجمالي المبيعات", "Go_TotalSales", 30, 240, 220, 40
    AddHomeButton wsHome, "ملخص حسابات العملاء", "Go_AccountsSummary", 30, 290, 220, 40

    'اختياري: قفل/فتح الحماية من الرئيسية
    AddHomeButton wsHome, "قفل الملف", "Lock_All", 280, 40, 220, 40
    AddHomeButton wsHome, "فتح الملف", "Unlock_All", 280, 90, 220, 40

    'إخفاء كل الشيتات ما عدا الرئيسية
    HideAllSheetsExceptHome

    'إضافة زر رجوع لكل الشيتات
    AddBackButtonToAllSheets

    wsHome.Activate
    MsgBox "? تم تجهيز شيت الرئيسية وإخفاء باقي الشيتات + إضافة زر الرجوع.", vbInformation
End Sub

'==============================
' أزرار الرئيسية (تفتح الشيت المطلوب)
'==============================
Public Sub Go_Invoice()
    ShowOnlySheet SHEET_INVOICE
End Sub

Public Sub Go_CustomersList()
    ShowOnlySheet SHEET_CUSTOMERS
End Sub


'========================================
' إضافة عميل جديد بالاسم + تحديث الملخص + الرجوع للرئيسية
'========================================
Public Sub Add_New_Customer_ByName(ByVal newName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nm As String

    On Error GoTo ErrH

    newName = Trim(newName)
    If newName = "" Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(SHEET_CUSTOMERS)

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
        MsgBox "يوجد شيت بنفس اسم العميل بالفعل.", vbExclamation
        Exit Sub
    End If

    '? منع تشغيل Worksheet_Change أثناء الإضافة من القائمة الرئيسية
    IsAddingCustomerFromMenu = True
    Application.EnableEvents = False

    'أول صف فاضي
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If lastRow < 2 Then lastRow = 2

    'إضافة الاسم
    ws.Cells(lastRow, "A").Value = newName

    'إنشاء شيت العميل (مخفي VeryHidden)
    CreateCustomerSheet nm

    'تسجيل اسم الشيت في C
    ws.Cells(lastRow, "C").Value = nm

CleanExit:
    Application.EnableEvents = True
    IsAddingCustomerFromMenu = False

    MsgBox "? تمت إضافة العميل وإنشاء الشيت بنجاح: " & newName, vbInformation
    Exit Sub

ErrH:
    Application.EnableEvents = True
    IsAddingCustomerFromMenu = False
    MsgBox "حدث خطأ أثناء إضافة العميل: " & Err.Description, vbExclamation
End Sub
    

'========================================
' زر إضافة عميل من الشيت الرئيسي (InputBox)
'========================================
Public Sub Go_AddCustomer()
    Dim nm As String
    nm = InputBox("اكتب اسم العميل الجديد:", "إضافة عميل")
    nm = Trim(nm)
    If nm = "" Then Exit Sub

    Add_New_Customer_ByName nm
End Sub



Public Sub Go_Kashf()
    ShowOnlySheet SHEET_CUSTOMER_STATEMENT
End Sub

Public Sub Go_TotalSales()
    'لو أنت عامل فتح بالباسورد + إخفاء تلقائي: استخدم الماكرو الموجود عندك
    Open_TotalSales_Sheet
End Sub

Public Sub Go_AccountsSummary()
    Open_AccountsSummary_Sheet
End Sub

'==============================
' زر الرجوع (يظهر الرئيسية ويخفي الشيت الحالي)
'==============================
Public Sub Back_To_Home()
    Dim wsCur As Worksheet
    Dim wasProtected As Boolean
    Set wsCur = ActiveSheet

    'لا تخفِ الرئيسية
    If wsCur.Name <> SHEET_HOME Then
        On Error GoTo CleanExit
        wasProtected = TryUnprotectWorkbook()
        wsCur.Visible = xlSheetVeryHidden
        RestoreProtectWorkbook wasProtected
    End If

    'اظهر الرئيسية فقط
    HideAllSheetsExceptHome
    ThisWorkbook.Worksheets(SHEET_HOME).Activate
    Exit Sub

CleanExit:
    RestoreProtectWorkbook wasProtected
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'==============================
' إظهار شيت واحد فقط (ويخفي باقي الشيتات)
'==============================
Public Sub ShowOnlySheet(ByVal SheetName As String)

    Dim ws As Worksheet
    Dim wasProtected As Boolean

    If SheetExists(SheetName) = False Then
        MsgBox "الشيت غير موجود: " & SheetName, vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    On Error GoTo CleanExit

    'هل Structure مقفولة؟
    wasProtected = TryUnprotectWorkbook()

    'فك القفل مؤقتًا
    If wasProtected Then
        'تأكد فعليًا أنها اتفكت
        If ThisWorkbook.ProtectStructure Then
            Application.EnableEvents = True
            MsgBox "لا يمكن تغيير إظهار/إخفاء الشيتات لأن بنية المصنف ما زالت محمية." & vbCrLf & _
                   "تأكد من كلمة المرور أو لا يوجد قفل يدوي بكلمة أخرى.", vbCritical
            Exit Sub
        End If
    End If

    'أظهر الشيت المطلوب
    ThisWorkbook.Worksheets(SheetName).Visible = xlSheetVisible

    'فعّله أولاً حتى لا تحاول إخفاء الشيت النشط بالغلط
    ThisWorkbook.Worksheets(SheetName).Activate

    'اخفِ باقي الشيتات (لا تخفِ الشيت النشط)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SheetName Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

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


'==============================
' إخفاء كل الشيتات ما عدا الرئيسية
'==============================
Public Sub HideAllSheetsExceptHome()
    Dim ws As Worksheet
    Dim wasProtected As Boolean

    On Error GoTo CleanExit
    wasProtected = TryUnprotectWorkbook()

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = SHEET_HOME Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    RestoreProtectWorkbook wasProtected
    Exit Sub

CleanExit:
    RestoreProtectWorkbook wasProtected
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

'==============================
' إضافة زر رجوع للرئيسية في كل الشيتات (حتى لا تتوه)
'==============================
Public Sub AddBackButtonToAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SHEET_HOME Then
            AddBackButton ws
        End If
    Next ws
End Sub

Private Sub AddBackButton(ByVal ws As Worksheet)
    'يحذف زر رجوع قديم ثم يضيف جديد
    On Error Resume Next
    ws.Shapes("btnBackHome").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 160, 32)
    shp.Name = "btnBackHome"
    shp.TextFrame2.TextRange.Text = "? رجوع للرئيسية"
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.OnAction = "Back_To_Home"
End Sub

'==============================
' أدوات مساعدة للشيت الرئيسية
'==============================
Private Function GetOrCreateHomeSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_HOME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = SHEET_HOME
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "لوحة التحكم - الرئيسية"
    ws.Range("A1").Font.Size = 18
    ws.Range("A1").Font.Bold = True

    Set GetOrCreateHomeSheet = ws
End Function

Private Sub AddHomeButton(ByVal ws As Worksheet, ByVal caption As String, ByVal macroName As String, _
                          ByVal L As Double, ByVal T As Double, ByVal W As Double, ByVal H As Double)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, L, T, W, H)
    shp.TextFrame2.TextRange.Text = caption
    shp.TextFrame2.TextRange.Font.Size = 13
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.OnAction = macroName
End Sub

Private Sub DeleteAllShapes(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub



Public Sub Reset_Excel_Environment()
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "? تم إعادة تفعيل Events وإعدادات Excel.", vbInformation
End Sub

