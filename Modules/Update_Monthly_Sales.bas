Attribute VB_Name = "Update_Monthly_Sales"
Option Explicit
Public Sub Update_Monthly_Sales_To_TotalsSheet()

    Dim wsSum As Worksheet
    Dim wsList As Worksheet
    Dim wsCust As Worksheet

    Dim monthsArr As Variant
    Dim lastCustRow As Long, i As Long
    Dim outRow As Long
    Dim custName As String, custSheet As String
    Dim m As Long
    Dim totalVal As Double

    '? عدّل لو اسم الشيت مختلف
    Set wsSum = ThisWorkbook.Worksheets(SHEET_TOTAL_SALES)
    Set wsList = ThisWorkbook.Worksheets(SHEET_CUSTOMERS)

    'أسماء الأشهر (لازم تطابق اللي مكتوب في عمود M داخل شيت العميل)
    monthsArr = Array("يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", _
                      "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر")

    '---- كتابة عناوين الجدول (بدون مسح تنسيق) ----
    wsSum.Range("A1").Value = "اسم العميل"
    For m = LBound(monthsArr) To UBound(monthsArr)
        wsSum.Cells(1, 2 + m).Value = monthsArr(m) 'B1..M1
    Next m

    '---- قراءة العملاء من قائمة_عملاء ----
    lastCustRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    If lastCustRow < 2 Then
        MsgBox "لا توجد أسماء عملاء في شيت قائمة_عملاء.", vbExclamation
        Exit Sub
    End If

    '---- اكتب أسماء العملاء في العمود A واملأ الإجماليات ----
    outRow = 2

    'ملاحظة: يمكن تنظيف منطقة الأرقام فقط (بدون هدم التصميم)
    wsSum.Range("A2:M10000").ClearContents

    For i = 2 To lastCustRow

        custName = Trim(CStr(wsList.Cells(i, "A").Value))
        If custName <> "" Then

            custSheet = SafeSheetName(custName)
            wsSum.Cells(outRow, 1).Value = custName

            If SheetExists(custSheet) Then
                Set wsCust = ThisWorkbook.Worksheets(custSheet)

                '? التجميع من العمود L حسب الشهر في العمود M
                For m = LBound(monthsArr) To UBound(monthsArr)
                    On Error Resume Next
                    totalVal = WorksheetFunction.SumIfs(wsCust.Range("L:L"), wsCust.Range("M:M"), monthsArr(m))
                    If Err.Number <> 0 Then totalVal = 0
                    Err.Clear
                    On Error GoTo 0

                    wsSum.Cells(outRow, 2 + m).Value = totalVal
                Next m
            Else
                'لو شيت العميل مش موجود
                For m = LBound(monthsArr) To UBound(monthsArr)
                    wsSum.Cells(outRow, 2 + m).Value = 0
                Next m
            End If

            outRow = outRow + 1
        End If

    Next i

   ' MsgBox "? تم تحديث شيت الإجماليات.", vbInformation

End Sub


