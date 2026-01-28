Attribute VB_Name = "Update_Monthly_Sales"
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

    '? ⁄œ¯· ·Ê «”„ «·‘Ì  „Œ ·›
    Set wsSum = ThisWorkbook.Worksheets("≈Ã„«·Ì_«·„»Ì⁄« ")
    Set wsList = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")

    '√”„«¡ «·√‘Â— (·«“„  ÿ«»ﬁ «··Ì „ﬂ Ê» ›Ì ⁄„Êœ M œ«Œ· ‘Ì  «·⁄„Ì·)
    monthsArr = Array("Ì‰«Ì—", "›»—«Ì—", "„«—”", "√»—Ì·", "„«ÌÊ", "ÌÊ‰ÌÊ", _
                      "ÌÊ·ÌÊ", "√€”ÿ”", "”» „»—", "√ﬂ Ê»—", "‰Ê›„»—", "œÌ”„»—")

    '---- ﬂ «»… ⁄‰«ÊÌ‰ «·ÃœÊ· (»œÊ‰ „”Õ  ‰”Ìﬁ) ----
    wsSum.Range("A1").Value = "«”„ «·⁄„Ì·"
    For m = LBound(monthsArr) To UBound(monthsArr)
        wsSum.Cells(1, 2 + m).Value = monthsArr(m) 'B1..M1
    Next m

    '---- ﬁ—«¡… «·⁄„·«¡ „‰ ﬁ«∆„…_⁄„·«¡ ----
    lastCustRow = wsList.Cells(wsList.Rows.Count, "A").End(xlUp).Row
    If lastCustRow < 2 Then
        MsgBox "·«  ÊÃœ √”„«¡ ⁄„·«¡ ›Ì ‘Ì  ﬁ«∆„…_⁄„·«¡.", vbExclamation
        Exit Sub
    End If

    '---- «ﬂ » √”„«¡ «·⁄„·«¡ ›Ì «·⁄„Êœ A Ê«„·√ «·≈Ã„«·Ì«  ----
    outRow = 2

    '„·«ÕŸ…: Ì„ﬂ‰  ‰ŸÌ› „‰ÿﬁ… «·√—ﬁ«„ ›ﬁÿ (»œÊ‰ Âœ„ «· ’„Ì„)
    wsSum.Range("A2:M10000").ClearContents

    For i = 2 To lastCustRow

        custName = Trim(CStr(wsList.Cells(i, "A").Value))
        If custName <> "" Then

            custSheet = SafeSheetName(custName)
            wsSum.Cells(outRow, 1).Value = custName

            If SheetExists(custSheet) Then
                Set wsCust = ThisWorkbook.Worksheets(custSheet)

                '? «· Ã„Ì⁄ „‰ «·⁄„Êœ L Õ”» «·‘Â— ›Ì «·⁄„Êœ M
                For m = LBound(monthsArr) To UBound(monthsArr)
                    On Error Resume Next
                    totalVal = WorksheetFunction.SumIfs(wsCust.Range("L:L"), wsCust.Range("M:M"), monthsArr(m))
                    If Err.Number <> 0 Then totalVal = 0
                    Err.Clear
                    On Error GoTo 0

                    wsSum.Cells(outRow, 2 + m).Value = totalVal
                Next m
            Else
                '·Ê ‘Ì  «·⁄„Ì· „‘ „ÊÃÊœ
                For m = LBound(monthsArr) To UBound(monthsArr)
                    wsSum.Cells(outRow, 2 + m).Value = 0
                Next m
            End If

            outRow = outRow + 1
        End If

    Next i

   ' MsgBox "?  „  ÕœÌÀ ‘Ì  «·≈Ã„«·Ì« .", vbInformation

End Sub


