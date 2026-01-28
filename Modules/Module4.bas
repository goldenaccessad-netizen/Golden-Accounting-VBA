Attribute VB_Name = "Module4"
Option Explicit

'========================================================
' 1)  ÕœÌÀ „·Œ’ Õ”«»«  «·⁄„·«¡ (·« ÌŒ›Ì «·‘Ì  ÊÂÊ ‰‘ÿ)
'========================================================
Public Sub UpdateCustomerAccountsSummary(Optional ByVal HideAfterUpdate As Boolean = True)

    Dim wsSum As Worksheet, wsList As Worksheet, wsCust As Worksheet
    Dim lastRow As Long, i As Long, outRow As Long
    Dim custName As String, shName As String
    Dim wsBack As Worksheet

    Set wsList = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")

    '«Õ›Ÿ «·‘Ì  «·Õ«·Ì ··—ÃÊ⁄ ·Â ·Ê «Õ Ã‰« ‰Œ›Ì „·Œ’ ÊÂÊ ‰‘ÿ
    Set wsBack = ActiveSheet

    '›ﬂ Structure „ƒﬁ «
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo 0

    '«Õ’· ⁄·Ï ‘Ì  «·„·Œ’ √Ê √‰‘∆Â
    On Error Resume Next
    Set wsSum = ThisWorkbook.Worksheets("„·Œ’_Õ”«»« _«·⁄„·«¡")
    On Error GoTo 0

    If wsSum Is Nothing Then
        Set wsSum = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        wsSum.Name = "„·Œ’_Õ”«»« _«·⁄„·«¡"
        On Error GoTo 0
    End If

    ' ÃÂÌ“ «·ÂÌœ— (’› 1)
    wsSum.Range("A1").Value = "«”„ «·⁄„Ì·"
    wsSum.Range("B1").Value = "≈Ã„«·Ì «·„»Ì⁄« "
    wsSum.Range("C1").Value = "≈Ã„«·Ì «·„œ›Ê⁄« "
    wsSum.Range("D1").Value = "«·—’Ìœ"
    wsSum.Rows(1).Font.Bold = True

    '«„”Õ »Ì«‰«  ›ﬁÿ „‰ ’› 2
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

    '·Ê „ÿ·Ê» ‰Œ›ÌÂ »⁄œ «· ÕœÌÀ: ·« Ì„ﬂ‰ ≈Œ›«¡ «·‘Ì  «·‰‘ÿ
    If HideAfterUpdate Then
        If ActiveSheet.Name = wsSum.Name Then
            wsBack.Activate
        End If
        wsSum.Visible = xlSheetVeryHidden
    End If

    '≈⁄«œ… ﬁ›· Structure
    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    On Error GoTo 0

End Sub


'========================================================
' 2) › Õ ‘Ì  „·Œ’/≈Ã„«·Ì«  »ﬂ·„… „—Ê— +  ÕœÌÀ  ·ﬁ«∆Ì
'========================================================
Public Sub OpenSummarySheet_WithPassword(ByVal SheetName As String)

    Dim frm As UserForm1

    If SheetExists(SheetName) = False Then
        MsgBox "«·‘Ì  €Ì— „ÊÃÊœ: " & SheetName, vbExclamation
        Exit Sub
    End If

    Set frm = New UserForm1
    frm.Show vbModal
    If frm.IsOk = False Then Exit Sub

    If frm.EnteredPassword <> ADMIN_PWD Then
        MsgBox "ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ…", vbCritical
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    '›ﬂ Structure
    ThisWorkbook.Unprotect Password:=ADMIN_PWD

    '·Ê ‘Ì  «·„·Œ’: ÕœÀ «·»Ì«‰« 
    If SheetName = "„·Œ’_Õ”«»« _«·⁄„·«¡" Then
        UpdateCustomerAccountsSummary False
        '? „Â„ Ãœ«: «·œ«·…  ﬁ›· Structure „—… √Œ—Ï° ›«› ÕÂ  «‰Ì Â‰«
        ThisWorkbook.Unprotect Password:=ADMIN_PWD
    End If

    ' √ﬂœ √‰ Structure „› ÊÕ ﬁ»·  €ÌÌ— Visible
    If ThisWorkbook.ProtectStructure Then
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        MsgBox "·« Ì„ﬂ‰ › Õ «·‘Ì  ·√‰ »‰Ì… «·„’‰› „« “«·  „Õ„Ì….", vbCritical
        Exit Sub
    End If

    With ThisWorkbook.Worksheets(SheetName)
        .Visible = xlSheetVisible
        .Activate
    End With

    TempOpenedSummarySheet = SheetName

    '≈⁄«œ… «·ﬁ›·
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub




'========================================================
' 3) „«ﬂ—Ê“ Ã«Â“… ··√“—«— («” œ⁄«¡ „»«‘—)
'========================================================
Public Sub Open_AccountsSummary_Sheet()
    OpenSummarySheet_WithPassword "„·Œ’_Õ”«»« _«·⁄„·«¡"
End Sub

Public Sub Open_TotalSales_Sheet()
    OpenSummarySheet_WithPassword "≈Ã„«·Ì_«·„»Ì⁄« "
End Sub


'========================================================
' 4) —ÃÊ⁄ + ≈Œ›«¡ «·‘Ì  «·Õ«·Ì ›Ê—« (··„·Œ’« )
'========================================================
Public Sub HideCurrentAndGo(ByVal TargetSheet As String)

    Dim wsCurrent As Worksheet
    Set wsCurrent = ActiveSheet

    If SheetExists(TargetSheet) = False Then
        MsgBox "«·‘Ì  €Ì— „ÊÃÊœ: " & TargetSheet, vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False

    '›ﬂ Structure „ƒﬁ «
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo 0

    '«–Â» ··Âœ› √Ê·«
    ThisWorkbook.Worksheets(TargetSheet).Activate

    '«Œ›ˆ «·‘Ì  «·–Ì ﬂ‰  ›ÌÂ
    wsCurrent.Visible = xlSheetVeryHidden

    '’›¯— «·„ €Ì— ·Ê ﬂ«‰ „·Œ’ „› ÊÕ „ƒﬁ «
    If TempOpenedSummarySheet = wsCurrent.Name Then TempOpenedSummarySheet = ""

    '≈⁄«œ… «·ﬁ›·
    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    On Error GoTo 0

    Application.EnableEvents = True
End Sub

Public Sub Back_To_CustomersList()
    HideCurrentAndGo "ﬁ«∆„…_⁄„·«¡"
End Sub

Public Sub Back_To_CustomerStatement()
    HideCurrentAndGo "ﬂ‘›_Õ”«»_«·⁄„·«¡"
End Sub

Public Sub Go_To_CustomersList()
    On Error GoTo ErrH

    If SheetExists("ﬁ«∆„…_⁄„·«¡") = False Then
        MsgBox "‘Ì  ﬁ«∆„…_⁄„·«¡ €Ì— „ÊÃÊœ.", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False

    ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡").Activate

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrH:
    Application.EnableEvents = True
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ «·«‰ ﬁ«· ≈·Ï ﬁ«∆„… «·⁄„·«¡.", vbCritical
End Sub

