Attribute VB_Name = "Module2"
Option Explicit

'=========================
' ﬂ·„… „—Ê— «·≈œ«—… (Ê«Õœ… ›ﬁÿ)
'=========================
' Public Const ADMIN_PWD As String = "mina2040"
    

'Flag ·„‰⁄ —”«∆·/√Õœ«À √À‰«¡ «· ›—Ì€
Public IsClearingInvoice As Boolean

'=========================
' ‰ÿ«ﬁ«  «·”„«Õ »«·ﬂ «»…
'=========================
Private Const INVOICE_UNLOCK As String = "B2,F2,I2,B3:J3,C7:C31,D7:E31,F7:F31,G7:G31,I7:I31"
Private Const KASHF_UNLOCK As String = "B2"

'=========================
' ﬁ›· «·Õ„«Ì… (··≈œ«—…)
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

    MsgBox " „ ﬁ›· «·Õ„«Ì… »‰Ã«Õ", vbInformation
End Sub


'=========================
' › Õ «·Õ„«Ì… (··≈œ«—…) - »«” Œœ«„ UserForm1
'=========================
Public Sub Unlock_All()
    Dim frm As UserForm1
    Set frm = New UserForm1

    frm.Show vbModal
    If frm.IsOk = False Then Exit Sub

    If frm.EnteredPassword <> ADMIN_PWD Then
        MsgBox "ﬂ·„… «·„—Ê— €Ì— ’ÕÌÕ…", vbCritical
        Exit Sub
    End If

    AdminMode = True

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo 0

    UnprotectSheet "≈œŒ«·_›« Ê—…"
    UnprotectSheet "ﬂ‘›_Õ”«»_«·⁄„·«¡"
    UnprotectSheet "ﬁ«∆„…_⁄„·«¡"
    UnprotectSheet "_ﬁ«·»_⁄„Ì·"

    UnprotectCustomerSheets

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox " „ › Õ «·Õ„«Ì… (Ê÷⁄ «·≈œ«—…) »‰Ã«Õ", vbInformation
End Sub

'====================================================
' Õ„«Ì… ‘Ì  ≈œŒ«·_›« Ê—… («·”„«Õ »«·≈œŒ«· ›ﬁÿ)
'====================================================
Private Sub ProtectInvoiceSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("≈œŒ«·_›« Ê—…")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        ws.Unprotect Password:=ADMIN_PWD

        ws.Cells.Locked = True
        ws.Range(INVOICE_UNLOCK).Locked = False

        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
        ws.EnableSelection = xlUnlockedCells
    End If
End Sub

'====================================================
' Õ„«Ì… ‘Ì  ﬂ‘›_Õ”«»_«·⁄„·«¡
'====================================================
Private Sub ProtectKashfSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ﬂ‘›_Õ”«»_«·⁄„·«¡")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        ws.Unprotect Password:=ADMIN_PWD

        ws.Cells.Locked = True
        ws.Range(KASHF_UNLOCK).Locked = False

        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
        ws.EnableSelection = xlUnlockedCells
    End If
End Sub

'====================================================
' Õ„«Ì… ‘Ì  ﬁ«∆„…_⁄„·«¡ («·”„«Õ »«·ﬂ «»… ›Ì A ›ﬁÿ)
'====================================================
Private Sub ProtectCustomersSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        ws.Unprotect Password:=ADMIN_PWD

        ws.Cells.Locked = True
        ws.Range("A2:A10000").Locked = False

        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
        ws.EnableSelection = xlUnlockedCells
    End If
End Sub

'====================================================
' Õ„«Ì… ‘Ì  «·ﬁ«·»
'====================================================
Private Sub ProtectTemplateSheet(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_ﬁ«·»_⁄„Ì·")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If doProtect Then
        ws.Unprotect Password:=ADMIN_PWD
        ws.Cells.Locked = True
        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
        ws.EnableSelection = xlUnlockedCells
    End If
End Sub

'====================================================
' Õ„«Ì… ‘Ì «  «·⁄„·«¡ (√Ì ‘Ì  €Ì— «·√”«”Ì…)
'====================================================
Private Sub ProtectCustomerSheets(ByVal doProtect As Boolean)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "≈œŒ«·_›« Ê—…" _
           And ws.Name <> "ﬂ‘›_Õ”«»_«·⁄„·«¡" _
           And ws.Name <> "ﬁ«∆„…_⁄„·«¡" _
           And ws.Name <> "_ﬁ«·»_⁄„Ì·" Then

            If doProtect Then
                ws.Unprotect Password:=ADMIN_PWD
                ws.Cells.Locked = True
                ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True
                ws.EnableSelection = xlUnlockedCells
            End If
        End If
    Next ws
End Sub

Private Sub UnprotectCustomerSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "≈œŒ«·_›« Ê—…" _
           And ws.Name <> "ﬂ‘›_Õ”«»_«·⁄„·«¡" _
           And ws.Name <> "ﬁ«∆„…_⁄„·«¡" _
           And ws.Name <> "_ﬁ«·»_⁄„Ì·" Then
            ws.Unprotect Password:=ADMIN_PWD
        End If
    Next ws
End Sub

Private Sub UnprotectSheet(ByVal SheetName As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(SheetName).Unprotect Password:=ADMIN_PWD
    On Error GoTo 0
End Sub

'=========================
'  ›—Ì€ «·›« Ê—… »œÊ‰ Õ›Ÿ („⁄  √ﬂÌœ)
'=========================
Public Sub Clear_Invoice_Without_Save()

    Dim wsI As Worksheet
    Set wsI = ThisWorkbook.Worksheets("≈œŒ«·_›« Ê—…")

    If MsgBox("Â·  —Ìœ  ›—Ì€ «·›« Ê—… »œÊ‰ Õ›Ÿø", vbYesNo + vbQuestion, " √ﬂÌœ «· ›—Ì€") = vbNo Then Exit Sub

    On Error GoTo SafeExit

    IsClearingInvoice = True
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' ›—Ì€ «·⁄„Ì· + «·ﬂÊ„»Ê
    wsI.Range("B2").ClearContents
    On Error Resume Next
    wsI.OLEObjects("ComboBox1").Object.Value = ""
    wsI.OLEObjects("ComboBox1").Object.Text = ""
    On Error GoTo SafeExit

    ' ›—Ì€ «·ÂÌœ—
    wsI.Range("F2").ClearContents
    wsI.Range("I2").ClearContents
    wsI.Range("B3:J3").ClearContents

    ' ›—Ì€ »‰Êœ «·≈œŒ«· ›ﬁÿ (»œÊ‰ „”Õ „⁄«œ·«  H Ê J)
    wsI.Range("C7:C31").ClearContents
    wsI.Range("D7:E31").ClearContents
    wsI.Range("F7:F31").ClearContents
    wsI.Range("G7:G31").ClearContents
    wsI.Range("I7:I31").ClearContents

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    IsClearingInvoice = False

    If Err.Number <> 0 Then
        MsgBox "ÕœÀ Œÿ√ √À‰«¡ «· ›—Ì€: " & Err.Description, vbExclamation
        Exit Sub
    End If

    MsgBox " „  ›—Ì€ «·›« Ê—… »‰Ã«Õ (»œÊ‰ Õ›Ÿ).", vbInformation, " „ «· ›—Ì€"
End Sub


Public Sub HideTempCustomerSheet()
    On Error GoTo ErrH

    If TempOpenedCustomerSheet = "" Then Exit Sub
    If SheetExists(TempOpenedCustomerSheet) = False Then
        TempOpenedCustomerSheet = ""
        Exit Sub
    End If

    '›ﬂ Structure „ƒﬁ «
    ThisWorkbook.Unprotect Password:=ADMIN_PWD

    '„Â„: ·« Ì„ﬂ‰ ≈Œ›«¡ «·‘Ì  «·‰‘ÿ
    If ActiveSheet.Name = TempOpenedCustomerSheet Then Exit Sub

    ThisWorkbook.Worksheets(TempOpenedCustomerSheet).Visible = xlSheetVeryHidden
    TempOpenedCustomerSheet = ""

    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    Exit Sub

ErrH:
    MsgBox "Œÿ√ √À‰«¡ ≈Œ›«¡ ‘Ì  «·⁄„Ì·:" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
End Sub


Public Sub HideTempSummarySheet()
    On Error GoTo ErrH

    If TempOpenedSummarySheet = "" Then Exit Sub
    If SheetExists(TempOpenedSummarySheet) = False Then
        TempOpenedSummarySheet = ""
        Exit Sub
    End If

    ThisWorkbook.Unprotect Password:=ADMIN_PWD

    If ActiveSheet.Name = TempOpenedSummarySheet Then Exit Sub

    ThisWorkbook.Worksheets(TempOpenedSummarySheet).Visible = xlSheetVeryHidden
    TempOpenedSummarySheet = ""

    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    Exit Sub

ErrH:
    MsgBox "Œÿ√ √À‰«¡ ≈Œ›«¡ ‘Ì  «·≈Ã„«·Ì« :" & vbCrLf & _
           Err.Number & " - " & Err.Description, vbCritical
    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
End Sub



