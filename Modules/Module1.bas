Attribute VB_Name = "Module1"
Option Explicit

Public Const ADMIN_PWD As String = "mina2040"
Public TempOpenedCustomerSheet As String
Public TempOpenedSummarySheet As String
Public AdminMode As Boolean
Public IsAddingCustomerFromMenu As Boolean



'========================
' √œÊ«  „”«⁄œ…
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

'Â· «·⁄„Ì· „ÊÃÊœ ›Ì ﬁ«∆„…_⁄„·«¡ø (»ÕÀ ⁄„Êœ A)
Public Function CustomerExistsInList(ByVal customerName As String) As Boolean
    Dim ws As Worksheet, lastRow As Long, i As Long
    Dim v As String

    customerName = Trim(customerName)
    If customerName = "" Then
        CustomerExistsInList = False
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")
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
' ≈‰‘«¡ ‘Ì  ⁄„Ì· „‰ «·ﬁ«·» („Œ›Ì VeryHidden)
' - ·« Ì› Õ «·‘Ì  »⁄œ «·≈‰‘«¡
' - Ì„‰⁄  —ﬂ _ﬁ«·»_⁄„Ì· (2) ⁄‰œ ›‘· «· ”„Ì…
' - Ì ⁄«„· „⁄ «·ﬁ›· Structure »‘ﬂ· „ƒﬂœ
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

    ' √ﬂœ „‰ ÊÃÊœ «·ﬁ«·»
    Set wsTemplate = Nothing
    On Error Resume Next
    Set wsTemplate = ThisWorkbook.Worksheets("_ﬁ«·»_⁄„Ì·")
    On Error GoTo 0
    If wsTemplate Is Nothing Then
        MsgBox "‘Ì  «·ﬁ«·» €Ì— „ÊÃÊœ: _ﬁ«·»_⁄„Ì·", vbCritical
        Exit Sub
    End If

    'Â· Structure „ﬁ›Ê·…ø
    wasProtected = ThisWorkbook.ProtectStructure

    '›ﬂ Structure „ƒﬁ « Ê»‹ √ﬂÌœ
    If wasProtected Then
        ThisWorkbook.Unprotect Password:=ADMIN_PWD

        If ThisWorkbook.ProtectStructure Then
            MsgBox "·« Ì„ﬂ‰ ≈‰‘«¡ ‘Ì  «·⁄„Ì· ·√‰ »‰Ì… «·„’‰› „« “«·  „Õ„Ì…." & vbCrLf & _
                   "«÷€ÿ (› Õ «·„·›) √Ê·« À„ √⁄œ «·„Õ«Ê·….", vbCritical
            Exit Sub
        End If
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    lastIndex = ThisWorkbook.Sheets.Count

    '‰”Œ «·ﬁ«·»
    wsTemplate.Copy After:=ThisWorkbook.Sheets(lastIndex)

    Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    On Error GoTo RenameFailed
    wsNew.Name = nm
    wsNew.Visible = xlSheetVeryHidden

CleanExit:
    '«—Ã⁄ ﬂ„« ﬂ‰ 
    On Error Resume Next
    wsBack.Activate
    On Error GoTo 0

    If wasProtected Then
        On Error Resume Next
        ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
        On Error GoTo 0
    End If

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

RenameFailed:
    On Error Resume Next
    wsNew.Delete
    On Error GoTo 0

    MsgBox "›‘· ≈‰‘«¡ ‘Ì  «·⁄„Ì· »”»» „‘ﬂ·… ›Ì «·«”„.", vbCritical
    Resume CleanExit
End Sub


'========================
' ›Õ’ «·›« Ê—… ﬁ»· «·Õ›Ÿ (Õ”» « ›«ﬁ‰« «·‰Â«∆Ì)
' - ·« Ì‰‘∆ ⁄„Ì·/‘Ì 
' - Ì √ﬂœ „‰: «·⁄„Ì· „ÊÃÊœ »«·ﬁ«∆„… + «· «—ÌŒ „ÊÃÊœ + «·‘Ì  „ÊÃÊœ
'========================
Public Function ValidateInvoiceForSave() As Boolean
    Dim ws As Worksheet, r As Long
    Dim customer As String, invDate As Variant
    Dim targetSheetName As String

    Set ws = ThisWorkbook.Worksheets("≈œŒ«·_›« Ê—…")

    customer = Trim(CStr(ws.Range("B2").Value))
    invDate = ws.Range("I2").Value

    '1) «·⁄„Ì·
    If customer = "" Then
        MsgBox "«Œ —/«ﬂ » «”„ «·⁄„Ì· √Ê·«.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '2) «·⁄„Ì· „ÊÃÊœ ›Ì ﬁ«∆„… «·⁄„·«¡ø
    If CustomerExistsInList(customer) = False Then
        MsgBox "Â–« «·⁄„Ì· €Ì— „ÊÃÊœ ›Ì ﬁ«∆„… «·⁄„·«¡." & vbCrLf & _
               "„‰ ›÷·ﬂ √÷› «·⁄„Ì· √Ê·« À„ ﬁ„ »«·Õ›Ÿ.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '3) «· «—ÌŒ I2
    If IsEmpty(invDate) Or Trim(CStr(invDate)) = "" Then
        MsgBox "√œŒ·  «—ÌŒ «·›« Ê—… ›Ì «·Œ·Ì… I2.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    If Not IsDate(invDate) Then
        MsgBox " «—ÌŒ «·›« Ê—… €Ì— ’ÕÌÕ. √œŒ·  «—ÌŒ ’ÕÌÕ ›Ì I2.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '4) —ﬁ„ «·›« Ê—… F2
    If Trim(CStr(ws.Range("F2").Value)) = "" Then
        MsgBox "«œŒ· —ﬁ„ «·›« Ê—….", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    '5) ›Õ’ »‰Êœ «·›« Ê—… (ﬂ„« ﬂ«‰ ⁄‰œﬂ)
    For r = 7 To 31
        If Trim(CStr(ws.Cells(r, "C").Value)) <> "" Then

            If Trim(CStr(ws.Cells(r, "F").Value)) = "" Then
                MsgBox "«Œ — «·ÊÕœ… (⁄œœ/ﬁÌ«”) («·”ÿ— " & r & ")", vbExclamation
                ValidateInvoiceForSave = False
                Exit Function
            End If

            If Val(ws.Cells(r, "G").Value) <= 0 Then
                MsgBox "«œŒ· ⁄œœ ’ÕÌÕ √ﬂ»— „‰ ’›— («·”ÿ— " & r & ")", vbExclamation
                ValidateInvoiceForSave = False
                Exit Function
            End If

            If ws.Cells(r, "F").Value = "ﬁÌ«”" Then
                If Val(ws.Cells(r, "D").Value) <= 0 Or Val(ws.Cells(r, "E").Value) <= 0 Then
                    MsgBox "›Ì Õ«·… (ﬁÌ«”) ÌÃ» ≈œŒ«· «·⁄—÷ Ê«·«— ›«⁄ («·”ÿ— " & r & ")", vbExclamation
                    ValidateInvoiceForSave = False
                    Exit Function
                End If
            End If

            If Val(ws.Cells(r, "I").Value) <= 0 Then
                MsgBox "«·”⁄— ·«“„ ÌﬂÊ‰ √ﬂ»— „‰ ’›— («·”ÿ— " & r & ")", vbExclamation
                ValidateInvoiceForSave = False
                Exit Function
            End If
        End If
    Next r

    '6) ‘Ì  «·⁄„Ì· ·«“„ ÌﬂÊ‰ „ÊÃÊœ (·√‰‰« ·‰ ‰‰‘∆Â ⁄‰œ «·Õ›Ÿ)
    targetSheetName = SafeSheetName(customer)
    If SheetExists(targetSheetName) = False Then
        MsgBox "‘Ì  «·⁄„Ì· €Ì— „ÊÃÊœ »⁄œ." & vbCrLf & _
               "„‰ ›÷·ﬂ √÷› «·⁄„Ì· (·Ì „ ≈‰‘«¡ ‘Ì  «·⁄„Ì·) À„ √⁄œ «·Õ›Ÿ.", vbExclamation
        ValidateInvoiceForSave = False
        Exit Function
    End If

    ValidateInvoiceForSave = True
End Function

'========================
' Õ›Ÿ/ —ÕÌ· «·›« Ê—… ≈·Ï ‘Ì  «·⁄„Ì·
' - ·« Ì‰‘∆ ⁄„Ì·/‘Ì 
' - Ì—Õ¯· ›ﬁÿ »⁄œ «· Õﬁﬁ
'========================
Public Sub AddInvoice()

    Dim wsI As Worksheet, wsTarget As Worksheet
    Dim customer As String, invNo As String
    Dim invDate As Date
    Dim r As Long, lastTarget As Long
    Dim targetSheetName As String
    Dim oldVis As XlSheetVisibility

    If ValidateInvoiceForSave = False Then Exit Sub

    Set wsI = ThisWorkbook.Worksheets("≈œŒ«·_›« Ê—…")

    customer = Trim(CStr(wsI.Range("B2").Value))
    invNo = Trim(CStr(wsI.Range("F2").Value))
    invDate = CDate(wsI.Range("I2").Value)

    targetSheetName = SafeSheetName(customer)
    Set wsTarget = ThisWorkbook.Worksheets(targetSheetName)

    '›ﬂ ﬁ›· »‰Ì… «·„’‰› „ƒﬁ «
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo 0

    '«› Õ «·‘Ì  „ƒﬁ « ·· —ÕÌ· À„ «—Ã⁄Â ﬂ„« ﬂ«‰
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

    '≈⁄«œ… «·ﬁ›·
    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    On Error GoTo 0

    ' ›—Ì€ »Ì«‰«  «·›« Ê—… »œÊ‰ „”Õ «·„⁄«œ·« 
    wsI.Range("F2").ClearContents
    wsI.Range("I2").ClearContents
    wsI.Range("B3:J3").ClearContents

    wsI.Range("C7:C31").ClearContents
    wsI.Range("D7:E31").ClearContents
    wsI.Range("F7:F31").ClearContents
    wsI.Range("G7:G31").ClearContents
    wsI.Range("I7:I31").ClearContents

    ' ›—Ì€ «·⁄„Ì· (B2) + ComboBox1 ≈‰ ÊÃœ
    wsI.Range("B2").ClearContents
    On Error Resume Next
    wsI.OLEObjects("ComboBox1").Object.Value = ""
    wsI.OLEObjects("ComboBox1").Object.Text = ""
    On Error GoTo 0

    MsgBox "?  „ Õ›Ÿ «·›« Ê—… Ê —ÕÌ·Â« ≈·Ï Õ”«» «·⁄„Ì·: " & customer, vbInformation
End Sub

'========================
' › Õ ﬂ‘› Õ”«» «·⁄„·«¡
'========================
Public Sub Open_Kashf_Hesab()
    On Error GoTo ErrH
    ThisWorkbook.Worksheets("ﬂ‘›_Õ”«»_«·⁄„·«¡").Activate
    Exit Sub
ErrH:
    MsgBox "‘Ì  (ﬂ‘›_Õ”«»_«·⁄„·«¡) €Ì— „ÊÃÊœ.  √ﬂœ „‰ «”„ «·‘Ì .", vbExclamation
End Sub

'› Õ ’›Õ… ≈œŒ«· «·›« Ê—…
Public Sub Open_InvoiceEntry()
    On Error GoTo ErrH
    ThisWorkbook.Worksheets("≈œŒ«·_›« Ê—…").Activate
    Exit Sub
ErrH:
    MsgBox "‘Ì  (≈œŒ«·_›« Ê—…) €Ì— „ÊÃÊœ.  √ﬂœ „‰ «·«”„.", vbExclamation
End Sub

'========================
' ≈÷«›… ⁄„Ì· ÃœÌœ (Ì÷Ì› «·«”„ ›Ì ﬁ«∆„…_⁄„·«¡ + Ì‰‘∆ ‘Ì  «·⁄„Ì· „Œ›Ì)
'========================
Public Sub Add_New_Customer()

    Dim ws As Worksheet
    Dim newName As String
    Dim lastRow As Long
    Dim nm As String

    On Error GoTo ErrH
    Set ws = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")

    newName = InputBox("«ﬂ » «”„ «·⁄„Ì· «·ÃœÌœ:", "≈÷«›… ⁄„Ì·")
    newName = Trim(newName)
    If newName = "" Then Exit Sub

    nm = SafeSheetName(newName)
    If nm = "" Then
        MsgBox "«”„ «·⁄„Ì· €Ì— ’«·Õ.", vbExclamation
        Exit Sub
    End If

    If CustomerExistsInList(newName) Then
        MsgBox "Â–« «·⁄„Ì· „ÊÃÊœ »«·›⁄· ›Ì «·ﬁ«∆„….", vbExclamation
        Exit Sub
    End If

    If SheetExists(nm) Then
        MsgBox "ÌÊÃœ ‘Ì  »‰›” «”„ «·⁄„Ì· »«·›⁄·. «Œ — «”„ „Œ ·›.", vbExclamation
        Exit Sub
    End If

    '›ﬂ «·ﬁ›· „ƒﬁ «
    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo ErrH

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If lastRow < 2 Then lastRow = 2

    ws.Cells(lastRow, "A").Value = newName

    '≈‰‘«¡ ‘Ì  «·⁄„Ì· („Œ›Ì)
    CreateCustomerSheet nm
    
    


    ' ”ÃÌ· «”„ «·‘Ì  ›Ì C («Œ Ì«—Ì)
    ws.Cells(lastRow, "C").Value = nm

    '≈⁄«œ… «·ﬁ›·
    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    On Error GoTo 0

    ws.Activate
    ws.Cells(lastRow, "A").Select
    MsgBox "?  „  ≈÷«›… «·⁄„Ì· »‰Ã«Õ: " & newName, vbInformation
    Exit Sub

ErrH:
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ ≈÷«›… «·⁄„Ì·.  √ﬂœ „‰ ÊÃÊœ ‘Ì  (ﬁ«∆„…_⁄„·«¡) Ê√‰ ﬂ·„… «·„—Ê— ’ÕÌÕ….", vbExclamation
    
    
End Sub

'========================
' › Õ Õ”«» «·⁄„Ì· „ƒﬁ « (ÌŸÂ— «·‘Ì  À„ ÌıŒ›Ï ⁄‰œ «·Œ—ÊÃ ñ ⁄»— ThisWorkbook)
'========================
Public Sub OpenCustomerSheet2()
    Dim shName As String
    Dim wsCust As Worksheet

    shName = Trim(ThisWorkbook.Worksheets("ﬂ‘›_Õ”«»_«·⁄„·«¡").Range("B2").Value)
    If shName = "" Then
        MsgBox "«Œ — «”„ «·⁄„Ì· √Ê·«", vbExclamation
        Exit Sub
    End If

    shName = SafeSheetName(shName)

    If SheetExists(shName) = False Then
        MsgBox "‘Ì  «·⁄„Ì· €Ì— „ÊÃÊœ: " & shName, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo 0

    Set wsCust = ThisWorkbook.Worksheets(shName)
    wsCust.Visible = xlSheetVisible

    TempOpenedCustomerSheet = shName

    wsCust.Activate

    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub


