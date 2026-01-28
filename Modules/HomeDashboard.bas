Attribute VB_Name = "HomeDashboard"
Option Explicit

Public Const HOME_SHEET As String = "«·—∆Ì”Ì…"

'==============================
'  ‘€Ì· „—… Ê«Õœ… ·≈‰‘«¡ «·—∆Ì”Ì… + «·√“—«— + ≈Œ›«¡ »«ﬁÌ «·‘Ì « 
'==============================
Public Sub Setup_Home_Dashboard()
    Dim wsHome As Worksheet

    '≈‰‘«¡/ ÃÂÌ“ ‘Ì  «·—∆Ì”Ì…
    Set wsHome = GetOrCreateHomeSheet

    '«„”Õ √“—«— ﬁœÌ„… »«·‘Ì  «·—∆Ì”Ì… («Œ Ì«—Ì)
    DeleteAllShapes wsHome

    '≈‰‘«¡ «·√“—«— (Shapes) ›Ì «·‘Ì  «·—∆Ì”Ì…
    AddHomeButton wsHome, "›« Ê—… „»Ì⁄« ", "Go_Invoice", 30, 40, 220, 40
    AddHomeButton wsHome, "ﬁ«∆„… «·⁄„·«¡", "Go_CustomersList", 30, 90, 220, 40
    AddHomeButton wsHome, "≈÷«›… ⁄„Ì·", "Go_AddCustomer", 30, 140, 220, 40
    AddHomeButton wsHome, "ﬂ‘› Õ”«» ⁄„Ì·", "Go_Kashf", 30, 190, 220, 40
    AddHomeButton wsHome, "≈Ã„«·Ì «·„»Ì⁄« ", "Go_TotalSales", 30, 240, 220, 40
    AddHomeButton wsHome, "„·Œ’ Õ”«»«  «·⁄„·«¡", "Go_AccountsSummary", 30, 290, 220, 40

    '«Œ Ì«—Ì: ﬁ›·/› Õ «·Õ„«Ì… „‰ «·—∆Ì”Ì…
    AddHomeButton wsHome, "ﬁ›· «·„·›", "Lock_All", 280, 40, 220, 40
    AddHomeButton wsHome, "› Õ «·„·›", "Unlock_All", 280, 90, 220, 40

    '≈Œ›«¡ ﬂ· «·‘Ì «  „« ⁄œ« «·—∆Ì”Ì…
    HideAllSheetsExceptHome

    '≈÷«›… “— —ÃÊ⁄ ·ﬂ· «·‘Ì « 
    AddBackButtonToAllSheets

    wsHome.Activate
    MsgBox "?  „  ÃÂÌ“ ‘Ì  «·—∆Ì”Ì… Ê≈Œ›«¡ »«ﬁÌ «·‘Ì «  + ≈÷«›… “— «·—ÃÊ⁄.", vbInformation
End Sub

'==============================
' √“—«— «·—∆Ì”Ì… ( › Õ «·‘Ì  «·„ÿ·Ê»)
'==============================
Public Sub Go_Invoice()
    ShowOnlySheet "≈œŒ«·_›« Ê—…"
End Sub

Public Sub Go_CustomersList()
    ShowOnlySheet "ﬁ«∆„…_⁄„·«¡"
End Sub


'========================================
' ≈÷«›… ⁄„Ì· ÃœÌœ »«·«”„ +  ÕœÌÀ «·„·Œ’ + «·—ÃÊ⁄ ··—∆Ì”Ì…
'========================================
Public Sub Add_New_Customer_ByName(ByVal newName As String)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim nm As String

    On Error GoTo ErrH

    newName = Trim(newName)
    If newName = "" Then Exit Sub

    Set ws = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")

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
        MsgBox "ÌÊÃœ ‘Ì  »‰›” «”„ «·⁄„Ì· »«·›⁄·.", vbExclamation
        Exit Sub
    End If

    '? „‰⁄  ‘€Ì· Worksheet_Change √À‰«¡ «·≈÷«›… „‰ «·ﬁ«∆„… «·—∆Ì”Ì…
    IsAddingCustomerFromMenu = True
    Application.EnableEvents = False

    '√Ê· ’› ›«÷Ì
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    If lastRow < 2 Then lastRow = 2

    '≈÷«›… «·«”„
    ws.Cells(lastRow, "A").Value = newName

    '≈‰‘«¡ ‘Ì  «·⁄„Ì· („Œ›Ì VeryHidden)
    CreateCustomerSheet nm

    ' ”ÃÌ· «”„ «·‘Ì  ›Ì C
    ws.Cells(lastRow, "C").Value = nm

CleanExit:
    Application.EnableEvents = True
    IsAddingCustomerFromMenu = False

    MsgBox "?  „  ≈÷«›… «·⁄„Ì· Ê≈‰‘«¡ «·‘Ì  »‰Ã«Õ: " & newName, vbInformation
    Exit Sub

ErrH:
    Application.EnableEvents = True
    IsAddingCustomerFromMenu = False
    MsgBox "ÕœÀ Œÿ√ √À‰«¡ ≈÷«›… «·⁄„Ì·: " & Err.Description, vbExclamation
End Sub
    

'========================================
' “— ≈÷«›… ⁄„Ì· „‰ «·‘Ì  «·—∆Ì”Ì (InputBox)
'========================================
Public Sub Go_AddCustomer()
    Dim nm As String
    nm = InputBox("«ﬂ » «”„ «·⁄„Ì· «·ÃœÌœ:", "≈÷«›… ⁄„Ì·")
    nm = Trim(nm)
    If nm = "" Then Exit Sub

    Add_New_Customer_ByName nm
End Sub



Public Sub Go_Kashf()
    ShowOnlySheet "ﬂ‘›_Õ”«»_«·⁄„·«¡"
End Sub

Public Sub Go_TotalSales()
    '·Ê √‰  ⁄«„· › Õ »«·»«”Ê—œ + ≈Œ›«¡  ·ﬁ«∆Ì: «” Œœ„ «·„«ﬂ—Ê «·„ÊÃÊœ ⁄‰œﬂ
    Open_TotalSales_Sheet
End Sub

Public Sub Go_AccountsSummary()
    Open_AccountsSummary_Sheet
End Sub

'==============================
' “— «·—ÃÊ⁄ (ÌŸÂ— «·—∆Ì”Ì… ÊÌŒ›Ì «·‘Ì  «·Õ«·Ì)
'==============================
Public Sub Back_To_Home()
    Dim wsCur As Worksheet
    Set wsCur = ActiveSheet

    '·«  Œ›ˆ «·—∆Ì”Ì…
    If wsCur.Name <> HOME_SHEET Then
        On Error Resume Next
        ThisWorkbook.Unprotect Password:=ADMIN_PWD
        wsCur.Visible = xlSheetVeryHidden
        ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
        On Error GoTo 0
    End If

    '«ŸÂ— «·—∆Ì”Ì… ›ﬁÿ
    HideAllSheetsExceptHome
    ThisWorkbook.Worksheets(HOME_SHEET).Activate
End Sub

'==============================
' ≈ŸÂ«— ‘Ì  Ê«Õœ ›ﬁÿ (ÊÌŒ›Ì »«ﬁÌ «·‘Ì « )
'==============================
Public Sub ShowOnlySheet(ByVal SheetName As String)

    Dim ws As Worksheet
    Dim wasProtected As Boolean

    If SheetExists(SheetName) = False Then
        MsgBox "«·‘Ì  €Ì— „ÊÃÊœ: " & SheetName, vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False

    'Â· Structure „ﬁ›Ê·…ø
    wasProtected = ThisWorkbook.ProtectStructure

    '›ﬂ «·ﬁ›· „ƒﬁ «
    If wasProtected Then
        ThisWorkbook.Unprotect Password:=ADMIN_PWD

        ' √ﬂœ ›⁄·Ì« √‰Â« « ›ﬂ 
        If ThisWorkbook.ProtectStructure Then
            Application.EnableEvents = True
            MsgBox "·« Ì„ﬂ‰  €ÌÌ— ≈ŸÂ«—/≈Œ›«¡ «·‘Ì «  ·√‰ »‰Ì… «·„’‰› „« “«·  „Õ„Ì…." & vbCrLf & _
                   " √ﬂœ „‰ ﬂ·„… «·„—Ê— √Ê ·« ÌÊÃœ ﬁ›· ÌœÊÌ »ﬂ·„… √Œ—Ï.", vbCritical
            Exit Sub
        End If
    End If

    '√ŸÂ— «·‘Ì  «·„ÿ·Ê»
    ThisWorkbook.Worksheets(SheetName).Visible = xlSheetVisible

    '›⁄¯·Â √Ê·« Õ Ï ·«  Õ«Ê· ≈Œ›«¡ «·‘Ì  «·‰‘ÿ »«·€·ÿ
    ThisWorkbook.Worksheets(SheetName).Activate

    '«Œ›ˆ »«ﬁÌ «·‘Ì «  (·«  Œ›ˆ «·‘Ì  «·‰‘ÿ)
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SheetName Then
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    '≈⁄«œ… «·ﬁ›·
    If wasProtected Then
        ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    End If

    Application.EnableEvents = True
End Sub


'==============================
' ≈Œ›«¡ ﬂ· «·‘Ì «  „« ⁄œ« «·—∆Ì”Ì…
'==============================
Public Sub HideAllSheetsExceptHome()
    Dim ws As Worksheet

    On Error Resume Next
    ThisWorkbook.Unprotect Password:=ADMIN_PWD
    On Error GoTo 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = HOME_SHEET Then
            ws.Visible = xlSheetVisible
        Else
            ws.Visible = xlSheetVeryHidden
        End If
    Next ws

    On Error Resume Next
    ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
    On Error GoTo 0
End Sub

'==============================
' ≈÷«›… “— —ÃÊ⁄ ··—∆Ì”Ì… ›Ì ﬂ· «·‘Ì «  (Õ Ï ·«   ÊÂ)
'==============================
Public Sub AddBackButtonToAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> HOME_SHEET Then
            AddBackButton ws
        End If
    Next ws
End Sub

Private Sub AddBackButton(ByVal ws As Worksheet)
    'ÌÕ–› “— —ÃÊ⁄ ﬁœÌ„ À„ Ì÷Ì› ÃœÌœ
    On Error Resume Next
    ws.Shapes("btnBackHome").Delete
    On Error GoTo 0

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 10, 10, 160, 32)
    shp.Name = "btnBackHome"
    shp.TextFrame2.TextRange.Text = "? —ÃÊ⁄ ··—∆Ì”Ì…"
    shp.TextFrame2.TextRange.Font.Size = 12
    shp.TextFrame2.TextRange.Font.Bold = msoTrue
    shp.OnAction = "Back_To_Home"
End Sub

'==============================
' √œÊ«  „”«⁄œ… ··‘Ì  «·—∆Ì”Ì…
'==============================
Private Function GetOrCreateHomeSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOME_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = HOME_SHEET
    End If

    ws.Cells.Clear
    ws.Range("A1").Value = "·ÊÕ… «· Õﬂ„ - «·—∆Ì”Ì…"
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
    MsgBox "?  „ ≈⁄«œ…  ›⁄Ì· Events Ê≈⁄œ«œ«  Excel.", vbInformation
End Sub

