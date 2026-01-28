Attribute VB_Name = "Module3"
Option Explicit

Sub Build_ActiveX_Buttons()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("ÅÏÎÇá_İÇÊæÑÉ")

    'ÍĞİ ÃÒÑÇÑ ŞÏíãÉ ÈäİÓ ÇáÃÓãÇÁ
    DeleteBtnIfExists ws, "btnSave"
    DeleteBtnIfExists ws, "btnClear"
    DeleteBtnIfExists ws, "btnAddCustomer"
    DeleteBtnIfExists ws, "btnOpenKashf"
    DeleteBtnIfExists ws, "btnLock"
    DeleteBtnIfExists ws, "btnUnlock"

    'ÅäÔÇÁ ÇáÃÒÑÇÑ (ÃáæÇä RGB ãÖãæäÉ)
    AddBtn ws, "btnSave", "ÍİÙ ÇáİÇÊæÑÉ", 420, 30, 170, 40, RGB(0, 160, 0), True
    AddBtn ws, "btnClear", "ÊİÑíÛ ÇáİÇÊæÑÉ", 240, 30, 170, 40, RGB(255, 140, 0), True
    AddBtn ws, "btnAddCustomer", "ÅÖÇİÉ Úãíá", 60, 30, 170, 40, RGB(0, 120, 215), True
    AddBtn ws, "btnOpenKashf", "İÊÍ ßÔİ ÍÓÇÈ ÇáÚãáÇÁ", 60, 80, 350, 40, RGB(128, 0, 128), True
    AddBtn ws, "btnLock", "Şİá Çáãáİ", 420, 80, 170, 40, RGB(220, 0, 0), True
    AddBtn ws, "btnUnlock", "İÊÍ Çáãáİ", 420, 130, 170, 40, RGB(0, 170, 255), True

    MsgBox "? Êã ÅäÔÇÁ ÃÒÑÇÑ ActiveX ÈäÌÇÍ.", vbInformation

End Sub


Private Sub AddBtn( _
    ws As Worksheet, _
    btnName As String, _
    cap As String, _
    L As Double, _
    T As Double, _
    W As Double, _
    H As Double, _
    backClr As Long, _
    boldFont As Boolean _
)

    Dim obj As OLEObject
    Set obj = ws.OLEObjects.Add(ClassType:="Forms.CommandButton.1", _
                                Left:=L, Top:=T, Width:=W, Height:=H)

    obj.Name = btnName

    With obj.Object
        .caption = cap
        .BackColor = backClr
        .ForeColor = vbWhite
        .Font.Name = "Arial"
        .Font.Size = 14
        .Font.Bold = boldFont
        .WordWrap = True
    End With

End Sub


Private Sub DeleteBtnIfExists(ws As Worksheet, btnName As String)
    On Error Resume Next
    ws.OLEObjects(btnName).Delete
    On Error GoTo 0
End Sub




Public Function CustomerHasAccount(ByVal customerSheetName As String) As Boolean
    Dim ws As Worksheet
    Dim bal As Double
    Dim lastRow As Long
    Dim rng As Range
    Dim c As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(customerSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        CustomerHasAccount = False
        Exit Function
    End If

    '1) ãÚíÇÑ ÇáÑÕíÏ (K4)
    If IsNumeric(ws.Range("K4").Value) Then
        bal = CDbl(ws.Range("K4").Value)
        If Abs(bal) > 0.00001 Then
            CustomerHasAccount = True
            Exit Function
        End If
    End If

    '2) ãÚíÇÑ ÇáÍÑßÉ ÇáÍŞíŞí:
    'ÇÚÊÈÑ ÇáÍÑßÉ ãæÌæÏÉ İŞØ ÅĞÇ æõÌÏ ÑŞã/äÕ İÚáí İí ÚãæÏ A ãä Õİ 7 Åáì ÂÎÑ ÇáÔíÊ
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 7 Then
        CustomerHasAccount = False
        Exit Function
    End If

    Set rng = ws.Range("A7:A" & lastRow)

    For Each c In rng
        If Trim(CStr(c.Value)) <> "" Then
            'áæ İíåÇ ŞíãÉ ÍŞíŞíÉ (ÑŞã İÇÊæÑÉ/ÃãÑ ÔÛá)
            CustomerHasAccount = True
            Exit Function
        End If
    Next c

    CustomerHasAccount = False
End Function

