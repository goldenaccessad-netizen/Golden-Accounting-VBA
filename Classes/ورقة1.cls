VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ê—ﬁ…1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private IsUpdating As Boolean

'? ⁄‰œ «·ÊﬁÊ› œ«Œ· «·ﬂ„»Ê »Êﬂ”:
' - Ì„”Õ «·«”„
' - Ì› Õ «·ﬁ«∆„… ﬂ«„·…
Private Sub ComboBox1_GotFocus()
    IsUpdating = True
    Me.ComboBox1.Value = ""
    Me.ComboBox1.Text = ""
    Me.Range("B2").ClearContents
    IsUpdating = False

    FillCustomersCombo True
End Sub

'? »ÕÀ –ﬂÌ √À‰«¡ «·ﬂ «»…
Private Sub ComboBox1_Change()
    If IsUpdating Then Exit Sub
    FilterCustomersCombo
End Sub

'? ⁄‰œ «Œ Ì«— «”„ ⁄„Ì·
Private Sub ComboBox1_Click()
    If Me.ComboBox1.Value <> "" Then
        IsUpdating = True
        Me.Range("B2").Value = Me.ComboBox1.Value
        IsUpdating = False
    End If
End Sub

Private Sub ComboBox1_LostFocus()
    If Me.ComboBox1.Value <> "" Then
        IsUpdating = True
        Me.Range("B2").Value = Me.ComboBox1.Value
        IsUpdating = False
    End If
End Sub

'--------------------------
'  ⁄»∆… «·ﬁ«∆„… ﬂ«„·…
'--------------------------
Private Sub FillCustomersCombo(Optional ByVal ShowDropDown As Boolean = False)
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long

    Set ws = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")

    IsUpdating = True
    Me.ComboBox1.Clear

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 2 To lastRow
        If Trim(ws.Cells(i, "A").Value) <> "" Then
            Me.ComboBox1.AddItem ws.Cells(i, "A").Value
        End If
    Next i

    If ShowDropDown Then Me.ComboBox1.DropDown
    IsUpdating = False
End Sub

'--------------------------
' ›· —… Õ”» «·„ﬂ Ê»
'--------------------------
Private Sub FilterCustomersCombo()
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim txt As String

    Set ws = ThisWorkbook.Worksheets("ﬁ«∆„…_⁄„·«¡")
    txt = Me.ComboBox1.Text

    IsUpdating = True
    Me.ComboBox1.Clear

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If Trim(txt) = "" Then
        For i = 2 To lastRow
            If Trim(ws.Cells(i, "A").Value) <> "" Then
                Me.ComboBox1.AddItem ws.Cells(i, "A").Value
            End If
        Next i
    Else
        For i = 2 To lastRow
            If InStr(1, ws.Cells(i, "A").Value, txt, vbTextCompare) > 0 Then
                Me.ComboBox1.AddItem ws.Cells(i, "A").Value
            End If
        Next i
    End If

    Me.ComboBox1.Text = txt
    Me.ComboBox1.SelStart = Len(txt)
    Me.ComboBox1.DropDown

    IsUpdating = False
End Sub



Private Sub btnSave_Click()
    Call AddInvoice
End Sub

Private Sub btnClear_Click()
    Call Clear_Invoice_Without_Save
End Sub

Private Sub btnAddCustomer_Click()
    Call Add_New_Customer
End Sub

Private Sub btnOpenKashf_Click()
    Call Open_Kashf_Hesab
End Sub

Private Sub btnLock_Click()
    Call Lock_All
End Sub

Private Sub btnUnlock_Click()
    Call Unlock_All
End Sub

