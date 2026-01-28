VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "æÑŞÉ2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private PrevName As String

'äÎÒä ÇáÇÓã ÇáŞÏíã ÚäÏ ÊÍÏíÏ ÇáÎáíÉ ŞÈá ÊÚÏíáåÇ
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.CountLarge > 1 Then Exit Sub
    If Not Intersect(Target, Me.Range("A2:A10000")) Is Nothing Then
        PrevName = Trim(CStr(Target.Value))
    Else
        PrevName = ""
    End If
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo CleanExit

    'ÇÔÊÛá İŞØ Úáì ÇáÚãæÏ A
    If Target.CountLarge > 1 Then GoTo CleanExit
    If Intersect(Target, Me.Range("A2:A10000")) Is Nothing Then GoTo CleanExit

    If IsAddingCustomerFromMenu Then GoTo CleanExit
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim r As Long
    Dim newName As String
    Dim safeSheet As String
    Dim oldSheet As String
    Dim wasProtected As Boolean

    r = Target.Row
    newName = Trim(CStr(Target.Value))
    If newName = "" Then GoTo CleanExit

    safeSheet = SafeSheetName(newName)
    oldSheet = Trim(CStr(Me.Cells(r, "C").Value))

    '--- İß Şİá ÈäíÉ ÇáãÕäİ ãÄŞÊğÇ (ãİÊÇÍ ÇáÔÛá İí æÖÚ ÇáŞİá) ---
    wasProtected = ThisWorkbook.ProtectStructure

    If wasProtected Then
        On Error Resume Next
        ThisWorkbook.Unprotect Password:=ADMIN_PWD
        On Error GoTo CleanExit

        If ThisWorkbook.ProtectStructure Then
            MsgBox "áÇ íãßä ÅÚÇÏÉ ÊÓãíÉ ÇáÔíÊ ÃËäÇÁ ÇáŞİá áÃä ÈäíÉ ÇáãÕäİ ãÇ ÒÇáÊ ãÍãíÉ." & vbCrLf & _
                   "ÊÃßÏ Ãä ÍãÇíÉ ÇáãÕäİ ÈßáãÉ mina2040 (æáÇ íæÌÏ Protect Workbook íÏæí ÈßáãÉ ÃÎÑì).", vbCritical
            GoTo CleanExit
        End If
    End If
    '------------------------------------------------------------

    'áæ Úãíá ÌÏíÏ (áÇ íæÌÏ ÔíÊ ŞÏíã)
    If oldSheet = "" Then

        If SheetExists(safeSheet) Then
            MsgBox "åĞÇ ÇáÇÓã ãæÌæÏ ÈÇáİÚá ßÔíÊ. ÇÎÊÑ ÇÓã ãÎÊáİ.", vbExclamation
            GoTo ReProtect
        End If

        CreateCustomerSheet safeSheet
        Me.Cells(r, "C").Value = safeSheet

    Else
        'Úãíá ãæÌæÏ: ÃÚÏ ÊÓãíÉ ÇáÔíÊ ÇáŞÏíã
        If SheetExists(oldSheet) Then

            '? ãäÚ ÊÚÏíá ÇáÇÓã áæ Úáíå ÍÓÇÈ/ÍÑßÉ
            If CustomerHasAccount(oldSheet) Then
                MsgBox "áÇ íãßä ÊÚÏíá ÇÓã ÇáÚãíá áÃäå Úáíå ÍÓÇÈ/ÍÑßÉ." & vbCrLf & _
                       "Şã ÈÊÕİíÉ ÇáÍÓÇÈ ÃæáÇğ Ëã ÃÚÏ ÇáãÍÇæáÉ.", vbExclamation

                'ÇÑÌÇÚ ÇáÇÓã ÇáŞÏíã ÏÇÎá ÇáÎáíÉ
                Target.Value = PrevName
                GoTo ReProtect
            End If

            If SheetExists(safeSheet) And safeSheet <> oldSheet Then
                MsgBox "áÇ íãßä ÊÛííÑ ÇáÇÓã áÃä ÔíÊ ÈÇÓã (" & safeSheet & ") ãæÌæÏ ÈÇáİÚá.", vbExclamation
                Target.Value = PrevName
                GoTo ReProtect
            End If

            ThisWorkbook.Worksheets(oldSheet).Name = safeSheet
            Me.Cells(r, "C").Value = safeSheet

        Else
            'áæ ÇáÔíÊ ÇáŞÏíã ÇÊÍĞİ ÈÇáÛáØ: ÃäÔÆ ÌÏíÏ
            If SheetExists(safeSheet) Then
                MsgBox "åĞÇ ÇáÇÓã ãæÌæÏ ÈÇáİÚá ßÔíÊ. ÇÎÊÑ ÇÓã ãÎÊáİ.", vbExclamation
                Target.Value = PrevName
                GoTo ReProtect
            End If

            CreateCustomerSheet safeSheet
            Me.Cells(r, "C").Value = safeSheet
        End If
    End If

ReProtect:
    'ÅÚÇÏÉ Şİá ÈäíÉ ÇáãÕäİ ßãÇ ßÇäÊ
    If wasProtected Then
        On Error Resume Next
        ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
        On Error GoTo 0
    End If

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


