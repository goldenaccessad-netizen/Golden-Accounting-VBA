Attribute VB_Name = "WorkbookProtection"
Option Explicit

Public Function TryUnprotectWorkbook() As Boolean
    Dim wasProtected As Boolean

    wasProtected = ThisWorkbook.ProtectStructure
    If wasProtected Then
        On Error Resume Next
        ThisWorkbook.Unprotect Password:=ADMIN_PWD
        On Error GoTo 0
    End If

    TryUnprotectWorkbook = wasProtected
End Function

Public Sub RestoreProtectWorkbook(ByVal wasProtected As Boolean)
    If wasProtected Then
        On Error Resume Next
        ThisWorkbook.Protect Password:=ADMIN_PWD, Structure:=True, Windows:=False
        On Error GoTo 0
    End If
End Sub

Public Function TryUnprotectSheet(ByVal ws As Worksheet) As Boolean
    Dim wasProtected As Boolean

    If ws Is Nothing Then Exit Function

    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect Password:=ADMIN_PWD
        On Error GoTo 0
    End If

    TryUnprotectSheet = wasProtected
End Function

Public Sub RestoreProtectSheet(ByVal ws As Worksheet, ByVal wasProtected As Boolean, _
                               Optional ByVal allowFiltering As Boolean = False, _
                               Optional ByVal allowSorting As Boolean = False)
    If ws Is Nothing Then Exit Sub

    If wasProtected Then
        On Error Resume Next
        ws.Protect Password:=ADMIN_PWD, UserInterfaceOnly:=True, AllowFiltering:=allowFiltering, AllowSorting:=allowSorting
        ws.EnableSelection = xlUnlockedCells
        On Error GoTo 0
    End If
End Sub
