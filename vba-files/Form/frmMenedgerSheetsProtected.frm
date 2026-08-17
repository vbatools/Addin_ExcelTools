VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsProtected 
   Caption         =   "Защита выделенных листов книги:"
   ClientHeight    =   8880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmMenedgerSheetsProtected.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsProtected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmSheetsMenedgerProtected- add description!
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   18-06-2026 09:08:21
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Const LOCK_ICON     As Long = 60379
Const LOCK_ICON_COLOR As Long = &HC0C0FF
Const UNLOCK_ICON   As Long = 62471
Const UNLOCK_ICON_COLOR As Long = &HC0FFC0

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSetPass_Click()
    Dim i           As Long
    Dim iCount      As Long
    Dim Sh          As Object
    With frmMenedgerSheets.listSheets
        iCount = .ListCount - 1
        For i = 0 To iCount
            If .Selected(i) And .List(i, 4) <> "защита" Then
                .List(i, 4) = "защита"
                Set Sh = ActiveWorkbook.Sheets(.List(i, 1))
                Select Case TypeName(Sh)
                    Case "Chart"
                        Sh.Protect Password:=txtPasword.TEXT, DrawingObjects:=Not chDrawing.Value, Contents:=True, Scenarios:=chScenarios.Value
                    Case Else
                        Sh.Protect Password:=txtPasword.TEXT, DrawingObjects:=Not chDrawing.Value, Contents:=True, Scenarios:=chScenarios.Value, AllowFormattingCells:=chFormatCells.Value, _
                                AllowFormattingColumns:=chFormattingColumns.Value, AllowFormattingRows:=chFormattingRows.Value, AllowInsertingColumns:=chInsertingColumns.Value, AllowInsertingRows:=chInsertingRows.Value, _
                                AllowInsertingHyperlinks:=chInsertingHyperlinks.Value, AllowDeletingColumns:=chDeletingColumns.Value, AllowDeletingRows:=chDeletingRows.Value, AllowSorting:=chSort.Value, AllowFiltering:=chAutoFilter.Value, AllowUsingPivotTables:=chUsingPivotTables.Value
                        If (Not chBlockCells.Value) And chUnBlockCells.Value Then Sh.EnableSelection = xlUnlockedCells
                        If (Not chBlockCells.Value) And (Not chUnBlockCells.Value) Then Sh.EnableSelection = xlNoSelection
                        If TypeName(Sh) <> "DialogSheet" Then
                            If chBlockCells.Value And chUnBlockCells.Value And Sh.Type <> 3 And Sh.Type <> 4 Then Sh.EnableSelection = xlNoRestrictions
                        Else
                            If chBlockCells.Value And chUnBlockCells.Value Then Sh.EnableSelection = xlNoRestrictions
                        End If
                End Select
            End If
        Next i
    End With
    Unload Me
End Sub

Private Sub btnUnSetPass_Click()
    Dim i           As Long
    Dim iCount      As Long
    With frmMenedgerSheets.listSheets
        iCount = .ListCount - 1
        For i = 0 To iCount
            If .Selected(i) And .List(i, 4) = "защита" Then
                On Error Resume Next
                ActiveWorkbook.Sheets(.List(i, 1)).Unprotect Password:=txtPasword.Value
                Select Case Err.Number
                    Case 0
                    Case 1004
                        Call MsgBox("Лист: " & ActiveWorkbook.Sheets(.List(i, 1)).Name & vbNewLine & Err.Description, vbCritical, "Ошибка:")
                        Exit Sub
                    Case Else
                        Call MsgBox(Err.Number & vbNewLine & Err.Description, vbCritical, "Ошибка:")
                        Exit Sub
                End Select
                .List(i, 4) = vbNullString
            End If
        Next i
    End With
    Unload Me
End Sub

Private Sub lbChangeChrPass_Click()
    If lbChangeChrPass.Caption = VBA.ChrW$(LOCK_ICON) Then
        lbChangeChrPass.Caption = VBA.ChrW$(UNLOCK_ICON)
        lbChangeChrPass.ForeColor = UNLOCK_ICON_COLOR
        txtPasword.PasswordChar = vbNullChar
    Else
        lbChangeChrPass.Caption = VBA.ChrW$(LOCK_ICON)
        lbChangeChrPass.ForeColor = LOCK_ICON_COLOR
        txtPasword.PasswordChar = "*"
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    lbChangeChrPass.Caption = VBA.ChrW$(LOCK_ICON)
    lbChangeChrPass.ForeColor = LOCK_ICON_COLOR
End Sub

