VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsFreezePane 
   Caption         =   "Закрепить область:"
   ClientHeight    =   3735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmMenedgerSheetsFreezePane.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsFreezePane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOK_Click()

    Dim sAddress    As String
    sAddress = txtInputRng.Value

    Dim shAct       As Worksheet

    Set shAct = ActiveSheet

    Dim i           As Long
    Dim iCount      As Long
    Dim Sh          As Worksheet
    Dim shVisible   As XlSheetVisibility
    Dim rng         As Range

    With frmMenedgerSheets.listSheets
        iCount = .ListCount - 1
        For i = 0 To iCount
            If .Selected(i) And .List(i, 4) = vbNullString And .List(i, 2) = "лист" Then
                Set Sh = ActiveWorkbook.Worksheets(.List(i, 1))
                shVisible = Sh.Visible
                Sh.Visible = XlSheetVisibility.xlSheetVisible
                Sh.Activate
                Set rng = Sh.Range(sAddress)
                rng.Select

                ActiveWindow.FreezePanes = False
                With ActiveWindow
                    .SplitColumn = 0
                    .SplitRow = 0
                End With

                Select Case True
                    Case optArea.Value
                        ActiveWindow.FreezePanes = True
                    Case optColumn.Value
                        With ActiveWindow
                            .SplitColumn = rng.Column - 1
                            .SplitRow = 0
                        End With
                        ActiveWindow.FreezePanes = True
                    Case optRow.Value
                        With ActiveWindow
                            .SplitColumn = 0
                            .SplitRow = rng.Row - 1
                        End With
                        ActiveWindow.FreezePanes = True
                End Select

                Sh.Visible = shVisible
            End If
        Next i
    End With
    shAct.Activate
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog(, False)
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    If TypeName(Selection) = "Range" Then txtInputRng.Value = Selection.Address
    Call ConfigureDropButton(txtInputRng)
End Sub
