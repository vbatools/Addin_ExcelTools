VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsGroop 
   Caption         =   "Группировка данных:"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmMenedgerSheetsGroop.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsGroop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim i           As Long
    Dim iCount      As Long
    Dim shAct       As Worksheet

    Set shAct = ActiveSheet

    Dim sAddress    As String
    With Range(txtInputRng.Value)
        If optColumn.Value Then
            sAddress = VBA.Split(Columns(.Column).Address, ":")(0) & ":"
            sAddress = sAddress & VBA.Split(Columns(.Column + .Columns.Count - 1).Address, ":")(0)
        Else
            sAddress = .Row & ":" & .Row + .Rows.Count - 1
        End If
    End With

    Dim shVisible   As XlSheetVisibility
    Dim Sh          As Worksheet
    With frmMenedgerSheets.listSheets
        iCount = .ListCount - 1
        For i = 0 To iCount
            If .Selected(i) And .List(i, 4) = vbNullString And .List(i, 2) = "лист" Then
                Set Sh = ActiveWorkbook.Worksheets(.List(i, 1))
                shVisible = Sh.Visible
                Sh.Visible = XlSheetVisibility.xlSheetVisible
                Sh.Activate
                If optColumn.Value Then
                    With Sh.Columns(sAddress)
                        If optGroop.Value Then
                            .Group
                        Else
                            .Ungroup
                        End If
                    End With
                Else
                    With Sh.Rows(sAddress)
                        If optGroop.Value Then
                            .Group
                        Else
                            .Ungroup
                        End If
                    End With
                End If
                Sh.Visible = shVisible
            End If
        Next i
    End With
    shAct.Activate
    Unload Me
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
