VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherToolsDeletePasswordWB 
   Caption         =   "Удаление паролей с книги:"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13515
   OleObjectBlob   =   "frmOtherToolsDeletePasswordWB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtherToolsDeletePasswordWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    lbValue.Caption = -1
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)

    Dim wb          As Workbook
    For Each wb In Workbooks
        cmbMain.AddItem wb.Name
    Next
    cmbMain.Value = ActiveWorkbook.Name
    Call loadDataToList
End Sub

Private Sub cmbMain_Change()
    Call loadDataToList
End Sub

Private Sub loadDataToList()
    Dim wb          As Workbook
    Dim Sh          As Worksheet
    Dim i           As Long
    Dim bHaveShProtect As Boolean
    Set wb = Workbooks(cmbMain.Value)
    lbMsg.Visible = wb.ProtectStructure

    With listMain
        .Clear
        For Each Sh In wb.Worksheets
            .AddItem Sh.Name
            .List(i, 1) = "нет"
            If Sh.ProtectContents Then
                .List(i, 1) = "есть"
                bHaveShProtect = True
            End If

            i = i + 1
        Next Sh
    End With
    btnOK.Enabled = (bHaveShProtect Or lbMsg.Visible)
End Sub
