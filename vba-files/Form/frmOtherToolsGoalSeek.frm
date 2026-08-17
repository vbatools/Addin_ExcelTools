VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherToolsGoalSeek 
   Caption         =   "Собр данных со сквозных листов:"
   ClientHeight    =   2475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10395
   OleObjectBlob   =   "frmOtherToolsGoalSeek.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtherToolsGoalSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    If txtSetCells.Value = vbNullString Or txtValues.Value = vbNullString Or txtChangeCells.Value = vbNullString Then
        Call MsgBox("Не все параметры выбраны", vbCritical)
        Exit Sub
    End If

    Dim rngSetCells As Range
    Dim rngValues   As Range
    Dim rngChangeCells As Range
    Dim errMsg      As String
    Dim iCount      As Long

    Set rngSetCells = Range(txtSetCells.Value)
    Set rngValues = Range(txtValues.Value)
    Set rngChangeCells = Range(txtChangeCells.Value)

    iCount = rngSetCells.Cells.Count

    If iCount <> rngValues.Cells.Count Then
        errMsg = "значения"
    End If

    If iCount <> rngChangeCells.Cells.Count Then
        If errMsg <> vbNullString Then errMsg = errMsg & vbNewLine
        errMsg = errMsg & "изменяемые ячейки"
    End If

    If errMsg <> vbNullString Then
        Call MsgBox("Не равное количество элементов:" & vbNewLine & errMsg, vbCritical)
        Exit Sub
    End If

    Dim i           As Long

    For i = 1 To iCount
        If Left(rngSetCells.Cells(i).formula, 1) = "=" And IsNumeric(rngValues.Cells(i).Value) Then
            rngSetCells.Cells(i).GoalSeek Goal:=rngValues.Cells(i).Value, ChangingCell:=rngChangeCells.Cells(i)
        End If
    Next i

    Me.Hide
End Sub

Private Sub txtSetCells_DropButtonClick()
    Me.Hide
    txtSetCells.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtSetCells_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtValues_DropButtonClick()
    Me.Hide
    txtValues.Value = SelectRangeViaDialog(, , False)
    Me.Show
End Sub

Private Sub txtChangeCells_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtChangeCells_DropButtonClick()
    Me.Hide
    txtChangeCells.Value = SelectRangeViaDialog(, , False)
    Me.Show
End Sub

Private Sub txtValues_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    Call ConfigureDropButton(txtSetCells)
    Call ConfigureDropButton(txtValues)
    Call ConfigureDropButton(txtChangeCells)
End Sub
