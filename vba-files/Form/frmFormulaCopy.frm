VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaCopy 
   Caption         =   "Точное копирование формул:"
   ClientHeight    =   2085
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmFormulaCopy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmFormulaCopy- add description!
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   15-06-2026 14:40:40
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub btnOK_Click()

    If txtFormulsRng.Value = vbNullString Or txtInputRng.Value = vbNullString Then
        Call MsgBox("Не выбран диапазон!", vbCritical)
        Exit Sub
    End If

    Dim rngPaste    As Range
    Set rngPaste = Range(txtInputRng.Value)
    If rngPaste.Parent.ProtectContents Then
        Call MsgBox("Лист [" & rngPaste.Parent.Name & "] - защищен от изменений, снимите пароль!", vbCritical)
        Exit Sub
    End If

    Application.ReferenceStyle = xlA1

    Dim arrData     As Variant
    arrData = Range(txtFormulsRng.Value).formula
    If Not IsArray(arrData) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrData
        arrData = arr
    End If

    rngPaste.Cells(1, 1).Resize(UBound(arrData, 1), UBound(arrData, 2)).formula = arrData
    rngPaste.Parent.Activate
    Unload Me
End Sub

Private Sub txtFormulsRng_DropButtonClick()
    Me.Hide
    txtFormulsRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog(, , False)
    Me.Show
End Sub

Private Sub txtFormulsRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    If TypeName(Selection) = "Range" Then txtFormulsRng.Value = Selection.Address
    Call ConfigureDropButton(txtFormulsRng)
    Call ConfigureDropButton(txtInputRng)
End Sub
