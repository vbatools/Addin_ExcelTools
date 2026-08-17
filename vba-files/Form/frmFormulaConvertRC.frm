VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaConvertRC 
   Caption         =   "Изменение ссылок в формуле:"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   OleObjectBlob   =   "frmFormulaConvertRC.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaConvertRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmFormulaConvertRC - изменение ссылок в формулах
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   11-06-2026 13:00:52
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub btnCancel_Click()
    Unload Me
End Sub




Private Sub btnOK_Click()
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Не выбран диапазон данных!", vbCritical)
        Exit Sub
    End If

    Dim rng         As Range
    Dim arrData     As Variant
    Application.ReferenceStyle = xlA1
    Set rng = Selection
    arrData = rng.formula

    If Not IsArray(arrData) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrData
        arrData = arr
    End If

    Dim ReferenceType As XlReferenceType


    Select Case True
        Case optColumn.Value
            ReferenceType = xlRelRowAbsColumn
        Case optRow.Value
            ReferenceType = xlAbsRowRelColumn
        Case optColumnRow.Value
            ReferenceType = xlAbsolute
        Case optNone.Value
            ReferenceType = xlRelative
    End Select

    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long
    iCount = UBound(arrData, 1)
    jCount = UBound(arrData, 2)

    For i = 1 To iCount
        For j = 1 To jCount
            If VBA.Left$(arrData(i, j), 1) = "=" Then arrData(i, j) = Application.ConvertFormula(arrData(i, j), xlA1, xlA1, ReferenceType)
        Next j
    Next i

    rng.Value2 = arrData

    Unload Me

End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
End Sub
