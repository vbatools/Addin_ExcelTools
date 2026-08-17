VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataMergeText 
   Caption         =   "Объединение ячеек с сохранением текста:"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataMergeText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataMergeText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmDataMergeText - объединение значений ячеек без потери данных
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   11-06-2026 10:14:25
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub btnCancel_Click()
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Procedure: btnOK_Click
' Purpose: Объединяет ячейки выделенного диапазона с сохранением текста
'          Текст объединяется через указанный разделитель
' Parameters: нет
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()

    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Не выбран диапазон данных!", vbCritical)
        Exit Sub
    End If

    If ActiveSheet.ProtectContents Then
        Call MsgBox("Лист [" & ActiveSheet.Name & "] - защищен от изменений, снимите пароль!", vbCritical)
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim rng         As Range
    Dim arr         As Variant
    Dim arrVal()    As String
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long
    Dim sDelimiter  As String
    Dim iStepMerg   As Integer
    Dim k           As Long

    Set rng = Selection
    arr = rng.Value2

    If Not IsArray(arr) Then
        Call MsgBox("Выбрана одна ячейка!", vbCritical)
        Exit Sub
    End If

    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)

    Select Case True
        Case optSpace.Value
            sDelimiter = " "
        Case optAltEnter.Value
            sDelimiter = vbNewLine
        Case optChr.Value
            sDelimiter = txtChr.Value
    End Select
    iStepMerg = 1
    If txtStepMerge.Value <> vbNullString Then iStepMerg = VBA.CInt(txtStepMerge.Value)

    Call DisableApplicationSettings
    Call SaveUndoInfo(rng, False, True)

    Dim sVal        As String
    If optRow.Value Then
        ReDim arrVal(1 To iCount, 1 To 1) As String
        If iStepMerg > iCount Then iStepMerg = iCount
        For i = 1 To iCount
            For j = 1 To jCount
                If sVal <> vbNullString Then sVal = sVal & sDelimiter
                sVal = sVal & arr(i, j)
            Next j
            If i Mod iStepMerg = 0 Then
                k = i - iStepMerg + 1
                arrVal(k, 1) = sVal
                sVal = vbNullString
                rng.Rows(k & ":" & i).Merge Across:=False
            End If
        Next i
        If sVal <> vbNullString Then
            k = k + iStepMerg
            arrVal(k, 1) = sVal
            rng.Rows(k & ":" & iCount).Merge Across:=False
        End If
    Else
        ReDim arrVal(1 To 1, 1 To jCount) As String
        If iStepMerg > jCount Then iStepMerg = jCount
        For j = 1 To jCount
            For i = 1 To iCount
                If sVal <> vbNullString Then sVal = sVal & sDelimiter
                sVal = sVal & arr(i, j)
            Next i
            If j Mod iStepMerg = 0 Then
                k = j - iStepMerg + 1
                arrVal(1, k) = sVal
                sVal = vbNullString
                rng.Columns(Split(Columns(k).Address, ":")(0) & ":" & Split(Columns(j).Address, ":")(0)).Merge Across:=False
            End If
        Next j
        If sVal <> vbNullString Then
            k = k + iStepMerg
            arrVal(1, k) = sVal
            rng.Columns(Split(Columns(k).Address, ":")(0) & ":" & Split(Columns(jCount).Address, ":")(0)).Merge Across:=False
        End If
    End If

    rng.Value2 = arrVal

    Call RestoreApplicationSettings
    Application.OnUndo "Отменить", "RestoreUndoInfo"
    Unload Me
    Exit Sub

ErrorHandler:
    Call RestoreApplicationSettings
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка"
End Sub

Private Sub txtStepMerge_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
End Sub
