VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataToSheets 
   Caption         =   "загрузка данных в сквозные листы:"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataToSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataToSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmDataToSheets - Модуль загрузки данных в сквозные листы
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   10-06-2026 09:36:16
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    ' --- Проверка диапазона ---
    Dim sAddress    As String
    sAddress = Trim$(txtInputRng.Value)

    If sAddress = vbNullString Then
        MsgBox "Не выбран диапазон данных!", vbCritical, "Ошибка"
        Exit Sub
    End If

    ' --- Подготовка переменных ---
    Call DisableApplicationSettings

    Dim wsTarget    As Worksheet
    Dim rngSource   As Range
    Dim rngTarget   As Range
    Dim arrData     As Variant
    Dim lRow        As Long
    Dim lCol        As Long
    Dim lColCount   As Long
    Dim lRowCount   As Long
    Dim i           As Long
    Dim wsCurrent   As Worksheet

    Set wsTarget = ActiveSheet

    ' --- Получение исходного диапазона ---
    On Error Resume Next
    Set rngSource = Range(sAddress)
    On Error GoTo 0

    If rngSource Is Nothing Then
        MsgBox "Некорректный диапазон!", vbCritical, "Ошибка"
        GoTo CleanUp
    End If

    ' --- Чтение данных ---
    lRow = rngSource.Row
    lCol = rngSource.Column
    lColCount = rngSource.Columns.Count + lCol - 1
    arrData = rngSource.Value2

    ' --- Проверка целевого диапазона ---
    On Error Resume Next
    Set rngTarget = Range(arrData(1, 2))
    On Error GoTo 0

    If Not IsArray(arrData) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrData
        arrData = arr
    End If

    ' --- Обработка данных ---
    lRowCount = rngTarget.Rows.Count

    For i = 1 To UBound(arrData, 1) Step lRowCount
        ' Проверка существования листа
        On Error Resume Next
        Set wsCurrent = Worksheets(arrData(i, 1))
        On Error GoTo 0

        If Not wsCurrent Is Nothing Then
            With wsCurrent.Range(arrData(1, 2))
                If optOnlyLinks.Value Then
                    ' Создание ссылок на исходные данные
                    .formula = "='" & wsTarget.Name & "'!" & _
                            wsTarget.Cells(lRow + i - 1, lCol + 2).Address(RowAbsolute:=False, columnAbsolute:=False)
                Else
                    ' Копирование значений
                    On Error Resume Next
                    .Value2 = wsTarget.Range( _
                            wsTarget.Cells(lRow + i - 1, lCol + 2), _
                            wsTarget.Cells(lRow + i + lRowCount - 2, lColCount)).Value2
                    If Err.Number <> 0 Then
                        Call MsgBox(Err.Description, vbCritical)
                        GoTo CleanUp
                    End If
                    On Error GoTo 0
                End If
            End With
        End If

        Set wsCurrent = Nothing
    Next i
    Call RestoreApplicationSettings
    Unload Me
    Exit Sub

CleanUp:
    Call RestoreApplicationSettings
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)

    ' Инициализация поля ввода адресом выделенного диапазона
    If TypeName(Selection) = "Range" Then
        txtInputRng.Value = Selection.Address
    End If

    Call ConfigureDropButton(txtInputRng)
End Sub




