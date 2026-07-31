VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataFromSheets 
   Caption         =   "Собр данных со сквозных листов:"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataFromSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataFromSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmDataFromSheets - Модуль сбора данных из сквозных листов
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   10-06-2026 09:37:02
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    On Error GoTo ErrorHandler

    ' --- Проверка диапазона ---
    Dim sAddress    As String
    sAddress = Trim$(txtInputRng.Value)

    If sAddress = vbNullString Then
        MsgBox "Не выбран диапазон данных!", vbCritical, "Ошибка"
        Exit Sub
    End If

    ' --- Определение целевого листа ---
    Dim wsTarget    As Worksheet

    If chAddNewSheet.Value Then
        Set wsTarget = ActiveWorkbook.Worksheets.Add(before:=ActiveWorkbook.Worksheets(1))
        wsTarget.Cells(1, 1).Select
    Else
        Set wsTarget = ActiveSheet
    End If

    ' --- Подготовка переменных ---
    Call DisableApplicationSettings

    Dim wsCurrent   As Worksheet
    Dim lRowCount   As Long
    Dim lRow        As Long
    Dim lCol        As Long
    Dim iShift      As Integer

    lRow = activeCell.Row
    lCol = activeCell.Column
    lRowCount = Range(sAddress).Rows.Count

    ' --- Обход листов и сбор данных ---
    For Each wsCurrent In ActiveWorkbook.Worksheets
        ' Пропуск текущего листа
        If wsCurrent.Name <> ActiveSheet.Name Then
            ' Проверка скрытых листов
            If chDontLoolHidenSheets.Value And wsCurrent.Visible <> xlSheetVisible Then
                ' Переход к следующему листу
            Else
                ' Запись имени листа и адреса диапазона
                If chAddSheetsNames.Value Then
                    wsTarget.Cells(lRow, lCol).Resize(lRowCount, 1).Value2 = wsCurrent.Name
                    wsTarget.Cells(lRow, lCol + 1).Resize(lRowCount, 1).Value2 = sAddress
                    iShift = 2
                Else
                    iShift = 0
                End If

                ' Копирование данных
                wsCurrent.Activate
                wsCurrent.Range(sAddress).Copy

                wsTarget.Activate
                wsTarget.Cells(lRow, lCol + iShift).Select

                ' Выбор метода вставки
                If optOnlyValues.Value Then
                    Selection.PasteSpecial Paste:=xlPasteValues
                ElseIf optOnlyFormuls.Value Then
                    Selection.PasteSpecial Paste:=xlPasteFormulas
                ElseIf optOnlyLinks.Value Then
                    ActiveSheet.Paste link:=True
                End If

                ' Смещение на следующую позицию
                lRow = lRow + lRowCount
            End If
        End If
    Next wsCurrent
    Call RestoreApplicationSettings
    Unload Me
    Exit Sub

ErrorHandler:
    Call RestoreApplicationSettings
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка"
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
