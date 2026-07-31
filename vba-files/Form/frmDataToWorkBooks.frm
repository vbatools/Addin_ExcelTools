VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataToWorkBooks 
   Caption         =   "загрузка данных в сквозные книги:"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataToWorkBooks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataToWorkBooks"
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

' ============================================================================
' КОНСТАНТЫ
' ============================================================================
Private Const MIN_COLUMNS As Long = 5
Private Const PATH_COL As Long = 1
Private Const FILENAME_COL As Long = 2
Private Const SHEETNAME_COL As Long = 3
Private Const TARGET_RANGE_COL As Long = 4

' ============================================================================
' ОБРАБОТЧИКИ СОБЫТИЙ
' ============================================================================

Private Sub btnCancel_Click()
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Procedure: btnOK_Click
' Purpose: Главная процедура — валидация, перенос данных в целевые файлы
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()
    Dim sAddress    As String
    Dim arrData     As Variant
    Dim excelApp    As Excel.Application
    Dim targetBook  As Workbook
    Dim i As Long, iCount As Long, batchSize As Long
    Dim sPath As String, sFormula As String

    ' --- Валидация ---
    sAddress = Trim$(txtInputRng.Value)
    If sAddress = vbNullString Then
        MsgBox "Не выбран диапазон данных!", vbCritical, "Ошибка"
        Exit Sub
    End If

    arrData = Range(sAddress).Value2
    If UBound(arrData, 2) < MIN_COLUMNS Then
        MsgBox "Недостаточно колонок в диапазоне!", vbCritical, "Ошибка"
        Exit Sub
    End If

    If Not arrData(1, TARGET_RANGE_COL) Like "$*$*[1-9]:$*$*[1-9]" Then
        MsgBox "Некорректный формат целевого диапазона!", vbCritical, "Ошибка"
        Exit Sub
    End If

    ' --- Инициализация ---
    Call DisableApplicationSettings

    Set excelApp = New Excel.Application
    excelApp.DisplayAlerts = False
    excelApp.Visible = False

    iCount = UBound(arrData, 1)
    batchSize = Range(arrData(1, TARGET_RANGE_COL)).Rows.Count
    sFormula = "='" & ActiveWorkbook.Path & Application.PathSeparator & _
            "[" & ActiveWorkbook.Name & "]" & ActiveSheet.Name & "'!"

    ' --- Основной цикл ---
    For i = 1 To iCount Step batchSize
        sPath = arrData(i, PATH_COL) & arrData(i, FILENAME_COL)

        ' Открываем файл при смене пути
        If Not FileHave(sPath, vbNormal) Then
            ' Файл не существует — пропуск записи
        ElseIf targetBook Is Nothing Then
            Set targetBook = excelApp.Workbooks.Open(sPath, False, False)
        ElseIf targetBook.FullName <> sPath Then
            targetBook.Close True
            Set targetBook = excelApp.Workbooks.Open(sPath, False, False)
        End If

        ' Запись данных
        If Not targetBook Is Nothing Then
            If HaveSheetInFile(targetBook, arrData(i, SHEETNAME_COL)) Then
                With targetBook.Worksheets(arrData(i, SHEETNAME_COL)).Range(arrData(i, TARGET_RANGE_COL))
                    .formula = sFormula & Cells(.Row + (i - 1) * batchSize, .Column + 4).Address(RowAbsolute:=False, columnAbsolute:=False)
                    If optOnlyValues.Value Then .Value2 = .Value2
                End With
            End If
        End If

        Call UpdateProgress(i / iCount)
    Next i

    ' --- Завершение ---
    If Not targetBook Is Nothing Then targetBook.Close True
    If Not excelApp Is Nothing Then
        excelApp.Quit
        Set excelApp = Nothing
    End If

    Call RestoreApplicationSettings
    Unload Me
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

'--------------------------------------------------------------------------------
' Procedure: UserForm_Initialize
' Purpose: Инициализация формы: центрирование и настройка полей ввода
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)

    If TypeName(Selection) = "Range" Then
        txtInputRng.Value = Selection.Address
    End If

    Call ConfigureDropButton(txtInputRng)
End Sub

'--------------------------------------------------------------------------------
' Procedure: UpdateProgress
' Purpose: Обновляет индикатор прогресса выполнения
' Parameters:
' pct - доля выполненной работы (от 0 до 1)
'--------------------------------------------------------------------------------
Private Sub UpdateProgress(ByVal pct As Single)
    On Error Resume Next
    With Me
        .Progress.Caption = format(pct, "0%")
        .Progress.Width = pct * .FrameProgress.Width
        .Repaint
    End With
End Sub



