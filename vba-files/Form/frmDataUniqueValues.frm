VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataUniqueValues 
   Caption         =   "Список уникальных значений:"
   ClientHeight    =   3495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataUniqueValues.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataUniqueValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmDataUniqueValues - Получение уникальных строк
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   10-06-2026 09:35:59
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    ' --- Проверка ввода ---
    Dim sAddress    As String
    Dim sAddressInsert As String
    Dim sColumns    As String
    Dim sMessage    As String

    sAddress = Trim$(txtInputRng.Value)
    sAddressInsert = Trim$(txtOutputRng.Value)
    sColumns = VBA.Replace(txtNumberColumn.Value, " ", vbNullString)

    ' Накопление ошибок
    If sAddress = vbNullString Then
        sMessage = "Не выбран диапазон данных!"
    End If

    If sAddressInsert = vbNullString Then
        If sMessage <> vbNullString Then sMessage = sMessage & vbCrLf
        sMessage = sMessage & "Не выбран диапазон вставки данных!"
    End If

    If sMessage <> vbNullString Then
        MsgBox sMessage, vbCritical, "Ошибка"
        Exit Sub
    End If

    Dim rngInsert   As Range
    Set rngInsert = Range(sAddressInsert)

    If rngInsert.Parent.ProtectContents Then
        Call MsgBox("Лист [" & rngInsert.Parent.Name & "] - защищен от изменений, снимите пароль!", vbCritical)
        Exit Sub
    End If

    ' --- Подготовка данных ---
    Dim arrData     As Variant
    Dim arrColumns  As Variant

    On Error Resume Next
    arrData = Range(sAddress).Value2
    On Error GoTo 0

    If Not IsArray(arrData) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrData
        arrData = arr
    End If

    ' Разбор номеров столбцов
    arrColumns = VBA.Split(sColumns, ",")

    ' --- Проверка номеров столбцов ---
    Dim lRowCount   As Long
    Dim lColCount   As Long
    Dim i           As Long

    lRowCount = UBound(arrData, 1)
    lColCount = UBound(arrData, 2)

    For i = LBound(arrColumns) To UBound(arrColumns)
        If arrColumns(i) <> vbNullString Then
            If lColCount < CLng(arrColumns(i)) Then
                MsgBox "Номер столбца [" & arrColumns(i) & "] превышает количество столбцов в таблице!", vbCritical, "Ошибка"
                Exit Sub
            End If
        End If
    Next i

    ' --- Получение уникальных строк ---
    Dim dict        As Object
    Dim arrResult() As Variant
    Dim sKey        As String
    Dim lUniqueCount As Long
    Dim lCol        As Long
    Dim lKeyCol     As Long

    ReDim arrResult(1 To lRowCount, 1 To lColCount)

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For i = 1 To lRowCount
        ' Формирование составного ключа
        sKey = vbNullString
        For lKeyCol = LBound(arrColumns) To UBound(arrColumns)
            If arrColumns(lKeyCol) <> vbNullString Then
                If lKeyCol > LBound(arrColumns) Then sKey = sKey & "|"
                sKey = sKey & arrData(i, CLng(arrColumns(lKeyCol)))
            End If
        Next lKeyCol

        ' Проверка уникальности
        If Not dict.Exists(sKey) Then
            lUniqueCount = lUniqueCount + 1
            dict.Add sKey, i

            ' Копирование строки
            For lCol = 1 To lColCount
                arrResult(lUniqueCount, lCol) = arrData(i, lCol)
            Next lCol
        End If
    Next i

    Call DisableApplicationSettings
    ' --- Выделение уникальных строк цветом ---
    If chbSelectColor.Value And lUniqueCount > 0 Then
        Call HighlightUniqueRows(dict, sAddress)
    End If

    Set dict = Nothing

    ' --- Вывод результата ---
    If lUniqueCount > 0 Then
        Call SaveUndoInfo(rngInsert.Cells(1, 1).Resize(lUniqueCount, lColCount), False, False)
        rngInsert.Cells(1, 1).Resize(lUniqueCount, lColCount).Value2 = arrResult
        rngInsert.Parent.Activate
        Application.OnUndo "Отменить", "RestoreUndoInfo"
    Else
        MsgBox "Уникальные строки не найдены!", vbExclamation, "Внимание"
    End If

    Call RestoreApplicationSettings
    Unload Me
End Sub

' ============================================
' Выделение уникальных строк цветом
' ============================================
Private Sub HighlightUniqueRows(ByRef dict As Object, ByVal sAddress As String)
    Dim rngSource   As Range
    Dim arrRows     As Variant
    Dim lFirstRow   As Long
    Dim lFirstCol   As Long
    Dim lColCount   As Long
    Dim i           As Long

    Set rngSource = Range(sAddress)

    With rngSource
        lFirstRow = .Row
        lFirstCol = .Column
        lColCount = .Columns.Count
    End With

    ' Получение номеров уникальных строк
    arrRows = dict.Items

    ' Выделение строк
    Call DisableApplicationSettings

    For i = LBound(arrRows) To UBound(arrRows)
        rngSource.Parent.Cells(lFirstRow + arrRows(i) - 1, lFirstCol) _
                .Resize(1, lColCount).Interior.Color = 5296274
    Next i

    Call RestoreApplicationSettings
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtOutputRng_DropButtonClick()
    Me.Hide
    txtOutputRng.Value = SelectRangeViaDialog(, , False)
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtOutputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtNumberColumn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Разрешаем: цифры (48-57), запятую (44), Backspace (8)
    ' Точка (46) заменяется на запятую
    Select Case KeyAscii
        Case 48 To 57, 44, 8
            ' Разрешено
        Case 46
            KeyAscii = 44  ' Замена точки на запятую
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)

    ' Инициализация поля ввода адресом выделенного диапазона
    If TypeName(Selection) = "Range" Then
        txtInputRng.Value = Selection.Address
    End If

    Call ConfigureDropButton(txtInputRng)
    Call ConfigureDropButton(txtOutputRng)
End Sub

