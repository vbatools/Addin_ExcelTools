VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataToJSON 
   Caption         =   "Создать JSON из диапазона:"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataToJSON.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataToJSON"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' UserForm     :   frmDataToSheets - Модуль загрузки данных в сквозные листы
' Author       :   VBATools
' Copyright    :   Apache License
' Created      :   10-06-2026 09:36:16
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Const QUOTE_CHAR As String = """"
Private Const QUOTE_CHAR_CON As String = """: "

'--------------------------------------------------------------------------------
' Event: btnCancel_Click
' Purpose: Обработчик закрытия формы без сохранения изменений
'--------------------------------------------------------------------------------
Private Sub btnCancel_Click()
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Event: btnOK_Click
' Purpose: Основная процедура формирования JSON из выбранного диапазона
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()
    If txtInputRng.Value = vbNullString Then Exit Sub

    Dim arrData     As Variant
    Dim i           As Long
    Dim j           As Long
    Dim iCount      As Long
    Dim jCount      As Long
    Dim fItem       As Long
    Dim k           As Long
    Dim sHeaders    As String
    Dim sItem       As String
    Dim sJSON       As String
    Dim arrVal()    As String

    ' Получение данных из диапазона
    arrData = Range(txtInputRng.Value).Value2

    ' Обработка случая одиночной ячейки
    If Not IsArray(arrData) Then
        If arrData = vbNullString Then
            Call MsgBox("Выбрана пустая ячейка!", vbCritical)
            Exit Sub
        End If
        ReDim arrData(1 To 1, 1 To 1)
        arrData(1, 1) = arrData
    End If

    iCount = UBound(arrData, 1)
    jCount = UBound(arrData, 2)

    ' Формирование заголовков (если есть)
    If chbHaveTitle.Value Then
        ReDim arrVal(1 To jCount) As String
        fItem = 2    ' Данные начинаются со второй строки

        sHeaders = vbTab & QUOTE_CHAR & "headers" & QUOTE_CHAR_CON & "["
        For j = 1 To jCount
            If VBA.Replace(arrData(1, j), " ", vbNullString) = vbNullString Then
                Call MsgBox("В заголовке не может быть пустых значений", vbCritical)
                Exit Sub
            End If
            ' Экранируем только заголовки
            arrVal(j) = QUOTE_CHAR & EscapeJSON(CStr(arrData(1, j))) & QUOTE_CHAR
        Next j
        sHeaders = sHeaders & VBA.Join(arrVal, ", ")

        ' Пересоздаем массив для строк данных
        ReDim arrVal(1 To iCount - 1) As String
    Else
        fItem = 1    ' Данные с первой строки
        ReDim arrVal(1 To iCount) As String
    End If

    ' Формирование массива строк данных
    For i = fItem To iCount
        sItem = vbNullString

        For j = 1 To jCount
            If sItem <> vbNullString Then sItem = sItem & ", "

            If optTypeJSONObject.Value Then
                ' Формат: "Ключ": "Значение"
                sItem = sItem & QUOTE_CHAR & EscapeJSON(CStr(arrData(1, j))) & QUOTE_CHAR_CON
                sItem = sItem & GetJSONFormattedValue(arrData(i, j))
            Else
                ' Формат: ["Значение1", "Значение2"]
                sItem = sItem & GetJSONFormattedValue(arrData(i, j))
            End If
        Next j

        k = k + 1
        If optTypeJSONObject.Value Then
            arrVal(k) = vbTab & vbTab & "{" & sItem & "}"
        Else
            arrVal(k) = vbTab & vbTab & "[" & sItem & "]"
        End If
    Next i

    ' Сборка итогового JSON
    sJSON = "{" & vbNewLine
    sJSON = sJSON & vbTab & QUOTE_CHAR & "title" & QUOTE_CHAR_CON & QUOTE_CHAR & ActiveSheet.Name & QUOTE_CHAR & ", " & vbNewLine

    If chbHaveTitle.Value Then
        sJSON = sJSON & sHeaders & "], " & vbNewLine
    End If

    sJSON = sJSON & vbTab & QUOTE_CHAR & "data" & QUOTE_CHAR_CON & "[" & vbNewLine
    sJSON = sJSON & VBA.Join(arrVal, ", " & vbNewLine) & vbNewLine
    sJSON = sJSON & vbTab & "]" & vbNewLine & "}"

    If optCopyBufer.Value Then
        Dim ClipBoard As New DataObject
        With ClipBoard
            .SetText sJSON
            .PutInClipboard
        End With
        Call MsgBox("Скопировано в буфер обмена!", vbInformation)
    Else
        If saveTextToFile(sJSON, ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & "_" & ActiveSheet.Name & ".json", "utf-8") Then
            Call MsgBox("Сохранено в файл!", vbInformation)
        Else
            Call MsgBox("Не удалось сохранить в файл!", vbCritical)
        End If
    End If
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Event: txtInputRng_DropButtonClick
' Purpose: Скрытие формы, вызов диалога выбора диапазона и повторный показ
'--------------------------------------------------------------------------------
Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

'--------------------------------------------------------------------------------
' Event: txtInputRng_KeyDown
' Purpose: Ограничение навигационных клавиш в поле ввода
'--------------------------------------------------------------------------------
Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

'--------------------------------------------------------------------------------
' Event: UserForm_Initialize
' Purpose: Инициализация формы при запуске (центрирование, заполнение полей)
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
' Function: EscapeJSON
' Purpose: Экранирует специальные символы для корректного представления в JSON
' Parameters:
'   sText - Исходная строка
' Returns: String - Экранированная строка
'--------------------------------------------------------------------------------
Private Function EscapeJSON(ByVal sText As String) As String
    If Len(Trim$(sText)) = 0 Then
        EscapeJSON = vbNullString
        Exit Function
    End If

    Dim sResult     As String
    sResult = sText

    ' Порядок замены важен: сначала обратный слеш, чтобы не экранировать его дважды
    sResult = Replace(sResult, "\", "\\")
    sResult = Replace(sResult, QUOTE_CHAR, "\" & QUOTE_CHAR)
    sResult = Replace(sResult, vbCr, "\r")
    sResult = Replace(sResult, vbLf, "\n")
    sResult = Replace(sResult, vbTab, "\t")
    sResult = Replace(sResult, vbBack, "\b")
    sResult = Replace(sResult, vbFormFeed, "\f")
    sResult = Replace(sResult, vbNullChar, "\u0000")

    ' Разделители строк Unicode
    sResult = Replace(sResult, ChrW(&H2028), "\u2028")
    sResult = Replace(sResult, ChrW(&H2029), "\u2029")

    EscapeJSON = sResult
End Function

'--------------------------------------------------------------------------------
' Function: GetJSONFormattedValue
' Purpose: Возвращает строковое представление значения в формате JSON
' Parameters:
'   vValue - Значение ячейки (Variant)
' Returns: String - Форматированная строка (в кавычках, null, число и т.д.)
'--------------------------------------------------------------------------------
Private Function GetJSONFormattedValue(ByVal vValue As Variant) As String
    Select Case TypeName(vValue)
        Case "String"
            GetJSONFormattedValue = QUOTE_CHAR & EscapeJSON(vValue) & QUOTE_CHAR
        Case "Boolean"
            GetJSONFormattedValue = VBA.LCase$(vValue)
        Case "Empty"
            GetJSONFormattedValue = "null"
        Case Else
            ' Проверка на дату и число выполняется для Variant типа Double/String, содержащего дату/число
            If IsDate(vValue) Then
                GetJSONFormattedValue = QUOTE_CHAR & vValue & QUOTE_CHAR
            ElseIf IsNumeric(vValue) Then
                GetJSONFormattedValue = VBA.Replace(vValue, ",", ".")
            Else
                ' Fallback для прочих типов
                GetJSONFormattedValue = QUOTE_CHAR & EscapeJSON(CStr(vValue)) & QUOTE_CHAR
            End If
    End Select
End Function
