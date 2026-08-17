VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataToCSV 
   Caption         =   "Создать CSV из диапазона:"
   ClientHeight    =   1935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataToCSV.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataToCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'--------------------------------------------------------------------------------
' Event: btnCancel_Click
' Purpose: Обработчик закрытия формы без сохранения изменений
'--------------------------------------------------------------------------------
Private Sub btnCancel_Click()
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Event: btnOK_Click
' Purpose: Основная процедура формирования CSV из выбранного диапазона с порционной записью в файл
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()
    Dim arrData     As Variant
    Dim sDelimiter  As String
    Dim sFullNameFile As String
    Dim ws          As Worksheet
    Dim rngData     As Range

    ' Проверка и инициализация разделителя
    sDelimiter = txtDelimiter.Value
    If sDelimiter = vbNullString Then
        Call MsgBox("Не задан символ разделитель", vbCritical)
        Exit Sub
    End If

    ' Проверка корректности введенного диапазона
    On Error Resume Next
    Set ws = ActiveSheet
    Set rngData = ws.Range(txtInputRng.Value)
    On Error GoTo 0

    If rngData Is Nothing Then
        Call MsgBox("Указан некорректный диапазон ячеек", vbCritical)
        Exit Sub
    End If

    arrData = rngData.Value2
    ' Если выбрана одна ячейка, Value2 возвращает не массив, а одно значение
    If Not IsArray(arrData) Then
        Call MsgBox("Выбрана одна ячейка. Укажите диапазон.", vbCritical)
        Exit Sub
    End If

    Dim i           As Long
    Dim j           As Long
    Dim iCount      As Long
    Dim jCount      As Long
    Dim sCSV        As String
    Dim sLine       As String
    Dim lCharCount  As Long

    iCount = UBound(arrData, 1)
    jCount = UBound(arrData, 2)

    ' Формирование пути к файлу
    sFullNameFile = ActiveWorkbook.Path & Application.PathSeparator & ActiveWorkbook.Name & "_" & ws.Name & ".csv"

 ' Удаление старого файла, если он существует
    If Dir(sFullNameFile, vbNormal) <> vbNullString Then
        On Error Resume Next
        Call Kill(sFullNameFile)
        If Err.Number <> 0 Then
            Call MsgBox("Не удалось удалить старый файл. Возможно, он открыт в другой программе." & vbCrLf & "Ошибка: " & Err.Description, vbCritical)
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
    End If

    ' Циклическая обработка массива данных
    For i = 1 To iCount
        sLine = vbNullString
        For j = 1 To jCount
            If sLine <> vbNullString Then sLine = sLine & sDelimiter
            ' Преобразование значения в строку и передача в функцию экранирования
            sLine = sLine & EscapeCsvValue(CStr(arrData(i, j)), sDelimiter)
        Next j

        ' Добавление переноса строки перед новой записью (кроме первой)
        If i <> 1 Then
            sCSV = sCSV & vbNewLine
            lCharCount = lCharCount + VBA.Len(vbNewLine)
        End If

        sCSV = sCSV & sLine
        lCharCount = lCharCount + VBA.Len(sLine)

        ' Проверка лимита символов и запись в файл
        If lCharCount >= 50000 Then
            If Not TXTAddIntoTXTFile(sFullNameFile, sCSV, True) Then
                Call MsgBox("Ошибка при записи в файл. Возможно, файл открыт в другой программе.", vbCritical)
                Exit Sub
            End If
            sCSV = vbNullString
            lCharCount = 0
        End If
    Next i

    ' Запись остатков данных, если они остались после завершения цикла
    If sCSV <> vbNullString Then
        If Not TXTAddIntoTXTFile(sFullNameFile, sCSV, True) Then
            Call MsgBox("Ошибка при записи в файл. Возможно, файл открыт в другой программе.", vbCritical)
            Exit Sub
        End If
    End If

    ' Очистка памяти
    Erase arrData
    Call MsgBox("Сохранено в файл!", vbInformation)
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Function: EscapeCsvValue
' Purpose: Экранирование значения строки для безопасной записи в CSV формат
' Parameters:
' Value - Исходное значение (String)
' Delimiter - Разделитель, используемый в CSV (String)
' Returns: String - Экранированное значение, соответствующее стандарту RFC 4180
'--------------------------------------------------------------------------------
Private Function EscapeCsvValue(ByVal Value As String, ByVal Delimiter As String) As String

    ' Дублирование двойных кавычек, если они присутствуют в значении
    If InStr(1, Value, """") > 0 Then
        Value = Replace(Value, """", """""")
    End If

    ' Обертка значения в кавычки, если оно содержит разделитель,
    ' двойную кавычку или символ переноса строки
    If InStr(1, Value, Delimiter) > 0 Or InStr(1, Value, """") > 0 Or _
            InStr(1, Value, vbCrLf) > 0 Or InStr(1, Value, vbCr) > 0 Or _
            InStr(1, Value, vbLf) > 0 Then
        Value = """" & Value & """"
    End If
    EscapeCsvValue = Value
End Function

'--------------------------------------------------------------------------------
' Function: TXTAddIntoTXTFile
' Purpose: Запись или добавление текста в файл с обработкой ошибок
' Parameters:
' FileName - Полный путь к файлу
' txt - Текст добавляемый в файл
' AddFile - True (по умолчанию) для создания файла, если он не существует
' Returns: Boolean - True если запись прошла успешно, False в случае ошибки
'--------------------------------------------------------------------------------
Public Function TXTAddIntoTXTFile(ByVal FileName As String, ByVal txt As String, Optional AddFile As Boolean = True) As Boolean
    Dim FSO         As Object
    Dim ts          As Object

    On Error GoTo ErrorHandler

    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' 8 = ForAppending
    Set ts = FSO.OpenTextFile(FileName, 8, AddFile)
    ts.Write txt
    ts.Close

    TXTAddIntoTXTFile = True
    GoTo CleanUp

ErrorHandler:
    TXTAddIntoTXTFile = False
    ' Запись ошибки в окно Immediate для отладки
    Debug.Print "Ошибка TXTAddIntoTXTFile: " & Err.Description

CleanUp:
    If Not ts Is Nothing Then Set ts = Nothing
    If Not FSO Is Nothing Then Set FSO = Nothing
End Function

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

