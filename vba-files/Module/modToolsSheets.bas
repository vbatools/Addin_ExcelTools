Attribute VB_Name = "modToolsSheets"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module       :   modToolsSheets - операции с листами книги
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   10-06-2026 09:24:42
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Enum TypeCaseString
    tpUCase
    tpLCase
    tpAsString
    tpAllWorldUCase
End Enum

'--------------------------------------------------------------------------------
' Function: caseString
' Purpose: Преобразование регистра строки в соответствии с указанным режимом
' Parameters:
'   - sName (String) - исходная строка
'   - TypeCase (TypeCaseString) - режим преобразования регистра
' Returns: String - преобразованная строка
' Remarks: Поддерживает четыре режима преобразования регистра
'--------------------------------------------------------------------------------

Public Function caseString(ByVal sName As String, ByVal TypeCase As TypeCaseString) As String
    Select Case TypeCase
        Case tpUCase
            caseString = VBA.UCase$(sName)
        Case tpLCase
            caseString = VBA.LCase$(sName)
        Case tpAsString
            caseString = VBA.UCase$(VBA.Left$(sName, 1)) & IIf(VBA.Len(sName) > 1, VBA.LCase$(VBA.Mid$(sName, 2, VBA.Len(sName) - 1)), vbNullString)
        Case tpAllWorldUCase
            caseString = Application.Proper(sName)
    End Select
End Function

'--------------------------------------------------------------------------------
' Sub: RenameSheets
' Purpose: Переименование листов в книге на основе значений из выбранного диапазона
' Parameters: Нет
' Returns: Ничего. Переименовывает листы в активной книге.
' Remarks:
'   - Количество ячеек в диапазоне должно соответствовать количеству листов
'   - Недопустимые символы автоматически заменяются на "_"
'   - Пустые ячейки вызывают ошибку и остановку выполнения
'--------------------------------------------------------------------------------
Public Sub RenameSheets()

    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлена защита структуры книги", vbCritical)
        Exit Sub
    End If

    Dim selectedRangeAddress As String
    Dim arrNames    As Variant
    Dim sheetIndex  As Long
    Dim sheetsCount As Long
    Dim newSheetName As String
    Dim renamedCount As Long

    Application.ReferenceStyle = xlA1

    selectedRangeAddress = SelectRangeViaDialog("Выберите список листов")

    If selectedRangeAddress = vbNullString Then Exit Sub

    On Error Resume Next
    arrNames = Range(selectedRangeAddress).Value2
    On Error GoTo 0

    If Not IsArray(arrNames) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrNames
        arrNames = arr
    End If

    sheetsCount = UBound(arrNames, 1)

    If sheetsCount <> ActiveWorkbook.Sheets.Count Then
        MsgBox "Количество элементов списка (" & sheetsCount & ")" & vbCrLf & _
                "не совпадает с количеством листов в книге (" & ActiveWorkbook.Sheets.Count & ")!", _
                vbCritical, "Ошибка"
        Exit Sub
    End If

    renamedCount = 0
    For sheetIndex = 1 To sheetsCount

        ' Получение нового имени
        newSheetName = CStr(arrNames(sheetIndex, 1))

        ' Проверка на пустое значение
        If Trim(newSheetName) = vbNullString Then
            MsgBox "Ячейка " & sheetIndex & " содержит пустое значение!", _
                    vbExclamation, "Ошибка"
            Exit Sub
        End If

        newSheetName = CleanSheetName(newSheetName)

        ' Переименование листа с обработкой ошибок
        On Error Resume Next
        ActiveWorkbook.Worksheets(sheetIndex).Name = newSheetName

        If Err.Number = 0 Then
            renamedCount = renamedCount + 1
        Else
            MsgBox "Не удалось переименовать лист " & sheetIndex & ":" & vbCrLf & _
                    Err.Description, vbExclamation, "Ошибка"
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0

    Next sheetIndex
End Sub

'--------------------------------------------------------------------------------
' Sub: AddSheetsByList
' Purpose: Создание новых листов на основе шаблона по списку названий
' Parameters: Нет
' Returns: Ничего. Создаёт листы в активной книге.
' Remarks:
'   - Сначала выбирается диапазон со списком названий
'   - Затем выбирается ячейка на листе-шаблоне
'   - Листы с существующими именами пропускаются
'   - При ошибке создания лист автоматически удаляется
'--------------------------------------------------------------------------------
Public Sub AddSheetsByList()
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлена защита структуры книги", vbCritical)
        Exit Sub
    End If
    
    Dim sListRange  As String
    Dim sTemplateSheet As String
    Dim arrNames    As Variant

    Dim vItem       As Variant
    Dim sSheetName  As String
    Dim wsStart     As Worksheet

    ' --- Выбор диапазона со списком названий ---
    sListRange = SelectRangeViaDialog()
    If sListRange = vbNullString Then Exit Sub

    ' --- Выбор ячейки на листе-шаблоне ---
    sTemplateSheet = SelectRangeViaDialog("Выберите ячейку на листе-шаблоне")
    If sTemplateSheet = vbNullString Then Exit Sub

    ' --- Извлечение имени листа-шаблона ---
    sTemplateSheet = ExtractSheetName(sTemplateSheet)

    ' --- Получение списка названий ---
    On Error Resume Next
    arrNames = Range(sListRange).Value2
    On Error GoTo 0

    If Not IsArray(arrNames) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrNames
        arrNames = arr
    End If

    ' --- Создание листов ---
    Call DisableApplicationSettings
    Set wsStart = ActiveSheet


    For Each vItem In arrNames
        sSheetName = CleanSheetName(CStr(vItem))
        ' Создание листа, если имя не пустое и лист не существует
        If sSheetName <> vbNullString Then
            If Not HaveSheetInFile(ActiveWorkbook, sSheetName) Then
                On Error Resume Next
                Worksheets(sTemplateSheet).Copy After:=Worksheets(Worksheets.Count)
                If Err.Number = 0 Then
                    ActiveSheet.Name = sSheetName
                Else
                    Application.DisplayAlerts = False
                    ActiveSheet.Delete
                    Application.DisplayAlerts = True
                End If
                On Error GoTo 0
            End If
        End If
    Next vItem

    ' --- Возврат на исходный лист ---
    wsStart.Activate
    Call RestoreApplicationSettings
End Sub

'--------------------------------------------------------------------------------
' Function: ExtractSheetName
' Purpose: Извлечение имени листа из полного адреса диапазона
' Parameters:
'   - sFullAddress (String) - полный адрес диапазона (например, "[Книга]Лист'!$A$1")
' Returns: String - имя листа
' Remarks:
'   - Удаляет имя книги в квадратных скобках
'   - Удаляет адрес ячейки после "'!"
'   - Удаляет начальный апостроф
'--------------------------------------------------------------------------------
Private Function ExtractSheetName(ByVal sFullAddress As String) As String
    ' Пример: "[Книга]Лист'!$A$1" или "Лист'!$A$1"
    Dim sTemp       As String

    ' Удаляем часть после "'!"
    sTemp = Split(sFullAddress, "'!")(0)

    ' Удаляем первый символ "'" если есть
    If Left$(sTemp, 1) = "'" Then
        sTemp = Mid$(sTemp, 2)
    End If

    ' Удаляем имя книги в квадратных скобках если есть
    If InStr(sTemp, "]") > 0 Then
        sTemp = Mid$(sTemp, InStr(sTemp, "]") + 1)
    End If

    ExtractSheetName = sTemp
End Function

'--------------------------------------------------------------------------------
' Sub: GetSheetsLists
' Purpose: Вывод списка имён всех листов книги в активную ячейку
' Parameters: Нет
' Returns: Ничего. Выводит список в рабочий лист.
' Remarks:
'   - Список выводится начиная с активной ячейки вниз по столбцу
'   - Поддерживает отмену действия (Ctrl+Z)
'--------------------------------------------------------------------------------
Public Sub GetSheetsLists()

    If ActiveSheet.ProtectContents Then
        Call MsgBox("Лист [" & ActiveSheet.Name & "] - защищен от изменений, снимите пароль!", vbCritical)
        Exit Sub
    End If

    Dim i           As Long
    Dim iCount      As Long
    iCount = ActiveWorkbook.Worksheets.Count
    ReDim arr(1 To ActiveWorkbook.Worksheets.Count, 1 To 1) As String

    For i = 1 To iCount
        arr(i, 1) = ActiveWorkbook.Worksheets(i).Name
    Next i
    Call SaveUndoInfo(activeCell.Resize(iCount, 1), False, False)
    activeCell.Resize(iCount, 1).Value2 = arr
    Application.OnUndo "Отменить", "RestoreUndoInfo"
End Sub
'--------------------------------------------------------------------------------
' Sub: SetSheetNameUCase
' Purpose: Преобразование имён всех листов в верхний регистр
' Parameters: Нет
' Returns: Ничего. Преобразует имена листов.
'--------------------------------------------------------------------------------
Public Sub SetSheetNameUCase()
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлена защита структуры книги", vbCritical)
        Exit Sub
    End If
    Dim Sh          As Worksheet
    For Each Sh In ActiveWorkbook.Worksheets
        Sh.Name = caseString(Sh.Name, tpUCase)
    Next Sh
End Sub
'--------------------------------------------------------------------------------
' Sub: SetSheetNameLCase
' Purpose: Преобразование имён всех листов в нижний регистр
' Parameters: Нет
' Returns: Ничего. Преобразует имена листов.
'--------------------------------------------------------------------------------
Public Sub SetSheetNameLCase()
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлена защита структуры книги", vbCritical)
        Exit Sub
    End If
    Dim Sh          As Worksheet
    For Each Sh In ActiveWorkbook.Worksheets
        Sh.Name = caseString(Sh.Name, tpLCase)
    Next Sh
End Sub
'--------------------------------------------------------------------------------
' Sub: SetSheetNameAsString
' Purpose: Преобразование имён листов: первая буква заглавная, остальные строчные
' Parameters: Нет
' Returns: Ничего. Преобразует имена листов.
'--------------------------------------------------------------------------------
Public Sub SetSheetNameAsString()
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлена защита структуры книги", vbCritical)
        Exit Sub
    End If
    Dim Sh          As Worksheet
    For Each Sh In ActiveWorkbook.Worksheets
        Sh.Name = caseString(Sh.Name, tpAsString)
    Next Sh
End Sub
'--------------------------------------------------------------------------------
' Sub: SetSheetNameAllWorldUCase
' Purpose: Преобразование имён листов: каждое слово с заглавной буквы
' Parameters: Нет
' Returns: Ничего. Преобразует имена листов.
'--------------------------------------------------------------------------------
Public Sub SetSheetNameAllWorldUCase()
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлена защита структуры книги", vbCritical)
        Exit Sub
    End If
    Dim Sh          As Worksheet
    For Each Sh In ActiveWorkbook.Worksheets
        Sh.Name = caseString(Sh.Name, tpAllWorldUCase)
    Next Sh
End Sub



