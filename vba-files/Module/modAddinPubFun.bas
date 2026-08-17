Attribute VB_Name = "modAddinPubFun"
Option Explicit
Option Private Module
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module       :   modAddinPubFun - утилиты для работы с книгами, листами и элементами управления
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   10-06-2026 09:29:15
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

#If VBA7 Then
    ' For 64-bit systems
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Dim stateCalculation As XlCalculation
Dim stateReferenceStyle As XlReferenceStyle

Public Sub DisableApplicationSettings()
    If Workbooks.Count = 0 Then Exit Sub
    With Application
        .ScreenUpdating = False
        stateCalculation = .Calculation
        stateReferenceStyle = .ReferenceStyle
        .ReferenceStyle = xlA1
        .Calculation = xlCalculationManual
        .DisplayAlerts = False
    End With
End Sub

Public Sub RestoreApplicationSettings()
    If Workbooks.Count = 0 Then Exit Sub
    With Application
        .Calculation = stateCalculation
        .ReferenceStyle = stateReferenceStyle
        .DisplayAlerts = True
        .StatusBar = False
        .ScreenUpdating = True
    End With
End Sub

Public Sub OpenPath(ByVal defaultPath As String)
    Dim ActivePath  As String
    ActivePath = """" & defaultPath & """"
    Shell "explorer.exe " & ActivePath, vbNormalFocus
End Sub

'--------------------------------------------------------------------------------
' Function: WorkbookIsOpen
' Purpose: Проверяет, открыта ли книга с указанным именем
' Parameters:
'   wbName - Имя книги (без пути)
' Returns: Boolean - True если книга открыта
'--------------------------------------------------------------------------------
Public Function WorkbookIsOpen(ByVal wbName As String) As Boolean
    Dim wb          As Workbook

    On Error Resume Next
    Set wb = Workbooks(wbName)
    On Error GoTo 0

    WorkbookIsOpen = Not wb Is Nothing
End Function

'--------------------------------------------------------------------------------
' Function: SheetExists
' Purpose: Проверяет наличие листа в указанной книге
' Parameters:
'   wb - Ссылка на книгу
'   sheetName - Имя листа
' Returns: Boolean - True если лист существует
'--------------------------------------------------------------------------------
Public Function HaveSheetInFile(ByRef wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws          As Worksheet

    If wb Is Nothing Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    HaveSheetInFile = Not ws Is Nothing
End Function

'--------------------------------------------------------------------------------
' Sub: ConfigureDropButton
' Purpose: Настраивает отображение кнопки выпадающего списка для TextBox
' Parameters:
'   txtBox - Ссылка на элемент управления TextBox
'--------------------------------------------------------------------------------
Public Sub ConfigureDropButton(ByRef txtBox As MSForms.TextBox)
    If txtBox Is Nothing Then Exit Sub

    With txtBox
        .DropButtonStyle = fmDropButtonStyleEllipsis
        .ShowDropButtonWhen = fmShowDropButtonWhenAlways
    End With
End Sub

Public Sub CenterUserForm(ByRef frm As Object)
    With frm
        .StartUpPosition = 0
        .Left = Application.Left + (Application.Width - .Width) \ 2
        .Top = Application.Top + (Application.Height - .Height) \ 2
        
        ' Защита от отрицательных координат
        If .Left < 0 Then .Left = 0
        If .Top < 0 Then .Top = 0
    End With
End Sub


'--------------------------------------------------------------------------------
' Function: SelectRangeViaDialog
' Purpose: Открывает диалог выбора диапазона и возвращает адрес
' Parameters:
'   promptText - Текст приглашения (по умолчанию "Выберите диапазон")
' Returns: String - Адрес диапазона в формате "'Лист'!$A$1" или vbNullString
'--------------------------------------------------------------------------------
Public Function SelectRangeViaDialog(Optional ByVal promptText As String = "Выберите диапазон", _
        Optional bApendSheetName As Boolean = True, Optional bSelectedRng As Boolean = True) As String
    Dim rng         As Range
    Dim defaultAddress As String

    ' Устанавливаем стиль ссылок A1
    Application.ReferenceStyle = xlA1

    ' Формируем адрес по умолчанию из текущего выделения
    If TypeName(Selection) = "Range" And bSelectedRng Then
        defaultAddress = Selection.Address
    End If

    ' Получаем диапазон от пользователя
    On Error Resume Next
    Set rng = Application.InputBox( _
            prompt:=promptText & ":", _
            Default:=defaultAddress, _
            Type:=8)
    On Error GoTo 0

    ' Обрабатаем результат
    If rng Is Nothing Then
        SelectRangeViaDialog = vbNullString
    Else
        ' Возвращаем адрес первого диапазона (при множественном выборе)
        Dim firstPart As String
        firstPart = VBA.Split(rng.Address, ",")(0)
        If bApendSheetName Then firstPart = "'" & rng.Parent.Name & "'!" & firstPart
        SelectRangeViaDialog = firstPart
    End If
End Function

'--------------------------------------------------------------------------------
' Sub: RestrictNavigationKeys
' Purpose: Ограничивает ввод только навигационными клавишами (Tab, Enter, Escape)
' Parameters:
'   KeyCode - Код нажатой клавиши (передаётся ByRef для модификации)
'   Shift - Состояние клавиш-модификаторов
'--------------------------------------------------------------------------------
Public Sub RestrictNavigationKeys(ByRef KeyCode As MSForms.ReturnInteger, ByRef Shift As Integer)
    Const KEY_TAB   As Long = 9
    Const KEY_ENTER As Long = 13
    Const KEY_ESCAPE As Long = 27

    Select Case KeyCode
        Case KEY_TAB, KEY_ENTER, KEY_ESCAPE
            ' Разрешённые клавиши - оставляем без изменений
        Case Else
            KeyCode = 0
    End Select
End Sub

Public Sub ValidateNumericKey(ByRef KeyAscii As MSForms.ReturnInteger)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

'--------------------------------------------------------------------------------
' Function: ValidateRealNumericKey
' Purpose: Проверяет ввод в текстовое поле, разрешая только цифры,
'          один десятичный разделитель (точка; запятая конвертируется в точку)
'          и знак минуса в начале строки.
' Parameters:
' txtBox - Текстовое поле для валидации ввода
' KeyAscii - Код нажатой клавиши (модифицируется для блокировки или конвертации)
'--------------------------------------------------------------------------------
Public Sub ValidateRealNumericKey(ByRef txtBox As MSForms.TextBox, ByRef KeyAscii As MSForms.ReturnInteger, ByRef bIsNegative As Boolean)
    Const BACKSPACE As Integer = 8

    ' Пропускаем Backspace
    If KeyAscii = BACKSPACE Then Exit Sub

    ' Пропускаем цифры
    If KeyAscii >= 48 And KeyAscii <= 57 Then Exit Sub

    ' Обработка десятичного разделителя
    If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        ' Если разделитель уже есть или строка пуста - блокируем ввод
        If InStr(1, txtBox.TEXT, ".") > 0 Or Len(txtBox.TEXT) = 0 Then
            KeyAscii = 0
            ' Если введена запятая - конвертируем код символа в точку
        ElseIf KeyAscii = Asc(",") Then
            KeyAscii = Asc(".")
        End If
        Exit Sub
    End If

    ' Обработка знака минуса (только в начале строки)
    If bIsNegative Then
        If KeyAscii = Asc("-") Then
            ' Блокируем, если минус уже есть или курсор стоит не в начале
            If Left(txtBox.TEXT, 1) <> "-" And txtBox.SelStart = 0 Then Exit Sub
        End If
    End If

    ' Блокируем все остальные символы
    KeyAscii = 0
End Sub

'--------------------------------------------------------------------------------
' Function: SanitizeSheetName
' Purpose: Заменяет недопустимые символы в имени листа на символ подчёркивания
' Parameters:
'   sheetName - Исходное имя листа (String)
' Returns: String - Очищенное имя листа, пригодное для использования в Excel
' Remarks:
'   Недопустимые символы для имени листа Excel: \ / * ? : [ ]
'   Также удаляются начальные и конечные пробелы
'--------------------------------------------------------------------------------
Public Function CleanSheetName(ByVal sheetName As String) As String
    Dim invalidChars As Variant
    Dim char        As Variant

    If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)
    invalidChars = Array("\", "/", "*", "?", ":", "[", "]")
    For Each char In invalidChars
        sheetName = Replace(sheetName, char, "_")
    Next char
    CleanSheetName = sheetName
End Function

'--------------------------------------------------------------------------------
' Function: CleanFileName
' Purpose: Удаление недопустимых символов из имени файла
' Parameters:
'   - fileName (String) - исходное имя файла
' Returns: String - очищенное имя файла
'--------------------------------------------------------------------------------
Public Function CleanFileName(ByVal FileName As String) As String

    Dim invalidChars As Variant
    Dim char        As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    FileName = Trim$(FileName)

    For Each char In invalidChars
        FileName = Replace(FileName, char, "_")
    Next char

    CleanFileName = FileName

End Function
Public Function PathDialogFun(Optional defaultPath As String = vbNullString)
    With Application.FileDialog(msoFileDialogFolderPicker)
        .ButtonName = "Выбрать"
        .Title = "Выберите папку:"
        .InitialFileName = defaultPath
        If .Show <> -1 Then Exit Function
        PathDialogFun = .SelectedItems(1)
    End With
End Function

Public Function FileDialogFun(ByVal sPath As String, _
        ByRef bMultiSelect As Boolean, _
        Optional sExpansion As String = "*.xlsm;*.xlsb;*.xlsx;*.xls", Optional bShowMsg As Boolean = True) As String()

    If sPath = vbNullString Or Not (Dir(sPath, vbDirectory) <> vbNullString) Then sPath = ThisWorkbook.Path

    Dim oFd         As FileDialog
    Set oFd = Application.FileDialog(msoFileDialogFilePicker)
    With oFd
        .AllowMultiSelect = bMultiSelect
        'заголовок окна диалога
        .Title = "Выбрать файлы:"
        'очищаем установленные ранее типы файлов
        .Filters.Clear
        'устанавливаем возможность выбора только файлов Excel
        .Filters.Add "Microsoft Excel Files", sExpansion, 1
        'назначаем папку отображения и имя файла по умолчанию
        .InitialFileName = sPath
        'вид диалогового окна(доступно 9 вариантов)
        .InitialView = msoFileDialogViewDetails
        If .Show = 0 Then
            If bShowMsg Then Call MsgBox("Не выбрано ни одного файла!", vbCritical, "Выбор файлов:")
            Exit Function
        End If
        Dim iCount  As Integer
        Dim i       As Integer
        iCount = .SelectedItems.Count
        'ReDim arr(1 To iCount) As String
        ReDim arr(1 To iCount, 1 To 1) As String
        For i = 1 To iCount
            'arr(i) = VBA.CStr(.SelectedItems.Item(i))
            arr(i, 1) = VBA.CStr(.SelectedItems.item(i))
        Next
    End With
    FileDialogFun = arr

    'Для процедуры
    'If (Not (Not (v))) = 0 Then Exit Sub
End Function

Public Function GetBaseName(ByVal sPathFile As String) As String
    'sPathFile - строка, путь.
    'возвращает имя (без расширения) последнего компонента в заданном пути.
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetBaseName = FSO.GetBaseName(sPathFile)
    Set FSO = Nothing
End Function

Public Function GetFileName(ByVal sPathFile As String) As String
    'sPathFile - строка, путь.
    'возвращает имя (с расширением) последнего компонента в заданном пути.
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetFileName = FSO.GetFileName(sPathFile)
    Set FSO = Nothing
End Function

Public Function GetParentFolderName(ByVal sPathFile As String) As String
    'sPathFile - строка, путь.
    'возвращает путь к последнему компоненту в заданном пути (его каталог).
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetParentFolderName = FSO.GetParentFolderName(sPathFile)
    Set FSO = Nothing
End Function

Public Function GetExtensionName(ByVal sPathFile As String) As String
    'sPathFile - строка, путь.
    'возвращает расширение последнего компонента в заданном пути.
    Dim FSO         As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    GetExtensionName = FSO.GetExtensionName(sPathFile)
    Set FSO = Nothing
End Function

Public Sub DeleteItemInListBox(ByRef listBox As MSForms.listBox)
    With listBox
        Dim i       As Long
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then Call .RemoveItem(i)
        Next i
    End With
End Sub

'--------------------------------------------------------------------------------
' Sub: SelectedItemListSheets
' Purpose: Выполняет фильтрацию элементов списка listSheets на основе шаблона.
' Parameters:
'   sValue - Шаблон для фильтрации (поддерживает символы подстановки Like).
'   iCol   - Индекс столбца в списке, по которому производится сравнение.
'--------------------------------------------------------------------------------
Public Sub SelectedItemListSheets(ByRef List As MSForms.listBox, ByRef sValue As String, ByRef iCol As Integer)
    
    Dim i           As Long
    With List
        On Error GoTo goExitSub
        For i = 0 To .ListCount - 1
            .Selected(i) = .List(i, iCol) Like sValue
        Next i
    End With
    Exit Sub
goExitSub:
    On Error GoTo 0
End Sub

Public Function FileHave(ByVal Path As String, ByVal fileAttribute As VbFileAttribute) As Boolean
    Dim FSO         As Object

    ' Проверка на пустоту
    If Path = vbNullString Then Exit Function
    ' Создаем объект FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ' В зависимости от значения параметра IsFolder выбираем метод проверки
    Select Case fileAttribute
        Case VbFileAttribute.vbDirectory
            ' Ищем папку
            FileHave = FSO.FolderExists(Path)
        Case VbFileAttribute.vbNormal
            ' Ищем файл
            FileHave = FSO.FileExists(Path)
    End Select
    ' Освобождаем память
    Set FSO = Nothing
End Function

'--------------------------------------------------------------------------------
' Function: IsArrayDimensioned
' Purpose: Проверяет, имеет ли массив указанное измерение
' Parameters:
' arr - Проверяемый массив
' dimension - Номер измерения (по умолчанию 1)
' Returns: Boolean - True, если измерение существует
'--------------------------------------------------------------------------------
Public Function IsArrayDimensioned(ByRef arr As Variant, _
        Optional ByVal dimension As Long = 1) As Boolean
    If Not IsArray(arr) Then Exit Function

    On Error Resume Next
    IsArrayDimensioned = (UBound(arr, dimension) >= 0)
    On Error GoTo 0
End Function

'--------------------------------------------------------------------------------
' Function: SortArray
' Purpose: Сортировка двумерного массива по указанному столбцу
' Parameters:
' SourceArr - Двумерный массив Variant для сортировки
' n - Номер столбца для сортировки (Integer)
' bFlagSortAs - Направление сортировки: True - по возрастанию, False - по убыванию (по умолчанию True)
' bFlagDigital - Тип данных: True - числовая сортировка, False - строковая (по умолчанию False)
' Returns: Variant - Отсортированный двумерный массив
'--------------------------------------------------------------------------------
Public Function SortArray(SourceArr As Variant, ByVal n As Integer, _
        Optional bFlagSortAs As Boolean = True, _
        Optional bFlagDigital As Boolean = False, _
        Optional bFlagCase As Boolean = False) As Variant
    Dim Check As Boolean, iCount As Integer, jCount As Integer
    Dim sVal1       As Variant
    Dim sVal2       As Variant

    If IsArrayDimensioned(SourceArr, 2) Then
        ReDim tmpArr(UBound(SourceArr, 2)) As Variant
        Do Until Check
            Check = True
            For iCount = LBound(SourceArr, 1) To UBound(SourceArr, 1) - 1
                sVal1 = SourceArr(iCount, n)
                sVal2 = SourceArr(iCount + 1, n)

                If bFlagDigital Then
                    sVal1 = VBA.Val(sVal1)
                    sVal2 = VBA.Val(sVal2)
                ElseIf bFlagCase Then
                    sVal1 = VBA.UCase$(sVal1)
                    sVal2 = VBA.UCase$(sVal2)
                End If

                If bFlagSortAs Then
                    If sVal1 > sVal2 Then
                        For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2)
                            tmpArr(jCount) = SourceArr(iCount, jCount)
                            SourceArr(iCount, jCount) = SourceArr(iCount + 1, jCount)
                            SourceArr(iCount + 1, jCount) = tmpArr(jCount)
                            Check = False
                        Next
                    End If
                Else
                    If sVal1 < sVal2 Then
                        For jCount = LBound(SourceArr, 2) To UBound(SourceArr, 2)
                            tmpArr(jCount) = SourceArr(iCount, jCount)
                            SourceArr(iCount, jCount) = SourceArr(iCount + 1, jCount)
                            SourceArr(iCount + 1, jCount) = tmpArr(jCount)
                            Check = False
                        Next
                    End If
                End If
            Next
        Loop
    End If

    SortArray = SourceArr
End Function

'--------------------------------------------------------------------------------
' Sub: SortColumnList
' Purpose: Сортировка содержимого ListBox по указанному столбцу с переключением направления
' Parameters:
' lMainList - Ссылка на элемент управления ListBox для сортировки
' oLabelBtn - Ссылка на элемент Label, используемый для отображения направления сортировки
' iCol - Номер столбца для сортировки (Integer)
' bFlagDigital - Тип сортировки: True - числовая, False - строковая (по умолчанию False)
' Returns: Нет
'--------------------------------------------------------------------------------
Public Sub SortColumnList(ByRef lMainList As MSForms.listBox, ByRef oLabelBtn As MSForms.Label, _
        ByVal iCol As Integer, Optional bFlagDigital As Boolean = False, Optional bFlagCase As Boolean = False)
    Dim arr()       As Variant
    Dim bFlag       As Boolean

    With oLabelBtn
        If .Caption = "p" Then
            bFlag = False
            .Caption = "q"
        Else
            bFlag = True
            .Caption = "p"
        End If
    End With

    With lMainList
        If .ListCount > 1 Then
            arr = .List
            arr = SortArray(arr, iCol, bFlag, bFlagDigital, bFlagCase)
            .List = arr
        End If
    End With
End Sub

Public Function CheckProtectStructure() As Boolean
    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Установлен пароль на структуру книги, выполнение операции не возножно!", vbCritical)
        CheckProtectStructure = True
    End If
End Function

Public Function GetColorFromDialog(Optional defaultColor As Long = 255) As Long
    Const COLOR_INDEX As Long = 56
    Dim originalColor As Long
    Dim result      As Long

    originalColor = ActiveWorkbook.Colors(COLOR_INDEX)

    If Application.Dialogs(xlDialogEditColor).Show(COLOR_INDEX, _
            defaultColor Mod 256, _
            (defaultColor \ 256) Mod 256, _
            (defaultColor \ 65536) Mod 256) Then

        result = ActiveWorkbook.Colors(COLOR_INDEX)
        GetColorFromDialog = result
    Else
        GetColorFromDialog = -1
    End If

    ActiveWorkbook.Colors(COLOR_INDEX) = originalColor
End Function

Public Function MoveFile(OldFile As String, NewPathFile As String) As String
    'Перемещение файлов
    Dim objFSO As Object, objFile As Object
    If Dir(OldFile, 16) = vbNullString Then MoveFile = "Нет такого файла" & OldFile: Exit Function
    'перемещаем файл
    Set objFSO = CreateObject("Scripting.FileSystemObject"): Set objFile = objFSO.GetFile(OldFile)
    objFile.Copy NewPathFile
    Set objFile = Nothing: Set objFSO = Nothing
    MoveFile = vbNullString
End Function

Public Sub URLLinks(ByVal url_str As String)
    On Error GoTo ErrorHandler

    Dim appEX       As Object
    Set appEX = CreateObject("Wscript.Shell")
    appEX.Run url_str
    Set appEX = Nothing
    Exit Sub
ErrorHandler:
    Select Case Err
        Case Else:
            Call MsgBox("Произошла ошибка в URLLinks" & vbNewLine & Err.Number & vbNewLine & Err.Description, vbOKOnly + vbCritical, "Ошибка в URLLinks")
    End Select
    Set appEX = Nothing
    Err.Clear
End Sub

' Main function: Returns a two-dimensional array with file information
Function GetFilesTable(ByVal folderPath As String) As Variant
    Dim FSO         As Object
    Dim folder      As Object
    Dim fileCount   As Long
    Dim varResult   As Variant
    Dim rowIndex    As Long

    ' Create FileSystemObject
    Set FSO = CreateObject("Scripting.FileSystemObject")

    ' 1. Check if folder exists
    If Not FSO.FolderExists(folderPath) Then
        MsgBox "Specified folder does not exist:" & folderPath, vbExclamation, "Error"
        GetFilesTable = Array()    ' Return empty array
        Exit Function
    End If

    Set folder = FSO.GetFolder(folderPath)

    ' 2. First pass: Count files
    fileCount = CountFilesRecursive(folder)

    ' If no files found
    If fileCount = 0 Then
        GetFilesTable = Array()
        Exit Function
    End If

    ' 3. Initialize two-dimensional array
    ' Rows: from 1 to fileCount
    ' Columns: from 1 to 4 (Path, Name, Size, Date)
    ReDim varResult(1 To fileCount, 1 To 4)

    ' 4. Second pass: Fill array
    rowIndex = 1    ' Start filling from the first row
    Call FillFilesTableRecursive(folder, varResult, rowIndex)

    ' Return filled table
    GetFilesTable = varResult

    ' Free memory
    Set folder = Nothing
    Set FSO = Nothing
End Function

' Helper function: Recursive file count
Private Function CountFilesRecursive(ByVal currentFolder As Object) As Long
    Dim subFolder   As Object
    Dim lCount      As Long

    lCount = currentFolder.Files.Count

    For Each subFolder In currentFolder.SubFolders
        lCount = lCount + CountFilesRecursive(subFolder)
    Next subFolder

    CountFilesRecursive = lCount
End Function

' Helper procedure: Recursive array filling
Private Sub FillFilesTableRecursive(ByVal currentFolder As Object, ByRef tableArray As Variant, ByRef currentRow As Long)
    Dim subFolder   As Object
    Dim file        As Object

    ' Fill data for each file in the current folder
    For Each file In currentFolder.Files
        tableArray(currentRow, 1) = file.Name                ' File name
        tableArray(currentRow, 2) = file.Path                ' Full path
        tableArray(currentRow, 3) = file.Size                ' Size (bytes)
        tableArray(currentRow, 4) = file.DateLastModified    ' Modification date

        currentRow = currentRow + 1
    Next file

    ' Recursive call for subfolders
    For Each subFolder In currentFolder.SubFolders
        Call FillFilesTableRecursive(subFolder, tableArray, currentRow)
    Next subFolder
End Sub

Public Function SaveTextToFile(ByVal txt As String, ByVal FileName As String, Optional ByVal encoding As String = "windows-1251") As Boolean
    ' функция сохраняет текст txt в кодировке Charset$ в файл filename$
    ' encoding: koi8-r, ascii, utf-7, utf-8, utf-8noBOM, utf-16, Windows-1251, unicode
    On Error Resume Next: Err.Clear
    Dim FSO         As Object
    Dim ts          As Object
    Select Case encoding$
        Case "windows-1251", "", "ansi"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(FileName, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing

        Case "utf-16", "utf-16LE"
            Set FSO = CreateObject("scripting.filesystemobject")
            Set ts = FSO.CreateTextFile(FileName, True, True)
            ts.Write txt: ts.Close
            Set ts = Nothing: Set FSO = Nothing

        Case "utf-8noBOM"
            Dim binaryStream As Object
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = "utf-8": .Open
                .WriteText txt$

                Set binaryStream = CreateObject("ADODB.Stream")
                binaryStream.Type = 1: binaryStream.Mode = 3: binaryStream.Open
                .Position = 3: .CopyTo binaryStream            'Skip BOM bytes
                .Flush: .Close
                binaryStream.SaveToFile FileName$, 2
                binaryStream.Close
            End With

        Case Else
            With CreateObject("ADODB.Stream")
                .Type = 2: .Charset = encoding$: .Open
                .WriteText txt$
                .SaveToFile FileName$, 2            ' сохраняем файл в заданной кодировке
                .Close
            End With
    End Select
    SaveTextToFile = Err = 0: DoEvents
End Function

