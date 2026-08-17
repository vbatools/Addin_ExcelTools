Attribute VB_Name = "modFilesTools"
Option Explicit
Option Private Module

Public Sub AddFilesList()
    Dim rng         As Range
    Dim arrPath()   As String
    Dim i           As Long
    Dim iCount      As Long
    Dim jCount      As Byte
    Dim j           As Integer
    Dim objFile     As Object
    Dim sPathName   As String
    Dim sFileName   As String
    Dim iLen        As Integer
    Dim arrVal      As Variant
    Dim arrValNum   As Variant
    'On Error GoTo AddFileNewName_Err
    Call DisableApplicationSettings

    arrPath = FileDialogFun(ActiveWorkbook.Path, True, "*.*")
    If (Not (Not (arrPath))) = 0 Then Exit Sub
    arrVal = Array("Размер файла:", "Тип:", "Дата изменения:", "Дата создания:", "Дата открытия:", "Дата изменения:", "Атрибуты:", "Тип:", "Ключевые слова:", "Оценка:", "Автор:", "Название:", "Тема:", "Категории:", "Комментарий:", "Авторские права:", "Размеры:", "Организация:", "Имя программы:", "Состояние содержимого:", "Язык:")
    arrValNum = Array(1, 2, 3, 4, 5, 128, 6, 11, 18, 19, 20, 21, 22, 23, 24, 25, 31, 33, 35, 126, 185)
    jCount = UBound(arrValNum)
    ActiveWorkbook.Sheets.Add After:=Sheets(Sheets.Count)
    Set rng = ActiveSheet.Cells(1, 1)
    With rng
        .Cells(1, 1).Value = "Директория к файлу:"
        .Cells(1, 2).Value = "Расширение файла:"
        For j = 0 To jCount
            .Cells(1, j + 4).Value = arrVal(j)
        Next j

        .Cells(1, j + 4).Value = "новое имя файла:"
        .Cells(1, j + 4).AddComment
        .Cells(1, j + 4).Comment.Visible = False
        .Cells(1, j + 4).Comment.Shape.DrawingObject.Font.Size = 14
        .Cells(1, j + 4).Comment.TEXT TEXT:="Переименование файлов" & vbNewLine & "Если не нужно переименовывать то оставить пустым!" & vbNewLine & vbNewLine & "При пустом названии файла и имени пути, файл будет скопирован в новую папку по умолчанию, без переименования!"
        .Cells(1, j + 4).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
        .Cells(1, j + 4).Comment.Shape.ScaleHeight 2.5, msoFalse, msoScaleFromTopLeft
        .Cells(1, j + 5).Value = "Переместить файл в папку:"
        .Cells(1, j + 5).AddComment
        .Cells(1, j + 5).Comment.Visible = False
        .Cells(1, j + 5).Comment.Shape.DrawingObject.Font.Size = 14
        .Cells(1, j + 5).Comment.TEXT TEXT:="Перемещение файла в папку" & vbNewLine & "Если не нужно перемещать то оставить пустым!" & vbNewLine & "Для Пустых будет создана новая папка по умолчанию!"
        .Cells(1, j + 5).Comment.Shape.ScaleWidth 3, msoFalse, msoScaleFromTopLeft
        .Cells(1, j + 5).Comment.Shape.ScaleHeight 2.5, msoFalse, msoScaleFromTopLeft
        .Cells(1, j + 6).Value = "Переместить файл в папку:"
        iCount = UBound(arrPath)
        ReDim arr(1 To iCount, 1 To jCount + 4)
        For i = 1 To iCount
            iLen = VBA.InStrRev(arrPath(i, 1), Application.PathSeparator)
            If iLen > 0 Then
                sPathName = VBA.Left$(arrPath(i, 1), iLen)
                sFileName = VBA.Replace(arrPath(i, 1), sPathName, vbNullString)
                Set objFile = CreateObject("Shell.Application").Namespace((sPathName))
                arr(i, 1) = sPathName
                iLen = VBA.InStrRev(sFileName, ".")
                arr(i, 3) = VBA.Left$(sFileName, iLen - 1)
                arr(i, 2) = VBA.Replace(sFileName, arr(i, 3) & ".", vbNullString)

                For j = 0 To jCount
                    arr(i, 4 + j) = objFile.GetDetailsOf(objFile.ParseName((sFileName)), arrValNum(j))
                Next j
                Set objFile = Nothing
            End If
            If i Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = "Выполнено: " & i & " из " & iCount & ", " & VBA.format$(i / iCount, "Percent")
            End If
        Next i
        .Cells(2, 1).Resize(iCount, jCount + 4).Value2 = arr
        .Columns("A:Z").EntireColumn.AutoFit
        .Columns("J:X").Columns.Group
        ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
        Application.StatusBar = False
        Call RestoreApplicationSettings
    End With
    Exit Sub
AddFileNewName_Err:
    Application.StatusBar = False
    Call RestoreApplicationSettings
    MsgBox Err.Description & vbCrLf & "в VBAProject.D_Macros.LoadFileName " & vbCrLf & "в строке " & Erl, vbExclamation + vbOKOnly, "Ошибка:"
End Sub

'Переименовывание файлов по списку
Public Sub MoveAndRenameFiles()
    Dim rng         As Range
    Dim OldFile     As String
    Dim NewPathFile As String
    Dim NewPath     As String
    Dim StrErr      As String
    Dim sFileName   As String
    Dim i           As Long
    Dim n           As Byte
    On Error GoTo Canceled
    Set rng = Application.InputBox(prompt:="Выберите диапазон", Title:="Выбор диапазона:", Default:=Selection.Address, Type:=8)
    On Error GoTo 0
    On Error GoTo AddFileNewName_Err
    With rng
        If .Cells(1, 1).Value <> "Директория к файлу:" And .Cells(1, 25).Value <> "Новое название файла:" Then
            Call MsgBox("Должен быть выбран весь диапазон данных!" & vbNewLine & "Начиная с ячейки [ Директория к файлу ] и заканчивая ячейкой [ Перемещение файла в папку ]", vbCritical, "Ошибка:")
            Exit Sub
        End If
        'создаю новую уникальную папку
        If .Cells(1, 1).Value = "Директория к файлу:" Then
            n = 2
        Else
            n = 1
        End If

        Dim sNewPath As String
        sNewPath = Application.PathSeparator & "new_" & VBA.Replace(VBA.Replace(VBA.Now(), ":", "."), " ", "_") & Application.PathSeparator

        'переношу и переименовую файлы
        On Error Resume Next
        For i = n To .Rows.Count
            If .Cells(i, 1) <> vbNullString And .Cells(i, 2) <> vbNullString And .Cells(i, 3) <> vbNullString Then
                OldFile = .Cells(i, 1) & Application.PathSeparator & .Cells(i, 3) & "." & .Cells(i, 2)

                'проверка папок
                If .Cells(i, 26) <> vbNullString Then
                    NewPath = CleanFileName(.Cells(i, 26)) & IIf(VBA.Right$(.Cells(i, 26), 1) = Application.PathSeparator, vbNullString, Application.PathSeparator)
                Else
                    NewPath = .Cells(n, 1) & sNewPath
                End If
                'проверка на сущ папки
                If VBA.Len(Dir(NewPath, vbDirectory)) = 0 Then
                    Call MkDir(NewPath)
                    StrErr = "создана папка: " & NewPath & vbNewLine
                End If

                If .Cells(i, 25) <> vbNullString Then
                    sFileName = .Cells(i, 25).Value
                Else
                    sFileName = .Cells(i, 3).Value
                End If

                NewPathFile = NewPath & sFileName & "." & .Cells(i, 2)
                If Err.Number = 0 Then
                    .Cells(i, 27).Value = StrErr & MoveFile(OldFile, NewPathFile)
                Else
                    .Cells(i, 27).Value = "Ошибка: " & Err.Description
                    Err.Clear
                End If
            Else
                .Cells(i, 27).Value = "ошибка в данных: " & .Cells(i, 1) & " " & .Cells(i, 2) & " " & .Cells(i, 3) & " " & .Cells(i, 25) & " " & .Cells(i, 26)
            End If
            StrErr = vbNullString
        Next i
        On Error GoTo 0
    End With
    Call MsgBox("Переименовывание и перемещение файлов завершено!", vbInformation + vbOKOnly, "Переименовывание и перемещение файлов:")
    Exit Sub
AddFileNewName_Err:
    MsgBox Err.Description & vbCrLf & "в VBAProject.D_Macros.AddFileNewName " & vbCrLf & "в строке " & Erl, vbExclamation + vbOKOnly, "Ошибка:"
    Exit Sub
Canceled:
    Call MsgBox("Диапазон данных не выбран!", vbInformation + vbOKOnly, "Выбор диапазона:")
End Sub
