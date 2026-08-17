VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataAddWorkBooks 
   Caption         =   "Создание шаблонов книг:"
   ClientHeight    =   6360
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8835.001
   OleObjectBlob   =   "frmDataAddWorkBooks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataAddWorkBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' UserForm     :   frmDataUniqueValues - создание книг Excel
' Author       :   VBATools
' Copyright    :   Apache License
' Created      :   10-06-2026 09:35:59
' Refactored   :   2024
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' ============================================================================
' ОБРАБОТЧИКИ СОБЫТИЙ
' ============================================================================

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()

    If Not ValidateInput Then Exit Sub

    Dim arrNames    As Variant
    If Not GetNamesArray(arrNames) Then Exit Sub

    Dim FileExt     As String
    Dim fileFormat  As XlFileFormat
    Call GetFileFormatInfo(FileExt, fileFormat)

    Dim excelApp    As Excel.Application
    Dim templateBook As Workbook

    On Error GoTo ErrorHandler

    ' --- Создание экземпляра Excel ---
    Set excelApp = New Excel.Application
    excelApp.DisplayAlerts = False
    excelApp.Visible = False


    ' --- Открытие шаблона ---
    Dim templatePath As String
    templatePath = BuildPath(txtPath.Value, txtWBookName.Value)

    Set templateBook = excelApp.Workbooks.Open(templatePath, False, False)

    ' --- Создание книг ---
    Call CreateWorkbooks(templateBook, arrNames, txtPath.Value, FileExt, fileFormat)

    ' --- Закрытие шаблона без сохранения ---
    templateBook.Close savechanges:=False

CleanUp:
    If Not excelApp Is Nothing Then
        excelApp.Quit
        Set excelApp = Nothing
    End If
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка выполнения"
    Resume CleanUp

End Sub

Private Sub UserForm_Initialize()

    ' Центрирование формы
    Call CenterUserForm(Me)

    ' --- Инициализация диапазона ---
    If TypeName(Selection) = "Range" Then
        txtListNameWBook.Value = Selection.Address
    End If

    ' --- Инициализация пути и имени файла ---
    txtWBookName.Value = ActiveWorkbook.Name
    txtPath.Value = ActiveWorkbook.Path & Application.PathSeparator

    ' --- Определение типа файла ---
    Call SelectFileTypeByExtension(ActiveWorkbook.Name)

    ' --- Настройка кнопки выбора диапазона ---
    Call ConfigureDropButton(txtListNameWBook)

End Sub

' ============================================================================
' ОБРАБОТЧИКИ ЭЛЕМЕНТОВ УПРАВЛЕНИЯ
' ============================================================================

Private Sub optActWBook_Change()

    If Not optActWBook.Value Then Exit Sub

    txtWBookName.Value = ActiveWorkbook.Name
    optPathActWBook.Value = True

End Sub

Private Sub optChoseWB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
        ByVal x As Single, ByVal y As Single)

    Dim initialPath As String
    initialPath = IIf(txtWBookName.Value = vbNullString, _
            ActiveWorkbook.Path & Application.PathSeparator, _
            txtWBookName.Value)

    Dim selectedFile As Variant
    selectedFile = SelectFileViaDialog(initialPath)

    If IsEmpty(selectedFile) Then Exit Sub

    ' --- Обновление полей ---
    txtWBookName.Value = GetFileName(CStr(selectedFile))
    txtPath.Value = GetParentFolderName(CStr(selectedFile)) & Application.PathSeparator

    ' --- Автовыбор типа файла ---
    Call SelectFileTypeByExtension(CStr(selectedFile))

    ' --- Снятие выбора с опции активной книги ---
    If optActWBook.Value Then optActWBook.Value = False
    optChoseWB.Value = True

End Sub

Private Sub optPathActWBook_Change()

    If optPathActWBook.Value Then
        txtPath.Value = ActiveWorkbook.Path & Application.PathSeparator
    Else
        Call SelectFolderViaDialog
    End If

End Sub

Private Sub txtListNameWBook_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtListNameWBook_DropButtonClick()

    Me.Hide
    txtListNameWBook.Value = SelectRangeViaDialog()
    Me.Show

End Sub

' ============================================================================
' БИЗНЕС-ЛОГИКА
' ============================================================================

Private Function ValidateInput() As Boolean

    ' --- Проверка диапазона ---
    If Trim$(txtListNameWBook.Value) = vbNullString Then
        MsgBox "Не выбран диапазон названий книг!", vbCritical, "Ошибка"
        Exit Function
    End If

    ' --- Проверка имени файла ---
    If Trim$(txtWBookName.Value) = vbNullString Then
        MsgBox "Не указан файл-шаблон!", vbCritical, "Ошибка"
        Exit Function
    End If

    ValidateInput = True
End Function

Private Function GetNamesArray(ByRef arrNames As Variant) As Boolean

    Dim sAddress    As String
    sAddress = Trim$(txtListNameWBook.Value)

    On Error Resume Next
    arrNames = Range(sAddress).Value2
    On Error GoTo 0

    If IsEmpty(arrNames) Then
        MsgBox "Не выбран диапазон названий книг!", vbCritical, "Ошибка"
        Exit Function
    End If

    ' --- Нормализация до массива ---
    If Not IsArray(arrNames) Then
        arrNames = Array(arrNames)
    End If

    ' --- Удаление пустых значений ---
    arrNames = FilterNonEmpty(arrNames)

    GetNamesArray = True

End Function

Private Sub CreateWorkbooks(ByVal templateBook As Workbook, _
        ByVal arrNames As Variant, _
        ByVal savePath As String, _
        ByVal FileExt As String, _
        ByVal fileFormat As XlFileFormat)

    Dim i           As Long
    Dim FileName    As String
    Dim fullPath    As String

    For i = LBound(arrNames) To UBound(arrNames)

        FileName = CleanFileName(CStr(arrNames(i)))

        If Len(FileName) > 0 Then
            fullPath = BuildPath(savePath, FileName & FileExt)

            templateBook.SaveAs _
                    FileName:=fullPath, _
                    fileFormat:=fileFormat, _
                    ReadOnlyRecommended:=False
        End If

    Next i

End Sub

' ============================================================================
' ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
' ============================================================================

Private Sub GetFileFormatInfo(ByRef extension As String, ByRef format As XlFileFormat)

    If optXLSM.Value Then
        extension = ".xlsm"
        format = xlOpenXMLWorkbookMacroEnabled
    ElseIf optXLSB.Value Then
        extension = ".xlsb"
        format = xlExcel12
    ElseIf optXLS.Value Then
        extension = ".xls"
        format = xlExcel8
    Else
        extension = ".xlsx"
        format = xlOpenXMLWorkbook
    End If

End Sub

Private Sub SelectFileTypeByExtension(ByVal FilePath As String)

    Dim ext         As String
    ext = LCase$(GetExtensionName(FilePath))

    Select Case ext
        Case "xlsm": optXLSM.Value = True
        Case "xlsb": optXLSB.Value = True
        Case "xls": optXLS.Value = True
        Case Else: optXLSX.Value = True
    End Select

End Sub

Private Function SelectFileViaDialog(ByVal initialPath As String) As Variant

    Dim arrResult() As String
    arrResult = FileDialogFun(initialPath, False, , False)

    If (Not (Not arrResult)) = 0 Then
        SelectFileViaDialog = Empty
    Else
        SelectFileViaDialog = arrResult(1, 1)
    End If

End Function

Private Sub SelectFolderViaDialog()

    With Application.FileDialog(msoFileDialogFolderPicker)
        .ButtonName = "Выбрать"
        .Title = NAME_ADDIN
        .InitialFileName = ActiveWorkbook.Path

        If .Show = -1 Then
            txtPath.Value = .SelectedItems(1) & Application.PathSeparator
        Else
            optPathActWBook.Value = True
        End If
    End With

End Sub

Private Function FilterNonEmpty(ByVal arr As Variant) As Variant

    Dim result()    As Variant
    Dim i As Long, Count As Long

    ReDim result(LBound(arr) To UBound(arr))

    For i = LBound(arr) To UBound(arr)
        If Not IsEmpty(arr(i)) And Len(Trim$(CStr(arr(i)))) > 0 Then
            result(Count) = arr(i)
            Count = Count + 1
        End If
    Next i

    If Count = 0 Then
        FilterNonEmpty = Array()
    Else
        ReDim Preserve result(0 To Count - 1)
        FilterNonEmpty = result
    End If

End Function


Private Function BuildPath(ByVal folder As String, ByVal FileName As String) As String

    If Right$(folder, 1) <> Application.PathSeparator Then
        folder = folder & Application.PathSeparator
    End If

    BuildPath = folder & FileName

End Function

