VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherToolsComments 
   Caption         =   "Создать примечания к ячейкам:"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17025
   OleObjectBlob   =   "frmOtherToolsComments.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtherToolsComments"
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
' Purpose: Обработчик подтверждения с проверкой и выполнением вставки примечаний
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()
    Dim errMsg      As String
    errMsg = checkRange()

    If errMsg <> vbNullString Then
        MsgBox errMsg, vbCritical
        Exit Sub
    End If

    ' Проверка защиты листа
    If Not CheckWorksheetProtection(ActiveSheet) Then
        Exit Sub
    End If

    Dim rngInput    As Range
    Set rngInput = Range(txtInputRng.Value)

    Select Case True
        Case optCommentOne.Value
            Call AddSingleComment(rngInput)
        Case optFormuls.Value
            Call AddFormulasComment(rngInput)
        Case optRange.Value
            Call AddRangeComments(rngInput)
    End Select
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Sub: AddSingleComment
' Purpose: Добавление одинакового текста примечания во все ячейки диапазона
' Parameters:
' rngInput - Диапазон ячеек для вставки примечаний
'--------------------------------------------------------------------------------
Private Sub AddSingleComment(ByVal rngInput As Range)
    Dim sText       As String
    Dim i           As Long
    Dim iCount      As Long

    sText = txtTextComment.Value
    If sText = vbNullString Then sText = " "

    iCount = rngInput.Count
    For i = 1 To iCount
        With rngInput.Cells(i)
            If Not (chbEmptyRngInput.Value And .Value = vbNullString) Then
                Call DeleteExistingComment(.Comment)
                Call AddCommentWithImage(.AddComment(sText), i)
            End If
        End With
    Next i
End Sub

'--------------------------------------------------------------------------------
' Sub: AddFormulasComment
' Purpose: Добавление формул в примечания, если ячейка содержит формулу
' Parameters:
' rngInput - Диапазон ячеек для вставки примечаний
'--------------------------------------------------------------------------------
Private Sub AddFormulasComment(ByVal rngInput As Range)
    Dim i           As Long
    Dim iCount      As Long

    iCount = rngInput.Count
    For i = 1 To iCount
        With rngInput.Cells(i)
            If Left(.formulaLocal, 1) = "=" Then
                Call DeleteExistingComment(.Comment)
                Call AddCommentWithImage(.AddComment(.formulaLocal), i)
            End If
        End With
    Next i
End Sub

'--------------------------------------------------------------------------------
' Sub: AddRangeComments
' Purpose: Добавление примечаний из указанного диапазона с возможностью пропуска пустых
' Parameters:
' rngInput - Диапазон ячеек для вставки примечаний
'--------------------------------------------------------------------------------
Private Sub AddRangeComments(ByVal rngInput As Range)
    Dim rngComments As Range
    Set rngComments = Range(txtRngComments.Value)

    Dim i           As Long
    Dim iCount      As Long
    iCount = rngInput.Count

    For i = 1 To iCount
        Dim sVal    As String
        sVal = rngComments.Cells(i).Value
        If sVal = vbNullString Then sVal = " "

        With rngInput.Cells(i)
            If Not (chbEmptyRngInput.Value And .Value = vbNullString) Then
                If Not (chbNotEmptyComments.Value And sVal = " ") Then
                    Call DeleteExistingComment(.Comment)
                    Call AddCommentWithImage(.AddComment(sVal), i)
                End If
            End If
        End With
    Next i
End Sub

'--------------------------------------------------------------------------------
' Sub: DeleteExistingComment
' Purpose: Удаление существующего примечания, если оно есть
' Parameters:
' oComment - Объект Comment для удаления
'--------------------------------------------------------------------------------
Private Sub DeleteExistingComment(ByVal oComment As Comment)
    If Not oComment Is Nothing Then
        oComment.Delete
    End If
End Sub

'--------------------------------------------------------------------------------
' Sub: AddCommentWithImage
' Purpose: Настройка видимости примечания и вставка изображения
' Parameters:
' oComment - Объект Comment для настройки
' i - Индекс текущей ячейки
'--------------------------------------------------------------------------------
Private Sub AddCommentWithImage(ByVal oComment As Comment, ByVal i As Long)
    With oComment
        .Visible = Not chbShowComments.Value
        Call InsertImageToComment(oComment, i)
    End With
End Sub

'--------------------------------------------------------------------------------
' Sub: InsertImageToComment
' Purpose: Вставка изображения в примечание
' Parameters:
' oComment - Объект Comment для вставки изображения
' i - Индекс текущей ячейки
'--------------------------------------------------------------------------------
Private Sub InsertImageToComment(ByRef oComment As Comment, ByRef i As Long)
    If Not chbImage.Value Then Exit Sub

    Dim sPath       As String
    sPath = GetImagePath(i)
    If sPath <> vbNullString Then
        With oComment.Shape
            If chbSaveProporce.Value Then
                Dim imgWidth As Long
                Dim imgHeight As Long
                Call GetImageSizeViaShape(sPath, imgWidth, imgHeight)
                .Fill.UserPicture sPath
                .Width = imgWidth
                .Height = imgHeight
            End If
        End With
    End If
End Sub

'--------------------------------------------------------------------------------
' Function: GetImageSizeViaShape
' Purpose: Получение размеров изображения через временный Shape
' Parameters:
' imagePath - Путь к файлу изображения
' imgWidth - (Out) Ширина изображения в пунктах
' imgHeight - (Out) Высота изображения в пунктах
' Returns: Boolean - True если успешно
'--------------------------------------------------------------------------------
Public Function GetImageSizeViaShape(ByVal imagePath As String, _
        ByRef imgWidth As Long, _
        ByRef imgHeight As Long) As Boolean

    On Error GoTo ErrorHandler

    If Dir(imagePath) = "" Then
        GetImageSizeViaShape = False
        Exit Function
    End If

    Dim tempShp     As Shape
    Set tempShp = ActiveSheet.Shapes.AddPicture(imagePath, _
            msoFalse, _
            msoTrue, _
            0, 0, -1, -1)
    imgWidth = tempShp.Width
    imgHeight = tempShp.Height
    tempShp.Delete

    GetImageSizeViaShape = True
    Exit Function

ErrorHandler:
    GetImageSizeViaShape = False

End Function

'--------------------------------------------------------------------------------
' Function: GetImagePath
' Purpose: Получение пути к изображению в зависимости от режима выбора
' Parameters:
' i - Индекс текущей ячейки для режима диапазона
' Returns: String - Путь к файлу изображения
'--------------------------------------------------------------------------------
Private Function GetImagePath(ByVal i As Long) As String
    If optSingleFile.Value Then
        GetImagePath = txtFileSingle.Value
    Else
        If txtRngFile.Value <> vbNullString Then
            GetImagePath = Range(txtRngFile.Value).Cells(i).Value
        End If
    End If
End Function

'--------------------------------------------------------------------------------
' Function: CheckWorksheetProtection
' Purpose: Проверка защиты листа и попытка снятия защиты
' Returns: Boolean - True если лист не защищен или защита успешно снята
'--------------------------------------------------------------------------------
Private Function CheckWorksheetProtection(ByVal ws As Worksheet) As Boolean

    If Not ws.ProtectContents Then
        CheckWorksheetProtection = True
        Exit Function
    End If

    Dim response    As VbMsgBoxResult
    response = MsgBox("Лист '" & ws.Name & "' защищен." & vbNewLine & _
            "Необходимо снять защиту для добавления примечаний." & vbNewLine & vbNewLine & _
            "Попытаться снять защиту?", _
            vbYesNo + vbQuestion, "Защищенный лист")

    If response = vbNo Then
        CheckWorksheetProtection = False
        Exit Function
    End If

    ' Попытка снятия защиты без пароля
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        MsgBox "Не удалось снять защиту листа." & vbNewLine & _
                "Возможно, установлен пароль.", vbExclamation
        CheckWorksheetProtection = False
        Exit Function
    End If

    CheckWorksheetProtection = True

End Function

'--------------------------------------------------------------------------------
' Event: chbImage_Change
' Purpose: Включение/отключение группы элементов для работы с изображениями
'--------------------------------------------------------------------------------
Private Sub chbImage_Change()
    frmImage.Enabled = chbImage.Value
End Sub

'--------------------------------------------------------------------------------
' Event: optCommentOne_Change
' Purpose: Переключение в режим одиночного комментария
'--------------------------------------------------------------------------------
Private Sub optCommentOne_Change()
    txtTextComment.Enabled = True
    txtRngComments.Enabled = False
End Sub

'--------------------------------------------------------------------------------
' Event: optFormuls_Change
' Purpose: Переключение в режим формул (оба поля отключены)
'--------------------------------------------------------------------------------
Private Sub optFormuls_Change()
    txtTextComment.Enabled = False
    txtRngComments.Enabled = False
End Sub

'--------------------------------------------------------------------------------
' Event: optRange_Change
' Purpose: Переключение в режим диапазона комментариев
'--------------------------------------------------------------------------------
Private Sub optRange_Change()
    txtTextComment.Enabled = False
    txtRngComments.Enabled = True
End Sub

'--------------------------------------------------------------------------------
' Event: optSingleFile_Change
' Purpose: Переключение режима выбора файла изображения
'--------------------------------------------------------------------------------
Private Sub optSingleFile_Change()
    txtFileSingle.Enabled = optSingleFile.Value
    txtRngFile.Enabled = Not txtFileSingle.Enabled
End Sub

'--------------------------------------------------------------------------------
' Function: checkRange
' Purpose: Проверка корректности выбранных диапазонов
' Returns: String - Описание ошибки или пустая строка при успешной проверке
'--------------------------------------------------------------------------------
Private Function checkRange() As String

    Dim errorMessage As String
    Dim iCount      As Long
    errorMessage = vbNullString

    ' Проверка заполнения поля ввода исходного диапазона
    If txtInputRng.Value = vbNullString Then
        errorMessage = "Не выбран диапазон для вставки примечаний"
        checkRange = errorMessage
        Exit Function
    End If

    iCount = Range(txtInputRng.Value).Count

    ' Проверка изображений
    If chbImage.Value Then
        Dim imgError As String
        imgError = CheckImageSettings(iCount)
        If imgError <> vbNullString Then
            If errorMessage <> vbNullString Then
                errorMessage = errorMessage & vbNewLine
            End If
            errorMessage = errorMessage & imgError
        End If
    End If

    ' Проверка для режима диапазона комментариев
    If optRange.Value Then
        Dim rangeError As String
        rangeError = CheckRangeSettings(iCount)
        If rangeError <> vbNullString Then
            If errorMessage <> vbNullString Then
                errorMessage = errorMessage & vbNewLine
            End If
            errorMessage = errorMessage & rangeError
        End If
    End If

    checkRange = errorMessage
End Function

'--------------------------------------------------------------------------------
' Function: CheckImageSettings
' Purpose: Проверка настроек изображений
' Parameters:
' iCount - Количество ячеек в исходном диапазоне
' Returns: String - Описание ошибки или пустая строка
'--------------------------------------------------------------------------------
Private Function CheckImageSettings(ByVal iCount As Long) As String

    If optSingleFile.Value Then
        If txtFileSingle.Value = vbNullString Then
            CheckImageSettings = "Не выбран файл картинки"
        End If
    Else
        If txtRngFile.Value = vbNullString Then
            CheckImageSettings = "Не выбран диапазон путей к файлам картинок"
        ElseIf iCount <> Range(txtRngFile.Value).Count Then
            CheckImageSettings = "Количество ячеек диапазона вставки не равно " & _
                    "количеству ячеек диапазона путей к файлам картинок"
        End If
    End If

End Function

'--------------------------------------------------------------------------------
' Function: CheckRangeSettings
' Purpose: Проверка настроек диапазона комментариев
' Parameters:
' iCount - Количество ячеек в исходном диапазоне
' Returns: String - Описание ошибки или пустая строка
'--------------------------------------------------------------------------------
Private Function CheckRangeSettings(ByVal iCount As Long) As String

    If txtRngComments.Value = vbNullString Then
        CheckRangeSettings = "Не выбран диапазон примечаний"
    ElseIf iCount <> Range(txtRngComments.Value).Count Then
        CheckRangeSettings = "Количество ячеек диапазона вставки не равно " & _
                "количеству ячеек диапазона комментариев"
    End If

End Function

'--------------------------------------------------------------------------------
' Event: txtInputRng_KeyDown
' Purpose: Ограничение навигационных клавиш в поле ввода
'--------------------------------------------------------------------------------
Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

'--------------------------------------------------------------------------------
' Event: txtRngComments_KeyDown
' Purpose: Ограничение навигационных клавиш в поле ввода комментариев
'--------------------------------------------------------------------------------
Private Sub txtRngComments_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

'--------------------------------------------------------------------------------
' Event: txtFileSingle_KeyDown
' Purpose: Ограничение навигационных клавиш в поле ввода пути к файлу
'--------------------------------------------------------------------------------
Private Sub txtFileSingle_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

'--------------------------------------------------------------------------------
' Event: txtRngFile_KeyDown
' Purpose: Ограничение навигационных клавиш в поле ввода диапазона файлов
'--------------------------------------------------------------------------------
Private Sub txtRngFile_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
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
' Event: txtRngComments_DropButtonClick
' Purpose: Скрытие формы, вызов диалога выбора диапазона комментариев
'--------------------------------------------------------------------------------
Private Sub txtRngComments_DropButtonClick()
    Me.Hide
    txtRngComments.Value = SelectRangeViaDialog()
    Me.Show
End Sub

'--------------------------------------------------------------------------------
' Event: txtFileSingle_DropButtonClick
' Purpose: Скрытие формы, вызов диалога выбора файла и повторный показ
'--------------------------------------------------------------------------------
Private Sub txtFileSingle_DropButtonClick()
    Me.Hide
    Dim arr()       As String
    arr = FileDialogFun(ActiveWorkbook.Path, False, "*.png;*.jpeg;*.jpg;*.gif")
    If (Not (Not (arr))) > 0 Then
        txtFileSingle.Value = arr(1, 1)
    End If
    Me.Show
End Sub

'--------------------------------------------------------------------------------
' Event: txtRngFile_DropButtonClick
' Purpose: Скрытие формы, вызов диалога выбора диапазона файлов и повторный показ
'--------------------------------------------------------------------------------
Private Sub txtRngFile_DropButtonClick()
    Me.Hide
    txtRngFile.Value = SelectRangeViaDialog()
    Me.Show
End Sub

'--------------------------------------------------------------------------------
' Event: UserForm_Initialize
' Purpose: Инициализация формы при запуске (центрирование, заполнение полей)
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Call CenterUserForm(Me)

    ' Заполнение поля ввода, если выделен диапазон
    If TypeName(Selection) = "Range" Then
        txtInputRng.Value = Selection.Address
    End If

    Call ConfigureDropButton(txtRngFile)
    Call ConfigureDropButton(txtFileSingle)
    Call ConfigureDropButton(txtRngComments)
    Call ConfigureDropButton(txtInputRng)
End Sub

