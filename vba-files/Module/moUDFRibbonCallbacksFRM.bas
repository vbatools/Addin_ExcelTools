Attribute VB_Name = "moUDFRibbonCallbacksFRM"
Option Explicit
Option Private Module

Private Sub btnCalculationOperations(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    frmFormulaCalculationOperations.Show
End Sub

'--------------------------------------------------------------------------------
' Sub: btnCatchError
' Purpose: Wraps formulas in the selected range with IFERROR function
'          for error handling in calculations
' Parameters:
' control - Ribbon object (IRibbonControl)
'--------------------------------------------------------------------------------
Private Sub btnCatchError(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    ' Проверка типа выделенного объекта
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Пожалуйста, выделите диапазон ячеек.", vbExclamation, "Внимание")
        Exit Sub
    End If

    Dim arrData     As Variant
    Dim arrResult   As Variant
    Dim sLength     As String

    ' Запрос значения для обработки ошибок
    sLength = InputBox("Укажите новое значение для замены в формулах:", "Обновление обработки ошибок")

    ' Подготовка значения: если это текст, берем в кавычки
    If Not IsNumeric(sLength) Then
        sLength = """" & sLength & """"
    End If

    Application.ScreenUpdating = False
    On Error GoTo ErrorHandler

    ' Получение данных (в VBA всегда в английском формате)
    arrData = Selection.formula

    ' Нормализация массива для одной ячейки
    If Not IsArray(arrData) Then
        Dim arrSingle(1 To 1, 1 To 1) As Variant
        arrSingle(1, 1) = arrData
        arrData = arrSingle
    End If

    Dim iCount      As Long
    Dim jCount      As Long
    Dim i           As Long
    Dim j           As Long
    Dim sFormula    As String

    iCount = UBound(arrData, 1)
    jCount = UBound(arrData, 2)

    ReDim arrResult(1 To iCount, 1 To jCount)

    ' Обработка каждой ячейки
    For i = 1 To iCount
        For j = 1 To jCount
            sFormula = arrData(i, j)

            ' Проверяем, является ли содержимое формулой
            If Left$(sFormula, 1) = "=" Then

                Dim sNewFormula As String
                Dim sTempFormula As String

                ' Убираем "="
                sTempFormula = Mid$(sFormula, 2)

                ' Проверяем: начинается ли формула с IFERROR
                ' В VBA .Formula всегда используется английское имя "IFERROR"
                If StrComp(Left$(sTempFormula, 8), "IFERROR(", vbTextCompare) = 0 Then

                    ' СЛУЧАЙ А: Формула уже есть (=IFERROR(...))

                    ' Удаляем "IFERROR(" (9 символов)
                    Dim sBody As String
                    sBody = Mid$(sTempFormula, 9)

                    ' Находим позицию последней запятой
                    Dim iSepPos As Long
                    iSepPos = InStrRev(sBody, ",")

                    If iSepPos > 0 Then
                        ' Заменяем аргумент после запятой
                        sNewFormula = "=IFERROR(" & Left$(sBody, iSepPos) & sLength & ")"
                    Else
                        ' Запятой не было, добавляем
                        sNewFormula = "=IFERROR(" & sBody & "," & sLength & ")"
                    End If
                Else
                    ' СЛУЧАЙ Б: Обычная формула -> Оборачиваем
                    sNewFormula = "=IFERROR(" & sTempFormula & "," & sLength & ")"
                End If
                arrResult(i, j) = sNewFormula
            Else
                arrResult(i, j) = sFormula
            End If
        Next j
    Next i

    ' Вывод результата
    Call SaveUndoInfo(Selection, False, False)
    Selection.formula = arrResult
    Application.OnUndo "Отменить", "RestoreUndoInfo"

ErrorHandler:
    Application.ScreenUpdating = True
End Sub


Private Sub btnAddFormulaPropSumm(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Call frmFormulaWordsNumbers.Show
End Sub

Private Sub btnAddFormulaPropDate(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Call frmFormulaWordsDate.Show
End Sub

Private Sub btnAddFormulaPropTime(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Call frmFormulaWordsTime.Show
End Sub

Private Sub btnAddFormulaPropQRCode(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Call frmFormulaQRCode.Show
End Sub

Private Sub btnAddFileQRCode(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Call frmFormulaQRCodeAsFile.Show
End Sub
