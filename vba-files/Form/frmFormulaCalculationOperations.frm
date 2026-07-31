VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaCalculationOperations 
   Caption         =   "Произвети вычисление над диапазоном:"
   ClientHeight    =   5565
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "frmFormulaCalculationOperations.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaCalculationOperations"
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
' Purpose: Выполняет арифметическую или логическую операцию над значениями в
'          выбранном диапазоне ячеек с учетом разделителя десятичных знаков
' Parameters:
' Нет (считывает значения из элементов управления формы)
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()

    ' Проверка обязательных полей
    If txtInputRng.Value = vbNullString Then
        Call MsgBox("Не выбран диапазон данных", vbCritical)
        Exit Sub
    End If

    Select Case txtValue.Value
        Case vbNullString, "-"
            Call MsgBox("Не выбрано значение", vbCritical)
            Exit Sub
    End Select

    ' Получение значения оператора
    Dim snVal       As Single
    If Application.DecimalSeparator = "," Then
        snVal = VBA.CSng(VBA.Replace(txtValue.Value, ".", ","))
    Else
        snVal = VBA.CSng(VBA.Replace(txtValue.Value, ",", "."))
    End If

    ' Проверка деления на ноль
    If (optDivide.Value Or optDivideInt.Value) And snVal = 0 Then
        Call MsgBox("Введено значение ноль, на ноль делить нельзя", vbCritical)
        Exit Sub
    End If

    ' Загрузка данных из диапазона в массив
    Dim rng         As Range
    Dim arr         As Variant
    Set rng = Range(txtInputRng.Value)
    arr = rng.FormulaR1C1
    
    If Not IsArray(arr) Then
        ReDim arrData(1 To 1, 1 To 1)
        arrData(1, 1) = arr
        arr = arrData
    End If
    
    ' Обработка массива
    Dim i As Long, j As Long
    Dim iCount As Long, jCount As Long
    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)

    For i = 1 To iCount
        For j = 1 To jCount
            ' Нормализация десятичного разделителя
            If Application.DecimalSeparator = "," Then
                arr(i, j) = VBA.Replace(arr(i, j), ".", ",")
            Else
                arr(i, j) = VBA.Replace(arr(i, j), ",", ".")
            End If

            ' Применение операции к числовым значениям
            If IsNumeric(arr(i, j)) Then
                Select Case True
                    Case optRound.Value
                        arr(i, j) = VBA.Round(arr(i, j), VBA.Abs(VBA.Fix(snVal)))
                    Case optMultiply.Value
                        arr(i, j) = arr(i, j) * snVal
                    Case optAdd.Value
                        arr(i, j) = arr(i, j) + snVal
                    Case optSubtract.Value
                        arr(i, j) = arr(i, j) - snVal
                    Case optDivide.Value
                        arr(i, j) = arr(i, j) / snVal
                    Case optDivideInt.Value
                        arr(i, j) = arr(i, j) \ snVal
                    Case optMod.Value
                        arr(i, j) = arr(i, j) Mod snVal
                End Select
            End If
        Next j
    Next i

    ' Запись результата обратно в диапазон
    rng.FormulaR1C1 = arr
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
        ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateRealNumericKey(txtValue, KeyAscii, True)
End Sub

'--------------------------------------------------------------------------------
' Event: UserForm_Initialize
' Purpose: Инициализация формы при запуске (центрирование, заполнение полей)
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Call CenterUserForm(Me)

    If TypeName(Selection) = "Range" Then
        txtInputRng.Value = Selection.Address
    End If
    Call ConfigureDropButton(txtInputRng)
End Sub

