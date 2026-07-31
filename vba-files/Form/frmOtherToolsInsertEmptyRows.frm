VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherToolsInsertEmptyRows 
   Caption         =   "Вставка данных:"
   ClientHeight    =   4110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235.001
   OleObjectBlob   =   "frmOtherToolsInsertEmptyRows.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtherToolsInsertEmptyRows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim errMsg      As String
    Dim iStep       As Integer
    Dim rngValue    As Range
    Dim rngCellInsert As Range
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long
    Dim k           As Long
    Dim m           As Long

    ' Проверка заполнения параметров
    If txtInputRng.Value = vbNullString Then
        errMsg = "Диапазон данных"
    End If

    If txtCellInsert.Value = vbNullString Then
        If errMsg <> vbNullString Then errMsg = errMsg & vbNewLine
        errMsg = errMsg & "Ячейка вставки"
    End If

    iStep = VBA.Val(txtStep.Value)
    If iStep < 1 Then
        If errMsg <> vbNullString Then errMsg = errMsg & vbNewLine
        errMsg = errMsg & "Шаг вставки"
    End If

    If errMsg <> vbNullString Then
        Call MsgBox("Не заполнены параметры:" & vbNewLine & errMsg, vbCritical)
        Exit Sub
    End If

    ' Безопасное получение ссылок на диапазоны
    On Error Resume Next
    Set rngValue = Range(txtInputRng.Value)
    Set rngCellInsert = Range(txtCellInsert.Value)
    On Error GoTo 0

    If rngValue Is Nothing Or rngCellInsert Is Nothing Then
        MsgBox "Указаны некорректные адреса диапазонов.", vbCritical
        Exit Sub
    End If
    
    Call DisableApplicationSettings

    iStep = iStep + 1
    Dim rng         As Range
    If optRow.Value Then

        iCount = rngValue.Rows.Count * iStep
        jCount = rngValue.Columns.Count
        Set rng = rngCellInsert
        Call SaveUndoInfo(rng.Resize(iCount, jCount), False, False)
        Set rng = Nothing
        For i = 1 To iCount Step iStep
            k = k + 1
            For j = 1 To jCount
                With rngCellInsert.Cells(i, j)
                    .FormulaR1C1 = "=" & rngValue.Cells(k, j).Address(, , xlR1C1, True)
                    If Not chbLinks.Value Then .Value2 = .Value2
                End With
            Next j

            If chDubleEmpty.Value Then
                For m = 1 To iStep - 1
                    For j = 1 To jCount
                        With rngCellInsert.Cells(i + m, j)
                            .FormulaR1C1 = "=" & rngValue.Cells(k, j).Address(, , xlR1C1, True)
                            If Not chbLinks.Value Then .Value2 = .Value2
                        End With
                    Next j
                Next m
            End If
        Next i
    Else
        ' Реализация для столбцов
        iCount = rngValue.Rows.Count
        jCount = rngValue.Columns.Count * iStep
        Set rng = rngCellInsert
        Call SaveUndoInfo(rng.Resize(iCount, jCount), False, False)
        Set rng = Nothing
        For j = 1 To jCount Step iStep
            k = k + 1
            For i = 1 To iCount
                rngCellInsert.Cells(i, j).FormulaR1C1 = "=" & rngValue.Cells(i, k).Address(, , xlR1C1, True)
            Next i

            If chDubleEmpty.Value Then
                For m = 1 To iStep - 1
                    For i = 1 To iCount
                        rngCellInsert.Cells(i, j + m).FormulaR1C1 = "=" & rngValue.Cells(i, k).Address(, , xlR1C1, True)
                    Next i
                Next m
            End If
        Next j
    End If
    
    Call RestoreApplicationSettings
    Application.OnUndo "Отменить", "RestoreUndoInfo"
    Me.Hide
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtCellInsert_DropButtonClick()
    Me.Hide
    txtCellInsert.Value = SelectRangeViaDialog()
    txtCellInsert.Value = VBA.Split(txtCellInsert.Value, ":")(0)
    Call msgSecondaryCell
    Me.Show
End Sub

Private Sub txtCellInsert_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtStep_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub optRow_Change()
    Call msgSecondaryCell
End Sub

Private Sub txtStep_Change()
    Call msgSecondaryCell
End Sub

Private Sub msgSecondaryCell()
    If txtCellInsert.Value <> vbNullString And txtStep.Value <> vbNullString Then
        Dim rng     As Range
        Dim iStep   As Integer
        iStep = VBA.Val(txtStep.Value) + 1
        Set rng = Range(txtCellInsert.Value)

        If optRow.Value Then
            lbSecond.Caption = rng.Offset(iStep, 0).Address(False, False)
        Else
            lbSecond.Caption = rng.Offset(0, iStep).Address(False, False)
        End If
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    If TypeName(Selection) = "Range" Then txtInputRng.Value = Selection.Address
    Call ConfigureDropButton(txtInputRng)
    Call ConfigureDropButton(txtCellInsert)
End Sub
