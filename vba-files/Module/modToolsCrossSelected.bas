Attribute VB_Name = "modToolsCrossSelected"
Option Explicit
Option Private Module
'================================================================================
' МОДУЛЬ: modCrossSelection
' Purpose: Управление перекрёстным выделением ячеек
'================================================================================

Private oCrossSelected As clsCrossSel

'--------------------------------------------------------------------------------
' Procedure: CrossSelection
' Purpose: Включает/выключает режим перекрёстного выделения
'--------------------------------------------------------------------------------
Public Sub CrossSelection()
    On Error GoTo ErrorHandler

    ' Проверка версии Excel 2007
    If Application.Version = "12.0" Then
        If MsgBox("Перекрёстное выделение в Excel 2007 не работает совместно с условным форматированием." & vbCrLf & _
                "Условное форматирование (если используется) будет удалено." & vbCrLf & vbCrLf & _
                "Продолжить?", vbYesNo Or vbExclamation Or vbDefaultButton2, "Перекрёстное выделение") = vbNo Then
            Exit Sub
        End If
    End If

    ' Проверка типа листа
    If Not TypeOf ActiveSheet Is Worksheet Then
        MsgBox "Перекрёстное выделение применяется только к рабочему листу.", vbCritical, "Ошибка"
        Exit Sub
    End If

    ' Проверка защиты листа
    If ActiveSheet.ProtectContents Then
        MsgBox "Активный лист защищён паролем. Перекрёстное выделение невозможно.", vbCritical, "Ошибка"
        If Not oCrossSelected Is Nothing Then
            oCrossSelected.IsStateOn = False
            Set oCrossSelected = Nothing
        End If
        Exit Sub
    End If

    ' Переключение состояния
    If oCrossSelected Is Nothing Then
        Set oCrossSelected = New clsCrossSel
        oCrossSelected.IsStateOn = True
    Else
        oCrossSelected.IsStateOn = False
        Set oCrossSelected = Nothing
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Перекрёстное выделение"
End Sub
