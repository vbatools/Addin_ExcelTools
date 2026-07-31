Attribute VB_Name = "modToolsUnmergeCells"
Option Explicit
Option Private Module

'--------------------------------------------------------------------------------
' Procedure: unmergeCells
' Purpose: Разъединяет объединённые ячейки и заполняет все ячейки бывшей
'          объединённой области значением из первой ячейки
' Parameters: нет
'--------------------------------------------------------------------------------
Public Sub unMergeCells()

    Dim targetRange As Range
    Dim currentArea As Range
    Dim currentCell As Range
    Dim cellValue   As Variant
    Dim mergedArea  As Range
    Dim mergedAreas() As Range
    Dim mergedValues() As Variant
    Dim mergedCount As Long
    Dim index       As Long
    Dim sRng        As String

    ' Получаем пересечение используемого диапазона листа и выделенного диапазона
    sRng = SelectRangeViaDialog()
    If sRng = vbNullString Then Exit Sub
    Set targetRange = Range(sRng)
    If targetRange Is Nothing Then Exit Sub

    ' Предварительное определение размера массива
    mergedCount = 0
    ReDim mergedAreas(1 To targetRange.Count)
    ReDim mergedValues(1 To targetRange.Count)

    ' Первый проход: сбор данных об объединённых ячейках
    For Each currentArea In targetRange.Areas
        For Each currentCell In currentArea
            If currentCell.MergeCells Then
                Set mergedArea = currentCell.MergeArea
                ' Проверяем, что это первая ячейка в объединённой области
                If currentCell.Address = mergedArea.Cells(1).Address Then
                    mergedCount = mergedCount + 1
                    Set mergedAreas(mergedCount) = mergedArea
                    mergedValues(mergedCount) = mergedArea.Cells(1).Value
                End If
            End If
        Next currentCell
    Next currentArea

    ' Если нет объединённых ячеек, завершаем
    If mergedCount = 0 Then Exit Sub

    ' Оптимизация производительности
    Call DisableApplicationSettings
    Call SaveUndoInfo(targetRange, True, False)
    ' Второй проход: разъединение и заполнение
    For index = 1 To mergedCount
        mergedAreas(index).UnMerge
        mergedAreas(index).Value = mergedValues(index)
    Next index
    Application.OnUndo "Отменить", "RestoreUndoInfo"
    ' Восстановление настроек
    Call RestoreApplicationSettings
End Sub
