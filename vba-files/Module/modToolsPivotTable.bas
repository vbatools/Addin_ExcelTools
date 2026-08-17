Attribute VB_Name = "modToolsPivotTable"
Option Explicit
Option Private Module

Public Sub DialogPivotFieldProperties()
    'диалог параметров поля сводной таблицы
    On Error Resume Next
    Application.Dialogs(xlDialogPivotFieldProperties).Show
End Sub

Public Sub DialogPivotShowPages()
    'диалог разбить сводную по листам фильтра
    On Error Resume Next
    Application.Dialogs(xlDialogPivotShowPages).Show
End Sub

'--------------------------------------------------------------------------------
' Sub: RefreshPivotCachesClearMissingItems
' Purpose: Выполняет очистку кэшей сводных таблиц в книге, удаляя устаревшие
'          элементы из фильтров (устанавливает MissingItemsLimit = None).
'--------------------------------------------------------------------------------
Public Sub RefreshPivotCachesClearMissingItems()
    Dim ws          As Worksheet
    Dim pvt         As PivotTable
    Dim lProtectedCount As Long
    Dim lProcessedCount As Long

    ' Инициализация счетчиков состояния
    lProtectedCount = 0
    lProcessedCount = 0

    Application.ScreenUpdating = False    ' Блокировка обновления экрана для ускорения

    ' Перебор всех листов в активной книге
    For Each ws In ActiveWorkbook.Worksheets
        If ws.ProtectContents Then
            ' Фиксация количества защищенных листов для отчета
            lProtectedCount = lProtectedCount + 1
        Else
            ' Обработка сводных таблиц на незащищенном листе
            For Each pvt In ws.PivotTables
                ' Локальная обработка ошибок для совместимости с Моделью Данных
                On Error Resume Next
                pvt.PivotCache.MissingItemsLimit = xlMissingItemsNone

                If Err.Number = 0 Then
                    lProcessedCount = lProcessedCount + 1
                End If
                On Error GoTo 0
            Next pvt
        End If
    Next ws

    ActiveWorkbook.RefreshAll

    Application.ScreenUpdating = True    ' Восстановление обновления экрана

    ' Вывод итогового сообщения пользователю
    If lProtectedCount > 0 Then
        MsgBox _
                "Операция завершена частично." & vbNewLine & vbNewLine & _
                "Обработано таблиц: " & lProcessedCount & "." & vbNewLine & _
                "Пропущено защищенных листов: " & lProtectedCount & "." & vbNewLine & vbNewLine & _
                "Для выполнения очистки на защищенных листах необходимо снять защиту " & _
                "через меню 'Рецензирование' -> 'Снять защиту листа'.", _
                vbExclamation, _
                "Внимание: Ограничение доступа"
    Else
        MsgBox _
                "Кэши всех сводных таблиц (" & lProcessedCount & ") успешно обновлены. " & _
                "Устаревшие элементы удалены.", _
                vbInformation + vbOKOnly, _
                "Успешное выполнение"
    End If

End Sub


