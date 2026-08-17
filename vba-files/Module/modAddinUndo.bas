Attribute VB_Name = "modAddinUndo"
Option Explicit
Option Private Module
' Модульная переменная для хранения данных отката
Private m_UndoData  As UndoCellData

' Тип данных для операции отката
Private Type UndoCellData
    uRange          As Range
    vValues         As Variant
    isMerge         As Boolean
    mergedAreas     As Collection    ' Коллекция адресов объединённых областей
End Type

'--------------------------------------------------------------------------------
' Sub: RestoreUndoInfo
' Purpose: Восстановление ранее сохранённых значений и структуры объединения ячеек
' Returns: Ничего. Восстанавливает значения и объединения в диапазоне.
' Remarks:
'   - Вызывается через Application.OnUndo
'   - При ошибке выводит сообщение пользователю
'   - Игнорирует ошибку, если книга или лист не найдены
'--------------------------------------------------------------------------------
Public Sub RestoreUndoInfo()
    On Error GoTo ErrorHandler

    With m_UndoData
        If .uRange Is Nothing Then Exit Sub

        ' Активируем лист
        .uRange.Parent.Activate

        If .isMerge Then
            .uRange.UnMerge
        Else
            ' Восстанавливаем структуру объединений (если была сохранена)
            RestoreMergedAreas .uRange.Parent, .mergedAreas
        End If

        ' Восстанавливаем значения
        .uRange.FormulaR1C1 = .vValues
    End With

ExitRoutine:
    ' Очищаем данные после восстановления
    ClearUndoData
    Exit Sub

ErrorHandler:
    ' Книга может быть закрыта - это допустимо
    If Err.Number = 9 Then  ' Subscript out of range
        Resume ExitRoutine
    End If

    ' Для других ошибок информируем пользователя
    MsgBox "Не удалось восстановить предыдущие значения:" & vbCrLf & _
            "Книга: " & m_UndoData.uRange.Parent.Parent.Name & vbCrLf & _
            "Лист: " & m_UndoData.uRange.Parent.Name & vbCrLf & _
            "Адрес: " & m_UndoData.uRange.Address & vbCrLf & vbCrLf & _
            "Ошибка: " & Err.Description, _
            vbExclamation, "Восстановление"
    Resume ExitRoutine
End Sub

'--------------------------------------------------------------------------------
' Sub: SaveUndoInfo
' Purpose: Сохранение текущих значений диапазона для последующего отката
' Parameters:
'   - rng (Range) - диапазон ячеек для сохранения
'   - saveMergeInfo (Boolean) - сохранять ли информацию об объединённых ячейках
' Returns: Ничего. Сохраняет данные в модульную переменную.
' Remarks:
'   - При saveMergeInfo = True сохраняется структура объединений
'   - При saveMergeInfo = False сохраняются только значения ячеек
'--------------------------------------------------------------------------------
Public Sub SaveUndoInfo(ByRef rng As Range, ByVal saveMergeInfo As Boolean, ByVal isMerge As Boolean)
    If rng Is Nothing Then Exit Sub

    ' Очищаем предыдущие данные
    ClearUndoData

    With m_UndoData
        Set .uRange = rng
        .vValues = rng.FormulaR1C1
        .isMerge = isMerge
        ' Сохраняем информацию об объединённых ячейках, если требуется
        If saveMergeInfo Then
            Set .mergedAreas = GetMergedAreasInfo(rng)
        End If
    End With
End Sub

'--------------------------------------------------------------------------------
' Function: GetMergedAreasInfo
' Purpose: Получение информации об объединённых ячейках в диапазоне
' Parameters:
'   - rng (Range) - диапазон для анализа
' Returns: Collection - коллекция строк с адресами объединённых областей
'--------------------------------------------------------------------------------
Private Function GetMergedAreasInfo(ByRef rng As Range) As Collection
    Dim result      As New Collection
    Dim ws          As Worksheet
    Dim cell        As Range
    Dim mergedArea  As Range
    Dim processedAreas As Object
    Dim areaKey     As String

    Set ws = rng.Parent
    Set processedAreas = CreateObject("Scripting.Dictionary")

    On Error Resume Next

    ' Перебираем все ячейки в диапазоне
    For Each cell In rng.Cells
        If cell.MergeCells Then
            Set mergedArea = cell.MergeArea

            ' Формируем уникальный ключ для области
            areaKey = mergedArea.Address

            ' Проверяем, не обрабатывали ли уже эту область
            If Not processedAreas.Exists(areaKey) Then
                ' Сохраняем адрес объединённой области
                result.Add areaKey
                processedAreas.Add areaKey, True
            End If
        End If
    Next cell

    On Error GoTo 0

    Set GetMergedAreasInfo = result
End Function

'--------------------------------------------------------------------------------
' Sub: RestoreMergedAreas
' Purpose: Восстановление структуры объединений ячеек
' Parameters:
'   - ws (Worksheet) - рабочий лист
'   - mergedAreas (Collection) - коллекция адресов объединённых областей
'--------------------------------------------------------------------------------
Private Sub RestoreMergedAreas(ByRef ws As Worksheet, ByRef mergedAreas As Collection)
    Dim areaAddress As Variant
    Dim rngToMerge  As Range

    ' Если коллекция пустая или не инициализирована - выходим
    If mergedAreas Is Nothing Then Exit Sub
    If mergedAreas.Count = 0 Then Exit Sub

    On Error Resume Next    ' Игнорируем ошибки при восстановлении объединений

    DisableApplicationSettings
    For Each areaAddress In mergedAreas
        Set rngToMerge = ws.Range(areaAddress)
        If Not rngToMerge Is Nothing Then
            rngToMerge.Merge
        End If
    Next areaAddress
    RestoreApplicationSettings

    On Error GoTo 0
End Sub

'--------------------------------------------------------------------------------
' Sub: ClearUndoData
' Purpose: Очистка сохранённых данных отката
'--------------------------------------------------------------------------------
Private Sub ClearUndoData()
    With m_UndoData
        Set .uRange = Nothing
        .vValues = Empty
        Set .mergedAreas = Nothing
        .isMerge = False
    End With
End Sub

