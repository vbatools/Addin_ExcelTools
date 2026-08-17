VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCrossSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' МОДУЛЬ КЛАССА: CrossSelectionManager
' Purpose: Управление перекрёстным выделением строк/столбцов в Excel
'================================================================================

'--------------------------------------------------------------------------------
' КОНСТАНТЫ МОДУЛЯ
'--------------------------------------------------------------------------------
Private Const FORMULA_MARKER As String = "=1*1*1*1"

'--------------------------------------------------------------------------------
' ПЕРЕМЕННЫЕ УРОВНЯ МОДУЛЯ
'--------------------------------------------------------------------------------
Private WithEvents AppExc As Excel.Application
Attribute AppExc.VB_VarHelpID = -1
Private m_stateOn   As Boolean
Private m_selectionType As CrossSelectionType
Private m_colorInterior As Long
Private m_colorFont As Long
Private m_settingsLoaded As Boolean

'--------------------------------------------------------------------------------
' Enum: CrossSelectionType
' Purpose: Определяет тип перекрёстного выделения
'--------------------------------------------------------------------------------
Private Enum CrossSelectionType
    cstRowColumn = 0
    cstRow = 1
    cstColumn = 2
End Enum

'================================================================================
' PUBLIC PROPERTIES
'================================================================================

Public Property Get IsStateOn() As Boolean
    IsStateOn = m_stateOn
End Property

Public Property Let IsStateOn(ByVal Value As Boolean)
    m_stateOn = Value
    If Not m_stateOn Then
        Call RemoveCrossSelectionsOnly(ActiveSheet)
    End If
End Property

'================================================================================
' PUBLIC METHODS
'================================================================================

'--------------------------------------------------------------------------------
' Procedure: RefreshSettings
' Purpose: Принудительное обновление настроек из реестра
'--------------------------------------------------------------------------------
Public Sub RefreshSettings()
    m_settingsLoaded = False
    Call LoadSettingsIfNeeded
End Sub

'--------------------------------------------------------------------------------
' Procedure: ClearAllSelections
' Purpose: Удаление всех правил перекрёстного выделения на активном листе
'--------------------------------------------------------------------------------
Public Sub ClearAllSelections()
    Call RemoveCrossSelectionsOnly(ActiveSheet)
End Sub

'================================================================================
' PRIVATE PROCEDURES - EVENTS
'================================================================================

Private Sub Class_Initialize()
    Set AppExc = Application
    m_colorInterior = &HC0FFC0
    m_colorFont = 0
    m_selectionType = cstRowColumn
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not AppExc Is Nothing Then
        Call RemoveCrossSelectionsOnly(ActiveSheet)
    End If
    Set AppExc = Nothing
End Sub

Private Sub AppExc_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next

    ' Пропускаем выделение нескольких ячеек
    If Target.CountLarge > 1 Then Exit Sub

    ' Проверяем тип листа
    If Not TypeOf Sh Is Worksheet Then Exit Sub

    Call ApplyCrossSelection(Sh, Target)
End Sub

Private Sub AppExc_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    On Error Resume Next
    Dim ws          As Worksheet
    For Each ws In wb.Worksheets
        Call RemoveCrossSelectionsOnly(ws)
    Next ws
End Sub

Private Sub AppExc_WorkbookOpen(ByVal wb As Workbook)
    On Error Resume Next
    Dim ws          As Worksheet
    For Each ws In wb.Worksheets
        Call RemoveCrossSelectionsOnly(ws)
    Next ws
End Sub

Private Sub AppExc_SheetDeactivate(ByVal Sh As Object)
    On Error Resume Next
    Call RemoveCrossSelectionsOnly(Sh)
End Sub


'================================================================================
' PRIVATE PROCEDURES - CORE LOGIC
'================================================================================

'--------------------------------------------------------------------------------
' Procedure: ApplyCrossSelection
' Purpose: Применяет перекрёстное выделение для указанного диапазона
' Parameters:
'   ws     - Рабочий лист
'   Target - Выбранная ячейка
'--------------------------------------------------------------------------------
Private Sub ApplyCrossSelection(ByVal ws As Worksheet, ByRef Target As Range)
    Dim rngDisplay  As Range
    Dim rngSelection As Range
    Dim fc          As FormatCondition

    On Error GoTo ErrorHandler

    If AppExc Is Nothing Then Set AppExc = Application
    If Not m_stateOn Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' Загружаем настройки при необходимости
    Call LoadSettingsIfNeeded

    ' Получаем используемый диапазон
    Set rngDisplay = GetUsedRange(ws)
    If rngDisplay Is Nothing Then GoTo CleanUp

    ' Удаляем старые правила
    Call RemoveCrossSelectionsOnly(ws)

    ' Определяем диапазон выделения
    Set rngSelection = GetSelectionRange(rngDisplay, Target)

    ' Применяем условное форматирование
    If Not rngSelection Is Nothing Then
        Set fc = rngSelection.FormatConditions.Add(Type:=xlExpression, Formula1:=FORMULA_MARKER)
        With fc
            .Interior.Color = m_colorInterior
            .Font.Color = m_colorFont
            .SetFirstPriority
            .StopIfTrue = False
        End With
    End If

CleanUp:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Resume CleanUp
End Sub

'--------------------------------------------------------------------------------
' Function: GetSelectionRange
' Purpose: Возвращает диапазон для выделения в зависимости от типа
' Parameters:
'   rngDisplay - Используемый диапазон листа
'   Target     - Выбранная ячейка
' Returns: Range - диапазон для применения форматирования
'--------------------------------------------------------------------------------
Private Function GetSelectionRange(ByVal rngDisplay As Range, ByVal Target As Range) As Range
    On Error Resume Next

    Select Case m_selectionType
        Case cstColumn
            Set GetSelectionRange = Intersect(rngDisplay, Target.EntireColumn)
        Case cstRow
            Set GetSelectionRange = Intersect(rngDisplay, Target.EntireRow)
        Case Else
            Set GetSelectionRange = Intersect(rngDisplay, Union(Target.EntireRow, Target.EntireColumn))
    End Select
End Function

'--------------------------------------------------------------------------------
' Procedure: RemoveCrossSelectionsOnly
' Purpose: Удаляет правила условного форматирования, созданные данным модулем
' Parameters:
'   ws - Рабочий лист
'--------------------------------------------------------------------------------
Private Sub RemoveCrossSelectionsOnly(ByVal ws As Worksheet)
    Dim fc          As FormatCondition
    Dim i           As Long

    On Error Resume Next

    If ws Is Nothing Then Exit Sub

    With ws.Cells
        For i = .FormatConditions.Count To 1 Step -1
            If TypeOf .FormatConditions(i) Is FormatCondition Then
                Set fc = .FormatConditions(i)
                If fc.Type = xlExpression Then
                    If fc.Formula1 = FORMULA_MARKER Then
                        fc.Delete
                    End If
                End If
            End If
        Next i
    End With
End Sub

'--------------------------------------------------------------------------------
' Procedure: LoadSettingsIfNeeded
' Purpose: Загружает настройки из реестра при первом обращении
'--------------------------------------------------------------------------------
Private Sub LoadSettingsIfNeeded()
    Dim settingValue As String

    If m_settingsLoaded Then Exit Sub

    On Error Resume Next

    ' Загрузка типа выделения
    settingValue = GetSetting(NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_Select", "RowColumn")
    Select Case settingValue
        Case "Row": m_selectionType = cstRow
        Case "Column": m_selectionType = cstColumn
        Case Else: m_selectionType = cstRowColumn
    End Select

    ' Загрузка цветов
    m_colorInterior = Val(GetSetting(NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_ColorInterior", "&HC0FFC0"))
    m_colorFont = Val(GetSetting(NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_ColorFont", "0"))

    m_settingsLoaded = True
End Sub

'================================================================================
' PRIVATE FUNCTIONS
'================================================================================

'--------------------------------------------------------------------------------
' Function: GetUsedRange
' Purpose: Возвращает используемый диапазон листа с точным определением границ
' Parameters:
'   ws - Рабочий лист
' Returns: Range - используемый диапазон от A1 до последней заполнённой ячейки
'--------------------------------------------------------------------------------
Private Function GetUsedRange(ByVal ws As Worksheet) As Range
    Dim LastRow     As Long
    Dim LastCol     As Long

    On Error GoTo ErrorHandler

    LastRow = ws.Cells.Find(What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlPrevious).Row

    LastCol = ws.Cells.Find(What:="*", _
            After:=ws.Cells(1, 1), _
            LookIn:=xlFormulas, _
            LookAt:=xlPart, _
            SearchOrder:=xlByColumns, _
            SearchDirection:=xlPrevious).Column

    If LastRow > 0 And LastCol > 0 Then
        Set GetUsedRange = ws.Range(ws.Cells(1, 1), ws.Cells(LastRow, LastCol))
    End If
    Exit Function

ErrorHandler:
    Set GetUsedRange = ws.UsedRange
End Function

