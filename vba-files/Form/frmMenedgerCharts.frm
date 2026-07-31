VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerCharts 
   Caption         =   "Менеджер диаграмм:"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16245
   OleObjectBlob   =   "frmMenedgerCharts.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const FILTER_ALL As String = "все"
Private Const FILTER_NONE As String = "ничего"
Private Const FILTER_REVERSE As String = "обратное выделение"

Private Sub btnDown_Click()
    Call moveShape(1, True)
End Sub

Private Sub btnUp_Click()
    Call moveShape(-1, True)
End Sub

Private Sub btnLeft_Click()
    Call moveShape(-1, False)
End Sub

Private Sub btnRight_Click()
    Call moveShape(1, False)
End Sub

Private Sub moveShape(ByVal ShiftMove As Integer, ByVal bHorizont As Boolean)
    Dim i           As Long
    Dim oChart      As ChartObject
    Dim snShag      As Single
    snShag = VBA.CSng(VBA.Replace(txtShag.TEXT, ".", ","))
    If snShag = 0 Then Exit Sub
    With listChart
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Set oChart = ActiveWorkbook.Worksheets(.List(i, 1)).ChartObjects(.List(i, 2))
                With oChart
                    If bHorizont Then
                        .Top = .Top + ShiftMove * snShag
                    Else
                        .Left = .Left + ShiftMove * snShag
                    End If
                End With
            End If
        Next i
    End With
End Sub

Private Sub btnOK_Click()
    Dim i           As Long
    With listChart
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Call ApplyColorsToChart(ActiveWorkbook.Worksheets(.List(i, 1)).ChartObjects(.List(i, 2)).Chart)
            End If
        Next i
    End With
End Sub

Private Sub btnSortNum_Click()
    Call sortColumnList(listChart, btnSortNum, 0, True)
End Sub

Private Sub btnSortName_Click()
    Call sortColumnList(listChart, btnSortName, 1, False)
End Sub

Private Sub btnSorChartName_Click()
    Call sortColumnList(listChart, btnSorChartName, 2, False)
End Sub

Private Sub btnSorChartType_Click()
    Call sortColumnList(listChart, btnSorChartType, 3, False)
End Sub

Private Sub lbOK_Click()
    Call btnOK_Click
End Sub

Private Sub listChart_Change()
    Dim arr()       As String
    arr = getSelectedItemMainList()
    If IsArrayDimensioned(arr) Then
        ActiveSheet.Shapes.Range(arr).Select
    End If
End Sub

Private Function getSelectedItemMainList() As String()
    Dim i           As Long
    Dim j           As Long
    Dim arr()       As String
    With listChart
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                ReDim Preserve arr(0 To j) As String
                arr(j) = .List(i, 2)
                j = j + 1
            End If
        Next i
    End With
    getSelectedItemMainList = arr
End Function

Private Sub listFilters_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With listFilters
        If .ListIndex < 0 Then Exit Sub
        Select Case .Value
            Case FILTER_REVERSE: Call reversSelected
        End Select
    End With
End Sub

Private Sub listFilters_Change()
    With listFilters
        If .ListIndex < 0 Then Exit Sub
        Select Case .Value
            Case FILTER_ALL: Call selectedItemListSheets(listChart, "*", 1)
            Case FILTER_NONE: Call selectedItemListSheets(listChart, vbNullString, 1)
            Case FILTER_REVERSE: Call reversSelected
            Case Else:
                If .ListIndex > -1 Then Call selectedItemListSheets(listChart, .List(.ListIndex, 0), 1)
        End Select
    End With
End Sub

Private Sub reversSelected()
    Dim i           As Long
    With listChart
        For i = 0 To .ListCount - 1
            .Selected(i) = Not .Selected(i)
        Next i
    End With
End Sub

Private Sub txtShag_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateRealNumericKey(txtShag, KeyAscii, False)
End Sub

'--------------------------------------------------------------------------------
' Event: UserForm_Initialize
' Purpose: Инициализация формы при запуске (центрирование, заполнение полей)
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Call CenterUserForm(Me)
    Call refreshList

    Const W         As Byte = 32
    Const h         As Byte = 32
    With Application.CommandBars
        btnOK.Picture = .GetImageMso("GroupWindow", W, h)
    End With

End Sub

Private Sub cmbFilter_Change()
    With cmbFilter
        If .ListIndex < 0 Then
            If .Value = vbNullString Then
                Call selectedItemListSheets(listChart, vbNullString, 2)
            Else
                Call selectedItemListSheets(listChart, "*" & .Value & "*", 2)
            End If
        Else
            ' Оптимизация: Direct access по имени без необходимости предварительной активации
            Call selectedItemListSheets(listChart, cmbFilter.Value, 2)
        End If
    End With
End Sub

Private Sub refreshList()

    Dim ch          As Chart
    Dim chtObj      As ChartObject
    Dim i           As Long

    With listFilters
        .Clear
        .AddItem FILTER_ALL
        .AddItem FILTER_NONE
        .AddItem FILTER_REVERSE
    End With

    cmbFilter.Clear

    ' Перебор всех листов в активной книге
    With listChart
        ' Проверка на наличие встроенных диаграмм
        If ActiveSheet.ChartObjects.Count > 0 Then
            listFilters.AddItem ActiveSheet.Name
            For Each chtObj In ActiveSheet.ChartObjects
                Set ch = chtObj.Chart
                .AddItem i + 1
                .List(i, 1) = ActiveSheet.Name
                .List(i, 2) = chtObj.Name
                .List(i, 3) = GetChartTypeName(chtObj.Chart.chartType)
                cmbFilter.AddItem chtObj.Name
                i = i + 1
            Next chtObj
        End If
    End With
End Sub

'--------------------------------------------------------------------------------
' Function: GetChartTypeName
' Purpose: Преобразует числовое значение типа диаграммы в текстовое описание
' Parameters:
' chartType - XlChartType - числовой тип диаграммы
' Returns: String - текстовое описание типа диаграммы
'--------------------------------------------------------------------------------
Private Function GetChartTypeName(ByVal chartType As XlChartType) As String
    Select Case chartType
        Case xlColumnClustered: GetChartTypeName = "Гистограмма с группировкой"
        Case xlColumnStacked: GetChartTypeName = "Гистограмма с накоплением"
        Case xlLine, xlLineStacked: GetChartTypeName = "График"
        Case xlLineMarkers, xlLineMarkersStacked: GetChartTypeName = "График с маркерами"
        Case xlPie: GetChartTypeName = "Круговая"
        Case xlBarClustered: GetChartTypeName = "Линейчатая с группировкой"
        Case xlArea: GetChartTypeName = "Область"
        Case xlXYScatter: GetChartTypeName = "Точечная"
        Case xlRadar: GetChartTypeName = "Лепестковая"
        Case xlDoughnut: GetChartTypeName = "Кольцевая"
        Case Else: GetChartTypeName = "Другой тип (" & CStr(chartType) & ")"
    End Select
End Function

Private Sub ApplyColorsToChart(ByRef oChart As Chart)

    ' Проверка наличия данных на диаграмме
    If oChart.SeriesCollection.Count = 0 Then Exit Sub

    Dim oSeries     As Series
    Dim iSeries     As Long

    ' Перебор всех серий на диаграмме
    For iSeries = 1 To oChart.SeriesCollection.Count
        Set oSeries = oChart.SeriesCollection(iSeries)

        ' Безопасное извлечение диапазона значений из формулы серии
        ' Формула обычно имеет вид: =SERIES(, , "Имя", 1)
        ' Нам нужен третий аргумент (значения), который находится после второй запятой
        Dim formulaParts() As String
        Dim valueRangeAddress As String

        ' Разделяем формулу по запятым
        formulaParts = Split(oSeries.formula, ",")

        ' Проверяем, что в формуле достаточно частей для извлечения диапазона
        If UBound(formulaParts) >= 2 Then
            ' Извлекаем адрес диапазона, удаляя лишние пробелы
            valueRangeAddress = Trim(formulaParts(2))

            ' Проверка валидности диапазона
            Dim oRng As Range
            On Error Resume Next
            Set oRng = Range(valueRangeAddress)
            On Error GoTo 0

            ' Если диапазон найден, обрабатываем точки
            If Not oRng Is Nothing Then
                Dim iPoint As Long
                Dim oPoint As Point
                Dim lColor As Long
                Dim bFlag As Boolean

                ' Цикл по точкам текущей серии
                For iPoint = 1 To oSeries.Points.Count
                    ' Проверка, что индекс точки не выходит за границы диапазона
                    If iPoint <= oRng.Cells.Count Then
                        Set oPoint = oSeries.Points(iPoint)
                        With oRng.Cells(iPoint).DisplayFormat.Interior
                            lColor = .Color
                            bFlag = Not (.ColorIndex = -4142)
                        End With

                        ' Применение цвета из ячейки
                        With oPoint.format
                            .Fill.ForeColor.RGB = lColor
                            .Fill.Visible = bFlag
                            Select Case oChart.chartType
                                Case xlLineMarkers, 63, 66, 4
                                    .Line.Visible = bFlag
                                    .Line.ForeColor.RGB = lColor
                            End Select
                        End With
                    End If
                Next iPoint
            End If
        End If
    Next iSeries
End Sub

