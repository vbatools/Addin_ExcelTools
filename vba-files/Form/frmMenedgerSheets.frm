VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheets 
   Caption         =   "Менеджер листов:"
   ClientHeight    =   10320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19695
   OleObjectBlob   =   "frmMenedgerSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' UserForm     :   frmSheetsManager
' Description  :   Интерфейс для управления, фильтрации и навигации по листам активной книги.
'                  Поддерживает сортировку, поиск, групповые операции и анализ свойств листов.
' Author       :   VBATools
' Copyright    :   Apache License
' Created      :   15-06-2026 15:36:42
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

'--------------------------------------------------------------------------------
' Константы фильтров и отображаемых значений
'--------------------------------------------------------------------------------
' Фильтры списка
Private Const FILTER_ALL As String = "все"
Private Const FILTER_NONE As String = "ничего"
Private Const FILTER_REVERSE As String = "обратное выделение"
Private Const FILTER_VISIBLE As String = "видимые"
Private Const FILTER_HIDDEN As String = "скрытые"
Private Const FILTER_VERY_HIDDEN As String = "супер скрыт"
Private Const FILTER_PROTECTED As String = "защищенные"
Private Const FILTER_UNPROTECTED As String = "не защищенные"
Private Const FILTER_EMPTY As String = "пустые"
Private Const FILTER_NOT_EMPTY As String = "не пустые"
Private Const FILTER_SHEETS As String = "листы"
Private Const FILTER_CHARTS As String = "диаграммы"
Private Const FILTER_TAB_COLOR As String = "цветные ярлыки"
Private Const FILTER_TAB_NOT_COLOR As String = "без цветные ярлыки"

' Значения свойств листов
Private Const VALUE_VISIBLE As String = "видим"
Private Const VALUE_HIDDEN As String = "скрыт"
Private Const VALUE_VERY_HIDDEN As String = "супер скрыт"
Private Const VALUE_PROTECTED As String = "защита"
Private Const VALUE_SHEET_TYPE As String = "лист"
Private Const VALUE_CHART_TYPE As String = "график"
Private Const VALUE_DIALOG_TYPE As String = "диалог"
Private Const VALUE_MACRO_TYPE As String = "макро"

' Варианты регистра текста
Private Const CASE_UPPER As String = "ВСЕ ПРОПИСНЫЕ"
Private Const CASE_LOWER As String = "все строчные"
Private Const CASE_SENTENCE As String = "Как в предложениях"
Private Const CASE_TITLE As String = "Начинать С Прописных"

Private Enum typeMoveItem
    miFerst
    miUp
    miDown
    miLast
End Enum

Private Sub btnColors_Click()
    Dim lSeletedColor As Long
    lSeletedColor = GetColorFromDialog()
    If lSeletedColor = -1 Then Exit Sub
    Dim i           As Long
    With listSheets
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                ActiveWorkbook.Sheets(.List(i, 1)).Tab.Color = lSeletedColor
            End If
        Next i
    End With
End Sub

'--------------------------------------------------------------------------------
' кнопки инструментов
'--------------------------------------------------------------------------------
Private Sub btnContent_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsLinks.Show
    Me.Show
End Sub

Private Sub btnExport_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsExportSheets.Show
    Me.Show
End Sub

Private Sub btnFreezePane_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsFreezePane.Show
    Me.Show
End Sub

Private Sub btnGrouping_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsGroop.Show
    Me.Show
End Sub

Private Sub btnImport_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsImpotrSheets.Show
    Call refreshForm
    Me.Show
End Sub

Private Sub btnSync_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsZoom.Show
    Me.Show
End Sub

Private Sub btnSyncSheets_Click()
    If CheckProtectStructure() Then Exit Sub

    Dim Sh          As Object
    Set Sh = ActiveSheet
    Dim shVisible   As XlSheetVisibility
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If MsgBox("Произвести синхронизацию листов?" & vbNewLine & "Выбран лист: " & Sh.Name, vbYesNo + vbQuestion, "Синхрогизация листов:") = vbNo Then Exit Sub

    Call DisableApplicationSettings

    shVisible = Sh.Visible
    Sh.Visible = XlSheetVisibility.xlSheetVisible
    Sh.Activate

    Dim scrolRow    As Long
    Dim scrolLeft   As Long
    Dim sAddresSlect As String
    Dim iZoom       As Integer

    With Sh
        scrolRow = ActiveWindow.ScrollRow
        scrolLeft = ActiveWindow.ScrollColumn
        sAddresSlect = ActiveWindow.RangeSelection.Address
        iZoom = ActiveWindow.Zoom
        .Visible = shVisible
    End With

    Dim i           As Long
    Dim iSheetCount As Long
    Dim shCurent    As Worksheet
    With listSheets
        For i = 0 To .ListCount - 1
            If .Selected(i) And .List(i, 2) = VALUE_SHEET_TYPE And .List(i, 1) <> Sh.Name Then
                Set shCurent = ActiveWorkbook.Sheets(.List(i, 1))
                shVisible = shCurent.Visible
                shCurent.Visible = XlSheetVisibility.xlSheetVisible
                shCurent.Activate
                ActiveWindow.ScrollRow = scrolRow
                ActiveWindow.ScrollColumn = scrolLeft
                shCurent.Range(sAddresSlect).Select
                ActiveWindow.Zoom = iZoom
                shCurent.Visible = shVisible
                iSheetCount = iSheetCount + 1
            End If
        Next i
    End With
    Call RestoreApplicationSettings
    If iSheetCount > 0 Then Call MsgBox("Синхронизировано " & iSheetCount & " листов!", vbInformation)
End Sub

Private Sub btnProtect_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsProtected.Show
    Me.Show
End Sub

Private Sub btnSort_Click()
    If CheckProtectStructure() Then Exit Sub
    Me.Hide
    frmMenedgerSheetsSortSheets.Show
    Me.Show
End Sub

Private Sub btnCopy_Click()
    If CheckProtectStructure() Then Exit Sub
    Dim iCopyNumber As Integer
    iCopyNumber = Application.InputBox("Введите количество копий:", "Копирователь листов", Default:=1, Type:=1)
    If iCopyNumber < 1 Then Exit Sub

    Dim i           As Long
    Dim j           As Byte
    Dim Sh          As Object
    Dim shVisible   As XlSheetVisibility
    Call DisableApplicationSettings
    With listSheets
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                Set Sh = ActiveWorkbook.Sheets(.List(i, 1))
                shVisible = Sh.Visible
                For j = 1 To iCopyNumber
                    Sh.Visible = XlSheetVisibility.xlSheetVisible
                    Sh.Copy After:=Sh
                    ActiveSheet.Visible = shVisible
                Next j
                Sh.Visible = shVisible
            End If
        Next i
    End With
    Call refreshForm
    Call RestoreApplicationSettings
End Sub


Private Sub btnDelete_Click()
    If CheckProtectStructure() Then Exit Sub
    If MsgBox("Удалить выбранные листы?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim i           As Long
    Call DisableApplicationSettings
    With listSheets
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                With ActiveWorkbook.Sheets(.List(i, 1))
                    .Visible = XlSheetVisibility.xlSheetVisible
                    .Delete
                End With
                Call .RemoveItem(i)
            End If
        Next i
    End With
    Call refreshForm
    Call RestoreApplicationSettings
End Sub

Private Sub btnShowHide_Click()
    If CheckProtectStructure() Then Exit Sub
    Dim i           As Long
    Dim sTypeVisible As String
    Call DisableApplicationSettings
    With listSheets
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                With ActiveWorkbook.Sheets(.List(i, 1))
                    Select Case .Visible
                        Case XlSheetVisibility.xlSheetHidden
                            .Visible = XlSheetVisibility.xlSheetVeryHidden
                            sTypeVisible = VALUE_VERY_HIDDEN
                        Case XlSheetVisibility.xlSheetVeryHidden
                            .Visible = XlSheetVisibility.xlSheetVisible
                            sTypeVisible = VALUE_VISIBLE
                        Case XlSheetVisibility.xlSheetVisible
                            .Visible = XlSheetVisibility.xlSheetHidden
                            sTypeVisible = VALUE_HIDDEN
                    End Select
                End With
                .List(i, 3) = sTypeVisible
            End If
        Next i
    End With
    Call RestoreApplicationSettings
End Sub

Private Sub btnAddNewSheet_Click()
    If CheckProtectStructure() Then Exit Sub
    Sheets.Add After:=ActiveSheet
    Call refreshForm
End Sub

Private Sub btnMoveAsList_Click()
    If CheckProtectStructure() Then Exit Sub
    Dim i           As Long
    Dim shVisible   As XlSheetVisibility
    With listSheets
        For i = .ListCount - 1 To 0 Step -1
            With ActiveWorkbook.Sheets(.List(i, 1))
                shVisible = .Visible
                If shVisible <> xlSheetVisible Then .Visible = xlSheetVisible
                .Move After:=ActiveWorkbook.Sheets(i + 1)
                .Visible = shVisible
            End With
        Next i
    End With
End Sub

'--------------------------------------------------------------------------------
' кнопки перемещения листов
'--------------------------------------------------------------------------------
Private Sub btnFirst_Click()
    Call moveItemList(listSheets, miFerst)
End Sub

Private Sub btnUp_Click()
    Call moveItemList(listSheets, miUp)
End Sub

Private Sub btnDown_Click()
    Call moveItemList(listSheets, miDown)
End Sub

Private Sub btnLast_Click()
    Call moveItemList(listSheets, miLast)
End Sub

Private Sub moveItemList(ByRef List As MSForms.listBox, ByVal tpMove As typeMoveItem)
    If CheckProtectStructure() Then Exit Sub
    Dim i           As Long
    Dim Sh          As Variant
    Dim index       As Integer
    Dim wb          As Workbook
    Dim k           As Long

    Dim iCount      As Long
    Dim itemFerst   As Long
    Dim itemLast    As Long
    Dim itemStep    As Long


    Set wb = ActiveWorkbook
    With List
        iCount = .ListCount - 1
        ReDim arr(0 To iCount)
        itemFerst = 0
        itemLast = iCount
        itemStep = 1
        If tpMove = miFerst Or tpMove = miDown Then
            itemFerst = iCount
            itemLast = 0
            itemStep = -1
        End If

        For i = itemFerst To itemLast Step itemStep
            If .Selected(i) Then
                Set Sh = wb.Sheets(.List(i, 1))
                index = Sh.index
                ' Определяем целевой индекс
                Select Case tpMove
                    Case typeMoveItem.miFerst, typeMoveItem.miUp
                        If index = 1 Then GoTo SkipItem
                    Case typeMoveItem.miDown, typeMoveItem.miLast
                        If index = wb.Sheets.Count Then GoTo SkipItem
                End Select
                ' Выполняем перемещение
                arr(k) = MoveSheet(wb, Sh, tpMove)
                k = k + 1
SkipItem:
            End If
        Next i
    End With

    If k = 0 Then Exit Sub
    Call refreshForm
    For i = 0 To k - 1
        Select Case tpMove
            Case typeMoveItem.miFerst
                List.Selected(i) = True
            Case typeMoveItem.miLast
                List.Selected(iCount - i) = True
            Case Else
                List.Selected(arr(i) - 1) = True
        End Select
    Next i
End Sub

' Вспомогательная процедура перемещения листа
Private Function MoveSheet(ByRef wb As Workbook, ByRef Sh As Variant, _
        ByVal moveType As typeMoveItem) As Long

    Dim targetSheet As Variant
    Dim targetVisible As XlSheetVisibility
    Dim shVisible   As XlSheetVisibility

    Select Case moveType
        Case typeMoveItem.miFerst
            Set targetSheet = wb.Sheets(1)
        Case typeMoveItem.miUp
            Set targetSheet = wb.Sheets(Sh.index - 1)
        Case typeMoveItem.miDown
            Set targetSheet = wb.Sheets(Sh.index + 1)
        Case typeMoveItem.miLast
            Set targetSheet = wb.Sheets(wb.Sheets.Count)
    End Select


    targetVisible = targetSheet.Visible
    shVisible = Sh.Visible

    ' Делаем листы видимыми для перемещения
    If targetVisible <> xlSheetVisible Then targetSheet.Visible = xlSheetVisible
    If shVisible <> xlSheetVisible Then Sh.Visible = xlSheetVisible
    ' Перемещаем
    Select Case moveType
        Case typeMoveItem.miDown
            targetSheet.Move before:=Sh
        Case typeMoveItem.miLast
            Sh.Move After:=targetSheet
        Case Else
            Sh.Move before:=targetSheet
    End Select
    MoveSheet = Sh.index

    ' Восстанавливаем видимость
    Sh.Visible = shVisible
    targetSheet.Visible = targetVisible

End Function

'--------------------------------------------------------------------------------
' кнопки сортировки
'--------------------------------------------------------------------------------
Private Sub btnSortNum_Click()
    Call SortColumnList(listSheets, btnSortNum, 0, True)
End Sub

Private Sub btnSortName_Click()
    Call SortColumnList(listSheets, btnSortName, 1, False, True)
End Sub

Private Sub btnSortType_Click()
    Call SortColumnList(listSheets, btnSortType, 2)
End Sub

Private Sub btnSortVisible_Click()
    Call SortColumnList(listSheets, btnSortVisible, 3)
End Sub

Private Sub btnSortProtect_Click()
    Call SortColumnList(listSheets, btnSortProtect, 4)
End Sub

Private Sub btnSortRange_Click()
    Call SortColumnList(listSheets, btnSortRange, 5)
End Sub

Private Sub btnSortCells_Click()
    Call SortColumnList(listSheets, btnSortCells, 6, True)
End Sub

Private Sub btnSortInfo_Click()
    Call SortColumnList(listSheets, btnSortInfo, 8, True)
End Sub

Private Sub listRegistr_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If CheckProtectStructure() Then Exit Sub
    Dim i           As Long
    Dim sName       As String
    Call DisableApplicationSettings
    With listSheets
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                With ActiveWorkbook.Sheets(.List(i, 1))
                    Select Case listRegistr.List(listRegistr.ListIndex, 0)
                        Case CASE_UPPER
                            sName = caseString(.Name, tpUCase)
                        Case CASE_LOWER
                            sName = caseString(.Name, tpLCase)
                        Case CASE_SENTENCE
                            sName = caseString(.Name, tpAsString)
                        Case CASE_TITLE
                            sName = caseString(.Name, tpAllWorldUCase)
                    End Select
                    .Name = sName
                End With
                .List(i, 1) = sName
            End If
        Next i
    End With
    Call RestoreApplicationSettings
End Sub

Private Sub cmbSheets_Change()
    With cmbSheets
        If .ListIndex < 0 Then
            If .Value = vbNullString Then
                Call SelectedItemListSheets(listSheets, vbNullString, 1)
            Else
                Call SelectedItemListSheets(listSheets, "*" & .Value & "*", 1)
            End If
        Else
            ' Оптимизация: Direct access по имени без необходимости предварительной активации
            ActiveWorkbook.Sheets(.Value).Activate
            Call SelectedItemListSheets(listSheets, cmbSheets.Value, 1)
        End If
    End With
End Sub

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
            Case FILTER_ALL: Call SelectedItemListSheets(listSheets, "*", 1)
            Case FILTER_NONE: Call SelectedItemListSheets(listSheets, vbNullString, 1)
            Case FILTER_REVERSE: Call reversSelected
            Case FILTER_VISIBLE: Call SelectedItemListSheets(listSheets, VALUE_VISIBLE, 3)
            Case FILTER_HIDDEN: Call SelectedItemListSheets(listSheets, VALUE_HIDDEN, 3)
            Case FILTER_VERY_HIDDEN: Call SelectedItemListSheets(listSheets, VALUE_VERY_HIDDEN, 3)
            Case FILTER_PROTECTED: Call SelectedItemListSheets(listSheets, VALUE_PROTECTED, 4)
            Case FILTER_UNPROTECTED: Call SelectedItemListSheets(listSheets, vbNullString, 4)
            Case FILTER_EMPTY: Call SelectedItemListSheets(listSheets, vbNullString, 6)
            Case FILTER_NOT_EMPTY: Call SelectedItemListSheets(listSheets, "*[1-9]*", 6)
            Case FILTER_SHEETS: Call SelectedItemListSheets(listSheets, VALUE_SHEET_TYPE, 2)
            Case FILTER_CHARTS: Call SelectedItemListSheets(listSheets, VALUE_CHART_TYPE, 2)
            Case FILTER_TAB_COLOR: Call SelectedItemListSheets(listSheets, "*[1-9]*", 9)
            Case FILTER_TAB_NOT_COLOR: Call SelectedItemListSheets(listSheets, 0, 9)
        End Select
    End With
End Sub

Private Sub reversSelected()
    Dim i           As Long
    With listSheets
        For i = 0 To .ListCount - 1
            .Selected(i) = Not .Selected(i)
        Next i
    End With
End Sub

Private Sub listSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With listSheets
        If .ListIndex < 0 Then Exit Sub
        ' Активация выбранного листа (имя хранится во втором столбце с индексом 1)
        ActiveWorkbook.Sheets(.List(.ListIndex, 1)).Activate
    End With
End Sub

'--------------------------------------------------------------------------------
' Sub: UserForm_Initialize
' Purpose: Инициализация формы при запуске. Настраивает расположение, наполняет
'          списки фильтров и регистров, а также назначает иконки кнопкам из галереи MSO.
'--------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Const W         As Byte = 32
    Const h         As Byte = 32

    ' Центрирование формы
    Call CenterUserForm(Me)

    ' Настройка списка регистров (текстовые варианты)
    With listRegistr
        .AddItem CASE_UPPER
        .AddItem CASE_LOWER
        .AddItem CASE_SENTENCE
        .AddItem CASE_TITLE
    End With

    ' Настройка списка фильтров
    With listFilters
        .AddItem FILTER_ALL
        .AddItem FILTER_NONE
        .AddItem FILTER_REVERSE
        .AddItem FILTER_VISIBLE
        .AddItem FILTER_HIDDEN
        .AddItem FILTER_VERY_HIDDEN
        .AddItem FILTER_PROTECTED
        .AddItem FILTER_UNPROTECTED
        .AddItem FILTER_EMPTY
        .AddItem FILTER_NOT_EMPTY
        .AddItem FILTER_SHEETS
        .AddItem FILTER_CHARTS
        .AddItem FILTER_CHARTS
        .AddItem FILTER_TAB_COLOR
        .AddItem FILTER_TAB_NOT_COLOR
    End With

    ' Назначение иконок кнопкам через CommandBars.GetImageMso
    With Application.CommandBars
        Set btnContent.Picture = .GetImageMso("FunctionsMathTrigInsertGallery", W, h)
        Set btnSort.Picture = .GetImageMso("SortDialog", W, h)
        Set btnShowHide.Picture = .GetImageMso("GroupWindow", W, h)
        Set btnCopy.Picture = .GetImageMso("NewOutlookDataFile", W, h)
        Set btnDelete.Picture = .GetImageMso("SketchpadToolDeleteBackground", W, h)
        Set btnProtect.Picture = .GetImageMso("DatabaseMakeMdeFile", W, h)
        Set btnSync.Picture = .GetImageMso("MailMergeCreateList", W, h)
        Set btnExport.Picture = .GetImageMso("ExportExcel", W, h)
        Set btnImport.Picture = .GetImageMso("ImportExcel", W, h)
        Set btnGrouping.Picture = .GetImageMso("OutlineShowDetail", W - 16, h - 16)
        Set btnFreezePane.Picture = .GetImageMso("DatasheetView", W, h)
        Set btnColors.Picture = .GetImageMso("QuickStepTemplateCategorizeAndMove", W - 16, h - 16)
        Set btnSyncSheets.Picture = .GetImageMso("DataRefreshAll", W, h)
        Set btnAddNewSheet.Picture = .GetImageMso("FilesToolAddFiles", W, h)
        Set btnMoveAsList.Picture = .GetImageMso("AfterUpdate", W, h)
    End With

    Call refreshForm
End Sub

'--------------------------------------------------------------------------------
' Sub: refreshForm
' Purpose: Основная процедура обновления данных формы. Сканирует все листы активной
'          книги, собирает их свойства (имя, тип, видимость, защита, заполненность)
'          и загружает их в массив listSheets. Также формирует строку визуализации
'          заполненности листа.
'--------------------------------------------------------------------------------
Private Sub refreshForm()
    If Workbooks.Count = 0 Then Exit Sub
    
    Dim i           As Integer
    Dim iCount      As Integer
    Dim Sh          As Variant
    Dim iMax        As Long
    Dim shActive    As Variant

    Call DisableApplicationSettings

    iCount = ActiveWorkbook.Sheets.Count
    ReDim arrSheets(1 To iCount, 1 To 2) As String
    ReDim arrListSheets(1 To iCount, 1 To 10) As String
    Set shActive = ActiveSheet
    For i = 1 To iCount
        Set Sh = ActiveWorkbook.Sheets(i)

        With Sh
            ' Базовая информация
            arrSheets(i, 1) = .Name
            arrSheets(i, 2) = i
            arrListSheets(i, 1) = i
            arrListSheets(i, 2) = arrSheets(i, 1)

            ' Определение типа листа
            Select Case TypeName(Sh)
                Case "DialogSheet"
                    arrListSheets(i, 3) = VALUE_DIALOG_TYPE
                Case "Chart"
                    arrListSheets(i, 3) = VALUE_CHART_TYPE
                Case Else
                    Select Case .Type
                        Case XlSheetType.xlChart
                            arrListSheets(i, 3) = VALUE_CHART_TYPE
                        Case XlSheetType.xlDialogSheet
                            arrListSheets(i, 3) = VALUE_DIALOG_TYPE
                        Case XlSheetType.xlExcel4IntlMacroSheet, XlSheetType.xlExcel4MacroSheet
                            arrListSheets(i, 3) = VALUE_MACRO_TYPE
                        Case XlSheetType.xlWorksheet
                            arrListSheets(i, 3) = VALUE_SHEET_TYPE
                            ' Адрес используемого диапазона
                            arrListSheets(i, 6) = .UsedRange.Address(RowAbsolute:=False, columnAbsolute:=False)
                            ' Подсчет заполненных ячеек
                            arrListSheets(i, 7) = WorksheetFunction.CountA(.UsedRange)

                            If arrListSheets(i, 7) = 0 Then
                                arrListSheets(i, 7) = vbNullString
                            Else
                                iMax = iMax + arrListSheets(i, 7)
                            End If
                    End Select
            End Select

            ' Определение видимости
            Select Case .Visible
                Case XlSheetVisibility.xlSheetHidden
                    arrListSheets(i, 4) = VALUE_HIDDEN
                Case XlSheetVisibility.xlSheetVeryHidden
                    arrListSheets(i, 4) = VALUE_VERY_HIDDEN
                Case XlSheetVisibility.xlSheetVisible
                    arrListSheets(i, 4) = VALUE_VISIBLE
            End Select

            ' Проверка защиты
            If .ProtectContents Then arrListSheets(i, 5) = VALUE_PROTECTED
            arrListSheets(i, 10) = VBA.CLng(.Tab.Color)
        End With
    Next i

    ' Формирование визуальной шкалы заполненности (столбец 8)
    If iMax > 0 Then
        For i = 1 To iCount
            If IsNumeric(arrListSheets(i, 7)) Then
                arrListSheets(i, 9) = arrListSheets(i, 7) / iMax
                arrListSheets(i, 8) = VBA.String$(VBA.Int(arrListSheets(i, 9) * 15), "|")
                arrListSheets(i, 9) = VBA.Round(arrListSheets(i, 9) * 100, 2)
            End If
        Next i
    End If

    ' Загрузка данных в контролы
    listSheets.List = arrListSheets
    cmbSheets.List = arrSheets
    shActive.Activate
    Call RestoreApplicationSettings
End Sub

