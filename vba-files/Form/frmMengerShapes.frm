VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMengerShapes 
   Caption         =   "Менеджер фигур:"
   ClientHeight    =   8685.001
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18360
   OleObjectBlob   =   "frmMengerShapes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMengerShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const FILTER_REVERSE As String = "обратное выделение"

Private Sub btnAddText_Click()
    Dim i           As Integer
    Dim bFlafMsg    As Boolean
    Dim sNewName    As String

    sNewName = Application.InputBox("Введите новый текст!:", "Изменение текста:", Type:=2)
    If sNewName = vbNullString Or sNewName = "False" Then
        Exit Sub
    End If

    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                On Error Resume Next
                ActiveSheet.Shapes.Range(.List(i, 0)).TextFrame2.TextRange.TEXT = sNewName
                On Error GoTo 0
                bFlafMsg = True
            End If
        Next i
    End With
    If bFlafMsg Then
        Call refreshMainList
    Else
        Call MsgBox("Ни чего не выбрано в главном списке!", vbCritical, "Удаление:")
    End If
End Sub

Private Sub btnAppendText_Click()
    Dim i           As Integer
    Dim bFlafMsg    As Boolean
    Dim sNewName    As String

    sNewName = Application.InputBox("Введите новый текст для добовления!:", "Изменение текста:", Type:=2)
    If sNewName = vbNullString Or sNewName = "False" Then
        Exit Sub
    End If

    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                On Error Resume Next
                ActiveSheet.Shapes.Range(.List(i, 0)).TextFrame2.TextRange.TEXT = ActiveSheet.Shapes.Range(.List(i, 0)).TextFrame2.TextRange.TEXT & sNewName
                On Error GoTo 0
                bFlafMsg = True
            End If
        Next i
    End With
    If bFlafMsg Then
        Call refreshMainList
    Else
        Call MsgBox("Ни чего не выбрано в главном списке!", vbCritical, "Удаление:")
    End If
End Sub

Private Sub btnCopySize_Click()
    Dim i           As Integer
    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                txtHeight.Value = lMainList.List(i, 7)
                txtWidth.Value = lMainList.List(i, 8)
                Exit Sub
            End If
        Next i
    End With
End Sub
Private Sub btnPasteSize_Click()
    Dim i           As Integer
    Dim snHeight    As String
    Dim snWidth     As String
    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If IsNumeric(txtHeight.Value) And IsNumeric(txtWidth.Value) Then
                    snHeight = txtHeight.Value
                    snWidth = txtWidth.Value
                    .List(i, 7) = snHeight
                    .List(i, 8) = snWidth
                    ActiveSheet.Shapes(.List(i, 0)).Height = snHeight
                    ActiveSheet.Shapes(.List(i, 0)).Width = snWidth
                End If
            End If
        Next i
    End With
End Sub

Private Sub txtHeight_Change()
    Call checkVal(txtHeight)
End Sub
Private Sub txtWidth_Change()
    Call checkVal(txtWidth)
End Sub

Private Sub checkVal(ByRef oTxt As MSForms.TextBox)
    With oTxt
        If Application.DecimalSeparator = "," Then
            .Value = VBA.Replace(.Value, ".", ",")
        Else
            .Value = VBA.Replace(.Value, ",", ".")
        End If
        If IsNumeric(.Value) Then
            .BorderColor = &H80000006
        Else
            .BorderColor = &HC0C0FF
        End If
    End With
End Sub
'tools
Private Sub btnExpor_Click()
    Dim sExtFile    As String
    sExtFile = cmbExFile.Value
    If sExtFile = "PDF" Then
        Call MsgBox("Для PDF выбирите выгрузку диапазонов!", vbCritical, "Ошибка:")
        Exit Sub
    End If
    If TypeName(Selection) = "Range" Then
        Call MsgBox("Пожалуйста, выберите фигуры перед запуском инструмента!", vbCritical, "Ошибка:")
        Exit Sub
    End If

    Dim sNewNameFile As String
    Dim sFileName   As String
    sNewNameFile = ActiveWorkbook.Name & "_" & ActiveSheet.Name
    sFileName = Application.GetSaveAsFilename(Title:="Введите имя файла")
    If sFileName = vbNullString Or sFileName = "False" Then Exit Sub

    Dim cht         As ChartObject
    Dim oShape      As Shape
    For Each oShape In Selection.ShapeRange
        Set cht = ActiveSheet.ChartObjects.Add( _
                Left:=activeCell.Left, _
                Width:=oShape.Width, _
                Top:=activeCell.Top, _
                Height:=oShape.Height)

        cht.ShapeRange.Fill.Visible = msoFalse
        cht.ShapeRange.Line.Visible = msoFalse

        oShape.Copy
        cht.Activate
        ActiveChart.Paste

        cht.Chart.Export sFileName & "_" & oShape.Name & "." & sExtFile
        cht.Delete
        oShape.Select
    Next oShape
    Call MsgBox("Выбранные фигуры выгружены!", vbInformation, "Экспорт:")
End Sub

Private Sub btnImport_Click()
    If cmbLoadType.Value = "в лист" Then
        Call loadImgInSheet
    Else
        Me.Hide
        Call loadImInComment
        Me.Show
    End If
End Sub
Private Sub loadImInComment()
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Выделите начальную ячейку для вставки комментарий", vbCritical, "Ошибка:")
        Exit Sub
    End If

    Dim objFileDialog As Object
    Dim iCount      As Long

    Set objFileDialog = Application.FileDialog(msoFileDialogFilePicker)
    With objFileDialog
        .Show
        iCount = .SelectedItems.Count
        If iCount = 0 Then Exit Sub
        .Filters.Add "Images", "*.png;*.jpeg;*.jpg;*.gif;*.ico;*.cur;*.wmf"
    End With

    Dim bFInserAs   As Boolean
    Select Case MsgBox("Вставлять в столбик или в строчу?" & vbNewLine & "[ДА]   - в столбец" & vbNewLine & "[НЕТ] - в строчку", vbYesNoCancel + vbQuestion, "Вставка картинок:")
        Case vbYes
            bFInserAs = True
        Case vbNo
            bFInserAs = False
        Case vbCancel
            Exit Sub
    End Select

    Dim i           As Long
    Dim iCol        As Long
    Dim iRow        As Long
    Dim objCell     As Range
    Set objCell = activeCell
    For i = 1 To iCount
        If bFInserAs Then
            iRow = i - 1
        Else
            iCol = i - 1
        End If
        With objCell.Offset(iRow, iCol)
            .ClearComments
            Call .AddComment.Shape.Fill.UserPicture(objFileDialog.SelectedItems(i))
        End With
    Next i
End Sub
Private Sub loadImgInSheet()
    Dim PicList()   As Variant
    Dim PicFormat   As String
    Dim rng         As Range
    Dim sShape      As Shape
    Dim xColIndex   As Long
    Dim xRowIndex   As Long
    Dim lLoop       As Long

    On Error Resume Next
    PicList = Application.GetOpenFilename(PicFormat, MultiSelect:=True)
    If IsArrayDimensioned(PicList) Then
        xColIndex = Application.activeCell.Column
        If IsArray(PicList) Then
            xRowIndex = Application.activeCell.Row
            For lLoop = LBound(PicList) To UBound(PicList)
                Set rng = Cells(xRowIndex, xColIndex)
                Set sShape = ActiveSheet.Shapes.AddPicture(PicList(lLoop), msoFalse, msoCTrue, _
                        rng.Left, rng.Top, -1, -1)
                With sShape
                    .LockAspectRatio = msoTrue
                    If .Height > .Width Then
                        .Height = rng.Height
                    Else
                        .Width = rng.Width
                    End If
                    .Top = rng.MergeArea.Top + (rng.MergeArea.Height - .Height) / 2
                    .Left = rng.MergeArea.Left + (rng.MergeArea.Width - .Width) / 2
                End With
                xRowIndex = xRowIndex + 1
            Next
        End If
        Call refreshMainList
    End If
End Sub

Private Sub btnRename_Click()
    Dim i           As Integer
    Dim j           As Integer
    Dim bFlafMsg    As Boolean
    Dim sNewName    As String

    sNewName = Application.InputBox("Введите новое название!:", "Переименовать фигуры:", Type:=2)
    If sNewName = vbNullString Or sNewName = "False" Then
        Exit Sub
    End If

    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                j = j + 1
                ActiveSheet.Shapes.Range(.List(i, 0)).Name = sNewName & "_" & j
                bFlafMsg = True
            End If
        Next i
    End With
    If bFlafMsg Then
        Call refreshMainList
    Else
        Call MsgBox("Ни чего не выбрано в главном списке!", vbCritical, "Удаление:")
    End If
End Sub
Private Sub btnShowHide_Click()
    Dim i           As Long
    Dim bFlafMsg    As Boolean

    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                ActiveSheet.Shapes.Range(.List(i, 0)).Visible = Not VBA.CBool(.List(i, 9))
                bFlafMsg = True
            End If
        Next i
    End With
    If bFlafMsg Then
        Call refreshMainList
    Else
        Call MsgBox("Ни чего не выбрано в главном списке!", vbCritical, "Удаление:")
    End If
End Sub

Private Sub btnCopy_Click()
    Dim arr()       As String
    arr = getSelectedItemMainList(, False)
    If IsArrayDimensioned(arr) Then
        Dim iCopyNumber As Integer
        Dim i       As Integer
        iCopyNumber = Application.InputBox("Введите количество копий:", "Копирователь листов", Default:=1, Type:=1)
        If iCopyNumber <= 0 Or iCopyNumber = False Then
            Exit Sub
        End If

        ActiveSheet.Shapes.Range(arr).Select
        Selection.Copy
        For i = 1 To iCopyNumber
            ActiveSheet.Paste
        Next i

        Call refreshMainList
    Else
        Call MsgBox("Ни чего не выбрано в главном списке!", vbCritical, "Удаление:")
    End If
End Sub

Private Sub btnDelete_Click()
    Dim arr()       As String
    arr = getSelectedItemMainList
    If IsArrayDimensioned(arr) Then
        If MsgBox("Продолжить удаление фигур?" & vbNewLine & "Данную операцию нельзя отменить!", vbYesNo + vbQuestion, "Удаление:") = vbYes Then

            Dim i   As Long
            Dim iCount As Long
            iCount = UBound(arr, 1)
            For i = 0 To iCount
                With ActiveSheet.Shapes.Range(arr(i))
                    If .Type <> MsoShapeType.msoComment Then .Delete
                End With
            Next i
            Call refreshMainList
        End If
    Else
        Call MsgBox("Ни чего не выбрано в главном списке!", vbCritical, "Удаление:")
    End If
End Sub

Private Sub btnSortName_Click()
    Call sortColumnList(lMainList, btnSortName, 0, False)
End Sub
Private Sub btnSortType_Click()
    Call sortColumnList(lMainList, btnSortType, 1, False)
End Sub
Private Sub btnSortVisible_Click()
    Call sortColumnList(lMainList, btnSortVisible, 2, False)
End Sub
Private Sub btnSortMacroName_Click()
    Call sortColumnList(lMainList, btnSortMacroName, 3, False)
End Sub
Private Sub btnSortText_Click()
    Call sortColumnList(lMainList, btnSortText, 4, False)
End Sub
Private Sub btnSortTop_Click()
    Call sortColumnList(lMainList, btnSortTop, 5, True)
End Sub
Private Sub btnSortLeft_Click()
    Call sortColumnList(lMainList, btnSortLeft, 6, True)
End Sub
Private Sub btnSorHeight_Click()
    Call sortColumnList(lMainList, btnSorHeight, 7, True)
End Sub
Private Sub btnSorWidth_Click()
    Call sortColumnList(lMainList, btnSorWidth, 8, True)
End Sub
'end tools
Private Sub lMainList_Change()
    Dim arr()       As String
    arr = getSelectedItemMainList(0, False)
    If IsArrayDimensioned(arr) Then
        ActiveSheet.Shapes.Range(arr).Select
    End If
End Sub

Private Sub listFilters_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With listFilters
        If .ListIndex < 0 Then Exit Sub
        Select Case .Value
            Case FILTER_REVERSE: Call reversSelected
        End Select
    End With
End Sub

Private Sub reversSelected()
    Dim i           As Long
    With lMainList
        For i = 0 To .ListCount - 1
            .Selected(i) = Not .Selected(i)
        Next i
    End With
End Sub

Private Sub listFilters_Change()
    Dim i           As Long
    Dim j           As Long
    Dim byCol       As Byte
    Dim sFilter     As String
    Dim arr()       As String

    sFilter = listFilters.List(listFilters.ListIndex, 0)
    byCol = 1
    Select Case sFilter
        Case "видим": byCol = 2
        Case "скрыт": byCol = 2
        Case "с макросом":
            byCol = 10
            sFilter = 1
        Case "без макроса":
            byCol = 10
            sFilter = 0
    End Select

    With lMainList
        For i = 0 To .ListCount - 1
            .Selected(i) = False
            If .List(i, byCol) = sFilter Then
                If .List(i, 3) <> "скрыт" And .List(i, 0) = ActiveSheet.Name And .List(i, 2) <> "Comment" Then
                    ReDim Preserve arr(0 To j) As String
                    arr(j) = .List(i, 1)
                    j = j + 1
                End If
                .Selected(i) = True
            End If
        Next i
    End With
    If IsArrayDimensioned(arr) Then
        ActiveSheet.Shapes.Range(arr).Select
    End If
End Sub
Private Sub refreshMainList()
    Dim arr()       As Variant
    Dim shp         As Shape
    Dim i           As Long
    Dim ItemCol     As Long
    Dim arrFilter() As Variant
    Dim sVal        As String
    Dim oColl       As Collection

    ReDim Preserve arrFilter(1 To 2, 1 To 5)

    arrFilter(1, 1) = "видим"
    arrFilter(1, 2) = "скрыт"
    arrFilter(1, 3) = "с макросом"
    arrFilter(1, 4) = "без макроса"
    ItemCol = 4
    Set oColl = New Collection
    For Each shp In ActiveSheet.Shapes
        With shp
            i = i + 1
            ReDim Preserve arr(1 To 11, 1 To i)
            arr(1, i) = .Name
            arr(2, i) = getTypeShape(.Adjustments.Parent.Type)

            sVal = arr(2, i)
            If addUnic(oColl, ItemCol, sVal) Then
                ItemCol = ItemCol + 1
                ReDim Preserve arrFilter(1 To 2, 1 To ItemCol)
                arrFilter(1, ItemCol) = sVal
                arrFilter(2, ItemCol) = 1
            Else
                arrFilter(2, oColl.item(sVal) + 1) = arrFilter(2, oColl.item(sVal) + 1) + 1
            End If


            arr(10, i) = .Visible
            If .Visible Then
                arr(3, i) = "видим"
                arrFilter(2, 1) = arrFilter(2, 1) + 1
            Else
                arr(3, i) = "скрыт"
                arrFilter(2, 2) = arrFilter(2, 2) + 1
            End If
            arr(4, i) = .OnAction
            If .OnAction <> vbNullString Then
                arr(11, i) = 1
                arrFilter(2, 3) = arrFilter(2, 3) + 1
            Else
                arr(11, i) = 0
                arrFilter(2, 4) = arrFilter(2, 4) + 1
            End If

            On Error Resume Next
            arr(5, i) = VBA.Left(.TextFrame2.TextRange.TEXT, 50)
            On Error GoTo 0

            arr(6, i) = VBA.Round(.Top, 3)
            arr(7, i) = VBA.Round(.Left, 3)
            arr(8, i) = VBA.Round(.Height, 3)
            arr(9, i) = VBA.Round(.Width, 3)
        End With
    Next shp

    If IsArrayDimensioned(arr, 1) Then
        With lMainList
            arr = WorksheetFunction.Transpose(arr)
            If IsArrayDimensioned(arr, 2) Then
                arr = SortArray(arr, 1)
                arr = SortArray(arr, 2)
                .List = arr
            Else
                ReDim arrVal(0 To 0, 0 To 10)
                For i = 1 To 11
                    arrVal(0, i - 1) = arr(i)
                Next i
                .List = arrVal
            End If
        End With
    Else
        lMainList.Clear
    End If
    listFilters.List = WorksheetFunction.Transpose(arrFilter)

    Erase arrVal
    Erase arr
    Erase arrFilter
    Set oColl = Nothing
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub UserForm_Activate()
    Call refreshMainList
End Sub
Private Sub UserForm_Initialize()

    ' Центрирование формы
    Call CenterUserForm(Me)
    
    With cmbExFile
        .AddItem "GIF"
        .AddItem "PNG"
        .AddItem "JPG"
        .AddItem "BMP"
        .AddItem "PDF"
        .ListIndex = 0
    End With
    With cmbLoadType
        .AddItem "в лист"
        .AddItem "в коммент."
        .ListIndex = 0
    End With

    Const W         As Byte = 32
    Const h         As Byte = 32
    With Application.CommandBars
        btnShowHide.Picture = .GetImageMso("GroupWindow", W, h)
        btnCopy.Picture = .GetImageMso("NewOutlookDataFile", W, h)
        btnDelete.Picture = .GetImageMso("SketchpadToolDeleteBackground", W, h)
        btnRename.Picture = .GetImageMso("FillOutThisForm", W, h)
        btnExpor.Picture = .GetImageMso("CheckOutFile", W, h)
        btnImport.Picture = .GetImageMso("CheckInMenu", W, h)
        btnAddText.Picture = .GetImageMso("PasteDestinationFormatting", W, h)
        btnAppendText.Picture = .GetImageMso("PasteDestinationTheme", W, h)
        btnCopySize.Picture = .GetImageMso("ContactCardCopy", W, h)
        btnPasteSize.Picture = .GetImageMso("ContactCardPaste", W, h)
    End With
End Sub
'functions
Private Function addUnic(ByRef oColl As Collection, ByVal ItemCol As Long, ByVal sKey As String) As Boolean
    On Error Resume Next
    oColl.Add ItemCol, sKey
    If Err.Number = 0 Then addUnic = True
    Err.Clear
    On Error GoTo 0
End Function
Private Function getSelectedItemMainList(Optional iCol As Byte = 0, Optional bFlagComment As Boolean = True) As String()
    Dim i           As Long
    Dim j           As Long
    Dim arr()       As String
    With lMainList
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                If bFlagComment Then
                    ReDim Preserve arr(0 To j) As String
                    arr(j) = .List(i, iCol)
                    j = j + 1
                Else
                    If .List(i, 2) <> "скрыт" And .List(i, 1) <> "Comment" Then
                        ReDim Preserve arr(0 To j) As String
                        arr(j) = .List(i, iCol)
                        j = j + 1
                    End If
                End If
            End If
        Next i
    End With
    getSelectedItemMainList = arr
End Function
Private Function getTypeShape(ByVal sMsoType As MsoShapeType) As String
    Select Case sMsoType
        Case -2: getTypeShape = "Mixed shape type"
        Case 1: getTypeShape = "AutoShape"
        Case 2: getTypeShape = "Callout"
        Case 3: getTypeShape = "Chart"
        Case 4: getTypeShape = "Comment"
        Case 5: getTypeShape = "Freeform"
        Case 6: getTypeShape = "Group"
        Case 7: getTypeShape = "Embedded OLE object"
        Case 8: getTypeShape = "Form control"
        Case 9: getTypeShape = "Line"
        Case 10: getTypeShape = "Linked OLE object"
        Case 11: getTypeShape = "Linked picture"
        Case 12: getTypeShape = "OLE control object"
        Case 13: getTypeShape = "Picture"
        Case 14: getTypeShape = "Placeholder"
        Case 15: getTypeShape = "Text effect"
        Case 16: getTypeShape = "Media"
        Case 17: getTypeShape = "Text box"
        Case 18: getTypeShape = "Script anchor"
        Case 19: getTypeShape = "Table"
        Case 20: getTypeShape = "Canvas"
        Case 21: getTypeShape = "Diagram"
        Case 22: getTypeShape = "Ink"
        Case 23: getTypeShape = "Ink comment"
        Case 24: getTypeShape = "SmartArt graphic"
        Case 25: getTypeShape = "Slicer"
        Case 26: getTypeShape = "Web video"
        Case 27: getTypeShape = "Content Office Add-in"
        Case 28: getTypeShape = "Graphic"
        Case 29: getTypeShape = "Linked graphic"
        Case 30: getTypeShape = "3D model"
        Case 31: getTypeShape = "Linked 3D model"

        Case Else: getTypeShape = "Unknown Type"
    End Select
End Function









