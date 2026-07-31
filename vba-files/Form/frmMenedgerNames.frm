VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerNames 
   Caption         =   "Менеджер имен:"
   ClientHeight    =   8310.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19455
   OleObjectBlob   =   "frmMenedgerNames.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerNames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------
' Константы фильтров и отображаемых значений
'--------------------------------------------------------------------------------
' Фильтры списка
Private Const FILTER_ALL As String = "все"
Private Const FILTER_NONE As String = "ничего"
Private Const FILTER_REVERSE As String = "обратное выделение"
Private Const FILTER_VISIBLE As String = "видимые"
Private Const FILTER_HIDDEN As String = "скрытые"
Private Const FILTER_ERROR As String = "с ошибками"
Private Const FILTER_OUT_ERROR As String = "без ошибок"
Private Const FILTER_WB As String = "в книге"
Private Const FILTER_SH As String = "в листе"

Private Const VALUE_VISIBLE As String = "видимое"
Private Const VALUE_HIDDEN As String = "скрытое"
Private Const VALUE_WB As String = "Книга"

Private Sub btnDelete_Click()
    If listNames.ListCount - 1 < 0 Then Exit Sub
    If MsgBox("Продолжить, удаление имен листов?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim i           As Long
    Dim wb          As Workbook
    Set wb = Workbooks(cmbWB.Value)
    With listNames
        On Error Resume Next
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                wb.Names(.List(i, 1)).Delete
                If Err.Number = 0 Then Call .RemoveItem(i)
            End If
        Next i
        On Error GoTo 0
    End With
End Sub

Private Sub btnShowHide_Click()
    If listNames.ListCount - 1 < 0 Then Exit Sub
    Dim i           As Long
    Dim wb          As Workbook
    Set wb = Workbooks(cmbWB.Value)
    With listNames
        On Error Resume Next
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                With wb.Names(.List(i, 1))
                    .Visible = Not .Visible
                    If .Visible Then
                        listNames.List(i, 3) = VALUE_VISIBLE
                    Else
                        listNames.List(i, 3) = VALUE_HIDDEN
                    End If
                    listNames.List(i, 7) = .Visible
                End With
            End If
        Next i
        On Error GoTo 0
    End With
End Sub

Private Sub btnSortNum_Click()
    Call sortColumnList(listNames, btnSortNum, 0, True)
End Sub

Private Sub btnSortName_Click()
    Call sortColumnList(listNames, btnSortName, 1, False)
End Sub

Private Sub btnSortFormula_Click()
    Call sortColumnList(listNames, btnSortFormula, 2, False)
End Sub

Private Sub btnSortVisible_Click()
    Call sortColumnList(listNames, btnSortVisible, 6, False)
End Sub

Private Sub btnSortParent_Click()
    Call sortColumnList(listNames, btnSortParent, 4, False)
End Sub

Private Sub btnSortError_Click()
    Call sortColumnList(listNames, btnSortError, 5, False)
End Sub

Private Sub btnSortComment_Click()
    Call sortColumnList(listNames, btnSortComment, 6, False)
End Sub

Private Sub cmbFilter_Change()
    With cmbFilter
        If .ListIndex < 0 Then
            If .Value = vbNullString Then
                Call selectedItemListSheets(listNames, vbNullString, 1)
            Else
                Call selectedItemListSheets(listNames, "*" & .Value & "*", 1)
            End If
        Else
            ' Оптимизация: Direct access по имени без необходимости предварительной активации
            Call selectedItemListSheets(listNames, cmbFilter.Value, 1)
        End If
    End With
End Sub

Private Sub cmbWB_Change()
    Call refreshForm
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


    ' Настройка списка фильтров
    With listFilters
        .AddItem FILTER_ALL
        .AddItem FILTER_NONE
        .AddItem FILTER_REVERSE
        .AddItem FILTER_VISIBLE
        .AddItem FILTER_HIDDEN
        .AddItem FILTER_ERROR
        .AddItem FILTER_OUT_ERROR
        .AddItem FILTER_WB
        .AddItem FILTER_SH
    End With

    Dim wb          As Workbook

    For Each wb In Workbooks
        cmbWB.AddItem wb.Name
    Next wb
    If Workbooks.Count > 0 Then cmbWB.Value = ActiveWorkbook.Name


    ' Назначение иконок кнопкам через CommandBars.GetImageMso
    With Application.CommandBars
        Set btnDelete.Picture = .GetImageMso("SketchpadToolDeleteBackground", W, h)
        Set btnShowHide.Picture = .GetImageMso("GroupWindow", W, h)
    End With

    Call refreshForm
End Sub

Private Sub refreshForm()

    Dim nm          As Name
    Dim iCount      As Long
    Dim i           As Long
    Dim sEmptyFormula As String

    Dim wb          As Workbook
    cmbFilter.Clear
    listNames.Clear
    If cmbWB.Value = vbNullString Then Exit Sub
    Set wb = Workbooks(cmbWB.Value)
    iCount = wb.Names.Count
    If iCount = 0 Then Exit Sub

    ReDim arrNames(1 To iCount, 1 To 9) As String
    cmbFilter.Clear
    
    sEmptyFormula = "=" & VBA.Chr$(34) & VBA.Chr$(34)

    For i = 1 To iCount
        Set nm = wb.Names(i)
        With nm
            arrNames(i, 1) = i
            arrNames(i, 2) = .Name

            cmbFilter.AddItem .Name

            arrNames(i, 3) = .RefersToLocal
            If arrNames(i, 3) = sEmptyFormula Then arrNames(i, 3) = vbNullString
            If .Visible Then
                arrNames(i, 4) = VALUE_VISIBLE
            Else
                arrNames(i, 4) = VALUE_HIDDEN
            End If
            If TypeOf .Parent Is Workbook Then
                arrNames(i, 5) = VALUE_WB
                arrNames(i, 9) = True
            Else
                arrNames(i, 5) = .Parent.Name
                arrNames(i, 9) = False
            End If

            arrNames(i, 7) = .Comment
            arrNames(i, 8) = .Visible

            If VBA.InStr(1, .Value, "#REF!") > 0 Then
                arrNames(i, 6) = "#REF!"
            ElseIf .Value = "=#NAME?" Then
                arrNames(i, 6) = "#NAME?"
            Else
                On Error Resume Next
                Select Case .RefersToRange.TEXT
                    Case "#NAME?", "#ИМЯ?"
                        arrNames(i, 6) = "#NAME?"
                End Select

                On Error GoTo 0
            End If
        End With
    Next i

    listNames.List = arrNames
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
            Case FILTER_ALL: Call selectedItemListSheets(listNames, "*", 1)
            Case FILTER_NONE: Call selectedItemListSheets(listNames, vbNullString, 1)
            Case FILTER_REVERSE: Call reversSelected
            Case FILTER_VISIBLE: Call selectedItemListSheets(listNames, VALUE_VISIBLE, 3)
            Case FILTER_HIDDEN: Call selectedItemListSheets(listNames, VALUE_HIDDEN, 3)
            Case FILTER_ERROR: Call selectedItemListSheets(listNames, "[#]*", 5)
            Case FILTER_OUT_ERROR: Call selectedItemListSheets(listNames, vbNullString, 5)
            Case FILTER_WB: Call selectedItemListSheets(listNames, True, 8)
            Case FILTER_SH: Call selectedItemListSheets(listNames, False, 8)

        End Select
    End With
End Sub

Private Sub reversSelected()
    Dim i           As Long
    With listNames
        For i = 0 To .ListCount - 1
            .Selected(i) = Not .Selected(i)
        Next i
    End With
End Sub
