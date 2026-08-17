VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerWBooks 
   Caption         =   "Менеджер книг:"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19455
   OleObjectBlob   =   "frmMenedgerWBooks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerWBooks"
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
Private Const FILTER_SAVE As String = "сохраненые"
Private Const FILTER_NOT_SAVE As String = "не сохраненые"
Private Const FILTER_PROTECT As String = "защищенные"
Private Const FILTER_NOT_PROTECT As String = "не защищенные"


Private Const VALUE_SAVED As String = "сохранена"
Private Const VALUE_NOT_SAVED As String = "не сохранена"
Private Const VALUE_PROTECT As String = "защита книги"

Private Sub btnClose_Click()
    If listNames.ListCount - 1 < 0 Then Exit Sub
    Dim i           As Long
    With listNames
        On Error Resume Next
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                Workbooks(.List(i, 1)).Close False
                Call .RemoveItem(i)
            End If
        Next i
        On Error GoTo 0
    End With
    If listNames.ListCount - 1 < 0 Then Unload Me
End Sub

Private Sub btnSaveWB_Click()
    If listNames.ListCount - 1 < 0 Then Exit Sub
    Dim i           As Long
    With listNames
        On Error Resume Next
        For i = .ListCount - 1 To 0 Step -1
            If .Selected(i) Then
                With Workbooks(.List(i, 1))
                    If .Path <> vbNullString Then
                        .Save
                        .List(i, 3) = VALUE_SAVED
                    End If
                End With
            End If
        Next i
        On Error GoTo 0
    End With
End Sub

Private Sub btnSortNum_Click()
    Call SortColumnList(listNames, btnSortNum, 0, True)
End Sub

Private Sub btnSortName_Click()
    Call SortColumnList(listNames, btnSortName, 1, False)
End Sub

Private Sub btnSortFormula_Click()
    Call SortColumnList(listNames, btnSortFormula, 2, False)
End Sub

Private Sub btnSortVisible_Click()
    Call SortColumnList(listNames, btnSortVisible, 6, False)
End Sub

Private Sub btnSortParent_Click()
    Call SortColumnList(listNames, btnSortParent, 4, False)
End Sub

Private Sub cmbFilter_Change()
    With cmbFilter
        If .ListIndex < 0 Then
            If .Value = vbNullString Then
                Call SelectedItemListSheets(listNames, vbNullString, 1)
            Else
                Call SelectedItemListSheets(listNames, "*" & .Value & "*", 1)
            End If
        Else
            ' Оптимизация: Direct access по имени без необходимости предварительной активации
            Call SelectedItemListSheets(listNames, cmbFilter.Value, 1)
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
        .AddItem FILTER_SAVE
        .AddItem FILTER_NOT_SAVE
        .AddItem FILTER_PROTECT
        .AddItem FILTER_NOT_PROTECT
    End With


    ' Назначение иконок кнопкам через CommandBars.GetImageMso
    With Application.CommandBars
        Set btnSaveWB.Picture = .GetImageMso("SaveAll", W, h)
        Set btnClose.Picture = .GetImageMso("CancelEditing", W, h)
    End With

    Call refreshForm
End Sub

Private Sub refreshForm()

    Dim i           As Long
    Dim iCount      As Long
    iCount = Workbooks.Count
    cmbFilter.Clear
    If iCount = 0 Then Exit Sub
    
    ReDim arrNames(1 To iCount, 1 To 5) As String
    For i = 1 To iCount
        With Workbooks(i)
            arrNames(i, 1) = i
            arrNames(i, 2) = .Name
            cmbFilter.AddItem .Name
            arrNames(i, 3) = .Sheets.Count
            If .Saved And .Path <> vbNullString Then
                arrNames(i, 4) = VALUE_SAVED
            Else
                arrNames(i, 4) = VALUE_NOT_SAVED
            End If
            If .ProtectStructure Then arrNames(i, 5) = VALUE_PROTECT
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
            Case FILTER_ALL: Call SelectedItemListSheets(listNames, "*", 1)
            Case FILTER_NONE: Call SelectedItemListSheets(listNames, vbNullString, 1)
            Case FILTER_REVERSE: Call reversSelected
            Case FILTER_SAVE: Call SelectedItemListSheets(listNames, VALUE_SAVED, 3)
            Case FILTER_NOT_SAVE: Call SelectedItemListSheets(listNames, VALUE_NOT_SAVED, 3)
            Case FILTER_PROTECT: Call SelectedItemListSheets(listNames, VALUE_PROTECT, 4)
            Case FILTER_NOT_PROTECT: Call SelectedItemListSheets(listNames, vbNullString, 4)
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
