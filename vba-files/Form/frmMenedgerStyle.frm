VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerStyle 
   Caption         =   "Менеджер стилей:"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19455
   OleObjectBlob   =   "frmMenedgerStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerStyle"
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
Private Const FILTER_SISTEM As String = "встроенные"
Private Const FILTER_CUSTOM As String = "пользовательские"

Private Const VALUE_SISTEM As String = "встроенный"
Private Const VALUE_CUSTOM As String = "пользовательский"


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
                wb.Styles(.List(i, 1)).Delete
                If Err.Number = 0 Then Call .RemoveItem(i)
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

Private Sub btnSortTypeStyle_Click()
    Call sortColumnList(listNames, btnSortTypeStyle, 2, False)
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
        .AddItem FILTER_SISTEM
        .AddItem FILTER_CUSTOM
    End With

    Dim wb          As Workbook

    For Each wb In Workbooks
        cmbWB.AddItem wb.Name
    Next wb
    If Workbooks.Count > 0 Then cmbWB.Value = ActiveWorkbook.Name


    ' Назначение иконок кнопкам через CommandBars.GetImageMso
    With Application.CommandBars
        Set btnDelete.Picture = .GetImageMso("SketchpadToolDeleteBackground", W, h)
    End With

    Call refreshForm
End Sub

Private Sub refreshForm()

    Dim st          As Style
    Dim iCount      As Long
    Dim i           As Long

    Dim wb          As Workbook
    cmbFilter.Clear
    listNames.Clear
    If cmbWB.Value = vbNullString Then Exit Sub
    Set wb = Workbooks(cmbWB.Value)
    iCount = wb.Styles.Count
    If iCount = 0 Then Exit Sub

    ReDim arrNames(1 To iCount, 1 To 3) As String

    For i = 1 To iCount
        Set st = wb.Styles(i)
        With st
            arrNames(i, 1) = i
            arrNames(i, 2) = .Name

            If .BuiltIn Then
                arrNames(i, 3) = VALUE_SISTEM
            Else
                arrNames(i, 3) = VALUE_CUSTOM
            End If

            cmbFilter.AddItem .Name
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
            Case FILTER_SISTEM: Call selectedItemListSheets(listNames, VALUE_SISTEM, 2)
            Case FILTER_CUSTOM: Call selectedItemListSheets(listNames, VALUE_CUSTOM, 2)
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
