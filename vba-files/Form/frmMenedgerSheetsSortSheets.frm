VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsSortSheets 
   Caption         =   "Сортировка листов:"
   ClientHeight    =   2235
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7905
   OleObjectBlob   =   "frmMenedgerSheetsSortSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsSortSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim arr         As Variant
    arr = frmMenedgerSheets.listSheets.List

    Select Case True
        Case optAsc.Value
            arr = SortArray(arr, 1, True, False, True)
        Case optDesc.Value
            arr = SortArray(arr, 1, False, False, True)
        Case optColorAsc.Value
            arr = SortArray(arr, 9, True, True, False)
        Case optColorDesc.Value
            arr = SortArray(arr, 9, False, True, False)
    End Select

    Dim i           As Long
    Dim shVisible   As XlSheetVisibility

    For i = UBound(arr, 1) To 0 Step -1
        With ActiveWorkbook.Sheets(arr(i, 1))
            shVisible = .Visible
            If shVisible <> xlSheetVisible Then .Visible = xlSheetVisible
            .Move After:=ActiveWorkbook.Sheets(i + 1)
            .Visible = shVisible
        End With
    Next i


End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    Const W         As Byte = 16
    Const h         As Byte = 16
    With Application.CommandBars
        Image1.Picture = .GetImageMso("JotFindSortAscending", W, h)
        Image2.Picture = .GetImageMso("JotFindSortDescending", W, h)
        Image3.Picture = .GetImageMso("GroupChartType", W, h)
        Image4.Picture = .GetImageMso("GroupBlogProperties", W, h)
    End With
End Sub
