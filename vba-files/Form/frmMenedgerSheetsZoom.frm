VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsZoom 
   Caption         =   "Установить масштаб:"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmMenedgerSheetsZoom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit



Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Call DisableApplicationSettings
    Dim i           As Long
    Dim Sh          As Object
    Dim shActive    As Object
    Dim shVisible   As XlSheetVisibility
    Set shActive = ActiveSheet
    With listSheets
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                Set Sh = ActiveWorkbook.Sheets(.List(i, 1))
                If TypeName(Sh) <> "DialogSheet" Then
                    shVisible = Sh.Visible
                    Sh.Visible = XlSheetVisibility.xlSheetVisible
                    Sh.Activate
                    ActiveWindow.Zoom = VBA.Val(txtZoom.Value)
                    .List(i, 2) = VBA.Val(txtZoom.Value)
                    Sh.Visible = shVisible
                End If
            End If
        Next i
    End With
    shActive.Activate
    Call RestoreApplicationSettings
End Sub

Private Sub btnSortName_Click()
    Call SortColumnList(listSheets, btnSortName, 1, False, True)
End Sub

Private Sub btnSortNum_Click()
    Call SortColumnList(listSheets, btnSortNum, 0, True)
End Sub

Private Sub btnSortZoom_Click()
    Call SortColumnList(listSheets, btnSortZoom, 2, True)
End Sub

Private Sub listSheets_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    With listSheets
        If .ListIndex < 0 Then Exit Sub
        ' Активация выбранного листа (имя хранится во втором столбце с индексом 1)
        ActiveWorkbook.Sheets(.List(.ListIndex, 1)).Activate
    End With
End Sub

Private Sub txtZoom_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 9, 13, 27
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    txtZoom.Value = ActiveWindow.Zoom
    Call refreshListSheets
End Sub

Private Sub ScrolZoom_Change()
    txtZoom.Value = ScrolZoom.Value & "%"
End Sub

Private Sub txtZoom_AfterUpdate()
    Select Case VBA.Val(txtZoom.Value)
        Case Is < 10
            txtZoom.Value = 10 & "%"
        Case Is > 400
            txtZoom.Value = 400 & "%"
        Case Else
            ScrolZoom.Value = VBA.Val(txtZoom.Value)
    End Select
End Sub

Private Sub refreshListSheets()
    Call DisableApplicationSettings
    Dim i           As Long
    Dim iCount      As Long
    Dim Sh          As Object
    Dim shVisible   As XlSheetVisibility
    iCount = ActiveWorkbook.Sheets.Count
    ReDim arr(1 To iCount, 1 To 3) As String
    Set Sh = ActiveSheet
    For i = 1 To iCount
        With ActiveWorkbook.Sheets(i)
            arr(i, 1) = i
            arr(i, 2) = .Name
            shVisible = .Visible
            .Visible = XlSheetVisibility.xlSheetVisible
            .Activate
            arr(i, 3) = ActiveWindow.Zoom
            .Visible = shVisible
        End With
    Next i
    listSheets.List = arr
    Call RestoreApplicationSettings
End Sub
