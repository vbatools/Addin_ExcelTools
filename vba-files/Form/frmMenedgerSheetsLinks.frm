VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsLinks 
   Caption         =   "Создание оглавления:"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmMenedgerSheetsLinks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmSheetsMenedgerLinks- add description!
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   17-06-2026 15:56:21
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Const SHEET_NAME_OG As String = "ОГЛАВЛЕНИЕ"
    If HaveSheetInFile(ActiveWorkbook, SHEET_NAME_OG) Then
        Call DisableApplicationSettings
        ActiveWorkbook.Sheets(SHEET_NAME_OG).Delete
        Call RestoreApplicationSettings
    End If

    Application.ReferenceStyle = xlA1

    Dim arr         As Variant
    arr = frmMenedgerSheets.listSheets.List

    Dim i           As Long
    Dim k           As Long
    Dim iCount      As Long
    iCount = UBound(arr, 1)
    ReDim arrRes(1 To iCount + 1, 1 To 8)
    For i = 0 To iCount
        If arr(i, 1) <> SHEET_NAME_OG Then
            k = k + 1
            arrRes(k, 1) = arr(i, 0)
            arrRes(k, 2) = arr(i, 1)
            arrRes(k, 3) = arr(i, 2)
            arrRes(k, 4) = arr(i, 3)
            arrRes(k, 5) = arr(i, 4)
            arrRes(k, 6) = arr(i, 5)
            arrRes(k, 7) = arr(i, 6)
            arrRes(k, 8) = arr(i, 8)
        End If
    Next i
    If k = 0 Then Exit Sub
    Dim Sh          As Worksheet
    Set Sh = ActiveWorkbook.Worksheets.Add(before:=ActiveWorkbook.Sheets(1))
    With Sh
        .Name = SHEET_NAME_OG
        .Cells(2, 1).Resize(iCount + 1, 8).Value2 = arrRes
        With .Cells(1, 1).Resize(1, 8)
            .Value2 = Array("№", "имя листа", "тип листа", "видим/скрыт", "защита", "диапазон", "кол-во ячеек", "ранжирование")
            .AutoFilter
        End With
        .Columns("A:I").EntireColumn.AutoFit
    End With

    If chbAddbBackLinks.Value Then
        Dim rng     As Range
        Dim sValue  As String
        Dim iRowCount As Long
        Dim iColCount As Integer
        Set rng = Range(txtInputRng.Value)
        iRowCount = rng.Row
        iColCount = rng.Column
    End If

    Dim lColor      As Long

    For i = 1 To k
        If arrRes(i, 3) = "лист" Then
            Sh.Hyperlinks.Add Anchor:=Sh.Cells(i + 1, 2), Address:="", SubAddress:= _
                    "'" & arrRes(i, 2) & "'!A1", TextToDisplay:=arrRes(i, 2)
            If chbAddbBackLinks.Value Then
                With Worksheets(arrRes(i, 2))
                    sValue = .Cells(i + 1, 2).Value
                    If sValue = vbNullString Then sValue = "<<" & SHEET_NAME_OG
                    .Hyperlinks.Add Anchor:=.Cells(iRowCount, iColCount), Address:="", SubAddress:= _
                            "'" & SHEET_NAME_OG & "'!A1", TextToDisplay:=sValue
                End With
            End If
        End If
        lColor = Sheets(arrRes(i, 2)).Tab.Color
        If lColor > 0 Then Sh.Cells(i + 1, 2).Interior.Color = lColor
    Next i




    Unload Me
End Sub

Private Sub chbAddbBackLinks_Click()
    txtInputRng.Enabled = chbAddbBackLinks.Value
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    If TypeName(Selection) = "Range" Then txtInputRng.Value = Selection.Address
    Call ConfigureDropButton(txtInputRng)
End Sub



