VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaWordsTime 
   Caption         =   "Создание формулы время прописью:"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmFormulaWordsTime.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaWordsTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SKOBKI_VAL1 As String = "(01:01:01 один час одна минута одна секунда)"
Private Const SKOBKI_VAL2 As String = "01:01:01 (один час одна минута одна секунда)"
Private Const SKOBKI_VAL3 As String = "(01:01:01) один час одна минута одна секунда"
Private Const SKOBKI_VAL4 As String = "(один час одна минута одна секунда)"

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim MayTime     As Date
    MayTime = VBA.TimeSerial(txtHour.TEXT, txtMinute.TEXT, txtSecond.TEXT)

    With activeCell
        .FormulaR1C1 = "=ВРЕМЯПРОПИСЬЮ(" & VBA.Chr$(34) & MayTime & VBA.Chr$(34) & "," & cmbTypeDate.ListIndex + 1 & "," & _
                cmbSkobki.ListIndex & "," & VBA.CBool(chcSkobki.Value) & "," & _
                cmbRegistr.ListIndex & "," & VBA.CBool(chcDublVal.Value) & ")"

        .Font.Bold = TogBtnFat.Value
        .Font.Italic = TogBtnKursiv.Value
        If TogBtnCherta.Value Then
            .Font.Underline = xlUnderlineStyleSingle
        Else
            .Font.Underline = xlUnderlineStyleNone
        End If
    End With
    Me.Hide
End Sub

Private Sub btnSetTimeNow_Click()
    txtHour.TEXT = Hour(Time)
    spHour.Value = VBA.Val(txtHour.TEXT)
    txtMinute.TEXT = Minute(Time)
    spMinute.Value = VBA.Val(txtMinute.TEXT)
    txtSecond.TEXT = Second(Time)
    spSecond.Value = VBA.Val(txtSecond.TEXT)
    Call addFormula
End Sub

Private Sub chcDublVal_Click()
    Call addFormula
End Sub

Private Sub chcSkobki_Click()
    Call addFormula
End Sub

Private Sub cmbRegistr_Change()
    Call addFormula
End Sub

Private Sub cmbSkobki_Change()
    Call addFormula
End Sub

Private Sub cmbTypeDate_Change()
    Call addFormula
End Sub

Private Sub spHour_Change()
    With spHour
        Select Case .Value
            Case 24: .Value = 0
            Case -1: .Value = 23
        End Select
        txtHour.TEXT = .Value
    End With
    Call addFormula
End Sub

Private Sub spMinute_Change()
    With spMinute
        Select Case .Value
            Case 60: .Value = 0
            Case -1: .Value = 59
        End Select
        txtMinute.TEXT = .Value
    End With
    Call addFormula
End Sub

Private Sub spSecond_Change()
    With spSecond
        Select Case .Value
            Case 60: .Value = 0
            Case -1: .Value = 59
        End Select
        txtSecond.TEXT = .Value
    End With
    Call addFormula
End Sub

Private Sub TogBtnFat_Change()
    txtPropis.Font.Bold = TogBtnFat.Value
End Sub

Private Sub TogBtnKursiv_Change()
    txtPropis.Font.Italic = TogBtnKursiv.Value
End Sub

Private Sub TogBtnCherta_Change()
    txtPropis.Font.Underline = TogBtnCherta.Value
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    With cmbTypeDate
        .AddItem "Часы, минуты, секунды"
        .AddItem "Часы, минуты"
        .AddItem "Часы"
        .ListIndex = 0
    End With
    With cmbRegistr
        .AddItem "все строчные"
        .AddItem "ВСЕ ПРОПИСНЫЕ"
        .AddItem "Как в предложениях"
        .AddItem "Начинать С Прописных"
        .ListIndex = 0
    End With
    With cmbSkobki
        .AddItem SKOBKI_VAL1
        .AddItem SKOBKI_VAL2
        .AddItem SKOBKI_VAL3
        .ListIndex = 0
    End With

    Dim MayHour     As Byte
    Dim MayMinut    As Byte
    MayHour = VBA.Hour(Time)
    txtHour.TEXT = MayHour
    spHour.Value = MayHour
    MayMinut = VBA.Minute(Time)
    txtMinute.TEXT = MayMinut
    spMinute.Value = MayMinut
    MayHour = VBA.Second(Time)
    txtSecond.TEXT = MayHour
    spSecond.Value = MayHour

    Call addFormula
End Sub

Private Sub addFormula()
    If txtHour.TEXT = vbNullString Or txtMinute.TEXT = vbNullString Or txtSecond.TEXT = vbNullString Then Exit Sub
    Dim MayTime     As Date
    MayTime = format(TimeSerial(txtHour.TEXT, txtMinute.TEXT, txtSecond.TEXT), "h:m:s")
    txtPropis.TEXT = ВРЕМЯПРОПИСЬЮ(MayTime, _
            cmbTypeDate.ListIndex + 1, _
            cmbSkobki.ListIndex, _
            chcSkobki.Value, _
            cmbRegistr.ListIndex, _
            chcDublVal.Value)
End Sub


