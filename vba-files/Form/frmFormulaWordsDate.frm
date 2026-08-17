VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaWordsDate 
   Caption         =   "Создание формулы дата прописью:"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmFormulaWordsDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaWordsDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const SKOBKI_VAL1 As String = "(01.01.2001 первое января две тысячи первого года)"
Private Const SKOBKI_VAL2 As String = "01.01.2001 (первое января две тысячи первого года)"
Private Const SKOBKI_VAL3 As String = "(01.01.2001) первое января две тысячи первого года"
Private Const SKOBKI_VAL4 As String = "(первое января две тысячи первого года)"

Dim clDate          As clsCalendarDate

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim dtDate      As Date
    Dim iSkobki     As Integer
    Dim bFlag       As Boolean

    iSkobki = cmbSkobki.ListIndex - 1
    bFlag = VBA.CBool(chcDublVal.Value)
    If iSkobki < 0 Then bFlag = False

    dtDate = VBA.CDate(txtValue.Value)
    With activeCell
        .FormulaR1C1 = "=ДАТАПРОПИСЬЮ(" & VBA.Chr$(34) & dtDate & VBA.Chr$(34) & "," & cmbCase.ListIndex & "," & _
                cmbTypeDate.ListIndex & "," & iSkobki & "," & bFlag & "," & _
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

Private Sub chcDublVal_Click()
    If chcDublVal Then
        With cmbSkobki
            .Clear
            .AddItem SKOBKI_VAL1
            .AddItem SKOBKI_VAL2
            .AddItem SKOBKI_VAL3
            .Value = SKOBKI_VAL1
        End With
    Else
        With cmbSkobki
            .Clear
            .AddItem SKOBKI_VAL4
            .Value = SKOBKI_VAL4

        End With
    End If
    Call addFormula
End Sub

Private Sub chcSkobki_Click()
    'cmbSkobki.Enabled = Not chcSkobki.Value
    Call addFormula
End Sub

Private Sub cmbCase_Change()
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

Private Sub TogBtnFat_Change()
    txtPropis.Font.Bold = TogBtnFat.Value
End Sub

Private Sub TogBtnKursiv_Change()
    txtPropis.Font.Italic = TogBtnKursiv.Value
End Sub

Private Sub TogBtnCherta_Change()
    txtPropis.Font.Underline = TogBtnCherta.Value
End Sub

Private Sub txtValue_Change()
    Call addFormula
End Sub

Private Sub txtValue_DropButtonClick()
    Me.Hide
    With txtValue
        Dim sTxt    As String
        sTxt = SelectRangeViaDialog()
        If sTxt <> vbNullString Then
            lbRC.Caption = Range(VBA.Split(sTxt, ":")(0)).Address(ReferenceStyle:=xlR1C1)
            sTxt = Range(VBA.Split(sTxt, ":")(0)).Value2
            If IsNumeric(sTxt) Then
                On Error Resume Next
                .Value = VBA.format(sTxt, "dd.mm.yyyy")
                On Error GoTo 0
            Else
                .Value = vbNullString
            End If
        End If
    End With
    Me.Show
End Sub

Private Sub txtValue_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    KeyCode = 0
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    With cmbRegistr
        .AddItem "все строчные"
        .AddItem "ВСЕ ПРОПИСНЫЕ"
        .AddItem "Как в предложениях"
        .AddItem "Начинать С Прописных"
        .ListIndex = 0
    End With

    With cmbCase
        .AddItem "Именительный"
        .AddItem "Родительный"
        .ListIndex = 0
    End With
    With cmbTypeDate
        .AddItem "Все прописью"
        .AddItem "Месяц прописью"
        .ListIndex = 0
    End With
    With cmbSkobki
        .AddItem SKOBKI_VAL1
        .AddItem SKOBKI_VAL2
        .AddItem SKOBKI_VAL3
        .ListIndex = 0
    End With
    txtValue.Value = VBA.format(VBA.Date, "dd.mm.yyyy")
    Set clDate = New clsCalendarDate
    Call clDate.AddDatePicker(txtValue, VBA.Date)

    Call ConfigureDropButton(txtValue)

    Call addFormula
End Sub

Private Sub addFormula()

    If cmbCase.ListIndex < 0 Or cmbTypeDate.ListIndex < 0 Or cmbRegistr.ListIndex < 0 _
            Or cmbSkobki.ListIndex < 0 Or txtValue.TEXT = vbNullString Then Exit Sub

    Dim dtDate      As Date
    dtDate = VBA.CDate(txtValue.TEXT)

    txtPropis.TEXT = ДАТАПРОПИСЬЮ(dtDate, _
            cmbCase.ListIndex, _
            cmbTypeDate.ListIndex, _
            cmbSkobki.ListIndex, _
            chcSkobki.Value, _
            cmbRegistr.ListIndex, _
            chcDublVal.Value)
End Sub

