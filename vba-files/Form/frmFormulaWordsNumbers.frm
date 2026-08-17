VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaWordsNumbers 
   Caption         =   "Создание формулы сумма прописью:"
   ClientHeight    =   5940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmFormulaWordsNumbers.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaWordsNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const SKOBKIVAL1 As String = "(100,01) сто целых 1 сотая"
Private Const SKOBKIVAL2 As String = "100,01 (сто целых 1 сотая)"
Private Const MAXLENGS As Integer = 16    '20-триллионы

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    txtValue.Value = vbNullString
End Sub

Private Sub btnOK_Click()
    Dim txt         As Variant
    Dim sFormat     As String

    txt = txtValue.TEXT
    txt = Replace(txt, " ", vbNullString)
    If chcTysych Then
        If txt Like "*,*" Then
            sFormat = "#,##0.00"
        Else
            sFormat = "#,##0"
        End If
    Else
        sFormat = "0"
    End If

    If txt = vbNullString Then txt = 0
    txt = VBA.CCur(txt)
    If lbRC.Caption <> vbNullString Then txt = lbRC.Caption
    With activeCell
        .FormulaR1C1 = "=СУММАПРОПИСЬЮ(" & txt & "," & _
                VBA.CInt(cmbCase.ListIndex) & "," & VBA.CInt(cmbTypeDate.ListIndex) & "," & _
                VBA.CBool(chcDublVal.Value) & "," & VBA.CBool(chcDrobProp.Value) & "," & _
                VBA.CInt(cmbRegistr.ListIndex) & "," & VBA.CInt(cmbSkobki.ListIndex) & "," & VBA.Chr$(34) & sFormat & VBA.Chr$(34) & ")"

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

Private Sub chcALL_Click()
    chcDublVal = chcALL.Value
    If chcDrobProp.Enabled = True Then chcDrobProp = chcALL.Value
End Sub

Private Sub chcDrobProp_Click()
    Call addFormula
End Sub

Private Sub chcDublVal_Click()
    Call addFormula
End Sub

Private Sub chcTysych_Click()
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
    Select Case cmbTypeDate.ListIndex
        Case 1, 2, 3, 8
            chcDrobProp.Enabled = True
            If chcALL.Value = True Then chcDrobProp.Value = True
        Case Else
            chcDrobProp.Enabled = False
            chcDrobProp.Value = False
    End Select
    Call addFormula
End Sub

Private Sub txtValue_Change()
    Dim txt         As String
    txt = Replace(txtValue.Value, " ", vbNullString)
    Dim MaxLeng     As Byte
    MaxLeng = MAXLENGS

    Select Case txt
        Case ""
            txt = vbNullString
            GoTo Ends
        Case "0", ","
            txt = "0,"
            GoTo Ends
        Case "-"
            txt = "-"
            GoTo Ends
        Case "-0", "-0,"
            txt = "-0,"
            GoTo Ends
    End Select

    Dim txtPatch()  As String
    txtPatch = Split(txt, ",")
    Dim txtFerst    As String
    txtFerst = txtPatch(0)
    Dim txtSecond   As String
    txtSecond = vbNullString
    If UBound(txtPatch) = 1 Then
        txtSecond = "," & Left(txtPatch(1), 4)
        MaxLeng = MaxLeng + 5
    End If
    Dim Minus       As String
    Minus = Left(txtFerst, 1)
    If Minus = "-" Then
        MaxLeng = MaxLeng + 1
        txtFerst = Right(txtFerst, Len(txtFerst) - 1)
    Else
        Minus = vbNullString
    End If
    If txtFerst = "0" Then
        GoTo Ends1
    End If
    Select Case Len(txtFerst)
        Case 1, 2, 3
            txtFerst = format(txtFerst, " ###")
        Case 4, 5, 6
            txtFerst = format(txtFerst, " ### ###")
        Case 7, 8, 9
            txtFerst = format(txtFerst, " ### ### ###")
        Case 10, 11, 12
            txtFerst = format(txtFerst, " ### ### ### ###")
        Case 13, 14, 15
            txtFerst = format(txtFerst, " ### ### ### ### ###")
    End Select
Ends1:
    txt = Minus & txtFerst & txtSecond
Ends:
    txtValue.Value = txt
    txtValue.MaxLength = MaxLeng

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
                .Value = sTxt
            Else
                .Value = vbNullString
            End If
        End If
    End With
    Me.Show
End Sub

Private Sub txtValue_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error Resume Next
    Select Case KeyAscii
            'цифры 1 - 9 и <Backspace> (эти символы всегда разрешены)
        Case 48 To 57, 8
            'обработка десятичного разделителя (44 - код запятой, 46 - код точки)
        Case 44, 46
            'если в поле  введена точка, то заменим ее на запятую
            KeyAscii = 44
            If txtValue.MaxLength = MAXLENGS Or txtValue.MaxLength = MAXLENGS + 1 Then txtValue.MaxLength = MAXLENGS + 5
            '------------если нужен дефис (минус перед числом), раскоментируйте код---------------
        Case 45
            If InStr(1, txtValue.TEXT, "-") Then KeyAscii = 0    'второй минус нельзя
            If txtValue.SelStart Then KeyAscii = 0    'минус допустим только перед числом
            '-------------------------------------------------------------------------------------
        Case Else
            'остальные символы запрещены
            KeyAscii = 0
    End Select
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
        .AddItem "Дательный"
        .AddItem "Винительный"
        .AddItem "Творительный"
        .AddItem "Предложный"
        .ListIndex = 0
    End With

    With cmbTypeDate
        .AddItem "Ничего"
        .AddItem "Рубли"
        .AddItem "Доллары"
        .AddItem "Евро"
        .AddItem "Календарные дни"
        .AddItem "Рабочие дни"
        .AddItem "Дни"
        .AddItem "Штуки"
        .AddItem "Целое + дробное"
        .AddItem "Рубли + 00 копейки"
        '.AddItem "Проценты"
        .ListIndex = 0
    End With
    With cmbSkobki
        .AddItem "НЕТ"
        .AddItem SKOBKIVAL1
        .AddItem SKOBKIVAL2
        .ListIndex = 0
    End With


    chcDrobProp.Enabled = False
    Call ConfigureDropButton(txtValue)
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

Private Sub addFormula()

    If cmbCase.ListIndex < 0 Or cmbTypeDate.ListIndex < 0 Or cmbRegistr.ListIndex < 0 Or cmbSkobki.ListIndex < 0 Then Exit Sub

    Dim txt         As String
    Dim sFormat     As String
    txt = txtValue.TEXT
    txt = Replace(txt, " ", vbNullString)
    If chcTysych Then
        If txt Like "*,*" Then
            sFormat = "#,##0.00"
        Else
            sFormat = "#,##0"
        End If
    Else
        sFormat = "0"
    End If
    If txt = vbNullString Then txt = 0
    txtPropis.TEXT = СУММАПРОПИСЬЮ(VBA.CCur(txt), cmbCase.ListIndex, _
            cmbTypeDate.ListIndex, chcDublVal.Value, chcDrobProp.Value, cmbRegistr.ListIndex, cmbSkobki.ListIndex, sFormat)
End Sub

