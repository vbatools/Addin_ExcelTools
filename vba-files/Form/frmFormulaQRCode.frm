VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaQRCode 
   Caption         =   "Создание формулы QR код:"
   ClientHeight    =   5220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   OleObjectBlob   =   "frmFormulaQRCode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    If txtMain.TEXT = vbNullString Then
        Me.Hide
        Call MsgBox("Не задан текст для QR кода!", vbCritical, "Ошибка:")
        Me.Show
        Exit Sub
    End If
    Dim bType       As Boolean
    If cmbType.Value = "Квадрат" Then bType = True
    Dim errQR       As QRCodegenEcc

    Select Case cmbError.Value
        Case "L": errQR = QRCodegenEcc_LOW
        Case "M": errQR = QRCodegenEcc_MEDIUM
        Case "Q": errQR = QRCodegenEcc_QUARTILE
        Case "H": errQR = QRCodegenEcc_HIGH
        Case Else: errQR = QRCodegenEcc_LOW
    End Select

    If lbAddress.Caption = vbNullString Then
        activeCell.FormulaR1C1 = "=Создать_QR(" & Chr$(34) & txtMain & Chr$(34) & "," & lbFrontColor.BackColor & "," & cmbSize.Value & "," & bType & "," & errQR & ")"
    Else
        activeCell.FormulaR1C1 = "=Создать_QR(" & lbAddress.Caption & "," & lbFrontColor.BackColor & "," & cmbSize.Value & "," & bType & "," & errQR & ")"
    End If
    Me.Hide
End Sub

Private Sub lbFrontColor_Click()
    Dim lSeletedColor As Long
    lSeletedColor = GetColorFromDialog()
    If lSeletedColor = -1 Then Exit Sub
    lbFrontColor.BackColor = lSeletedColor
    txtFrontColor.TEXT = VBA.Right$("000000" & VBA.Hex(lSeletedColor), 6)
End Sub

Private Sub txtFrontColor_DropButtonClick()
    Me.Hide
    Dim sAddress    As String
    sAddress = SelectRangeViaDialog()
    If sAddress <> vbNullString Then
        sAddress = VBA.Split(sAddress, ":")(0)
        lbFrontColor.BackColor = Range(sAddress).Interior.Color
        txtFrontColor.TEXT = VBA.Right$("000000" & VBA.Hex(lbFrontColor.BackColor), 6)
    End If
    Me.Show
End Sub

Private Sub txtCell_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtCell_DropButtonClick()
    Me.Hide
    txtCell.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtFrontColor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 65 To 70, 97 To 102
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtMain_DropButtonClick()
    Me.Hide
    With lbAddress
        .Caption = vbNullString
        .Caption = SelectRangeViaDialog(, False)
        If .Caption <> vbNullString Then
            .Caption = VBA.Split(.Caption, ":")(0)
            txtMain.Value = Range(.Caption).Value
            .Caption = Range(.Caption).Address(True, True, xlR1C1)
        End If
    End With
    Me.Show
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    
    Dim i           As Integer
    txtFrontColor.TEXT = "000000"
    For i = 1 To 10
        If i = 10 Then
            cmbSize.AddItem 1000
        Else
            cmbSize.AddItem 150 + 100 * (i - 1)
        End If
    Next i
    cmbError.AddItem "L"
    cmbError.AddItem "M"
    cmbError.AddItem "Q"
    cmbError.AddItem "H"
    cmbType.AddItem "Квадрат"
    cmbType.AddItem "Плавно"
    cmbType.Value = cmbType.List(1)
    cmbSize.Value = cmbSize.List(1)
    cmbError.Value = cmbError.List(0)
    Call ConfigureDropButton(txtMain)
    Call ConfigureDropButton(txtFrontColor)
    Call ConfigureDropButton(txtCell)
    txtCell.TEXT = activeCell.Address
End Sub
