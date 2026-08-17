VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaQRCodeAsFile 
   Caption         =   "Создание файлов QR код:"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   OleObjectBlob   =   "frmFormulaQRCodeAsFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaQRCodeAsFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim errMsg      As String
    Dim sPath       As String
    If txtPath.Value = vbNullString Then
        errMsg = "Не задана папка для файлов"
    End If
    sPath = txtPath.Value & Application.PathSeparator

    If txtRange.Value = vbNullString Then
        If errMsg <> vbNullString Then errMsg = errMsg & vbNewLine
        errMsg = "Не выбран диапазон с данными"
    End If

    If errMsg <> vbNullString Then
        Call MsgBox(errMsg, vbCritical)
        Exit Sub
    End If

    Dim i           As Long
    Dim iCount      As Long
    Dim bName       As Boolean

    Dim arrQR       As Variant
    arrQR = ActiveSheet.Range(txtRange.Value).Value2
    arrQR = addArray(arrQR)
    iCount = UBound(arrQR, 1)
    Dim arrName     As Variant
    If txtNameFile.Value <> vbNullString Then
        arrName = ActiveSheet.Range(txtNameFile.Value).Value2
        arrName = addArray(arrName)
        bName = True
        If iCount <> UBound(arrName, 1) Then
            If errMsg <> vbNullString Then errMsg = errMsg & vbNewLine
            errMsg = "Выбраны разные размеры диапазанов"
        End If
    End If

    If errMsg <> vbNullString Then
        Call MsgBox(errMsg, vbCritical, "Создание QR файлов:")
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

    Dim sFullName   As String

    For i = 1 To iCount
        If arrQR(i, 1) <> vbNullString Then
            If bName And arrName(i, 1) <> vbNullString Then
                sFullName = sPath & CleanFileName(arrName(i, 1))
            Else
                sFullName = sPath & "QR_" & i
            End If
            Call SavePicture(QRCodegenBarcode(arrQR(i, 1), lbFrontColor.BackColor, 120, bType, errQR, VERSION_MIN, VERSION_MAX, QRCodegenMask_AUTO, True), sFullName & ".emf")
        End If
        lbProgress.Width = i / iCount * 234
    Next i
    Call Unload(Me)
End Sub

Private Function addArray(ByVal arr As Variant) As Variant
    If Not IsArray(arr) Then
        ReDim arrNew(1 To 1, 1 To 1) As String
        arrNew(1, 1) = arr
        addArray = arrNew
    Else
        addArray = arr
    End If
    Exit Function
errMsg:
    Err.Clear
    ReDim arrNew(1 To 1, 1 To 1) As String
    arrNew(1, 1) = vbNullChar
    addArray = arrNew
End Function

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

Private Sub txtFrontColor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 48 To 57, 65 To 70, 97 To 102
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPath_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtNameFile_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtRange_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub txtPath_DropButtonClick()
    Me.Hide
    txtPath.Value = PathDialogFun(ActiveWorkbook.Path)
    Me.Show
End Sub

Private Sub txtNameFile_DropButtonClick()
    Me.Hide
    txtNameFile.Value = SelectRangeViaDialog(, , False)
    Me.Show
End Sub

Private Sub txtRange_DropButtonClick()
    Me.Hide
    txtRange.Value = SelectRangeViaDialog(, , False)
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
    Call ConfigureDropButton(txtFrontColor)
    Call ConfigureDropButton(txtPath)
    Call ConfigureDropButton(txtRange)
    Call ConfigureDropButton(txtNameFile)
    txtPath.Value = ActiveWorkbook.Path

    lbProgress.Width = 0
End Sub
