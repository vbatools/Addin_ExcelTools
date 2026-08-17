VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOtherToolsBackUpFile 
   Caption         =   "Резервная копия книги (настройки):"
   ClientHeight    =   3900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12060
   OleObjectBlob   =   "frmOtherToolsBackUpFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOtherToolsBackUpFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDeleteSettings_Click()
    If delFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_PATH) And _
            delFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION) Then
        txtPath.Value = vbNullString
        txtVersionFix.Value = vbNullString
        txtVersionMain.Value = vbNullString
        txtVersionSub.Value = vbNullString
        chAddDateFileName.Value = False
        Call MsgBox("Настройки резервного копирования удалены", vbInformation)
    End If
    Call delFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_DATE)
    Call delFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_DATE_ADD)
End Sub

Private Sub btnAddSettings_Click()
    If txtPath.Value = vbNullString Then
        Call MsgBox("Не выбрана папка для резервный копий", vbCritical)
        Exit Sub
    End If

    Dim sPath       As String
    sPath = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_PATH)
    If sPath = vbNullString Then
        Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_PATH, txtPath.Value)
        Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_DATE, lbDateVersion.Caption)
        Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_DATE_ADD, chAddDateFileName.Value)
        Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION, addVersion())
        Call MsgBox("Настроки для резервного копирования созданы", vbInformation)
    Else
        If sPath = txtPath.Value Then
            Call MsgBox("Настроки для резервного копирования были созданы ранее", vbInformation)
        Else
            If MsgBox("ранее выбраная папка для копий не совпадает с текущей, изменить папку?", vbYesNo + vbQuestion, "") = vbYes Then
                Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_PATH, txtPath.Value)
                Call MsgBox("Настроки для резервного копирования изменены", vbInformation)
            Else
                txtPath.Value = sPath
            End If
        End If
    End If
End Sub

Private Sub btnOpentPath_Click()
    Dim sPath       As String
    sPath = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_PATH)
    If sPath = vbNullString Then Exit Sub
    Call OpenPath(sPath)
End Sub

Private Sub txtPath_DropButtonClick()
    Me.Hide
    Dim sPath As String
    sPath = txtPath.Value
    If sPath = vbNullString Then sPath = ActiveWorkbook.Path
    txtPath.Value = PathDialogFun(sPath)
    Me.Show
End Sub

Private Sub txtPath_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub



Private Sub txtVersionFix_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtVersionMain_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)

End Sub

Private Sub txtVersionSub_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)

    Call ConfigureDropButton(txtPath)

    Dim sPath       As String
    sPath = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_PATH)
    txtPath.Value = sPath
    Dim sVersion    As String
    sVersion = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION)
    If sVersion <> vbNullString Then
        Dim arr     As Variant
        arr = VBA.Split(sVersion, ".")
        txtVersionMain.Value = arr(0)
        If UBound(arr, 1) > 0 Then
            txtVersionSub.Value = arr(1)
            If UBound(arr, 1) > 1 Then txtVersionFix.Value = arr(2)
        End If
    End If
    lbDateVersion.Caption = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_DATE)
    If lbDateVersion.Caption = vbNullString Then lbDateVersion.Caption = "нет файлов"
    chAddDateFileName.Value = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_DATE_ADD) = "True"
End Sub

Private Function addVersion() As String
    If txtVersionMain.Value = vbNullString Then txtVersionMain.Value = 0
    addVersion = txtVersionMain.Value
    If txtVersionSub.Value <> vbNullString Then
        addVersion = addVersion & "." & txtVersionSub.Value
        If txtVersionFix.Value <> vbNullString Then addVersion = addVersion & "." & txtVersionFix.Value
    End If
End Function

