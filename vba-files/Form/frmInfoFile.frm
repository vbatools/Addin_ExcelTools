VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmInfoFile 
   Caption         =   "Свойства файлов:"
   ClientHeight    =   7095
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11325
   OleObjectBlob   =   "frmInfoFile.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmInfoFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmInfoFile - Manage file properties
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   20-07-2020 15:34
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub cmbMain_Change()
    If cmbMain.Value = vbNullString Then Exit Sub

    On Error Resume Next
    Dim arr         As Variant
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.Value)
    arr = getFilePropertiesList(wb)
    If Not IsEmpty(arr) Then
        ListProp.List = arr
    Else
        ListProp.Clear
    End If

    arr = getFilePropertiesCustomList(wb)
    If Not IsEmpty(arr) Then
        ListCustomProp.List = arr
    Else
        ListCustomProp.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub LbDelAllProper_Click()
    If cmbMain.Value = vbNullString Then
        Call MsgBox("Нет выбранных книг!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.Value)

    If MsgBox("Удалить все свойства?", vbYesNo + vbQuestion, "Удаление свойств:") = vbYes Then
        Dim iCount  As Byte
        iCount = delFilePropertiesAll(wb)
        Call cmbMain_Change
        Call MsgBox("Свойств удалено:" & iCount, vbInformation, "Удаление свойств:")
    End If
End Sub
Private Sub LbEdit_Click()
    Call editProperty
End Sub

Private Sub ListProp_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Call editProperty
End Sub
Private Sub editProperty()
    If cmbMain.Value = vbNullString Then
        Call MsgBox("Нет выбранных книг!", vbCritical)
        Exit Sub
    End If
    On Error Resume Next
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.Value)

    Dim i           As Long
    With ListProp
        i = .ListIndex
        If i < 0 Then Exit Sub
        Dim sNameProp As String
        Dim sValueProp As String
        sNameProp = .List(i, 1)
        sValueProp = VBA.Trim$(.List(i, 2))
    End With
    Dim sNewValueProp As String

    sNewValueProp = InputBox("Изменить свойство [" & sNameProp & " ] ?", "Изменение свойства:", sValueProp)
    If sNewValueProp <> sValueProp Then
        If addFileProperty(wb, sNameProp, sNewValueProp) Then Call cmbMain_Change
    End If
End Sub

Private Sub lbAddCustProp_Click()
    Call AddCustProp(vbNullString, vbNullString)
End Sub

Private Sub lbEditCustProp_Click()

    Dim i           As Long
    With ListCustomProp
        i = .ListIndex
        If i < 0 Then Exit Sub
        Dim sNameProp As String
        Dim sValueProp As String
        sNameProp = .List(i, 1)
        sValueProp = VBA.Trim$(.List(i, 2))
    End With

    Call AddCustProp(sNameProp, sValueProp)
End Sub
Private Sub lbDelOneCustProp_Click()

    If cmbMain.Value = vbNullString Then
        Call MsgBox("Нет выбранных книг!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.Value)

    Dim i           As Long
    With ListCustomProp
        i = .ListIndex
        If i < 0 Then Exit Sub
        Dim sNameProp As String
        sNameProp = .List(i, 1)
    End With
    If MsgBox("Удаление свойства [" & sNameProp & " ] ?", vbYesNo + vbQuestion, "Удаление свойства:") = vbYes Then
        Call delFilePropertyCustom(wb, sNameProp)
        Call cmbMain_Change
    End If
End Sub
Private Sub AddCustProp(ByVal txtPropName As String, ByVal txtPropValue As String)

    If cmbMain.Value = vbNullString Then
        Call MsgBox("Нет выбранных книг!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.Value)

    txtPropName = InputBox("Ведите название свойства", "Создание свойства:", txtPropName)
    If txtPropName <> vbNullString Then
        txtPropValue = InputBox("Введите значение свойства", "Создание свойства:", txtPropValue)
        If txtPropValue <> vbNullString Then
            Call addFilePropertyCustom(wb, txtPropName, txtPropValue)
            Call cmbMain_Change
        End If
    End If
End Sub

Private Sub lbDelAllCustomProp_Click()
    If cmbMain.Value = vbNullString Then
        Call MsgBox("Нет выбранных книг!", vbCritical)
        Exit Sub
    End If
    Dim wb          As Workbook
    Set wb = Workbooks(cmbMain.Value)

    If MsgBox("Удалить все свойсва?", vbYesNo + vbQuestion, "Удаление свойств:") = vbYes Then
        Dim iCount  As Byte
        iCount = delFilePropertiesCustomAll(wb)
        Call cmbMain_Change
        Call MsgBox("Свойства удалены: " & iCount, vbInformation, "Удаление свойств:")
    End If
End Sub

Private Sub UserForm_Activate()
    If Workbooks.Count = 0 Then
        Unload Me
        Call MsgBox("Нет открытых" & Chr(34) & "Excel файлов" & Chr(34) & "!", vbOKOnly + vbExclamation)
        Exit Sub
    End If
    With Me.cmbMain
        .Clear
        On Error Resume Next
        Dim wb      As Workbook
        For Each wb In Workbooks
            Call .AddItem(wb.Name)
        Next
        .Value = ActiveWorkbook.Name
        On Error GoTo 0
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
End Sub
