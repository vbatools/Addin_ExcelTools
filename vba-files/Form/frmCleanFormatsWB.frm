VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCleanFormatsWB 
   Caption         =   "Удаление лишних стилей:"
   ClientHeight    =   8880.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9090.001
   OleObjectBlob   =   "frmCleanFormatsWB.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCleanFormatsWB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim bFlagProtect    As Boolean
Dim wbTarget        As Workbook

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDiagnostic_Click()

    If cmbWBFiles.Value = vbNullString Then
        Call MsgBox("Не выбрана книга для анализа!", vbCritical)
        Exit Sub
    Else
        Set wbTarget = Workbooks(cmbWBFiles.Value)
    End If

    On Error Resume Next

    Dim iProtectSheets As Integer
    Dim objSH       As Worksheet
    With wbTarget

        For Each objSH In wbTarget.Sheets
            If objSH.ProtectContents Then
                iProtectSheets = iProtectSheets + 1
            End If
        Next objSH

        ListMain.AddItem "": ListMain.AddItem ""
        ListMain.AddItem "Диагностика:"
        ListMain.AddItem "Именованных диапазонов " & vbTab & vbTab & vbTab & CStr(.Names.Count) & vbTab & " шт."
        ListMain.AddItem "Стилей " & vbTab & vbTab & vbTab & vbTab & vbTab & CStr(.Styles.Count) & vbTab & " шт."
        If iProtectSheets > 0 Then
            ListMain.AddItem ""
            ListMain.AddItem "Включена защита листов паролем" & vbTab & vbTab & CStr(iProtectSheets) & vbTab & " шт."
            ListMain.AddItem "Необходимо снять ПАРОЛИ со всех листов!"
            bFlagProtect = True
        End If
        ListMain.AddItem ""
        If .Names.Count + .Styles.Count > 500 Then
            ListMain.AddItem "Рекомендовано лечение. Нажмите кнопку ""Лечить"""
        Else
            ListMain.AddItem "Лечение не требуется"
        End If
    End With
End Sub

Private Sub btnFixStyles_Click()

    If wbTarget Is Nothing Then
        Call MsgBox("Не произведена диагностика файла!", vbCritical)
        Exit Sub
    End If

    If bFlagProtect Then
        Call MsgBox("Необходимо снять ПАРОЛИ с листов!", vbCritical, "Ошибка:")
        Exit Sub
    End If

    On Error Resume Next
    Dim sNewName    As String
    Dim sExt        As String
    Dim iKey        As Integer

    With wbTarget
        ListMain.AddItem "": ListMain.AddItem ""
        ListMain.AddItem "Лечение:"
        sExt = GetExtensionName(wbTarget.Name)
        sNewName = GetBaseName(wbTarget.Name) & "_CURED." & sExt
        ListMain.AddItem "Файл сохраняется под новым именем " & sNewName & " ..."
        Err.Clear
        wbTarget.SaveAs wbTarget.Path & Application.PathSeparator & sNewName
        If Err = 0 Then
            ListMain.AddItem "Сохранение прошло успешно."

            ListMain.AddItem ""
            iKey = MsgBox("Предпринимаем попытку удалить избыточные стили?", vbCritical + vbQuestion + vbYesNo, "Подтверждение операции")
            If iKey = vbYes Then
                ListMain.AddItem "Удаление избыточных стилей..."
                Call DeleteStyles
            End If

            iKey = MsgBox("Предпринимаем попытку удалить избыточные именованные диапазоны?", vbCritical + vbQuestion + vbYesNo, "Подтверждение операции")
            If iKey = vbYes Then
                ListMain.AddItem ""
                ListMain.AddItem "Удаление избыточных именованных диапазонов..."
                Call DeleteNames
            End If
            ListMain.AddItem ""
            wbTarget.Save
            ListMain.AddItem "Файл сохранен."
        Else
            ListMain.AddItem "Возникла ошибка: " & Err.Description
            ListMain.AddItem "Лечение отменено."
        End If
    End With

End Sub

Private Sub DeleteStyles()
    Dim oStyle As Style, i As Long
    Dim TotalCount  As Long
    Dim BuiltInCount As Long, FailedCount As Long, SuccessCount As Long
    On Error Resume Next
    With wbTarget
        TotalCount = .Styles.Count
        For i = .Styles.Count - 1 To 0 Step -1
            Set oStyle = .Styles(i)
            If Not oStyle.BuiltIn Then
                Err.Clear
                Application.StatusBar = "Элемент " & i
                DoEvents
                oStyle.Delete
                If Err Then
                    FailedCount = FailedCount + 1
                Else
                    SuccessCount = SuccessCount + 1
                End If
            Else
                BuiltInCount = BuiltInCount + 1
            End If
        Next
    End With
    ListMain.AddItem "Результаты:"
    ListMain.AddItem vbTab & "Успешно удалено " & vbTab & vbTab & vbTab & vbTab & vbTab & SuccessCount & vbTab & " шт."
    ListMain.AddItem vbTab & "Ошибка при попытке удаления " & vbTab & vbTab & vbTab & FailedCount & vbTab & " шт."
    ListMain.AddItem vbTab & "Пропушены встроенные стили " & vbTab & vbTab & vbTab & BuiltInCount & vbTab & " шт."
    ListMain.AddItem vbTab & "Удалено " & format(CDbl(SuccessCount / TotalCount), "0.0%")
    Application.StatusBar = ""
End Sub

Private Sub DeleteNames()
    Dim oName As Name, i As Long
    Dim FailedCount As Long, SuccessCount As Long, TotalCount As Long
    On Error Resume Next
    With wbTarget
        TotalCount = .Names.Count
        For i = .Names.Count - 1 To 0 Step -1
            Set oName = .Names(i)
            Err.Clear
            Application.StatusBar = "Элемент " & i
            DoEvents
            oName.Delete
            If Err Then
                FailedCount = FailedCount + 1
            Else
                SuccessCount = SuccessCount + 1
            End If
        Next
    End With
    ListMain.AddItem "Результаты:"
    ListMain.AddItem vbTab & "Успешно удалено " & vbTab & vbTab & vbTab & vbTab & vbTab & SuccessCount & vbTab & " шт."
    ListMain.AddItem vbTab & "Ошибка при попытке удаления " & vbTab & vbTab & vbTab & FailedCount & vbTab & " шт."
    ListMain.AddItem vbTab & "Удалено " & format(CDbl(SuccessCount / TotalCount), "0.0%")
    Application.StatusBar = ""
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)

    Dim wb          As Workbook
    For Each wb In Workbooks
        With wb
            'If .Name <> ThisWorkbook.Name And .Name <> "PERSONAL.XLSB" Then
            cmbWBFiles.AddItem .Name
            'End If
        End With
    Next wb

End Sub
