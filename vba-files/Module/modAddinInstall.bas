Attribute VB_Name = "modAddinInstall"
Option Explicit
Option Private Module

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Sub        : Installation - add-in installation procedure
'* Created    : 22-03-2023 15:14
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub InstallationAddinExcelTools()
    Dim addFolder   As String
    Dim sFullName   As String
    Dim existingAddIn As AddIn

    On Error GoTo InstallationAdd_Err
    addFolder = VBA.Replace(Application.UserLibraryPath & Application.PathSeparator, _
            Application.PathSeparator & Application.PathSeparator, _
            Application.PathSeparator)

    If Dir(addFolder, vbDirectory) = vbNullString Then
        MsgBox "К сожалению, программа не может установить надстройку на этот компьютер." & vbCrLf & _
                "Отсутствует каталог надстроек." & vbCrLf & _
                "Пожалуйста, свяжитесь с разработчиком программы.", vbCritical, "Не удалось установить надстройку:"
        Exit Sub
    End If

    sFullName = addFolder & modAddinConst.NAME_ADDIN & ".xlam"
    On Error Resume Next
    Set existingAddIn = Application.AddIns(modAddinConst.NAME_ADDIN)
    On Error GoTo InstallationAdd_Err

    If Not existingAddIn Is Nothing Then
        If existingAddIn.Installed Then
            existingAddIn.Installed = False
        End If
    End If

    If WorkbookIsOpen(modAddinConst.NAME_ADDIN & ".xlam") Then
        MsgBox "Файл надстройки уже открыт." & vbCrLf & _
                "Возможно, он был установлен ранее.", vbCritical, "Не удалось установить надстройку:"
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    ThisWorkbook.SaveAs FileName:=sFullName, fileFormat:=xlOpenXMLAddIn

    Call AddIns.Add(FileName:=sFullName)
    AddIns(modAddinConst.NAME_ADDIN).Installed = True

    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Программа была успешно установлена!" & vbCrLf & _
            "Пожалуйста, откройте или создайте новый документ.", vbInformation, _
            "Установка надстройки:" & modAddinConst.NAME_ADDIN

    ThisWorkbook.Close False
    Exit Sub

InstallationAdd_Err:
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    If Err.Number = 1004 Then
        MsgBox "Чтобы установить надстройку, пожалуйста, закройте этот файл и запустите его снова.", _
                vbCritical, "Установка:"
    Else
        MsgBox Err.Description, vbCritical
    End If
End Sub



