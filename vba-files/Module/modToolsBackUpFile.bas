Attribute VB_Name = "modToolsBackUpFile"
Option Explicit
Option Private Module

Public Const NAME_PROP_VERSION As String = "version"
Public Const NAME_PROP_VERSION_PATH As String = "version_path"
Public Const NAME_PROP_VERSION_DATE As String = "version_date"
Public Const NAME_PROP_VERSION_DATE_ADD As String = "version_date_add"
Public Const NAME_PATH As String = "versions_file"

Public Sub AddBackupFile()
    Dim sPath       As String
    Dim sVersion    As String
    Dim bAddDate    As Boolean
    Dim sOldWB       As String
    
    sPath = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_PATH)
    sVersion = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION)
    If sPath = vbNullString Then
        Call MsgBox("Не найдены настройки, пути сохранения файлов!", vbCritical)
        Exit Sub
    End If
    If Not FileHave(sPath, vbDirectory) Then
        Call MsgBox("Задан не существующий путь к папке для сохранения файлов!", vbCritical)
        Exit Sub
    End If
    Dim sExtension  As String
    Call DisableApplicationSettings
    With ActiveWorkbook
        sOldWB = .FullName
        sExtension = sGetExtensionName(.Name)
        bAddDate = GetOneCustomProp(ActiveWorkbook, NAME_PROP_VERSION_DATE_ADD)
        sPath = sPath & Application.PathSeparator
        If bAddDate Then sPath = sPath & VBA.format$(VBA.Date(), "yyyy_mm_dd_")
        sVersion = getVersion(sVersion)
        sPath = sPath & sGetBaseName(.Name) & "_v" & sVersion & "." & sExtension
        Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION, sVersion)
        Call addFilePropertyCustom(ActiveWorkbook, NAME_PROP_VERSION_DATE, VBA.format$(VBA.Date(), "dd.mm.yyyy"))
        .Save
        If MoveFile(sOldWB, sPath) <> vbNullString Then
            Call MsgBox("Не удалось создать резервную копию!", vbExclamation)
        Else
            Call MsgBox("Копия создан, версия: " & sVersion, vbInformation)
        End If
    End With
    Call RestoreApplicationSettings
End Sub

Private Function getVersion(ByVal sVal As String) As String
    Dim i           As Integer
    i = VBA.InStrRev(sVal, ".")
    If i > 0 Then
        getVersion = VBA.Left$(sVal, i) & VBA.Val(VBA.Right(sVal, VBA.Len(sVal) - i)) + 1
    Else
        getVersion = VBA.Val(sVal) + 1
    End If
End Function
