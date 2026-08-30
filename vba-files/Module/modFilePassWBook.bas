Attribute VB_Name = "modFilePassWBook"
Option Explicit
Option Private Module

'--------------------------------------------------------------------------------
' Sub: DelPasswordWBook
' Purpose:  Удаляет пароли защиты книги и листов через манипуляцию ZIP-архивом
'   - Открывает диалог выбора книги для снятия защиты
'   - Распаковывает файл как ZIP-архив
'   - Удаляет элементы защиты из XML-структуры
'   - Запаковывает файл обратно и открывает его
'--------------------------------------------------------------------------------
Public Sub DelPasswordWBook()
    Dim oForma      As frmOtherToolsDeletePasswordWB
    Set oForma = New frmOtherToolsDeletePasswordWB
    oForma.Show
    Call DisableApplicationSettings
    Dim bFlag       As Boolean
    bFlag = VBA.CBool(oForma.lbValue.Caption)
    If Not bFlag Then Exit Sub
    Dim wb          As Workbook
    Set wb = Workbooks(oForma.cmbMain.Value)
    Dim sFullName   As String
    With wb
        sFullName = .FullName
        wb.Close True
    End With

    Dim clsZIP      As clsOfficeArchiveManager
    Set clsZIP = New clsOfficeArchiveManager
    With clsZIP
        If .Initialize(sFullName, True) Then
            If .UnZipFile Then
                If oForma.lbMsg.Visible Then Call .DelPasswordWBook
                Dim arr As Variant
                arr = .getArraySheetsName()
                If Not IsEmpty(arr) Then
                    Dim i As Long
                    For i = 1 To UBound(arr, 1)
                        Call .delPasswordSheet(arr(i, 4))
                    Next i
                End If
                Call .ZipFilesInFolder
            End If
        End If
    End With
    Set clsZIP = Nothing
    Call Workbooks.Open(sFullName)
    Call RestoreApplicationSettings
    If bFlag Then
        Set oForma = New frmOtherToolsDeletePasswordWB
        oForma.Show
    End If
End Sub