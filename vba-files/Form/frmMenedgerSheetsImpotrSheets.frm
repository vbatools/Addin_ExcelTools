VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsImpotrSheets 
   Caption         =   "Импорт листов из книг:"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15450
   OleObjectBlob   =   "frmMenedgerSheetsImpotrSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsImpotrSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit



'================================================================================
' ОБРАБОТЧИКИ КНОПОК
'================================================================================

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnChoseFiles_Click()
    Dim arrFiles()  As String
    Dim i As Long, iCount As Long
    Dim arrVal()    As String

    arrFiles = FileDialogFun(ActiveWorkbook.Path, True, , False)

    ' Проверка пустого массива
    If Not IsArray(arrFiles) Then Exit Sub
    On Error Resume Next
    If UBound(arrFiles) < LBound(arrFiles) Then Exit Sub
    On Error GoTo 0

    listSheets.Clear
    iCount = UBound(arrFiles, 1)
    ReDim arrVal(1 To iCount, 1 To 2) As String

    For i = 1 To iCount
        arrVal(i, 1) = arrFiles(i, 1)
        arrVal(i, 2) = arrFiles(i, 1)
        If VBA.Len(arrVal(i, 2)) > 50 Then
            arrVal(i, 2) = ".. " & VBA.Right$(arrFiles(i, 1), 50)
        End If
    Next i

    listWBooks.List = arrVal
End Sub

Private Sub btnClearListFiles_Click()
    listWBooks.Clear
    listSheets.Clear
End Sub

Private Sub btnDelItemListFiles_Click()
    Call DeleteItemInListBox(listWBooks)
End Sub

Private Sub btnDelItemListSheets_Click()
    Call DeleteItemInListBox(listSheets)
End Sub

Private Sub btnGetSheetsNameFiles_Click()
    If listWBooks.ListCount = 0 Then Exit Sub

    listSheets.Clear
    cmbSheetName.Clear
    optAllItemsList.Value = True

    Dim arr         As Variant
    Dim Cnn         As ADODB.Connection
    Dim rs          As ADODB.Recordset
    Dim arrSheet    As Variant
    Dim dict        As Object
    Dim sheetName   As String
    Dim i As Long, j As Long
    Dim item As Long, itemSheet As Long
    Dim iCount      As Long

    arr = listWBooks.List
    iCount = UBound(arr, 1)
    Set dict = CreateObject("Scripting.Dictionary")

    On Error GoTo CleanUp

    For i = 0 To iCount
        Set Cnn = New ADODB.Connection
        Set rs = New ADODB.Recordset

        Cnn.Open "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & arr(i, 0) & ";ReadOnly=1"
        Set rs = Cnn.OpenSchema(adSchemaTables)
        arrSheet = rs.GetRows()

        rs.Close
        Cnn.Close
        itemSheet = 0

        For j = 0 To UBound(arrSheet, 2)
            sheetName = arrSheet(2, j)

            If VBA.Left$(sheetName, 1) = "'" Then sheetName = VBA.Mid$(sheetName, 2)

            If VBA.Len(sheetName) < VBA.InStr(1, sheetName, "$") + 2 Then
                itemSheet = itemSheet + 1
                sheetName = VBA.Left$(sheetName, InStr(1, sheetName, "$") - 1)
                sheetName = Replace(sheetName, "#", ".")

                With listSheets
                    .AddItem GetParentFolderName(arr(i, 0)) & Application.PathSeparator
                    .List(item, 1) = itemSheet
                    .List(item, 2) = GetFileName(arr(i, 0))
                    .List(item, 3) = sheetName
                    .Selected(item) = True
                    If Not dict.Exists(sheetName) Then
                        dict.Add sheetName, sheetName
                        cmbSheetName.AddItem sheetName
                    End If
                End With

                item = item + 1
            End If
        Next j

        Set rs = Nothing
        Set Cnn = Nothing
    Next i

CleanUp:
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not Cnn Is Nothing Then
        If Cnn.State = 1 Then Cnn.Close
        Set Cnn = Nothing
    End If
    Set dict = Nothing
End Sub

Private Sub btnOK_Click()

    Call DisableApplicationSettings

    Dim i           As Long
    Dim wb          As Workbook
    Dim sNameWB     As String
    Dim shCurent    As Object
    Dim iCount      As Long
    Dim wbAct       As Workbook
    Dim sMsg        As String
    Dim bFlag       As Boolean
    Set wbAct = ActiveWorkbook

    With listSheets
        iCount = .ListCount - 1
        For i = 0 To iCount
            If .Selected(i) Then
                If sNameWB <> .List(i, 0) & .List(i, 2) Then
                    sNameWB = .List(i, 0) & .List(i, 2)
                    If Not wb Is Nothing Then wb.Close False
                    Set wb = Workbooks.Open(sNameWB, False, True)
                    bFlag = False
                End If
                If wb.ProtectStructure And Not bFlag Then
                    If sMsg <> vbNullString Then sMsg = sMsg & vbNewLine
                    sMsg = "Структура книги защищена " & .List(i, 2) & " - импорт не возможен"
                    bFlag = True
                Else
                    Set shCurent = wb.Sheets(.List(i, 3))
                    shCurent.Visible = XlSheetVisibility.xlSheetVisible
                    If chAddNewSheet.Value Then shCurent.Name = VBA.Left$(shCurent.Name & "_" & .List(i, 2), 30)
                    shCurent.Copy After:=wbAct.Sheets(wbAct.Sheets.Count)
                End If
            End If
            lbProgress.Width = i / iCount * listWBooks.Width
            Me.Repaint
        Next i
    End With
    If Not wb Is Nothing Then wb.Close False

    Call RestoreApplicationSettings
    If sMsg <> vbNullString Then Call MsgBox(sMsg, vbCritical)
    Unload Me
End Sub

'================================================================================
' ОБРАБОТЧИКИ ОПЦИЙ ВЫБОРА
'================================================================================

Private Sub optAllItemsList_Click()
    Dim i           As Long
    If Not optAllItemsList.Value Then Exit Sub

    For i = 0 To listSheets.ListCount - 1
        listSheets.Selected(i) = True
    Next i
End Sub

Private Sub cmbSheetName_Change()
    optSheetByName.Value = True
    Call optSheetByName_Click
End Sub

Private Sub optSheetByName_Click()
    Dim i           As Long
    If Not optSheetByName.Value Then Exit Sub

    For i = 0 To listSheets.ListCount - 1
        listSheets.Selected(i) = (VBA.UCase$(listSheets.List(i, 3)) Like VBA.UCase$(cmbSheetName.Value))
    Next i
End Sub

Private Sub txtFerstNum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    optSheetByNumber.Value = True
    Call optSheetByNumber_Click
End Sub

Private Sub txtEndNum_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    optSheetByNumber.Value = True
    Call optSheetByNumber_Click
End Sub

Private Sub optSheetByNumber_Click()
    If Not optSheetByNumber.Value Then Exit Sub
    Dim i           As Long

    Dim ferstNum    As Integer
    Dim endNum      As Integer
    If txtFerstNum.Value <> vbNullString Then ferstNum = txtFerstNum.Value
    If txtEndNum.Value <> vbNullString Then endNum = txtEndNum.Value

    If ferstNum > endNum Then
        endNum = ferstNum
        ferstNum = endNum
        If txtEndNum.Value <> vbNullString Then ferstNum = txtEndNum.Value

    End If

    With listSheets
        For i = 0 To .ListCount - 1
            .Selected(i) = (VBA.CInt(.List(i, 1)) >= ferstNum And VBA.CInt(.List(i, 1)) <= endNum)
        Next i
    End With
End Sub

'================================================================================
' ВАЛИДАЦИЯ ВВОДА (используется дважды)
'================================================================================

Private Sub txtEndNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtFerstNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

'================================================================================
' ИНИЦИАЛИЗАЦИЯ
'================================================================================

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    lbProgress.Width = 0
End Sub
