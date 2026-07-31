VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataFromWorkBooks 
   Caption         =   "—Ó· ‰‡ÌÌ˚ı ËÁ ÍÌË„:"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15450
   OleObjectBlob   =   "frmDataFromWorkBooks.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataFromWorkBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' UserForm      :   frmDataFromSheets
' Author        :   VBATools
' Copyright     :   Apache License
' Created       :   10-06-2026 09:37:02
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *



'================================================================================
' Œ¡–¿¡Œ“◊» »  ÕŒœŒ 
'================================================================================

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnChoseFiles_Click()
    Dim arrFiles()  As String
    Dim i As Long, iCount As Long
    Dim arrVal()    As String

    arrFiles = fileDialogFun(ActiveWorkbook.Path, True, , False)

    ' œÓ‚ÂÍ‡ ÔÛÒÚÓ„Ó Ï‡ÒÒË‚‡
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
    Call deleteItemInListBox(listWBooks)
End Sub

Private Sub btnDelItemListSheets_Click()
    Call deleteItemInListBox(listSheets)
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
                    .AddItem sGetParentFolderName(arr(i, 0)) & Application.PathSeparator
                    .List(item, 1) = itemSheet
                    .List(item, 2) = sGetFileName(arr(i, 0))
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
    On Error GoTo ErrorHandler

    Application.ReferenceStyle = xlA1

    Dim targetSheet As Worksheet
    Dim rng         As Range
    Dim iRowCount As Long, iColCount As Integer
    Dim Shift       As Byte
    Dim i           As Long

    Set targetSheet = ActiveSheet
    If chAddNewSheet.Value Then
        Set targetSheet = ActiveWorkbook.Worksheets.Add
    End If

    Set rng = Range(txtRng.Value)
    iRowCount = rng.Rows.Count
    iColCount = rng.Columns.Count

    With listSheets
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                Shift = 0

                If chbFileName.Value Then
                    activeCell.Resize(iRowCount, 1).Value = .List(i, 0)
                    activeCell.Offset(0, 1).Resize(iRowCount, 1).Value = .List(i, 2)
                    activeCell.Offset(0, 2).Resize(iRowCount, 1).Value = .List(i, 3)
                    activeCell.Offset(0, 3).Resize(iRowCount, 1).Value = txtRng.Value
                    Shift = 4
                End If

                activeCell.Offset(0, Shift).Resize(iRowCount, iColCount).formula = _
                        "='" & .List(i, 0) & "[" & .List(i, 2) & "]" & .List(i, 3) & "'!" & VBA.Split(txtRng.Value, ":")(0)

                If optOnlyValues.Value Then
                    activeCell.Offset(0, Shift).Resize(iRowCount, iColCount).Value2 = _
                            activeCell.Offset(0, Shift).Resize(iRowCount, iColCount).Value2
                End If

                activeCell.Offset(iRowCount, 0).Select
            End If
        Next i
    End With

    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Œ¯Ë·Í‡: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Œ¡–¿¡Œ“◊» » Œœ÷»… ¬€¡Œ–¿
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
' ¬¿À»ƒ¿÷»ﬂ ¬¬Œƒ¿ (ËÒÔÓÎ¸ÁÛÂÚÒˇ ‰‚‡Ê‰˚)
'================================================================================

Private Sub txtEndNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtFerstNum_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

'================================================================================
' Œ¡–¿¡Œ“◊» » œŒÀﬂ ƒ»¿œ¿«ŒÕ¿
'================================================================================

Private Sub txtRng_DropButtonClick()
    Me.Hide
    txtRng.Value = SelectRangeViaDialog(, False)
    Me.Show
End Sub

Private Sub txtRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

'================================================================================
' »Õ»÷»¿À»«¿÷»ﬂ
'================================================================================

Private Sub UserForm_Initialize()
    ' ÷ÂÌÚËÓ‚‡ÌËÂ ÙÓÏ˚
    Call CenterUserForm(Me)

    If TypeName(Selection) = "Range" Then txtRng.Value = Selection.Address
    Call ConfigureDropButton(txtRng)
End Sub



