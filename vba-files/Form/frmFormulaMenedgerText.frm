VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFormulaMenedgerText 
   Caption         =   "Операции с текстом:"
   ClientHeight    =   8970.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmFormulaMenedgerText.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFormulaMenedgerText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Const sNEW_LINE As String = "|.-.|"

Enum TypeData
    isOtherType
    isDateType
    isTextType
    isNumerType
End Enum

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnChangFormat_Click()
    Call DisableApplicationSettings
    Dim actRng      As Range
    Dim CurentRng   As Range
    Dim bFlag       As Boolean
    Set actRng = activeCell
    Set CurentRng = Cells(1000000, 16000)
    CurentRng.Select

    With CurentRng
        Select Case True
            Case TogBtnCherta.Value: .Font.Underline = xlUnderlineStyleSingle
            Case TogBtnTwoCherta.Value: .Font.Underline = xlUnderlineStyleDouble
        End Select
        .Font.Size = txtFontSize.Value
        .Font.Name = txtFontName.Value
        .Font.Color = txtFontName.ForeColor
        .Font.Bold = txtFontName.Font.Bold
        .Font.Italic = txtFontName.Font.Italic
    End With

    If Application.Dialogs(xlDialogFontProperties).Show Then
        bFlag = True
    End If
    If Not bFlag Then Set CurentRng = actRng

    With CurentRng
        TogBtnFat.Value = .Font.Bold
        TogBtnKursiv.Value = .Font.Italic
        Select Case .Font.Underline
            Case xlUnderlineStyleSingle: TogBtnCherta.Value = True
            Case xlUnderlineStyleDouble: TogBtnTwoCherta.Value = True
            Case Else:
                TogBtnCherta.Value = False
                TogBtnTwoCherta.Value = False
        End Select
        txtFontSize.Value = .Font.Size
        txtFontName.Value = .Font.Name
        txtFontName.ForeColor = .Font.Color
        txtFontName.Font.Bold = .Font.Bold
        txtFontName.Font.Italic = .Font.Italic
    End With
    If bFlag Then Call CurentRng.ClearFormats
    actRng.Select
    Call RestoreApplicationSettings
End Sub

Private Sub btnOK_Click()
    Dim rng         As Range
    Dim arrData     As Variant
    Dim typeDataVal As TypeData
    Dim bFlag       As Boolean

    Set rng = Range(txtInputRng.Value)
    arrData = rng.formula

    If Not IsArray(arrData) Then
        ReDim arr(1 To 1, 1 To 1)
        arr(1, 1) = arrData
        arrData = arr
    End If

    Select Case MultiPageMain.Value
        Case 0
            arrData = RegistrWords(arrData, chbFormulsCells.Value)
        Case 1
            'удаление
            arrData = deleteCharts(arrData, chbFormulsCells.Value)
        Case 2
            'вставка
            arrData = insertCharts(arrData, chbFormulsCells.Value)
        Case 3
            'исправить
            arrData = fixCharts(arrData, chbFormulsCells.Value, typeDataVal)
        Case 4
            'Прочее
            arrData = otherCharts(arrData, chbFormulsCells.Value)
        Case 5
            bFlag = True
            If optFindeAndReplace.Value Then
                Call findeAndReplaceCharts(rng)
            Else
                Call findeAndReplaceSaveFormatsCharts(rng, txtTextFormat.Value)
            End If
    End Select

    Call SaveUndoInfo(rng, False, False)
    If Not bFlag Then rng.formula = arrData
    
    If MultiPageMain.Value = 4 Then
        Select Case True
            Case (optOtherShowEngLetter.Value And Not chbOtherTranslit.Value)
                Call showColorLetter(rng, "[A-Za-z]")
            Case (optOtherShowRusLetter.Value And Not chbOtherTranslit.Value)
                Call showColorLetter(rng, "[А-Яа-яЁё]")
        End Select
    End If

    Select Case typeDataVal
        Case TypeData.isDateType
            rng.NumberFormat = "mm/dd/yyyy"
        Case TypeData.isTextType
            rng.NumberFormat = "@"
        Case TypeData.isNumerType
            rng.NumberFormat = "General"
    End Select
    Application.OnUndo "Отменить", "RestoreUndoInfo"
End Sub

Private Sub findeAndReplaceCharts(ByRef rng As Range)
    Dim sValReplace As String
    Dim sValFinde   As String
    sValFinde = txtFindeVal.Value
    sValReplace = txtReplaceVal.Value
    If sValFinde = sValReplace Then
        Exit Sub
    End If

    Dim oCell       As Range
    For Each oCell In rng
        If oCell.Value <> vbNullString Then
            If VBA.Left$(oCell.Value, 1) = "=" Then GoTo skipeValue
            Call replaceFotn(oCell, sValFinde, sValReplace)
        End If
skipeValue:
    Next oCell
End Sub

Public Sub findeAndReplaceSaveFormatsCharts(ByRef rng As Range, ByVal sVal As String)
    If sVal = vbNullString Then Exit Sub

    Dim oCell       As Range
    For Each oCell In rng
        If oCell.Value <> vbNullString Then
            If VBA.Left$(oCell.Value, 1) = "=" Then GoTo skipeValue
            Dim iStart As Integer
            Dim iLength As Integer
            iStart = 1
            iLength = VBA.Len(sVal)
            With oCell
                Do While iStart <> 0
                    iStart = VBA.InStr(iStart, .Value, sVal)
                    If iStart > 0 Then
                        Call changeFotn(oCell, iStart, iLength)
                        iStart = iStart + iLength
                    End If
                Loop
            End With
        End If
skipeValue:
    Next oCell
End Sub
Private Sub changeFotn(ByRef oCell As Range, ByVal iStart As Integer, ByVal iLength As Integer)
    With oCell.Characters(Start:=iStart, Length:=iLength).Font
        .Name = txtFontName.Value
        .Bold = txtFontName.Font.Bold
        .Italic = txtFontName.Font.Italic
        .Size = txtFontSize.Value
        .Color = txtFontName.ForeColor
        Select Case True
            Case TogBtnCherta.Value: .Underline = xlUnderlineStyleSingle
            Case TogBtnTwoCherta.Value: .Underline = xlUnderlineStyleDouble
        End Select
    End With
End Sub

Private Sub showColorLetter(ByRef rng As Range, ByVal alfavit As String)
    Dim rngItem     As Range
    Dim sValue      As String
    Dim i           As Integer
    For Each rngItem In rng
        If (Not IsEmpty(rngItem.Value2)) And (Not IsError(rngItem.Value2)) Then
            sValue = rngItem.Value2
            For i = 1 To Len(sValue)
                If Mid(sValue, i, 1) Like alfavit Then
                    rngItem.Characters(Start:=i, Length:=1).Font.ColorIndex = 3
                End If
            Next i
        End If
    Next rngItem
End Sub

Private Function otherCharts(ByVal arr As Variant, ByVal bFormuls As Boolean) As Variant
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long

    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)
    ReDim arrRes(1 To iCount, 1 To jCount)
    arrRes = arr
    For i = 1 To iCount
        For j = 1 To jCount
            If arrRes(i, j) <> vbNullString Then
                If bFormuls Then
                    If VBA.Left$(arrRes(i, j), 1) = "=" Then GoTo skipeValue
                End If
                If Not chbOtherTranslit.Value Then
                    Select Case True
                        Case optOtherShowNotPrintSimbol.Value
                            arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(32), VBA.Chr$(149))
                            arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(10), VBA.Chr$(182) & VBA.Chr$(10))
                            arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(160), VBA.Chr$(176))
                        Case optOtherNoShowNotPrintSimbol.Value
                            arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(149), VBA.Chr$(32))
                            arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(182) & VBA.Chr$(10), VBA.Chr$(10))
                            arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(176), VBA.Chr$(160))
                    End Select
                Else
                    Dim iStandart As Byte
                    Select Case True
                        Case optCommon.Value: iStandart = 0
                        Case optIKAO.Value: iStandart = 1
                        Case optGOST.Value: iStandart = 2
                        Case optPCGN.Value: iStandart = 3
                        Case optPCGNY.Value: iStandart = 4
                    End Select
                    arrRes(i, j) = TRANSLIT(arr(i, j), iStandart)
                End If
            End If
skipeValue:
        Next j
    Next i
    otherCharts = arrRes
End Function

Private Function fixCharts(ByVal arr As Variant, ByVal bFormuls As Boolean, ByRef typeDataVal As TypeData) As Variant
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long

    Dim sChr        As String * 1
    Dim k           As Integer

    Const RUS       As String = "асекорхуАВСЕНКМОРТХ"
    Const ENG = "acekopxyABCEHKMOPTX"


    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)
    ReDim arrRes(1 To iCount, 1 To jCount)
    arrRes = arr
    For i = 1 To iCount
        For j = 1 To jCount
            If arrRes(i, j) <> vbNullString Then
                If bFormuls Then
                    If VBA.Left$(arrRes(i, j), 1) = "=" Then GoTo skipeValue
                End If
                Select Case True
                    Case optFixNumberWithMinus.Value
                        typeDataVal = isNumerType
                        If IsNumeric(arrRes(i, j)) Then arrRes(i, j) = VBA.CDbl(arrRes(i, j))
                    Case optFixTextAsNumber.Value
                        typeDataVal = isNumerType
                        arrRes(i, j) = VBA.Replace(arrRes(i, j), Chr(160), vbNullString)
                        arrRes(i, j) = VBA.Replace(arrRes(i, j), " ", vbNullString)
                        Select Case cmbFixListNumber.ListIndex
                            Case 0
                                'Авто
                            Case 1
                                '9 876,54
                                arrRes(i, j) = VBA.Replace(arrRes(i, j), ",", ".")
                            Case 2
                                '9.876,54
                                arrRes(i, j) = VBA.Replace(arrRes(i, j), ".", vbNullString)
                                arrRes(i, j) = VBA.Replace(arrRes(i, j), ",", ".")
                            Case 3
                                '9,876.54
                                arrRes(i, j) = VBA.Replace(arrRes(i, j), ",", vbNullString)
                            Case 4
                                '9 876.54
                                arrRes(i, j) = VBA.Replace(arrRes(i, j), " ", vbNullString)
                            Case 5
                                '9'876.54
                                arrRes(i, j) = VBA.Replace(arrRes(i, j), "'", vbNullString)
                        End Select
                        If IsNumeric(arrRes(i, j)) Then arrRes(i, j) = VBA.CDbl(arrRes(i, j))
                    Case optFixTextAsDate.Value
                        typeDataVal = isDateType
                        Dim arrDate As Variant
                        Dim sVal As String
                        sVal = VBA.Trim$(arrRes(i, j))
                        sVal = VBA.Split(sVal, ".")
                        sVal = VBA.Replace(sVal, ",", ".")
                        sVal = VBA.Replace(sVal, "/", ".")
                        sVal = VBA.Replace(sVal, "\", ".")
                        sVal = VBA.Replace(sVal, "-", ".")
                        sVal = VBA.Replace(sVal, " ", ".")
                        sVal = VBA.Replace(sVal, "'", ".")
                        arrDate = VBA.Split(sVal, ".")
                        Select Case cmbFixListDate
                            Case 0
                                'Д.М.Г
                                arrRes(i, j) = VBA.DateSerial(arrDate(2), arrDate(1), arrDate(0))
                            Case 1
                                'М.Д.Г
                                arrRes(i, j) = VBA.DateSerial(arrDate(2), arrDate(0), arrDate(1))
                            Case 2
                                'Г.М.Д
                                arrRes(i, j) = VBA.DateSerial(arrDate(0), arrDate(1), arrDate(2))
                            Case 3
                                'ГГГГММДД
                                If VBA.Len(sVal) = 8 Then arrRes(i, j) = VBA.DateSerial(VBA.Left$(sVal, 4), VBA.Mid$(sVal, 5, 2), VBA.Right$(sVal, 2))
                            Case 4
                                'ГГММДД
                                If VBA.Len(sVal) = 6 Then arrRes(i, j) = VBA.DateSerial(VBA.Left$(sVal, 2), VBA.Mid$(sVal, 3, 2), VBA.Right$(sVal, 2))
                            Case 5
                                'Другой
                                If IsDate(arrRes(i, j)) Then arrRes(i, j) = VBA.DateValue(arrRes(i, j))
                        End Select
                    Case optFixNumberAsText.Value
                        typeDataVal = isTextType
                        Select Case cmbFixListTextAsNumber.ListIndex
                            Case 0
                            Case 1
                                arrRes(i, j) = "'" & arrRes(i, j)
                        End Select
                    Case optFixEngLattersRus.Value
                        If arrRes(i, j) Like "*[" & ENG & "]*" Then
                            For k = 1 To VBA.Len(arrRes(i, j))
                                sChr = VBA.Mid$(arrRes(i, j), i, 1)
                                If sChr Like "[" & ENG & "]" Then
                                    arrRes(i, j) = VBA.Replace(arrRes(i, j), sChr, VBA.Mid$(RUS, VBA.InStr(1, ENG, sChr), 1))
                                End If
                            Next k
                        End If
                    Case optFixRusLattersEng.Value
                        If arrRes(i, j) Like "*[" & RUS & "]*" Then
                            For k = 1 To VBA.Len(arrRes(i, j))
                                sChr = VBA.Mid$(arrRes(i, j), i, 1)
                                If sChr Like "[" & RUS & "]" Then
                                    arrRes(i, j) = VBA.Replace(arrRes(i, j), sChr, VBA.Mid$(ENG, VBA.InStr(1, RUS, sChr), 1))
                                End If
                            Next k
                        End If
                    Case optChngeCommaOnPointDigital.Value
                        typeDataVal = isTextType
                        arrRes(i, j) = "'" & VBA.Replace(arrRes(i, j), ",", ".")
                End Select
            End If
skipeValue:
        Next j
    Next i
    fixCharts = arrRes
End Function

Private Function insertCharts(ByVal arr As Variant, ByVal bFormuls As Boolean) As Variant
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long

    Dim sText       As String
    sText = txtInsText.Value
    If sText = vbNullString Then
        insertCharts = arr
        Exit Function
    End If

    Dim sBefore     As String
    Dim sAfter      As String

    sBefore = txtInsertBefore.Value
    sAfter = txtInserAfter.Value

    Dim iPos        As Integer
    Dim k           As Integer
    iPos = VBA.Val(txtInsertPosiotion.Value)

    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)
    ReDim arrRes(1 To iCount, 1 To jCount)
    arrRes = arr
    For i = 1 To iCount
        For j = 1 To jCount
            If arr(i, j) <> vbNullString Then
                If bFormuls Then
                    If VBA.Left$(arr(i, j), 1) = "=" Then GoTo skipeValue
                End If

                Select Case True
                    Case optInsertAfterText.Value
                        arrRes(i, j) = arrRes(i, j) & sText
                    Case optInsertBeforeText.Value
                        arrRes(i, j) = sText & arrRes(i, j)
                    Case optInsertTextPosion.Value
                        If iPos = 0 Then GoTo skipeValue
                        If iPos > VBA.Len(arrRes(i, j)) Then iPos = VBA.Len(arrRes(i, j))
                        arrRes(i, j) = VBA.Left$(arrRes(i, j), iPos) & sText & VBA.Right$(arrRes(i, j), VBA.Len(arrRes(i, j)) - iPos + 1)
                    Case optInsertUntilTextChr.Value

                        If sBefore = vbNullString Then GoTo skipeValue
                        k = VBA.InStr(1, arr(i, j), sBefore)
                        If k = 0 Then GoTo skipeValue

                        If k - 1 = 0 Then
                            arrRes(i, j) = sText & arrRes(i, j)
                        Else
                            arrRes(i, j) = VBA.Left(arrRes(i, j), k - 1) & sText & VBA.Right(arrRes(i, j), VBA.Len(arrRes(i, j)) - k + 1)
                        End If

                    Case optInsertAfterTextChr.Value
                        If sAfter = vbNullString Then GoTo skipeValue
                        k = VBA.InStr(1, arr(i, j), sAfter)
                        If k = 0 Then GoTo skipeValue

                        If VBA.Len(sAfter) + k - 1 = VBA.Len(arrRes(i, j)) Then
                            arrRes(i, j) = arrRes(i, j) & sText
                        Else
                            arrRes(i, j) = VBA.Left(arrRes(i, j), VBA.Len(sAfter) + k - 1) & sText & VBA.Right(arrRes(i, j), VBA.Len(arrRes(i, j)) - VBA.Len(sAfter) - k + 1)
                        End If
                End Select
            End If
skipeValue:
        Next j
    Next i
    insertCharts = arrRes
End Function

Private Function deleteCharts(ByVal arr As Variant, ByVal bFormuls As Boolean) As Variant
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long
    Dim iNumDel     As Integer
    Dim iNumDelLast As Integer

    Dim sText       As String

    Dim m           As Integer
    Dim k           As Integer
    Dim sLeter      As String * 1


    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)
    ReDim arrRes(1 To iCount, 1 To jCount)
    arrRes = arr
    sText = txtUntilText.Value
    If sText = vbNullString Then sText = txtAfterText.Value
    For i = 1 To iCount
        For j = 1 To jCount
            If arr(i, j) <> vbNullString Then

                If bFormuls Then
                    If VBA.Left$(arr(i, j), 1) = "=" Then
                        arrRes(i, j) = arr(i, j)
                        GoTo skipeValue
                    End If
                End If
                Select Case True
                    Case optDeleteLeftCharts.Value
                        If iNumDel = 0 Then iNumDel = VBA.Val(txtDeleteLeftSymbols.Value)
                        If VBA.Len(arr(i, j)) > iNumDel Then arrRes(i, j) = VBA.Left(arr(i, j), iNumDel)
                    Case optDeleteRightCharts.Value
                        If iNumDel = 0 Then iNumDel = VBA.Val(txtDeleteRightSymbols.Value)
                        If VBA.Len(arr(i, j)) - iNumDel >= 0 Then arrRes(i, j) = VBA.Right(arr(i, j), VBA.Len(arr(i, j)) - iNumDel)
                    Case optDeleteMidel.Value
                        If iNumDel = 0 Then iNumDel = VBA.Val(txtDeleteMidleStartPosition.Value)
                        If iNumDelLast = 0 Then iNumDelLast = VBA.Val(txtDelMidleLetters.Value)
                        If VBA.Len(arr(i, j)) > iNumDel Then
                            If chbUndo.Value Then
                                arrRes(i, j) = VBA.Mid(arr(i, j), iNumDel, iNumDelLast)
                            Else
                                arrRes(i, j) = VBA.Left(arr(i, j), iNumDel)
                                If VBA.Len(arr(i, j)) - (iNumDelLast + iNumDel) >= 0 Then arrRes(i, j) = arrRes(i, j) & VBA.Right(arr(i, j), VBA.Len(arr(i, j)) - iNumDelLast - iNumDel)
                            End If
                        End If
                    Case optDeleteTrim.Value
                        arrRes(i, j) = VBA.Replace(arr(i, j), Chr(160), " ")
                        arrRes(i, j) = WorksheetFunction.Trim(arrRes(i, j))
                    Case optDeleteNotPrintSimbols.Value
                        arrRes(i, j) = VBA.Replace(arr(i, j), Chr(160), " ")
                        arrRes(i, j) = WorksheetFunction.Clean(arrRes(i, j))
                    Case optDeleteApostrofs.Value
                        If VBA.Left$(arrRes(i, j), 1) = "'" And VBA.Len(arrRes(i, j)) > 1 Then arrRes(i, j) = VBA.Right(arrRes(i, j), VBA.Len(arrRes(i, j)) - 1)
                    Case optDeleteAllLetters.Value
                        If arrRes(i, j) Like "*[0-9]" Then
                            Dim bFlag As Boolean
                            ReDim arrWord(0 To VBA.Len(arrRes(i, j)))
                            k = 0
                            For m = 1 To VBA.Len(arrRes(i, j))
                                sLeter = VBA.Mid$(arrRes(i, j), m, 1)
                                Select Case VBA.Asc(sLeter)
                                    Case 48 To 57
                                        bFlag = True
                                        arrWord(k) = sLeter
                                        k = k + 1
                                    Case 46, 44
                                        If bFlag Then
                                            arrWord(k) = sLeter
                                            k = k + 1
                                        End If
                                        bFlag = False
                                    Case Else
                                        bFlag = False
                                End Select
                            Next m
                            arrRes(i, j) = VBA.Join(arrWord, vbNullString)
                        Else
                            arrRes(i, j) = vbNullString
                        End If
                    Case optDeleteAllNumbers.Value
                        If arrRes(i, j) Like "*[0-9]" Then
                            ReDim arrWord(0 To VBA.Len(arrRes(i, j)))
                            k = 0
                            For m = 1 To VBA.Len(arrRes(i, j))
                                sLeter = VBA.Mid$(arrRes(i, j), m, 1)
                                Select Case VBA.Asc(sLeter)
                                    Case 32, 65 To 90, 97 To 122, 192 To 255
                                        arrWord(k) = sLeter
                                        k = k + 1
                                End Select
                            Next m
                            arrRes(i, j) = VBA.Join(arrWord, vbNullString)
                        Else
                            arrRes(i, j) = vbNullString
                        End If
                    Case optDeleteUntilText.Value
                        If sText = vbNullString Then Exit For
                        k = VBA.InStr(1, arrRes(i, j), sText) - 1
                        If k > 0 Then arrRes(i, j) = VBA.Left(arrRes(i, j), k)

                    Case optDeleteAfterText.Value
                        If sText = vbNullString Then Exit For
                        k = VBA.InStr(1, arrRes(i, j), sText)
                        If k > 0 Then arrRes(i, j) = VBA.Right(arrRes(i, j), VBA.Len(arrRes(i, j)) - k - VBA.Len(sText) + 1)
                    Case optDeletetAltEnter.Value
                        arrRes(i, j) = VBA.Replace(arrRes(i, j), VBA.Chr$(10), vbNullString)
                        k = VBA.InStr(1, arrRes(i, j), sText)
                End Select
            End If
skipeValue:
        Next j
    Next i
    deleteCharts = arrRes
End Function

Private Function RegistrWords(ByVal arr As Variant, ByVal bFormuls As Boolean) As Variant

    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long


    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)
    ReDim arrRes(1 To iCount, 1 To jCount)
    arrRes = arr
    For i = 1 To iCount
        For j = 1 To jCount
            If arr(i, j) <> vbNullString Then

                If bFormuls Then
                    If VBA.Left$(arr(i, j), 1) = "=" Then GoTo skipeValue
                End If
                Select Case True
                    Case optUCase.Value
                        arrRes(i, j) = caseString(arr(i, j), tpUCase)
                    Case optLCase.Value
                        arrRes(i, j) = caseString(arr(i, j), tpLCase)
                    Case optTitle.Value
                        arrRes(i, j) = caseString(arr(i, j), tpAsString)
                    Case optAllWordsTitle.Value
                        arrRes(i, j) = caseString(arr(i, j), tpAllWorldUCase)
                End Select
            End If
skipeValue:
        Next j
    Next i
    RegistrWords = arrRes
End Function

Private Sub optFindeAndReplace_Change()
    txtTextFormat.Enabled = Not optFindeAndReplace.Value
    btnChangFormat.Enabled = Not optFindeAndReplace.Value
    txtFindeVal.Enabled = optFindeAndReplace.Value
    txtReplaceVal.Enabled = optFindeAndReplace.Value
End Sub

Private Sub TogBtnTwoCherta_Click()
    txtFontName.Font.Strikethrough = TogBtnTwoCherta.Value
End Sub

Private Sub TogBtnCherta_Click()
    txtFontName.Font.Underline = TogBtnCherta.Value
End Sub

Private Sub TogBtnFat_Click()
    txtFontName.Font.Bold = Not txtFontName.Font.Bold
End Sub

Private Sub TogBtnKursiv_Click()
    txtFontName.Font.Italic = Not txtFontName.Font.Italic
End Sub

Private Sub txtDeleteLeftSymbols_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtDeleteMidleStartPosition_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtDeleteRightSymbols_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtDelMidleLetters_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtFontSize_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtInsertPosiotion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtInputRng_DropButtonClick()
    Me.Hide
    txtInputRng.Value = SelectRangeViaDialog()
    Me.Show
End Sub

Private Sub txtInputRng_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Public Sub replaceFotn(ByRef oCell As Range, ByVal sValFinde As String, ByVal sValReplace As String)
    Dim xmlVal      As String
    If Not oCell.Value Like "*" & sValFinde & "*" Then Exit Sub
    xmlVal = oCell.Value(xlRangeValueXMLSpreadsheet)
    xmlVal = VBA.Replace(xmlVal, "&#10;", sNEW_LINE)

    Dim oXMLDoc     As MSXML2.DOMDocument
    Set oXMLDoc = New MSXML2.DOMDocument
    Call oXMLDoc.loadXML(xmlVal)
    Dim oXMLDataNode As MSXML2.IXMLDOMNode
    Set oXMLDataNode = oXMLDoc.SelectSingleNode("//ss:Data")
    If oXMLDataNode Is Nothing Then Set oXMLDataNode = oXMLDoc.SelectSingleNode("//Data")
    If oXMLDataNode Is Nothing Then Exit Sub
    Dim i           As Long
    Dim iCount      As Long
    Dim sText       As String
    With oXMLDataNode
        iCount = .ChildNodes.Length
        If iCount = 0 Then Exit Sub
        For i = 0 To iCount - 1
            sText = .ChildNodes(i).TEXT
            If VBA.InStr(1, sText, sValFinde) > 0 Then
                Dim oXMLDataNodeF As MSXML2.IXMLDOMNode
                Dim oXMLDataNodeF1 As MSXML2.IXMLDOMNode
                Set oXMLDataNodeF = findeLastChildNode(.ChildNodes(i), oXMLDataNodeF1)
                oXMLDataNodeF1.TEXT = VBA.Replace(sText, sValFinde, sValReplace)
            End If
        Next i
    End With
    xmlVal = oXMLDoc.XML
    xmlVal = VBA.Replace(xmlVal, sNEW_LINE, "&#10;")
    oCell.Value(xlRangeValueXMLSpreadsheet) = xmlVal
End Sub
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Function   : findeLastChildNode - поиск самого последнего node
'* Created    : 22-03-2022 10:18
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* Argument(s):                                 Description
'*
'* ByRef oXMLDataNode As MSXML2.IXMLDOMNode :
'* ByRef findeNode As MSXML2.IXMLDOMNode    :
'*
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function findeLastChildNode(ByRef oXMLDataNode As MSXML2.IXMLDOMNode, ByRef findeNode As MSXML2.IXMLDOMNode) As MSXML2.IXMLDOMNode
    With oXMLDataNode
        If .ChildNodes.Length = 0 Then
            Set findeNode = oXMLDataNode
            Exit Function
        Else
            Call findeLastChildNode(.LastChild, findeNode)
        End If
    End With
End Function


Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    If TypeName(Selection) = "Range" Then txtInputRng.Value = Selection.Address
    Call ConfigureDropButton(txtInputRng)

    With cmbFixListTextAsNumber
        .AddItem "форматом"
        .AddItem "добавить апостроф"
        .ListIndex = 0
    End With

    With cmbFixListNumber
        .AddItem "Авто"
        .AddItem "9 876,54"
        .AddItem "9.876,54"
        .AddItem "9,876.54"
        .AddItem "9 876.54"
        .AddItem "9'876.54"
        .ListIndex = 0
    End With

    With cmbFixListDate
        .AddItem "Д.М.Г"
        .AddItem "М.Д.Г"
        .AddItem "Г.М.Д"
        .AddItem "ГГГГММДД"
        .AddItem "ГГММДД"
        .AddItem "Другой"
        .ListIndex = 0
    End With

    With activeCell
        If Not IsNull(.Font.Size) Then
            txtFontSize.Value = .Font.Size
        Else
            txtFontSize.Value = 12
        End If
        txtFontName.Value = .Font.Name
        If Not IsNull(.Font.Color) Then txtFontName.ForeColor = .Font.Color
        TogBtnFat.Value = .Font.Bold
        TogBtnKursiv.Value = .Font.Italic
        Select Case .Font.Underline
            Case xlUnderlineStyleSingle: TogBtnCherta.Value = True
            Case xlUnderlineStyleDouble: TogBtnTwoCherta.Value = True
            Case Else:
                TogBtnCherta.Value = False
                TogBtnTwoCherta.Value = False
        End Select
    End With
End Sub
