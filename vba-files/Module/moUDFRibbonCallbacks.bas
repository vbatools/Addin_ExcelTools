Attribute VB_Name = "moUDFRibbonCallbacks"
Option Explicit
Option Private Module

' ============================================================================ '
' PROCEDURES FOR RIBBON BUTTONS - PROCEDURES ƒÀﬂ  ÕŒœŒ  À≈Õ“€                  '
' ============================================================================ '

Private Sub FunctionWizardShow()
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("¬˚·ËÂÚÂ ˇ˜ÂÈÍÛ!", vbExclamation)
        Exit Sub
    End If
    If Not Application.Dialogs(xlDialogFunctionWizard).Show Then activeCell.Clear
    Calculate
End Sub

' ============================================================================ '
' SECTION 1: TEXT AND STRING PROCESSING                                        '
' ============================================================================ '

Private Sub btnReplaceChars(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=«¿Ã≈Õ»“‹_—»Ã¬ŒÀ€()"
        Case Else: activeCell.FormulaR1C1 = "=REPLACE_CHARS()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextLeft(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=À≈¬Œ_“≈ —“()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_LEFT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextBetween(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=Ã≈∆ƒ”_“≈ —“()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_BETWEEN()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnFindReplace(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=Õ¿…“»_«¿Ã≈Õ»“‹()"
        Case Else: activeCell.FormulaR1C1 = "=FIND_REPLACE()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextRight(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=œ–¿¬Œ_“≈ —“()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_RIGHT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnSplitString(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=–¿«¡»“‹_—“–Œ ”()"
        Case Else: activeCell.FormulaR1C1 = "=SPLIT_STRING()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnConcatMulti(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=—÷≈œ»“‹_Ã”À‹“»()"
        Case Else: activeCell.FormulaR1C1 = "=CONCAT_MULTI()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextMatchesPattern(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=“≈ —“_—ŒŒ“¬≈“—“¬”≈“_ÿ¿¡ÀŒÕ”()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_MATCHES_PATTERN()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTranslit(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=“–¿Õ—À»“()"
        Case Else: activeCell.FormulaR1C1 = "=TRANSLIT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnRemoveChars(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=”ƒ¿À»“‹_—»Ã¬ŒÀ€()"
        Case Else: activeCell.FormulaR1C1 = "=REMOVE_CHARS()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 2: DATA EXTRACTION FROM CELLS                                       '
' ============================================================================ '

Private Sub btnGetComment(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=œŒÀ”◊ ŒÃÃ≈Õ“()"
        Case Else: activeCell.FormulaR1C1 = "=GET_COMMENT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnGetText(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=œŒÀ”◊“≈ —“()"
        Case Else: activeCell.FormulaR1C1 = "=GET_TEXT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnGetNumber(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=œŒÀ”◊◊»—ÀŒ()"
        Case Else: activeCell.FormulaR1C1 = "=GET_NUMBER()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnFormulaText(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=“≈ —“‘Œ–Ã”À€()"
        Case Else: activeCell.FormulaR1C1 = "=FORMULA_TEXT()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 3: FORMATTING OPERATIONS (FILL AND FONT COLOR)                       '
' ============================================================================ '

Private Sub btnSumByColor(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=—”ÃÃ«¿À»¬ ¿()"
        Case Else: activeCell.FormulaR1C1 = "=SUM_BY_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnSumByFontColor(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=—”ÃÃÿ–»‘“()"
        Case Else: activeCell.FormulaR1C1 = "=SUM_BY_FONT_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnCountByColor(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=—◊≈“«¿À»¬ ¿()"
        Case Else: activeCell.FormulaR1C1 = "=COUNT_BY_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnCountByFontColor(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=—◊≈“ÿ–»‘“()"
        Case Else: activeCell.FormulaR1C1 = "=COUNT_BY_FONT_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 4: INFORMATION FUNCTIONS (WORKBOOK, SHEET, USER)                     '
' ============================================================================ '

Private Sub btnWorkbookName(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=»Ãﬂ Õ»√»()"
        Case Else: activeCell.FormulaR1C1 = "=WORKBOOK_NAME()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnSheetName(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=»ÃﬂÀ»—“¿()"
        Case Else: activeCell.FormulaR1C1 = "=SHEET_NAME()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnUserName(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=»ÃﬂœŒÀ‹«Œ¬¿“≈Àﬂ()"
        Case Else: activeCell.FormulaR1C1 = "=USER_NAME()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnWorkbookFullPath(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=œŒÀÕ€…œ”“‹ Õ»√»()"
        Case Else: activeCell.FormulaR1C1 = "=WORKBOOK_FULL_PATH()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 5: DATA VALIDATION AND ANALYSIS                                      '
' ============================================================================ '

Private Sub btnHasLatin(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=≈À¿“»Õ()"
        Case Else: activeCell.FormulaR1C1 = "=HAS_LATIN()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnHasCyrillic(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=≈ »–»ÀÀ()"
        Case Else: activeCell.FormulaR1C1 = "=HAS_CYRILLIC()"
    End Select
    Call FunctionWizardShow
End Sub


' ============================================================================ '
' SECTION 6: QR CODE GENERATION                                                '
' ============================================================================ '

Private Sub btnCreateQR(control As IRibbonControl)
    If isNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=—Œ«ƒ¿“‹_QR()"
        Case Else: activeCell.FormulaR1C1 = "=CREATE_QR()"
    End Select
    Call FunctionWizardShow
End Sub
