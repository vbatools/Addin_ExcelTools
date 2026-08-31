Attribute VB_Name = "moUDFRibbonCallbacks"
Option Explicit
Option Private Module

' ============================================================================ '
' PROCEDURES FOR RIBBON BUTTONS - PROCEDURES ДЛЯ КНОПОК ЛЕНТЫ                  '
' ============================================================================ '

Private Sub FunctionWizardShow()
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Выбирете ячейку!", vbExclamation)
        Exit Sub
    End If
    If Not Application.Dialogs(xlDialogFunctionWizard).Show Then activeCell.Clear
    Calculate
End Sub

' ============================================================================ '
' SECTION 1: TEXT AND STRING PROCESSING                                        '
' ============================================================================ '

Private Sub btnReplaceChars(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ЗАМЕНИТЬ_СИМВОЛЫ()"
        Case Else: activeCell.FormulaR1C1 = "=REPLACE_CHARS()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextLeft(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ЛЕВО_ТЕКСТ()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_LEFT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextBetween(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=МЕЖДУ_ТЕКСТ()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_BETWEEN()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnFindReplace(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=НАЙТИ_ЗАМЕНИТЬ()"
        Case Else: activeCell.FormulaR1C1 = "=FIND_REPLACE()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextRight(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ПРАВО_ТЕКСТ()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_RIGHT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnSplitString(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=РАЗБИТЬ_СТРОКУ()"
        Case Else: activeCell.FormulaR1C1 = "=SPLIT_STRING()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnConcatMulti(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=СЦЕПИТЬ_МУЛЬТИ()"
        Case Else: activeCell.FormulaR1C1 = "=CONCAT_MULTI()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTextMatchesPattern(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ТЕКСТ_СООТВЕТСТВУЕТ_ШАБЛОНУ()"
        Case Else: activeCell.FormulaR1C1 = "=TEXT_MATCHES_PATTERN()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnTranslit(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ТРАНСЛИТ()"
        Case Else: activeCell.FormulaR1C1 = "=TRANSLIT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnRemoveChars(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=УДАЛИТЬ_СИМВОЛЫ()"
        Case Else: activeCell.FormulaR1C1 = "=REMOVE_CHARS()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 2: DATA EXTRACTION FROM CELLS                                       '
' ============================================================================ '

Private Sub btnGetComment(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ПОЛУЧКОММЕНТ()"
        Case Else: activeCell.FormulaR1C1 = "=GET_COMMENT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnGetText(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ПОЛУЧТЕКСТ()"
        Case Else: activeCell.FormulaR1C1 = "=GET_TEXT()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnGetNumber(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ПОЛУЧЧИСЛО()"
        Case Else: activeCell.FormulaR1C1 = "=GET_NUMBER()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnFormulaText(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ТЕКСТФОРМУЛЫ()"
        Case Else: activeCell.FormulaR1C1 = "=FORMULA_TEXT()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 3: FORMATTING OPERATIONS (FILL AND FONT COLOR)                       '
' ============================================================================ '

Private Sub btnSumByColor(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=СУММЗАЛИВКА()"
        Case Else: activeCell.FormulaR1C1 = "=SUM_BY_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnSumByFontColor(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=СУММШРИФТ()"
        Case Else: activeCell.FormulaR1C1 = "=SUM_BY_FONT_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnCountByColor(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=СЧЕТЗАЛИВКА()"
        Case Else: activeCell.FormulaR1C1 = "=COUNT_BY_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnCountByFontColor(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=СЧЕТШРИФТ()"
        Case Else: activeCell.FormulaR1C1 = "=COUNT_BY_FONT_COLOR()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 4: INFORMATION FUNCTIONS (WORKBOOK, SHEET, USER)                     '
' ============================================================================ '

Private Sub btnWorkbookName(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ИМЯКНИГИ()"
        Case Else: activeCell.FormulaR1C1 = "=WORKBOOK_NAME()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnSheetName(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ИМЯЛИСТА()"
        Case Else: activeCell.FormulaR1C1 = "=SHEET_NAME()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnUserName(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ИМЯПОЛЬЗОВАТЕЛЯ()"
        Case Else: activeCell.FormulaR1C1 = "=USER_NAME()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnWorkbookFullPath(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ПОЛНЫЙПУТЬКНИГИ()"
        Case Else: activeCell.FormulaR1C1 = "=WORKBOOK_FULL_PATH()"
    End Select
    Call FunctionWizardShow
End Sub

' ============================================================================ '
' SECTION 5: DATA VALIDATION AND ANALYSIS                                      '
' ============================================================================ '

Private Sub btnHasLatin(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ЕЛАТИН()"
        Case Else: activeCell.FormulaR1C1 = "=HAS_LATIN()"
    End Select
    Call FunctionWizardShow
End Sub

Private Sub btnHasCyrillic(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=ЕКИРИЛЛ()"
        Case Else: activeCell.FormulaR1C1 = "=HAS_CYRILLIC()"
    End Select
    Call FunctionWizardShow
End Sub


' ============================================================================ '
' SECTION 6: QR CODE GENERATION                                                '
' ============================================================================ '

Private Sub btnCreateQR(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Select Case Application.International(xlCountryCode)
        Case 7: activeCell.FormulaR1C1 = "=СОЗДАТЬ_QR()"
        Case Else: activeCell.FormulaR1C1 = "=CREATE_QR()"
    End Select
    Call FunctionWizardShow
End Sub
