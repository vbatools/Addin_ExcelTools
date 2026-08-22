Attribute VB_Name = "modAddinRibbonCallbacks"
Option Explicit
Option Private Module

Private Sub btnDataFromSheets(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    frmDataFromSheets.Show
End Sub

Private Sub btnDataToSheets(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    frmDataToSheets.Show
End Sub

Private Sub btnDataFromWBooks(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataFromWorkBooks.Show
End Sub

Private Sub btnDataToWBooks(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataToWorkBooks.Show
End Sub

Private Sub btnAddWorkBooks(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataAddWorkBooks.Show
End Sub

Private Sub btnUniqeValue(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataUniqueValues.Show
End Sub

Private Sub btnMergeValue(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataMergeText.Show
End Sub

Private Sub btnMergeDuplicatesValue(control As IRibbonControl)
    frmDataMergeTextDuplicates.Show
End Sub

Private Sub btnUnMergeValue(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call UnMergeCells
End Sub

Private Sub btnCrossSelectionOn(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call CrossSelection
End Sub

Private Sub btnCrossSelectionSettings(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmCrossSelectionSettings.Show
End Sub

Private Sub btnReferenceStyle(control As IRibbonControl)
    With Application
        If .ReferenceStyle = xlR1C1 Then
            .ReferenceStyle = xlA1
        Else
            .ReferenceStyle = xlR1C1
        End If
    End With
End Sub

Private Sub btnNewStream(control As IRibbonControl)
    If MsgBox("Запустить в новом потоке?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim App         As Application
    Set App = CreateObject("Excel.Application")
    App.Visible = True
    Call App.Workbooks.Open(FileName:=ActiveWorkbook.FullName, ReadOnly:=True)
End Sub

Private Sub btnOpenFolder(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call OpenPath(ActiveWorkbook.Path)
End Sub

Private Sub btnCopyFullNameFile(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    On Error GoTo CopyFullWay

    Dim ClipBoard   As New DataObject
    With ClipBoard
        .SetText ActiveWorkbook.FullName
        .PutInClipboard
    End With
    Exit Sub
CopyFullWay:
    MsgBox Err.Description
End Sub

Private Sub btnZeroShowHiden(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    On Error GoTo Zeros_Err
    With ActiveWindow
        .DisplayZeros = Not .DisplayZeros
    End With
    Exit Sub
Zeros_Err:
    MsgBox Err.Description & vbCrLf & "в VBAProject.D_Macros.Zeros_ " & vbCrLf & "в строке " & Erl, vbExclamation + vbOKOnly, "Ошибка:"
End Sub

Private Sub btnSplitPages(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    ActiveSheet.DisplayPageBreaks = Not ActiveSheet.DisplayPageBreaks
End Sub

Private Sub btnDisplayOutTop(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    With ActiveSheet.Outline
        .AutomaticStyles = False
        If .SummaryRow = xlBelow Then
            .SummaryRow = xlAbove
        Else
            .SummaryRow = xlBelow
        End If
    End With
End Sub

Private Sub btnDisplayOutLeft(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    With ActiveSheet.Outline
        .AutomaticStyles = False
        If .SummaryColumn = xlLeft Then
            .SummaryColumn = xlRight
        Else
            .SummaryColumn = xlLeft
        End If
    End With
End Sub

Private Sub btnDisplayOutline(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    ActiveWindow.DisplayOutline = Not ActiveWindow.DisplayOutline
End Sub


'Private Sub btnZeroColor(control As IRibbonControl)
'    MsgBox  control.id
'End Sub

Private Sub btnCopyFormuls(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmFormulaCopy.Show
End Sub

Private Sub btnConvertRC(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmFormulaConvertRC.Show
End Sub

Private Sub btnMenedgerText(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    frmFormulaMenedgerText.Show
End Sub

Private Sub btnMenegerSheets(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmMenedgerSheets.Show
End Sub

Private Sub btnMenegerWBook(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    frmMenedgerWBooks.Show
End Sub

Private Sub btnMenegerNames(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmMenedgerNames.Show
End Sub

Private Sub btnMenegerStyles(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmMenedgerStyle.Show
End Sub

Private Sub btnMenegerCharts(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmMenedgerCharts.Show
End Sub


Private Sub btnAddSheets(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call GetSheetsLists
End Sub

Private Sub btnRenameSheets(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call RenameSheets
End Sub

Private Sub btnAddSheetsFromList(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call AddSheetsByList
End Sub

Private Sub btnUCase(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call SetSheetNameUCase
End Sub

Private Sub btnLCase(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call SetSheetNameLCase
End Sub

Private Sub btnAsStrings(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call SetSheetNameAsString
End Sub

Private Sub btnWordsUCase(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call SetSheetNameAllWorldUCase
End Sub

Private Sub btnMenedgerShapes(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmMengerShapes.Show
End Sub

Private Sub btnMenedgerShapesTools(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmMengerShapesTools.Show(0)
End Sub

Private Sub btnDialogPivotShowPages(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call DialogPivotShowPages
End Sub

Private Sub btnDialogPivotFieldProperties(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call DialogPivotFieldProperties
End Sub

Private Sub btnRefreshPivotCachesClearMissingItems(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call RefreshPivotCachesClearMissingItems
End Sub

Private Sub btnAddIn(control As IRibbonControl)
    On Error GoTo ErrorHandler
    Application.Dialogs(xlDialogAddinManager).Show
    Exit Sub
ErrorHandler:
    Err.Clear
    Call MsgBox("Нет открытых книг эксель!" & Chr(34) & "Files Excel" & Chr(34) & "!", vbOKOnly + vbExclamation, "Ошибка:")
End Sub

Private Sub btnAboutAddin(control As IRibbonControl)
    Call frmAboutAddin.Show
End Sub

'--------------------------------------
Private Sub btnVBAPasswordDelete(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call unProtectVBA
End Sub

Private Sub btnExcelPasswordDelete(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call DelPasswordWBook
End Sub

Private Sub btnFileInfo(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmInfoFile.Show
End Sub

Private Sub btnCleanFormats(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmCleanFormatsWB.Show
End Sub

Private Sub btnCleanNames(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call DeleteHiddenNames
End Sub

Private Sub btnCleanLinks(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call ExternalLinkUtility
End Sub

Private Sub btnExctractingFiles(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call FileExtractorFromExcelFile
End Sub

Private Sub btnAddFilesList(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call AddFilesList
End Sub

Private Sub btnRenameFilesList(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call MoveAndRenameFiles
End Sub

Private Sub btnCommentsSheet(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call AddSheetsComments
End Sub

Private Sub btnCommentsAdd(control As IRibbonControl)
    frmOtherToolsComments.Show
End Sub

Private Sub btnCommentsShow(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call ShowHiddenComment(True)
End Sub

Private Sub btnCommentsHiden(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call ShowHiddenComment(False)
End Sub

Private Sub btnCommentsFontSize(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call SizeTextComment
End Sub

Private Sub btnAddBackUpFile(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call AddBackupFile
End Sub

Private Sub btnBackUpSettings(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmOtherToolsBackUpFile.Show
End Sub

Private Sub btnGroup(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmOtherToolsGoalSeek.Show
End Sub

Private Sub btnInsertDataUnderColumnOrRow(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmOtherToolsInsertEmptyRows.Show
End Sub

Private Sub btnRangToJSON(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataToJSON.Show
End Sub

Private Sub btnRangToCSV(control As IRibbonControl)
    If IsNotOpenWBooks Then Exit Sub
    Call frmDataToCSV.Show
End Sub

Public Function IsNotOpenWBooks() As Boolean
    IsNotOpenWBooks = Workbooks.Count = 0
    If IsNotOpenWBooks Then
        Call MsgBox("Нет открытых книг!", vbCritical)
    End If
End Function







