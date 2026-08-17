VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMenedgerSheetsExportSheets 
   Caption         =   "Экспорт листов как книги:"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmMenedgerSheetsExportSheets.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMenedgerSheetsExportSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmSheetsMenedgeSaveSheetAsWB- add description!
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   18-06-2026 11:15:58
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()

    Dim sPathParent As String
    sPathParent = txtDir.Value & Application.PathSeparator

    If Not FileHave(sPathParent, vbDirectory) Then
        Call MsgBox("Задан не существующий путь к файлу!", vbCritical)
        Exit Sub
    End If

    If chbParentName.Value Then
        sPathParent = sPathParent & GetBaseName(ActiveWorkbook.Name) & Application.PathSeparator
        If Not FileHave(sPathParent, vbDirectory) Then Call MkDir(sPathParent)
    End If

    If chDateExport.Value Then
        sPathParent = sPathParent & VBA.format$(VBA.Now(), "dd.mm.yyyy_hh.mm.ss") & Application.PathSeparator
        If Not FileHave(sPathParent, vbDirectory) Then Call MkDir(sPathParent)
    End If

    Dim wbCurent    As Workbook
    Dim wb          As Workbook
    Dim Sh          As Object
    Dim iCountFiles As Long
    Set wbCurent = ActiveWorkbook

    Call DisableApplicationSettings

    Dim i           As Long
    Dim iCount      As Long
    Dim sPath       As String
    Dim shVisible   As XlSheetVisibility
    With frmMenedgerSheets.listSheets
        iCount = .ListCount - 1
        For i = 0 To iCount
            If .Selected(i) Then
                sPath = sPathParent & CleanFileName(.List(i, 1))
                If chAddSheetsNames.Value Then
                    sPath = sPath & VBA.format$(VBA.Now(), "_dd.mm.yyyy_hh.mm.ss")
                End If
                Set Sh = wbCurent.Sheets(.List(i, 1))
                shVisible = Sh.Visible
                Sh.Visible = XlSheetVisibility.xlSheetVisible
                Sh.Copy
                Set wb = ActiveWorkbook


                If chbBreakLinks.Value Then
                    Dim WbLinks As Variant
                    Dim j As Long
                    WbLinks = wb.LinkSources(Type:=xlLinkTypeExcelLinks)
                    If Not IsEmpty(WbLinks) Then
                        For j = 1 To UBound(WbLinks)
                            wb.BreakLink Name:=WbLinks(j), Type:=xlLinkTypeExcelLinks
                        Next
                    End If
                End If

                If chbDelFormuls.Value Then
                    wb.ActiveSheet.Cells.Copy
                    wb.ActiveSheet.Range("A1").PasteSpecial Paste:=xlValues
                End If

                Call SaveWBook(wb, sPath)
                Sh.Visible = shVisible
                iCountFiles = iCountFiles + 1
            End If
        Next i
    End With
    Call RestoreApplicationSettings
    If iCountFiles > 0 Then Call MsgBox("Выгружено [" & iCountFiles & "] - файлов", vbInformation)
    If chbOpenPath.Value Then Call OpenPath(sPathParent)
    Unload Me
End Sub

Private Sub txtDir_DropButtonClick()
    Me.Hide
    txtDir.Value = PathDialogFun(txtDir.Value)
    Me.Show
End Sub

Private Sub txtDir_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call RestrictNavigationKeys(KeyCode, Shift)
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    txtDir.Value = ThisWorkbook.Path
    Call ConfigureDropButton(txtDir)
    Call SelectFileTypeByExtension(ActiveWorkbook.Name)
End Sub

Private Sub SelectFileTypeByExtension(ByVal FilePath As String)

    Dim ext         As String
    ext = LCase$(GetExtensionName(FilePath))

    Select Case ext
        Case "xlsm": optXLSM.Value = True
        Case "xlsb": optXLSB.Value = True
        Case "xls": optXLS.Value = True
        Case Else: optXLSX.Value = True
    End Select

End Sub

Private Sub SaveWBook(ByRef wb As Workbook, ByRef sPathFile As String)

    Select Case True
        Case optXLS.Value
            wb.SaveAs sPathFile & ".xls", fileFormat:=xlExcel8
        Case optXLSB.Value
            wb.SaveAs sPathFile & ".xlsb", fileFormat:=xlExcel12
        Case optXLSM.Value
            wb.SaveAs sPathFile & ".xlsm", fileFormat:=xlOpenXMLWorkbookMacroEnabled
        Case optXLSX.Value
            wb.SaveAs sPathFile & ".xlsx", xlOpenXMLWorkbook
        Case optPDF.Value
            On Error Resume Next
            wb.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:= _
                    sPathFile & ".pdf", Quality:=xlQualityStandard, _
                    IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
            On Error GoTo 0
    End Select
    wb.Close True
End Sub

