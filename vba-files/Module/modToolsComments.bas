Attribute VB_Name = "modToolsComments"
Option Explicit
Option Private Module

Public Sub addSheetsComments()

    If ActiveWorkbook.ProtectStructure Then
        Call MsgBox("Включина защита структуры книги!", vbCritical)
        Exit Sub
    End If

    Dim ShKomm      As Worksheet
    Dim ShW         As Worksheet
    Dim i           As Long
    Dim Komm        As Comment
    Dim result      As Integer

    Call DisableApplicationSettings

    If HaveSheetInFile(ActiveWorkbook, "Комментарии") Then
        result = MsgBox("Старый лист с комментариями будет удален." _
                & vbCrLf & "Хотите продолжить?" _
                , vbYesNo Or vbQuestion Or vbDefaultButton2, "Список комментариев")
        If result = vbNo Then Exit Sub
        ActiveWorkbook.Sheets("Комментарии").Delete
    End If

    ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)

    Set ShKomm = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
    With ShKomm
        .Name = "Комментарии"

        .Columns("D:D").ColumnWidth = 21.5
        .Columns("E:E").ColumnWidth = 44.5
        .Columns("F:F").ColumnWidth = 11
        .Columns("B:B").ColumnWidth = 14.88
        With .Range("A1:F1")
            .Interior.Color = 2604537
            .Font.Bold = True
            .Font.Color = 16777215
        End With

        i = 1
        .Cells(i, 1) = "№"
        .Cells(i, 2) = "Лист"
        .Cells(i, 3) = "Адрес"
        .Cells(i, 4) = "Автор"
        .Cells(i, 5) = "Комментарий"
        .Cells(i, 6) = "Ссылка"

        i = i + 1
        For Each ShW In ActiveWorkbook.Worksheets
            If ShW.Name <> .Name Then
                For Each Komm In ShW.Comments
                    .Cells(i, 1) = i - 1
                    .Cells(i, 2) = ShW.Name
                    .Cells(i, 3) = Komm.Parent.Address
                    .Cells(i, 4) = Komm.Author
                    .Cells(i, 5) = Komm.Shape.AlternativeText
                    .Cells(i, 6).FormulaR1C1 = "=HYPERLINK(""[" & ActiveWorkbook.Name & "]'""&RC[-4]&""'!""&RC[-3],""Перейти"")"
                    i = i + 1
                Next
            End If
        Next ShW

        .Range("A1:F1").AutoFilter

        If i = 2 Then
            Call MsgBox("В текущей книге примечаний нет!", vbInformation, "Собщение:")
            ShKomm.Delete
        End If
    End With

    Call RestoreApplicationSettings
End Sub

Public Sub sizeTextComment()
    Dim sinSize     As Single
    sinSize = Application.InputBox(prompt:="Ведите размер шрифта", Type:=1)
    If sinSize <= 0 Then Exit Sub
    Dim oComm       As Comment
    For Each oComm In ActiveSheet.Comments
        oComm.Shape.TextFrame.Characters.Font.Size = sinSize
    Next oComm
End Sub

Public Sub showHidenComment(ByVal bShow As Boolean)
    Dim oComm       As Comment
    For Each oComm In ActiveSheet.Comments
        oComm.Visible = bShow
    Next oComm
End Sub




