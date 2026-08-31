Attribute VB_Name = "modClearStyles"
Option Explicit
Option Private Module

'--------------------------------------------------------------------------------
' Sub: DeleteHiddenNames
' Purpose:  Удаляет все скрытые имена из активной книги
'   - Запрашивает подтверждение пользователя
'   - Перебирает все имена в книге
'   - Удаляет только скрытые имена (Visible = False)
'   - Выводит сообщение о количестве удаленных имен
'--------------------------------------------------------------------------------
Public Sub DeleteHiddenNames()
    If MsgBox("Вы хотите удалить скрытые имена в файле [" & ActiveWorkbook.Name & "]?" & vbNewLine & "Количество имен: " & ActiveWorkbook.Names.Count, vbYesNo + vbQuestion, "Удаление скрытых имен:") = vbNo Then
        Exit Sub
    End If
    Dim n           As Name
    Dim Count       As Integer
    On Error Resume Next
    For Each n In ActiveWorkbook.Names
        If Not n.Visible Then
            n.Delete
            Count = Count + 1
        End If
    Next n
    MsgBox "Скрытые имена в количестве " & Count & " удалены, из книги: [" & ActiveWorkbook.Name & "]"
End Sub
