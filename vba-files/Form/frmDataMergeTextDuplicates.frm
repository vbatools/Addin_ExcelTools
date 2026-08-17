VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDataMergeTextDuplicates 
   Caption         =   "Объединение ячеек с дублирующим текстом:"
   ClientHeight    =   2070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmDataMergeTextDuplicates.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDataMergeTextDuplicates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* UserForm     :   frmDataMergeText - объединение значений ячеек без потери данных
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   11-06-2026 10:14:25
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub btnCancel_Click()
    Unload Me
End Sub

'--------------------------------------------------------------------------------
' Procedure: btnOK_Click
' Purpose: Объединяет ячейки выделенного диапазона с сохранением текста
'          Текст объединяется через указанный разделитель
' Parameters: нет
'--------------------------------------------------------------------------------
Private Sub btnOK_Click()

    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Не выбран диапазон данных!", vbCritical)
        Exit Sub
    End If

    If ActiveSheet.ProtectContents Then
        Call MsgBox("Лист [" & ActiveSheet.Name & "] - защищен от изменений, снимите пароль!", vbCritical)
        Exit Sub
    End If

    On Error GoTo ErrorHandler

    Dim rng         As Range
    Dim arr         As Variant
    Dim i           As Long
    Dim iCount      As Long
    Dim j           As Long
    Dim jCount      As Long
    Dim k           As Long

    Set rng = Selection
    arr = rng.Value2

    If Not IsArray(arr) Then
        Call MsgBox("Выбрана одна ячейка!", vbCritical)
        Exit Sub
    End If

    iCount = UBound(arr, 1)
    jCount = UBound(arr, 2)

    Call DisableApplicationSettings
    Call SaveUndoInfo(rng, False, True)

    Dim sVal        As String
    
    If optRow.Value Then
        ' --- ОБЪЕДИНЕНИЕ ПО СТРОКАМ (вниз по столбцам) ---
        For j = 1 To jCount
            k = 1    ' Начальная строка группы
            sVal = arr(k, j)

            For i = 2 To iCount
                If arr(i, j) <> sVal Then
                    ' Значение изменилось, объединяем предыдущую группу
                    If i - 1 > k Then
                        Range(rng.Cells(k, j), rng.Cells(i - 1, j)).Merge
                    End If
                    k = i
                    sVal = arr(i, j)
                End If
            Next i

            ' Объединение последней группы в столбце после выхода из цикла
            If iCount > k Then
                Range(rng.Cells(k, j), rng.Cells(iCount, j)).Merge
            End If
        Next j
        
    Else
        ' --- ОБЪЕДИНЕНИЕ ПО СТОЛБЦАМ (вправо по строкам) ---
        For i = 1 To iCount
            k = 1    ' Начальный столбец группы
            sVal = arr(i, k)

            For j = 2 To jCount
                If arr(i, j) <> sVal Then
                    ' Значение изменилось, объединяем предыдущую группу
                    If j - 1 > k Then
                        Range(rng.Cells(i, k), rng.Cells(i, j - 1)).Merge
                    End If
                    k = j
                    sVal = arr(i, j)
                End If
            Next j

            ' Объединение последней группы в строке после выхода из цикла
            If jCount > k Then
                Range(rng.Cells(i, k), rng.Cells(i, jCount)).Merge
            End If
        Next i
    End If

    Call RestoreApplicationSettings
    Application.OnUndo "Отменить", "RestoreUndoInfo"
    Unload Me
    Exit Sub

ErrorHandler:
    Call RestoreApplicationSettings
    MsgBox "Ошибка: " & Err.Description, vbCritical, "Ошибка"
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
End Sub
