VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMengerShapesTools 
   Caption         =   "Менеджер фигур:"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6435
   OleObjectBlob   =   "frmMengerShapesTools.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMengerShapesTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit



Private Sub btnAlingHorizont_Click()
    Dim shp         As Shape
    Dim lCnt        As Long
    Dim dTop        As Double
    Dim dLeft       As Double
    Dim dHeight     As Double
    Dim dWidth      As Double
    Dim dSPACE      As Variant
    Dim lColCnt     As Variant
    Dim dStart      As Double
    Dim dMaxWidth   As Double

    If TypeName(Selection) = "Range" Then
        Call MsgBox("Пожалуйста, выберите фигуры перед запуском инструмента", vbCritical, "Ошибка:")
        Exit Sub
    End If

    lColCnt = VBA.Val(txtCountColumn.Value)
    dSPACE = VBA.Val(txtPaddingShapes.Value)

    If lColCnt = 0 Then lColCnt = 1
    lCnt = 1

    For Each shp In Selection.ShapeRange
        With shp
            'If first shape then store top position
            If lCnt = 1 Then
                dStart = .Top
            Else
                If lCnt Mod lColCnt = 1 Or lColCnt = 1 Then
                    'New column, move shape right
                    .Top = dStart
                    .Left = dLeft + dMaxWidth + dSPACE
                    dMaxWidth = .Width
                Else
                    'Same column, move shape down
                    .Top = dTop + dHeight + dSPACE
                    .Left = dLeft
                End If
            End If

            dTop = .Top
            dLeft = .Left
            dHeight = .Height
            dWidth = .Width
            dMaxWidth = WorksheetFunction.Max(dMaxWidth, .Width)
        End With
        lCnt = lCnt + 1
    Next shp
End Sub

Private Sub btnAlingVert_Click()
    Dim shp         As Shape
    Dim lCnt        As Long
    Dim dTop        As Double
    Dim dLeft       As Double
    Dim dHeight     As Double
    Dim dWidth      As Double
    Dim dSPACE      As Variant
    Dim lRowCnt     As Variant
    Dim dStart      As Double
    Dim dMaxHeight  As Double

    If TypeName(Selection) = "Range" Then
        Call MsgBox("Пожалуйста, выберите фигуры перед запуском инструмента", vbCritical, "Ошибка:")
        Exit Sub
    End If

    lRowCnt = VBA.Val(txtCountColumn.Value)
    dSPACE = VBA.Val(txtPaddingShapes.Value)

    If lRowCnt = 0 Then lRowCnt = 1
    lCnt = 1
    For Each shp In Selection.ShapeRange
        With shp
            If lCnt = 1 Then
                dStart = .Left
            Else
                If lCnt Mod lRowCnt = 1 Or lRowCnt = 1 Then
                    .Top = dTop + dMaxHeight + dSPACE
                    .Left = dStart
                    dMaxHeight = .Height
                Else
                    .Top = dTop
                    .Left = dLeft + dWidth + dSPACE
                End If
            End If

            dTop = .Top
            dLeft = .Left
            dHeight = .Height
            dWidth = .Width
            dMaxHeight = WorksheetFunction.Max(dMaxHeight, .Height)
        End With
        lCnt = lCnt + 1
    Next shp
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnText_Click()
    Dim oShape      As Shape
    On Error Resume Next
    For Each oShape In ActiveSheet.Shapes
        With oShape.TextFrame2
            .AutoSize = msoAutoSizeShapeToFitText
            .WordWrap = False
        End With
    Next
End Sub

Private Sub btnCell_Click()
    If TypeName(Selection) = "Range" Then
        Call MsgBox("Пожалуйста, выберите фигуры перед запуском инструмента", vbCritical, "Выровнять по вертикальной сетке:")
        Exit Sub
    End If

    Dim oShape      As Shape
    For Each oShape In Selection.ShapeRange
        Dim oCell   As Range: Set oCell = Cells(oShape.TopLeftCell.Row, oShape.TopLeftCell.Column)
        With oShape

            .LockAspectRatio = False
            If .Height > .Width Then
                .Width = oCell.Width
                .Height = oCell.Height
            Else
                .Height = oCell.Height
                .Width = oCell.Width
            End If

            .Top = oCell.MergeArea.Top + (oCell.MergeArea.Height - .Height) / 2
            .Left = oCell.MergeArea.Left + (oCell.MergeArea.Width - .Width) / 2
        End With
    Next
End Sub

Private Sub btnCopyRangeAsLink_Click()
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Выделите диапазон для копирования!", vbCritical, "Ошибка:")
        Exit Sub
    End If
    Dim coll        As New Collection
    Dim i           As Long

    For i = 1 To Selection.Areas.Count
        coll.Add Selection.Areas(i).Address
    Next
    Dim element     As Variant
    Range(coll(1)).Select
    For Each element In coll
        Application.CutCopyMode = False
        Range(element).Copy
        On Error Resume Next
        ActiveSheet.Pictures.Paste link:=True
        On Error GoTo 0
    Next
    Application.CutCopyMode = False
End Sub

Private Sub btnCopyRangeAsShape_Click()
    If TypeName(Selection) <> "Range" Then
        Call MsgBox("Выделите диапазон для копирования!", vbCritical, "Ошибка:")
        Exit Sub
    End If
    Dim i           As Long
    For i = 1 To Selection.Areas.Count
        Application.CutCopyMode = False
        Selection.Areas(i).Copy
        On Error Resume Next
        ActiveSheet.Pictures.Paste
        On Error GoTo 0
    Next
    Application.CutCopyMode = False
End Sub

Private Sub btnSelectAllShapes_Click()
    Dim shp         As Shape
    Dim i           As Long
    With ActiveSheet.Shapes
        If .Count - 1 < 0 Then Exit Sub
        ReDim arr(0 To .Count - 1) As String
        For Each shp In ActiveSheet.Shapes
            If shp.Adjustments.Parent.Type <> msoComment Then
                arr(i) = shp.Name
                i = i + 1
            End If
        Next shp
        .Range(arr).Select
    End With
End Sub

Private Sub txtCountColumn_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub txtPaddingShapes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Call ValidateNumericKey(KeyAscii)
End Sub

Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.98 * Application.Width - .Width
        .Top = Application.Top + 0.95 * Application.Height - .Height
    End With
End Sub
