VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCrossSelectionSettings 
   Caption         =   "Настройки перекрестного выделения:"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5805
   OleObjectBlob   =   "frmCrossSelectionSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCrossSelectionSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    On Error GoTo btnOK_Click_Err

    Select Case True
        Case OptBtn_RowColumn
            SaveSetting NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_Select", "RowColumn"
        Case OptBtn_Row
            SaveSetting NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_Select", "Row"
        Case OptBtn_Column
            SaveSetting NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_Select", "Column"
    End Select
    SaveSetting NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_ColorInterior", lbl_ColorInterior.BackColor
    SaveSetting NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_ColorFont", lbl_ColorFont.BackColor

    Unload Me
    Exit Sub
btnOK_Click_Err:
    Unload Me
    MsgBox Err.Description & vbCrLf & "в VBAProject.B1_CrossSelectionSet.btnOK_Click " & vbCrLf & "в строке " & Erl, vbExclamation + vbOKOnly, "Ошибка:"
End Sub

Private Sub lbl_ColorFont_Click()
    Dim lSeletedColor As Long
    lSeletedColor = GetColorFromDialog()
    If lSeletedColor = -1 Then Exit Sub
    lbl_ColorFont.BackColor = lSeletedColor
End Sub

Private Sub lbl_ColorInterior_Click()
    Dim lSeletedColor As Long
    lSeletedColor = GetColorFromDialog()
    If lSeletedColor = -1 Then Exit Sub
    lbl_ColorInterior.BackColor = lSeletedColor
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
End Sub

Private Sub UserForm_Activate()
    On Error GoTo UserForm_Activate_Err

    Dim CrossSelection_Select
    Dim CrossSelection_ColorInterior
    Dim CrossSelection_ColorFont

    CrossSelection_Select = GetSetting(NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_Select")
    CrossSelection_ColorInterior = GetSetting(NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_ColorInterior")
    CrossSelection_ColorFont = GetSetting(NAME_ADDIN, "AddIn" & NAME_ADDIN, "CrossSelection_ColorFont")

    If CrossSelection_Select = "" Then CrossSelection_Select = "RowColumn"
    If CrossSelection_ColorInterior = "" Then CrossSelection_ColorInterior = &HC0FFC0
    If CrossSelection_ColorFont = "" Then CrossSelection_ColorFont = &H0&

    Select Case CrossSelection_Select
        Case "RowColumn"
            OptBtn_RowColumn = True
        Case "Row"
            OptBtn_Row = True
        Case "Column"
            OptBtn_Column = True
    End Select

    lbl_ColorInterior.BackColor = CrossSelection_ColorInterior
    lbl_ColorFont.BackColor = CrossSelection_ColorFont

    Exit Sub
UserForm_Activate_Err:
    MsgBox Err.Description & vbCrLf & "в VBAProject.B1_CrossSelectionSet.UserForm_Activate " & vbCrLf & "в строке " & Erl, vbExclamation + vbOKOnly, "Ошибка:"
End Sub
