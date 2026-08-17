VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAboutAddin 
   Caption         =   "О надстройке:"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8850.001
   OleObjectBlob   =   "frmAboutAddin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAboutAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit



Private Sub btnCancel_Click()
    Unload Me
End Sub


Private Sub Label1_Click()
    Call URLLinks("https://github.com/vbatools/Addin_ExcelTools")
End Sub

Private Sub UserForm_Initialize()
    ' Центрирование формы
    Call CenterUserForm(Me)
    lbAbout.Caption = Version(enAll)
End Sub
