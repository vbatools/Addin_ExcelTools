Attribute VB_Name = "modAddinConst"
Option Explicit
Option Private Module
' __      ______       _______         ,_
' \ \    / /  _ \   /\|__   __|        | |
'  \ \  / /| |_) | /  \  | | ___   ___ | |___
'   \ \/ / |  _ < / /\ \ | |/ _ \ / _ \| / __|
'    \  /  | |_) / ____ \| | (_) | (_) | \__ \
'     \/   |____/_/    \_\_|\___/ \___/|_|___/

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : modAddinConst - глобальные константы надстройки
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Copyright  : Apache License
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Const NAME_ADDIN As String = "EXCELTools"
Public Const FORMAT_DATE As String = "dd-mm-yyyy hh:mm:ss"

Public Enum enumParametrVersion
    enName = 1
    enAuthor
    enVersion
    enLicense
    enDateOfCreation
    enDateOfUpdate
    enDescription
    enAll
    [_First] = enName
    [_Last] = enAll
End Enum

Public Function Version(ByVal Parametr As enumParametrVersion, Optional bOnlyValue As Boolean = False) As String
    Dim sRes        As String
    Dim arr         As Variant
    arr = shLists.ListObjects("TB_ABOUT").DataBodyRange.Value2
    Select Case Parametr
        Case enumParametrVersion.enAll:
            Dim i   As Byte
            For i = enumParametrVersion.[_First] To enumParametrVersion.[_Last] - 1
                If sRes <> vbNullString Then sRes = sRes & vbNewLine
                If arr(i, 3) = 1 Then arr(i, 2) = VBA.format$(arr(i, 2), FORMAT_DATE)
                sRes = sRes & arr(i, 1) & ": " & arr(i, 2)
            Next i
        Case Else:
            If arr(Parametr, 3) = 1 Then arr(i, 2) = VBA.format$(arr(Parametr, 2), FORMAT_DATE)
            If bOnlyValue Then
                sRes = arr(Parametr, 2)
            Else
                sRes = arr(Parametr, 1) & ": " & arr(Parametr, 2)
            End If
    End Select
    Version = sRes
End Function



