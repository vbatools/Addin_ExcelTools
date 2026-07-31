Attribute VB_Name = "modUDFPropisSum"
Option Explicit
Option Private Module
'***********************************************************************************************************
' Author         : VBATools
' Date           : 13.11.2019
' Обратная связь : info@VBATools.ru
' Copyright      : VBATools.ru
'***********************************************************************************************************


Public Function СУММАПРОПИСЬЮ(ByVal ЧИСЛО As Currency, Optional Падеж As Byte = 1, _
        Optional ТИП_ДАННЫХ As Byte = 0, _
        Optional ДУБЛИРОВАТЬ_ЧИСЛО As Boolean = True, _
        Optional ДРОБЬ_ПРОПИСЬЮ As Boolean = False, _
        Optional ЗАГЛАВНАЯ As Byte = 0, _
        Optional СКОБКИ_МЕСТО As Byte = 0, _
        Optional ФОРМАТ As String = "0") As String

    Dim snNumeric   As Currency
    snNumeric = ЧИСЛО

    If Val(snNumeric) < -1E+15 Or Val(snNumeric) > 1E+15 Then
        MsgBox ("Число выходит за допустимые пределы [-1 000 000 000 000 000, 1 000 000 000 000 000]!")
        Exit Function
    End If

    Dim sNumeric    As String
    Dim sNewNum     As String
    Dim sNewNumDrob As String
    Dim iDrob       As Integer
    Dim iType       As Byte
    Dim bRodWomenCel As Boolean
    Dim bRodWomenDrob As Boolean
    Dim sNum        As String

    iDrob = 0
    sNum = VBA.format$(snNumeric, ФОРМАТ)

    Select Case ТИП_ДАННЫХ
        Case 1:
            sNum = VBA.format$(snNumeric, "#,##0.00")
            bRodWomenCel = False
        Case 2, 3:
            bRodWomenCel = False
            bRodWomenDrob = False
            iType = 4
            iDrob = (snNumeric - VBA.Fix(snNumeric)) * 100
            sNum = VBA.format$(snNumeric, "#,##0.00")
        Case 4, 5, 6:
            bRodWomenCel = False
        Case 7:
            bRodWomenCel = True
        Case 8:
            bRodWomenCel = True
            bRodWomenDrob = True
            iDrob = (snNumeric - VBA.Fix(snNumeric)) * 1000
            iType = 8
            sNum = VBA.format$(snNumeric, "0.000")
            If iDrob Mod 100 = 0 Then iDrob = iDrob / 100: iType = 4: sNum = VBA.format$(snNumeric, "0.0")
            If iDrob Mod 10 = 0 Then iDrob = iDrob / 10: iType = 6: sNum = VBA.format$(snNumeric, "0.00")
        Case 9:
            bRodWomenCel = False
            bRodWomenDrob = True
            iType = 4
            iDrob = (snNumeric - VBA.Fix(snNumeric)) * 100
            sNum = VBA.format$(snNumeric, "#,##0.00")
    End Select
    sNewNum = AddNumerString(VBA.Fix(snNumeric), Падеж, ТИП_ДАННЫХ, 1, bRodWomenCel)
    If iDrob > 0 Then
        sNewNum = sNewNum & AddNumerString(VBA.format$(iDrob, VBA.String$(15, "0")), Падеж, ТИП_ДАННЫХ, iType, bRodWomenDrob, ДРОБЬ_ПРОПИСЬЮ)
    End If

    If ДУБЛИРОВАТЬ_ЧИСЛО Then
        Select Case СКОБКИ_МЕСТО
            Case 1:
                sNumeric = "(" & sNum & ") " & sNewNum
            Case 2:
                sNumeric = sNum & " " & "(" & VBA.Left$(sNewNum, VBA.Len(sNewNum) - 1) & ")"
            Case Else:
                sNumeric = sNum & " " & sNewNum
        End Select
    Else
        sNumeric = sNewNum
    End If
    СУММАПРОПИСЬЮ = AddRegistr(WorksheetFunction.Trim(sNumeric), ЗАГЛАВНАЯ)
End Function

Private Function AddNumerString(ByVal lNumeric As Currency, _
        Optional byPadeg As Byte = 1, _
        Optional TypeData As Byte = 1, _
        Optional arrCol As Byte = 1, _
        Optional bRodWomen As Boolean = False, _
        Optional bNumDrob As Boolean = True) As String
    Dim sZero       As String
    Dim arrSot      As Variant
    Dim arrDes      As Variant
    Dim arrEd       As Variant
    Dim arrRazr     As Variant
    Dim arrNin      As Variant

    Dim arrType     As Variant


    Dim i           As Integer
    Dim sNewNum     As String
    Dim sNumeric    As String
    Dim sMinus      As String

    Select Case byPadeg
            'родительный
        Case 1:
            sZero = "ноля "
            arrEd = Array("", "одного ", "двух ", "трех ", "четырех ", "пяти ", "шести ", "семи ", "восьми ", "девяти ", "", "одной ", "двух ")
            arrNin = Array("десяти ", "одиннадцати ", "двенадцати ", "тринадцати ", "четырнадцати ", "пятнадцати ", "шестнадцати ", "семнадцати ", "восемнадцати ", "девятнадцати ")
            arrDes = Array("", "", "двадцати ", "тридцати ", "сорока ", "пятидесяти ", "шестидесяти ", "семидесяти ", "восьмидесяти ", "девяноста ")
            arrSot = Array("", "ста ", "двухсот ", "трехсот ", "четырехсот ", "пятисот ", "шестисот ", "семисот ", "восьмисот ", "девятисот ")
            arrRazr = Array("триллиона ", "триллионов ", "триллионов ", "миллиарда ", "миллиардов ", "миллиардов ", "миллиона ", "миллионов ", "миллионов ", "тысячи ", "тысяч ", "тысяч ", "", "", "")
            Select Case TypeData
                Case 1, 9:
                    arrType = Array("рубля ", "рублей ", "рублей ", "копейки ", "копеек ", "копеек ")
                Case 2:
                    arrType = Array("доллара ", "долларов ", "долларов ", "цента ", "центов ", "центов ")
                Case 3:
                    arrType = Array("евро ", "евро ", "евро ", "цента ", "центов ", "центов ")
                Case 4:
                    arrType = Array("календарного дня ", "календарных дней ", "календарных дней ")
                Case 5:
                    arrType = Array("рабочего дня ", "рабочих дней ", "рабочих дней ")
                Case 6:
                    arrType = Array("дня ", "дней ", "дней ")
                Case 7:
                    arrType = Array("штуки ", "штук ", "штук ")
                Case 8:
                    arrType = Array("целой ", "целых ", "целых ", "десятой ", "десятых ", "сотой ", "сотых ", "тысячной ", "тысячных ")
                Case Else:
                    arrType = Array("", "", "")
            End Select
            'дательный
        Case 2:
            sZero = "нолю "
            arrEd = Array("", "одному ", "двум ", "трем ", "четырем ", "пяти ", "шести ", "семи ", "восьми ", "девяти ", "", "одной ", "двум ")
            arrNin = Array("десяти ", "одиннадцати ", "двенадцати ", "тринадцати ", "четырнадцати ", "пятнадцати ", "шестнадцати ", "семнадцати ", "восемнадцати ", "девятнадцати ")
            arrDes = Array("", "", "двадцати ", "тридцати ", "сорока ", "пятидесяти ", "шестидесяти ", "семидесяти ", "восьмидесяти ", "девяноста ")
            arrSot = Array("", "ста ", "двухстам ", "тремстам ", "четыремстам ", "пятистам ", "шестистам ", "семистам ", "восьмистам ", "девятистам ")
            arrRazr = Array("триллиону ", "триллионам ", "триллионам ", "миллиарду ", "миллиардам ", "миллиардам ", "миллиону ", "миллионам ", "миллионам ", "тысяче ", "тысячам ", "тысячам ", "", "", "")
            Select Case TypeData
                Case 1, 9:
                    arrType = Array("рублю ", "рублям ", "рублям ", "копейке ", "копейкам ", "копейкам ")
                Case 2:
                    arrType = Array("доллару ", "долларам ", "долларам ", "центу ", "центам ", "центам ")
                Case 3:
                    arrType = Array("евро ", "евро ", "евро ", "центу ", "центам ", "центам ")
                Case 4:
                    arrType = Array("календарному дню ", "календарным дням ", "календарным дням ")
                Case 5:
                    arrType = Array("рабочему дню ", "рабочим дням ", "рабочим дням ")
                Case 6:
                    arrType = Array("дню ", "дням ", "дням ")
                Case 7:
                    arrType = Array("штуке ", "штукам ", "штукам ")
                Case 8:
                    arrType = Array("целой ", "целым ", "целым ", "десятой ", "десятым ", "сотой ", "сотым ", "тысячной ", "тысячным ")
                Case Else:
                    arrType = Array("", "", "")
            End Select
            'винительный
        Case 3:
            sZero = "ноль "
            arrEd = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ", "", "одну ", "две ")
            arrNin = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
            arrDes = Array("", "", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
            arrSot = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
            arrRazr = Array("триллион ", "триллиона ", "триллионов ", "миллиард ", "миллиарда ", "миллиардов ", "миллион ", "миллиона ", "миллионов ", "тысячу ", "тысячи ", "тысяч ", "", "", "")
            Select Case TypeData
                Case 1, 9:
                    arrType = Array("рубль ", "рублей ", "рубля ", "копейку ", "копеек ", "копейки ")
                Case 2:
                    arrType = Array("доллар ", "доллара ", "долларов ", "цент ", "центов ", "цента ")
                Case 3:
                    arrType = Array("евро ", "евро ", "евро ", "цент ", "цент ", "центов ", "цента ")
                Case 4:
                    arrType = Array("календарный день ", "календарных дней ", "календарных дня ")
                Case 5:
                    arrType = Array("рабочий день ", "рабочих дней ", "рабочих дня ")
                Case 6:
                    arrType = Array("день ", "дней ", "дня ")
                Case 7:
                    arrType = Array("штуку ", "штук ", "штуки ")
                Case 8:
                    arrType = Array("целую ", "целых ", "целых ", "десятую ", "десятых ", "сотую ", "сотых ", "тысячную ", "тысячных ")
                Case Else:
                    arrType = Array("", "", "")
            End Select
            'творительный
        Case 4:
            sZero = "нолем "
            arrEd = Array("", "одним ", "двумя ", "тремя ", "четырьмя ", "пятью ", "шестью ", "семью ", "восемью ", "девятью ", "", "одной ", "двумя ")
            arrNin = Array("десятью ", "одиннадцатью ", "двенадцатью ", "тринадцатью ", "четырнадцатью ", "пятнадцатью ", "шестнадцатью ", "семнадцатью ", "восемнадцатью ", "девятнадцатью ")
            arrDes = Array("", "", "двадцатью ", "тридцатью ", "сорока ", "пятьюдесятью ", "шестьюдесятью ", "семьюдесятью ", "восемьюдесятью ", "девяноста ")
            arrSot = Array("", "ста ", "двумястами ", "тремястами ", "четырьмястами ", "пятьюстами ", "шестьюстами ", "семьюстами ", "восьмьюстами ", "девятьюстами ")
            arrRazr = Array("триллионом ", "триллионами ", "триллионами ", "миллиардом ", "миллиардами ", "миллиардами ", "миллионом ", "миллионами ", "миллионами ", "тысячей ", "тысячами ", "тысячами ", "", "", "")
            Select Case TypeData
                Case 1, 9:
                    arrType = Array("рублем ", "рублями ", "рублями ", "копейкой ", "копейками ", "копейками ")
                Case 2:
                    arrType = Array("долларом ", "долларами ", "долларами ", "центом ", "центами ", "центами ")
                Case 3:
                    arrType = Array("евро ", "евро ", "евро ", "центом ", "центами ", "центами ")
                Case 4:
                    arrType = Array("календарным днем ", "календарными днями  ", "календарными днями  ")
                Case 5:
                    arrType = Array("рабочим днем ", "рабочими днями ", "рабочими днями ")
                Case 6:
                    arrType = Array("днем ", "днями ", "днями ")
                Case 7:
                    arrType = Array("штукой ", "штуками ", "штуками ")
                Case 8:
                    arrType = Array("целой ", "целыми ", "целыми ", "десятой ", "десятыми ", "сотой ", "сотыми ", "тысячной ", "тысячными ")
                Case Else:
                    arrType = Array("", "", "")
            End Select
            'предложный
        Case 5:
            sZero = "ноле "
            arrEd = Array("", "одном ", "двух ", "трех ", "четырех ", "пяти ", "шести ", "семи ", "восьми ", "девяти ", "", "одной ", "двух ")
            arrNin = Array("десяти ", "одиннадцати ", "двенадцати ", "тринадцати ", "четырнадцати ", "пятнадцати ", "шестнадцати ", "семнадцати ", "восемнадцати ", "девятнадцати ")
            arrDes = Array("", "", "двадцати ", "тридцати ", "сорока ", "пятидесяти ", "шестидесяти ", "семидесяти ", "восьмидесяти ", "девяноста ")
            arrSot = Array("", "ста ", "двухстах ", "трехстах ", "четырехстах ", "пятистах ", "шестистах ", "семистах ", "восьмистах ", "девятистах ")
            arrRazr = Array("триллионе ", "триллионах ", "триллионах ", "миллиарде ", "миллиардах ", "миллиардах ", "миллионе ", "миллионах ", "миллионах ", "тысяче ", "тысячах ", "тысячах ", "", "", "")
            Select Case TypeData
                Case 1, 9:
                    arrType = Array("рубле ", "рублях ", "рублях ", "копейке ", "копейках ", "копейках ")
                Case 2:
                    arrType = Array("долларе ", "долларах ", "долларах ", "центе ", "центах ", "центах ")
                Case 3:
                    arrType = Array("евро ", "евро ", "евро ", "центе ", "центах ", "центах ")
                Case 4:
                    arrType = Array("календарном дне ", "календарных днях ", "календарных днях ")
                Case 5:
                    arrType = Array("рабочем дне ", "рабочих днях ", "рабочих днях ")
                Case 6:
                    arrType = Array("дне ", "днях ", "днях ")
                Case 7:
                    arrType = Array("штуке ", "штуках ", "штуках ")
                Case 8:
                    arrType = Array("целой ", "целых ", "целых ", "десятой ", "десятых ", "сотой ", "сотых ", "тысячной ", "тысячных ")
                Case Else:
                    arrType = Array("", "", "")
            End Select
        Case Else:
            sZero = "ноль "
            arrEd = Array("", "один ", "два ", "три ", "четыре ", "пять ", "шесть ", "семь ", "восемь ", "девять ", "", "одна ", "две ")
            arrNin = Array("десять ", "одиннадцать ", "двенадцать ", "тринадцать ", "четырнадцать ", "пятнадцать ", "шестнадцать ", "семнадцать ", "восемнадцать ", "девятнадцать ")
            arrDes = Array("", "", "двадцать ", "тридцать ", "сорок ", "пятьдесят ", "шестьдесят ", "семьдесят ", "восемьдесят ", "девяносто ")
            arrSot = Array("", "сто ", "двести ", "триста ", "четыреста ", "пятьсот ", "шестьсот ", "семьсот ", "восемьсот ", "девятьсот ")
            arrRazr = Array("триллион ", "триллиона ", "триллионов ", "миллиард ", "миллиарда ", "миллиардов ", "миллион ", "миллиона ", "миллионов ", "тысяча ", "тысячи ", "тысяч ", "", "", "")
            Select Case TypeData
                Case 1, 9:
                    arrType = Array("рубль ", "рублей ", "рубля ", "копейка ", "копеек ", "копейки ")
                Case 2:
                    arrType = Array("доллар ", "долларов ", "доллара ", "цент ", "центов ", "цента ")
                Case 3:
                    arrType = Array("евро ", "евро ", "евро ", "цент ", "центов ", "цента ")
                Case 4:
                    arrType = Array("календарный день ", "календарных дней ", "календарных дня ")
                Case 5:
                    arrType = Array("рабочий день ", "рабочих дней ", "рабочих дня ")
                Case 6:
                    arrType = Array("день ", "дней ", "дня ")
                Case 7:
                    arrType = Array("штука ", "штук ", "штуки ")
                Case 8:
                    arrType = Array("целая ", "целых ", "целых ", "десятая ", "десятых ", "сотая ", "сотых ", "тысячная ", "тысячных ")
                Case Else:
                    arrType = Array("", "", "")
            End Select
    End Select

    If lNumeric < 0 Then
        sNumeric = VBA.format$(lNumeric * -1, VBA.String$(15, "0"))
        sMinus = "минус "
    Else
        sNumeric = VBA.format$(lNumeric, VBA.String$(15, "0"))
    End If

    If VBA.CDbl(sNumeric) = 0 Then sNewNum = sZero
    For i = 1 To VBA.Len(sNumeric) Step 3

        If Mid(sNumeric, i, 3) <> "000" Or i = Len(sNumeric) - 2 Then

            sNewNum = sNewNum & arrSot(CInt(Mid(sNumeric, i, 1)))

            If Mid(sNumeric, i + 1, 1) = "1" Then
                sNewNum = sNewNum & arrNin(CInt(Mid(sNumeric, i + 2, 1)))
            ElseIf bRodWomen And Val(Mid(sNumeric, i + 2, 1)) < 3 And i = Len(sNumeric) - 2 Then
                sNewNum = sNewNum & arrDes(CInt(Mid(sNumeric, i + 1, 1))) & arrEd(CInt(Mid(sNumeric, i + 2, 1)) + 10)
            ElseIf i = Len(sNumeric) - 5 And CInt(Mid(sNumeric, i + 2, 1)) < 3 Then
                sNewNum = sNewNum & arrDes(CInt(Mid(sNumeric, i + 1, 1))) & arrEd(CInt(Mid(sNumeric, i + 2, 1)) + 10)
            Else
                sNewNum = sNewNum & arrDes(CInt(Mid(sNumeric, i + 1, 1))) & arrEd(CInt(Mid(sNumeric, i + 2, 1)))
            End If

            If Mid(sNumeric, i + 1, 1) = "1" Or (Mid(sNumeric, i + 2, 1) + 9) Mod 10 >= 4 Then
                sNewNum = sNewNum & arrRazr(i + 1)
            ElseIf Mid(sNumeric, i + 2, 1) = "1" Then
                sNewNum = sNewNum & arrRazr(i - 1)
            Else
                sNewNum = sNewNum & arrRazr(i)
            End If
        End If

        If Mid(sNumeric, i + 2, 1) = "1" And i = Len(sNumeric) - 2 And (Not sNewNum Like "*цат[ьи] " And Not sNewNum Like "*цатью ") Then arrCol = arrCol - 1
        Select Case TypeData
            Case 1 To 7, 9:
                If i = Len(sNumeric) - 2 And Val(Mid(sNumeric, i + 2, 1)) >= 2 And Mid(sNumeric, i + 2, 1) <= 4 _
                        And (Not sNewNum Like "*цат[ьи] " And Not sNewNum Like "*цатью ") Then arrCol = arrCol + 1
        End Select
    Next i

    If Not bNumDrob Then
        sNewNum = VBA.CInt(lNumeric) & " "
    End If

    AddNumerString = sMinus & sNewNum & arrType(arrCol)

End Function

Public Function AddRegistr(ByVal txt As String, ByVal Registr As Integer) As String
    Dim txtPart()   As String
    Dim i           As Integer
    Dim n           As Integer
    If Registr Then
        Select Case Registr
            Case 0
                txt = LCase(txt)
            Case 1
                txt = UCase(txt)
            Case 2
                txtPart = Split(txt, " ")
                txt = ""
                Dim NumFlag As Boolean
                NumFlag = False
                For i = 0 To UBound(txtPart)
                    n = 1
                    If IsNumeric(txtPart(i)) = False And NumFlag = False Then
                        If Left(txtPart(i), 1) = "(" Then n = 2
                        txt = txt & UCase(Left(txtPart(i), n)) & LCase(Mid(txtPart(i), n + 1, Len(txtPart(i)) - n)) & " "
                        NumFlag = True
                    Else
                        txt = txt & txtPart(i) & " "
                    End If
                Next i
            Case 3
                txtPart = Split(Trim(txt), " ")
                txt = ""
                For i = 0 To UBound(txtPart)
                    n = 1
                    If Left(txtPart(i), 1) = "(" Then n = 2
                    If Len(txtPart(i)) > 2 Then
                        txt = txt & UCase(Left(txtPart(i), n)) & LCase(Mid(txtPart(i), n + 1, Len(txtPart(i)) - n)) & " "
                    Else
                        txt = txt & txtPart(i) & " "
                    End If
                Next i
        End Select
    End If
    AddRegistr = txt
End Function



