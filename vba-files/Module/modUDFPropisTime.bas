Attribute VB_Name = "modUDFPropisTime"
Option Explicit
Option Private Module
'***********************************************************************************************************
' Author         : VBATools
' Date           : 13.11.2019
' Обратная связь : info@VBATools.ru
' Copyright      : VBATools.ru
'***********************************************************************************************************
'

Public Function ВРЕМЯПРОПИСЬЮ(ByVal Время As Date, _
        Optional ByVal ФОРМАТ As Integer = 1, _
        Optional ByVal Тип_скобок As Integer = 0, _
        Optional ByVal Скобки_добавить As Boolean = True, _
        Optional ByVal РЕГИСТР As Integer = 0, _
        Optional ByVal Дублировать_число As Boolean = False) As String

    Dim L10(9)      As String   ' Десятки
    Dim L1(22)      As String   ' Единицы

    Dim h As Integer, m As Integer, s As Integer
    Dim LETTERS As String, LETTHOUR As String, LETTMINUTE As String, LETTSECOND As String
    Dim n10 As Integer, n1 As Integer
    Dim MyDate      As Date
    ' ДЕСЯТКИ
    L10(0) = ""
    L10(1) = ""
    L10(2) = "двадцать"
    L10(3) = "тридцать"
    L10(4) = "сорок"
    L10(5) = "пятьдесят"
    L10(6) = "шестьдесят"
    L10(7) = "семьдесят"
    L10(8) = "восемьдесят"
    L10(9) = "девяносто"
    ' ЕДИНИЦЫ
    L1(0) = "ноль"
    L1(1) = "один"
    L1(2) = "два"
    L1(3) = "три"
    L1(4) = "четыре"
    L1(5) = "пять"
    L1(6) = "шесть"
    L1(7) = "семь"
    L1(8) = "восемь"
    L1(9) = "девять"
    L1(10) = "десять"
    L1(11) = "одиннадцать"
    L1(12) = "двенадцать"
    L1(13) = "тринадцать"
    L1(14) = "четырнадцать"
    L1(15) = "пятнадцать"
    L1(16) = "шестнадцать"
    L1(17) = "семнадцать"
    L1(18) = "восемнадцать"
    L1(19) = "девятнадцать"
    L1(20) = "двадцать"
    L1(21) = "одна"
    L1(22) = "две"
    MyDate = Время
    h = Hour(MyDate)

    ' выделение десятков
    n10 = Fix(h / 10)
    ' выделение единиц
    n1 = h - n10 * 10

    If h <= 20 Then
        LETTHOUR = L1(h)
        If h = 1 Then
            LETTHOUR = LETTHOUR & " час"
        ElseIf h < 5 And h > 0 Then
            LETTHOUR = LETTHOUR & " часа"
        Else
            LETTHOUR = LETTHOUR & " часов"
        End If
    Else
        LETTHOUR = L10(n10) & " " & L1(n1)
        If n1 = 1 Then
            LETTHOUR = LETTHOUR & " час"
        ElseIf n1 < 5 Then
            LETTHOUR = LETTHOUR & " часа"
        Else
            LETTHOUR = LETTHOUR & " часов"
        End If
    End If

    m = Minute(MyDate)

    ' выделение десятков
    n10 = Fix(m / 10)
    ' выделение единиц
    n1 = m - n10 * 10

    If m <= 20 Then
        If m = 1 Or m = 2 Then
            LETTMINUTE = L1(m + 20)
        Else
            LETTMINUTE = L1(m)
        End If
        If m = 1 Then
            LETTMINUTE = LETTMINUTE & " минута"
        ElseIf m < 5 And m > 0 Then
            LETTMINUTE = LETTMINUTE & " минуты"
        Else
            LETTMINUTE = LETTMINUTE & " минут"
        End If
    Else
        If n1 = 1 Or n1 = 2 Then
            LETTMINUTE = L10(n10) & " " & L1(n1 + 20)
        Else
            LETTMINUTE = L10(n10) & IIf(n1 = 0, "", " " & L1(n1))
        End If
        If n1 = 1 Then
            LETTMINUTE = LETTMINUTE & " минута"
        ElseIf n1 > 0 And n1 < 5 Then
            LETTMINUTE = LETTMINUTE & " минуты"
        Else
            LETTMINUTE = LETTMINUTE & " минут"
        End If
    End If

    s = Second(MyDate)

    ' выделение десятков
    n10 = Fix(s / 10)
    ' выделение единиц
    n1 = s - n10 * 10

    If s <= 20 Then
        If s = 1 Or s = 2 Then
            LETTSECOND = L1(s + 20)
        Else
            LETTSECOND = L1(s)
        End If
        If s = 1 Then
            LETTSECOND = LETTSECOND & " секунда"
        ElseIf s < 5 And s > 0 Then
            LETTSECOND = LETTSECOND & " секунды"
        Else
            LETTSECOND = LETTSECOND & " секунд"
        End If
    Else
        If n1 = 1 Or n1 = 2 Then
            LETTSECOND = L10(n10) & " " & L1(n1 + 20)
        Else
            LETTSECOND = L10(n10) & IIf(n1 = 0, "", " " & L1(n1))
        End If
        If n1 = 1 Then
            LETTSECOND = LETTSECOND & " секунда"
        ElseIf n1 > 0 And n1 < 5 Then
            LETTSECOND = LETTSECOND & " секунды"
        Else
            LETTSECOND = LETTSECOND & " секунд"
        End If
    End If

    Select Case ФОРМАТ
        Case 3
            LETTERS = LETTHOUR
        Case 2
            LETTERS = LETTHOUR & " " & LETTMINUTE
        Case 1
            LETTERS = LETTHOUR & " " & LETTMINUTE & " " & LETTSECOND
    End Select

    If Скобки_добавить Then
        Select Case Тип_скобок
            Case 0
                If Дублировать_число Then
                    LETTERS = "(" & MyDate & " " & LETTERS & ")"
                Else
                    LETTERS = "(" & LETTERS & ")"
                End If
            Case 1
                LETTERS = MyDate & " (" & LETTERS & ")"
            Case 2
                LETTERS = " (" & MyDate & ") " & LETTERS
        End Select
    Else
        If Дублировать_число Then LETTERS = MyDate & " " & LETTERS
    End If

    ВРЕМЯПРОПИСЬЮ = AddRegistr(LETTERS, РЕГИСТР)
End Function

