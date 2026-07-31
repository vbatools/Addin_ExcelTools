Attribute VB_Name = "moUDF"
Option Explicit
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module       :   modUserDefinedFunction - User-defined functions with dual naming (Cyrillic + English)
'* Author       :   VBATools
'* Copyright    :   Apache License
'* Created      :   23-06-2026 09:37:09
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


' ============================================================================
' SECTION 1: TEXT AND STRING PROCESSING
' ============================================================================

'--------------------------------------------------------------------------------
' Function: REPLACE_CHARS / «¿Ã≈Õ»“‹_—»Ã¬ŒÀ€
' Purpose: Performs character-by-character replacement in a string
' Parameters:
' TEXT_STR - Source string for processing
' CHARS_FIND - String of characters to search for
' CHARS_REPLACE - String of characters to replace with
' CASE_SENSITIVE - Case sensitivity (False by default)
' Returns: String - Transformed string
'--------------------------------------------------------------------------------
Public Function REPLACE_CHARS(ByVal TEXT_STR As String, ByVal CHARS_FIND As String, ByVal CHARS_REPLACE As String, Optional CASE_SENSITIVE As Boolean = False) As String
    Dim iLen        As Integer
    iLen = VBA.Len(CHARS_FIND)
    If iLen <> VBA.Len(CHARS_REPLACE) Then
        Select Case Application.International(xlCountrySetting)
            Case 7
                REPLACE_CHARS = "Œ¯Ë·Í‡: ÕÂ‡‚ÌÓÂ ÍÓÎË˜ÂÒÚ‚Ó ÒËÏ‚ÓÎÓ‚ ÔË ÔÓËÒÍÂ Ë Á‡ÏÂÌÂ"
            Case Else
                REPLACE_CHARS = "Error: Unequal number of characters in search and replace"
        End Select
        Exit Function
    End If
    Dim i           As Integer
    Dim sResult     As String
    If Not CASE_SENSITIVE Then
        CHARS_FIND = VBA.UCase(CHARS_FIND)
        CHARS_REPLACE = VBA.UCase(CHARS_REPLACE)
    End If

    sResult = TEXT_STR
    For i = 1 To iLen
        sResult = VBA.Replace(sResult, VBA.Mid(CHARS_FIND, i, 1), VBA.Mid(CHARS_REPLACE, i, 1))
        If Not CASE_SENSITIVE Then sResult = VBA.Replace(sResult, VBA.LCase$(VBA.Mid(CHARS_FIND, i, 1)), VBA.LCase$(VBA.Mid(CHARS_REPLACE, i, 1)))
    Next i
    REPLACE_CHARS = sResult
End Function

Public Function «¿Ã≈Õ»“‹_—»Ã¬ŒÀ€(ByVal —“–Œ ¿ As String, ByVal —»Ã¬ŒÀ€_Õ¿…“» As String, ByVal —»Ã¬ŒÀ€_«¿Ã≈Õ»“‹ As String, Optional ”◊»“¿“‹_–≈√»—“– As Boolean = False) As String
    «¿Ã≈Õ»“‹_—»Ã¬ŒÀ€ = REPLACE_CHARS(—“–Œ ¿, —»Ã¬ŒÀ€_Õ¿…“», —»Ã¬ŒÀ€_«¿Ã≈Õ»“‹, ”◊»“¿“‹_–≈√»—“–)
End Function

'--------------------------------------------------------------------------------
' Function: TEXT_LEFT / À≈¬Œ_“≈ —“
' Purpose: Extracts text to the left of the specified delimiter occurrence
' Parameters:
' TEXT_STR - Source string
' DELIMITER - Delimiter character
' DELIMITER_NUM - Ordinal number of delimiter (1 by default)
' COMPARE_MODE - Comparison mode (0 - binary, 1 - text)
' Returns: String - Text to the left of delimiter
'--------------------------------------------------------------------------------
Public Function TEXT_LEFT(ByVal TEXT_STR As String, ByVal Delimiter As String, Optional DELIMITER_NUM As Integer = 1, Optional COMPARE_MODE As Byte = 0) As String
    Dim i           As Integer
    Dim j           As Integer
    Dim k           As Integer
    Dim sResult     As String
    k = 1
    For i = 1 To DELIMITER_NUM
        j = VBA.InStr(k, TEXT_STR, Delimiter, COMPARE_MODE)
        If j <= 0 Then
            k = VBA.Len(TEXT_STR) + 2
            Exit For
        End If
        k = j + VBA.Len(Delimiter)
    Next i
    If k > 1 Then sResult = VBA.Left(TEXT_STR, k - VBA.Len(Delimiter) - 1)
    TEXT_LEFT = sResult
End Function

Public Function À≈¬Œ_“≈ —“(ByVal —“–Œ ¿ As String, ByVal –¿«ƒ≈À»“≈À‹ As String, Optional ÕŒÃ≈–_–¿«ƒ≈À»“≈Àﬂ As Integer = 1, Optional –≈∆»Ã_—–¿¬Õ≈Õ»ﬂ As Byte = 0) As String
    À≈¬Œ_“≈ —“ = TEXT_LEFT(—“–Œ ¿, –¿«ƒ≈À»“≈À‹, ÕŒÃ≈–_–¿«ƒ≈À»“≈Àﬂ, –≈∆»Ã_—–¿¬Õ≈Õ»ﬂ)
End Function

'--------------------------------------------------------------------------------
' Function: TEXT_BETWEEN / Ã≈∆ƒ”_“≈ —“
' Purpose: Extracts text between left and right delimiters
' Parameters:
' TEXT_STR - Source string
' LEFT_DELIM - Left boundary for extraction
' RIGHT_DELIM - Right boundary for extraction
' Returns: String - Text between delimiters
'--------------------------------------------------------------------------------
Public Function TEXT_BETWEEN(ByVal TEXT_STR As String, ByVal LEFT_DELIM As String, ByVal RIGHT_DELIM As String) As String
    Dim sResult     As String
    Dim i           As Integer
    sResult = TEXT_STR
    i = VBA.InStr(1, sResult, LEFT_DELIM)
    If i > 0 Then sResult = VBA.Right(sResult, VBA.Len(sResult) - i - VBA.Len(LEFT_DELIM) + 1)
    i = VBA.InStr(1, sResult, RIGHT_DELIM) - 1
    If i > 0 Then sResult = VBA.Left(sResult, i)
    TEXT_BETWEEN = sResult
End Function

Public Function Ã≈∆ƒ”_“≈ —“(ByVal —“–Œ ¿ As String, ByVal À≈¬€…_–¿«ƒ≈À»“≈À‹ As String, ByVal œ–¿¬€…_–¿«ƒ≈À»“≈À‹ As String) As String
    Ã≈∆ƒ”_“≈ —“ = TEXT_BETWEEN(—“–Œ ¿, À≈¬€…_–¿«ƒ≈À»“≈À‹, œ–¿¬€…_–¿«ƒ≈À»“≈À‹)
End Function

'--------------------------------------------------------------------------------
' Function: FIND_REPLACE / Õ¿…“»_«¿Ã≈Õ»“‹
' Purpose: Find and replace text in string
' Parameters:
' TEXT_STR - Source string
' FIND_STR - Text to find
' REPLACE_STR - Text to replace with
' REPLACE_COUNT - Number of replacements (-1 for all)
' Returns: String - Modified string
'--------------------------------------------------------------------------------
Public Function FIND_REPLACE(ByVal TEXT_STR As String, _
        ByVal FIND_STR As String, _
        ByVal REPLACE_STR As String, _
        Optional REPLACE_COUNT As Integer = -1) As String

    FIND_REPLACE = VBA.Replace(TEXT_STR, FIND_STR, REPLACE_STR, , REPLACE_COUNT)
End Function

Public Function Õ¿…“»_«¿Ã≈Õ»“‹(ByVal “≈ —“ As String, _
        ByVal Õ¿…“» As String, _
        ByVal «¿Ã≈Õ»“‹ As String, _
        Optional  ŒÀ»◊≈—“¬Œ_«¿Ã≈Õ As Integer = -1) As String

    Õ¿…“»_«¿Ã≈Õ»“‹ = FIND_REPLACE(“≈ —“, Õ¿…“», «¿Ã≈Õ»“‹,  ŒÀ»◊≈—“¬Œ_«¿Ã≈Õ)
End Function

'--------------------------------------------------------------------------------
' Function: TEXT_RIGHT / œ–¿¬Œ_“≈ —“
' Purpose: Extracts text to the right of the specified delimiter occurrence
'          (search is performed from right to left)
' Parameters:
' TEXT_STR - Source string
' DELIMITER - Delimiter character
' DELIMITER_NUM - Ordinal number of delimiter from right (1 by default)
' COMPARE_MODE - Comparison mode (0 - binary, 1 - text)
' Returns: String - Text to the right of delimiter
'--------------------------------------------------------------------------------
Public Function TEXT_RIGHT(ByVal TEXT_STR As String, ByVal Delimiter As String, Optional DELIMITER_NUM As Integer = 1, Optional COMPARE_MODE As Byte = 0) As String
    Dim i           As Integer
    Dim j           As Integer
    j = -1

    For i = 1 To DELIMITER_NUM
        j = VBA.InStrRev(TEXT_STR, Delimiter, j, COMPARE_MODE)
        If j = 0 Then Exit For
        j = j - 1
    Next i
    If j = 0 Then
        TEXT_RIGHT = TEXT_STR
    Else
        TEXT_RIGHT = VBA.Right(TEXT_STR, VBA.Len(TEXT_STR) - j - VBA.Len(Delimiter))
    End If
End Function

Public Function œ–¿¬Œ_“≈ —“(ByVal —“–Œ ¿ As String, ByVal –¿«ƒ≈À»“≈À‹ As String, Optional ÕŒÃ≈–_–¿«ƒ≈À»“≈Àﬂ As Integer = 1, Optional –≈∆»Ã_—–¿¬Õ≈Õ»ﬂ As Byte = 0) As String
    œ–¿¬Œ_“≈ —“ = TEXT_RIGHT(—“–Œ ¿, –¿«ƒ≈À»“≈À‹, ÕŒÃ≈–_–¿«ƒ≈À»“≈Àﬂ, –≈∆»Ã_—–¿¬Õ≈Õ»ﬂ)
End Function

'--------------------------------------------------------------------------------
' Function: SPLIT_STRING / –¿«¡»“‹_—“–Œ ”
' Purpose: Splits string by delimiter and returns specified element
' Parameters:
' TEXT_STR - Source string
' DELIMITER - Delimiter character (space by default)
' ELEMENT_NUM - Element number to return (1 by default)
' LIMIT - Maximum number of splits (-1 for all)
' Returns: String - Specified element from split array
'--------------------------------------------------------------------------------
Public Function SPLIT_STRING(ByVal TEXT_STR As String, _
        Optional Delimiter As String = " ", _
        Optional ELEMENT_NUM As Integer = 1, _
        Optional LIMIT As Integer = -1) As String

    SPLIT_STRING = VBA.Split(TEXT_STR, Delimiter, LIMIT)(ELEMENT_NUM - 1)
End Function

Public Function –¿«¡»“‹_—“–Œ ”(ByVal “≈ —“ As String, _
        Optional –¿«ƒ≈À»“≈À‹ As String = " ", _
        Optional ÕŒÃ≈–_›À≈Ã≈Õ“¿ As Integer = 1, _
        Optional À»Ã»“ As Integer = -1) As String

    –¿«¡»“‹_—“–Œ ” = SPLIT_STRING(“≈ —“, –¿«ƒ≈À»“≈À‹, ÕŒÃ≈–_›À≈Ã≈Õ“¿, À»Ã»“)
End Function

'--------------------------------------------------------------------------------
' Function: CONCAT_MULTI / —÷≈œ»“‹_Ã”À‹“»
' Purpose: Combines values from multiple ranges with specified delimiter
' Parameters:
' DELIMITER - Delimiter character between values
' RANGES - Array of ranges to combine (ParamArray)
' Returns: String - Combined string
'--------------------------------------------------------------------------------
Public Function CONCAT_MULTI(ByVal Delimiter As String, ParamArray RANGES() As Variant) As String
    Dim arr         As Variant
    Dim item        As Variant
    Dim sResult     As String

    For Each arr In RANGES
        For Each item In arr
            If sResult <> vbNullString Then sResult = sResult & Delimiter
            sResult = sResult & item
        Next item
    Next arr
    CONCAT_MULTI = sResult
End Function

Public Function —÷≈œ»“‹_Ã”À‹“»(ByVal –¿«ƒ≈À»“≈À‹ As String, ParamArray ƒ»¿œ¿«ŒÕ€() As Variant) As String
    —÷≈œ»“‹_Ã”À‹“» = CONCAT_MULTI(–¿«ƒ≈À»“≈À‹, ƒ»¿œ¿«ŒÕ€)
End Function

'--------------------------------------------------------------------------------
' Function: TEXT_MATCHES_PATTERN / “≈ —“_—ŒŒ“¬≈“—“¬”≈“_ÿ¿¡ÀŒÕ”
' Purpose: Checks if text matches a pattern (Like operator)
' Parameters:
' TEXT_STR - Source string
' PATTERN - Pattern to match (supports wildcards * ? # [])
' Returns: Boolean - True if text matches pattern
'--------------------------------------------------------------------------------
Public Function TEXT_MATCHES_PATTERN(ByVal TEXT_STR As String, ByVal PATTERN As String) As Boolean
    TEXT_MATCHES_PATTERN = TEXT_STR Like PATTERN
End Function

Public Function “≈ —“_—ŒŒ“¬≈“—“¬”≈“_ÿ¿¡ÀŒÕ”(ByVal “≈ —“ As String, ByVal ÿ¿¡ÀŒÕ As String) As Boolean
    “≈ —“_—ŒŒ“¬≈“—“¬”≈“_ÿ¿¡ÀŒÕ” = TEXT_MATCHES_PATTERN(“≈ —“, ÿ¿¡ÀŒÕ)
End Function

'--------------------------------------------------------------------------------
' Function: TRANSLIT
' Purpose: œÂÓ·‡ÁÛÂÚ ÍËËÎÎË˜ÂÒÍËÈ ÚÂÍÒÚ ‚ Î‡ÚËÌËˆÛ ÒÓ„Î‡ÒÌÓ ‚˚·‡ÌÌÓÏÛ ÒÚ‡Ì‰‡ÚÛ.
'          œÓ‰‰ÂÊË‚‡ÂÚ Ó·‡·ÓÚÍÛ ÓÍÓÌ˜‡ÌËÈ Ë ÍÓÌÚÂÍÒÚÓÁ‡‚ËÒËÏÛ˛ Á‡ÏÂÌÛ ÒËÏ‚ÓÎÓ‚.
' Parameters:
'   sText        - String. »ÒıÓ‰Ì˚È ÚÂÍÒÚ ‰Îˇ ÍÓÌ‚ÂÚ‡ˆËË.
'   iStandard    - Integer (Optional).  Ó‰ ÒÚ‡Ì‰‡Ú‡ Ú‡ÌÒÎËÚÂ‡ˆËË (ÔÓ ÛÏÓÎ˜‡ÌË˛ 0).
'                  0 - Œ·˘ÂÔËÌˇÚ˚È ÒÚ‡Ì‰‡Ú (Ò Á‡ÏÂÌÓÈ "˚È" -> "iy").
'                  1 - » ¿Œ (ICAO Doc 9303, Á‡„‡ÌÔ‡ÒÔÓÚ).
'                  2 - √Œ—“ 7.79-2000 (ISO 9, Ò ÍÓÌÚÂÍÒÚÌÓÈ Á‡ÏÂÌÓÈ "÷").
'                  3 - BGN/PCGN.
'                  4 - BGN/PCGN (Ò Á‡ÏÂÌÓÈ ÓÍÓÌ˜‡ÌËÈ "ËÈ"/"˚È" -> "y").
' Returns: String - —ÚÓÍ‡ Ò ÔÂÓ·‡ÁÓ‚‡ÌÌ˚Ï ÚÂÍÒÚÓÏ.
'--------------------------------------------------------------------------------
Public Function TRANSLIT(ByVal sText As String, Optional ByVal iStandard As Integer = 0) As String
    Dim sRusAlphabet As String
    Dim vEngMap As Variant
    Dim i As Long, iIndex As Integer
    Dim sCharIn As String, sCharOut As String
    Dim sResult As String
    Dim sNextChar As String
    Dim bIsUpper As Boolean
    
    ' ¡‡ÁÓ‚˚È ‡ÎÙ‡‚ËÚ
    sRusAlphabet = "‡·‚„‰Â∏ÊÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˜¯˘˙˚¸˝˛ˇ"
    
    ' »ÌËˆË‡ÎËÁ‡ˆËˇ Ï‡ÒÒË‚Ó‚ Á‡ÏÂÌ ‚ Á‡‚ËÒËÏÓÒÚË ÓÚ ÒÚ‡Ì‰‡Ú‡
    Select Case iStandard
        Case 0 ' Œ·˘ÂÔËÌˇÚ˚È
            vEngMap = Array("a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "i", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "", "y", "", "e", "u", "ya")
        Case 1 ' » ¿Œ (ICAO)
            vEngMap = Array("a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "i", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "shch", "ie", "y", "", "e", "iu", "ia")
        Case 2 ' √Œ—“ 7.79-2000 (ISO 9)
            vEngMap = Array("a", "b", "v", "g", "d", "e", "yo", "zh", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "x", "cz", "ch", "sh", "shh", """", "y'", "'", "e'", "yu", "ya")
        Case 3 ' BGN/PCGN
            vEngMap = Array("a", "b", "v", "g", "d", "e", "yo", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "shch", "", "y", "", "e", "yu", "ya")
        Case 4 ' BGN/PCGN Ò Á‡ÏÂÌÓÈ ÓÍÓÌ˜‡ÌËÈ
            vEngMap = Array("a", "b", "v", "g", "d", "e", "yo", "zh", "z", "i", "y", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "shch", "", "y", "", "e", "yu", "ya")
        Case Else
            ' œÓ ÛÏÓÎ˜‡ÌË˛ Standard 0
            vEngMap = Array("a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "i", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "sch", "", "y", "", "e", "u", "ya")
    End Select
    
    ' ŒÒÌÓ‚ÌÓÈ ˆËÍÎ Ó·‡·ÓÚÍË ÒËÏ‚ÓÎÓ‚
    For i = 1 To Len(sText)
        sCharIn = Mid(sText, i, 1)
        
        ' œÓËÒÍ ÒËÏ‚ÓÎ‡ ‚ ÛÒÒÍÓÏ ‡ÎÙ‡‚ËÚÂ (·ÂÁ Û˜ÂÚ‡ Â„ËÒÚ‡ ‰Îˇ ÔÓËÒÍ‡)
        iIndex = InStr(1, sRusAlphabet, sCharIn, vbTextCompare)
        
        If iIndex > 0 Then
            ' —ËÏ‚ÓÎ Ì‡È‰ÂÌ - ·ÂÂÏ ·‡ÁÓ‚Û˛ Á‡ÏÂÌÛ
            sCharOut = vEngMap(iIndex - 1)
            
            ' --- —ÔÂˆËÙË˜ÂÒÍ‡ˇ ÎÓ„ËÍ‡ ‰Îˇ —Ú‡Ì‰‡Ú‡ 2 (√Œ—“) ---
            If iStandard = 2 Then
                ' ¡ÛÍ‚‡ "ˆ" Á‡ÏÂÌˇÂÚÒˇ Ì‡ "c", ÂÒÎË Á‡ ÌÂÈ ÒÎÂ‰Û˛Ú e, i, y, j. »Ì‡˜Â "cz".
                If (LCase(sCharIn) = "ˆ") Then
                    If i < Len(sText) Then
                        sNextChar = LCase(Mid(sText, i + 1, 1))
                        If InStr(1, "ÂËÈ˛ˇei", sNextChar) > 0 Then
                            sCharOut = "c"
                        End If
                    End If
                End If
            End If
            
            ' ŒÔÂ‰ÂÎÂÌËÂ Â„ËÒÚ‡ ËÒıÓ‰ÌÓ„Ó ÒËÏ‚ÓÎ‡
            bIsUpper = (StrComp(sCharIn, UCase(sCharIn), vbBinaryCompare) = 0) And (StrComp(sCharIn, LCase(sCharIn), vbBinaryCompare) <> 0)
            
            If bIsUpper Then
                ' ‘ÓÏËÛÂÏ ‚ÂıÌËÈ Â„ËÒÚ ‰Îˇ Á‡ÏÂÌ˚ (“ÓÎ¸ÍÓ ÔÂ‚‡ˇ ·ÛÍ‚‡ Á‡„Î‡‚Ì‡ˇ)
                If Len(sCharOut) > 0 Then
                    sCharOut = UCase(Left(sCharOut, 1)) & LCase(Mid(sCharOut, 2))
                End If
            Else
                sCharOut = LCase(sCharOut)
            End If
            
        Else
            ' —ËÏ‚ÓÎ ÌÂ Ì‡È‰ÂÌ (Î‡ÚËÌËˆ‡, ˆËÙ˚, ÁÌ‡ÍË) - ÓÒÚ‡‚ÎˇÂÏ Í‡Í ÂÒÚ¸
            sCharOut = sCharIn
        End If
        
        sResult = sResult & sCharOut
    Next i
    
    ' --- œÓÒÚ-Ó·‡·ÓÚÍ‡ ÓÍÓÌ˜‡ÌËÈ (ÒÔÂˆËÙËÍ‡ Standard 0, 4) ---
    If iStandard = 0 Then
        ' «‡ÏÂÌ‡ "˚È" Ì‡ "iy"
        sResult = Replace(sResult, "yy ", "iy ", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "yi ", "iy ", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "yi.", "iy.", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "yi,", "iy,", 1, -1, vbTextCompare)
        
        ' œÓ‚ÂÍ‡ ÍÓÌˆ‡ ÒÚÓÍË
        If Right(sResult, 2) = "yi" Then sResult = Left(sResult, Len(sResult) - 2) & "iy"
        If Right(sResult, 2) = "YI" Then sResult = Left(sResult, Len(sResult) - 2) & "IY"
        
    ElseIf iStandard = 4 Then
        ' «‡ÏÂÌ‡ "ËÈ" Ë "˚È" Ì‡ "y"
        sResult = Replace(sResult, "iy ", "y ", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "iy.", "y.", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "iy,", "y,", 1, -1, vbTextCompare)
        
        sResult = Replace(sResult, "yy ", "y ", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "yy.", "y.", 1, -1, vbTextCompare)
        sResult = Replace(sResult, "yy,", "y,", 1, -1, vbTextCompare)
        
        ' œÓ‚ÂÍ‡ ÍÓÌˆ‡ ÒÚÓÍË
        If Right(sResult, 2) = "iy" Then sResult = Left(sResult, Len(sResult) - 2) & "y"
        If Right(sResult, 2) = "yy" Then sResult = Left(sResult, Len(sResult) - 2) & "y"
        
        If Right(sResult, 2) = "IY" Then sResult = Left(sResult, Len(sResult) - 2) & "Y"
        If Right(sResult, 2) = "YY" Then sResult = Left(sResult, Len(sResult) - 2) & "Y"
    End If
    
    TRANSLIT = sResult

End Function

Public Function “–¿Õ—À»“(ByVal “≈ —“ As String, Optional ByVal —“¿Õƒ¿–“ As Integer = 0) As String
    “–¿Õ—À»“ = TRANSLIT(“≈ —“, —“¿Õƒ¿–“)
End Function

'--------------------------------------------------------------------------------
' Function: REMOVE_CHARS / ”ƒ¿À»“‹_—»Ã¬ŒÀ€
' Purpose: Removes specified characters from a string
' Parameters:
' TEXT_STR - Source string
' CHARS_REMOVE - String with characters to remove
' CASE_SENSITIVE - Case sensitivity (False by default)
' Returns: String - Cleaned string
'--------------------------------------------------------------------------------
Public Function REMOVE_CHARS(ByVal TEXT_STR As String, ByVal CHARS_REMOVE As String, Optional CASE_SENSITIVE As Boolean = False) As String
    Dim i           As Integer
    Dim sResult     As String
    sResult = TEXT_STR
    If Not CASE_SENSITIVE Then CHARS_REMOVE = VBA.UCase(CHARS_REMOVE)
    For i = 1 To VBA.Len(CHARS_REMOVE)
        sResult = VBA.Replace(sResult, VBA.Mid(CHARS_REMOVE, i, 1), vbNullString)
        If Not CASE_SENSITIVE Then sResult = VBA.Replace(sResult, VBA.LCase$(VBA.Mid(CHARS_REMOVE, i, 1)), vbNullString)
    Next i
    REMOVE_CHARS = sResult
End Function

Public Function ”ƒ¿À»“‹_—»Ã¬ŒÀ€(ByVal —“–Œ ¿ As String, ByVal —»Ã¬ŒÀ€_”ƒ¿À»“‹ As String, Optional ”◊»“¿“‹_–≈√»—“– As Boolean = False) As String
    ”ƒ¿À»“‹_—»Ã¬ŒÀ€ = REMOVE_CHARS(—“–Œ ¿, —»Ã¬ŒÀ€_”ƒ¿À»“‹, ”◊»“¿“‹_–≈√»—“–)
End Function

' ============================================================================
' SECTION 2: DATA EXTRACTION FROM CELLS
' ============================================================================

'--------------------------------------------------------------------------------
' Function: GET_COMMENT / œŒÀ”◊ ŒÃÃ≈Õ“
' Purpose: Returns comment text from a cell
' Parameters:
' CELL - Range object (single cell)
' Returns: String - Comment text or empty string
'--------------------------------------------------------------------------------
Public Function GET_COMMENT(cell As Range) As String
    On Error Resume Next
    GET_COMMENT = cell.Comment.TEXT
End Function

Public Function œŒÀ”◊ ŒÃÃ≈Õ“(ByVal ﬂ◊≈… ¿ As Range) As String
    œŒÀ”◊ ŒÃÃ≈Õ“ = GET_COMMENT(ﬂ◊≈… ¿)
End Function

'--------------------------------------------------------------------------------
' Function: GET_TEXT / œŒÀ”◊“≈ —“
' Purpose: Extracts only text characters (letters) from cell
' Parameters:
' CELL - Range object (single cell)
' Returns: String - Text characters only
'--------------------------------------------------------------------------------
Public Function GET_TEXT(cell As Range) As String
    Dim LenStr      As Long
    For LenStr = 1 To Len(cell)
        Select Case Asc(VBA.Mid(cell, LenStr, 1))
            Case 65 To 90
                GET_TEXT = GET_TEXT & VBA.Mid(cell, LenStr, 1)
            Case 97 To 122
                GET_TEXT = GET_TEXT & VBA.Mid(cell, LenStr, 1)
            Case 192 To 255
                GET_TEXT = GET_TEXT & VBA.Mid(cell, LenStr, 1)
        End Select
    Next
End Function

Public Function œŒÀ”◊“≈ —“(ByVal ﬂ◊≈… ¿ As Range) As String
    œŒÀ”◊“≈ —“ = GET_TEXT(ﬂ◊≈… ¿)
End Function

'--------------------------------------------------------------------------------
' Function: GET_NUMBER / œŒÀ”◊◊»—ÀŒ
' Purpose: Extracts only numeric characters from cell
' Parameters:
' CELL - Range object (single cell)
' Returns: String - Numeric characters only
'--------------------------------------------------------------------------------
Public Function GET_NUMBER(cell As Range) As String
    Dim LenStr      As Long
    For LenStr = 1 To Len(cell)
        Select Case Asc(VBA.Mid(cell, LenStr, 1))
            Case 48 To 57
                GET_NUMBER = GET_NUMBER & VBA.Mid(cell, LenStr, 1)
        End Select
    Next
End Function

Public Function œŒÀ”◊◊»—ÀŒ(ByVal ﬂ◊≈… ¿ As Range) As String
    œŒÀ”◊◊»—ÀŒ = GET_NUMBER(ﬂ◊≈… ¿)
End Function

'--------------------------------------------------------------------------------
' Function: FORMULA_TEXT / “≈ —“‘Œ–Ã”À€
' Purpose: Returns formula from cell as text
' Parameters:
' CELL - Range object (single cell)
' Returns: String - Formula text
'--------------------------------------------------------------------------------
Public Function FORMULA_TEXT(cell As Range) As String
    FORMULA_TEXT = cell.formula
End Function

Public Function “≈ —“‘Œ–Ã”À€(ByVal ﬂ◊≈… ¿ As Range) As String
    “≈ —“‘Œ–Ã”À€ = FORMULA_TEXT(ﬂ◊≈… ¿)
End Function

' ============================================================================
' SECTION 3: FORMATTING OPERATIONS (FILL AND FONT COLOR)
' ============================================================================

'--------------------------------------------------------------------------------
' Function: SUM_BY_COLOR / —”ÃÃ«¿À»¬ ¿
' Purpose: Sums values in cells with specified fill color
' Parameters:
' RANGE_DATA - Range to sum
' COLOR_SAMPLE - Cell with sample fill color
' Returns: Double - Sum of matching cells
'--------------------------------------------------------------------------------
Public Function SUM_BY_COLOR(RANGE_DATA As Range, COLOR_SAMPLE As Range) As Double
    Dim sinSum      As Double
    Dim oRng        As Range

    Application.Volatile True
    sinSum = 0
    For Each oRng In RANGE_DATA
        If oRng.Interior.ColorIndex = COLOR_SAMPLE.Interior.ColorIndex Then
            sinSum = sinSum + oRng.Value
        End If
    Next oRng
    SUM_BY_COLOR = sinSum
End Function

Public Function —”ÃÃ«¿À»¬ ¿(ByVal ƒ»¿œ¿«ŒÕ As Range, ByVal œ–»Ã≈–_«¿À»¬ » As Range) As Double
    —”ÃÃ«¿À»¬ ¿ = SUM_BY_COLOR(ƒ»¿œ¿«ŒÕ, œ–»Ã≈–_«¿À»¬ »)
End Function

'--------------------------------------------------------------------------------
' Function: SUM_BY_FONT_COLOR / —”ÃÃÿ–»‘“
' Purpose: Sums values in cells with specified font color
' Parameters:
' RANGE_DATA - Range to sum
' COLOR_SAMPLE - Cell with sample font color
' Returns: Double - Sum of matching cells
'--------------------------------------------------------------------------------
Public Function SUM_BY_FONT_COLOR(RANGE_DATA As Range, COLOR_SAMPLE As Range) As Double
    Dim sinSum      As Double
    Dim oRng        As Range

    Application.Volatile True
    sinSum = 0
    For Each oRng In RANGE_DATA
        If oRng.Font.ColorIndex = COLOR_SAMPLE.Font.ColorIndex Then
            sinSum = sinSum + oRng.Value
        End If
    Next oRng
    SUM_BY_FONT_COLOR = sinSum
End Function

Public Function —”ÃÃÿ–»‘“(ByVal ƒ»¿œ¿«ŒÕ As Range, ByVal œ–»Ã≈–_ÿ–»‘“¿ As Range) As Double
    —”ÃÃÿ–»‘“ = SUM_BY_FONT_COLOR(ƒ»¿œ¿«ŒÕ, œ–»Ã≈–_ÿ–»‘“¿)
End Function

'--------------------------------------------------------------------------------
' Function: COUNT_BY_COLOR / —◊≈“«¿À»¬ ¿
' Purpose: Counts cells with specified fill color
' Parameters:
' RANGE_DATA - Range to count
' COLOR_SAMPLE - Cell with sample fill color
' Returns: Long - Count of matching cells
'--------------------------------------------------------------------------------
Public Function COUNT_BY_COLOR(RANGE_DATA As Range, COLOR_SAMPLE As Range) As Long
    Dim lResult     As Long
    Dim oRng        As Range

    Application.Volatile True
    lResult = 0
    For Each oRng In RANGE_DATA
        If oRng.Interior.ColorIndex = COLOR_SAMPLE.Interior.ColorIndex Then
            lResult = lResult + 1
        End If
    Next oRng
    COUNT_BY_COLOR = lResult
End Function

Public Function —◊≈“«¿À»¬ ¿(ByVal ƒ»¿œ¿«ŒÕ As Range, ByVal œ–»Ã≈–_«¿À»¬ » As Range) As Long
    —◊≈“«¿À»¬ ¿ = COUNT_BY_COLOR(ƒ»¿œ¿«ŒÕ, œ–»Ã≈–_«¿À»¬ »)
End Function

'--------------------------------------------------------------------------------
' Function: COUNT_BY_FONT_COLOR / —◊≈“ÿ–»‘“
' Purpose: Counts cells with specified font color
' Parameters:
' RANGE_DATA - Range to count
' COLOR_SAMPLE - Cell with sample font color
' Returns: Long - Count of matching cells
'--------------------------------------------------------------------------------
Public Function COUNT_BY_FONT_COLOR(RANGE_DATA As Range, COLOR_SAMPLE As Range) As Long
    Dim lResult     As Long
    Dim oRng        As Range

    Application.Volatile True
    lResult = 0
    For Each oRng In RANGE_DATA
        If oRng.Font.ColorIndex = COLOR_SAMPLE.Font.ColorIndex Then
            lResult = lResult + 1
        End If
    Next oRng
    COUNT_BY_FONT_COLOR = lResult
End Function

Public Function —◊≈“ÿ–»‘“(ByVal ƒ»¿œ¿«ŒÕ As Range, ByVal œ–»Ã≈–_ÿ–»‘“¿ As Range) As Long
    —◊≈“ÿ–»‘“ = COUNT_BY_FONT_COLOR(ƒ»¿œ¿«ŒÕ, œ–»Ã≈–_ÿ–»‘“¿)
End Function

' ============================================================================
' SECTION 4: INFORMATION FUNCTIONS (WORKBOOK, SHEET, USER)
' ============================================================================

'--------------------------------------------------------------------------------
' Function: WORKBOOK_NAME / »Ãﬂ Õ»√»
' Purpose: Returns active workbook name
' Returns: String - Workbook name
'--------------------------------------------------------------------------------
Public Function WORKBOOK_NAME() As String
    WORKBOOK_NAME = ActiveWorkbook.Name
End Function

Public Function »Ãﬂ Õ»√»() As String
    »Ãﬂ Õ»√» = WORKBOOK_NAME()
End Function

'--------------------------------------------------------------------------------
' Function: SHEET_NAME / »ÃﬂÀ»—“¿
' Purpose: Returns active sheet name
' Returns: String - Sheet name
'--------------------------------------------------------------------------------
Public Function SHEET_NAME() As String
    SHEET_NAME = ActiveSheet.Name
End Function

Public Function »ÃﬂÀ»—“¿() As String
    »ÃﬂÀ»—“¿ = SHEET_NAME()
End Function

'--------------------------------------------------------------------------------
' Function: USER_NAME / »ÃﬂœŒÀ‹«Œ¬¿“≈Àﬂ
' Purpose: Returns current Windows user name
' Returns: String - User name
'--------------------------------------------------------------------------------
Public Function USER_NAME() As String
    USER_NAME = Environ("UserName")
End Function

Public Function »ÃﬂœŒÀ‹«Œ¬¿“≈Àﬂ() As String
    »ÃﬂœŒÀ‹«Œ¬¿“≈Àﬂ = USER_NAME()
End Function

'--------------------------------------------------------------------------------
' Function: WORKBOOK_FULL_PATH / œŒÀÕ€…œ”“‹ Õ»√»
' Purpose: Returns full path to active workbook
' Returns: String - Full path
'--------------------------------------------------------------------------------
Public Function WORKBOOK_FULL_PATH() As String
    WORKBOOK_FULL_PATH = ActiveWorkbook.FullName
End Function

Public Function œŒÀÕ€…œ”“‹ Õ»√»() As String
    œŒÀÕ€…œ”“‹ Õ»√» = WORKBOOK_FULL_PATH()
End Function

' ============================================================================
' SECTION 5: DATA VALIDATION AND ANALYSIS
' ============================================================================

'------------------------------------------------------------------------------
' Function: hasIs
' Purpose:  Base helper function for pattern matching (case-insensitive)
' Parameters:
'   sText   - String to check
'   sMaska  - Pattern mask for Like operator
' Returns: Boolean - True if pattern matches
'------------------------------------------------------------------------------
Public Function hasIs(sText As String, ByVal sMaska As String) As Boolean
    hasIs = UCase(sText) Like sMaska
End Function

'------------------------------------------------------------------------------
' Function: HAS_LATIN / ≈À¿“»Õ
' Purpose:  Checks if string contains Latin characters
' Parameters:
'   CELL - String to check
' Returns:  Boolean - True if Latin characters found
'------------------------------------------------------------------------------
Public Function HAS_LATIN(cell As String) As Boolean
    Const MASKA     As String = "*[ABCDEFGHIJKLMNOPQRSTUVWXYZ]*"
    HAS_LATIN = hasIs(cell, MASKA)
End Function

Public Function ≈À¿“»Õ(ByVal ﬂ◊≈… ¿ As String) As Boolean
    ≈À¿“»Õ = HAS_LATIN(ﬂ◊≈… ¿)
End Function

'------------------------------------------------------------------------------
' Function: HAS_CYRILLIC / ≈ »–»ÀÀ
' Purpose:  Checks if string contains Cyrillic characters
' Parameters:
'   CELL - String to check
' Returns:  Boolean - True if Cyrillic characters found
'------------------------------------------------------------------------------
Public Function HAS_CYRILLIC(cell As String) As Boolean
    Const MASKA     As String = "*[¿¡¬√ƒ≈®∆«»… ÀÃÕŒœ–—“”‘’÷◊ÿŸ⁄€‹›ﬁﬂ]*"
    HAS_CYRILLIC = hasIs(cell, MASKA)
End Function

Public Function ≈ »–»ÀÀ(ByVal ﬂ◊≈… ¿ As String) As Boolean
    ≈ »–»ÀÀ = HAS_CYRILLIC(ﬂ◊≈… ¿)
End Function

' ============================================================================
' SECTION 6: QR CODE GENERATION
' Note: Requires QRCodegen library reference
' ============================================================================

'--------------------------------------------------------------------------------
' Function: CREATE_QR / —Œ«ƒ¿“‹_QR
' Purpose: Generates QR code image in cell and returns source text
' Dependencies: QRCodegen library (QRCodegenEcc enum, QRCodegenBarcode function)
' Parameters:
' TEXT_STR - Text to encode in QR code
' QR_COLOR - QR code color (black by default)
' QR_SIZE - QR code size in pixels (200 by default)
' QR_TYPE - QR code type flag
' QR_ERROR - Error correction level (QRCodegenEcc_LOW by default)
' Returns: String - Source text (QR code displayed as image in cell)
'--------------------------------------------------------------------------------
Public Function CREATE_QR(ByVal TEXT_STR As String, _
        Optional QR_COLOR As OLE_COLOR = vbBlack, _
        Optional QR_SIZE As Integer = 200, _
        Optional QR_TYPE As Boolean, _
        Optional QR_ERROR As QRCodegenEcc = QRCodegenEcc_LOW) As String

    Dim MyCell      As Range
    Set MyCell = Application.Caller
    Dim sPath       As String
    sPath = ThisWorkbook.Path & Application.PathSeparator & "QR.emf"
    Call SavePicture(QRCodegenBarcode(TEXT_STR, QR_COLOR, 120, QR_TYPE, QR_ERROR, VERSION_MIN, VERSION_MAX, QRCodegenMask_AUTO, True), sPath)
    On Error Resume Next
    ActiveSheet.Pictures("My_QR_" & MyCell.Address(False, False)).Delete
    On Error GoTo 0
    Dim objPict     As Shape
    With MyCell
        Set objPict = .Parent.Shapes.AddPicture(sPath, msoFalse, msoTrue, .Left, .Top, QR_SIZE, QR_SIZE)
    End With
    objPict.Name = "My_QR_" & MyCell.Address(False, False)
    If Not objPict.Name Like "My_QR_*" Then objPict.Delete

    Call Kill(sPath)
    CREATE_QR = TEXT_STR
End Function

Public Function —Œ«ƒ¿“‹_QR(ByVal “≈ —“ As String, _
        Optional ÷¬≈“_QR As OLE_COLOR = vbBlack, _
        Optional –¿«Ã≈–_QR As Integer = 200, _
        Optional “»œ_QR As Boolean, _
        Optional Œÿ»¡ ¿_QR As QRCodegenEcc = QRCodegenEcc_LOW) As String

    —Œ«ƒ¿“‹_QR = CREATE_QR(“≈ —“, ÷¬≈“_QR, –¿«Ã≈–_QR, “»œ_QR, Œÿ»¡ ¿_QR)
End Function



