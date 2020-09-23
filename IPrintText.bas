Attribute VB_Name = "IPrintText"
Option Explicit 'Handles the font for the iPod
'The iPod uses the Mac font Chicago, which I was unable to find myself.
'All functions are tailor made to fit a bitmap of the font I made
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SpecialChars As String = "> <up> <down> <play> <pause> : - ( ) . [ ] { } \ &"
Private Const SpecialCharWidth As String = "7 8 8 10 7 3 6 4 4 3 4 4 5 5 5 9"
Public Sub TransBLT(srcHdc As Long, xSrc As Long, ySrc As Long, MaskHDC As Long, Xmsk As Long, Ymsk As Long, width As Long, Height As Long, DestHdc As Long, x As Long, y As Long)
    Const SRCPAINT = &HEE0086 'Assumes the mask matches the source's coordinates
    BitBlt DestHdc, x, y, width, Height, MaskHDC, Xmsk, Ymsk, SRCPAINT
    BitBlt DestHdc, x, y, width, Height, srcHdc, xSrc, ySrc, vbSrcAnd
End Sub
Public Function iPrint(text As String, srcHdc As Long, DestHdc As Long, x As Long, ByVal y As Long, Hi As Boolean)
    If InStr(text, vbNewLine) = 0 Then
        PrintLine text, srcHdc, DestHdc, x, y, Hi
    Else
        Dim temp As Long, tempstr() As String
        tempstr = Split(text, vbNewLine)
        For temp = 0 To UBound(tempstr)
            PrintLine tempstr(temp), srcHdc, DestHdc, x, y, Hi
            y = y + StringHeight(tempstr(temp))
        Next
    End If
End Function
Private Function PrintLine(ByVal text As String, srcHdc As Long, DestHdc As Long, ByVal x As Long, y As Long, Hi As Boolean)
    Dim tempstr As String
    Do Until Len(text) = 0
        tempstr = StripWord(text)
        DrawChar tempstr, DestHdc, srcHdc, x, y, Hi
        x = x + CharWidth(tempstr)
    Loop
End Function
Private Function DrawChar(letter As String, DestHdc As Long, srcHdc As Long, x As Long, y As Long, Highlite As Boolean) As Long
    Dim Xany As Long, Ymsk As Long, ySrc As Long, width As Long, Height As Long
    If Len(letter) = 0 Then letter = " "
    width = CharWidth(letter)
    If letter >= "a" And letter <= "z" Then
        SetLoc Xany, ySrc, Ymsk, Height, (Asc(letter) - 97) * 11 + 1, 1, 14, 12
    Else
        If letter >= "A" And letter <= "Z" Then
            SetLoc Xany, ySrc, Ymsk, Height, (Asc(letter) - 65) * 11 + 1, 27, 40, 12
        Else
            If letter >= "0" And letter <= "9" Then
                SetLoc Xany, ySrc, Ymsk, Height, (Asc(letter) - 48) * 11 + 1, 55, 68, 12
            Else
                letter = LCase(letter)
                Select Case letter
                    Case ">", "<up>", "<down>", "<play>", "<pause>", ":", "-", "(", ")", ".", "[", "]", "{", "}", "\", "&"
                        SetLoc Xany, ySrc, Ymsk, Height, 111 + (GetIndex(SpecialChars, letter) * 11), 55, 68, 12
                    Case "<b0>", "<b1>", "<b2>", "<b3>", "<b4>", "<b5>", "<b6>", "<b7>", "<b8>", "<b9>"
                        SetLoc Xany, ySrc, Ymsk, Height, (Asc(mid(letter, 3, 1)) - 48) * 19 + 1, 83, 111, 27
                    Case "<s0>", "<s1>", "<s2>", "<s3>", "<s4>", "<s5>", "<s6>", "<s7>", "<s8>", "<s9>"
                        SetLoc Xany, ySrc, Ymsk, Height, (Asc(mid(letter, 3, 1)) - 65) * 5 + 212, 83, 90, 6
                    Case " ":           SetLoc Xany, ySrc, Ymsk, Height, 275, 83, 96, 12
                    Case "<dir>":       SetLoc Xany, ySrc, Ymsk, Height, 275, 109, 122, 12
                    Case "<b:>":        SetLoc Xany, ySrc, Ymsk, Height, 191, 83, 111, 27
                    Case "<repeat>":    SetLoc Xany, ySrc, Ymsk, Height, 233, 115, 141, 7
                    Case "<shuffle>":   SetLoc Xany, ySrc, Ymsk, Height, 254, 115, 141, 7
                    Case "<sun>":       SetLoc Xany, ySrc, Ymsk, Height, 212, 99, 125, 7
                    Case "<mon>":       SetLoc Xany, ySrc, Ymsk, Height, 233, 99, 125, 7
                    Case "<tue>":       SetLoc Xany, ySrc, Ymsk, Height, 254, 99, 125, 7
                    Case "<wed>":       SetLoc Xany, ySrc, Ymsk, Height, 212, 107, 133, 7
                    Case "<thu>":       SetLoc Xany, ySrc, Ymsk, Height, 233, 107, 133, 7
                    Case "<fri>":       SetLoc Xany, ySrc, Ymsk, Height, 254, 107, 133, 7
                    Case "<sat>":       SetLoc Xany, ySrc, Ymsk, Height, 212, 115, 141, 7
                End Select
            End If
        End If
    End If
    
    If Not Highlite Then
        TransBLT srcHdc, Xany, ySrc, srcHdc, Xany, Ymsk, width, Height, DestHdc, x, y
    Else
        TransBLT srcHdc, Xany, Ymsk, srcHdc, Xany, ySrc, width, Height, DestHdc, x, y
    End If
    DrawChar = x + width - 1
End Function
Public Function CharHeight(letter As String) As Long
    'If (letter >= "a" And letter <= "z") Or (letter >= "A" And letter <= "Z") Or (letter >= "0" And letter <= "9") Then
    Select Case letter
        Case "<b0>", "<b1>", "<b2>", "<b3>", "<b4>", "<b5>", "<b6>", "<b7>", "<b8>", "<b9>", "<b:>": CharHeight = 27
        Case "<s0>", "<s1>", "<s2>", "<s3>", "<s4>", "<s5>", "<s6>", "<s7>", "<s8>", "<s9>": CharHeight = 6
        Case "<repeat>", "<shuffle>", "<sun>", "<mon>", "<tue>", "<wed>", "<thu>", "<fri>", "<sat>": CharHeight = 7
        Case Else: CharHeight = 12
    End Select
End Function

Public Function CharExists(letter As String) As Boolean
    CharExists = CharWidth(letter) > 0
End Function
Private Function CharWidth(ByVal letter As String) As Long
If Left(letter, 1) = "<" And Right(letter, 1) = ">" Then letter = LCase(letter)
Select Case letter
    'Lower Case
    Case "i", "l":                          CharWidth = 3
    Case "t":                               CharWidth = 5
    Case "c", "f", "k", "r", "s", "z":      CharWidth = 6
    Case "q":                               CharWidth = 8
    Case "m", "w":                          CharWidth = 11
    Case "a", "b", "d", "e", "g", "h", "j", "n", "o", "p", "u", "v", "x", "y": CharWidth = 7

    'Upper Case
    Case "E", "F", "L", "S", "Z":       CharWidth = 6
    Case "K", "N", "Q":                 CharWidth = 8
    Case "M", "W":                      CharWidth = 11
    Case "A", "B", "C", "D", "G", "H", "I", "J", "O", "P", "R", "T", "U", "V", "X", "Y": CharWidth = 7

    'Numbers and Special Characters
    Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9": CharWidth = 7 'Normal sized numbers
    Case "<b0>", "<b1>", "<b2>", "<b3>", "<b4>", "<b5>", "<b6>", "<b7>", "<b8>", "<b9>": CharWidth = 18  'Large numbers
    Case "<b:>": CharWidth = 8 'Big semicolon (:)
    Case "<s0>", "<s1>", "<s2>", "<s3>", "<s4>", "<s5>", "<s6>", "<s7>", "<s8>", "<s9>": CharWidth = 4 'Small numbers
    Case " ": CharWidth = 4 'Space
    Case "<dir>": CharWidth = 11
    Case ">", "<up>", "<down>", "<play>", "<pause>", ":", "-", "(", ")", ".", "[", "]", "{", "}", "\", "&": CharWidth = Val(GetFromIndex(SpecialChars, letter, SpecialCharWidth)) 'Special Chars (Punctuation)
    Case "<repeat>", "<shuffle>", "<sun>", "<mon>", "<tue>", "<wed>", "<thu>", "<fri>", "<sat>": CharWidth = 20 'Special Chars (Days of the week and play modes)
End Select
End Function
Private Sub SetLoc(Xany As Long, ySrc As Long, Ymsk As Long, Height As Long, x As Long, Y1 As Long, Y2 As Long, hit As Long)
    Xany = x
    ySrc = Y1
    Ymsk = Y2
    Height = hit
End Sub
Private Function GetIndex(text As String, word As String, Optional delimeter As String = " ") As Long
    Dim tempstr() As String, temp As Long
    GetIndex = -1
    tempstr = Split(text, delimeter)
    For temp = 0 To UBound(tempstr)
        If tempstr(temp) = word Then
            GetIndex = temp
            Exit For
        End If
    Next
End Function
Private Function GetFromIndex(text As String, word As String, text2 As String, Optional delimeter As String = " ") As String
    Dim temp As Long, tempstr() As String
    temp = GetIndex(text, word, delimeter)
    If temp > -1 Then
        tempstr = Split(text2, delimeter)
        GetFromIndex = tempstr(temp)
    End If
End Function
Private Function WordLength(text As String, start As Long) As Long
    Dim temp As Long, tempstr As String
    WordLength = 1
    If mid(text, start, 1) = "<" Then
        temp = InStr(start, text, ">")
        If temp > start Then
            tempstr = mid(text, start, temp - start + 1)
            If CharExists(tempstr) Then WordLength = Len(tempstr)
        End If
    End If
End Function
Private Function StripWord(ByRef text As String) As String
    Dim temp As Long
    temp = WordLength(text, 1)
    StripWord = Left(text, temp)
    text = Right(text, Len(text) - temp)
End Function
Public Function StringWidth(ByVal text As String) As Long
    Dim temp As Long
    Do Until Len(text) = 0
        temp = temp + CharWidth(StripWord(text))
    Loop
    StringWidth = temp
End Function
Public Function StringHeight(ByVal text As String) As Long
    Dim temp As Long, temp2 As Long
    Do Until Len(text) = 0
        temp2 = CharHeight(StripWord(text))
        If temp2 > temp Then temp = temp2
    Loop
    StringHeight = temp
End Function
Public Function GetTime() As String
    Dim temp As String, tempstr As String
    temp = time
    Do Until Left(temp, 1) = " "
        tempstr = tempstr & "<b" & Left(temp, 1) & ">"
        temp = Right(temp, Len(temp) - 1)
    Loop
    GetTime = tempstr
End Function
Public Function containsword(text As String, word As String) As Boolean
    containsword = InStr(1, text, word, vbTextCompare) > 0
End Function

Public Function GetNextChar(text As String, chars As String, Optional start As Long = 1) As Long
Dim temp As Long
temp = start
Do Until containsword(chars, mid(text, temp, 1)) Or temp > Len(text)
    temp = temp + 1
Loop
GetNextChar = temp
End Function
Public Function LineCount(ByVal text As String, width As Long) As Long
    Dim temp As Long
    Do Until Len(text) = 0
        WrapLine text, width
        temp = temp + 1
    Loop
    LineCount = temp
End Function
Public Function WrapLine(ByRef text As String, width As Long) As String
    Const SepChars As String = " ,." & vbNewLine
    Dim continue As Boolean, start As Long, finish As Long, tempstr As String, returnme As String
    start = 1
    Do Until continue
        finish = GetNextChar(text, SepChars, start + 1)
        tempstr = mid(text, start, finish - start)
        If Len(returnme) = 0 Then
            If StringWidth(tempstr) > width Then
                returnme = Truncate(tempstr, width, "-")
                continue = True
            Else
                returnme = tempstr
                start = start + Len(tempstr)
            End If
        Else
            If StringWidth(tempstr) + StringWidth(returnme) <= width Then
                returnme = returnme + tempstr
                start = start + Len(tempstr)
            Else
                continue = True
            End If
        End If
        If Len(tempstr) = 0 Then continue = True
    Loop
    text = Right(text, Len(text) - Len(returnme))
    WrapLine = returnme
End Function
