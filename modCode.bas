Attribute VB_Name = "modCode"
Option Explicit

'(stolen from a different prog i made)

'Code Parsing Helpers - By Lewis Miller (aka Deth)
'some functions to make code parsing life easier
'****************************************

'used by the tabtrim function
Enum TrimType
    Trim_Normal = 0
    Trim_Left = 1
    Trim_Right = 2
End Enum

'looks at a string to see if it contains a word in wordlist...
'wordlist is a comma delimited list like so: "word here,next word,word,word up,yo"
Function ContainsWord(ByVal WordList As String, ByVal strCheck As String) As Boolean

  Dim Arr() As String, X As Long
    strCheck = Trim$(strCheck)
    
    'remove any line numbers
    strCheck = RemoveLineNumbers(strCheck)
    
    If (Len(WordList) > 0) And (Len(strCheck) > 0) Then
        If InStr(WordList, ",") Then
            Arr = Split(WordList, ",")
          Else
            ReDim Arr(0)
            Arr(0) = WordList
        End If
        For X = 0 To UBound(Arr)
            If InStr(1, strCheck, Arr(X), vbTextCompare) Then
                ContainsWord = True
                Exit Function
            End If
        Next X
    End If

End Function

'looks at the begining of a string to see if it starts with a word in wordlist...
'wordlist is a comma delimited list like so: "word here,next word,word,word up,yo"
Function StartWord(ByVal WordList As String, ByVal strCheck As String) As Boolean

  Dim Arr() As String, X As Long

    strCheck = Trim$(strCheck)
    
    'remove any line numbers
    strCheck = RemoveLineNumbers(strCheck)
    
    If (Len(WordList) > 0) And (Len(strCheck) > 0) Then
        If InStr(WordList, ",") Then
            Arr = Split(WordList, ",")
          Else
            ReDim Arr(0)
            Arr(0) = WordList
        End If
        For X = 0 To UBound(Arr)
            If LeftCheck(Arr(X), strCheck) Then
                StartWord = True
                Exit Function
            End If
        Next X
    End If

End Function

'looks at the end of a string to see if it ends with a word in wordlist...
'wordlist is a comma delimited list like so: "word here,next word,word,word up,yo"
Function EndWord(ByVal WordList As String, ByVal strCheck As String) As Boolean

  Dim Arr() As String, X As Long

    strCheck = Trim$(strCheck)
    
    If (Len(WordList) > 0) And (Len(strCheck) > 0) Then
        If InStr(",", WordList) Then
            Arr = Split(WordList, ",")
          Else
            ReDim Arr(0)
            Arr(0) = WordList
        End If
        For X = 0 To UBound(Arr)
            If RightCheck(Arr(X), strCheck) Then
                EndWord = True
                Exit Function
            End If
        Next X
    End If

End Function

'checks to see if a string starts with a word
Function LeftCheck(ByVal strItem As String, ByVal strCheck As String) As Boolean

    strCheck = TabTrim(strCheck)
    strCheck = RemoveLineNumbers(strCheck)
    If Len(strCheck) >= Len(strItem) Then
        If StrComp(Left$(strCheck, Len(strItem)), strItem, 1) = 0 Then
            LeftCheck = True
        End If
    End If

End Function

'checks to see if a string ends with a word
Function RightCheck(ByVal strItem As String, ByVal strCheck As String) As Boolean

    strCheck = TabTrim(strCheck)
    If Len(strCheck) >= Len(strItem) Then
        If StrComp(Right$(strCheck, Len(strItem)), strItem, 1) = 0 Then
            RightCheck = True
        End If
    End If

End Function

'removes line numbers from the beginning of a line of code
Function RemoveLineNumbers(ByVal strItem As String) As String
    
    Dim Place As Long
    
    'remove any line numbers
    Place = InStr(strItem, " ")
    If Place > 1 Then
      If IsNumeric(Left$(strItem, Place - 1)) Then
          strItem = Trim$(Mid$(strItem, Place + 1))
      End If
    End If
    
    RemoveLineNumbers = strItem

End Function

'gets the function name from a line of code, or api call
'depends on the isapifunction() function below
Function ParseFunctionName(ByVal strLine As String) As String

  Dim Place As Long

    Place = IsAPIFunction(strLine)
    If Place > 0 Then
        strLine = Trim$(Left$(strLine, Place - 1))
        If InStr(strLine, " ") Then
            ParseFunctionName = Trim$(Mid$(strLine, InStrRev(strLine, " ") + 1))
        End If
      Else
        Place = InStr(strLine, "(")
        If Place Then
            strLine = Trim$(Left$(strLine, Place - 1))
            If InStr(strLine, " ") Then
                ParseFunctionName = Trim$(Mid$(strLine, InStrRev(strLine, " ") + 1))
            End If
        End If
    End If

End Function

'checks a line of code to see whether its an api call or regular function,
'returns the end position of api function name
Function IsAPIFunction(ByVal strLine As String) As Long

    If LenB(strLine) > 0 Then
        IsAPIFunction = InStr(strLine, "Declare ")
        If IsAPIFunction Then
            IsAPIFunction = InStr(strLine, " Lib " & Chr$(34))
        End If
    End If

End Function

'changes intrinsic String() variant functions to String$() functions
Function FixVariants(ByVal StrFix As String) As String

    StrFix = Trim$(StrFix)
    If Len(StrFix) > 2 Then
        StrFix = Replace$(StrFix, " Chr(", " Chr$(")
        StrFix = Replace$(StrFix, " Dir(", " Dir$(")
        StrFix = Replace$(StrFix, " Oct(", " Oct$(")
        StrFix = Replace$(StrFix, " Str(", " Str$(")
        StrFix = Replace$(StrFix, " Mid(", " Mid$(")
        StrFix = Replace$(StrFix, " Hex(", " Hex$(")
        StrFix = Replace$(StrFix, " Trim(", " Trim$(")
        StrFix = Replace$(StrFix, " Left(", " Left$(")
        StrFix = Replace$(StrFix, " Space(", " Space$(")
        StrFix = Replace$(StrFix, " LTrim(", " LTrim$(")
        StrFix = Replace$(StrFix, " RTrim(", " RTrim$(")
        StrFix = Replace$(StrFix, " UCase(", " UCase$(")
        StrFix = Replace$(StrFix, " LCase(", " LCase$(")
        StrFix = Replace$(StrFix, " Right(", " Right$(")
        StrFix = Replace$(StrFix, " String(", " String$(")
        StrFix = Replace$(StrFix, " Format(", " Format$(")
        StrFix = Replace$(StrFix, " Environ(", " Environ$(")
        StrFix = Replace$(StrFix, " Replace(", " Replace$(")
        StrFix = Replace$(StrFix, " InputBox(", " InputBox$(")
        StrFix = Replace$(StrFix, " StrReverse(", " StrReverse$(")

        'look for the funky
        If EndWord(" Dir, Dir()", StrFix) Then
            StrFix = Left$(StrFix, InStrRev(StrFix, "Dir") - 1) & "Dir$"
        End If
        If EndWord(" CurDir, CurDir()", StrFix) Then
            StrFix = Left$(StrFix, InStrRev(StrFix, "CurDir") - 1) & "CurDir$"
        End If
        If EndWord(" Environ, Environ()", StrFix) Then
            StrFix = Left$(StrFix, InStrRev(StrFix, "Environ") - 1) & "Environ$"
        End If
        If EndWord(" Time, Time()", StrFix) Then
            StrFix = Left$(StrFix, InStrRev(StrFix, "Time") - 1) & "Time$"
        End If
        If EndWord(" Command, Command()", StrFix) Then
            StrFix = Left$(StrFix, InStrRev(StrFix, "Command") - 1) & "Command$"
        End If

        StrFix = Replace$(StrFix, "If Dir ", "If Dir$ ")
        StrFix = Replace$(StrFix, "If CurDir ", "If CurDir$ ")
        StrFix = Replace$(StrFix, "If Environ ", "If Environ$ ")
        StrFix = Replace$(StrFix, "If Time ", "If Time$ ")
        StrFix = Replace$(StrFix, "If Command ", "If Command$ ")

        FixVariants = StrFix
    End If

End Function

'checks a character to see if it is a letter of the english alphabet
Function isAlpha(ByVal strChar As String) As Boolean

  Dim intChar As Long

    If Len(strChar) > 0 Then
        intChar = Asc(UCase$(strChar))
        If intChar > 64 And intChar < 91 Then
            isAlpha = True
        End If
    End If

End Function

'checks a character to see if it is a number or letter of the english alphabet
Function isAlphaNumeric(ByVal strChar As String) As Boolean

  Dim intChar As Long

    If Len(strChar) > 0 Then
        intChar = Asc(UCase$(strChar))
        If ((intChar > 64) And (intChar < 91)) Or ((intChar > 47) And (intChar < 58)) Then
            isAlphaNumeric = True
        End If
    End If

End Function

'checks a line of code to see if its a single or multiple line If...Then statement
Function isSingleLineIf(ByVal strCheck As String) As Boolean

  Dim strLength As Long

    strCheck = Trim$(strCheck)
    strCheck = RemoveLineNumbers(strCheck)
    strLength = Len(strCheck)

    If strLength > 0 Then
        If InStr(InStr(strCheck, " Then "), strCheck, " Else ") Then
            isSingleLineIf = True
          Else
            If InStrRev(strCheck, " Then ") Then
                strCheck = Mid$(strCheck, InStrRev(strCheck, " Then ") + 6)
                strLength = Len(strCheck)
                If strLength > 0 Then
                    If isAlpha(strCheck) Then
                        isSingleLineIf = True
                      Else
                        If Not StartWord("' ,rem ", strCheck) Then
                            isSingleLineIf = True
                        End If
                    End If
                End If
            End If
        End If
    End If

End Function

'not used in this program
'Function FindEndOfLine(Arr As Code_List, ByVal StartIndex As Long, ByVal MaxIndex As Long) As Long

'    If StartIndex <= MaxIndex Then
'        Do While RightCheck(" _", Arr.Lines(StartIndex).Text)
'            If (StartIndex >= MaxIndex) Then Exit Do
'            StartIndex = StartIndex + 1
'            'Arr.Lines(StartIndex) = Space$(48) & TabTrim(Trim$(Arr.Lines(StartIndex)))
'        Loop
'    End If

'    FindEndOfLine = StartIndex

'End Function

'trims tabs and spaces from either end of a string
Function TabTrim(ByVal strItem As String, Optional ByVal eTrim As TrimType = Trim_Normal) As String

    strItem = Trim$(strItem)
    If Len(strItem) > 0 Then
        If eTrim < Trim_Right Then
            Do While Left$(strItem, 1) = vbTab
                strItem = Trim$(Mid$(strItem, 2))
            Loop
        End If
    End If
    If Len(strItem) > 0 Then
        If (eTrim <> Trim_Left) Then
            Do While Right$(strItem, 1) = vbTab
                strItem = Trim$(Left$(strItem, Len(strItem) - 1))
            Loop
        End If
    End If

    TabTrim = Trim$(strItem)

End Function

'inserts a string into a string, at position specified
Function StringInsert(ByVal strInsert As String, ByVal strItem As String, ByVal lngPosition As Long) As String

  Dim lngLength As Long

    lngLength = Len(strItem)
    If (lngPosition < lngLength) And (lngPosition > 1) Then
        StringInsert = Left$(strItem, lngPosition - 1) & strInsert & Right$(strItem, lngLength - (lngPosition - 1))
      Else
        If lngPosition = 1 Then
            StringInsert = strItem & strInsert
          Else
            StringInsert = strInsert & strItem
        End If
    End If

End Function

'returns the string in between 2 items. empty if nothing found
Function StrBetween(ByVal strExpression As String, ByVal strBegin As String, ByVal strEnd As String, Optional ByVal lngStart As Long = 1, Optional ByVal Compare As VbCompareMethod = vbTextCompare) As String

  Dim StartPlace As Long, EndPlace As Long, lngLength As Long

    lngLength = Len(strExpression)
    If lngStart <= lngLength Then
        StartPlace = InStr(lngStart, strExpression, strBegin, vbTextCompare)
        If StartPlace > 0 Then
            StartPlace = StartPlace + Len(strBegin)
            If StartPlace < lngLength Then
                EndPlace = InStr(StartPlace + 1, strExpression, strEnd, vbTextCompare)
                If EndPlace > 0 Then
                    StrBetween = Mid$(strExpression, StartPlace, EndPlace - StartPlace)
                    Exit Function
                End If
            End If
            'StrBetween = Mid$(strExpression, StartPlace)
          Else
            'StrBetween = strExpression
        End If
    End If

End Function

'finds the starting point of a vb comment in a string, returns zero if nothing found
'if found and autoremove is true then it removes the comment from the line by reference
Function FindComment(strCodeLine As String, Optional ByVal AutoRemove As Boolean) As Long

  Dim X As Long, Length As Long, Quote As String
  Dim InQuote As Boolean, CurrentChar As String

    Length = Len(strCodeLine)
    If Length > 0 Then
        Quote = Chr$(34)
        Do While X < Length
            X = X + 1
            CurrentChar = Mid$(strCodeLine, X, 1)
            If CurrentChar = Quote Then
                InQuote = Not InQuote
            ElseIf CurrentChar = "'" Then
                If Not InQuote Then
                    FindComment = X
                    Exit Do
                End If
            ElseIf X < Length - 3 Then
                If Not InQuote Then
                    If StrComp(Mid$(strCodeLine, X, 4), "Rem ") = 0 Then
                        FindComment = X
                        Exit Do
                    End If
                End If
            End If
        Loop
    End If

    If AutoRemove Then
        If (FindComment > 0) Then
            strCodeLine = Trim$(Left$(strCodeLine, FindComment - 1))
        End If
    End If

End Function

'used to trim brackets, quotes, or braces from both ends of a strings
Function TrimBrackets(ByVal strItem As String, ByVal strBracket As String) As String

  Dim Length As String
    
    strItem = Trim$(strItem)
    Length = Len(strItem)
    If Length > 1 Then
        If (Left$(strItem, 1) = strBracket) Then
            strItem = Trim$(Right$(strItem, Length - 1))
            Length = Len(strItem)
        End If
        If (Right$(strItem, 1) = strBracket) Then
            strItem = Trim$(Left$(strItem, Length - 1))
        End If
    End If
    TrimBrackets = strItem
    
End Function


