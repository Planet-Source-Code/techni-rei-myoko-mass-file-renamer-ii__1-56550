Attribute VB_Name = "RenameFunctions"
Option Explicit
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Const chars = "_.- "

Public Function long2dos(ByRef Filename As String) As String
    Dim sFile As String, sShortFile As String * 67, lRet As Long
    sFile = Filename
    lRet = GetShortPathName(sFile, sShortFile, Len(sShortFile))
    sFile = Left(sShortFile, lRet)
    long2dos = sFile
End Function

Public Sub SeedListView(Source As FileListBox, Destination As ListView, IML As ImageList, PIC As PictureBox)
    Dim temp As Long, tempstr As String
    Destination.Tag = Source.Path
    For temp = 0 To Source.ListCount - 1
        tempstr = LCase(chkdir(Source.Path, Source.List(temp)))
        Destination.ListItems.Add , tempstr, Source.List(temp), , GetIcon(tempstr, IML, PIC)
    Next
End Sub

Public Function ResetList(Destination As ListView, Optional action As Long) As Boolean
    Dim temp As Long, tempstr As String
    ResetList = True
    With Destination
        For temp = 1 To .ListItems.count
            Select Case action
                Case 0
                    .ListItems(temp).SubItems(1) = .ListItems(temp).text
                    .ListItems(temp).SubItems(2) = getFileDateTime(.ListItems(temp).Key)
                Case 1: .ListItems(temp).text = .ListItems(temp).SubItems(1)
                Case 2
                    If StrComp(.ListItems.Item(temp).SubItems(1), .ListItems.Item(temp).text, vbTextCompare) <> 0 Then
                        tempstr = GetPath(.ListItems.Item(temp).Key)
                        If RenameFile(.ListItems.Item(temp).Key, chkdir(tempstr, .ListItems.Item(temp).text)) Then
                            .ListItems.Item(temp).SubItems(1) = .ListItems.Item(temp).text
                        Else
                            .ListItems.Item(temp).ForeColor = vbRed
                            ResetList = False
                        End If
                    End If
            End Select
        Next
    End With
End Function

Public Function getFileDateTime(Filename As String) As Date
    On Error Resume Next
    getFileDateTime = Now
    getFileDateTime = FileDateTime(Filename)
End Function

Public Function RenameFile(Source As String, Destination As String) As Boolean
    On Error Resume Next
    Name Source As Destination
    RenameFile = True
End Function

Public Function FileExists(Filename As String) As Boolean
On Error Resume Next
    If Dir(Filename, vbNormal + vbHidden + vbSystem + vbDirectory) <> Empty And Filename <> Empty Then FileExists = True Else Filename = False
End Function

Public Function CopyFile(Source As String, Destination As String) As Boolean
    On Error Resume Next
    If StrComp(Left(Source, 1), Left(Destination, 1), vbTextCompare) = 0 Then
        CopyFile = RenameFile(Source, Destination)
    Else
        FileCopy Source, Destination
        Kill Source
        CopyFile = True
    End If
End Function

Public Function UniqueFilename(Filename As String) As String
    Dim temp1 As String, temp2 As String, temp3 As Long
    UniqueFilename = Filename
    
    If FileExists(Filename) Then
        Dim count As Long
        count = 1
        temp3 = InStrRev(Filename, ".")
        temp1 = Filename
        If temp3 > 0 Then
            temp1 = Left(Filename, temp3 - 1)
            temp2 = Right(Filename, Len(Filename) - temp3 + 1)
        End If
        Do Until FileExists(temp1 & "(" & count & ")" & temp2) = False
            count = count + 1
        Loop
        UniqueFilename = temp1 & "(" & count & ")" & temp2
    End If
End Function

Public Sub ReplaceText(LST As ListView, Find As String, Replacement As String, Optional MustbeSelected As Boolean)
    Dim temp As Long, DOIT As Boolean
    With LST
        For temp = 1 To .ListItems.count
            DOIT = True
            If MustbeSelected Then DOIT = .ListItems.Item(temp).Selected
            If DOIT Then .ListItems.Item(temp).text = Replace(.ListItems.Item(temp).text, Find, Replacement, , , vbTextCompare)
        Next
    End With
End Sub

Public Function NumericIndex(LST As ListView, Pattern As String, Value As Long, Optional Digits As Long, Optional MustbeSelected As Boolean) As Long
    Dim temp As Long, DOIT As Boolean, tempstr As String
    With LST
        For temp = 1 To .ListItems.count
            DOIT = True
            If MustbeSelected Then DOIT = .ListItems.Item(temp).Selected
            If DOIT Then
                tempstr = .ListItems.Item(temp).text
                If Digits <= 0 Then
                    .ListItems.Item(temp).text = SeedIndex(.ListItems.Item(temp).text, Pattern, killallceptnumber(.ListItems.Item(temp).text, Abs(Digits)), 0)
                Else
                    .ListItems.Item(temp).text = SeedIndex(.ListItems.Item(temp).text, Pattern, CStr(Value), Digits)
                    Value = Value + 1
                End If
                .ListItems.Item(temp).text = DoDateTime(.ListItems.Item(temp).text, .ListItems.Item(temp).SubItems(2))
                If InStr(1, .ListItems.Item(temp).text, "[OLD]", vbTextCompare) > 0 Then
                    If InStr(tempstr, ".") > 0 Then tempstr = Left(tempstr, InStrRev(tempstr, ".") - 1)
                    .ListItems.Item(temp).text = Replace(.ListItems.Item(temp).text, "[OLD]", tempstr, , , vbTextCompare)
                End If
            End If
        Next
    End With
    If Digits > 0 Then NumericIndex = Value
End Function

Public Function DoDateTime(ByVal text As String, DateTime As Date) As String
    Dim temp As Long
    text = Replace(text, "[DAY]", Format(DateTime, "dd"), , , vbTextCompare)
    
    text = Replace(text, "[MONTH]", Format(DateTime, "MM"), , , vbTextCompare)
    text = Replace(text, "[SMONTH]", Format(DateTime, "MMM"), , , vbTextCompare)
    text = Replace(text, "[LMONTH]", Format(DateTime, "MMMM"), , , vbTextCompare)
    
    text = Replace(text, "[YEAR]", Format(DateTime, "yyyy"), , , vbTextCompare)
    text = Replace(text, "[SYEAR]", Format(DateTime, "yy"), , , vbTextCompare)
    
    temp = Format(DateTime, "hh")
    If temp = 0 Then temp = 12
    text = Replace(text, "[12H]", temp, , , vbTextCompare)
    text = Replace(text, "[24H]", Format(DateTime, "hh"), , , vbTextCompare)
    
    text = Replace(text, "[MIN]", Format(DateTime, "nn"), , , vbTextCompare)
    text = Replace(text, "[AMPM]", IIf(Format(DateTime, "hh") > 12, "PM", "AM"), , , vbTextCompare)
    
    text = Replace(text, "[SEC]", Format(DateTime, "ss"), , , vbTextCompare)
    
    DoDateTime = text
End Function

Public Function SeedIndex(Filename As String, ByVal Pattern As String, Value As String, Digits As Long, Optional Number As String = "#", Optional Old As String = "[OLD]", Optional Extention As String = "[EXT]")
    Pattern = Replace(Pattern, Old, GetFilenoext(Filename), , , vbTextCompare)
    Pattern = Replace(Pattern, Extention, GetExtention(Filename), , , vbTextCompare)
    If Digits = 0 Then
        Pattern = Replace(Pattern, Number, Value, , , vbTextCompare)
    Else
        Pattern = Replace(Pattern, Number, Format(Value, String(Digits, "0")), , , vbTextCompare)
    End If
    SeedIndex = Pattern
End Function

Public Function MakeUnique(NewPath As String, Filename As String) As String
    MakeUnique = UniqueFilename(chkdir(NewPath, Filename))
End Function

Public Sub MakeAllUnique(LST As ListView, Optional MustbeSelected As Boolean, Optional hWnd As Long)
    Dim temp As Long, tempstr As String, DOIT As Boolean, DOALL As Boolean, tempstr2 As String
    tempstr = BrowseForFolder(hWnd, "Select the destination")
    If Len(tempstr) = 0 Then Exit Sub
    DOALL = MsgBox("Do you want to move the file(s) to " & tempstr & " after being renamed?", vbYesNo, "Move files?") = vbYes
    With LST
        For temp = 1 To .ListItems.count
            DOIT = True
            If MustbeSelected Then DOIT = .ListItems.Item(temp).Selected
            If DOIT Then
                tempstr2 = MakeUnique(tempstr, .ListItems.Item(temp).text)
                .ListItems.Item(temp).text = GetFilename(tempstr2)
                If DOALL Then CopyFile .ListItems.Item(temp).Key, tempstr2
            End If
        Next
    End With
End Sub

Public Sub MakeAllDOS(LST As ListView, Optional MustbeSelected As Boolean)
    Dim temp As Long, DOIT As Boolean
    With LST
        For temp = 1 To .ListItems.count
            DOIT = True
            If MustbeSelected Then DOIT = .ListItems.Item(temp).Selected
            If DOIT Then .ListItems.Item(temp).text = GetFilename(long2dos(.ListItems.Item(temp).Key))
        Next
    End With
End Sub

Public Function killallceptnumber(text As String, Optional Digits As Long) As String
    Dim temp As Long, tempstr As String, tempstr2() As String
    For temp = 1 To Len(text)
        If IsNumeric(Mid(text, temp, 1)) Then
            tempstr = tempstr & Mid(text, temp, 1)
        Else
            If Len(tempstr) > 0 Then
                If Right(tempstr, 1) <> " " Then tempstr = tempstr & " "
            End If
        End If
    Next
    If tempstr = Empty Then tempstr = 1
    If Digits > 0 Then
        If InStr(tempstr, " ") = 0 Then
            tempstr = Format(tempstr, String(Digits, "0"))
        Else
            tempstr2 = Split(tempstr, " ")
            For temp = 0 To UBound(tempstr2)
                tempstr2(temp) = Format(tempstr2(temp), String(Digits, "0"))
            Next
            tempstr = Join(tempstr2, " ")
        End If
    End If
    killallceptnumber = tempstr
End Function

Public Function removetext(text As String, Start As Long, finish As Long, Optional exclusive As Boolean = True) As String
    If exclusive = True Then
        removetext = Left(text, Start - 1) & Right(text, Len(text) - finish)
    Else
        removetext = Mid(text, Start, finish - Start)
    End If
End Function

Public Function RemoveBrackets(ByVal text As String, leftb As String, rightb As String) As String
    Do While InStr(text, leftb) > 0 And InStr(text, rightb) > InStr(text, leftb)
        text = removetext(text, InStr(text, leftb), InStr(text, rightb))
    Loop
    RemoveBrackets = text
End Function

Public Function removeallbutlast(ByVal text As String, char As String, replacewith As String) As String
    removeallbutlast = Replace(text, char, replacewith, 1, countchars(text, char) - 1, vbTextCompare)
End Function

Public Function killchars(ByVal text As String, filter As String, Optional replacewith As String = Empty) As String
    Dim count As Long
    For count = 1 To Len(text)
        If Replace(filter, Mid(text, count, 1), Empty) <> filter Then
            text = Left(text, count - 1) & replacewith & Right(text, Len(text) - count)
        End If
    Next
    killchars = text
End Function

Public Function replacedoubles(ByVal text As String, char As String) As String
    Do While InStr(text, char & char) > 0
        text = Replace(text, char & char, char)
    Loop
    replacedoubles = text
End Function

Public Function removefirst(ByVal text As String, char As String) As String
    Do Until Left(text, Len(char)) <> char
        text = Right(text, Len(text) - Len(char))
    Loop
    removefirst = text
End Function

Public Function killnonalpha(ByVal text As String) As String
    Dim temp As Long
    Do Until temp >= Len(text)
        temp = temp + 1
        If isalphanumeric(Mid(text, temp, 1)) = False Then text = Replace(text, Mid(text, temp, 1), Empty)
    Loop
    killnonalpha = text
End Function

Public Function isalphanumeric(text As String, Optional includenumeric As Boolean = True, Optional includepunctuation As Boolean) As Boolean
    isalphanumeric = False
    text = Left(LCase(text), 1)
    If text >= "a" And text <= "z" Then isalphanumeric = True
    If includenumeric Then If text >= "0" And text <= "9" Then isalphanumeric = True
    If includepunctuation Then If InStr(chars, text) > 0 Then isalphanumeric = True
End Function

Public Function countchars(text As String, char As String) As Long
    Dim count As Long, counter As Long
    counter = 0
    For count = 1 To Len(text)
        If Mid(text, count, Len(char)) = char Then counter = counter + 1
    Next
    countchars = counter
End Function

Public Function ChangeExtention(ByVal Filename As String, Extention As String) As String
    Dim temp As Long
    temp = InStrRev(Filename, ".")
    If temp = 0 Then
        ChangeExtention = Filename & "." & Extention
    Else
        ChangeExtention = Left(Filename, temp) & Extention
    End If
End Function

Public Function AnimeRename(ByVal Filename As String, Optional doRemoveBrackets As Boolean, Optional doUnderscores As Boolean, Optional doCommas As Boolean, Optional doPeriods As Boolean, Optional doCorrectExtentions As Boolean, Optional doPunctuation As Boolean, Optional doWholeWords As Boolean, Optional WholeWords As String, Optional doCapitalize As Boolean, Optional doDigits As Boolean, Optional Digits As Long, Optional doRemoveWords As Boolean, Optional Words2Remove As String, Optional doCapitalizeWords As Boolean, Optional Words2Capitalize As String, Optional doLowerCaseWords As Boolean, Optional Words2LowerCase As String, Optional doForceLowerExt, Optional Prepend As String = "-") As String
    Dim tempstr As String, tempstr2() As String, temp As Long, temp2 As Long, ext As String
    temp = InStr(Filename, ".")
    If temp > 0 Then
        ext = Right(Filename, Len(Filename) - temp + 1)
        Filename = Left(Filename, temp - 1)
    End If
    If doRemoveBrackets Then
        Filename = RemoveBrackets(Filename, "<", ">")
        Filename = RemoveBrackets(Filename, "(", ")")
        Filename = RemoveBrackets(Filename, "[", "]")
        Filename = RemoveBrackets(Filename, "{", "}")
    End If
    If doCorrectExtentions Then
        Select Case LCase(ext)
            Case ".ogm": ext = ".avi"
            Case ".nfo", ".diz": ext = ".txt"
        End Select
    End If
    If doForceLowerExt Then ext = LCase(ext)
    
    If doUnderscores Then Filename = Replace(Filename, "_", " ")
    If doCommas Then Filename = Replace(Filename, ",", " ")
    If doPeriods Then Filename = Replace(Filename, ".", " ")
    
    If Not doWholeWords Then WholeWords = Empty
    temp2 = SplitByCharType(Filename, tempstr2, WholeWords) 'Split by char type
    If temp2 > 0 Then
        For temp = 0 To temp2 - 1
            If doCapitalize Then If tempstr2(temp) = LCase(tempstr2(temp)) Or tempstr2(temp) = UCase(tempstr2(temp)) Then tempstr2(temp) = UCase(Left(tempstr2(temp), 1)) & LCase(Right(tempstr2(temp), Len(tempstr2(temp)) - 1))    'Capitalize first letter of each word
            If IsNumeric(tempstr2(temp)) And doDigits Then
                If Len(tempstr2(temp)) < Digits Then
                    tempstr2(temp) = Prepend & Format(tempstr2(temp), String(Digits, "0")) 'Force 2 digits
                Else
                    tempstr2(temp) = Prepend & tempstr2(temp)
                End If
            End If
            If doRemoveWords Then tempstr2(temp) = MultiWords(tempstr2(temp), Words2Remove, 0) 'Remove words
            If doCapitalizeWords Then tempstr2(temp) = MultiWords(tempstr2(temp), Words2Capitalize, 1) 'Upper/Lower case words
            If doLowerCaseWords Then tempstr2(temp) = MultiWords(tempstr2(temp), Words2LowerCase, 2)
        Next
        Filename = Join(tempstr2, " ")  'end split
    End If
    
    Filename = Filename & ext
    If doPunctuation Then
        Filename = replacedoubles(Filename, " ") 'Remove Double spaces
        Filename = removefirst(Filename, " ") 'Remove the first char if its a space or minus
        Filename = removefirst(Filename, "-")
        Filename = Replace(Filename, " -", "-") 'Makes sure each '-' is surrounded by spaces, and never 2 in a row
        Filename = Replace(Filename, "- ", "-")
        Filename = replacedoubles(Filename, "-")
        Filename = Replace(Filename, "-", " - ")
        Filename = Replace(Filename, " .", ".") 'Remove spaces before extentions
        Filename = Replace(Filename, ". ", ".") 'Remove spaces after extentions
    End If
    
    AnimeRename = Filename
End Function

Public Function MultiWords(text As String, Words As String, action As Long) As String
    '0 = remove, 1 = capitalize, 2 = lowercase
    Dim temp As Long, tempstr() As String
    If InStr(Words, " ") = 0 Then
        ReDim tempstr(1)
        tempstr(0) = Words
    Else
        tempstr = Split(Words, " ")
    End If
    For temp = 0 To UBound(tempstr)
        If StrComp(text, tempstr(temp), vbTextCompare) = 0 Then
            Select Case action
                Case 0
                Case 1: MultiWords = UCase(text)
                Case 2: MultiWords = LCase(text)
            End Select
            Exit Function
        End If
    Next
    MultiWords = text
End Function

Public Function StringFormat(ByVal text As String, Optional numofzeros As Long = 2) As String
    Dim tempstr As String, tempstr2 As String, oldtype As Boolean
    tempstr2 = Left(text, 1)
    oldtype = IsNumeric(tempstr2)
    text = Right(text, Len(text) - 1)
    Do Until Len(text) = 0
        If IsNumeric(Left(text, 1)) <> oldtype Then
            If oldtype Then tempstr2 = Format(tempstr2, String(numofzeros, "0"))
            tempstr = tempstr & tempstr2
            tempstr2 = Left(text, 1)
            oldtype = Not oldtype
        Else
            tempstr2 = tempstr2 & Left(text, 1)
        End If
        text = Right(text, Len(text) - 1)
    Loop
    StringFormat = tempstr
End Function

Public Function GetEndOfWord(text As String, Optional Start As Long = 1, Optional ByVal WholeWords As String) As Long
    'Check for whole words
    Dim temp As Long, tempstr() As String, temp2 As Long
    If Len(WholeWords) > 0 Then
        If InStr(WholeWords, ";") = 0 Then
            If isOverlay(text, Start, WholeWords) Then
                GetEndOfWord = Start + Len(WholeWords)
                Exit Function
            End If
        Else
            tempstr = Split(WholeWords, ";")
            For temp = 0 To UBound(tempstr)
                If isOverlay(text, Start, tempstr(temp)) Then
                    GetEndOfWord = Start + Len(tempstr(temp))
                    Exit Function
                End If
            Next
        End If
    End If
    'Checks for end of normal words
    temp2 = CharType(Mid(text, Start, 1))
    temp = Start + 1
    Do Until temp > Len(text) Or CharType(Mid(text, temp, 1)) <> temp2
        temp = temp + 1
    Loop
    GetEndOfWord = temp
End Function

Public Function SplitByCharType(ByVal text As String, strArray, Optional WholeWords As String) As Long
    Dim temp As Long, temp2 As Long
    Do Until Len(text) = 0
        'Get the word
        temp = GetEndOfWord(text, , WholeWords) - 1
        
        If Left(text, temp) <> " " Then 'if the word was not a space
            'Redim the array, and push the word in to the new cell
            temp2 = temp2 + 1
            ReDim Preserve strArray(temp2)
            strArray(temp2 - 1) = Left(text, temp)
        End If
        
        'Remove the word from the string
        text = Right(text, Len(text) - temp)
    Loop
    SplitByCharType = temp2
End Function

Public Function isOverlay(text As String, Start As Long, Word As String) As Boolean
    isOverlay = StrComp(Mid(text, Start, Len(Word)), Word, vbTextCompare) = 0
End Function

Public Function CharType(text As String) As Long
    CharType = -1
    text = LCase(text)
    If text >= "a" And text <= "z" Then CharType = 0
    If text >= "0" And text <= "9" Then CharType = 1
    If InStr(chars, text) > 0 Then CharType = 2
End Function

Public Function getfromquotes(ByVal text As String) As String
    If Left(text, 1) = """" Then text = Right(text, Len(text) - 1)
    If Right(text, 1) = """" Then text = Left(text, Len(text) - 1)
    getfromquotes = text
End Function
