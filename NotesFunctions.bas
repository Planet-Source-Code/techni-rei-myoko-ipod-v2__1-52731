Attribute VB_Name = "NotesFunctions"
Option Explicit 'heres where i wish i didnt put special characters inside html brackets
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFilename As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
Private Const MenuWidth As Long = 145
Private Type HTMLLink
    start As Long
    lines As Long
    Destination As String
End Type
Private HTMLLinks() As HTMLLink, HTMLLinkCount As Long, currloc As String
Private HistoryList() As String, HistoryCount As Long, HistoryItem As Long
Public Sub NotesHistory(Index As Long, MNU)
    If Index <> 0 And HistoryCount > 1 Then 'if there are history items, and index isnt 0
        HistoryItem = HistoryItem + Index
        If HistoryItem < 0 Then HistoryItem = 0
        If HistoryItem >= HistoryCount Then HistoryItem = HistoryCount
        ProcessHTMLFile HistoryList(HistoryItem), MNU
    End If
End Sub
Public Sub AddtoHistory(filename As String)
    If HistoryCount > 0 Then If StrComp(HistoryList(HistoryItem), filename, vbTextCompare) = 0 Then Exit Sub 'if not viewing a history item
    If HistoryItem < HistoryCount - 1 Then 'is not the last item, cut off the rest
        ReDim Preserve HistoryList(HistoryItem + 1)
    End If
    HistoryCount = HistoryCount + 1
    ReDim Preserve HistoryList(HistoryCount)
    HistoryList(HistoryCount - 1) = filename
    HistoryItem = HistoryCount - 1
End Sub
Public Sub ClearHistory()
    HistoryCount = 0
    ReDim HistoryList(HistoryCount)
End Sub
Public Sub ExecuteNote(filename As String, MNU)
    Dim tempstr As String
    tempstr = chkpath(App.path, "Notes\" & filename & "\aTitle.txt")
    currloc = tempstr
    AddtoHistory filename
    ProcessHTMLFile tempstr, MNU
End Sub
Public Sub ExecuteLink(MNU)
    Dim tempstr As String
    tempstr = LCase(GetDestination(MNU))
    If Len(tempstr) > 0 Then
        If Left(tempstr, 5) <> "song=" And Left(tempstr, 5) <> "ipod:" Then
            If currloc Like "?:\*" Then
                currloc = Left(currloc, InStrRev(currloc, "\") - 1)
                If Left(tempstr, 7) <> "http://" Then tempstr = chkpath(currloc, tempstr)
            End If
            If Left(currloc, 7) = "http://" Then 'is a web page
                If InStrRev(currloc, "/") > 7 Then currloc = Left(currloc, InStrRev(currloc, "/") - 1)
                If Not (currloc Like "?:\*") Then tempstr = chkurl(currloc, tempstr)
            End If
        End If
        ProcessHTMLFile tempstr, MNU
        AddtoHistory tempstr
    End If
End Sub
'Brains of the whole sub section
Public Sub ProcessHTMLFile(ByVal filename As String, MNU)
    Dim tempstr As String, temp As Long
    On Error Resume Next
    filename = LCase(filename)
    currloc = filename
    
    If Left(filename, 7) = "http://" Then 'is a web page
        tempstr = chkpath(App.path, "iPodCache.html") 'set the filename to the cache file
        If DownloadFile(filename, tempstr) Then 'download the webpage to a cache file
            ProcessHTML HN.loadwholefile(tempstr), MNU
            Kill tempstr 'clean up the mess i made
        End If
    End If
    
    If filename Like "?:\*" Then 'is a local file
        MNU.ClearItems 'just in case the file doesnt exist
        If FileLen(filename) > 0 Then ProcessHTML HN.loadwholefile(filename), MNU 'process local file
    End If
    
    If Left(filename, 5) = "song=" Then 'is a link to a song via filename or its title
        filename = Right(filename, Len(filename) - 5) 'cut 'song=' off the front
        tempstr = GetFilenameFromTitle(filename)
        If Len(tempstr) > 0 Then filename = tempstr
        AddPlayItem filename
        PlayItem PlayCount - 1
        MenuMode = ipod_nowplaying
    End If
    
    If Left(filename, 5) = "ipod:" Then 'logical link to a section in the media database
        filename = Right(filename, Len(filename) - 5) 'cut 'ipod:' off the front
        temp = InStr(filename, "?")
        If temp = 0 Then temp = Len(filename) + 1
        tempstr = Left(filename, temp - 1)
        filename = Right(filename, Len(filename) - temp)
        temp = InStr(filename, "&") 'remove multiple searching as HINI has no SQL at the moment
        If temp > 0 Then filename = Left(filename, temp - 1) 'I tried but SQL is not good enough to handle a 3 dimensional database
        Select Case tempstr
            Case "music"
                temp = InStr(filename, "=")
                tempstr = "Browse\" & Left(filename, temp - 1) & "s\" & Right(filename, Len(filename) - temp)
                MainMenu lC, MNU, Tim, tempstr
        End Select
    End If
End Sub
Public Sub ProcessHTML(ByVal HTMLCode As String, MNU)
    Dim temp As Long, temp2 As Long, tempstr As String
    ClearHTMLLinks
    SetTitle
    MNU.ClearItems
    MNU.Locked = True
    Do Until Len(HTMLCode) = 0
        Select Case Left(HTMLCode, 1)
            Case "<" 'might be an html tag
                temp = InStr(HTMLCode, ">") 'location of the end of the tag
                If temp = 0 Then 'Randomly placed "<" messing up my code
                    'No idea what to do with it. dump it all as text?
                    tempstr = tempstr & HTMLCode
                    HTMLCode = Empty 'erase dumped text from HTMLcode
                Else
                    If CharExists(Left(HTMLCode, temp)) Then 'is a special char (pass as text)
                        tempstr = tempstr & Left(HTMLCode, temp)
                    Else 'is an html tag
                        If Len(tempstr) > 0 Then
                            ProcessText tempstr, MNU
                            tempstr = Empty
                        End If
                        temp2 = LocationOfEndTag(HTMLCode, tagname(HTMLCode))
                        If temp2 > 0 And temp2 < Len(HTMLCode) Then temp = InStr(temp2, HTMLCode, ">")
                        ProcessHTMLtag Left(HTMLCode, temp), MNU
                    End If
                    HTMLCode = Right(HTMLCode, Len(HTMLCode) - temp)
                End If
            Case Else 'definetly is not
                temp = InStr(HTMLCode, "<") 'location of first tag
                If temp = 0 Then 'no tags found, dump the HTMLCode to the buffer
                    tempstr = tempstr & HTMLCode
                    HTMLCode = Empty 'erase dumped text from HTMLcode
                Else 'Tag found, dump the HTMLcode before the tag into the buffer
                    tempstr = tempstr & Left(HTMLCode, temp - 1)
                    HTMLCode = Right(HTMLCode, Len(HTMLCode) - temp + 1) 'erase dumped text from HTMLcode
                End If
        End Select
        DoEvents
    Loop
    If Len(tempstr) > 0 Then ProcessText tempstr, MNU
    MNU.Locked = False
End Sub
Public Function GetDestination(MNU) As String
    Dim temp As Long
    temp = Selected2Link(MNU)
    If temp > -1 Then GetDestination = HTMLLinks(temp).Destination
End Function

'Logistics
Private Function Selected2Link(MNU) As Long
    Dim temp As Long
    Selected2Link = -1
    If HTMLLinkCount > 0 Then
        For temp = 0 To HTMLLinkCount
            With HTMLLinks(temp)
                If MNU.selecteditem >= .start And MNU.selecteditem < .start + .lines Then
                    Selected2Link = temp
                    Exit For
                End If
            End With
        Next
    End If
End Function
Private Function AddHTMLLink(start As Long, lines As Long, Dest As String) As Long
    HTMLLinkCount = HTMLLinkCount + 1
    ReDim Preserve HTMLLinks(HTMLLinkCount)
    With HTMLLinks(HTMLLinkCount - 1)
        .Destination = Dest
        .start = start
        .lines = lines
    End With
    AddHTMLLink = HTMLLinkCount - 1
End Function
Private Sub ClearHTMLLinks()
    HTMLLinkCount = 0
    ReDim HTMLLinks(0)
End Sub
Private Function LocationOfEndTag(HTMLCode As String, Tag As String) As Long
    LocationOfEndTag = InStr(1, HTMLCode, "</" & tagname(HTMLCode) & ">", vbTextCompare)
End Function
Public Sub ProcessText(text As String, MNU, Optional Underline As Boolean)
    Dim tempstr As String, temp As Long
    Do Until Len(text) = 0
        tempstr = WrapLine(text, MenuWidth)
        temp = GetNextChar(tempstr, vbNewLine & Chr(13))
        Select Case Left(tempstr, 1)
            Case " ", vbNewLine: tempstr = Right(tempstr, Len(tempstr) - 1)
        End Select
        MNU.Underline(MNU.NewItem(tempstr)) = Underline
    Loop
End Sub
Private Sub ProcessHTMLtag(ByVal text As String, MNU)
    Dim tempstr As String, Destination As String, temp As Long
    Select Case LCase(tagname(text))
        Case "br"
        Case "p": MNU.NewItem Empty
        Case "title": SetTitle addfrom(text, "node")
        Case "a"
            tempstr = addfrom(text, "node") 'Caption
            Destination = addfrom(text, "href") 'Destination
            AddHTMLLink MNU.itemcount, LineCount(tempstr, MenuWidth), Destination
            ProcessText tempstr, MNU, True
        Case Else: LCase (tagname(text)) & " was not recognised"
    End Select
End Sub
Private Sub SetTitle(Optional title As String = "Notes")
    With lC
        .ClearText
        TitleBar lC, title
        .LCDRefresh
    End With
End Sub

'Code borrowed from my NetDownloader
Public Function DownloadFile(url As String, filename As String) As Boolean
On Error Resume Next 'Downloads the file from URL and saves it as filename
If Len(url) > 0 And Len(filename) > 0 Then DownloadFile = URLDownloadToFile(0, url, filename, 0, 0) = 0
End Function
Public Function chkurl(ByVal basehref As String, url As String) As String
'check for absolute (is like *://*)
'check for relative (contains ../)
'check for additive (else)
Dim spoth As Long
If Left(url, 1) = "#" Then Exit Function 'is not a file
If Left(url, 1) = "/" Then url = Right(url, Len(url) - 1)
'If Len(basehref) - Len(Replace(basehref, "/", Empty)) = 2 Then basehref = basehref & "/"
If containsword(basehref, "://") = False Then basehref = "http://" & basehref
If LCase(url) <> LCase(basehref) And url <> Empty And basehref <> Empty Then
If url Like "*://*" Then 'is absolute
    chkurl = url
Else
    If containsword(url, "../") Then 'is relative
        If Right(basehref, 1) = "/" And Len(basehref) - Len(Replace(basehref, "/", Empty)) > 2 Then basehref = Left(basehref, Len(basehref) - 1)
        If containsword(Replace(basehref, "://", ""), "/") = True Then
            For spoth = 1 To countwords(basehref, "../")
                url = Right(url, Len(url) - Len("../"))
                basehref = Left(basehref, InStrRev(basehref, "/"))
            Next
        Else
            url = Replace(url, "../", "")
        End If
        If Right(basehref, 1) <> "/" Then chkurl = basehref & "/" & url Else chkurl = basehref & url
    Else 'is additive
        If Right(basehref, 1) <> "/" Then chkurl = basehref & "/" & url Else chkurl = basehref & url
    End If
End If
End If
End Function
Public Function addfrom(content As String, Tag As String) As String
    Dim temp As Long, location As Long, temp2 As Long
    If LCase(Tag) <> "node" Then
        location = InStr(1, content, Tag, vbTextCompare)
        If location > 0 Then
            location = InStr(location, content, "=") + 1
            Select Case mid(content, location, 1)
                Case """", "'"
                    location = location + 1
                    temp = InStr(location, content, """")
                    If temp = 0 Then temp = InStr(location, content, "'")
                    temp2 = InStr(location, content, ">")
                Case Else
                    temp = InStr(location, content, " ")
                    temp2 = InStr(location, content, ">")
            End Select
            If temp2 < temp And temp2 > 0 Then temp = temp2
            If temp = 0 Then temp = InStr(location, content, ">")
            If temp = 0 Then temp = Len(content)
            addfrom = mid(content, location, temp - location)
        End If
    Else
        addfrom = removebrackets(content, "<", ">")
    End If
End Function

Public Function countwords(text As String, word As String) As Long
    countwords = (Len(text) - Len(Replace(text, word, Empty, , , vbTextCompare))) \ Len(word)
End Function
Public Function hasHTMLtag(content As String, Tag As String) As Boolean
    hasHTMLtag = InStr(1, content, Tag & "=", vbTextCompare) > 0
End Function
Public Function tagname(content As String) As String
    Dim temp As Long, temp2 As Long
    temp = InStr(content, " ")
    temp2 = InStr(content, ">")
    If temp > 0 And temp < temp2 Then temp2 = temp
    tagname = mid(content, 2, temp2 - 2)
End Function
Public Function removetext(text As String, start As Long, finish As Long, Optional exclusive As Boolean = True) As String
    If exclusive = True Then
        removetext = Left(text, start - 1) & Right(text, Len(text) - finish)
    Else
        removetext = mid(text, start, finish - start)
    End If
End Function
Public Function removebrackets(ByVal text As String, leftb As String, rightb As String) As String
    Do While InStr(text, leftb) > 0 And InStr(text, rightb) > InStr(text, leftb)
        text = removetext(text, InStr(text, leftb), InStr(text, rightb))
    Loop
    removebrackets = text
End Function
Public Function isCancel(ByVal Tag As String) As Boolean
    If Left(Tag, 1) = "<" Then Tag = Right(Tag, Len(Tag) - 1)
    isCancel = Left(Tag, 1) = "/"
End Function
Public Function CleanTag(ByVal Tag As String) As String
    Tag = tagname(Tag)
    If Left(Tag, 1) = "/" Then Tag = Right(Tag, Len(Tag) - 1)
    CleanTag = LCase(Tag)
End Function

