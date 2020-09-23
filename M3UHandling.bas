Attribute VB_Name = "M3UHandling"
Option Explicit
Public Const PlayListFiles As String = "*.m3u;*.pls;*.b4s;*.asx;*.smil;*.wax"
Public Sub LoadPlayList(ByVal filename As String, STRarray() As String, itemcount As Long)
    Dim tempfile As Long, tempstr As String, foundparam As Boolean
    'On Error Resume Next
    tempfile = FreeFile
    If FileLen(filename) > 0 Then
        Open filename For Input As #tempfile
            filename = left(filename, InStrRev(filename, "\") - 1)
            Line Input #tempfile, tempstr
            tempstr = Replace(Trim(tempstr), " = """, "=""")
            tempstr = Replace(Trim(tempstr), """ >", """>")
            Select Case Trim(LCase(tempstr))
                Case "#extm3u" 'm3u playlist format
                    Do Until EOF(tempfile)
                        Line Input #tempfile, tempstr
                        If StrComp(left(tempstr, 4), "#EXT", vbTextCompare) <> 0 Then '%filename%
                            tempstr = Replace(tempstr, "\.\", "\") 'whos bloody idea was it to seperate with a period for no good damned reason
                            AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount
                        End If
                    Loop
                    
                Case "[playlist]" 'pls playlist format
                    Do Until EOF(tempfile)
                        Line Input #tempfile, tempstr
                        If StrComp(left(tempstr, 4), "file", vbTextCompare) = 0 Then 'File#=%filepath%\.\%filetitle%
                            tempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, "="))
                            tempstr = Replace(tempstr, "\.\", "\") 'whos bloody idea was it to seperate with a period for no good damned reason
                            AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount
                        End If
                    Loop
                    
                Case "<?xml version=""1.0"" encoding='utf-8' standalone=""yes""?>", "<?xml version=""1.0"" encoding=""utf-8""?>" 'Some god forsaken xml format (b4s, and itunes)
                    Do Until EOF(tempfile)
                        Line Input #tempfile, tempstr
                        tempstr = Trim(Replace(tempstr, vbTab, Empty))
                        If StrComp(left(tempstr, 18), "<entry Playstring=", vbTextCompare) = 0 Then '<entry Playstring="file:%filepath%\.\%filetitle%">
                            tempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, """"))
                            If StrComp(left(tempstr, 5), "file:", vbTextCompare) = 0 Then
                                tempstr = Right(tempstr, Len(tempstr) - 5)
                            End If
                            tempstr = left(tempstr, InStrRev(tempstr, """") - 1)
                            tempstr = Replace(tempstr, "\.\", "\") 'whos bloody idea was it to seperate with a period for no good damned reason
                            AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount
                        Else 'The ipod format is useless as you cant get them off the thing
                            'Does not support multiline statements
                            If StrComp(left(tempstr, 27), "<key>Location</key><string>", vbTextCompare) = 0 Then '<key>Location</key><string>%filename%</string>
                                tempstr = Mid(tempstr, 28, Len(tempstr) - 36)
                                AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount
                            End If
                        End If
                    Loop
                    
                Case "<asx version=""3.0"">", "<smil>" 'asx and smil format
                    Do Until EOF(tempfile)
                        Line Input #tempfile, tempstr
                        tempstr = Replace(Trim(tempstr), " = """, "=""")
                        If InStr(1, tempstr, "Name=""SourceURL""", vbTextCompare) > 1 Then '<Param Name="SourceURL" Value="%filename%" />
                            tempstr = Right(tempstr, Len(tempstr) - InStr(1, tempstr, "Value", vbTextCompare))
                            tempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, """"))
                            tempstr = left(tempstr, InStrRev(tempstr, """") - 1)
                            AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount
                            foundparam = True
                        Else
                            If Not foundparam Then 'damned asx files have 2 listings of the filename >:(
                                If InStr(1, tempstr, "HREF=""", vbTextCompare) > 1 Or StrComp(left(tempstr, 6), "<audio", vbTextCompare) = 0 Then ' <REF HREF="%filename%" /> or <audio src="%filename%"/>
                                    tempstr = Right(tempstr, Len(tempstr) - InStr(tempstr, """"))
                                    tempstr = left(tempstr, InStrRev(tempstr, """") - 1)
                                    AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount
                                End If
                            End If
                        End If
                    Loop
                    
                Case Else 'Simple m3u format (ie: raw file list, the best format of all)
                    AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount 'First line was an item too
                    Do Until EOF(tempfile)
                        Line Input #tempfile, tempstr
                        AddPlayListItem chkpath(filename, tempstr), STRarray, itemcount '%filename%
                    Loop
                'God I hate XML, all the formats are too similar to warrant seperate formats
                'rdf is just too damned stupid to implement
            End Select
        Close #tempfile
    End If
End Sub
Private Sub AddPlayListItem(ByVal filename As String, STRarray() As String, itemcount As Long)
    itemcount = itemcount + 1
    ReDim Preserve STRarray(itemcount)
    STRarray(itemcount - 1) = filename
End Sub
Private Function containsword(phrase As String, word As String) As Boolean
    If Replace(phrase, word, Empty) <> phrase Then containsword = True Else containsword = False
End Function
Private Function countwords(phrase As String, word As String) As Long
    countwords = (Len(phrase) - Len(Replace(phrase, word, Empty))) / Len(word)
End Function
Private Function chkpath(ByVal basehref As String, ByVal url As String) As String
'Debug.Print basehref & " " & URL
Const goback As String = "..\"
Const slash As String = "\"
Dim spoth As Long
If left(url, 2) = ".\" Then url = Right(url, Len(url) - 2)
If left(url, 1) = slash Then url = Right(url, Len(url) - 1)
If Right(basehref, 1) = slash And Len(basehref) > 3 Then basehref = left(basehref, Len(basehref) - 1)
If LCase(url) <> LCase(basehref) And url <> Empty And basehref <> Empty Then
If url Like "?:*" Then 'is absolute
    chkpath = url
Else
    If containsword(url, goback) Then 'is relative
        If containsword(Right(basehref, Len(basehref) - 3), slash) = True Then
            For spoth = 1 To countwords(url, goback)
                If countwords(basehref, slash) > 0 Then
                    url = Right(url, Len(url) - Len(goback))
                    basehref = left(basehref, InStrRev(basehref, slash) - 1)
                Else
                    url = Replace(url, goback, "")
                End If
            Next
        Else
            url = Replace(url, goback, "")
        End If
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & url Else chkpath = basehref & url
    Else 'is additive
        If Right(basehref, 1) <> slash Then chkpath = basehref & slash & url Else chkpath = basehref & url
    End If
End If
End If
End Function
