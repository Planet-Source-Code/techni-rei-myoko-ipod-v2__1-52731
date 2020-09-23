Attribute VB_Name = "DatabaseFunctions"
Option Explicit
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Public Enum DriveConstants
    SectorsPerCluster = 0
    BytesPerSector = 1
    NumberOfFreeClusters = 2
    TotalNumberOfClusters = 3
    sectorsize = 4
    TotalSize = 5
    FreeSpace = 6
    UsedSpace = 7
End Enum

Public HN As Hini, CurrPlayList As String
'Used for ToDo list or section key dumping
Public Sub DumpKeys(Section As String, MNU)
    Dim tempstr() As String, temp As Long, temp2 As Long
    HN.enumeratekeys Section, tempstr
    temp2 = HN.sectioncount(Section)
    For temp = 1 To temp2
        MNU.NewItem tempstr(1, temp), tempstr(2, temp)
    Next
End Sub
'Used for Calender
Public Function HasPlans(Day As Date) As Boolean
    HasPlans = PlanCount(Day) > 0
End Function
Public Function PlanCount(Day As Date) As Long
    PlanCount = HN.keycount("Calender\" & DateValue(Day))
End Function
Public Sub ExecuteDate(lcdm, MNU, Day As Date)
    lcdm.ClearText
    MNU.Visible = True
    MNU.ClearItems
    MNU.HideSelected = True
    TitleBar lcdm, Format(Day, "mMmm dd, yyyy")
    Dim temp As Long, tempstr() As String, temp2 As Long
    temp2 = PlanCount(Day)
    HN.enummultikey "Calender\" & DateValue(Day), "Event", tempstr
    For temp = 1 To temp2
        MNU.NewItem tempstr(temp)
    Next
    lcdm.LCDRefresh
End Sub

'Used to obtain hard drive details
Public Function GetSize(ByVal Size, Optional Bytes As String = "B", Optional Kilo As String = "K", Optional Mega As String = "M", Optional Giga As String = "G") As String
    Select Case Val(Size)
        Case 0 To 1023
            GetSize = Val(Size) & Bytes
        Case 1024 To 1048576
            GetSize = Round(Val(Size) / 1024, 2) & Kilo
        Case 1048576 To 1073741824
            GetSize = Round(Val(Size) / 1048576, 2) & Mega
        Case Is > 1073741824
            GetSize = Round(Val(Size) / 1073741824, 2) & Giga
    End Select
End Function
Public Function DriveDetail(DriveLetter As String, Detail As DriveConstants) As Long
    Dim temp(0 To 7) As Long
    GetDiskFreeSpace UCase(Left(DriveLetter, 1)) & ":\", temp(0), temp(1), temp(2), temp(3)
    temp(sectorsize) = temp(SectorsPerCluster) * temp(BytesPerSector) / 1024
    temp(TotalSize) = temp(sectorsize) * temp(TotalNumberOfClusters)
    temp(FreeSpace) = temp(sectorsize) * temp(NumberOfFreeClusters)
    temp(UsedSpace) = temp(sectorsize) * (temp(TotalNumberOfClusters) - temp(NumberOfFreeClusters))
    DriveDetail = temp(Detail)
End Function
Public Function Detail2Size(temp As Long) As String
    Detail2Size = GetSize(temp, " KB", " MB", " GB", " TB")
End Function
Public Function MyDrive() As String
    MyDrive = Left(App.path, 1)
End Function
Public Function iPodSize() As Long
    Dim temp As Long
    temp = DriveDetail(MyDrive, TotalSize) \ 1048576
    temp = temp \ 5 + 1
    iPodSize = temp
End Function

'Used to reset the profile
Public Sub ResetProfile()
    SaveOption "Main", "Username", UserName
    
    SaveOption "Solitaire", "Deal", "3"
    SaveOption "Solitaire", "Max Rotations", "3"
    SaveOption "Solitaire", "Scoring", "Off"
    SaveOption "Solitaire", "Timed Game", "Off"
    
    SaveOption "Puzzle", "Drop at once", "1"
    
    SetCash iPodSize * 50
    HN.setkeycontents "Games", "Brick", "0"
    HN.setkeycontents "Games", "Parachute", "0"
    HN.setkeycontents "Games", "Puzzle", "0"
End Sub

'Used to get the initial UserName
Public Function UserName() As String
    Dim strUserName As String
    strUserName = String(255, Chr$(0))
    GetUserName strUserName, 255
    UserName = Left$(strUserName, InStr(strUserName, Chr$(0)) - 1)
End Function

'Used for option/settings management
Public Sub SaveOption(Section As String, key As String, Value As String)
    If Not HN.existancesection("Options") Then HN.creationsection "Options"
    If Not HN.existancesection("Options\" & Section) Then HN.creationsection "Options\" & Section
    HN.setkeycontents "Options\" & Section, key, Value
End Sub
Public Function LoadOption(Section As String, key As String, Optional default As String) As String
    Dim temp As String
    temp = HN.getkeycontents("Options\" & Section, key, default)
    If Len(temp) = 0 Then temp = default
    LoadOption = temp
End Function
Public Function Option2Bool(Section As String, key As String, default As String) As Boolean
    Dim temp As String
    temp = HN.getkeycontents("Options\" & Section, key, default)
    Option2Bool = StrComp(temp, "On", vbTextCompare) = 0
End Function
Public Function OptionIndex(Section As String, key As String, default As String, ParamArray options() As Variant) As Long
    Dim temp As Long, tempstr As String
    tempstr = LoadOption(Section, key, default)
    OptionIndex = -1
    For temp = 0 To UBound(options)
        If StrComp(tempstr, CStr(options(temp)), vbTextCompare) = 0 Then
            OptionIndex = temp
            Exit For
        End If
    Next
End Function

'Used for high score and cash record keeping
Public Function GetHighScore(Game As String) As Long
        GetHighScore = Val(HN.getkeycontents("Games", Game, "0"))
End Function
Public Sub HighScore(Game As String, Score As Long, Optional Reward As Long = 10)
    Dim temp As Long
    If Not HN.existancesection("Games") Then HN.creationsection "Games"
    temp = GetHighScore(Game)
    If Score > temp Then
        HN.setkeycontents "Games", Game, CStr(Score)
        AddCash Reward
    End If
End Sub
Public Function GetCash() As Long
    GetCash = Val(HN.getkeycontents("Games", "TotalCash", "0"))
End Function
Public Sub SetCash(Amount As Long)
    If Not HN.existancesection("Games") Then HN.creationsection "Games"
    HN.setkeycontents "Games", "TotalCash", CStr(Amount)
End Sub
Public Sub AddCash(Amount As Long)
    SetCash GetCash + Amount
End Sub

'Used for Last Played
Public Sub SaveLastPlayed()
    Dim temp As Long
    DeleteLastPlayed
    If PlayCount > 0 Then
        HN.creationsection "Last Played"
        For temp = 0 To PlayCount
            HN.setkeycontents "Last Played", "Item", PlayList(temp), temp + 1
        Next
    End If
End Sub
Public Sub DeleteLastPlayed()
    HN.deletesection "Last Played"
End Sub
Public Function LastPlayedCount() As Long
    LastPlayedCount = HN.multikeycount("Last Played", "Item")
End Function
Public Sub LoadLastPlayed(Optional ClearList As Boolean)
    Dim temp As Long, tempstr() As String
    temp = LastPlayedCount
    If temp > 0 Then
            If ClearList Then ClearPlaylist
    End If
End Sub

'Used for media databasing
Public Function getRank(Optional ByVal filename As String) As Long
    If Len(filename) = 0 And PlayCount > 0 Then filename = PlayList(CurrItem)
    filename = Replace(filename, "\", "|")
    If Len(filename) > 0 Then rank = Val(HN.getkeycontents("Songs\" & filename, "Rank", "3")): getRank = rank
End Function
Public Sub SetRank(Optional ByVal filename As String, Optional ByVal CurrRank As Long = -1)
    If Len(filename) = 0 And PlayCount > 0 Then filename = PlayList(CurrItem)
    If CurrRank < 0 Then CurrRank = rank
    filename = Replace(filename, "\", "|")
    If Len(filename) > 0 Then HN.setkeycontents "Songs\" & filename, "Rank", CStr(CurrRank)
End Sub
Public Function SongName(Index As Long) As String
    SongName = Replace(HN.sectionatindex("Songs", Index), "|", "\")
End Function
Public Sub NewPlaylist(filename As String)
    Dim tempstr As String
    If HN.existsancekey("Playlists", filename) = False Then
        tempstr = Right(filename, Len(filename) - InStrRev(filename, "\"))
        If InStr(tempstr, ".") > 0 Then tempstr = Left(tempstr, InStrRev(tempstr, ".") - 1)
        HN.createkey "Playlists", filename, tempstr
    End If
End Sub
Public Function GetPlaylist(Index As Long) As String
    GetPlaylist = HN.GetKeyAtIndex("Playlists", Index, 1) 'HN.keyindex(HN.qualifiedsectionhandle("Playlists"), Index, 1)
End Function
Public Function SongCount() As Long
    SongCount = HN.sectioncount("Songs")
End Function
Public Function ScanFile(filename As String) As Boolean
    Dim temp As String
    temp = Replace(filename, "\", "|")
    If HN.existsancekey("Songs\" & temp, "Date") = True Then
        temp = HN.getkeycontents("Songs\" & temp, "Date")
        If HasBeenModified(filename, temp) Then
            DeleteFileDetails filename
            AddFileDetails filename
        End If
    Else
        AddFileDetails filename
    End If
    GetFileDetails filename
End Function

Public Function HasBeenModified(filename As String, OldDate As String) As Boolean
    On Error Resume Next
    Dim temp As String, temp2 As Long
    temp = FileDateTime(filename)
    temp2 = DateDiff("s", temp, OldDate)
    HasBeenModified = temp2 < 0
End Function

Public Sub AddFileDetails(filename As String)
    Dim temp As String
    temp = Replace(filename, "\", "|")
    ReadID3 filename, MP3Info
    With MP3Info
        HN.creationsection "Songs\" & temp
        HN.setkeycontents "Songs\" & temp, "Date", Now
        If Len(.sArtist) > 0 Then HN.setkeycontents "Songs\" & temp, "Artist", .sArtist
        If Len(.sAlbum) > 0 Then HN.setkeycontents "Songs\" & temp, "Album", .sAlbum
        If Len(.sGenre) > 0 Then HN.setkeycontents "Songs\" & temp, "Genre", .sGenre
        If Len(.sTitle) > 0 Then HN.setkeycontents "Songs\" & temp, "Title", .sTitle
        If .sTrack > 0 Then HN.setkeycontents "Songs\" & temp, "Track", .sTrack & Empty
        If Len(.sYear) > 0 Then HN.setkeycontents "Songs\" & temp, "Year", .sYear
        If Len(.sComment) > 0 Then HN.setkeycontents "Songs\" & temp, "Comment", .sComment
        
        If Len(.sArtist) > 0 Then
            HN.creationsection "Artists\" & .sArtist
            HN.setkeycontents "Artists\" & .sArtist, temp, Empty
        End If
        
        If Len(.sAlbum) > 0 Then
            HN.creationsection "Albums\" & .sAlbum
            HN.setkeycontents "Albums\" & .sAlbum, temp, Empty
        End If
        
        If Len(.sGenre) > 0 Then
            HN.creationsection "Genres\" & .sGenre
            HN.setkeycontents "Genres\" & .sGenre, temp, Empty
        End If
    End With
End Sub
Public Sub GetFileDetails(filename As String)
    Dim temp As String
    temp = "Songs\" & Replace(filename, "\", "|")
    With MP3Info
        .sArtist = HN.getkeycontents(temp, "Artist")
        .sAlbum = HN.getkeycontents(temp, "Album")
        .sGenre = HN.getkeycontents(temp, "Genre")
        .sTitle = HN.getkeycontents(temp, "Title")
        .sTrack = Val(HN.getkeycontents(temp, "Track"))
        .sYear = HN.getkeycontents(temp, "Year")
        .sComment = HN.getkeycontents(temp, "Comment")
    End With
End Sub
Public Sub DeleteFileDetails(filename As String)
    Dim temp As String
    temp = Replace(filename, "\", "|")
    GetFileDetails filename
    With MP3Info
        HN.deletekey "Artists\" & .sArtist, temp
        HN.deletekey "Albums\" & .sAlbum, temp
        HN.deletekey "Genres\" & .sGenre, temp
        HN.deletekey "Scanned", filename
        
        If HN.keycount("Artists\" & .sArtist) = 0 Then HN.deletesection "Artists\" & .sArtist
        If HN.keycount("Albums\" & .sAlbum) = 0 Then HN.deletesection "Albums\" & .sAlbum
        If HN.keycount("Genres\" & .sGenre) = 0 Then HN.deletesection "Genres\" & .sGenre
        
        HN.deletesection "Songs\" & temp
    End With
End Sub

Public Sub ListSection(ByVal path As String, Mnumain)
    Dim tempstr() As String, temp As Long, count As Long
    path = LCase(path)
    Select Case path
        Case "songs", "artists", "albums", "genres"
            HN.enumeratesections path, tempstr
            temp = HN.sectioncount(path)
            For count = 1 To temp
                If path = "songs" Then
                    Mnumain.NewItem GetSongTitle(tempstr(count))
                Else
                    Mnumain.NewItem tempstr(count), ">"
                End If
            Next
        Case Else
            If InStr(path, "\") = InStrRev(path, "\") Then
                HN.enumeratekeys path, tempstr
                temp = HN.keycount(path)
                For count = 1 To temp
                    Mnumain.NewItem GetSongTitle(tempstr(1, count))
                Next
            End If
    End Select
End Sub
Public Function GetSongTitle(filename As String)
    Dim temp As String
    temp = Replace(filename, "\", "|")
    GetSongTitle = HN.getkeycontents("Songs\" & temp, "Title")
End Function
Public Sub ListSections(path As String, Mnumain, Optional Rside As String)
    Dim tempstr() As String, temp As Long, count As Long
    HN.enumeratesections path, tempstr
    temp = HN.sectioncount(path)
    For count = 1 To temp
        Mnumain.NewItem tempstr(count), Rside
    Next
End Sub
Public Sub ListKeys(path As String, Mnumain, Index As Long, Optional Rside As String)
    Dim tempstr() As String, temp As Long, count As Long
    HN.enumeratekeys path, tempstr
    temp = HN.keycount(path)
    For count = 1 To temp
        Mnumain.NewItem tempstr(Index, count), Rside
    Next
End Sub
Public Function GetFilename(path As String) As String
    Dim tempstr() As String, tempstr2() As String, temp As Long, count As Long, key As String
    tempstr2 = Split(path, "\")
    tempstr2(0) = Left(tempstr2(0), Len(tempstr2(0)) - 1)
    If UBound(tempstr2) < 2 Then Exit Function
    HN.enumeratesections "Songs", tempstr
    temp = HN.sectioncount("Songs")
    
    For count = 1 To temp
        If StrComp(tempstr2(1), HN.getkeycontents("Songs\" & tempstr(count), tempstr2(0)), vbTextCompare) = 0 Then
            If StrComp(tempstr2(2), HN.getkeycontents("Songs\" & tempstr(count), "Title"), vbTextCompare) = 0 Then
                GetFilename = Replace(tempstr(count), "|", "\")
                Exit For
            End If
        End If
    Next
End Function
Public Function GetFilenameFromTitle(title As String) As String
    Dim tempstr() As String, temp As Long, count As Long, key As String
    HN.enumeratesections "Songs", tempstr
    temp = HN.sectioncount("Songs")
    
    For count = 1 To temp
        If StrComp(title, HN.getkeycontents("Songs\" & tempstr(count), "Title"), vbTextCompare) = 0 Then
            GetFilenameFromTitle = Replace(tempstr(count), "|", "\")
            Exit For
        End If
    Next
End Function

