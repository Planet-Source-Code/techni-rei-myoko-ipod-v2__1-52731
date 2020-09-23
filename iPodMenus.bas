Attribute VB_Name = "iPodMenus"
Option Explicit
Public Enum Ipod_Mode 'this is used to determine where the controls are being sent
    ipod_menu = 0       'Menu Mode
    ipod_nowplaying = 1 'Now Playing Mode
    ipod_block = 2      'Block Game
    ipod_parachute = 3  'Parachute Game
    ipod_solitaire = 4  'Solitaire Game
    ipod_calender = 5   'Calender Mode
    ipod_clock = 6      'Clock Mode
    ipod_blackjack = 7  'Black Jack Game
    ipod_trivia = 8     'Music Trivia Game
    ipod_notes = 9      'Notes Mode
    ipod_slots = 10     'Slow Machine Game
    ipod_puzzle = 11    'Puzzle Game
    ipod_poker = 12     'Poker Game
    ipod_president = 13 'President Game
    ipod_hat = 14       'Honour Among Theives Game
    ipod_psychic = 15   'Psychic Test Game
End Enum
'Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowPos Lib "User32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Public Const SWP_NOMOVE = &H2
    Public Const SWP_NOSIZE = &H1
    'Used to set window to always be on top or not
    Public Const HWND_NOTOPMOST = -2
    Public Const HWND_TOPMOST = -1
    
Public CurrDir As String, OnTop As Boolean, rank As Long
Public DrvBox As DriveListBox, DirBox As DirListBox, FilBox As FileListBox, lC, Bat, Bar, Tim
Public PlayList() As String, PlayCount As Long, CurrItem As Long, TitleCaption As String, MenuMode As Ipod_Mode
Public Function chkpath(path As String, filename As String) As String
    chkpath = path & IIf(InStrRev(path, "\") = Len(path), Empty, "\") & filename
End Function
Public Sub setAlwaysOnTop(hwnd As Long, Optional OnTop As Boolean = True)
On Error Resume Next
If OnTop = False Then Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
If OnTop = True Then Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Public Sub Seekto(Direction As Long)
    Dim isdone As Boolean
    If MediaIsPlaying And Direction = -1 Then
        If MediaCurrentPosition > 4 Then
            MediaSeekto 0
            isdone = True
        End If
    End If
    If PlayCount > 0 And isdone = False Then
        CurrItem = CurrItem + Direction
        If CurrItem < 0 Then CurrItem = PlayCount - 1
        If CurrItem >= PlayCount Then CurrItem = 0
        PlayItem CurrItem
    End If
End Sub
Public Function isadir(filename As String) As Boolean
On Error Resume Next
If filename <> Empty Then isadir = (GetAttr(filename) And vbDirectory) = vbDirectory
End Function
Public Function fileexists(filename As String) As Boolean
    On Error Resume Next
    Dim temp As Long
    temp = FileLen(filename)
    fileexists = temp > 0
End Function
Public Sub AddPlayItem(filename As String)
    PlayCount = PlayCount + 1
    ReDim Preserve PlayList(PlayCount)
    PlayList(PlayCount - 1) = filename
End Sub
Public Sub AddPlayList(filename As String, Optional Index As Long = -1)
    Dim tempstr() As String, tempcount As Long, temp As Long
    LoadPlayList filename, tempstr, tempcount
    For temp = 0 To tempcount - 1
        AddPlayItem tempstr(temp)
    Next
    If Index > -1 Then CurrItem = PlayCount + Index
End Sub
Public Sub ErasePlayList(Index As Long, range As Long)
    Dim temp As Long
    If CurrItem >= Index And CurrItem < Index + range Then
        MediaStop
        CurrItem = Index - 1
    End If
    For temp = Index To PlayCount - range
        PlayList(temp) = PlayList(temp) + range
    Next
    PlayCount = PlayCount - range
    ReDim Preserve PlayList(PlayCount)
End Sub
Public Sub PlayItem(ByVal Index As Long)
    If PlayCount = 0 Then Exit Sub
    If Index < 0 Then
        Do Until Index >= 0
            Index = Index + PlayCount
        Loop
    End If
    If Index >= PlayCount Then Index = Index Mod PlayCount
    If CurrItem = Index Or MediaIsPlaying Then MediaSeekto 0
    CurrItem = Index
    MediaLoad PlayList(Index)
    ScanFile PlayList(Index)
    MediaResize 161, 85, 2, 21
    MediaPlay
End Sub
Public Sub ClearPlaylist()
    PlayCount = 0
    ReDim PlayList(0)
    CurrItem = 0
End Sub

Public Sub TitleBar(LCDMain, text As String, Optional Playstate As String)
    Dim temp As Long
    Static OldText As String
    If Len(text) = 0 Then text = OldText Else OldText = text
    temp = StringWidth(text)
    LCDMain.PrintText text, (LCDMain.Width - temp) / 2, 5
    LCDMain.PrintText Playstate, 13, 5
    LCDMain.DrawLine 2, 20, LCDMain.Width - 7, 1
    TitleCaption = text
End Sub
Public Sub ShowPlayList(MNU, filename As String)
Dim tempstr() As String, tempcount As Long, temp As Long
If filename = "on the go" Then
    For temp = 0 To PlayCount - 1
        MNU.NewItem afterslash(PlayList(temp))
    Next
Else
    LoadPlayList filename, tempstr, tempcount
    For temp = 0 To tempcount - 1
        MNU.NewItem afterslash(tempstr(temp))
    Next
End If
End Sub
Public Function afterslash(text As String) As String
    If InStr(text, "\") = 0 Then
        afterslash = text
    Else
        afterslash = Right(text, Len(text) - InStrRev(text, "\"))
    End If
End Function
Public Sub KillGhosts(ByRef location As String)
    Dim lef As String, rit As String 'Attempt to kill god damned effen ghost menus
    location = LCase(location)
    If Len(location) = 0 Then Exit Sub
    If InStr(location, "\") = 0 Then
        lef = location
    Else
        lef = Left(location, InStr(location, "\") - 1)
        rit = Right(location, Len(location) - InStr(location, "\"))
    End If
    Select Case lef
        Case "playlists", "browse", "last played", "settings", "extra" 'THESE ARE THE ONLY VALID LOCATIONS!!!!!!!!!!!!!!!
        Case Else: location = rit 'A GOD DAMNED EFFEN GHOST MENU WAS FOUND!!!!!!!!!!!!!!!!!!!!!!
    End Select
End Sub
Public Sub MainMenu(lcdm, MNU, Tim, Optional ByVal location As String, Optional Index As Long = -1)
    On Error Resume Next
    Dim temp As String, count As Long
    Bar.Move 120, 1620, 2295, 135
    With MNU
        .Locked = True
        .Move 30, 390, 2460, 1620

        .HideSelected = False
        
        MenuMode = ipod_menu
        lcdm.ClearText
        lcdm.DoubleBuffer = True
        .ClearItems True
        location = LCase(Trim(location))
        If Right(location, 1) = "\" Then location = Left(location, Len(location) - 1)
        
        KillGhosts location
        
        Select Case location
            Case Empty
                TitleBar lcdm, "MyPod", PlayPause
                .NewItem "Playlists", ">"
                .NewItem "Browse", ">"
                .NewItem "Last Played", ">"
                .NewItem "Settings", ">"
                .NewItem "Extra", ">"
            
            Case "playlists"
                TitleBar lcdm, "Playlists", PlayPause
                ListKeys "Playlists", MNU, 2, ">"
                .NewItem "On the Go", ">"
            
            Case "last played"
                TitleBar lcdm, "Last Played", PlayPause
                count = LastPlayedCount
                .NewItem "Songs", CStr(count)
                If count > 0 Then
                    .NewItem "Queue"
                    .NewItem "Play"
                    .NewItem "Clear"
                End If
                If PlayCount > 0 Then .NewItem "Reset"
            
            Case "browse"
                TitleBar lcdm, "Browse", PlayPause
                .NewItem "Artists", ">"
                .NewItem "Albums", ">"
                .NewItem "Songs", ">"
                .NewItem "Genres", ">"
                .NewItem "System", ">"
                
            Case "extra"
                TitleBar lcdm, "Extra", PlayPause
                .NewItem "Clock", ">"
                .NewItem "Contacts", ">"
                .NewItem "Calender", ">"
                .NewItem "Notes", ">"
                .NewItem "Games", ">"
                
                Case "extra\clock"
                    .Height = 810
                    .top = 1200
                    TitleBar lcdm, Format(Date, "MMM d YYYY"), PlayPause
                    .NewItem "Alarm Clock", ">"
                    .NewItem "Sleep & Timer", ">"
                    .NewItem "Date & Time", ">"
                    Tim.Enabled = True
            
                    Case "extra\clock\alarm clock"
                        TitleBar lcdm, "Alarm Clock", PlayPause
                        .NewItem "Alarm", "On"
                        .NewItem "Time", LoadOption("Main", "Alarm Time", "12:00 PM")
                        .NewItem "Sound"
            
                    Case "extra\clock\sleep & timer"
                        TitleBar lcdm, "Sleep & Timer"
                        .NewItem "1 minute"
                        .NewItem "2 minutes"
                        .NewItem "3 minutes"
                        .NewItem "4 minutes"
                        .NewItem "5 minutes"
                        .NewItem "10 minutes"
                        .NewItem "15 minutes"
                        .NewItem "Never"
                        
                    Case "extra\clock\date & time"
                        TitleBar lcdm, "Date & Time"
                        .NewItem "Set Timezone"
                        .NewItem "Set Time & Date"
                        .NewItem "Time", "12-hour"
                        .NewItem "Time in title", "Off"
                        
                Case "extra\calender"
                    TitleBar lcdm, "Calender", PlayPause
                    .NewItem "All"
                    .NewItem "To do", ">"
                    .NewItem "Alarms"
                
                    Case "extra\calender\to do"
                        TitleBar lcdm, "To do list", PlayPause
                        .NewItem "Item", "Is done"
                        .NewItem Empty
                        DumpKeys "Options\Todo", MNU
                        .NewItem Empty
                        .NewItem "Clear all"
                        
                Case "extra\notes"
                    TitleBar lcdm, "Notes"
                    ClearHistory
                    If direxists(chkpath(App.path, "Notes")) Then
                        DirBox.path = chkpath(App.path, "Notes")
                        DumpContents MNU, DirBox
                    End If
            
                Case "extra\games"
                    TitleBar lcdm, "Games", PlayPause
                    
                    .NewItem "Brick"
                    .NewItem "Parachute"
                    .NewItem "Solitaire"
                    .NewItem "Music Trivia"
                    
                    .NewItem "Blackjack"
                    .NewItem "Poker"
                    .NewItem "Puzzle"
                    .NewItem "Slot Machine"
                    
                    .NewItem "Honour Among Theives"
                    .NewItem "President"
                    .NewItem "Psychic Test"
                    
                    '.NewItem "Puzzle Bubble"
                    '.NewItem "Snake"
                    
                
                Case "extra\contacts"
                    TitleBar lcdm, "Contacts", PlayPause
                    ListSections "Contacts", MNU, ">"
                
            Case "settings"
                TitleBar lcdm, "Settings", PlayPause
                .NewItem "About", ">"
                .NewItem "Shuffle", LoadOption("Main", "Shuffle", "Off")
                .NewItem "Repeat", LoadOption("Main", "Repeat", "On")
                .NewItem "Clicker", LoadOption("Main", "Clicker", "On")
                .NewItem "Sleep Timer", LoadOption("Main", "Sleep Timer", "Off")
                .NewItem "Alarms", LoadOption("Main", "Alarms", "Off")
                .NewItem "Games", ">"
                
                Case "settings\about"
                    TitleBar lcdm, "About", PlayPause
                    .NewItem "Songs", SongCount
                    .NewItem "Capacity", Detail2Size(DriveDetail(MyDrive, TotalSize))
                    .NewItem "Available", Detail2Size(DriveDetail(MyDrive, FreeSpace))
                    .NewItem "Version", App.Major & "." & App.Minor & "." & App.Revision
                    .NewItem "S/N", "U" & Round(Rnd * 100000) & "TRM"
                    .NewItem "Format", "Windows"
                    .NewItem "Legal", ">"

                Case "settings\about\legal"
                    ProcessHTML "<TITLE>Legal</TITLE>iPod clone was programmed by Techni Myoko<P>Most of it is based off screenshots from Apple's website<P>Some is from my own memory of using a demo unit at FutureShop<P>And the rest is made up by me<P>It goes without saying that I have no affiliation with Apple<P>And iPod is a copyright of Apple", MNU
                
                Case "settings\games"
                    TitleBar lcdm, "Games", PlayPause
                    .NewItem "Profile", ">"
                    .NewItem "Solitaire", ">"
                    .NewItem "Puzzle", ">"
                    .NewItem "Slot Machine", ">"
    
                Case "settings\games\profile"
                    TitleBar lcdm, "Profile", PlayPause
                    .NewItem "Name", LoadOption("Main", "Username", UserName)
                    .NewItem "Cash", "$" & GetCash
                    .NewItem "Brick", GetHighScore("Brick") & " blocks"
                    .NewItem "Parachute", GetHighScore("Parachute") & " kills"
                    .NewItem "Puzzle", GetHighScore("Puzzle") & " Pts"
                    .NewItem "Reset"

                Case "settings\games\solitaire"
                    TitleBar lcdm, "Solitaire", PlayPause
                    .NewItem "Deal", LoadOption("Solitaire", "Deal", "3")
                    .NewItem "Max Rotations", LoadOption("Solitaire", "Max Rotations", "3")
                    .NewItem "Scoring", LoadOption("Solitaire", "Scoring", "Off")
                    .NewItem "Timed Game", LoadOption("Solitaire", "Timed Game", "Off")

                Case "settings\games\puzzle"
                    TitleBar lcdm, "Solitaire", PlayPause
                    .NewItem "Drop at once", LoadOption("Puzzle", "Drop at once", "1")
                    .NewItem "Board width", LoadOption("Puzzle", "Board width", "8")
                    
                Case "settings\games\slot machine"
                    TitleBar lcdm, "Slot Machine", PlayPause
                    .NewItem "Max Value", LoadOption("Slot Machine", "Max Value", "5")
                
            Case Else

                If isinSection("Playlists", location) Then
                    If location <> "playlists\on the go" Then
                        CurrPlayList = GetPlaylist(Index + 1)
                        TitleBar lcdm, afterslash(CurrPlayList), PlayPause
                        ShowPlayList MNU, CurrPlayList
                    Else
                        TitleBar lcdm, "On the Go", PlayPause
                        ShowPlayList MNU, "on the go"
                    End If
                End If
                If isinSection("Browse", location) Then
                    If isinSection("Browse\System", location) Then
                        location = Replace(location, "<dir>", Empty, , , vbTextCompare)
                        If StrComp(location, "Browse\System", vbTextCompare) <> 0 Then temp = Right(location, Len(location) - 14)
                        If Len(temp) = 0 Then
                            TitleBar lcdm, "System", PlayPause
                            For count = 0 To DrvBox.ListCount - 1
                                .NewItem "<dir>" & Left(DrvBox.List(count), 2), ">"
                            Next
                        Else
                            TitleBar lcdm, afterslash(temp), PlayPause
                            If Len(temp) = 2 Then temp = temp & "\"
                            DirBox.path = temp
                            FilBox.path = temp
                            DumpContents MNU, DirBox, "<dir>", ">"
                            FilBox.Pattern = PlayListFiles
                            DumpContents MNU, FilBox, Empty
                            FilBox.Pattern = AudioFiles & ";" & VideoFiles
                            DumpContents MNU, FilBox, Empty
                        End If
                    Else
                        location = Right(location, Len(location) - 7)
                        TitleBar lcdm, afterslash(sCase(location)), PlayPause
                        ListSection location, MNU
                    End If
                End If
        End Select
        .Locked = False
        .Pacman
    End With
    lcdm.LCDRefresh
End Sub
Public Function sCase(text As String) As String
    sCase = UCase(Left(text, 1)) & LCase(Right(text, Len(text) - 1))
End Function
Public Sub DumpContents(MNUMain, Obj, Optional Icon As String, Optional Rightside As String)
    Dim temp As Long, tempstr As String
    For temp = 0 To Obj.ListCount - 1
        tempstr = Obj.List(temp)
        If InStr(tempstr, "\") > 0 Then tempstr = Right(tempstr, Len(tempstr) - InStrRev(tempstr, "\"))
        MNUMain.NewItem Icon & tempstr, Rightside
    Next
End Sub
Public Function isinSection(current As String, Section As String) As Boolean
    isinSection = StrComp(current, Left(Section, Len(current)), vbTextCompare) = 0
End Function
Public Function GetMenu(ByVal current As String, Relative As String) As String
    If Right(current, 1) = "\" Then current = Left(current, Len(current) - 1)
    If Left(Relative, 1) = "\" Then Relative = Right(Relative, Len(Relative) - 1)
    If Relative = "..\" Then
        If Len(current) > 0 Then
            If InStrRev(current, "\") > 0 Then
                GetMenu = Left(current, InStrRev(current, "\") - 1)
            End If
        End If
    Else
        GetMenu = current & "\" & Relative
        If Len(current) = 0 Then GetMenu = Relative
        If Len(Relative) = 0 Then GetMenu = current
    End If
End Function

Public Sub Execute(frm As Form, MNU, ByVal current As String, filename As String, Optional Index As Long)
lC.DoubleBuffer = True
With MNU
    KillGhosts current
    If isinSection("extra\notes", current) Then
        MenuMode = ipod_notes
        ExecuteNote filename, MNU
    End If
    If isinSection("extra\clock\alarm clock", current) Then
        If LCase(filename) = "time" Then
            MNU.Visible = False
            Bar.Visible = False
            MenuMode = ipod_clock
            INITTime LoadOption("Main", "Alarm Time", "12:00 PM"), lC
        End If
    End If
    If isinSection("extra\calender", current) Then
        If isinSection("extra\calender\to do", current) Then
            If Len(filename) > 0 And LCase(filename) <> "item" Then
                If LCase(filename) = "clear all" Then
                    HN.deletesection "Options\Todo"
                    MainMenu lC, MNU, Tim, "Extra\Calender\To do"
                Else
                    ChangeSetting MNU, "Todo", MNU.GetItem(MNU.selecteditem, True), "Yes", "No"
                End If
            End If
        Else
            If LCase(filename) = "all" Then
                INITDate lC
                MenuMode = ipod_calender
                MNU.Visible = False
                Bar.Visible = False
                Bat.Visible = False
            Else
                If IsDate(filename) Then
                    If HasPlans(CDate(filename)) Then ExecuteDate lC, MNU, CDate(filename)
                End If
            End If
        End If
    End If
    If isinSection("Extra\games", current) Then
        MNU.Visible = False
        Bar.Visible = False
        Bat.Visible = False
        Select Case LCase(filename)
            Case "brick"
                AddCash -1
                Tim.Interval = 10
                MenuMode = ipod_block
                lC.DoubleBuffer = False
                InitializeDefaults lC
            Case "parachute"
                AddCash -1
                MenuMode = ipod_parachute
                Tim.Interval = 100
                InitParachuteVars
            Case "solitaire"
                Tim.Interval = 1000
                MenuMode = ipod_solitaire
                InitSolitaire lC
            Case "blackjack"
                MenuMode = ipod_blackjack
                InitBlackJack lC, MNU
            Case "music trivia"
                Tim.Interval = 1000
                MenuMode = ipod_trivia
                INITMusicTrivia lC, MNU
            Case "slot machine"
                MenuMode = ipod_slots
                INITSlotMachine lC
            Case "puzzle"
                MenuMode = ipod_puzzle
                PZLInit lC
            Case "poker"
                MenuMode = ipod_poker
                PKRInit lC
            Case "president"
                MenuMode = ipod_president
                PRESinit lC, MNU, True
            Case "honour among theives"
                MenuMode = ipod_hat
                HATInit lC, MNU
            Case "psychic test"
                MenuMode = ipod_psychic
                PSYInit lC, MNU
            Case Else 'Not installed yet
                MNU.Visible = True
                Bar.Visible = True
                Bat.Visible = True
        End Select
        Tim.Enabled = True
    End If
    If isinSection("Browse", current) Then
        If isinSection("Browse\System", current) Then
            current = Right(current, Len(current) - 14)
            current = GetMenu(current, filename)
        Else
            'current = current & "\" & filename
            'current = Right(current, Len(current) - 7)
            'current = GetFilename(current)
            current = GetFilenameFromTitle(filename)
        End If
        If Len(current) > 0 Then
            current = Replace(current, "<dir>", Empty, , , vbTextCompare)
            If isaPlaylist(current) Then
                CurrItem = PlayCount
                AddPlayList current
                NewPlaylist current
                PlayItem CurrItem
                MenuMode = ipod_nowplaying
            Else
                AddPlayItem current
                PlayItem PlayCount - 1
                MenuMode = ipod_nowplaying
            End If
        End If
    End If
    If isinSection("Last Played", current) Then
        Select Case LCase(filename)
            Case "queue": LoadLastPlayed
            Case "play":  LoadLastPlayed True
            Case "clear": DeleteLastPlayed
            Case "reset": SaveLastPlayed
        End Select
        MenuMode = ipod_nowplaying
    End If
    If isinSection("Settings", current) Then
        ChangeSetting MNU, "Main", "Shuffle", "On", "Off"
        ChangeSetting MNU, "Main", "Repeat", "On", "Off"
        ChangeSetting MNU, "Main", "Clicker", "On", "Off"
        ChangeSetting MNU, "Main", "Sleep Timer", "On", "Off"
        ChangeSetting MNU, "Main", "Alarms", "On", "Off"
        
        ChangeSetting MNU, "Solitaire", "Deal", "1", "2", "3"
        ChangeSetting MNU, "Solitaire", "Max Rotations", "1", "3", "5", "10", "50", "100", "500"
        ChangeSetting MNU, "Solitaire", "Scoring", "Off", "Standard", "Vegas"
        ChangeSetting MNU, "Solitaire", "Timed Game", "On", "Off"
        
        ChangeSetting MNU, "Puzzle", "Drop at once", "1", "2", "3", "4", "5"
        ChangeSetting MNU, "Puzzle", "Board width", "8", "12", "16"
        
        ChangeSetting MNU, "Slot Machine", "Max Value", "1", "3", "5", "7", "9"
        
        If isinSection("Settings\Games\Profile", current) And filename = "Reset" Then ResetProfile
    End If
    If isinSection("Playlists\On the Go", current) Then
        PlayItem Index
        MenuMode = ipod_nowplaying
    Else
        If isinSection("Playlists", current) Then
            AddPlayList CurrPlayList, Index
            PlayItem CurrItem
            MenuMode = ipod_nowplaying
        End If
    End If
End With
End Sub
Public Sub ChangeSetting(MNU, Section As String, Setting As String, ParamArray Settings() As Variant)
'This is an example of bad coding. It should've been split into seperate functions
'HOWEVER, VB wouldnt let me pass off the parameter array correctly, and it all HAD to be kept local
    Dim temp As String, temp2 As Long, Index As Long
    With MNU 'if selected item = Setting
        If StrComp(.GetItem(.selecteditem, True), Setting, vbTextCompare) = 0 Then
            temp = .GetItem(.selecteditem, False) 'Get the right side of the selected item
            Do Until StrComp(CStr(Settings(Index)), temp, vbTextCompare) = 0 Or Index > UBound(Settings)
                Index = Index + 1 'Get the index of that item from Settings Paramarray
            Loop
            Index = Index + 1 'Get the next item
            If Index > UBound(Settings) Then Index = 0
            temp = Settings(Index)
            .SetItem .selecteditem, Setting, temp 'Refresh the menu
            SaveOption Section, Setting, temp 'Save the setting
        End If
    End With
End Sub
Public Sub DoOnTop(frm As Form)
    setAlwaysOnTop frm.hwnd, OnTop
End Sub
Public Function isaPlaylist(filename As String) As Boolean
    isaPlaylist = islike(PlayListFiles, filename)
End Function
Public Function islike(filter As String, ByVal expression As String) As Boolean
Dim tempstr() As String, count As Long
tempstr = Split(LCase(filter), ";")
expression = LCase(expression)
islike = False
For count = 0 To UBound(tempstr)
    If expression Like tempstr(count) Then islike = True: Exit For
Next
End Function
Public Function Truncate(ByVal text As String, Width As Long, Optional append As String = "...") As String
    If StringWidth(text) <= Width Then
        Truncate = text
    Else
        Do Until StringWidth(text & append) <= Width
            text = Left(text, Len(text) - 1)
        Loop
        Truncate = text & append
    End If
End Function

Public Function PlayPause() As String
    If MediaIsPlaying Then PlayPause = "<play>"
    If MediaIsPaused Then PlayPause = "<pause>"
End Function
Public Function findXY(x As Single, y As Single, Distance As Single, Angle As Double, Optional isx As Boolean = True) As Single
    If isx = True Then findXY = x + Sin(Angle) * Distance Else findXY = y + Cos(Angle) * Distance
End Function
Public Function rad2deg(Radians As Double) As Double
    rad2deg = Radians * 180
End Function
Public Sub MoveForm(frm As Form)
    If frm.Left < 0 Then frm.Left = 0
    If frm.top < 0 Then frm.top = 0
    If frm.Left + frm.Width > Screen.Width Then frm.Left = Screen.Width - frm.Width
    If frm.top + frm.Height > Screen.Height Then frm.top = Screen.Height - frm.Height
End Sub
Public Function direxists(directory As String) As Boolean
On Error Resume Next
direxists = Len(dir(directory, vbDirectory + vbHidden)) > 0
End Function
Public Function long2text(Index As Long, ParamArray texts() As Variant) As String
    long2text = CStr(texts(Index))
End Function
