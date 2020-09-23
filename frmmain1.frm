VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "MyPod"
   ClientHeight    =   5370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   Icon            =   "frmmain1.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   Picture         =   "frmmain1.frx":0E42
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   247
   ShowInTaskbar   =   0   'False
   Begin IPod.iPodMenu Mnumain 
      Height          =   1350
      Left            =   720
      TabIndex        =   4
      Top             =   960
      Width           =   2295
      _extentx        =   4048
      _extenty        =   2381
      direction       =   -1  'True
      interval        =   0
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   2520
      Width           =   600
      _extentx        =   1058
      _extenty        =   1058
   End
   Begin IPod.ThumbWheel ThWmain 
      Height          =   1800
      Left            =   960
      TabIndex        =   2
      Top             =   3240
      Width           =   1800
      _extentx        =   3175
      _extenty        =   3175
      size            =   -1  'True
   End
   Begin IPod.LCD LCDmain 
      Height          =   2055
      Left            =   600
      TabIndex        =   3
      Top             =   360
      Width           =   2535
      _extentx        =   4471
      _extenty        =   3625
      Begin IPod.BatteryLevel BatMain 
         Height          =   150
         Left            =   2160
         TabIndex        =   5
         Top             =   100
         Width           =   300
         _extentx        =   529
         _extenty        =   265
      End
      Begin IPod.StatusBar barmain 
         Height          =   135
         Left            =   120
         Top             =   1620
         Width           =   2295
         _extentx        =   4048
         _extenty        =   238
         max             =   360
         value           =   0
      End
   End
   Begin VB.FileListBox Filmain 
      Height          =   285
      Left            =   840
      Pattern         =   "*.wav;*.mp3;*.wma;*.wax;*.mid;*.midi;*.rmi;*.au;*.snd;*.aif;*.aifc;*.aiff"
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DirListBox Dirmain 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DriveListBox Drvmain 
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   270
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   2520
      Width           =   600
      _extentx        =   1058
      _extenty        =   1058
      image           =   1
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   2
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   600
      _extentx        =   1058
      _extenty        =   1058
      image           =   2
      interval        =   2000
   End
   Begin IPod.Button btnmain 
      Height          =   600
      Index           =   3
      Left            =   2640
      TabIndex        =   11
      Top             =   2520
      Width           =   600
      _extentx        =   1058
      _extenty        =   1058
      image           =   3
   End
   Begin VB.Timer TimerMain 
      Interval        =   800
      Left            =   840
      Top             =   1320
   End
   Begin VB.PictureBox PicParachute 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1320
      Picture         =   "frmmain1.frx":2262
      ScaleHeight     =   720
      ScaleWidth      =   1050
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin IPod.Hini Hinmain 
      Left            =   1800
      Top             =   1320
      _extentx        =   847
      _extenty        =   847
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents systray As clsSysTray
Attribute systray.VB_VarHelpID = -1
Private doseek As Boolean
Const iPod_Green As Long = &HC8DDC1
Const MobilePhile_Blue As Long = 13514752
Public Sub Disable()
    'Disables thumbwheel support, cause it interferes with the IDE (I think)
    ThWmain.UseWheel = False
End Sub
Private Sub BatMain_Click()
Form_Unload 0
End Sub

Private Sub btnmain_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub btnmain_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Form_KeyUp KeyCode, Shift
End Sub

Public Sub btnmain_MouseClick(Index As Integer)
    ThWmain.UseWheel = False
    KillGhosts CurrDir
    Select Case Index
        Case 0
            Select Case MenuMode
                Case ipod_president, ipod_psychic 'GNDN
                Case ipod_slots
                    AddBet -1, LCDMain
                Case ipod_notes
                    NotesHistory -1, Mnumain
                Case ipod_block, ipod_parachute, ipod_trivia, ipod_puzzle
                    TimerMain.Interval = TimerMain.Interval + 50
                    If TimerMain.Interval > 3000 Then TimerMain.Interval = 3000
                Case ipod_solitaire
                    RemoveCardFromStack
                Case ipod_poker
                    PKRBet -1, LCDMain
                Case Else
                    PlayItem (CurrItem - 1)
            End Select
        Case 1
            Select Case MenuMode
                Case ipod_notes
                    MainMenu LCDMain, Mnumain, TimerMain, "Extra\Notes"
                Case ipod_menu
                    If Len(CurrDir) = 0 Then
                        MenuMode = ipod_nowplaying
                        TimerMain.Enabled = True
                        Mnumain.Visible = False
                    Else
                        Dim temp As String, temp2 As String
                        temp = CurrDir
                        temp2 = Right(temp, Len(temp) - InStrRev(temp, "\"))
                        CurrDir = GetMenu(CurrDir, "..\")
                        If StrComp(CurrDir, temp, vbTextCompare) <> 0 Then
                            Mnumain.Direction = False
                            MainMenu LCDMain, Mnumain, TimerMain, CurrDir
                            Mnumain.SetSelectedItem temp2
                            Mnumain.Direction = True
                        End If
                    End If
                Case Else
                        LCDMain.ClearText
                        If MenuMode = ipod_clock Then
                            If isinSection("extra\clock\alarm clock", CurrDir) Then
                                SaveOption "Main", "Alarm Time", ActualTime(Currtime) & " " & IIf(isAMorPM, "PM", "AM")
                                Mnumain.SetItem 1, "Time", ActualTime(Currtime) & " " & IIf(isAMorPM, "PM", "AM")
                            End If
                        End If
                        
                        If InStr(CurrDir, "\") > 0 Then
                            TitleBar LCDMain, Right(CurrDir, Len(CurrDir) - InStrRev(CurrDir, "\")), PlayPause
                        Else
                            TitleBar LCDMain, CurrDir, PlayPause
                        End If
                        
                        Mnumain.Visible = True
                        barmain.Visible = True
                        BatMain.Visible = True
                        
                        TimerMain.Enabled = False
                        MenuMode = ipod_menu
                        
                        Select Case LCase(CurrDir)
                            Case "extra\calender": Execute Me, Mnumain, "extra\calender", "all"
                            Case "extra\games": MainMenu LCDMain, Mnumain, TimerMain, "Extra\Games"
                        End Select
                        
                        LCDMain.LCDRefresh
            End Select
        Case 2
            Select Case MenuMode
                Case ipod_trivia, ipod_notes, ipod_poker, ipod_president, ipod_psychic 'GNDN
                Case ipod_block, ipod_parachute, ipod_puzzle
                    TimerMain.Enabled = Not TimerMain.Enabled
                Case ipod_solitaire
                    ForceWin
                Case ipod_slots
                    RandomizeSlots LCDMain
                Case ipod_president
                    Mnumain.Visible = Not Mnumain.Visible
                Case Else
                    MediaPlay
                    TimerMain.Enabled = True
                    Mnumain.Visible = False
                    MenuMode = ipod_nowplaying
                    If MediaIsPaused Then
                        TimerMain_Timer
                        TimerMain.Enabled = False
                    End If
            End Select
        Case 3
            Select Case MenuMode
                Case ipod_president, ipod_psychic 'GNDN
                Case ipod_slots
                    AddBet 1, LCDMain
                Case ipod_notes
                    NotesHistory 1, Mnumain
                Case ipod_block, ipod_parachute, ipod_trivia, ipod_puzzle
                    If TimerMain.Interval > 60 Then TimerMain.Interval = TimerMain.Interval - 50
                Case ipod_solitaire
                    SelectLastCard
                Case ipod_poker
                    PKRBet 1, LCDMain
                Case Else
                    PlayItem CurrItem + 1
            End Select
    End Select
ThWmain.UseWheel = True
End Sub

Private Sub btnmain_MouseStillDown(Index As Integer, x As Single, y As Single)
If Index = 2 Then BatMain_Click
End Sub

Private Sub btnmain_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
DoEvents
End Sub
Public Sub MoveMNUmain()
    'MNUmain is not a child of LCDmain by default because:
    'Randomly, VB will choose to think LCDmain is NOT a container
    'And will fail to load the entire form, cause it doesnt want
    'to make mnumain a child of it
    
    'SetParent Mnumain.hWnd, LCDmain.hWnd 'This method prevents the creation of autoredraw images
    'Set Mnumain.Parent = LCDmain 'This method fails to work
    
    Set Mnumain.Container = LCDMain
    Mnumain.Move 30, 390, 2460, 1620
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 17, 82: btnmain_MouseClick 0
    Case 8, 46, 77: btnmain_MouseClick 1
    Case 80: btnmain_MouseClick 2
    Case 96, 70: btnmain_MouseClick 3
    Case 13, 32: THWmain_ThumbClick
    Case 37, 38: THWmain_PodChangeCounterClockWise 15
    Case 39, 40: THWmain_PodChangeClockWise 15
End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 37, 38, 39, 40: ThWmain_PodUp 0, Shift, 0
End Select
End Sub

Private Sub Form_Load()
    Dim temp As String
    If Me.Picture <> 0 Then Call SetAutoRgn(Me)
    MoveMNUmain
    Set HN = Hinmain
    Set Tim = TimerMain
    Set lC = LCDMain
    Set Bar = barmain
    Set Bat = BatMain
    Set DrvBox = Me.Drvmain
    Set DirBox = Me.Dirmain
    Set FilBox = Me.Filmain

    Set systray = New clsSysTray
    Set systray.SourceWindow = Me
    systray.Icon = Me.Icon
    systray.ToolTip = Me.Caption
    systray.IconInSysTray
    
    MediaContainersHwnd LCDMain.LCDhwnd
    ALIAS = "MyPod"
    
    setAlwaysOnTop Me.hwnd
    LoadSettings
    
    If Len(command) = 0 Then
        MainMenu LCDMain, Mnumain, TimerMain
    Else
        temp = command
        If Left(temp, 1) = """" Then temp = Right(temp, Len(temp) - 1)
        If Right(temp, 1) = """" Then temp = Left(temp, Len(temp) - 1)
        AddPlayItem temp
        PlayItem PlayCount - 1
        MenuMode = ipod_nowplaying
        NowPlaying
    End If
    
    ThWmain.UseWheel = True
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    dragform Me.hwnd
    MoveForm Me
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'Effect 1 list 7 filenames
Dim temp As Integer, temp2 As Long
temp2 = PlayCount
If Data.Files.count > 0 Then
For temp = 1 To Data.Files.count
    If (GetAttr(Data.Files(temp)) And vbDirectory) <> vbDirectory Then
        AddPlayItem Data.Files.Item(temp)
    Else
        AddFolder Data.Files.Item(temp)
    End If
    If temp = 1 Then
        PlayItem temp2
        MenuMode = ipod_nowplaying
        NowPlaying
    End If
Next
End If
End Sub
Public Sub AddFolder(path As String)
    Dim temp As Long
    Filmain.path = path
    For temp = 0 To Filmain.ListCount - 1
        AddPlayItem chkpath(path, Filmain.List(temp))
    Next
End Sub

Private Sub Form_Resize()
    'Forces a doublebuffer refresh
    LCDMain.LCDRefresh
    Mnumain.DrawMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
    systray.RemoveFromSysTray
    Set systray = Nothing
    MediaClose
    SaveSettings
    End
End Sub

Private Sub LCDmain_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub LCDmain_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub LCDmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub Mnumain_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub Mnumain_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub mnumain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub


Private Sub ThWmain_KeyDown(KeyCode As Integer, Shift As Integer)
Form_KeyDown KeyCode, Shift
End Sub

Private Sub ThWmain_KeyUp(KeyCode As Integer, Shift As Integer)
Form_KeyUp KeyCode, Shift
End Sub

Private Sub ThWmain_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Form_OLEDragDrop Data, Effect, Button, Shift, x, y
End Sub

Private Sub THWmain_PodChangeClockWise(Angle As Long)
    Select Case MenuMode
        Case ipod_menu, ipod_blackjack, ipod_trivia, ipod_notes, ipod_hat, ipod_psychic: Mnumain.selecteditem = Mnumain.selecteditem + 1
        Case ipod_nowplaying
            If barmain.Visible Then
                barmain.Value = barmain.Value + Angle: doseek = True
            Else
                If rank < 5 Then
                    rank = rank + 1
                    DrawNowPlaying
                End If
            End If
        Case ipod_block: MovePaddle Angle / AngleFactor, 1
        Case ipod_parachute: MoveTurret -Angle
        Case ipod_solitaire: SolitaireMoveCursor 1
        Case ipod_calender: SeekDate 1, LCDMain
        Case ipod_clock: MoveClock Angle, LCDMain
        Case ipod_puzzle: PZLMoveCursor 1, LCDMain
        Case ipod_poker: PKRMoveSelected 1, LCDMain
        Case ipod_president: PRESmove LCDMain, 1
    End Select
End Sub

Private Sub THWmain_PodChangeCounterClockWise(Angle As Long)
    Select Case MenuMode
        Case ipod_menu, ipod_blackjack, ipod_trivia, ipod_notes, ipod_hat, ipod_psychic: Mnumain.selecteditem = Mnumain.selecteditem - 1
        Case ipod_nowplaying
            If barmain.Visible Then
                barmain.Value = barmain.Value - Angle: doseek = True
            Else
                If rank > 0 Then
                    rank = rank - 1
                    DrawNowPlaying
                End If
            End If
        Case ipod_block: MovePaddle Angle / AngleFactor, -1
        Case ipod_parachute: MoveTurret Angle
        Case ipod_solitaire: SolitaireMoveCursor -1
        Case ipod_calender: SeekDate -1, LCDMain
        Case ipod_clock: MoveClock -Angle, LCDMain
        Case ipod_puzzle: PZLMoveCursor -1, LCDMain
        Case ipod_poker: PKRMoveSelected -1, LCDMain
        Case ipod_president: PRESmove LCDMain, -1
    End Select
End Sub

Private Sub ThWmain_PodUp(Button As Integer, Shift As Integer, Angle As Long)
    If MenuMode = ipod_nowplaying And doseek And barmain.Visible Then MediaSeekto barmain.Value
    doseek = False
End Sub

Private Sub THWmain_ThumbClick()
    Dim temp As String, go As Boolean
    ThWmain.UseWheel = False 'Disable its effects while the menu is in operation
    Select Case MenuMode
        Case ipod_menu
            temp = Mnumain.GetItem(Mnumain.selecteditem, True)
            go = Mnumain.GetItem(Mnumain.selecteditem, False) = ">"
            If go Then
                CurrDir = GetMenu(CurrDir, temp)
                MainMenu LCDMain, Mnumain, TimerMain, CurrDir, Mnumain.selecteditem
            Else
                Execute Me, Mnumain, CurrDir, temp, Mnumain.selecteditem
                NowPlaying
            End If
        
        Case ipod_nowplaying
            barmain.Visible = Not barmain.Visible
            If barmain.Visible Then SetRank Else getRank
            DrawNowPlaying
            
        Case ipod_block:        StartGame
        Case ipod_parachute:    CreateBullet
        Case ipod_solitaire:    SolitaireAction
        Case ipod_calender:     Execute Me, Mnumain, CurrDir, CStr(DateValue(CurrDate))
        Case ipod_blackjack:    BlackJackExecute Mnumain.GetItem(Mnumain.selecteditem, True), LCDMain, Mnumain
        Case ipod_trivia:       CheckSong LCDMain, Mnumain
        Case ipod_notes:        ExecuteLink Mnumain
        Case ipod_slots:        RandomizeSlots LCDMain
        Case ipod_puzzle:       PZLSwitch LCDMain
        Case ipod_poker:        PKRChooseSelected LCDMain
        Case ipod_president:    PRESselect LCDMain
        Case ipod_hat:          HATSelect LCDMain, Mnumain.GetItem(Mnumain.selecteditem, True), Mnumain
        Case ipod_psychic:      PSYSelect LCDMain, Mnumain.GetItem(Mnumain.selecteditem, True), Mnumain
    End Select
    
    If MenuMode <> ipod_block And MenuMode <> ipod_parachute Then TimerMain.Interval = 800
    ThWmain.UseWheel = True
End Sub
Public Sub NowPlaying()
        If MenuMode = ipod_nowplaying Then
            Mnumain.Visible = False
            TimerMain.Enabled = True
        End If
End Sub
Private Sub TimerMain_Timer()
Dim temp As String
BatMain.Percent = BatMain.BatteryPercent
If isReady And MenuMode = ipod_nowplaying Then
    DrawNowPlaying
    If MediaTimeRemaining = 0 Then
        If Option2Bool("Main", "Repeat", "On") Then PlayItem CurrItem + 1
    End If
End If
If MenuMode = ipod_menu And StrComp(CurrDir, "Extra\Clock", vbTextCompare) = 0 Then
    LCDMain.ClearText
    temp = Format(Date, "MMM d YYYY")
    TitleBar LCDMain, temp, PlayPause
    temp = GetTime
    LCDMain.PrintText temp, (169 - StringWidth(temp)) / 2, 37
    LCDMain.DrawLine 2, 78, LCDMain.Width - 7, 1
    LCDMain.LCDRefresh
End If
If MenuMode = ipod_block Then MovePuck
If MenuMode = ipod_solitaire Then OneSecondPassed
If MenuMode = ipod_parachute Then
    MoveParachuters
    MoveBullets
    DrawParachuteScreen Me.PicParachute.hdc, LCDMain
End If
If MenuMode = ipod_trivia Then IncrementTimer Mnumain
If MenuMode = ipod_puzzle Then PZLDrawScreen LCDMain
End Sub
Public Sub DrawRank(rank As Long, Total As Long, ByVal x As Long, y As Long, WhiteSpace As Long)
    Dim temp As Long
    For temp = 1 To Total
        LCDMain.DrawStar rank >= temp, x, y
        x = x + WhiteSpace + RankWidth(1, 0)
    Next
End Sub
Public Function RankWidth(Total As Long, WhiteSpace As Long) As Long
    RankWidth = 29 * Total + (WhiteSpace - 1) * Total
End Function

Public Sub DrawNowPlaying()
Dim temp As String
    Mnumain.Visible = False
    
    LCDMain.ClearText
    TitleBar LCDMain, "Now Playing", PlayPause
    
    If Not ThWmain.MouseIsDown Then
        barmain.Max = MediaDuration
        barmain.Value = MediaCurrentPosition
    End If
    If Not barmain.Visible Then
        DrawRank rank, 5, 10, 90, 1
    End If
    
    temp = sec2time(MediaCurrentPosition)
    LCDMain.PrintText temp, 8, 120
    
    temp = "-" & sec2time(MediaTimeRemaining)
    LCDMain.PrintText temp, 160 - StringWidth(temp), 120
    
    temp = CurrItem + 1 & " of " & PlayCount
    If PlayCount = 0 Then
        temp = "0 of 0"
        If MediaIsPlaying Then temp = "Ghost file"
    End If
    LCDMain.PrintText temp, 4, 25
    
    If Option2Bool("Main", "Shuffle", "Off") Then LCDMain.PrintText "<shuffle>", 130, 25
    If Option2Bool("Main", "Repeat", "On") Then LCDMain.PrintText "<repeat>", 145, 25
    
    temp = Truncate(MP3Info.sTitle, 165)
    If Len(temp) = 0 Then temp = Truncate(Right(MP3Info.sFilename, Len(MP3Info.sFilename) - InStrRev(MP3Info.sFilename, "\")), 165)
    LCDMain.PrintText temp, (169 - StringWidth(temp)) / 2, 46
    
    temp = Truncate(MP3Info.sArtist, 165)
    LCDMain.PrintText temp, (169 - StringWidth(temp)) / 2, 64
    
    temp = Truncate(MP3Info.sAlbum, 165)
    LCDMain.PrintText temp, (169 - StringWidth(temp)) / 2, 82
    
    LCDMain.LCDRefresh
End Sub

Public Sub LoadSettings()
    Hinmain.loadfile chkpath(App.path, "iPod.hini")
    Me.Left = Val(LoadOption("Main", "Main Left", 0))
    Me.top = Val(LoadOption("Main", "Main Top", 0))
    frmremote.Left = Val(LoadOption("Main", "Remote Left", Screen.Width - frmremote.Width))
    frmremote.top = Val(LoadOption("Main", "Remote Top", Screen.Height - frmremote.Height))
    MoveForm Me
    MoveForm frmremote
End Sub

Public Sub SaveSettings()
    SaveLastPlayed
    SaveOption "Main", "Username", LoadOption("Main", "Username", UserName)
    SaveOption "Main", "Main Left", Me.Left
    SaveOption "Main", "Main Top", Me.top
    SaveOption "Main", "Remote Left", frmremote.Left
    SaveOption "Main", "Remote Top", frmremote.top
    Hinmain.savefile chkpath(App.path, "iPod.hini")
End Sub

Private Sub systray_LButtonUp()
On Error Resume Next
    App.TaskVisible = Not App.TaskVisible
    Me.Visible = App.TaskVisible
End Sub

Private Sub systray_RButtonUp()
    frmremote.Visible = Not frmremote.Visible
End Sub
Private Sub systray_LButtonDblClk()
    systray.IconInSysTray
End Sub

Private Sub systray_RButtonDblClk()
    systray.IconInSysTray
End Sub
