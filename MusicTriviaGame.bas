Attribute VB_Name = "MusicTriviaGame"
Option Explicit
Private Const MaxChoices As Long = 5, MaxTime = 10
Private Songs(0 To MaxChoices) As Long, CorrectSong As Long, Currtime As Long, CorrectSongIndex As Long

Public Sub INITMusicTrivia(LCDmain, Mnumain)
    Dim temp As Long
    Bat.Visible = False
    With Bar
        .Max = MaxTime
        .Visible = True
        .Value = 0
        .Move Bat.Left, Bat.top, Bat.Width, Bat.Height
    End With
    If SongCount >= MaxChoices Then
        RandomizeSongs
        Currtime = 0
        With LCDmain
            .ClearText
            TitleBar LCDmain, "Music Trivia"
            .LCDRefresh
        End With
        With Mnumain
            .ClearItems
            .Locked = True
            .Visible = True
            .HideSelected = False
            For temp = 0 To MaxChoices
                .NewItem GetSongTitle(SongName(Songs(temp)))
            Next
            .Locked = False
            MediaClose
            MediaLoad SongName(Songs(CorrectSong))
            MediaResize 26, 20, 0, 0 '4:3 aspect ratio
            MediaPlay
        End With
    Else
        With Mnumain
            .ClearItems
            .Visible = True
            .HideSelected = True
            .NewItem "There are not enough"
            .NewItem "songs to play."
            .NewItem Empty
            .NewItem "Have more scanned"
            .NewItem "before trying again."
            .NewItem Empty
            .NewItem Empty, ":)"
            Currtime = MaxTime
        End With
    End If
End Sub
Public Sub IncrementTimer(Mnumain)
    Currtime = Currtime + 1
    If Currtime <= MaxTime Then Bar.Value = Currtime
    Select Case Currtime
        Case 4, 6, 8: RemoveRandomSong Mnumain
        Case MaxTime: doWrong Mnumain, True
    End Select
End Sub
Private Sub doWrong(Mnumain, Optional TimeExpired As Boolean)
    With Mnumain
        MediaClose
        .ClearItems
        .Locked = True
        .HideSelected = True
        If TimeExpired Then
            .NewItem "Your time expired"
        Else
            .NewItem "You were incorrect"
        End If
        .NewItem "You lost 5 credits"
        .NewItem "The correct answer is:"
        .NewItem Empty
        .NewItem GetSongTitle(SongName(Songs(CorrectSong)))
        .NewItem Empty, ":("
        .Locked = False
        AddCash -5
    End With
End Sub
Private Sub doRight(Mnumain)
    With Mnumain
        MediaClose
        .ClearItems
        .Locked = True
        .HideSelected = True
        .NewItem "You were Correct!"
        .NewItem "You won 1 credit"
        .NewItem Empty
        .NewItem GetSongTitle(SongName(Songs(CorrectSong)))
        .NewItem Empty
        .NewItem Empty, ":)"
        .Locked = False
        AddCash 1
    End With
End Sub
Public Sub CheckSong(LCDmain, Mnumain)
    With Mnumain
        If .HideSelected Then
            INITMusicTrivia LCDmain, Mnumain
        Else
            If Mnumain.SelectedItem = CorrectSongIndex Then
                doRight Mnumain
            Else
                doWrong Mnumain
            End If
            Currtime = MaxTime
        End If
    End With
End Sub
Private Function SongChoosen(Song As Long) As Boolean
    Dim temp As Long
    If Song = 0 Then
        SongChoosen = True
    Else
        For temp = 0 To MaxChoices
            If Songs(temp) = Song Then
                SongChoosen = True
                Exit Function
            End If
        Next
    End If
End Function
Private Function RandomSong() As Long
    Dim temp As Long
    Do Until Not SongChoosen(temp)
        Randomize Timer
        temp = Rnd * (SongCount - 1) + 1
    Loop
    RandomSong = temp
End Function
Private Function RandomizeSongs()
    Dim temp As Long
    For temp = 0 To MaxChoices
        Songs(temp) = 0
    Next
    For temp = 0 To MaxChoices
        Songs(temp) = RandomSong
    Next
    Randomize Timer
    CorrectSong = Rnd * MaxChoices
    CorrectSongIndex = CorrectSong
End Function
Private Sub RemoveRandomSong(MNU)
    Dim temp As Long
    temp = CorrectSongIndex
    Do Until temp <> CorrectSongIndex
        temp = Rnd * (MNU.itemcount - 1)
    Loop
    If temp < CorrectSongIndex Then CorrectSongIndex = CorrectSongIndex - 1
    MNU.RemoveItem temp
End Sub
