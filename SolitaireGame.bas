Attribute VB_Name = "SolitaireGame"
Option Explicit
Private LCDmain 'As LCD
Private MaxRotations As Long, MaxCards As Long, ScoringMethod As Long, TimedGame As Boolean, TimePassed As Long
Public SelectedPile As Long, SelectedCard As Long, Rotations As Long, BufferPile As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Sub SolitaireScore(srcPile As Long, destPile As Long)
    Select Case PileList(srcPile).Style
        Case Dealt
            If PileList(destPile).Style = Aces Then
                AddCash ScorePoints(5, 10)
            Else
                AddCash ScorePoints(0, 5)
            End If
        Case Aces:  AddCash ScorePoints(-5, -10)
        Case Cards: If PileList(destPile).Style = Aces Then AddCash ScorePoints(5, 10)
    End Select
End Sub
Private Function ScorePoints(Vegas As Long, Optional Standard As Long) As Long
    If ScoringMethod > 0 Then ScorePoints = IIf(ScoringMethod = 1, Vegas, Standard)
End Function
Private Sub UnflipACard()
    If ScoringMethod = 2 Then AddCash 5
End Sub
Private Sub CompletedRotation()
    If ScoringMethod = 2 And GetCash > 0 Then AddCash -20
End Sub
Public Sub OneSecondPassed()
    TimePassed = TimePassed + 1
    If TimePassed Mod 10 = 0 And ScoringMethod = 2 And TimedGame And GetCash > 0 Then AddCash -2
    If TimedGame Then DrawSolitaire
End Sub

Public Sub SolitaireMoveCursor(Direction As Long)
    If Rotations = 0 Or PileCount = 0 Then Exit Sub 'GameOver
        If SelectedCard = -1 Then SelectedCard = 0
    If SelectedPile = 14 Then
        SelectedCard = SelectedCard + Direction
        If SelectedCard < 0 Then SelectedCard = PileList(14).CardCount - 1
        If SelectedCard >= PileList(14).CardCount Then SelectedCard = 0
    Else
        SelectedPile = SelectedPile + Direction
        If SelectedPile < 0 Then SelectedPile = 12
        If SelectedPile > 12 Then SelectedPile = 0
        If SelectedPile = 1 Then SelectedCard = PileList(1).CardCount - 1
    End If
    If SelectedPile = 1 And PileList(1).CardCount = 0 And BufferPile <> 1 Then
        SolitaireMoveCursor Direction
    Else
        DrawSolitaire
    End If
End Sub

Public Sub InitSolitaire(lcdm) ' As LCD)
    Set LCDmain = lcdm
    Const TopRow As Long = 22
    Const MiddleRow As Long = 52
    Const BottomRow As Long = 103
    
    TimePassed = 0
    TimedGame = Option2Bool("Solitaire", "Timed Game", "Off")
    ScoringMethod = OptionIndex("Solitaire", "Scoring", "Off", "Off", "Vegas", "Standard")
    If ScoringMethod = 1 Then AddCash -52
    MaxRotations = Val(LoadOption("Solitaire", "Max Rotations", "3"))
    MaxCards = Val(LoadOption("Solitaire", "Deal", "3"))
    
    Rotations = MaxRotations
    SelectedPile = 0
    DeletePiles
    
    AddPile 5, TopRow, Deck '0 Deck
    AddPile 27, TopRow, Dealt '1 3 dealt cards
    
    AddPile 71, TopRow, Aces '2 ace pile 1
    AddPile 93, TopRow, Aces '3 ace pile 2
    AddPile 115, TopRow, Aces '4 ace pile 3
    AddPile 137, TopRow, Aces '5 ace pile 4
    
    AddPile 5, MiddleRow, Cards '6 card pile 1
    AddPile 27, MiddleRow, Cards '7 card pile 2
    AddPile 49, MiddleRow, Cards '8 card pile 3
    AddPile 71, MiddleRow, Cards '9 card pile 4
    AddPile 93, MiddleRow, Cards '10 card pile 5
    AddPile 115, MiddleRow, Cards '11 card pile 6
    AddPile 137, MiddleRow, Cards '12 card pile 7
    
    AddPile 0, 0, Deck '13 discard pile
    AddPile 5, BottomRow, Dealt '14 Selected cards
    
    ShuffleDeck PileList(0)
    
    'Deal cards to the 7 piles
    Dim temp As Long, temp2 As Long
    For temp = 1 To 7
        For temp2 = 1 To temp
            MoveCard PileList(0), PileList(0).CardCount - 1, PileList(5 + temp)
            With PileList(5 + temp)
                .CardList(.CardCount - 1).Face = temp2 = temp
            End With
        Next
    Next
    
    Deal3Cards MaxCards
    
End Sub
Private Function DegreesToRadians(ByVal Degrees As Double) As Double 'Converts Degrees to Radians.
    Const Pi As Double = 3.14159265358979
    DegreesToRadians = Degrees * (Pi / 180)
End Function
Public Function GetCardY(Factor As Long, Angle As Long) As Long
    GetCardY = Sin(DegreesToRadians(Angle)) * Factor
End Function
Public Sub DrawSolitaire()
    Dim temp As Long, IsSelected As Long, title As String
    LCDmain.ClearText
    title = "Solitaire"
    If ScoringMethod > 0 Then title = title & " ($" & GetCash & ")"
    If TimedGame Then title = title & " " & sec2time(TimePassed) 'Thank you to my MCI handler code again
    TitleBar LCDmain, title
    For temp = 0 To 12
        IsSelected = IIf(SelectedPile = temp, SelectedCard, -1)
        If SelectedPile = 1 And temp = 1 And PileList(1).CardCount = 0 Then IsSelected = 0
        DrawCardPile PileList(temp), LCDmain.CardHdc, LCDmain.hdc, PileList(temp).x, PileList(temp).y, IsSelected
    Next
    If PileList(14).CardCount > 0 Then
        DrawCardPile PileList(14), LCDmain.CardHdc, LCDmain.hdc, PileList(14).x, PileList(14).y, IIf(SelectedPile = 14, SelectedCard, -1)
    End If
    LCDmain.LCDRefresh
End Sub
Public Sub SolitaireAction()
    If PileCount > 0 Then
        If SelectedCard = -1 Then SelectedCard = 0
        Select Case SelectedPile
            Case 0: Deal3Cards MaxCards 'deal 3 cards from the deck
            Case 1, 2, 3, 4, 5: SelectCard 'dealt cards and the ace pile
            Case 6, 7, 8, 9, 10, 11, 12: SelectCards 'card piles
        End Select 'selected cards pile cant be the selected pile so ignore
        If HasWon Then WINGAME
    Else
        InitSolitaire LCDmain 'If gameover, start again
    End If
End Sub
Public Sub WINGAME()
    Static IsWinning As Boolean
    If IsWinning Then Exit Sub 'prevents multiple instances as theyd interfere
    IsWinning = True
    Dim temp As Long, temp2 As Long
    temp2 = 2
    For temp = 1 To 52
        DropCard PileList(temp2)
        temp2 = temp2 + 1
        If temp2 = 6 Then temp2 = 2
        If MenuMode <> ipod_solitaire Then
            LCDmain.ClearText
            Exit For
        End If
    Next
    DeletePiles
    Rotations = 0
    IsWinning = False
End Sub
Public Sub DropCard(Pile As CardPile)
    Dim tempcard As Card, temp As Long, temp2 As Long, Factor As Long, Angle As Long, Delay As Long
    With Pile
        If .CardCount = 0 Then Exit Sub
        tempcard = .CardList(.CardCount - 1)
        DeleteCard Pile, .CardCount - 1
        Factor = 103 - .y
        Angle = 90
        Randomize Timer
        Delay = Rnd * 100 'Randomize speed
        For temp = .x To -20 Step -(2 + Rnd * 5) 'Randomize horizontal distance between redraws
            Angle = Angle + 15
            If Angle >= 180 Then Angle = 0
            temp2 = GetCardY(Factor, Angle)
            If temp2 = 0 Then Factor = Factor * (0.25 + Rnd * 0.5) 'randomize % of energy lost per bounce (in between 50% and 75%)
            LCDmain.DrawSquare temp + 1, 104 - temp2, 18, 27, LCDmain.BackColor, True
            DrawCardStyle LCDmain.CardHdc, LCDmain.hdc, temp, 103 - temp2, fullFront, tempcard.Value, tempcard.Suite
            LCDmain.LCDRefresh
            Sleep Delay
            DoEvents
            If MenuMode <> ipod_solitaire Then Exit Sub
        Next
    End With
End Sub

Public Function HasWon() As Boolean
    Dim temp As Long, buffer As Boolean
    buffer = True
    For temp = 2 To 5
        If PileList(temp).CardCount < 13 Then buffer = False
    Next
    HasWon = buffer
End Function
    
Public Sub SelectCard()
    If PileList(14).CardCount = 0 Then 'if there are no selected cards already, select one
        BufferPile = SelectedPile 'Used to undo selection
        MoveCard PileList(SelectedPile), PileList(SelectedPile).CardCount - 1, PileList(14)
        If SelectedCard >= PileList(SelectedPile).CardCount Then SelectedCard = SelectedCard - 1
        If SelectedPile = 1 And PileList(1).CardCount = 0 Then 'SolitaireMoveCursor -1 'dealt cards is empty,
            If PileList(13).CardCount = 0 Then 'take the cursor off of it cause its empty
                SolitaireMoveCursor -1
            Else 'move one from discard pile back into the dealt cards
                MoveCard PileList(13), PileList(13).CardCount - 1, PileList(1)
            End If
        End If
        DrawSolitaire
    Else 'place them down if you can, or are clicking the undo pile
        If SelectedPile = BufferPile Then
            RemoveCardFromStack
        Else
            If CanPlaceStack Then
                SolitaireScore BufferPile, SelectedPile
                PurgeStack SelectedPile
            End If
        End If
    End If
    If SelectedCard = -1 Then SelectedCard = PileList(SelectedPile).CardCount - 1: DrawSolitaire
End Sub
Public Function CanPlaceStack() As Boolean
If PileList(14).CardCount > 0 Then 'if there is a selected card
    Select Case SelectedPile
        Case 2, 3, 4, 5 'ace piles
            'Can only place one card at a time
            If PileList(14).CardCount = 1 Then
                If PileList(SelectedPile).CardCount = 0 Then
                    'is the ace pile empty and the selected card is an ace then
                    If PileList(14).CardList(0).Value = "a" Then CanPlaceStack = True
                Else
                    If Face2Value(PileList(14).CardList(0).Value) = Face2Value(PileList(SelectedPile).CardList(PileList(SelectedPile).CardCount - 1).Value) + 1 Then
                        'if the selected card is 1 higher than the last card on the pile, and matches the suite then
                        If PileList(14).CardList(0).Suite = PileList(SelectedPile).CardList(0).Suite Then
                            CanPlaceStack = True
                        End If
                    End If
                End If
            End If
        Case 6, 7, 8, 9, 10, 11, 12 'card piles
            'can place multiple cards at a time
            If PileList(SelectedPile).CardCount = 0 Then
                'if pile is empty, and the selected card is a king then
                If PileList(14).CardList(0).Value = "k" Then CanPlaceStack = True
            Else
                'if the selected card is 1 lower than the last card on the pile, and is opposite in color then
                If GetColor(PileList(14).CardList(0).Suite) = Not GetColor(PileList(SelectedPile).CardList(PileList(SelectedPile).CardCount - 1).Suite) Then
                    If Face2Value(PileList(14).CardList(0).Value) = Face2Value(PileList(SelectedPile).CardList(PileList(SelectedPile).CardCount - 1).Value) - 1 Then
                        CanPlaceStack = True
                    End If
                End If
            End If
    End Select
End If
End Function

Public Sub SelectCards() 'GNDN
    If PileList(14).CardCount = 0 Then 'if there are no selected cards already,select as much as possible
        BufferPile = SelectedPile 'Used to undo selection
        'if the top card is face down, make it face up
        With PileList(SelectedPile)
            If .CardCount = 0 Then Exit Sub
            If .CardList(.CardCount - 1).Face = False Then
                UnflipACard
                .CardList(.CardCount - 1).Face = True
                DrawSolitaire
            Else
                Dim temp As Long, temp2 As Long
                temp = TopCard
                For temp2 = temp To .CardCount - 1
                    MoveCard PileList(SelectedPile), temp, PileList(14)
                Next
                DrawSolitaire
            End If
        End With
    Else 'place them down if you can, or are clicking the undo pile
        If SelectedPile = BufferPile Then
            PurgeStack
        Else
            If CanPlaceStack Then
                SolitaireScore BufferPile, SelectedPile
                PurgeStack SelectedPile
            End If
        End If
    End If
End Sub
Public Function TopCard() As Long
    Dim temp As Long, temp2 As Long
    With PileList(SelectedPile)
        temp2 = .CardCount - 1
        For temp = .CardCount - 2 To 0 Step -1
            'If the card is face down, or the same color as the one on top of it, or not 1 higher in value than the one on top of it then cancel
            If .CardList(temp).Face = False Then Exit For
            If GetColor(.CardList(temp).Suite) = GetColor(.CardList(temp + 1).Suite) Then Exit For
            If Face2Value(.CardList(temp).Value) - 1 <> Face2Value(.CardList(temp + 1).Value) Then Exit For
            temp2 = temp
        Next
    End With
    TopCard = temp2
End Function
Public Sub RemoveCardFromStack()
    If PileList(14).CardCount > 0 Then
        PileList(14).CardList(0).Face = True
        MoveCard PileList(14), 0, PileList(BufferPile)
        SelectedCard = PileList(SelectedPile).CardCount - 1
        DrawSolitaire
    End If
End Sub
Public Sub SelectLastCard()
    If PileList(14).CardCount > 1 Then
        Do Until PileList(14).CardCount = 1
            RemoveCardFromStack
        Loop
    End If
End Sub
Public Sub PurgeStack(Optional destPile As Long = -1)
    If destPile = -1 Then destPile = BufferPile
    Do Until PileList(14).CardCount = 0
        PileList(14).CardList(0).Face = True
        MoveCard PileList(14), 0, PileList(destPile)
    Loop
    DrawSolitaire
End Sub
Public Sub Deal3Cards(Optional Amount As Long = 3)
    Dim temp As Long, gameover As Boolean
    If PileList(14).CardCount > 0 Then Exit Sub 'Cards are selected
    Do Until PileList(1).CardCount = 0
        MoveCard PileList(1), 0, PileList(13)
    Loop
    If PileList(0).CardCount > 0 Then
        For temp = 1 To Amount
            If PileList(0).CardCount > 0 Then
                MoveCard PileList(0), PileList(0).CardCount - 1, PileList(1)
            End If
        Next
    Else
        'Full rotation
        Rotations = Rotations - 1
        If Rotations <= 0 Then
            'Game Over
            LCDmain.ClearText
            TitleBar LCDmain, "Game Over"
            SelectedPile = -1
            gameover = True
        Else
            CompletedRotation
            Do Until PileList(13).CardCount = 0
                MoveCard PileList(13), PileList(13).CardCount - 1, PileList(0)
            Loop
        End If
    End If
    If Not gameover Then DrawSolitaire
End Sub

Public Sub ForceWin()
    Static IsWinning As Boolean
    If IsWinning Then Exit Sub 'prevents multiple instances as theyd interfere
    IsWinning = True
    Dim temp As Long
    For temp = 0 To 14
        EmptyPile PileList(temp)
    Next
    For temp = 1 To 13
        AddCard PileList(2), Value2Face(temp), "<heart>"
        AddCard PileList(3), Value2Face(temp), "<diamond>"
        AddCard PileList(4), Value2Face(temp), "<club>"
        AddCard PileList(5), Value2Face(temp), "<spade>"
    Next
    SelectedPile = -1
    DrawSolitaire
    WINGAME
    IsWinning = False
End Sub
