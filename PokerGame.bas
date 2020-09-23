Attribute VB_Name = "PokerGame"
Option Explicit 'First game to use naming conventions
Private CanBet As Boolean, Yourbet As Long, SelectedDeck As Long, SelectedCard As Long, CurrMSG As String, TimesDealt As Long
Private Enum PKRHand
    Zilch = 0
    Pair = 1
    TwoPair = 2 'Requires two card values to check
    ThreeOfAKind = 3
    Straight = 4
    Flush = 5
    FullHouse = 6 'Requires two card values to check
    FourOfAKind = 7
    StraightFlush = 8
    RoyalFlush = 9 'A-K-Q-J-10, all same suit
End Enum
Private Const PlaceBet As String = "Please place your bet"
Private Const TwoOrMore As String = "Must hold 2 or more"
Private Function HandName(Hand As PKRHand) As String
    Dim temp() As String
    temp = Split("Nothing,Pair,Two Pair,Three of a Kind,Straight,Flush,Full House,Four of a Kind,Straight Flush,Royal Flush", ",")
    HandName = temp(Hand)
End Function
Private Function PKREvalHand(Pile As CardPile, Optional ByRef RetValIndex As Long, Optional ByRef RetValIndex2 As Long) As PKRHand
Dim temp As Long, temp2 As Long, temp3 As PKRHand
For temp = 0 To Pile.CardCount - 1
    For temp2 = 0 To Pile.CardCount - 1
        If temp <> temp2 Then
            Checkforhand Pile, TwoPair, temp3, temp, temp2, RetValIndex, RetValIndex2
            Checkforhand Pile, FullHouse, temp3, temp, temp2, RetValIndex, RetValIndex2
        End If
    Next
    Checkforhand Pile, Pair, temp3, temp, RetValIndex, RetValIndex2
    Checkforhand Pile, ThreeOfAKind, temp3, temp, RetValIndex, RetValIndex2
    Checkforhand Pile, Flush, temp3, temp, RetValIndex, RetValIndex2
    Checkforhand Pile, FourOfAKind, temp3, temp, RetValIndex, RetValIndex2
    Checkforhand Pile, StraightFlush, temp3, temp, RetValIndex, RetValIndex2
    Checkforhand Pile, RoyalFlush, temp3, temp, RetValIndex, RetValIndex2
Next
PKREvalHand = temp3
End Function
Private Function Checkforhand(Pile As CardPile, Hand As PKRHand, ByRef RetVal As PKRHand, Index As Long, Optional Index2 As Long, Optional ByRef RetValIndex As Long, Optional ByRef RetValIndex2 As Long)
    Dim temp As String, temp2 As String
    temp = Pile.CardList(Index).Value
    
    If Hand = TwoPair Or Hand = FullHouse Then
        temp2 = Pile.CardList(Index2).Value
    Else
        temp2 = Pile.CardList(Index).Suite
    End If
    
    If HasHand(Pile, Hand, temp, temp2) = 100 And RetVal < Hand Then
        RetVal = Hand
        RetValIndex = Index
        RetValIndex2 = Index2
    End If
End Function
Private Function HasHand(Pile As CardPile, Hand As PKRHand, Optional StartValue As String, Optional Suite As String) As Long
Dim temp As Long, temp2 As Long, temp3 As Long
Select Case Hand
    
    Case Pair 'two cards, same value
        temp2 = 2
        temp = CardCount(Pile, StartValue)
    
    Case TwoPair 'two pairs, d'uh
        temp2 = 4 'Suite is treated as the second card value
        If Suite <> StartValue Then temp = CardCount(Pile, StartValue) + CardCount(Pile, Suite)
    
    Case ThreeOfAKind 'three cards, same value
        temp2 = 3
        temp = CardCount(Pile, StartValue)
        
    Case Straight 'any five consecutive cards
        temp2 = 5
        If Face2Value(StartValue) <= 10 Then
            temp = HasCardVal(Pile, StartValue)
            For temp3 = 1 To 4
                temp = temp + HasCardVal(Pile, RelativeCard(StartValue, temp))
            Next
        End If
                
    Case Flush 'any five cards of the same suit
        temp2 = 5
        temp = CardCount(Pile, Empty, Suite)
        
    Case FullHouse 'Three-of-a-Kind and a Pair
        temp2 = 5 'Suite is treated as the second card value
        If Suite <> StartValue Then temp = CardCount(Pile, StartValue) + CardCount(Pile, Suite)
    
    Case FourOfAKind 'four cards, same value
        temp2 = 4
        temp = CardCount(Pile, StartValue)
    
    Case StraightFlush 'any five consecutive cards, all same suit
        temp2 = 5
         If Face2Value(StartValue) <= 10 Then
            temp = HasCardVal(Pile, StartValue, Suite)
            For temp3 = 1 To 4
                temp = temp + HasCardVal(Pile, RelativeCard(StartValue, temp), Suite)
            Next
        End If
        
    Case RoyalFlush 'A-K-Q-J-10, all same suit
        temp2 = 5
        temp = HasCardVal(Pile, "10", Suite)
        temp = temp + HasCardVal(Pile, "j", Suite)
        temp = temp + HasCardVal(Pile, "q", Suite)
        temp = temp + HasCardVal(Pile, "k", Suite)
        temp = temp + HasCardVal(Pile, "a", Suite)
        
End Select
HasHand = Round(temp / temp2 * 100)
If Hand = Pair Or Hand = TwoPair Then If temp < temp2 Then HasHand = 0  'I dont want the AI trying to get pairs
End Function
Private Function HasCardVal(Pile As CardPile, Value As String, Optional Suite As String) As Long
    HasCardVal = Abs(HasCard(Pile, Value, Suite))
End Function
Public Sub PKRBet(direction As Long, LCDmain)
    If CanBet Then
        Yourbet = Yourbet + direction
        If Yourbet < 0 Then Yourbet = 0
    End If
    PKRDrawScreen LCDmain
End Sub
Public Sub PKRChooseSelected(LCDmain)
    Select Case SelectedDeck
        Case 0 'Deck
            TimesDealt = TimesDealt + 1
            CanBet = False
            Select Case TimesDealt
                Case 0, 1, 2: DealNonHelds
                Case 3: PKREndround LCDmain
                Case Else: PKRInit LCDmain
            End Select
        Case 3 'not held
            MoveCard PileList(3), SelectedCard, PileList(4)
            checkforemptypile 3, 4
        Case 4 'held
            MoveCard PileList(4), SelectedCard, PileList(3)
            checkforemptypile 4, 3
    End Select
    PKRDrawScreen LCDmain
End Sub
Private Sub checkforemptypile(srcpile As Long, destpile As Long)
    If SelectedCard >= PileList(srcpile).CardCount Then SelectedCard = SelectedCard - 1
    If PileList(srcpile).CardCount = 0 Then SelectedDeck = destpile
End Sub
Private Sub DealNonHelds()
    Dim temp As Long
    temp = PileList(3).CardCount
    If temp > 3 Then
        CurrMSG = TwoOrMore
    Else
        CurrMSG = "You got " & PileList(3).CardCount & " card" & IIf(PileList(3).CardCount = 1, Empty, "s")
        EmptyPile PileList(3)
        DealCards PileList(0), PileList(3), temp
        RevealHand PileList(3)
    End If
End Sub

Private Sub PKREndround(LCDmain)
    'Move non held cards to the held card pile
    Dim temp As Long
    DealCards PileList(3), PileList(4), PileList(3).CardCount
    'DealCards PileList(1), PileList(2), PileList(3).CardCount
    temp = PKREvalHand(PileList(4))
    If temp > 0 Then
        CurrMSG = HandName(temp) & " (" & temp * Yourbet & ")"
        AddCash temp * Yourbet
    Else
        CurrMSG = "You lost " & Yourbet & " credits"
        AddCash -Yourbet
    End If
    PKRDrawScreen LCDmain
End Sub
Public Sub PKRMoveSelected(direction As Long, LCDmain)
    If TimesDealt <= 2 Then
        SelectedCard = SelectedCard + direction
        If SelectedCard >= PileList(SelectedDeck).CardCount Or SelectedCard < 0 Or SelectedDeck = 0 Then SelectNextDeck direction
        If PileList(SelectedDeck).CardCount = 0 Then SelectNextDeck direction
    Else
        SelectedDeck = 0
    End If
    PKRDrawScreen LCDmain
End Sub
Private Sub SelectNextDeck(direction As Long)
    SelectedDeck = SelectedDeck + direction
    Select Case SelectedDeck
        Case -1: SelectedDeck = 4 'went from 0 down
        Case 1: SelectedDeck = 3 'went from 0 up
        Case 2: SelectedDeck = 0 'went from 3 down
        Case 5: SelectedDeck = 0 'went from 4 up
    End Select
    SelectedCard = 0
    If direction < 0 Then SelectedCard = PileList(SelectedDeck).CardCount - 1
End Sub

Public Sub PKRDrawScreen(LCDmain)
    Dim temp As Long
    LCDmain.ClearText
    TitleBar LCDmain, "Current Bet: $" & Yourbet
    PileList(2).y = 131 - PileHeight(PileList(2))
    For temp = 0 To 4
        If temp < 1 Or temp > 2 Then
            DrawCardPile PileList(temp), LCDmain.CardHdc, LCDmain.hdc, PileList(temp).x, PileList(temp).y, IIf(SelectedDeck = temp, SelectedCard, -1)
        End If
    Next
    
    'LCDmain.DrawLine 22, 22, 1, 111
    'LCDmain.PrintText "AIs", 2, 52
    LCDmain.PrintText "Your hand", 24, 22
    LCDmain.PrintText "Held cards", 24, 62
    LCDmain.PrintText CurrMSG, 80 - StringWidth(CurrMSG) / 2, 120
    
    LCDmain.LCDRefresh
End Sub

Public Sub PKRInit(LCDmain)
    SelectedPile = 0
    TimesDealt = 0
    DeletePiles
    
    AddPile 1, 22, Deck '0) Deck
    AddPile 1, 63, Cards '1) AI's NON-held cards
    AddPile 1, 131, Cards '2) AI's held cards (must move up 131-pileheight)
    
    AddPile 24, 32, Horizontal '3) Your NON-held cards
    AddPile 24, 72, Horizontal '4) Your held cards
    
    ShuffleDeck PileList(0)
    DealCardsMulti PileList(0), 5, 1, 3
    RevealHand PileList(3)
    
    CurrMSG = PlaceBet
    CanBet = True
    
    PKRDrawScreen LCDmain
End Sub
