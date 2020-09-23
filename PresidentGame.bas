Attribute VB_Name = "PresidentGame"
Option Explicit
'this is the first game to use the menu as a seperate screen

'as a note, the AI's rank is not kept track of
'if you are pres, ai 1 will give you 2 cards
'if vice pres ai 1 will give you 1 cardd
'if vice bum, you give ai 1, a single card
'if bum, you give ai 1, two cards
'simple ist it?

Private Const bufferdeck As Long = 1
Private Enum Pres_Ranks
    Bum = 0
    ViceBum = 1
    VicePresident = 2
    President = 3
End Enum
Private SelectedPile As Long, SelectedCard As Long, rank As Pres_Ranks, MNU, passes As Long, availrank As Pres_Ranks

Public Sub PRESinit(LCDmain, Mnumain, Optional ClearVars As Boolean)
    DeletePiles
    Set MNU = Mnumain
    MNU.ClearItems
    MNU.HideSelected = True
    LCDmain.ClearText
    TitleBar LCDmain, "President", PlayPause
    
    SelectedPile = 5
    SelectedCard = 0
    
    If ClearVars Then rank = President
    
    AddPile 1, 103, deck '0) Deck
    AddPile 23, 103, Horizontal  '1) current hand to be beaten
    
    AddPile 0, 0, Aces '2)AI number 1
    AddPile 0, 0, Aces '3)AI number 2
    AddPile 0, 0, Aces '4)AI number 3
    AddPile 1, 23, Dealt '5) Your hand
    AddPile 1, 73, Dealt '6) Your selected cards
        
    ShuffleDeck PileList(0)
    DealCardsMulti PileList(0), 7, 2, 3, 4, 5
    availrank = President
    passes = 0
    
    Notify "You are the " & getRank(rank)
    Select Case rank
        Case President
            AIgivecard 2   'AI gives 3 cards to you
            AIgivecard 1, 3, 4 'AI gives 1 card to someone else
            Notify "Its your turn to lead off"
        Case VicePresident
            AIgivecard 1  'AI gives 1 card to you
            AIgivecard 2, 3, 4  'AI gives 2 cards to someone else
            PRESAILeadOff 4
        Case ViceBum
            AIgivecard 1, 5, 2 'You give AI 1, 1 card
            AIgivecard 2, 3, 4 'AI gives 2 cards to someone else
            PRESAILeadOff 4
            PRESAI 2
        Case Bum
            AIgivecard 2, 5, 2 'You give AI 1, 2 cards
            AIgivecard 1, 3, 4 'AI gives 1 card to someone else
            PRESAILeadOff 2
            PRESAI 3
            PRESAI 4
    End Select
    RevealHand PileList(5)
    RevealHand PileList(1)
    MNU.Visible = True
    MNU.DrawMenu
    LCDmain.LCDRefresh
End Sub
Private Function getRank(urank As Pres_Ranks) As String
    Const ranks As String = "Bum|ViceBum|VicePresident|President"
    Dim tempstr() As String
    tempstr = Split(ranks, "|")
    getRank = tempstr(urank)
End Function
Public Sub PRESmove(LCDmain, Optional Direction As Long = 1)
If MNU.Visible Then
    MNU.selecteditem = MNU.selecteditem + Direction
Else
    SelectedCard = SelectedCard + Direction
    If SelectedCard >= PileList(SelectedPile).CardCount Or SelectedCard < 0 Or SelectedPile = 0 Then SelectNextDeck Direction
    If PileList(SelectedPile).CardCount = 0 Then SelectNextDeck Direction
    PRESdrawscreen LCDmain
End If
End Sub
Private Sub SelectNextDeck(Direction As Long)
    SelectedPile = SelectedPile + Direction
    Select Case SelectedPile        'valids are 0 5 and 6
        Case -1: SelectedPile = 6
        Case 1: SelectedPile = 5
        Case 4, 7: SelectedPile = 0
    End Select
    SelectedCard = 0
    If Direction = -1 Then SelectedCard = PileList(SelectedPile).CardCount - 1
End Sub
Public Sub PRESselect(LCDmain)
Dim buffer As Boolean
If MNU.Visible Then
    If passes = 4 Then
        EmptyPile PileList(1)
        MNU.ClearItems
        Notify "New Round!"
        Select Case rank
            Case President
                Notify "Its your turn to lead off"
            Case VicePresident
                PRESAILeadOff 4
            Case ViceBum
                PRESAILeadOff 4
                PRESAI 2
            Case Bum
                PRESAILeadOff 2
                PRESAI 3
                PRESAI 4
        End Select
    Else
        MNU.Visible = False
    End If
    passes = 0
Else
    Select Case SelectedPile
        Case 0 'is the deck (no cards in 6 is a pass)
            If PileList(bufferdeck).CardCount = 0 Then 'You are leading off, you must throw down card(s)
                If PileList(6).CardCount > 0 Then
                    DealCards PileList(6), PileList(bufferdeck), PileList(6).CardCount
                    lC.ClearText
                    lC.LCDRefresh
                    MNU.ClearItems
                    MNU.Visible = True
                    MNU.NewItem "You put down " & PileList(bufferdeck).CardCount & " " & UCase(DeckValue(bufferdeck)) & IIf(PileList(bufferdeck).CardCount <> 1, "'s", Empty)
                    DoAI LCDmain
                End If
            Else 'You are not leading off, check the deck
                'to see if you are passing, or that your cardcount matches, and the value is higher, or if you put a 2 down
                'if its a 2, make sure the deck doesnt have a 2, if it does you need to put more down
                'if you pass, you get another card
                If PileList(6).CardCount = 0 Then
                    MNU.ClearItems
                    MNU.Visible = True
                    If PileList(5).CardCount < 14 Then
                        DealCards PileList(0), PileList(5), 1
                        Notify "You were dealt a " & UCase(PileList(5).CardList(PileList(5).CardCount - 1).Value)
                    Else
                        Notify "I took mercy on you"
                        Notify "You were dealt no cards"
                    End If
                    passes = passes + 1
                    CheckEmptyDeck
                    DoAI LCDmain
                Else
                    'if the cards you put down are 2 then
                        'if the deck has a 2, then you must have more 2's than it does
                        'if not, then you can place it
                    'if not then you must have the same number of cards as the pile, unless its 2's
                    
                    buffer = (PileList(6).CardCount = PileList(1).CardCount) And (PRESValue(DeckValue(1)) < 15)
                    buffer = buffer Or (PRESValue(DeckValue(6)) = 15) And (PRESValue(DeckValue(1)) < 15)
                    buffer = buffer Or (PRESValue(DeckValue(6)) = 15) And (PRESValue(DeckValue(1)) = 15) And (PileList(6).CardCount > PileList(1).CardCount)
                    If buffer Then
                        EmptyPile PileList(1)
                        MNU.ClearItems
                        MNU.Visible = True
                        Notify "You threw down " & PileList(6).CardCount & " " & UCase(PileList(6).CardList(0).Value) & IIf(PileList(6).CardCount = 1, Empty, "s")
                        DealCards PileList(6), PileList(1), PileList(6).CardCount
                        If PileList(5).CardCount = 0 Then
                            Notify "You are out"
                            Notify "You are the " & getRank(availrank)
                            rank = availrank
                            availrank = availrank - 1
                            passes = 4
                        Else
                            MNU.Visible = True
                            DoAI LCDmain
                        End If
                    End If
                End If
            End If
        Case 5 'is non held cards (your hand)
            If CanMoveCard(PileList(5).CardList(SelectedCard).Value) Then
                MoveCard PileList(5), SelectedCard, PileList(6)
                CheckSelected 5, 6, SelectedCard
            End If
        Case 6 'is held cards (your buffer)
            MoveCard PileList(6), SelectedCard, PileList(5)
            CheckSelected 6, 5, SelectedCard
    End Select
End If
If passes = 4 Then
    MNU.NewItem "Everyone passed"
    MNU.NewItem "This round is over"
End If
If Not MNU.Visible Then PRESdrawscreen LCDmain
End Sub
Private Sub CheckSelected(Selected As Long, NextDeck As Long, Card As Long)
    If PileList(Selected).CardCount = 0 Then
        SelectedPile = NextDeck
        Card = 0
    Else
        If Card >= PileList(Selected).CardCount Then
            Card = PileList(Selected).CardCount - 1
        End If
    End If
End Sub
Private Function CanMoveCard(Value As String) As Boolean
    'Dim tempstr As String 'if pile is empty, or pile value matches selected card
    'tempstr = DeckValue(6) 'if selected card is 2, or pile value is 2 was removed
    CanMoveCard = (DeckValue(6) = Value) Or (PileList(6).CardCount = 0) ' Or (tempstr = "2") or (Value = "2")
End Function
Private Function DeckValue(deck As Long) As String
    Dim temp As Long, tempstr As String
    tempstr = "0"
    If PileList(deck).CardCount > 0 Then
        tempstr = PileList(deck).CardList(0).Value
        'If tempstr = "2" Then
        '    For temp = 1 To PileList(deck).CardCount - 1
        '        If PileList(deck).CardList(temp).Value <> "2" Then
        '            tempstr = PileList(deck).CardList(temp).Value
        '            Exit For
        '        End If
        '    Next
        'End If
    End If
    DeckValue = tempstr
End Function

Public Sub PRESdrawscreen(LCDmain)
    With LCDmain
        .ClearText
        TitleBar LCDmain, "President", PlayPause
        AUTOdrawdeck 0, LCDmain, IIf(SelectedPile = 0, 0, -1)
        AUTOdrawdeck 1, LCDmain
        AUTOdrawdeck 5, LCDmain, IIf(SelectedPile = 5, SelectedCard, -1)
        AUTOdrawdeck 6, LCDmain, IIf(SelectedPile = 6, SelectedCard, -1)
        .LCDRefresh
    End With
End Sub
Private Sub Notify(text As String, Optional text2 As String)
    MNU.NewItem text, text2
End Sub
Private Function HighestCard(pileindex As Long) As Long
    Dim temp As Long, temp2 As Long '2 A K Q J 10 9 8 7 6 5 4 3 (greatest to least)
    temp2 = 0
    If PileList(pileindex).CardCount > 0 Then
        For temp = 1 To PileList(pileindex).CardCount - 1
            If PREScardvalue(PileList(pileindex).CardList(temp).Value) > PREScardvalue(PileList(pileindex).CardList(temp2).Value) Then
                temp2 = temp
            End If
        Next
    End If
    HighestCard = temp2
End Function
Private Function PREScardvalue(cardvalue As String) As Long
    Dim temp As Long
    temp = Face2Value(cardvalue)
    If temp = 1 Then temp = 14 'ace is second highest
    If temp = 2 Then temp = 15 '2 card is highest
    PREScardvalue = temp
End Function
Private Function AIgivecard(Optional Cards As Long = 1, Optional AIPile As Long = 2, Optional destPile As Long = 5)
    Dim temp As Long
    For temp = 1 To Cards
        MoveCard PileList(AIPile), HighestCard(AIPile), PileList(destPile)
    Next
    Notify PlayerName(AIPile) & " gave " & PlayerName(destPile) & ", " & Cards & " card" & IIf(Cards = 1, Empty, "s")
End Function
Private Function PlayerName(pileindex As Long) As String
    Select Case pileindex
        Case 2, 3, 4: PlayerName = "AI" & pileindex - 1
        Case 5: PlayerName = "You"
    End Select
End Function
Private Function PRESValue(Value As String) As Long
    Select Case LCase(Value)
        Case "a": PRESValue = 14
        Case "2": PRESValue = 15
        Case Else: PRESValue = Face2Value(Value)
    End Select
End Function
Private Function unPRESvalue(Value As Long) As String
    Select Case Value
        Case 14: unPRESvalue = "a"
        Case 15: unPRESvalue = "2"
        Case Else: unPRESvalue = Value2Face(Value)
    End Select
End Function
Private Function HighestCardCount(indexpile As Long) As Long
    Dim temp As Long, temp2 As Long, temp3 As Long, temp4 As Long
    For temp = 0 To PileList(indexpile).CardCount - 1
        If PRESValue(PileList(indexpile).CardList(temp).Value) <> temp2 And PileList(indexpile).CardList(temp).Value <> "2" Then 'I dont want the AI throwing a bunch of 2's down
            temp3 = CardCount(PileList(indexpile), PileList(indexpile).CardList(temp).Value)
            If temp3 > temp4 Then
                temp4 = temp3
                temp2 = temp
            End If
        End If
    Next
    HighestCardCount = temp2
End Function
Private Sub PRESAILeadOff(Optional AIDeck As Long = 2)
    Dim temp As Long, temp2 As Long, temp3 As Long, tempstr As String
    temp = HighestCardCount(AIDeck)
    tempstr = PileList(AIDeck).CardList(temp).Value
    For temp2 = PileList(AIDeck).CardCount - 1 To 0 Step -1
        If Face2Value(PileList(AIDeck).CardList(temp2).Value) = Face2Value(tempstr) Then
            MoveCard PileList(AIDeck), temp2, PileList(bufferdeck)
            temp3 = temp3 + 1
        End If
    Next
    Notify PlayerName(AIDeck) & " lead off with " & temp3 & " " & UCase(tempstr) & IIf(temp3 = 1, Empty, "s")
End Sub
Private Sub DoAI(LCDmain) 'true is before you, false is after you
    PRESAI 2
    PRESAI 3
    PRESAI 4
End Sub
Private Function PRESAI(AIDeck As Long) As Boolean
    Dim temp As Long, Value As Long, Quantity As Long, tempQuantity As Long, tempValue As Long, pileValue As Long
    pileValue = PRESValue(DeckValue(1))
    If PileList(AIDeck).CardCount > 0 And passes < 4 Then ' The AI may have won already
        'get the lowest card value with the closest quantity to the number of cards in pile 1
        If pileValue = 15 Then 'The buffer deck is all 2's
            tempQuantity = CardCount(PileList(AIDeck), "2")
            If tempQuantity > PileList(1).CardCount Then
                Quantity = tempQuantity
                Value = 15
            End If
        Else

        For temp = 0 To PileList(AIDeck).CardCount - 1
            tempValue = PRESValue(PileList(AIDeck).CardList(temp).Value)
            If tempValue > pileValue And tempValue < 15 Then 'If the card is greater than the value of the buffer deck
                tempQuantity = CardCount(PileList(AIDeck), PileList(AIDeck).CardList(temp).Value)
                'If the buffer is empty,
                'or the quantity is greater than or equal to the card count in the buffer pile
                'and the value is smaller than the buffer card
                If (tempQuantity >= PileList(1).CardCount And ((tempValue < Value) Or (Value = 0))) Then
                    Quantity = tempQuantity
                    Value = tempValue
                End If
            End If
        Next
        End If
        AITransfer AIDeck, Face2Value(unPRESvalue(Value))
    End If
End Function
Private Sub endgame(LCDmain)
    Notify "Everyone passed"
    Notify "This round is over"
    PRESinit LCDmain, MNU, False
End Sub
Private Sub CheckEmptyDeck()
    If PileList(0).CardCount = 0 Then
        Notify "New deck shuffled"
        ShuffleDeck PileList(0)
    End If
End Sub
Private Sub AITransfer(AIDeck As Long, Value As Long)
    Dim temp As Long, temp2 As Long, temp3 As String
    If Value = 0 Then 'AI passed/couldnt find cards to use
        DealCards PileList(0), PileList(AIDeck), 1
        Notify PlayerName(AIDeck) & " passed its turn"
        passes = passes + 1
        CheckEmptyDeck
    Else 'AI must throw down (the number of cards in the buffer pile) * (the cards matching Value)
        temp2 = PileList(1).CardCount
        temp3 = Value2Face(Value)
        If temp3 = "2" Then temp2 = temp2 + 1
        EmptyPile PileList(1)
        Notify PlayerName(AIDeck) & " threw down " & temp2 & " " & UCase(temp3) & IIf(temp2 = 1, Empty, "s")
        For temp = PileList(AIDeck).CardCount - 1 To 0 Step -1
            If temp2 > 0 Then
                If PileList(AIDeck).CardList(temp).Value = temp3 Then
                    MoveCard PileList(AIDeck), temp, PileList(1)
                    temp2 = temp2 - 1
                End If
            End If
        Next
        RevealHand PileList(1)
    End If
    If PileList(AIDeck).CardCount = 0 Then
        Notify PlayerName(AIDeck) & " is out"
        availrank = availrank - 1
        If availrank = Bum Then
            Notify "You lost"
            Notify "You are the bum"
            rank = Bum
            passes = 4
        End If
    End If
End Sub
