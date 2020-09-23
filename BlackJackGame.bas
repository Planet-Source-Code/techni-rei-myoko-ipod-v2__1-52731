Attribute VB_Name = "BlackJackGame"
Option Explicit
Private LCDmain, HasStayed As Boolean, Currbid As Long

Public Sub InitBlackJack(lcdm, MNU)
    DeletePiles
    Set LCDmain = lcdm
    HasStayed = False
    
    With MNU
        .ClearItems
        .top = 1470
        .Height = 600
        .Left = 1800
        .Width = 700
        .Visible = True
        lcdm.ClearText
        TitleBar lcdm, "Place your bid"
        lcdm.PrintText "You have: $" & GetCash, 30, 30
        
        lcdm.DrawDealer 0, 21
        
        .NewItem "$1"
        .NewItem "$2"
        .NewItem "$5"
        .NewItem "$10"
        .NewItem "$20"
        .NewItem "$50"
        .NewItem "$100"
        .NewItem "$200"
        .NewItem "$500"
        
        If Currbid > 0 Then MNU.SetSelectedItem "$" & Currbid
        
        lcdm.LCDRefresh
    End With
    
    AddPile 0, 0, Dealt 'Deck
    AddPile 105, 45, Dealt 'Your hand
    AddPile 0, 0, Dealt 'AI's hand
    
    ShuffleDeck PileList(0)
    DealCards PileList(0), PileList(1), 2
    DealCards PileList(0), PileList(2), 2
    
    'DrawBlackJack lcdm, MNU
End Sub
Public Function AIBlackJack(Optional risk As Long = 15) As Boolean
    Dim temp As Long
    temp = BestPileValue(PileList(2))
    If temp <= risk Then
        DealCards PileList(0), PileList(2)
        AIBlackJack = True
        If HasStayed Then AIBlackJack = AIBlackJack(risk)
    End If
End Function
Public Sub BlackJackExecute(ByVal text As String, lcdm, MNU)
Dim temp As Long, temp2 As Boolean
text = LCase(text)
If Left(text, 1) = "$" Then text = Right(text, Len(text) - 1)

If text = "hit" Or text = "stay" Then
    If text = "hit" Then
        DealCards PileList(0), PileList(1)
        If BestPileValue(PileList(1)) >= 21 Then text = "stay"
    End If
    If text = "stay" Then
        HasStayed = True
        MNU.Visible = False
    End If
    temp = CheckWinner(temp2)
    
    If temp > 0 Then
        EndGame IIf(temp = 1, "You win!", "You lost"), lcdm, MNU
        AddCash IIf(temp = 1, Currbid * 3, -Currbid)
        Exit Sub
    Else
        If Not temp2 And HasStayed Then
            EndGame "You Tied", lcdm, MNU
            AddCash Currbid
            Exit Sub
        End If
    End If
    
    DrawBlackJack lcdm, MNU
Else
    If text = "start" Then InitBlackJack lcdm, MNU
    If text = "quit" Then MainMenu lcdm, MNU, Tim, "Extra\Games"
    If IsNumeric(text) Then
        Currbid = Val(text)
        
        With MNU
            .ClearItems
            .NewItem "Hit"
            .NewItem "Stay"
        End With
        
        DrawBlackJack lcdm, MNU
    End If
End If
End Sub
Private Sub EndGame(ByVal text As String, lcdm, MNU)
    lcdm.ClearText
    TitleBar lcdm, text
    
    lcdm.PrintText "Your hand: " & BestPileValue(PileList(1)), 10, 25
    DrawCardPile PileList(1), lcdm.CardHdc, lcdm.hdc, 10, 40
    lcdm.PrintText "AI's hand: " & BestPileValue(PileList(2)), 10, 74
    DrawCardPile PileList(2), lcdm.CardHdc, lcdm.hdc, 10, 89
    
    MNU.Locked = True
    MNU.ClearItems
    MNU.NewItem "Start"
    MNU.NewItem "Quit"
    lcdm.LCDRefresh
    MNU.Visible = True
    MNU.Locked = False
    MNU.DrawMenu
End Sub

Public Function CheckWinner(temp As Boolean)
temp = AIBlackJack
If Not temp And HasStayed Then
    Dim Yours As Long, AIs As Long
    Yours = BestPileValue(PileList(1))
    AIs = BestPileValue(PileList(2))
    
    'MsgBox "AI: " & AIs & vbNewLine & "Yours: " & Yours
    
    If Yours = AIs Then CheckWinner = 0: Exit Function
    
    If Yours = 21 Then CheckWinner = 1
    If AIs = 21 Then CheckWinner = 2
    
    If Yours <= 21 And Yours > AIs Then CheckWinner = 1
    If AIs <= 21 And AIs > Yours Then CheckWinner = 2
    
    If AIs > 21 Then CheckWinner = 1
    If Yours > 21 Then CheckWinner = 2
    
    If Yours > 21 And AIs > 21 Then CheckWinner = 0
    If Yours = AIs Then CheckWinner = 0
End If
End Function
Public Sub DrawBlackJack(lcdm, MNU)
    With lcdm
        .ClearText
        MNU.Locked = True
        .DrawDealer 0, 21
        TitleBar lcdm, "Black Jack"
        .PrintText "Your hand", 105, 30
        .PrintText "Value: " & BestPileValue(PileList(1)), 105, 80
        DrawCardPile PileList(1), lcdm.CardHdc, lcdm.hdc, PileList(1).x, PileList(1).y
        .LCDRefresh
        MNU.Locked = False
        MNU.DrawMenu
        DoEvents
    End With
End Sub

Public Function BasePileValue(Pile As CardPile) As Long
    Dim temp As Long, temp2 As Long, temp3 As String, temp4 As Long
    If Pile.CardCount > 0 Then
        For temp = 0 To Pile.CardCount - 1
            If Pile.CardList(temp).value <> "a" Then
                temp3 = Pile.CardList(temp).value
                temp4 = IIf(IsNumeric(temp3), Val(temp3), 10)
                temp2 = temp2 + temp4
            End If
        Next
    End If
    BasePileValue = temp2
End Function
Public Function BestPileValue(Pile As CardPile) As Long
    Dim AceCount As Long, BaseValue As Long, temp As Long, temp3 As Long
    BaseValue = BasePileValue(Pile)
    AceCount = CardCount(Pile, "a") 'Also base ace value
    BestPileValue = BaseValue
    For temp = 0 To AceCount
        temp3 = BaseValue + AceCount + temp * 10
        If temp3 <= 21 Then BestPileValue = temp3
    Next
End Function
