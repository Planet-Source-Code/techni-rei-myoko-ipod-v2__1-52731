Attribute VB_Name = "PsychicGame"
Option Explicit
Private Total As Long, Correct As Long, Bet As Long

Public Sub PSYInit(LCDMain, MNU)
    With LCDMain
        .ClearText
        TitleBar LCDMain, "Place your bet"
        .LCDRefresh
    End With
    With MNU
        .ClearItems
        .NewItem "1"
        .NewItem "2"
        .NewItem "5"
        .NewItem "10"
        .NewItem "20"
        .NewItem "50"
        .NewItem "100"
        .NewItem "200"
        .NewItem "500"
        .Visible = True
        .DrawMenu
    End With
    Total = 0
    Correct = 0
        
    DeletePiles
    
    AddPile 0, 0, deck 'Not Displayed
    AddPile 114, 45, horizontal  'Your hand
End Sub
Public Sub PSYSelect(LCDMain, Item As String, MNU)
    If IsNumeric(Item) Then
        With MNU
            .ClearItems
            .top = 1470
            .Height = 540
            .NewItem "Lower"
            .NewItem "Higher"
        End With
        If Item <> "0" Then Bet = Val(Item)
        PSYRound LCDMain, MNU
    Else
        If Item = "Lower" Or Item = "Higher" Then
            FlipCard LCDMain
            MNU.ClearItems
            Total = Total + 1
            If PSYValue(PileList(1).CardList(1).Value) < PSYValue(PileList(1).CardList(0).Value) Then
                If Item = "Lower" Then WinRound MNU Else LoseRound MNU
            Else
                If Item = "Lower" Then LoseRound MNU Else WinRound MNU
            End If
        Else
            With MNU
                .ClearItems
                .NewItem "Lower"
                .NewItem "Higher"
            End With
            PSYRound LCDMain, MNU
        End If
    End If
End Sub
Private Sub PSYRound(LCDMain, MNU)
    EmptyPile PileList(0)
    EmptyPile PileList(1)
    ShuffleDeck PileList(0)
    
    DealCards PileList(0), PileList(1), 2
    PileList(1).CardList(0).Face = True
    Do Until PileList(1).CardList(1).Value <> PileList(1).CardList(0).Value
        DeleteCard PileList(1), 1
        DealCards PileList(0), PileList(1)
    Loop
    
    DrawScreen LCDMain
End Sub
Private Sub DrawScreen(LCDMain)
    With LCDMain
        .ClearText
        .DrawDealer 0, 21
        TitleBar LCDMain, "Psychic Test"
        .PrintText "Your hand", 105, 30
        AUTOdrawdeck 1, LCDMain
        .LCDRefresh
        DoEvents
    End With
End Sub
Private Sub FlipCard(LCDMain)
    PileList(1).CardList(1).Face = True
    DrawScreen LCDMain
End Sub
Private Sub WinRound(MNU)
    Correct = Correct + 1
    AddCash Bet * (Correct / Total)
    MNU.NewItem "You Won"
End Sub
Private Sub LoseRound(MNU)
    AddCash -(Bet * (1 - (Correct / Total)))
    MNU.NewItem "You Lost"
End Sub
Private Function PSYValue(Value As String) As Long
    If Value = "a" Then
        PSYValue = 14
    Else
        PSYValue = Face2Value(Value)
    End If
End Function
