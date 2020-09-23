Attribute VB_Name = "LCDCards"
Option Explicit
Private Const IconOrder As String = "<front> <frontleft> <back> <fronttop> <backtop> <heart> <diamond> <spade> <club> 2 3 4 5 6 7 8 9 10 j q k a <hand>"
Private Const IconLeft As String = "0 21 32 0 32 21 29 37 45 0 6 12 18 24 30 0 6 12 21 26 0 6 36"
Private Const IconTop As String = "0 0 0 30 30 34 34 34 34 42 42 42 42 42 42 48 48 48 48 48 54 54 42"
Private Const IconWidth As String = "20 10 20 20 20 7 7 7 7 5 5 5 5 5 5 5 5 8 4 5 5 5 16"
Private Const IconHeight As String = "29 29 29 10 3 7 7 7 7 5 5 5 5 5 5 5 5 5 5 6 5 5 17"
Private Const MaskOffset As Long = 52

Private CardOrder() As String, CardLeft() As String, CardTop() As String, CardWidth() As String, CardHeight() As String

Public Enum iCardConstants
    fullFront = 0
    LeftFront = 1
    TopFront = 2
    fullback = 3
    Topback = 4
End Enum

Public Enum iStackConstants
    deck = 0 'The deck in the upper left corner
    Dealt = 1 'The dealt cards (between the deck and the ace piles) and selecting to move
    Aces = 2 'The Aces Pile in the upper right
    Cards = 3 '(Covered) top backs, then (Uncovered) top fronts, then 1 (Uncovered) front
    Horizontal = 4 'just do the cards across all in full
End Enum

Public Type Card
    Value As String
    Suite As String
    Face As Boolean 'true = show front, false = show back
End Type

Public Type CardPile
    CardList() As Card
    CardCount As Long
    x As Long
    y As Long
    Style As iStackConstants
End Type

Public PileList() As CardPile, PileCount As Long
Public Sub AUTOdrawdeck(pileindex As Long, LCDmain, Optional selecteditem As Long = -1)
    DrawCardPile PileList(pileindex), LCDmain.CardHdc, LCDmain.hdc, PileList(pileindex).x, PileList(pileindex).y, selecteditem
End Sub
Public Sub AddPile(x As Long, y As Long, Style As iStackConstants)
    PileCount = PileCount + 1
    ReDim Preserve PileList(PileCount)
    With PileList(PileCount - 1)
        .x = x
        .y = y
        .Style = Style
    End With
End Sub
Public Sub DeletePiles()
    PileCount = 0
    ReDim PileList(0)
End Sub
Public Sub DealCards(srcPile As CardPile, destPile As CardPile, Optional Cards As Long = 1)
    Dim temp As Long
    For temp = 1 To Cards
        MoveCard srcPile, srcPile.CardCount - 1, destPile
    Next
End Sub
Public Function CardCount(Pile As CardPile, Optional Value As String, Optional Suite As String) As Long
    Dim temp As Long, temp2 As Long
    If Pile.CardCount > 0 Then
        For temp = 0 To Pile.CardCount - 1
            If Len(Value) > 0 Then 'has a value
                If Len(Suite) > 0 Then 'has a suite too
                    If Pile.CardList(temp).Value = Value And Pile.CardList(temp).Suite = Suite Then temp2 = temp2 + 1
                Else 'use any suite
                    If Pile.CardList(temp).Value = Value Then temp2 = temp2 + 1
                End If
            Else 'has only a suite
                If Pile.CardList(temp).Suite = Suite Then temp2 = temp2 + 1
            End If
        Next
    End If
    CardCount = temp2
End Function
Public Function FindCard(Pile As CardPile, Optional start As Long = 0, Optional Value As String, Optional Suite As String) As Long
    Dim temp As Long
    FindCard = -1
    If Pile.CardCount > 0 Then
        For temp = start To Pile.CardCount - 1
            If Len(Value) > 0 Then 'has a value
                If Len(Suite) > 0 Then 'has a suite too
                    If Pile.CardList(temp).Value = Value And Pile.CardList(temp).Suite = Suite Then
                        FindCard = temp
                        Exit Function
                    End If
                Else 'use any suite
                    If Pile.CardList(temp).Value = Value Then
                        FindCard = temp
                        Exit Function
                    End If
                End If
            Else 'has only a suite
                If Pile.CardList(temp).Suite = Suite Then
                    FindCard = temp
                    Exit Function
                End If
            End If
        Next
    End If
End Function
Public Function HasCard(Pile As CardPile, Value As String, Optional Suite As String) As Boolean
    HasCard = CardCount(Pile, Value, Suite) > 0
End Function
Public Sub EmptyPile(Pile As CardPile)
    With Pile
        .CardCount = 0
        ReDim .CardList(.CardCount)
    End With
End Sub
Public Sub AddCard(Pile As CardPile, Optional Value As String = "2", Optional Suite As String = "<heart>", Optional Face As Boolean)
    On Error Resume Next
    With Pile
        .CardCount = .CardCount + 1
        ReDim Preserve .CardList(.CardCount)
    End With
    With Pile.CardList(Pile.CardCount - 1)
        .Face = Face
        .Suite = Suite
        .Value = Value
    End With
End Sub
Public Sub DeleteCard(Pile As CardPile, Index As Long)
    Dim temp As Long
    With Pile
        If .CardCount = 1 Then 'delete only card
            .CardCount = 0
            ReDim .CardList(0)
        Else
            If .CardCount - 1 <> Index Then 'if its not the last card, shift all of them down
                For temp = Index + 1 To .CardCount - 1
                    .CardList(temp - 1) = .CardList(temp)
                Next
            End If
            .CardCount = .CardCount - 1
            ReDim Preserve .CardList(.CardCount)
        End If
    End With
End Sub
Public Function GetColor(Suite As String) As Boolean
    Select Case LCase(Suite)
        Case "<heart>", "<diamond>": GetColor = False
        Case "<club>", "<spade>": GetColor = True
    End Select
End Function

Public Sub ShuffleDeck(destPile As CardPile, Optional Decks As Long = 1)
    'Creates a pile of cards in order, then randomly moves all cards to destPile
    Dim srcPile As CardPile, temp As Long, temp2 As Long
    For temp = 1 To Decks
        For temp2 = 1 To 13
            AddCard srcPile, Value2Face(temp2), "<heart>"
            AddCard srcPile, Value2Face(temp2), "<diamond>"
            AddCard srcPile, Value2Face(temp2), "<club>"
            AddCard srcPile, Value2Face(temp2), "<spade>"
        Next
    Next
    Do Until srcPile.CardCount = 0
        Randomize Timer
        MoveCard srcPile, Rnd * (srcPile.CardCount - 1), destPile
    Loop
End Sub
Public Function Value2Face(Index As Long) As String
    If Index > 1 And Index < 11 Then
        Value2Face = CStr(Index)
    Else
        Select Case Index
            Case 11: Value2Face = "j"
            Case 12: Value2Face = "q"
            Case 13: Value2Face = "k"
            Case 1, 14: Value2Face = "a"
        End Select
    End If
End Function
Public Function Value2Suite(Index As Long) As String
    Select Case Index
            Case 1: Value2Suite = "<heart>"
            Case 2: Value2Suite = "<diamond>"
            Case 3: Value2Suite = "<club>"
            Case 4: Value2Suite = "<spade>"
    End Select
End Function
Public Function RelativeCard(Value As String, Optional Direction As Long = 1) As String
    Dim temp As Long
    temp = Face2Value(Value) + Direction
    If temp > 14 Then temp = 14
    If temp < 1 Then temp = temp Mod 14 + 13
    RelativeCard = Value2Face(temp)
End Function
Public Function Face2Value(Value As String) As Long
    If IsNumeric(Value) Then
        Face2Value = Val(Value)
    Else
        Select Case LCase(Value)
            Case "j": Face2Value = 11
            Case "q": Face2Value = 12
            Case "k": Face2Value = 13
            Case "a": Face2Value = 1
        End Select
    End If
End Function
Public Sub MoveCard(srcPile As CardPile, Index As Long, destPile As CardPile)
    If Index < 0 Or srcPile.CardCount <= Index Then Exit Sub
    With srcPile.CardList(Index)
        AddCard destPile, .Value, .Suite, .Face
    End With
    DeleteCard srcPile, Index
End Sub
Public Sub MoveCards(srcPile As CardPile, Index As Long, destPile As CardPile)
    Dim temp As Long
    For temp = Index To srcPile.CardCount - 1
        MoveCard srcPile, Index, destPile
    Next
End Sub
Public Sub DealCardsMulti(srcPile As CardPile, Cards As Long, ParamArray destPiles() As Variant)
    Dim temp As Long, temp2 As Long
    For temp = 1 To Cards
        For temp2 = 0 To UBound(destPiles)
            MoveCard srcPile, srcPile.CardCount - 1, PileList(CInt(destPiles(temp2)))
        Next
    Next
End Sub
Public Sub RevealHand(Pile As CardPile, Optional Face As Boolean = True)
    Dim temp As Long
    If Pile.CardCount > 0 Then
        For temp = 0 To Pile.CardCount - 1
            Pile.CardList(temp).Face = Face
        Next
    End If
End Sub
'insignificant graphics routines
Public Function PileWidth(Pile As CardPile, Optional selecteditem As Long = -1) As Long
    Select Case Pile.Style
        Case deck, Aces, Cards: PileWidth = 20 '[]
        Case Horizontal: PileWidth = 21 * (Pile.CardCount - 1) '[] [] [] []
        Case Dealt '[[[[]
            If selecteditem > -1 And selecteditem < Pile.CardCount - 1 Then
                PileWidth = 39 + (Pile.CardCount - 2) * 20
            Else
                PileWidth = 20 + (Pile.CardCount - 1) * 20
            End If
    End Select
End Function
Public Function PileHeight(Pile As CardPile, Optional selecteditem As Long = -1) As Long
    Dim temp As Long, Height As Long
    Select Case Pile.Style
        Case deck, Aces, Horizontal, Dealt: PileHeight = 29
        Case Cards
            Height = 29
            For temp = 0 To Pile.CardCount - 2
                With Pile.CardList(temp)
                    Height = Height + IIf(.Face, 10, 3)
                End With
            Next
            PileHeight = Height
    End Select
End Function
Public Sub DrawCardPile(Pile As CardPile, srcHdc As Long, DestHdc As Long, ByVal x As Long, ByVal y As Long, Optional selecteditem As Long = -1)
    Dim temp As Long 'With Pile.CardList(Pile.CardCount - 1)
    If Pile.CardCount = 0 And (Pile.Style <> Aces And Pile.Style <> deck) Then
        If selecteditem > -1 Then DrawCardIcon "<hand>", srcHdc, DestHdc, x + 2, y + 10
        Exit Sub
    End If
    Select Case Pile.Style
        Case deck
            If Pile.CardCount > 0 Then
                DrawCardStyle srcHdc, DestHdc, x, y, fullback, , , selecteditem > -1
            Else
                If selecteditem > -1 Then DrawCardIcon "<hand>", srcHdc, DestHdc, x + 2, y + 10
            End If
        Case Aces
            If Pile.CardCount = 0 Then
                DrawCardStyle srcHdc, DestHdc, x, y, fullFront, Empty, Empty, selecteditem > -1
            Else
                With Pile.CardList(Pile.CardCount - 1)
                    DrawCardStyle srcHdc, DestHdc, x, y, fullFront, .Value, .Suite, selecteditem > -1
                End With
            End If
        Case Dealt
            For temp = 0 To Pile.CardCount - 1
                With Pile.CardList(temp)
                    If selecteditem = temp Or temp = Pile.CardCount - 1 Then
                        DrawCardStyle srcHdc, DestHdc, x, y, fullFront, .Value, .Suite, selecteditem = temp
                        x = x + 19 'shift right the width of the card
                    Else
                        DrawCardStyle srcHdc, DestHdc, x, y, LeftFront, .Value, .Suite
                        x = x + 10 'shift right the width of the card
                    End If
                End With
            Next
        Case Cards
            For temp = 0 To Pile.CardCount - 1
                With Pile.CardList(temp)
                    If temp = Pile.CardCount - 1 Then 'Last card in pile
                        If .Face Then 'Is uncovered
                            DrawCardStyle srcHdc, DestHdc, x, y, fullFront, .Value, .Suite, selecteditem > -1
                        Else
                            DrawCardStyle srcHdc, DestHdc, x, y, fullback, , , selecteditem > -1
                        End If
                    Else 'Is not the last card
                        If .Face Then 'Is uncovered
                            DrawCardStyle srcHdc, DestHdc, x, y, TopFront, .Value, .Suite 'cannot be selected
                            y = y + 10 'shift down the height of the card
                        Else
                            DrawCardStyle srcHdc, DestHdc, x, y, Topback
                            y = y + 3 'shift down the height of the card
                        End If
                    End If
                End With
            Next
        Case Horizontal
            For temp = 0 To Pile.CardCount - 1
                With Pile.CardList(temp)
                    If .Face Then 'Is uncovered
                        DrawCardStyle srcHdc, DestHdc, x, y, fullFront, .Value, .Suite, selecteditem = temp
                    Else
                        DrawCardStyle srcHdc, DestHdc, x, y, fullback, , , selecteditem = temp
                    End If
                    x = x + 21
                End With
            Next
    End Select
End Sub
Private Sub InitCards()
    Static hasinit As Boolean
    If Not hasinit Then
        hasinit = True
        CardOrder = Split(IconOrder, " ")
        CardLeft = Split(IconLeft, " ")
        CardTop = Split(IconTop, " ")
        CardWidth = Split(IconWidth, " ")
        CardHeight = Split(IconHeight, " ")
    End If
End Sub

Private Function GetOrder(ByVal text As String) As Long
    InitCards
    GetOrder = -1
    Dim temp As Long
    text = LCase(text)
    For temp = 0 To UBound(CardOrder)
        If text = CardOrder(temp) Then
            GetOrder = temp
            Exit For
        End If
    Next
End Function
Public Function DrawCardIcon(text As String, srcHdc As Long, DestHdc As Long, x As Long, y As Long) As Boolean
    Dim temp As Long
    temp = GetOrder(text)
    If temp > -1 Then
        TransBLT srcHdc, Val(CardLeft(temp)), Val(CardTop(temp)), srcHdc, Val(CardLeft(temp)) + MaskOffset, Val(CardTop(temp)), Val(CardWidth(temp)), Val(CardHeight(temp)), DestHdc, x, y
        DrawCardIcon = True
    End If
End Function

Public Sub DrawCardStyle(srcHdc As Long, DestHdc As Long, x As Long, y As Long, Style As iCardConstants, Optional Value As String = "2", Optional Suite As String = "<heart>", Optional Hand As Boolean)
    Select Case Style
        Case fullFront
            DrawCardIcon "<front>", srcHdc, DestHdc, x, y
            DrawCardIcon Value, srcHdc, DestHdc, x + 2, y + 2
            DrawCardIcon Suite, srcHdc, DestHdc, x + 11, y + 2
            If Hand Then DrawCardIcon "<hand>", srcHdc, DestHdc, x + 2, y + 10
        Case LeftFront
            DrawCardIcon "<frontleft>", srcHdc, DestHdc, x, y
            DrawCardIcon Value, srcHdc, DestHdc, x + 2, y + 2
            DrawCardIcon Suite, srcHdc, DestHdc, x + 2, y + 9
        Case TopFront
            DrawCardIcon "<fronttop>", srcHdc, DestHdc, x, y
            DrawCardIcon Value, srcHdc, DestHdc, x + 2, y + 2
            DrawCardIcon Suite, srcHdc, DestHdc, x + 11, y + 2
        Case fullback
            DrawCardIcon "<back>", srcHdc, DestHdc, x, y
            If Hand Then DrawCardIcon "<hand>", srcHdc, DestHdc, x + 2, y + 10
        Case Topback
            DrawCardIcon "<backtop>", srcHdc, DestHdc, x, y
    End Select
End Sub

