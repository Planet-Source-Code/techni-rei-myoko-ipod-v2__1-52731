Attribute VB_Name = "SlotMachineGame"
Option Explicit
Private Const CurrDelay As Long = 10
Private Type Slot
    Value As Long
    Offset As Long
End Type
Private Slots(0 To 2) As Slot, IsRandomizing As Boolean, CurrBet As Long, MaxValue As Long
Private Function GetSlot(Index As Long, Optional Direction As Long) As Long
    Dim temp As Long
    temp = Slots(Index).Value + Direction
    If temp < 0 Then temp = temp + MaxValue + 1
    If temp > MaxValue Then temp = temp - MaxValue - 1
    GetSlot = temp
End Function
Public Sub AddBet(Direction As Long, LCDmain)
    If IsRandomizing Then Exit Sub
    CurrBet = CurrBet + Direction
    If CurrBet < 0 Then CurrBet = 0
    If CurrBet > 100 Then CurrBet = 100
    LCDmain.ClearText
    TitleBar LCDmain, "Current Bet: " & CurrBet
    DrawSlots LCDmain
End Sub
Public Sub INITSlotMachine(LCDmain)
    Randomize Timer
    MaxValue = Val(LoadOption("Slot Machine", "Max Value", "5"))
    CurrBet = 5
    Slots(0).Value = Rnd * MaxValue
    Slots(1).Value = Rnd * MaxValue
    Slots(2).Value = Rnd * MaxValue
    'LCDmain.DoubleBuffer = False
    LCDmain.ClearText
    TitleBar LCDmain, "Current Bet: " & CurrBet
    DrawSlots LCDmain
End Sub

Public Sub RandomizeSlots(LCDmain)
    If IsRandomizing Then Exit Sub
    IsRandomizing = True
    Dim temp As Long, temp2 As Long, temp3 As Long, temp4 As Long
    Const BigRotate As Long = 4
    Const NumofRotations As Long = 2
    AddCash -CurrBet
    For temp = 0 To 2
        'Randomize the result
        Randomize Timer
        temp3 = Rnd * MaxValue
        temp4 = 0

        'Spin for till the slot passes the random number 5 times
        Do Until temp4 = NumofRotations
            OffsetSlots temp, BigRotate
            If Slots(temp).Value = temp3 And Slots(temp).Offset = 0 Then temp4 = temp4 + 1
            Sleep CurrDelay 'delay
            DrawSlots LCDmain 'draw
            DoEvents
        Loop

    Next
    If didwin Then
        LCDmain.DrawSquare 67, 94, 30, 6, vbBlack, True 'Second outer box
        LCDmain.LCDRefresh
        AddCash (Slots(0).Value + 1) * CurrBet
    End If
    IsRandomizing = False
End Sub
Public Function didwin() As Boolean
    If Slots(0).Value = Slots(1).Value And Slots(1).Value = Slots(2).Value Then didwin = True
    If CurrBet >= 5 And GetSlot(0, -1) = Slots(1).Value And Slots(1).Value = GetSlot(2, 1) Then didwin = True
    If CurrBet >= 10 And GetSlot(0, 1) = Slots(1).Value And Slots(1).Value = GetSlot(2, -1) Then didwin = True
End Function
Private Sub OffsetSlots(Index As Long, Speed As Long)
    Dim temp As Long
    For temp = Index To 2
        Offset temp, Speed
    Next
End Sub
Private Sub DrawSlots(LCDmain)
    With LCDmain
        .DrawSquare 67, 54, 30, 39, .BackColor, True 'Clear the area
        .DrawSquare 63, 67, 38, 13 'Second middle box
        .DrawSquare 67, 54, 30, 39 'Draw  box
        DrawSlot 0, 70, LCDmain 'left digit
        DrawSlot 1, 79, LCDmain 'middle digit
        DrawSlot 2, 88, LCDmain 'right digit
        .DrawSquare 67, 45, 30, 9, .BackColor, True 'Clear the top
        .DrawSquare 67, 93, 30, 9, .BackColor, True 'Clear the bottom
        .DrawSquare 65, 52, 34, 50 'Second outer box
        .DrawSquare 67, 94, 30, 6 'Second outer box
        If .DoubleBuffer Then
            .LCDRefresh
            DoEvents
        End If
    End With
End Sub
Private Sub Offset(Index As Long, Optional Value As Long = 1)
    With Slots(Index)
        .Offset = .Offset + Value
        If .Offset >= 13 Then
            .Offset = 0 '.Offset Mod 13
            .Value = .Value - 1
            If .Value < 0 Then .Value = MaxValue
            If .Value > MaxValue Then .Value = 0
        End If
    End With
End Sub
Private Sub DrawSlot(Index As Long, Left As Long, LCDmain)
    With LCDmain
        If Slots(Index).Offset > 0 Then .PrintText GetSlot(Index, -2), Left, 45 + Slots(Index).Offset
        .PrintText GetSlot(Index, -1), Left, 57 + Slots(Index).Offset
        .PrintText GetSlot(Index), Left, 69 + Slots(Index).Offset
        .PrintText GetSlot(Index, 1), Left, 81 + Slots(Index).Offset
    End With
End Sub

