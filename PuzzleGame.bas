Attribute VB_Name = "PuzzleGame"
Option Explicit
Private Const PZLTILEWIDTH As Long = 11
Private Type PZLTILE
    Shape As Long
    Falling As Boolean
    Dirty As Boolean
End Type
Private PZLmap() As PZLTILE, PZLWidth As Long, PZLHeight As Long, PZLScore As Long, PZLLevel As Long, PZLX As Long, PZLY As Long, PZLCursor As Long
Public Sub PZLMoveCursor(Direction As Long, LCDmain)
    If Not Tim.Enabled Then Exit Sub
    ClearCursor LCDmain
    PZLCursor = PZLCursor + Direction
    If PZLCursor < 1 Then PZLCursor = 1
    If PZLCursor > PZLWidth - 1 Then PZLCursor = PZLWidth - 1
    DrawCursor LCDmain
    LCDmain.LCDRefresh
End Sub
Private Sub ClearCursor(LCDmain)
    LCDmain.DrawSquare GetPos(PZLX, PZLCursor), GetPos(PZLY, PZLHeight + 1) + 2, PZLTILEWIDTH * 2 - 1, 3, LCDmain.BackColor, True
End Sub
Private Sub DrawCursor(LCDmain)
    LCDmain.DrawSquare GetPos(PZLX, PZLCursor), GetPos(PZLY, PZLHeight + 1) + 2, PZLTILEWIDTH * 2 - 1, 3
End Sub
Private Sub SetTile(x As Long, y As Long, Optional Shape As Long, Optional IsFalling As Boolean)
    With PZLmap(y, x)
        .Shape = Shape
        .Falling = IsFalling
        .Dirty = True
    End With
End Sub
Private Function IsEmpty(x As Long, y As Long) As Boolean
    IsEmpty = PZLmap(y, x).Shape = 0
End Function
Private Function DropTile(x As Long, y As Long) As Boolean
    Dim temp As Long
    If IsEmpty(x, y) Then Exit Function
    If y < PZLHeight Then
        If Not PZLmap(y, x).Falling Then 'is a map block, drop to bottom
            For temp = y + 1 To PZLHeight
                If IsEmpty(x, temp) Then
                    SetTile x, temp, PZLmap(temp - 1, x).Shape, PZLmap(temp - 1, x).Falling
                    SetTile x, temp - 1
                    DropTile = True
                Else
                    CheckLine x
                    Exit For
                End If
            Next
        Else 'is a dropped block, drop slowly
            If IsEmpty(x, y + 1) Then
                SetTile x, y + 1, PZLmap(y, x).Shape, PZLmap(y, x).Falling
                SetTile x, y
                DropTile = True
            Else
                CheckLine x
                PZLmap(y, x).Falling = False
            End If
        End If
    Else
        PZLmap(y, x).Falling = False
    End If
End Function
Private Sub CheckLine(x As Long)
    Dim temp As Long, temp2 As Long, temp3 As Long
    For temp = PZLHeight To 1 Step -1
        'Check vertical'going up
        temp2 = CountTiles(x, temp)
        If temp2 > 2 Then
            PZLScore = PZLScore + calcscore(temp2)
            For temp3 = temp To temp - temp2 + 1 Step -1
                If PZLmap(temp3, x).Falling Then CreateRandomBlock
                SetTile x, temp3, 5
            Next
        End If
        'Check horizontal going left
        temp2 = CheckTiles2(x, temp)
        If temp2 > 2 Then
            PZLScore = PZLScore + calcscore(temp2)
            For temp3 = x To x - temp2 + 1 Step -1
                If PZLmap(temp, temp3).Falling Then CreateRandomBlock
                SetTile temp3, temp, 5
            Next
        End If
        'Check horizontal going right
        temp2 = CheckTiles3(x, temp)
        If temp2 > 2 Then
            PZLScore = PZLScore + calcscore(temp2)
            For temp3 = x To x + temp2 - 1
                If PZLmap(temp, temp3).Falling Then CreateRandomBlock
                SetTile temp3, temp, 5
            Next
        End If
    Next
End Sub
Private Function CheckTiles2(x As Long, y As Long) As Long
    Dim temp As Long
    If PZLmap(y, x).Shape = 0 Then Exit Function
    For temp = x To 1 Step -1
        If PZLmap(y, temp).Shape <> PZLmap(y, x).Shape Then
            CheckTiles2 = x - temp
            Exit For
        Else
            If temp = 1 Then
                CheckTiles2 = x - temp
            End If
        End If
    Next
End Function
Private Function CheckTiles3(x As Long, y As Long) As Long
    Dim temp As Long
    If PZLmap(y, x).Shape = 0 Then Exit Function
    For temp = x To PZLWidth
        If PZLmap(y, temp).Shape <> PZLmap(y, x).Shape Then
            CheckTiles3 = temp - x
            Exit For
        Else
            If temp = 1 Then
                CheckTiles3 = temp - x
            End If
        End If
    Next
End Function
Private Function calcscore(Tiles As Long) As Long
    calcscore = 10 ^ (Tiles - 2)
End Function
Private Function CountTiles(x As Long, y As Long) As Long
    Dim temp As Long
    If PZLmap(y, x).Shape = 0 Then Exit Function
    For temp = y To 1 Step -1
        If PZLmap(temp, x).Shape <> PZLmap(y, x).Shape Then
            CountTiles = y - temp
            Exit For
        Else
            If temp = 1 Then
                CountTiles = y - temp
            End If
        End If
    Next
End Function
Private Sub SwitchTile(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    Dim temp As PZLTILE
    PZLmap(Y1, X1).Dirty = True
    PZLmap(Y2, X2).Dirty = True
    temp = PZLmap(Y1, X1)
    PZLmap(Y1, X1) = PZLmap(Y2, X2)
    PZLmap(Y2, X2) = temp
End Sub
Private Sub MoveTile(x As Long, Y1 As Long, Y2 As Long)
    PZLmap(Y2, x) = PZLmap(Y1, x)
    SetTile x, Y1
End Sub
Private Function GetPos(start As Long, Position As Long) As Long
    GetPos = start + (Position - 1) * (PZLTILEWIDTH - 1)
End Function
Public Sub PZLDrawScreen(LCDmain, Optional Drop As Boolean = True)
    Dim temp As Long, temp2 As Long, temp3 As Boolean, temp4 As Boolean
    LCDmain.ClearText
    LCDmain.DrawSquare PZLX, PZLY, PZLWidth * PZLTILEWIDTH - 7, PZLHeight * PZLTILEWIDTH - 3
    TitleBar LCDmain, PZLScore & " Pts"
    For temp = PZLWidth To 1 Step -1
        For temp2 = PZLHeight To 1 Step -1
            If Not IsEmpty(temp, temp2) Then
                temp4 = PZLmap(temp2, temp).Falling
                DrawTile temp, temp2, LCDmain
                If Drop Then
                    temp3 = DropTile(temp, temp2)
                    If temp4 And Not temp3 Then CreateRandomBlock
                End If
            End If
        Next
        CheckLine temp
    Next
    DrawCursor LCDmain
    LCDmain.LCDRefresh
    If IsGameOver Then
        Tim.Enabled = False
        HighScore "Puzzle", PZLScore
    End If
End Sub
Private Sub DrawTile(x As Long, y As Long, LCDmain)
    With PZLmap(y, x)
            If Not IsEmpty(x, y) Then
                If .Shape < 5 Then
                    LCDmain.DrawSquare GetPos(PZLX, x), GetPos(PZLY, y), PZLTILEWIDTH, PZLTILEWIDTH
                    DrawCardIcon Value2Suite(PZLmap(y, x).Shape), LCDmain.CardHdc, LCDmain.hdc, GetPos(PZLX, x) + 2, GetPos(PZLY, y) + 2
                Else
                    LCDmain.DrawSquare GetPos(PZLX, x), GetPos(PZLY, y), PZLTILEWIDTH, PZLTILEWIDTH, vbBlack, True
                    SetTile x, y
                End If
            End If
            .Dirty = False
    End With
End Sub
Private Function IsGameOver() As Boolean
    Dim temp As Long
    IsGameOver = True
    For temp = 1 To PZLWidth
        If IsEmpty(temp, 1) Or PZLmap(1, temp).Shape = 5 Then
            IsGameOver = False
            Exit For
        End If
    Next
End Function
Public Sub PZLInit(LCDmain)
    Dim temp As Long
    PZLWidth = Val(LoadOption("Puzzle", "Board width", "8"))
    PZLHeight = 10
    PZLX = 1
    PZLY = 25
    PZLScore = 0
    PZLCursor = PZLWidth \ 2
    ReDim PZLmap(1 To PZLHeight, 1 To PZLWidth)
    PZLDrawScreen LCDmain
    For temp = 1 To Val(LoadOption("Puzzle", "Drop at once", "1"))
        CreateRandomBlock
    Next
End Sub
Public Sub CreateRandomBlock()
    Dim x As Long
    If IsGameOver Then Exit Sub
    Randomize Timer
    x = (Rnd * (PZLWidth - 1)) + 1
    Do Until IsEmpty(x, 1)
        x = (Rnd * (PZLWidth - 1)) + 1
    Loop
    SetTile x, 1, Rnd * 3 + 1, True
End Sub
Public Sub PZLSwitch(LCDmain)
    Dim temp As Long, LeftFall As Boolean, RightFall As Boolean, doSwitch As Boolean, doDraw As Boolean
    If Not Tim.Enabled Then Exit Sub
    For temp = 1 To PZLHeight
        LeftFall = PZLmap(temp, PZLCursor).Falling
        RightFall = PZLmap(temp, PZLCursor + 1).Falling
        doSwitch = False
        
        'If both are standard non moving tiles
        If LeftFall = False And RightFall = False Then doSwitch = True
        'one is a falling tile, if its not next to an empty block, switch the tiles
        If LeftFall And Not IsEmpty(PZLCursor + 1, temp) And Not RightFall Then doSwitch = True
        If RightFall And Not IsEmpty(PZLCursor, temp) And Not LeftFall Then doSwitch = True
        
        If doSwitch Then
            SwitchTile PZLCursor, temp, PZLCursor + 1, temp
            doDraw = True
        End If
    Next
    If doDraw Then
        PZLDrawScreen LCDmain, False
        CheckLine PZLCursor
        CheckLine PZLCursor + 1
    End If
End Sub
