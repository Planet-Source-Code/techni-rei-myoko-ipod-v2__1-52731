Attribute VB_Name = "TetrisGame"
Option Explicit

Private X As Long, Y As Long, LCDm
Private TileWidth As Long, XTiles As Long, YTiles As Long, Score As Long
Private NextBlock As Long, CurrBlock As Long, XCurr As Long, YCurr As Long, RCurr As Long
Private Grid() As Boolean

Private Type TetrisBlock
    Grid(0 To 15) As Boolean
    Width As Long
    Height As Long
End Type

Private blocklist(0 To 6) As TetrisBlock
Private Function BlockLeft() As Long
    Dim temp As Long
    For temp = 0 To 3
        If Bottom(blocklist(CurrBlock), temp, RCurr) > -1 Then
            BlockLeft = temp
            Exit For
        End If
    Next
End Function
Public Sub RotateBlock()
    If CanRotate Then
        RCurr = RCurr + 1
        If RCurr = 4 Then RCurr = 0
        MoveBlock 0
    End If
End Sub
Public Sub INITTetris(destLCD)
    X = 2
    Y = 23
    Score = 0
    INITBlocks
    INITGrid 15, 27
    TileWidth = 4
    Set LCDm = destLCD
    RandomNextBlock
    CreateBlock
    
    DRAWTetris destLCD
End Sub
Public Sub MoveBlock(direction As Long)
    XCurr = XCurr + direction

    If XCurr + BlockLeft < 0 Then XCurr = -BlockLeft
    If XCurr >= XTiles - BlockWidth Then XCurr = XTiles - BlockWidth
    DRAWTetris LCDm
End Sub
Public Sub DRAWTetris(destLCD)
    Dim temp As Long, temp2 As Long
    destLCD.ClearText
    TitleBar destLCD, "Tetris"
    destLCD.DrawLine X - 1, Y - 1, XTiles * TileWidth + 2, 1
    destLCD.DrawLine X - 1, Y + YTiles * TileWidth, XTiles * TileWidth + 2, 1
    destLCD.DrawLine X - 1, Y, 1, YTiles * TileWidth + 2
    destLCD.DrawLine X + XTiles * TileWidth, Y, 1, YTiles * TileWidth + 2
    
    destLCD.PrintText "Next", X + (XTiles + 1) * TileWidth, Y
    DrawBlock destLCD, blocklist(NextBlock), X + (XTiles + 1) * TileWidth, Y + 10, TileWidth, 0
    
    destLCD.PrintText "Score", X + (XTiles + 1) * TileWidth, Y + 40
    destLCD.PrintText CStr(Score), X + (XTiles + 1) * TileWidth, Y + 50
    
    DrawBlock destLCD, blocklist(CurrBlock), GetX(XCurr), GetY(YCurr), TileWidth, RCurr
    For temp = 0 To YTiles - 1
        For temp2 = 0 To XTiles - 1
            If Grid(temp, temp2) Then DrawTile destLCD, GetX(temp2), GetY(temp), TileWidth
        Next
    Next
    destLCD.LCDRefresh
End Sub
Private Function GetX(Tile As Long) As Long
    GetX = X + Tile * TileWidth
End Function
Private Function GetY(Tile As Long) As Long
    GetY = Y + Tile * TileWidth
End Function
Private Sub CreateBlock()
    XCurr = (XTiles - 1) \ 2
    YCurr = 0
    CurrBlock = NextBlock
    RCurr = 0
    RandomNextBlock
End Sub
Private Sub RandomNextBlock()
    Randomize Timer
    NextBlock = Rnd * 6
End Sub
Private Sub INITGrid(Width As Long, Height As Long)
    XTiles = Width
    YTiles = Height
    ReDim Grid(0 To YTiles - 1, 0 To XTiles - 1)
End Sub
Private Sub DeclareBlock(Block As TetrisBlock, Width As Long, Height As Long, ParamArray Tiles() As Variant)
    With Block
        .Width = Width
        .Height = Height
        Dim temp As Long
        For temp = 0 To UBound(Tiles)
            .Grid(temp) = Tiles(temp)
        Next
    End With
End Sub

Private Sub INITBlocks()
    'Block, Line, L shape, backwards L, cross, Z shape, backwards Z
    DeclareBlock blocklist(0), 2, 2, True, True, False, False, True, True ' Square
    DeclareBlock blocklist(1), 1, 4, True, False, False, False, True, False, False, False, True, False, False, False, True 'Line
    DeclareBlock blocklist(2), 3, 2, False, True, False, False, True, True, True ' Cross
    DeclareBlock blocklist(3), 3, 2, True, False, False, False, True, True, True 'L shape
    DeclareBlock blocklist(4), 3, 2, False, False, True, False, True, True, True 'backwards L shape
    DeclareBlock blocklist(5), 3, 2, False, True, True, False, True, True 'Z shape
    DeclareBlock blocklist(6), 3, 2, True, True, False, False, False, True, True 'backwards Z shape
End Sub
Private Function XYRtoIndex(X As Long, Y As Long, R As Long) As Long
    Select Case R
        Case 0: XYRtoIndex = Y * 4 + X
        Case 1: XYRtoIndex = X * 4 + 3 - Y
        Case 2: XYRtoIndex = (3 - Y) * 4 + 3 - X
        Case 3: XYRtoIndex = (3 - X) * 4 + 3 - Y
    End Select
End Function
Private Sub DrawBlock(destLCD, Block As TetrisBlock, X As Long, Y As Long, Width As Long, Rotation As Long)
    Dim temp As Long, temp2 As Long, temp3 As Long
    For temp = 0 To 3 'x
        For temp2 = 0 To 3 'y
            temp3 = XYRtoIndex(temp, temp2, Rotation)
            If Block.Grid(temp3) Then DrawTile destLCD, X + temp * Width, Y + temp2 * Width, Width
        Next
    Next
End Sub
Private Sub DrawTile(destLCD, X As Long, Y As Long, Width As Long)
    destLCD.DrawSquare X, Y, Width, Width, vbBlack, True
End Sub
Private Function Bottom(Block As TetrisBlock, X As Long, R As Long) As Long
    Dim temp As Long, temp2 As Long
    temp2 = -1
    For temp = 0 To 3
        If Block.Grid(XYRtoIndex(X, temp, R)) Then temp2 = temp
    Next
    Bottom = temp2
End Function
Private Function CanMoveDown() As Boolean
    Dim temp As Long, temp2 As Long, buffer As Boolean
    buffer = True
    For temp = 0 To 3
        temp2 = Bottom(blocklist(CurrBlock), temp, RCurr)
        temp2 = YCurr + temp2
        If temp2 >= 0 Then
            If temp2 + 1 >= YTiles Then 'Detect bottom of the screen
                buffer = False
            Else
                If temp + XCurr >= 0 Then If temp + XCurr < XTiles Then If Grid(temp2 + 1, temp + XCurr) Then buffer = False                 'Detect tiles below
            End If
        End If
    Next
    CanMoveDown = buffer
End Function
Private Function BlockHeight() As Long
    If RCurr = 0 Or RCurr = 2 Then
        BlockHeight = blocklist(CurrBlock).Height
    Else
        BlockHeight = blocklist(CurrBlock).Width
    End If
End Function
Private Function BlockWidth() As Long
    If RCurr = 0 Or RCurr = 2 Then
        BlockHeight = blocklist(CurrBlock).Width
    Else
        BlockHeight = blocklist(CurrBlock).Height
    End If
End Function

Private Sub DestroyCurrentBlock()
    Dim temp As Long, temp2 As Long
    For temp = 0 To 3 'Transfer the block to the grid
        For temp2 = 0 To 3
            If blocklist(CurrBlock).Grid(XYRtoIndex(temp, temp2, RCurr)) Then
                Grid(YCurr + temp2, XCurr + temp) = True
            End If
        Next
    Next
End Sub
Private Function CanRotate() As Boolean
    Dim tempR As Long
    tempR = RCurr + 1
    If tempR = 4 Then tempR = 0
    CanRotate = CollisionDetect(CurrBlock, XCurr, YCurr, tempR)
End Function
Private Function IsAFullLine(Y As Long) As Boolean
    Dim temp As Long, buffer As Boolean
    buffer = True
    For temp = 0 To XTiles - 1
        If Not Grid(Y, temp) Then buffer = False
        Exit For
    Next
    IsAFullLine = buffer
End Function
Private Function CheckAllLines()
    Dim temp As Long
    For temp = YTiles - 1 To 0 Step -1
        Do Until Not IsAFullLine(temp)
            DestroyLine temp
        Loop
    Next
End Function
Private Function DestroyLine(Line As Long)
    Dim temp As Long, temp2 As Long
    For temp = Line - 1 To 0 Step -1
        For temp2 = 0 To XTiles - 1
            Grid(temp + 1, temp2) = Grid(temp, temp2)
            If temp = 0 Then Grid(0, temp2) = False
            Score = Score + 1
        Next
    Next
End Function
Private Function CollisionDetect(Index As Long, X As Long, Y As Long, R As Long) As Boolean
    Dim temp As Long, temp2 As Long, buffer As Boolean
    buffer = True
    For temp = 0 To 3
        For temp2 = 0 To 3
            If blocklist(Index).Grid(XYRtoIndex(temp, temp2, R)) Then
                If X + temp >= 0 And X + temp < XTiles And Y + temp2 >= 0 And Y + temp2 < YTiles Then
                    If Grid(Y + temp2, X + temp) Then buffer = False
                End If
            End If
        Next
    Next
    CollisionDetect = buffer
End Function
Public Sub MoveBlockDown()
    If CanMoveDown Then
        YCurr = YCurr + 1
    Else
        DestroyCurrentBlock
        CheckAllLines
        CreateBlock
    End If
    DRAWTetris LCDm
End Sub
