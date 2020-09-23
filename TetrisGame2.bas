Attribute VB_Name = "VirusGame"
Option Explicit

Private Type VirusBlock
    Lside As Long
    Rside As Long
End Type
Private Type Space
    Side As Long
    Box As Boolean
End Type
Private X As Long, Y As Long, LCDm
Private TileWidth As Long, XTiles As Long, YTiles As Long, Score As Long
Public NextBlock As VirusBlock, CurrBlock As VirusBlock, XCurr As Long, YCurr As Long, RCurr As Long
Private Grid() As Space

Public Function Side2String(Side As Long) As String
    Select Case Side
        Case 0: Side2String = "<heart>"
        Case 1: Side2String = "<diamond>"
        Case 2: Side2String = "<club>"
        Case 3: Side2String = "<spade>"
    End Select
End Function
Public Sub DrawBlock(destlcd, X As Long, Y As Long, Side As Long, Box As Boolean)
    DrawCardIcon Side2String(Side), destlcd.CardHdc, destlcd.hdc, X + 1, Y + 1
    If Box Then destlcd.DrawSquare X, Y, TileWidth + 2, TileWidth + 2
End Sub

Public Sub INITVirus(destlcd)
    X = 1
    Y = 1
    Score = 0
    INITBlocks
    INITGrid 5, 10
    TileWidth = 10
    Set LCDm = destlcd
    RandomNextBlock
    CreateBlock
    DRAWVirus destlcd
End Sub
Public Sub MoveBlock(Optional direction As Long)
    XCurr = XCurr + direction
    DRAWTetris LCDm
End Sub
Public Sub RandomNextBlock()
    Randomize Timer
    With NextBlock
        .Lside = Rnd * 3
        .Rside = Rnd * 3
    End With
End Sub
Public Sub CreateBlock()
    XCurr = XTiles \ 2
    YCurr = 0
    CurrBlock = NextBlock
    RandomNextBlock
End Sub
Public Sub DrawWholeBlock(block As VirusBlock, destlcd)
    With block
        
    End With
End Sub
Public Sub DRAWVirus(destlcd)
    Dim temp As Long, temp2 As Long
    destlcd.ClearText
    TitleBar destlcd, "Virus"
    destlcd.DrawLine X - 1, Y - 1, XTiles * TileWidth + 2, 1
    destlcd.DrawLine X - 1, Y + YTiles * TileWidth, XTiles * TileWidth + 2, 1
    destlcd.DrawLine X - 1, Y, 1, YTiles * TileWidth + 2
    destlcd.DrawLine X + XTiles * TileWidth, Y, 1, YTiles * TileWidth + 2
    destlcd.PrintText "Next", X + (XTiles + 1) * TileWidth, Y
    
    'DrawBlock destLCD, blocklist(NextBlock), X + (XTiles + 1) * TileWidth, Y + 10, TileWidth, 0
    
    destlcd.PrintText "Score", X + (XTiles + 1) * TileWidth, Y + 40
    destlcd.PrintText CStr(Score), X + (XTiles + 1) * TileWidth, Y + 50
    
    'DrawBlock destLCD, blocklist(CurrBlock), GetX(XCurr), GetY(YCurr), TileWidth, RCurr
    
    For temp = 0 To YTiles - 1
        For temp2 = 0 To XTiles - 1
            If Grid(temp, temp2) Then DrawTile destlcd, GetX(temp2), GetY(temp), TileWidth
        Next
    Next
    
    destlcd.LCDRefresh
End Sub
Private Function GetX(Tile As Long) As Long
    GetX = X + Tile * TileWidth
End Function
Private Function GetY(Tile As Long) As Long
    GetY = Y + Tile * TileWidth
End Function
Private Sub INITGrid(Width As Long, Height As Long)
    XTiles = Width
    YTiles = Height
    ReDim Grid(0 To YTiles - 1, 0 To XTiles - 1)
End Sub
