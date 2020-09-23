Attribute VB_Name = "ParachuteGame"
Option Explicit
Private Const Pi As Double = 3.14159265358979
Public Type Parachute_Parachuter
    ParaX As Long
    ParaY As Long
    ParaZ As Long
    HeliX As Long
    HeliY As Long
    HeliDir As Long
    Released As Boolean
    HeliIsAlive As Boolean
    ParaIsAlive As Boolean
End Type
Public Type Parachute_Bullet
    x As Long
    y As Long
    Angle As Double
End Type

    Private TurretX As Long, TurretY As Long
    Private TurretLength As Long, TurretAngle As Long
    
    Private Bullets() As Parachute_Bullet
    Private BulletCount As Long, BulletSpeed As Long
    
    Private Parachuters() As Parachute_Parachuter
    Private ParachuteCount As Long, ParachuteSpeed As Long
    
    Private BoardWidth As Long, Lives As Long, Kills As Long
    
Private Sub doDeath()
    If Lives > 0 Then Lives = Lives - 1
End Sub
Public Sub InitParachuteVars()
        On Error Resume Next
        TurretAngle = 0
        TurretX = 75
        TurretY = 120
        BoardWidth = TurretX * 2 + 14
        
        ParachuteCount = 4
        ParachuteSpeed = 2
        InitParachuters
        
        BulletSpeed = 5
        Lives = 5
        Kills = 0
End Sub
Public Function CreateBullet() As Long
    On Error Resume Next
        If Lives = 0 Then
            InitParachuteVars
            Exit Function
        End If
        BulletCount = BulletCount + 1
        ReDim Preserve Bullets(BulletCount)
    With Bullets(BulletCount - 1)
        .Angle = DegreesToRadians(TurretAngle + 180)
        .x = findXY(TurretX + 7, TurretY - 1, 10, .Angle, True)
        .y = findXY(TurretX + 7, TurretY - 1, 10, .Angle, False)
    End With
    CreateBullet = BulletCount - 1
End Function
Public Sub NewShrapnel(HelicopterIndex As Long, Angle As Long)
    On Error Resume Next
    With Bullets(CreateBullet)
        .x = Parachuters(HelicopterIndex).HeliX
        .y = Parachuters(HelicopterIndex).HeliY
        If Parachuters(HelicopterIndex).HeliDir = -1 Then
            .Angle = 180 - Angle
        Else
            .Angle = Angle
        End If
    End With
End Sub
Public Sub MoveBullets()
    On Error Resume Next
    Dim temp As Long, x As Long, y As Long
        If BulletCount = 0 Then Exit Sub
        For temp = BulletCount - 1 To 0 Step -1
            x = findXY(Bullets(temp).x + 0, Bullets(temp).y + 0, BulletSpeed + 0, Bullets(temp).Angle + 0, True)
            y = findXY(Bullets(temp).x + 0, Bullets(temp).y + 0, BulletSpeed + 0, Bullets(temp).Angle + 0, False)
            If x > 0 And x < BoardWidth And y > 20 Then
                Bullets(temp).x = x
                Bullets(temp).y = y
                CollisionCheck temp
            Else
                DeleteBullet temp
            End If
        Next
End Sub
Private Sub CollisionCheck(Index As Long)
    Dim temp As Long
        For temp = 0 To ParachuteCount - 1
            If Parachuters(temp).Released And Parachuters(temp).ParaIsAlive Then
                If IsBulletWithin(Index, Parachuters(temp).ParaX, Parachuters(temp).ParaY, 16, 10) Then
                    Parachuters(temp).ParaZ = TurretY + 20 - Parachuters(temp).ParaY
                End If
                If IsBulletWithin(Index, Parachuters(temp).ParaX, Parachuters(temp).ParaY + 10, 16, 10) Then
                    Parachuters(temp).ParaIsAlive = False
                    NewKill
                End If
            End If
            If Parachuters(temp).HeliIsAlive Then
                If IsBulletWithin(Index, Parachuters(temp).HeliX, Parachuters(temp).HeliY, 35, 14) Then
                    NewShrapnel temp, 20
                    NewShrapnel temp, 60
                    Parachuters(temp).HeliIsAlive = False
                    NewKill
                    If Not Parachuters(temp).Released Then
                        Parachuters(temp).ParaIsAlive = False
                        NewKill
                    End If
                End If
            End If
            If Not Parachuters(temp).ParaIsAlive And Not Parachuters(temp).HeliIsAlive Then
                CreateRandomParachuter temp
            End If
        Next
End Sub
Private Function NewKill()
    Kills = Kills + 1
    If Kills Mod 100 = 0 Then Lives = Lives + 1
End Function
Private Function IsWithin(lo As Long, mid As Long, hi As Long) As Boolean
    IsWithin = mid >= lo And mid <= hi
End Function
Private Function IsBulletWithin(Index As Long, x As Long, y As Long, Width As Long, Height As Long) As Boolean
    With Bullets(Index)
        IsBulletWithin = IsWithin(x, .x, x + Width - 1) And IsWithin(y, .y, y + Height - 1)
    End With
End Function
Private Sub DeleteBullet(Index As Long)
        If BulletCount > 0 Then 'If there will be one or more left afterwards
            If Index < BulletCount - 1 Then Bullets(Index) = Bullets(BulletCount - 1) 'Switch the one to be deleted with the last one if the one to be deleted isnt the last one
            BulletCount = BulletCount - 1 'then delete the last one
            ReDim Preserve Bullets(BulletCount)
        Else
            BulletCount = 0
            ReDim Bullets(0)
        End If
End Sub
Private Sub DrawBullets(lcdm)
    Dim temp As Long
        If BulletCount = 0 Then Exit Sub
        For temp = 0 To BulletCount
            If Bullets(temp).x > 0 And Bullets(temp).y > 0 Then lcdm.DrawSquare Bullets(temp).x, Bullets(temp).y, 2, 2
        Next
End Sub
Private Sub DrawParaTrooper(srcHdc As Long, lcdm, Index As Long)
    With Parachuters(Index)
        If .HeliIsAlive Then DrawHelicopter srcHdc, lcdm, .HeliX, .HeliY, .HeliDir = -1, .HeliX Mod 10
        If .Released And .ParaIsAlive Then DrawParachuter srcHdc, lcdm, .ParaX, .ParaY, .ParaZ = 0
    End With
End Sub
Private Sub InitParachuters()
        ReDim Parachuters(ParachuteCount)
        Dim temp As Long
        For temp = 0 To ParachuteCount - 1
            CreateRandomParachuter temp
        Next
End Sub
Public Sub DrawParachuteScreen(srcHdc As Long, lcdm)
    On Error Resume Next
    Static haschecked As Boolean
        lcdm.ClearText
        DrawAllTroppers srcHdc, lcdm
        If Lives = 0 Then
            lcdm.PrintText "Game Over", TurretX + 7 - StringWidth("Game Over") / 2, 40
            lcdm.PrintText Kills & " kills", TurretX + 7 - StringWidth(Kills & " kills") / 2, 60
            If Not haschecked Then
                haschecked = True
                HighScore "Parachute", Kills
            End If
        Else
            haschecked = False
            TitleBar lcdm, Lives & " lives, & " & Kills & " kills"
            DrawCannon srcHdc, lcdm, TurretX, TurretY, TurretAngle
            DrawBullets lcdm
        End If
        lcdm.LCDRefresh
End Sub
Private Sub DrawAllTroppers(srcHdc As Long, lcdm)
    Dim temp As Long
        For temp = 0 To ParachuteCount - 1
            DrawParaTrooper srcHdc, lcdm, temp
        Next
End Sub
Public Sub MoveTurret(Direction As Long)
        On Error Resume Next
        TurretAngle = TurretAngle + Direction
        If TurretAngle > 90 And TurretAngle < 180 Then TurretAngle = 90
        If TurretAngle < 0 Then TurretAngle = 360 + TurretAngle
        If TurretAngle < 270 And TurretAngle > 180 Then TurretAngle = 270
        If TurretAngle > 359 Then TurretAngle = TurretAngle - 360
End Sub
Public Sub MoveParachuters()
    On Error Resume Next
    Dim temp As Long
        For temp = 0 To ParachuteCount
            Parachuters(temp).HeliX = Parachuters(temp).HeliX + ParachuteSpeed * Parachuters(temp).HeliDir
            If Parachuters(temp).Released Then
                Parachuters(temp).ParaY = Parachuters(temp).ParaY + ParachuteSpeed
                If Parachuters(temp).ParaZ > 0 Then 'has no shoot, double the speed
                    Parachuters(temp).ParaY = Parachuters(temp).ParaY + ParachuteSpeed
                    Parachuters(temp).ParaZ = Parachuters(temp).ParaZ - 1
                End If
            Else
                If Parachuters(temp).HeliIsAlive Then
                    If Parachuters(temp).HeliDir > 0 Then
                        If Parachuters(temp).ParaX < Parachuters(temp).HeliX Then Parachuters(temp).Released = True
                    Else
                        If Parachuters(temp).ParaX > Parachuters(temp).HeliX Then Parachuters(temp).Released = True
                    End If
                End If
            End If
        Next
        For temp = ParachuteCount - 1 To 0 Step -1
            If Parachuters(temp).ParaY >= TurretY + 15 Then 'If parachuter is off screen
                If Parachuters(temp).HeliX <= -35 Or Parachuters(temp).HeliX > BoardWidth Then 'if helicopter is off screen
                    If Parachuters(temp).ParaIsAlive Then
                        doDeath
                    End If
                    CreateRandomParachuter temp
                End If
            End If
        Next
End Sub
Private Sub CreateRandomParachuter(Index As Long)
    Dim x As Long, y As Long, Direction As Long
        Randomize Timer
        x = Rnd * (BoardWidth - 16)
        Randomize Timer
        y = 20 + Rnd * 25
        Direction = 1
        Randomize Timer
        If Rnd < 0.5 Then Direction = -1
        NewParachuter Index, x, y, Direction, 5
End Sub
Private Sub NewParachuter(Index As Long, x As Long, y As Long, Direction As Long, shoot As Long)
    With Parachuters(Index)
        .HeliDir = Direction
        If Direction = -1 Then .HeliX = BoardWidth Else .HeliX = -35
        .ParaX = x
        .ParaY = y
        .ParaZ = shoot
        .HeliY = y
        .HeliIsAlive = True
        .ParaIsAlive = True
        .Released = False
    End With
End Sub
Private Function findXY(x As Single, y As Single, Distance As Single, Angle As Double, Optional isx As Boolean = True) As Single
    If isx = True Then findXY = x + Sin(Angle) * Distance Else findXY = y + Cos(Angle) * Distance
End Function
Private Function rad2deg(Radians As Double) As Double
    rad2deg = Radians * 180
End Function
Private Function DegreesToRadians(ByVal Degrees As Double) As Double 'Converts Degrees to Radians.
    DegreesToRadians = Degrees * (Pi / 180)
End Function
Private Function DrawHelicopter(srcHdc As Long, lcdm, x As Long, y As Long, Direction As Boolean, Bladewidth As Long)
    Dim top As Long, leftblade As Long, rightblade As Long, otherblade As Long
    otherblade = 32 'top=6
    leftblade = 9
    If Not Direction Then
        otherblade = 2
        top = 14
        leftblade = 21
    End If
    rightblade = leftblade + 5
    TransBLT srcHdc, 0, top, srcHdc, 35, top, 35, 14, lcdm.hdc, x, y
    lcdm.DrawLine x + leftblade, y, -Bladewidth, 1
    lcdm.DrawLine x + rightblade, y, Bladewidth, 1
End Function
Private Sub DrawParachuter(srcHdc As Long, lcdm, x As Long, y As Long, Optional HasPara As Boolean = True)
    If HasPara Then
        TransBLT srcHdc, 0, 28, srcHdc, 35, 28, 16, 20, lcdm.hdc, x, y
    Else
        TransBLT srcHdc, 4, 38, srcHdc, 39, 38, 8, 10, lcdm.hdc, x + 5, y + 11
    End If
End Sub
Private Sub DrawCannon(srcHdc As Long, lcdm, ByVal x As Long, ByVal y As Long, Angle As Long)
    Dim temp As Double, lefts As Long, top As Long
    TransBLT srcHdc, 17, 33, srcHdc, 52, 33, 14, 15, lcdm.hdc, x, y
    x = x + 7
    temp = DegreesToRadians(Angle)
    lefts = findXY(x + 0, y + 0, 10, temp)
    top = findXY(x + 0, y + 0, 10, temp, False)
    lcdm.DrawLine x, y, x - lefts + 1, y - top + 1
    lcdm.DrawLine x - 1, y, x - lefts + 1, y - top + 1
End Sub
