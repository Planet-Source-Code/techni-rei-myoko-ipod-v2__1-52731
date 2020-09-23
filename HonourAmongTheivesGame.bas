Attribute VB_Name = "HonourAmongTheivesGame"
Option Explicit

Private Honour As Long, Cash As Long
Private Security As Long, Loot As Long, Present As Long
Private Age As Long, Poverty As Long

Public Sub HATInit(LCDmain, mnu)
    Honour = iPodSize
    Cash = Honour * 50
    LCDmain.ClearText
    TitleBar LCDmain, "Honour Among Theives", PlayPause
    LCDmain.LCDRefresh
    mnu.Visible = True
    HatRound LCDmain, mnu
End Sub

Public Sub HATSelect(LCDmain, Item As String, mnu)
    Dim temp As Long
    With mnu
    .ClearItems
    .Locked = True
    Select Case LCase(Item)
        Case "rob it"
            Randomize Timer
            temp = Rnd * 12
            Select Case temp
                Case Security
                    .NewItem "You barely escaped"
                    .NewItem "You gained no loot"
                    .NewItem "You gained no honour"
                Case Is > Security
                    temp = (Security / 2.5) - (Present + Age + Poverty) 'present + age + poverty can be from 0 to 5,security can be from 1 to 10
                    .NewItem "You succeeded"
                    .NewItem "You gained " & Loot & " dollars"
                    .NewItem "You " & IIf(temp = 0, "gained no", IIf(temp > 0, "gained ", "lost ") & Abs(temp)) & " honour"
                    Cash = Cash + Loot
                    Honour = Honour + temp
                Case Is < Security
                    .NewItem "You were caught"
                    .NewItem "You lost 500 dollars"
                    .NewItem "You lost 5 honour"
                    Honour = Honour - 5
                    Cash = Cash - 500
            End Select
            
            If Cash <= 0 Then
                .NewItem "You are broke"
                .NewItem "Try Again"
            Else
                .NewItem Empty
                .NewItem "Continue"
            End If
            .NewItem "Quit"
            .selecteditem = 4
            
        Case "skip it"
            Honour = Honour - 1
            .NewItem "You coward"
            .NewItem "You gained no loot"
            .NewItem "You lost 1 honour"
            .NewItem Empty
            .NewItem "Continue"
            .NewItem "Quit"
            .selecteditem = 4
        Case "continue"
            HatRound LCDmain, mnu
        Case "quit"
            AddCash Round(Cash / Honour * 10)
            MainMenu LCDmain, mnu, Tim, "Extra\Games"
        Case "try again"
            HATInit LCDmain, mnu
    End Select
    .Locked = False
    End With
End Sub

Private Sub HatRound(LCDmain, mnu)
    Randomize Timer
    Security = Rnd * 9 + 1
    Age = Rnd * 2
    Present = Rnd * 1
    Loot = Round((Rnd * 9) + 1) * 100
    Poverty = Rnd * 2
    With mnu
        .ClearItems
        .Locked = True
        .NewItem "Security", CStr(Security)
        .NewItem "Difficulty", long2text(Security \ 4, "Easy", "Medium", "Hard")
        .NewItem "Loot", CStr(Loot)
        .NewItem "Age", long2text(Age, "Young", "Middle", "Elderly")
        .NewItem "Is home", long2text(Present, "No", "Yes")
        .NewItem "Class", long2text(Poverty, "Rich", "Middle", "Poor")
        .NewItem Empty
        .NewItem "Cash", CStr(Cash)
        .NewItem "Honour", CStr(Honour)
        .NewItem Empty
        .NewItem "Rob it"
        .Locked = False
        .NewItem "Skip it"
    End With
End Sub

