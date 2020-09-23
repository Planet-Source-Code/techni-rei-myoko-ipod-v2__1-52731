Attribute VB_Name = "CalenderFunctions"
Option Explicit
Public CurrDate As Date, Currtime As Long, isAMorPM As Boolean
'Clock
Public Sub INITTime(Ctime As String, lcdm)
    Dim hour As Long, minute As Long
    hour = Val(Format(Ctime, "h"))
    isAMorPM = hour > 12
    If isAMorPM Then hour = hour - 12
    minute = Val(Format(Ctime, "nn"))
    Currtime = hour * 60 + minute
    DrawClock lcdm, Currtime, isAMorPM
End Sub
Public Sub MoveClock(direction As Long, lcdm)
    Currtime = Currtime + direction
    If Currtime > 720 Then
        Currtime = Currtime - 720
        isAMorPM = Not isAMorPM
    End If
    If Currtime < 0 Then
        Currtime = Currtime + 720
        isAMorPM = Not isAMorPM
    End If
    DrawClock lcdm, Currtime, isAMorPM
End Sub
Public Function ActualTime(Tim As Long) As String
    Dim actual As String
    actual = sec2time(Tim)
    If Left(actual, 2) = "0:" Then actual = "12:" & Right(actual, 2)
    ActualTime = actual
End Function
Public Sub DrawClock(lcdm, Tim As Long, isPM As Boolean)
    Dim actual As String, lef As Long
    actual = ActualTime(Currtime)
    lef = 85 - StringWidth(actual)
    With lcdm
        .ClearText
        TitleBar lcdm, "Alarm Time"
        .PrintText actual, lef, 68, True
        .PrintText IIf(isPM, "PM", "AM"), 87, 68, True
        .PrintText "<up>", 77, 54
        .PrintText "<down>", 77, 84
        
        .DrawSquare 48, 66, 64, 2, vbBlack, True
        .DrawSquare 48, 68, lef - 48, 12, vbBlack, True
        .DrawSquare 105, 68, 7, 12, vbBlack, True
        .DrawSquare 85, 68, 2, 12, vbBlack, True
        .LCDRefresh
    End With
End Sub

Public Sub FORMATTEST() 'USed to test format strings, useful to me
    Dim temp As Long
    For temp = Asc("a") To Asc("z")
        If Format("4:08:20 PM", Chr(temp) & Chr(temp)) = "08" Then
            Debug.Print Chr(temp)
        End If
    Next
End Sub

'Calender
Public Sub INITDate(lcdm)
    CurrDate = Now
    SeekDate 0, lcdm
End Sub
Public Function GetDay(Day As Date) As Long
    GetDay = Val(Format(Day, "dd"))
End Function
Public Function GetYear(Day As Date) As Long
    GetYear = Val(Format(Day, "yyyy"))
End Function
Public Function GetMonth(Day As Date) As Long
    GetMonth = Val(Format(Day, "mm"))
End Function
Public Function GetShortDate(Day As Date) As String
    GetShortDate = Format(Day, "MMM") & " " & Format(Day, "yyyy")
End Function
Public Function DaysInMonth(ByVal month As Long, ByVal year As Long) As Long
    month = month + 1
    If month > 12 Then
        month = 1
        year = year + 1
    End If
    DaysInMonth = Val(Format(DateSerial(year, month, 0), "dd"))
End Function
Public Function FirstDay(month As Long, year As Long) As Long
    FirstDay = DayOfWeek(month & "/1/" & year & " 12:00:00 AM")
End Function
Public Function IsLeapYear(year As Long) As Boolean
    IsLeapYear = year Mod 400 = 0 Or (year Mod 4 = 0 And year Mod 100 <> 0)
End Function
Public Function DayOfWeek(Day As Date) As Long
    DayOfWeek = Val(Format(Day, "w"))
End Function
Public Function NameOfWeekDay(Day As Date) As String
    NameOfWeekDay = Format(Day, "dddd")
End Function
Public Function MakeDate(Day As Long, month As Long, year As Long, Optional hour As Long = "12", Optional minute As Long, Optional Second As Long, Optional isAM As Boolean) As String
    MakeDate = month & "/" & Day & "/" & year & " " & hour & ":" & Format(minute, "00") & ":" & Format(Second, "00") & " " & IIf(isAM, "AM", "PM")
End Function

Public Sub DrawCalender(lcdm, Day As Long, month As Long, year As Long)
    Dim temp As Long
    lcdm.ClearText 'Im not making another font :Þ
    TitleBar lcdm, GetShortDate(MakeDate(Day, month, year))
    lcdm.PrintText "SunMonTueWedThuFri  Sat", 5, 27
    For temp = 1 To DaysInMonth(month, year)
        DrawDate lcdm, temp, month, year, Day = temp, HasPlans(MakeDate(temp, month, year))
    Next
    lcdm.LCDRefresh
End Sub
Public Sub DrawDate(lcdm, Day As Long, month As Long, year As Long, Style As Boolean, Optional HasPlan As Boolean)
    Dim x As Long, y As Long, Daye As Date, wid As Long, lef As Long
    Daye = MakeDate(Day, month, year)
    x = 5 + (DayOfWeek(Daye) - 1) * 22
    y = 37 + (GetWeek(Day, month, year) - 1) * 16
    lcdm.DrawSquare x, y, 23, 17
    wid = StringWidth(CStr(Day))
    lef = 2
    lcdm.PrintText CStr(Day), x + lef, y + lef, Style
    If Style Then
        lcdm.DrawSquare x + 1, y + 1, 22, 1, vbBlack, True
        lcdm.DrawSquare x + 1, y + 1, lef - 1, 15, vbBlack, True
        lcdm.DrawSquare x + wid + lef, y + 1, 22 - wid - lef, 15, vbBlack, True
        lcdm.DrawSquare x + lef, y + 12, wid, 4, vbBlack, True
    End If
    If HasPlan Then lcdm.DrawSquare x + 18, y + 12, 3, 3, IIf(Style, lcdm.BackColor, vbBlack), True
End Sub
Public Sub DrawDateNew(lcdm, Day As Long, month As Long, year As Long, Style As Boolean, Optional HasPlan As Boolean)
    Dim x As Long, y As Long, Daye As Date, wid As Long, lef As Long
    Daye = MakeDate(Day, month, year)
    x = 5 + (DayOfWeek(Daye) - 1) * 22
    y = 37 + (GetWeek(Day, month, year) - 1) * 16
    lcdm.DrawSquare x, y, 23, 17
    wid = StringWidth(CStr(Day))
    lef = 12 - wid / 2
    lcdm.PrintText CStr(Day), x + lef, y + 4, Style
    If Style Then
        lcdm.DrawSquare x + 1, y + 1, 22, 3, vbBlack, True
        lcdm.DrawSquare x + 1, y + 1, lef - 1, 15, vbBlack, True
        lcdm.DrawSquare x + wid + lef, y + 1, lef - 1, 15, vbBlack, True
    End If
    If HasPlan Then
        lcdm.DrawSquare x + 2, y + 2, 19, 13, IIf(Style, lcdm.BackColor, vbBlack), False
        lcdm.DrawLine x + 1, y + 15, 21, 1, 8421504
        lcdm.DrawLine x + 21, y + 1, 1, 16, 8421504
    End If
End Sub
Public Function GetWeek(Day As Long, month As Long, year As Long) As Long
    Dim temp As Long, temp2 As Long
    temp2 = 1
    For temp = 2 To Day
        If DayOfWeek(MakeDate(temp, month, year)) = 1 Then temp2 = temp2 + 1
    Next
    GetWeek = temp2
End Function
Public Sub SeekDate(direction As Long, lcdm)
    If direction <> 0 Then CurrDate = DateAdd("d", direction, CurrDate)
    DrawCalender lcdm, GetDay(CurrDate), GetMonth(CurrDate), GetYear(CurrDate)
End Sub
