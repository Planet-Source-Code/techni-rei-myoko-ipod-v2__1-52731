Attribute VB_Name = "MCIHandler"
Option Explicit
Public Const AudioFiles As String = "*.wav;*.mp3;*.wma;*.mid;*.midi;*.rmi;*.au;*.snd;*.aif;*.aifc;*.aiff"
Public Const VideoFiles As String = "*.avi;*.mpg;*.mpeg;*.asf;*.wm;*.wmx;*.wmp;*.ivf;*.wmv;*.wvx;*.mpe;*.m1v;*.mp2;*.mpv2;*.mp2v;*.mpa"
Public ContainersHwnd As Long, currentstate As Long, s As String * 30, Mode As String
Public defaultwidth As Long, defaultheight As Long, ALIAS As String, currentfile As String
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Function GetMCIError(lError As Long) As String
Dim sBuffer As String 'We need this to store the returned error
sBuffer = String$(255, Chr(0)) 'This fills out buffer with null characters so the MCI has something to write the error on
mciGetErrorString lError, sBuffer, Len(sBuffer)
sBuffer = Replace$(sBuffer, Chr(0), Empty)
GetMCIError = sBuffer
End Function
Public Function mci(command As String, Optional Ret As String = Empty, Optional val1 As Long = 0, Optional val2 As Long = 0, Optional src As String = Empty, Optional ignore As Boolean = False) As String
Dim temp As Long
Static lasterr As Long
temp = mciSendString(command, Ret, val1, val2)
mci = Ret
If temp > 0 And temp <> 289 And ignore = False Then MsgBox GetMCIError(temp) & vbNewLine & "Please report this to Techni" & vbNewLine & "Problem function: " & src, vbCritical, "MCI error " & temp & ": " & command: lasterr = temp
End Function
Public Sub PlayFile(Filename As String)
    sndPlaySound Filename, 1
End Sub
Public Sub MediaContainersHwnd(hwnd As Long)
    ContainersHwnd = hwnd
End Sub
Public Function trackname(Filename As String) As Long
    trackname = Val(Mid(Filename, 9, 2))
End Function
Public Sub MediaLoad(Optional Filename As String)
On Error Resume Next
currentfile = Filename
defaultwidth = 0
defaultheight = 0
If isReady Then MediaClose
Dim tempstr() As String
   If LCase(Right(Filename, 3)) = "cda" Or Len(Filename) <= 3 Then
        mci "open " & Left(Filename, 1) & ":\ type CDAUDIO alias " & ALIAS & " wait shareable", , , , "Open CD"
        Mode = "CDAUDIO"
        mci "set " & ALIAS & " time format milliseconds", , , , "Set time format"
        If LCase(Right(Filename, 3)) = "cda" Then seektotrack trackname(Filename)
   Else
        mci "open """ & Filename & """ Type MPEGVIDEO Alias " & ALIAS & " parent " & ContainersHwnd & " Style " & &H40000000 & " wait", , , , "Open File"
        mci "set " & ALIAS & " time format milliseconds", , , , "Set time format"
        Mode = "MPEGVIDEO"
        
        mci "where " & ALIAS & " source", s, Len(s), , "Get Default Size", True
        If s = Empty Then s = "0 0 0 0"
        tempstr = Split(s, " ")
        defaultwidth = tempstr(2)
        defaultheight = tempstr(3)
        
   End If
End Sub

Public Function isReady() As Boolean
    isReady = MediaState <> Empty
End Function

Public Sub MediaPlay()
If Not isReady Then Exit Sub
    Select Case MediaState
        Case "stopped": mci "play " & ALIAS, , , , "Play"
        Case "playing": mci "pause " & ALIAS, , , , "Pause"
        Case "paused": mci "resume " & ALIAS, , , , "Resume"
        Case Empty: MsgBox "File is not loaded"
        Case Else: MsgBox "Error"
    End Select
End Sub

Public Sub MediaStop()
If Not isReady Then Exit Sub
    If Mode <> "CDAUDIO" Then MediaSeekto
    mci "stop " & ALIAS, , , , "Stop"
End Sub

Public Sub MediaClose()
If Not isReady Then Exit Sub
   mci "close " & ALIAS, , , , "Close", True
End Sub

Public Sub MediaResize(Width As Long, Height As Long, Optional x As Long, Optional y As Long)
If Not isReady Then Exit Sub
    If defaultheight > 0 Then mci "put " & ALIAS & " window at " & x & " " & y & " " & Width & " " & Height, , , , "Resize"
End Sub
Public Sub MediaSeekto(Optional Second As Long)
If Not isReady Then Exit Sub
    mci "play " & ALIAS & " from " & CStr(Second * 1000), , , , "Seek to"
End Sub

Public Function MediaCurrentPosition() As Long
If Not isReady Then Exit Function
    Dim tempstr As String
    tempstr = mci("status " & ALIAS & " position", s, Len(s), 0, "Current Position")
    If InStr(tempstr, ":") = 0 Then
        MediaCurrentPosition = Round(tempstr / 1000)
    Else
        MediaCurrentPosition = time2sec(tempstr)
    End If
End Function
Public Function time2sec(text As String) As Long
If Not isReady Then Exit Function
    Dim tempstr() As String
    tempstr = Split(text, ":")
    time2sec = Val(tempstr(0)) * 60 + Val(tempstr(1))
End Function

Public Function MediaDuration() As Long
If Not isReady Then Exit Function
    Dim tempstr As String
    tempstr = mci("status " & ALIAS & " length", s, Len(s), 0, "Duration")
    If InStr(tempstr, ":") = 0 Then
        MediaDuration = Round(tempstr / 1000)
    Else
        MediaDuration = time2sec(tempstr)
    End If
End Function

Public Function MediaTimeRemaining() As Long
    MediaTimeRemaining = MediaDuration - MediaCurrentPosition
End Function
    
Public Function MediaIsPlaying() As Boolean
    MediaIsPlaying = MediaState = "playing"
End Function
Public Function MediaIsPaused() As Boolean
    MediaIsPaused = MediaState = "paused"
End Function
Public Function MediaIsStopped() As Boolean
    MediaIsStopped = MediaState = "stopped"
End Function

Public Function MediaState() As String
If ALIAS <> Empty Then MediaState = Replace(mci("status " & ALIAS & " mode", s, Len(s), 0, "CD Status", True), Chr(0), Empty)
End Function

Public Function sec2time(ByVal whattime) As String
On Error Resume Next
If InStr(whattime, ".") > 0 Then whattime = Left(whattime, ".") - 1
Const time_min = 60
Const time_hour = 3600

Dim time_hours As Byte
Dim time_minutes As Byte
Dim time_seconds As Byte

time_hours = intdiv(whattime, time_hour)
time_minutes = intdiv(whattime, time_min)
time_seconds = whattime

If time_hours = 0 Then
    sec2time = Format(time_minutes, "#0") & ":" & Format(time_seconds, "00") 'Dont care about hours as it'll mess with the skin
Else
    sec2time = Format(time_hours, "#0:") & Format(time_minutes, "00") & ":" & Format(time_seconds, "00")
End If
End Function
Public Function intdiv(number, bywhat)
On Error Resume Next
If IsNumeric(number) And IsNumeric(bywhat) Then
On Error Resume Next
Dim temp As Integer
temp = 0
    Do While number >= bywhat
        temp = temp + 1
        number = number - bywhat
    Loop
intdiv = temp
Else
intdiv = 0
End If
End Function

Public Function MediaPositionScale(x As Long, Width As Long) As Long
    If x < 0 Then Exit Function
    If x > Width Then x = Width
    MediaPositionScale = x / Width * MediaDuration
End Function

Public Sub MediaCloseAll()
    mci "close all"
End Sub
Public Function lengthoftracks(upto As Long) As Long
    Static wasupto As Long, wasresult As Long
    Dim temp2 As Long, temp As Long

    If upto = wasupto Then
        temp2 = wasresult
        GoTo endfunct
    End If

    For temp = 1 To upto
        temp2 = time2sec(lengthoftrack(temp)) + temp2
    Next

endfunct:
    wasresult = temp2
    lengthoftracks = temp2
    wasupto = upto
End Function

Public Function iscdpresent() As Boolean
    iscdpresent = CBool(mci("status " & ALIAS & " media present wait", s, Len(s), 0, "Is CD present"))
End Function
Public Function numtracks() As Long
    numtracks = CInt(Mid$(mci("status " & ALIAS & " number of tracks wait", s, Len(s), 0, "Number of tracks"), 1, 2))
End Function
Public Function lengthofcd() As String
    lengthofcd = mci("status " & ALIAS & " length wait", s, Len(s), 0, "Length of CD")
End Function
Public Function lengthoftrack(whichtrack As Long) As String
    lengthoftrack = mci("status " & ALIAS & " length track " & whichtrack, s, Len(s), 0, "Length of track")
End Function
Public Sub seektotrack(TRACK As Long)
    If MediaIsPlaying Or MediaIsPaused Then mci "play " & ALIAS & " from " & TRACK, , , , "Seek to CD line if playing"
    If MediaIsStopped Then mci "seek " & ALIAS & " to " & TRACK, , , , "Seek to CD if stopped"
End Sub

Public Sub CDDOOR(status As Boolean)
    mci "set " & ALIAS & " door " & IIf(status, "open", "close"), 0, 0, 0, "CD Door open/close"
End Sub

Public Function MEDIAfileDuration(Filename As String, Optional ALIAS As String = "MEDIAFILEDURATION") As Long
    mci "open """ & Filename & """ Type MPEGVIDEO Alias " & ALIAS & " wait", , , , "Open File"
    mci "set " & ALIAS & " time format milliseconds", , , , "Set time format"
    MEDIAfileDuration = Val(mci("status " & ALIAS & " length", s, Len(s), 0, "Duration")) / 1000
    mci "close " & ALIAS, , , , "Close", True
End Function
