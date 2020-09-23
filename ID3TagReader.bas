Attribute VB_Name = "ID3v1"
Option Explicit 'Complete ID3 version 1 and 1.1 tag reading and writing capabilities

Const sGenreMatrix As String = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" & _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" & _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" & _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" & _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" & _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" & _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" & _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" & _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" & _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" & _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" & _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" & _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" & _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" & _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" & _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" & _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"

Private Type ID3TAG
    sFilename As String
    sScanned As Boolean
    
    sTitle As String '* 30
    sArtist As String ' * 30
    sAlbum As String '* 30
    sComment As String ' * 30
    sYear As String '* 4
    sGenre As String '* 21
    sTrack As Long
End Type

Public MP3Info As ID3TAG

Public Function Genre(index)
    Dim tempstr() As String, count As Long
    tempstr = Split(sGenreMatrix, "|")
    If IsNumeric(index) Then
        Genre = tempstr(index)
    Else
        Genre = -1
        index = Trim(index)
        For count = 0 To UBound(tempstr)
            If LCase(tempstr(count)) = LCase(index) Then
                Genre = count
                Exit For
            End If
        Next
    End If
End Function
Public Sub WriteID3(filename As String, memory As ID3TAG)
With memory
    .sFilename = filename
    Dim tempfile As Long, ID3TAG As String
    tempfile = FreeFile
    
    ID3TAG = "ÿTAG" & textedit(.sTitle, 30, Chr(0)) & textedit(.sArtist, 30, Chr(0)) & textedit(.sAlbum, 30, Chr(0))
    ID3TAG = ID3TAG & textedit(.sYear, 4, Chr(0)) & textedit(.sComment, 28, Chr(0)) & TRACK(memory) & Chr(Genre(.sGenre))
    
    Open filename For Binary As #tempfile
        If hastag(filename) Then
            Put #tempfile, FileLen(filename) - 128, ID3TAG
        Else
            Put #tempfile, FileLen(filename), ID3TAG
        End If
    Close tempfile
End With
End Sub
Public Function TRACK(memory As ID3TAG) As String
    Dim tempstr As String * 2
    tempstr = Chr(Int(memory.sTrack / 256))
    tempstr = tempstr & Chr(memory.sTrack - (Int(memory.sTrack / 256) * 256))
    TRACK = tempstr
End Function

Public Function textedit(text As String, length As Long, fillchar As String) As String
If Len(text) > length Then text = Left(text, length)
textedit = text & String(length - Len(text), fillchar)
End Function
Public Function hastag(filename As String)
    Dim tempfile As Long, ID3TAG(4) As Byte
    tempfile = FreeFile
    Open filename For Binary As #tempfile
        Get #tempfile, FileLen(filename) - 128, ID3TAG
    Close tempfile
    hastag = combine(ID3TAG, 1, 3) = "TAG"
End Function
Public Sub ReadID3(filename As String, memory As ID3TAG)
On Error Resume Next
With memory
    .sFilename = filename
    Dim tempfile As Long, ID3TAG(128) As Byte
    tempfile = FreeFile
    Open filename For Binary As #tempfile
        Get #tempfile, FileLen(filename) - 128, ID3TAG
    Close tempfile
    .sScanned = False
    If combine(ID3TAG, 1, 3) = "TAG" Then 'If it starts with 'TAG' it's a valid ID3 tag
        .sScanned = True
        
        .sTitle = combine(ID3TAG, 4, 30)
        .sArtist = combine(ID3TAG, 34, 30)
        .sAlbum = combine(ID3TAG, 64, 30)
        .sYear = combine(ID3TAG, 94, 4)
        .sComment = combine(ID3TAG, 98, 30)
        If Len(Trim(.sComment)) <= 28 Then .sTrack = ID3TAG(127) + ID3TAG(126) * 8
        .sGenre = Genre(Asc(combine(ID3TAG, 128, 1)))
    Else
    
        .sTitle = Right(filename, Len(filename) - InStrRev(filename, "\"))
        If InStr(.sTitle, ".") > 0 Then .sTitle = Left(.sTitle, InStrRev(.sTitle, ".") - 1)
        .sArtist = Empty
        .sAlbum = Empty
        .sYear = Empty
        .sComment = Empty
        .sTrack = 0
        .sGenre = Empty
         
    End If
End With
End Sub

Public Function combine(bytearray, start As Long, length As Long, Optional quiton0 As Boolean = True) As String
Dim temp As Long, tempstr As String
For temp = start To start + length - 1
    If bytearray(temp) = 0 And quiton0 = True Then Exit For
    tempstr = tempstr & Chr(bytearray(temp))
Next
combine = Replace(Replace(Trim(tempstr), Chr(0), Empty), Chr(10), vbNewLine)
End Function

