Attribute VB_Name = "Mp3"
Option Explicit
Dim i As Integer
Dim playlistitem As String
Public Type ID3v1Tag
    id As String * 3
    Title As String * 30
    Artist As String * 30
    Album As String * 30
    Year As String * 4
    Comment As String * 30
    Genre As Byte
End Type
'
Public Tag1 As ID3v1Tag
Public strFilePath As String
Public GenreArray() As String

Public Const strGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
Public Sub Play(strFileName As String)
    With frmMedia
        Call Mp3.ReadTag(strFileName)
        .lblTitle.Caption = Trim(Tag1.Title)
        .lblArtist.Caption = Trim(Tag1.Artist)
        .MediaPlayer1.Filename = strFileName
        .hscPosition.Max = .MediaPlayer1.Duration
        .MediaPlayer1.Play
        .lblDuration.Caption = .MediaPlayer1.Duration / 60
    End With
End Sub
Public Function ReadTag(strFileName As String)
    Dim lFilesize As Long
    Dim fn As Integer
    Dim sFileExt, sData As String
    Dim i As Integer
    fn = FreeFile
    Open strFileName For Binary As #fn
        lFilesize = LOF(fn)
        sFileExt = UCase(Right(strFileName, 3))
        If sFileExt = "MP3" Then
                Get #fn, lFilesize - 127, Tag1.id
                If Tag1.id = "TAG" Then
                    Get #fn, lFilesize - 127, Tag1
                End If
        End If
    Close #fn
    Tag1.Title = Replace(Tag1.Title, Chr(0), "", , , vbBinaryCompare)
    Tag1.Artist = Replace(Tag1.Artist, Chr(0), "", , , vbBinaryCompare)
    Tag1.Album = Replace(Tag1.Album, Chr(0), "", , , vbBinaryCompare)
    Tag1.Year = Replace(Tag1.Year, Chr(0), "", , , vbBinaryCompare)
    Tag1.Comment = Replace(Tag1.Comment, Chr(0), "", , , vbBinaryCompare)
End Function
Public Sub Forw()
On Error Resume Next
    With frmMedia
        .MediaPlayer1.CurrentPosition = .MediaPlayer1.CurrentPosition + 5
        .hscPosition.Value = .MediaPlayer1.CurrentPosition
    End With
End Sub
Public Sub Prev()
On Error Resume Next
    With frmMedia
        .MediaPlayer1.CurrentPosition = .MediaPlayer1.CurrentPosition - 5
        .hscPosition.Value = .MediaPlayer1.CurrentPosition
    End With
End Sub
Public Sub NextTrack()
On Error Resume Next
    With frmMedia
        If .lstList.ListCount <= 1 Then
            Exit Sub
        Else
            If .lstList1.ListIndex = .lstList1.ListCount Then
                Exit Sub
            Else
                .MediaPlayer1.Stop
                If .mnuTypeC(0).Checked = False Then
                    .lstList1.ListIndex = .lstList1.ListIndex + 1
                Else
                    Randomize .lstList1.ListCount
                    .lstList1.ListIndex = .lstList1.ListCount * Rnd
                End If
                    .lstList.ListIndex = .lstList1.ListIndex
                    .txtFileName.Text = .lstList1.Text
                    Call Mp3.Play(.txtFileName.Text)
            End If
        End If
        If .WindowState = vbMinimized Then Call mdlSys.SysTip("MediaPHT 1.0 - " & .lblTitle.Caption)
    End With
End Sub
Public Sub BackTrack()
    With frmMedia
    On Error Resume Next
        If .lstList.ListCount <= 1 Then
            Exit Sub
        Else
            If .lstList.ListIndex = 0 Then
                Exit Sub
            Else
                .MediaPlayer1.Stop
                If .mnuTypeC(0).Checked = False Then
                    .lstList1.ListIndex = .lstList1.ListIndex - 1
                Else
                    Randomize .lstList1.ListCount
                    .lstList1.ListIndex = .lstList1.ListCount * Rnd
                End If
                    .lstList.ListIndex = .lstList1.ListIndex
                    .txtFileName.Text = .lstList1.Text
                    Call Mp3.Play(.txtFileName.Text)
            End If
        End If
            If .WindowState = vbMinimized Then Call mdlSys.SysTip("MediaPHT 1.0 - " & .lblTitle.Caption)
    End With
End Sub
Public Sub VolDec()
With frmMedia
.lblVol = Val(.lblVol.Caption)
If .lblVol > 5 Then
    .lblVol = .lblVol - 5
End If
If .imgVol2.Width > 50 Then
    .imgVol2.Width = .imgVol2.Width - 50
End If
Select Case .lblVol
    Case 100
        .MediaPlayer1.Volume = 0
        Exit Sub
    Case 95
        .MediaPlayer1.Volume = -300
        Exit Sub
    Case 90
        .MediaPlayer1.Volume = -600
        Exit Sub
    Case 85
        .MediaPlayer1.Volume = -900
        Exit Sub
    Case 80
        .MediaPlayer1.Volume = -1200
        Exit Sub
    Case 75
        .MediaPlayer1.Volume = -1500
        Exit Sub
    Case 70
        .MediaPlayer1.Volume = -1800
        Exit Sub
    Case 65
        .MediaPlayer1.Volume = -2100
        Exit Sub
    Case 60
        .MediaPlayer1.Volume = -2400
        Exit Sub
    Case 55
        .MediaPlayer1.Volume = -2700
        Exit Sub
    Case 50
        .MediaPlayer1.Volume = -3000
        Exit Sub
    Case 45
        .MediaPlayer1.Volume = -3300
        Exit Sub
    Case 40
        .MediaPlayer1.Volume = -3600
        Exit Sub
    Case 35
        .MediaPlayer1.Volume = -3900
        Exit Sub
    Case 30
        .MediaPlayer1.Volume = -4200
        Exit Sub
    Case 25
        .MediaPlayer1.Volume = -4500
        Exit Sub
    Case 20
        .MediaPlayer1.Volume = -4800
        Exit Sub
    Case 15
        .MediaPlayer1.Volume = -5100
        Exit Sub
    Case 10
        .MediaPlayer1.Volume = -5400
        Exit Sub
    Case 5
        .MediaPlayer1.Volume = -5700
        Exit Sub
    Case 0
        .MediaPlayer1.Volume = -6000
        Exit Sub
End Select
End With
End Sub
Public Sub VolInc()
With frmMedia
.lblVol = Val(.lblVol.Caption)
If .lblVol < 100 Then
    .lblVol = .lblVol + 5
End If
If .imgVol2.Width < 1000 Then
    .imgVol2.Width = .imgVol2.Width + 50
End If
Select Case .lblVol
    Case 100
        .MediaPlayer1.Volume = 0
        Exit Sub
    Case 95
        .MediaPlayer1.Volume = -300
        Exit Sub
    Case 90
        .MediaPlayer1.Volume = -600
        Exit Sub
    Case 85
        .MediaPlayer1.Volume = -900
        Exit Sub
    Case 80
        .MediaPlayer1.Volume = -1200
        Exit Sub
    Case 75
        .MediaPlayer1.Volume = -1500
        Exit Sub
    Case 70
        .MediaPlayer1.Volume = -1800
        Exit Sub
    Case 65
        .MediaPlayer1.Volume = -2100
        Exit Sub
    Case 60
        .MediaPlayer1.Volume = -2400
        Exit Sub
    Case 55
        .MediaPlayer1.Volume = -2700
        Exit Sub
    Case 50
        .MediaPlayer1.Volume = -3000
        Exit Sub
    Case 45
        .MediaPlayer1.Volume = -3300
        Exit Sub
    Case 40
        .MediaPlayer1.Volume = -3600
        Exit Sub
    Case 35
        .MediaPlayer1.Volume = -3900
        Exit Sub
    Case 30
        .MediaPlayer1.Volume = -4200
        Exit Sub
    Case 25
        .MediaPlayer1.Volume = -4500
        Exit Sub
    Case 20
        .MediaPlayer1.Volume = -4800
        Exit Sub
    Case 15
        .MediaPlayer1.Volume = -5100
        Exit Sub
    Case 10
        .MediaPlayer1.Volume = -5400
        Exit Sub
    Case 5
        .MediaPlayer1.Volume = -5700
        Exit Sub
    Case 0
        .MediaPlayer1.Volume = -6000
        Exit Sub
End Select
End With
End Sub
Public Function WriteTag(strFileName As String, Tag As ID3v1Tag)
On Error Resume Next
    Dim lFilesize As Long
    Dim fn As Integer
    fn = FreeFile
        Open strFileName For Binary As #fn
        lFilesize = LOF(fn)
        Get #fn, lFilesize - 127, Tag.id
            If Tag.id = "TAG" Then
                    Put #fn, , Tag
            Else
                Put #fn, lFilesize - 127, "TAG"
                Close #fn
                Call WriteTag(strFileName, Tag)
            End If
        Close #fn
End Function
Public Sub StopPlayer()
    With frmMedia
        .MediaPlayer1.Stop
        .hscPosition.Value = .hscPosition.Max
        .allowplay = "no"
        .allowpause = False
        .lblDuration.Caption = "00:00"
        .lblPosition.Caption = "0:00"
        If .mnuSpectrumC(1).Checked = True Then
            For i = 0 To .pM.Count - 1
                .pM(i).Visible = False
                .p(i).Visible = False
            Next i
        End If
        If .mnuSpectrumC(2).Checked = True Then
            For i = 0 To .prowM.Count - 1
                .prow(i).Visible = False
                .prowM(i).Visible = False
            Next i
        End If
        If .mnuSpectrumC(3).Checked = True Then
            frmVisual.Timer1.Enabled = False
            For i = 0 To frmVisual.pctEquaM.Count - 1
                frmVisual.pctEquaM(i).Height = 3400
            Next i
        End If
    End With
End Sub

