Attribute VB_Name = "mdlPlaylist"
Option Explicit
Dim i As Integer
Dim playlistitem As String
Public Function GetShortName(str As String)
    GetShortName = Mid(str, InStrRev(str, "\", -1, vbBinaryCompare) + 1)
End Function
Public Sub GoBottom()
On Error Resume Next
    With frmMedia
        .lstList.ListIndex = .lstList.ListCount - 1
        .lstList1.ListIndex = .lstList.ListIndex
        .txtFileName.Text = .lstList1.Text
        Call Mp3.Play(.txtFileName.Text)
        .timerVisual.Enabled = True
        .paused = False
    End With
End Sub
Public Sub GoTop()
On Error Resume Next
    With frmMedia
        .lstList.ListIndex = 0
        .lstList1.ListIndex = .lstList.ListIndex
        .txtFileName.Text = .lstList1.Text
        Call Mp3.Play(.txtFileName.Text)
        .timerVisual.Enabled = True
        .paused = False
    End With
End Sub
Public Sub SubFile()
On Error Resume Next
    With frmMedia
        If .lstList.ListIndex <> -1 Then
            If .lstList.Text = .MediaPlayer1.Filename Then
                    i = .lstList.ListIndex
                    .MediaPlayer1.Stop
                    .hscPosition.Value = .hscPosition.Max
                    .allowplay = "no"
                    .allowpause = False
                    .lblDuration.Caption = "00:00"
                    .lblTitle.Caption = ""
                    .lblArtist.Caption = ""
                    .lstList1.RemoveItem (.lstList.ListIndex)
                    .lstList.RemoveItem (.lstList.ListIndex)
                    If i < .lstList.ListCount And .lstList.ListCount >= 1 Then
                        .txtFileName.Text = .lstList1.List(i)
                        .lstList.ListIndex = i
                        .allowplay = "yes"
                        Call Mp3.Play(.txtFileName.Text)
                        .timerVisual.Enabled = True
                        .paused = False
                    End If
            Else
                    .lstList1.RemoveItem (.lstList.ListIndex)
                    .lstList.RemoveItem (.lstList.ListIndex)
            End If
            .lstList.Clear
            For i = 0 To .lstList1.ListCount - 1
                Call Mp3.ReadTag(.lstList1.List(i))
                If bShowNumber = True Then
                    .lstList.AddItem (i + 1 & "-" & Trim(Tag1.Artist) & "-" & Trim(Tag1.Title))
                Else
                    .lstList.AddItem (Trim(Tag1.Artist) & "-" & Trim(Tag1.Title))
                End If
            Next i
        End If
    End With
End Sub
Public Sub SubFolder()
On Error Resume Next
    With frmMedia
           If .lstList.ListCount > 0 Then
                   .MediaPlayer1.Stop
                   .lblTitle.Caption = ""
                   .lblArtist.Caption = ""
                   .lblDuration = "00:00"
                   .lblPosition.Caption = "0:00"
                   .timerVisual.Enabled = False
                   .Timer1.Enabled = False
               If .mnuSpectrumC(1).Checked = True Then
                   For i = 0 To .pM.Count - 1
                       .pM(i).Visible = False
                       .p(i).Visible = False
                   Next i
               End If
               If .mnuSpectrumC(2).Checked = True Then
                   For i = 0 To .prow.Count - 1
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
               .lstList.Clear
               .lstList1.Clear
        End If
    End With
End Sub
Public Sub AddPl()
On Error Resume Next
    Dim strFilePath As String
    'Mo hop thoai Open
    With frmMedia
            With .cdloLoad
                .InitDir = strPathP
                .DefaultExt = "m3u"
                .Filter = "MediaPlaylist (*.m3u,*.pls) |*.m3u;*.pls | m3u Playlist (*.m3u)|*.m3u| Pl Pls Playlist (*.pls )| *.pls"
                .ShowOpen
                .CancelError = True
                strFilePath = .Filename
            End With
            If Len(strFilePath) > 0 Then
                    .lstList1.Clear
                    .lstList.Clear
                    If UCase(Right(strFilePath, 3)) = "M3U" Then Call LoadPlaylistM3U(strFilePath)
                    If UCase(Right(strFilePath, 3)) = "PLS" Then Call LoadPlaylistPLS(strFilePath)
            Else
                Exit Sub
            End If
            If .cdloLoad.FileTitle = "" Then Exit Sub
    End With
End Sub
Sub AddFile()
On Error Resume Next
    Dim strFilePath As String
    'Mo hop thoai Open
    With frmMedia
            With .cdloLoad
                .InitDir = strPathM
                .DialogTitle = "Mo file nhac"
                .Filter = "All Supported Files (*.mp3,*.wma,*.wav,*.mid)|*.mp3;*.wma;*.wav;*.mid|MP3 Files (*.mp3)|*.mp3|Wav Files (*.wma)|*.wma|Wave Files (*.wav)|*.wav|Midi Files (*.mid)|*.mid"
                .ShowOpen
                .CancelError = True
                strFilePath = .Filename
            End With
            If Len(strFilePath) > 0 Then
                .lstList1.AddItem (strFilePath)
                If UCase(Right(strFilePath, 3)) = "MP3" Then
                    Call Mp3.ReadTag(strFilePath)
                    If bShowNumber = True Then
                        .lstList.AddItem (.lstList.ListCount + 1 & "-" & Trim(Tag1.Artist) & "_" & Trim(Tag1.Title))
                    Else
                        .lstList.AddItem (Trim(Tag1.Artist) & "_" & Trim(Tag1.Title))
                    End If
                Else
                    If bShowNumber = True Then
                        .lstList.AddItem (.lstList.ListCount + 1 & "-" & GetShortName(strFilePath))
                    Else
                        .lstList.AddItem (GetShortName(strFilePath))
                    End If
                End If
            End If
        End With
End Sub
Public Sub LoadPlaylistM3U(Filename As String)
    With frmMedia
        Open Filename For Input As #1
        i = 0
        Do Until EOF(1)
            Line Input #1, playlistitem
            If Mid(playlistitem, 1, 8) <> "#EXTINF:" And Mid(playlistitem, 1, 8) <> "#EXTM3U" Then
                .lstList1.AddItem (playlistitem)
            End If
            If Mid(playlistitem, 1, 8) = "#EXTINF:" Then
                If bShowNumber = True Then
                        i = i + 1
                        .lstList.AddItem (i & "-" & Mid(playlistitem, InStr(1, playlistitem, ",", vbBinaryCompare) + 1))
                Else
                        .lstList.AddItem (Mid(playlistitem, InStr(1, playlistitem, ",", vbBinaryCompare) + 1))
                End If
            End If
        Loop
        Close #1
    End With
End Sub
Public Sub LoadPlaylistPLS(Filename As String)
With frmMedia
    Open Filename For Input As #1
    i = 0
    Do Until EOF(1)
       Input #1, playlistitem
       If UCase(Mid(playlistitem, 1, 4)) = "FILE" Then
        .lstList1.AddItem (Mid(playlistitem, InStr(1, playlistitem, "=", vbBinaryCompare) + 1))
        End If
        If UCase(Mid(playlistitem, 1, 5)) = "TITLE" Then
            If bShowNumber = True Then
                i = i + 1
                .lstList.AddItem (i & "-" & Mid(playlistitem, InStr(1, playlistitem, "=", vbBinaryCompare) + 1))
            Else
                .lstList.AddItem (Mid(playlistitem, InStr(1, playlistitem, "=", vbBinaryCompare) + 1))
            End If
        End If
    Loop
    Close #1
End With
End Sub
Public Sub SavePlaylist()
    On Error Resume Next
    With frmMedia
    Dim intRecord As Integer
    Dim strFilePath As String
    Dim fn As Integer
    fn = FreeFile
    If Mid(.lstList1.List(1), 2, 1) = ":" Then
    'Mo hop thoai Save
        With .cdloSave
            .InitDir = strPathP
            .DialogTitle = "Luu tap tin danh sach"
            .DefaultExt = "m3u"
            .Filter = "MediaPHT Playlist (*.m3u)|*.m3u"
            .ShowSave
            strFilePath = .Filename
        End With
        'Neu Save
            If strFilePath <> "" Then
                Open strFilePath For Output As #fn
                Print #fn, "#EXTM3U"
                For intRecord = 0 To .lstList1.ListCount - 1
                    Call Mp3.ReadTag(.lstList1.List(intRecord))
                    Print #fn, "#EXTINF:" & "," & Trim(Tag1.Artist) & "_" & Trim(Tag1.Title)
                    Print #fn, .lstList1.List(intRecord)
                    .Enabled = False
                    Next intRecord
            Close #fn
            End If
    Else
         MsgBox "Chuong trinh khong cho phep luu lai mot tap m3u da co san", vbOKOnly, "Thong bao"
         Exit Sub
    End If
    .Enabled = True
    End With
End Sub

