Attribute VB_Name = "mdlBoot"
Option Explicit
' API Get Window Vesions
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
    (lpVersionInformation As OSVersionInfo) As Long
    'Type
    Private Type OSVersionInfo
        OSVSize       As Long
        dwVerMajor    As Long
        dwVerMinor    As Long
        dwBuildNumber As Long
        PlatformID    As Long
        szCSDVersion  As String * 128
    End Type

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long

'API make Opacity form
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
   'Const
    Public Const WS_EX_LAYERED As Long = &H80000
    Public Const LWA_ALPHA As Long = &H2
    Public Const GWL_EXSTYLE = (-20)

'Variables
Public strPathM As String, strPathP As String
Public bShowList As Boolean
Public bShowNumber As Boolean
Public bLClickChoice As Boolean
Public bHidePL As Boolean
Public strOldPathM As String
Public strOldPathP As String
Public bOldShowPl As Boolean, bOldShowNumber As Boolean, bOldLClickChoice As Boolean
Public nAlpha As Integer
Public nVisual As Byte
Public nColorIndex As Byte
Public strRGBbg As String
Public strRGBfg As String
Public strPathSkin As String
Public Sub SaveConfig()
Dim strFileConfig As String
On Error GoTo loi
    strFileConfig = App.Path & "\config.ini"
    Open strFileConfig For Output As #1
        Print #1, "+-------------------------------+"
        Print #1, "+  MediaPHT SystemConfig        +"
        Print #1, "+-------------------------------+ "
        Print #1, "MusicFolder=" & strPathM
        Print #1, "PlaylistFolder=" & strPathP
        Print #1, "Skin=" & strPathSkin
        If bShowList = True Then
            Print #1, "ShowPL=1"
        Else
            Print #1, "ShowPL=0"
        End If
        If bShowNumber = True Then
            Print #1, "ShowNumber=1"
        Else
            Print #1, "ShowNumber=0"
        End If
        If bLClickChoice = True Then
            Print #1, "LeftClickChoice=1"
        Else
            Print #1, "LeftClickChoice=0"
        End If
        Print #1, ""
        Print #1, "[- Memories -]"
        If bHidePL = True Then
            Print #1, "PLHide=1"
        Else
            Print #1, "PLHide=0"
        End If
        Print #1, "ColorStyle=" & nColorIndex
        Print #1, "BG=" & strRGBbg
        Print #1, "FG=" & strRGBfg
        Print #1, "Alpha=" & nAlpha
        Print #1, "Visual=" & nVisual
   Close #1
loi:
    Exit Sub
End Sub

Sub Main()
    Call LoadConfig
    frmSplash.Show
End Sub
Public Sub LoadConfig()
    Dim strPath As String
    On Error GoTo default
    Open App.Path & "\config.ini" For Input As #1
        Do Until EOF(1)
            Line Input #1, strPath
            'Ñoïc ñöôøng daãn chöùa thö muïc Playlist
            If UCase(Mid(strPath, 1, 8)) = "PLAYLIST" Then
                strPathP = Mid(strPath, InStr(1, strPath, "=", vbBinaryCompare) + 1)
            End If
            'Ñoïc ñöôøng daãn chöùa thö muïc Music
            If UCase(Mid(strPath, 1, 5)) = "MUSIC" Then
                strPathM = Mid(strPath, InStr(1, strPath, "=", vbBinaryCompare) + 1)
            End If
            If UCase(Mid(strPath, 1, 4)) = "SKIN" Then
                strPathSkin = Mid(strPath, InStr(1, strPath, "=", vbBinaryCompare) + 1)
            End If
            'Xem caáu hình ñeå khôûi ñoäng
            If UCase(Mid(strPath, 1, 6)) = "PLHIDE" Then
                If Right(strPath, 1) = "1" Then
                    bHidePL = True
                Else
                    bHidePL = False
                End If
            End If
            If UCase(Mid(strPath, 1, 6)) = "SHOWPL" Then
                If Right(strPath, 1) = "1" Then
                    bShowList = True
                Else
                    bShowList = False
                End If
            End If
            If UCase(Mid(strPath, 1, 10)) = "SHOWNUMBER" Then
                If Right(strPath, 1) = "1" Then
                    bShowNumber = True
                Else
                    bShowNumber = False
                End If
            End If
            If UCase(Mid(strPath, 1, 4)) = "LEFT" Then
                If Right(strPath, 1) = "1" Then
                    bLClickChoice = True
                Else
                    bLClickChoice = False
                End If
            End If
            If UCase(Mid(strPath, 1, 5)) = "ALPHA" Then
                nAlpha = CInt(Mid(strPath, 7))
                If nAlpha > 100 Or nAlpha < 0 Then nAlpha = 100
            End If
            If UCase(Mid(strPath, 1, 6)) = "VISUAL" Then
                nVisual = CInt(Mid(strPath, 8))
                If nVisual > 3 Or nVisual < 0 Then nVisual = 0
            End If
            If UCase(Mid(strPath, 1, 10)) = "COLORSTYLE" Then
                nColorIndex = CInt(Mid(strPath, 12))
                If nColorIndex > 5 Or nColorIndex < 0 Then nColorIndex = 0
            End If
            If UCase(Mid(strPath, 1, 2)) = "BG" Then
                strRGBbg = Mid(strPath, 4)
            End If
            If UCase(Mid(strPath, 1, 2)) = "FG" Then
                strRGBfg = Mid(strPath, 4)
            End If
            If strRGBbg = "" Or strRGBfg = "" Then
                nColorIndex = 0
                strRGBbg = "RGB(0, 192, 192)"
                strRGBfg = "RGB(0, 0, 64)"
            End If
       Loop
    Close #1
    Exit Sub
default:
    If Err.Number <> 0 Then
        MsgBox "Error loading config.ini file !!! This program will start with value default", vbOKOnly + vbInformation, "Error"
        strPathP = "C:\"
        strPathM = "C:\"
        strPathSkin = App.Path & "\Skins\Default"
        bHidePL = False
        bShowList = False
        bShowNumber = False
        bLClickChoice = False
        nAlpha = 100
        nVisual = 0
        nColorIndex = 0
        strRGBbg = "RGB(0, 192, 192)"
        strRGBfg = "RGB(0, 0, 64)"
    End If
End Sub
Public Sub LoadSkin(strPathSkin As String)
Dim X As Integer
Dim ctr As Control
'Load image
    With frmMedia
        .imgMedia.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
        .imgVisual.Picture = LoadPicture(strPathSkin & "\Frame\DisplayVisual.jpg")
        .imgTitle.Picture = LoadPicture(strPathSkin & "\Frame\Title.jpg")
        .imgControl.Picture = LoadPicture(strPathSkin & "\Frame\Control.jpg")
        .imgVol2.Picture = LoadPicture(strPathSkin & "\Frame\Vol.jpg")
        .imgMixer.Picture = LoadPicture(strPathSkin & "\Frame\mixer.jpg")
        .pctShuffe.Picture = LoadPicture(strPathSkin & "\Frame\Shuffe.jpg")
        .pctRepeat.Picture = LoadPicture(strPathSkin & "\Frame\Repeat.jpg")
        For X = 0 To .p.Count - 1
            .p(X).Picture = LoadPicture(strPathSkin & "\Frame\colvisual4.jpg")
        Next X
        For X = 0 To .prow.Count - 1
            .prow(X).Picture = LoadPicture(strPathSkin & "\Frame\rowvisual4.jpg")
            .prowM(X).Picture = LoadPicture(strPathSkin & "\Frame\rowvisual4M.jpg")
        Next X
        .pctListUp.Picture = LoadPicture(strPathSkin & "\Frame\listUp.jpg")
        .pctListDown.Picture = LoadPicture(strPathSkin & "\Frame\listDown.jpg")
    'Load Button
        .cmdPrev.Picture = LoadPicture(strPathSkin & "\Button\Pre1.jpg")
        .cmdPrevTrack.Picture = LoadPicture(strPathSkin & "\Button\Previous.jpg")
        .cmdPlay.Picture = LoadPicture(strPathSkin & "\Button\Play.jpg")
        .cmdForw.Picture = LoadPicture(strPathSkin & "\Button\Next1.jpg")
        .cmdForwTrack.Picture = LoadPicture(strPathSkin & "\Button\Next.jpg")
        .cmdPause.Picture = LoadPicture(strPathSkin & "\Button\Pause.jpg")
        .cmdStop.Picture = LoadPicture(strPathSkin & "\Button\Stop.jpg")
        .cmdPower.Picture = LoadPicture(strPathSkin & "\Button\Power.jpg")
        .cmdLoad.Picture = LoadPicture(strPathSkin & "\Button\Add.jpg")
        .cmdSub.Picture = LoadPicture(strPathSkin & "\Button\Sub.jpg")
        .cmdSave.Picture = LoadPicture(strPathSkin & "\Button\Save.jpg")
        For Each ctr In .Controls
            If TypeOf ctr Is Image Then
                If Mid(ctr.Name, 1, 3) = "cmd" Then ctr.MousePointer = 10
            End If
        Next
    .Refresh
    End With
End Sub
Sub RefreshSkin()
    With frmTagID
        If .bShow = True Then
            .imgTagID.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
        End If
    End With
    With frmRate
        .imgRate.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
    End With
    With frmInput
        .imgInput.Picture = LoadPicture(strPathSkin & "\Frame\Control.jpg")
    End With
    With frmAlphaCustom
        .imgAlpha.Picture = LoadPicture(strPathSkin & "\Frame\Control.jpg")
    End With
    With frmVisual
        .imgVisual.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
    End With
    frmMedia.Enabled = True
End Sub
' This sub write by RAUL MARTINEZ HERNANDEZ _ Developer of MMPlayerX
Sub Make_Transparent(hWnd As Long, Percent As Integer)
 On Error Resume Next
  Dim OSV As OSVersionInfo
  OSV.OSVSize = Len(OSV)
  If GetVersionEx(OSV) <> 1 Then Exit Sub
  If OSV.PlatformID = 1 And OSV.dwVerMinor >= 10 Then Exit Sub '/* Win 98/ME
  If OSV.PlatformID = 2 And OSV.dwVerMajor >= 5 Then '/* Win 2000/XP
    Call SetWindowLong(hWnd, GWL_EXSTYLE, GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(hWnd, 0, (Percent * 255) / 100, LWA_ALPHA)
  End If
Exit Sub
End Sub

