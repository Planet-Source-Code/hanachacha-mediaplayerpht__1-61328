VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMedia 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   7470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmMedia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pctBalance 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   105
      Left            =   3000
      ScaleHeight     =   105
      ScaleWidth      =   1005
      TabIndex        =   47
      Top             =   1125
      Width           =   1000
      Begin VB.HScrollBar hscVol 
         Height          =   100
         Left            =   -240
         Max             =   5000
         Min             =   -5000
         TabIndex        =   52
         Top             =   0
         Width           =   1480
      End
   End
   Begin VB.PictureBox pctRepeat 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4350
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   46
      Top             =   720
      Width           =   300
   End
   Begin VB.PictureBox pctShuffe 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   45
      Top             =   720
      Width           =   300
   End
   Begin VB.PictureBox pctPosition 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   100
      Left            =   150
      ScaleHeight     =   105
      ScaleWidth      =   4605
      TabIndex        =   44
      Top             =   1800
      Width           =   4600
      Begin MSForms.ScrollBar hscPosition 
         Height          =   100
         Left            =   -240
         TabIndex        =   51
         Top             =   0
         Width           =   5085
         ForeColor       =   65280
         BackColor       =   16761024
         Size            =   "8969;176"
         Max             =   100
         Orientation     =   1
      End
   End
   Begin VB.PictureBox pctInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   2535
      TabIndex        =   43
      Top             =   1300
      Width           =   2535
      Begin MSForms.Label lblTitle 
         Height          =   375
         Left            =   0
         TabIndex        =   50
         Top             =   180
         Width           =   3000
         ForeColor       =   12632064
         BackColor       =   12648384
         VariousPropertyBits=   8388627
         Size            =   "5292;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblArtist 
         Height          =   375
         Left            =   0
         TabIndex        =   49
         Top             =   0
         Width           =   3000
         ForeColor       =   12632064
         BackColor       =   12648384
         VariousPropertyBits=   8388627
         Size            =   "5292;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.ListBox lstList1 
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pctListUp 
      Height          =   375
      Left            =   4350
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   2640
      Width           =   375
   End
   Begin VB.PictureBox pctListDown 
      Height          =   375
      Left            =   4350
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   6690
      Width           =   375
   End
   Begin VB.PictureBox pctList 
      BackColor       =   &H00000000&
      Height          =   4440
      Left            =   4350
      ScaleHeight     =   4380
      ScaleWidth      =   315
      TabIndex        =   8
      Top             =   2625
      Width           =   375
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   3480
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   8040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton optTimer 
      Height          =   375
      Left            =   3360
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optShuffe 
      Height          =   375
      Left            =   3360
      TabIndex        =   40
      Top             =   5400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optRepeat 
      Height          =   375
      Left            =   3360
      TabIndex        =   39
      Top             =   5040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.OptionButton optCont 
      Height          =   375
      Left            =   3360
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   9
      Left            =   2400
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   32
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   8
      Left            =   2175
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   31
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   1935
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   30
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   6
      Left            =   1710
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   1485
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   1260
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   27
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   1035
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   26
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   810
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   585
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   24
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox pM 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   360
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   23
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   9
      Left            =   2400
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   8
      Left            =   2175
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   7
      Left            =   1935
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   6
      Left            =   1710
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   5
      Left            =   1485
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   4
      Left            =   1260
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   3
      Left            =   1035
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   2
      Left            =   810
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   15
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   1
      Left            =   585
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox p 
      BorderStyle     =   0  'None
      Height          =   600
      Index           =   0
      Left            =   360
      ScaleHeight     =   600
      ScaleWidth      =   150
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Timer Timer5 
      Interval        =   500
      Left            =   3120
      Top             =   3960
   End
   Begin VB.Timer TimerListDown 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2760
      Top             =   6720
   End
   Begin VB.Timer TimerListUp 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2760
      Top             =   6240
   End
   Begin VB.Timer timerVisual 
      Interval        =   100
      Left            =   2640
      Top             =   5640
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   3000
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3600
      Top             =   3360
   End
   Begin MSComDlg.CommonDialog cdloSave 
      Left            =   4680
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cdloLoad 
      Left            =   4200
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox lstList 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   4440
      IntegralHeight  =   0   'False
      Left            =   220
      TabIndex        =   1
      Top             =   2625
      Width           =   4455
   End
   Begin VB.PictureBox prow 
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   2205
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.PictureBox prow 
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   2205
      TabIndex        =   33
      Top             =   720
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.PictureBox prowM 
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   1
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   2205
      TabIndex        =   36
      Top             =   960
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.PictureBox prowM 
      BorderStyle     =   0  'None
      Height          =   135
      Index           =   0
      Left            =   360
      ScaleHeight     =   135
      ScaleWidth      =   2205
      TabIndex        =   35
      Top             =   720
      Visible         =   0   'False
      Width           =   2200
   End
   Begin VB.PictureBox pctPlaylist 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   4000
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   48
      Top             =   1080
      Width           =   300
   End
   Begin VB.Image cmdSave 
      Height          =   195
      Left            =   4320
      Top             =   7180
      Width           =   405
   End
   Begin VB.Image cmdSub 
      Height          =   195
      Left            =   600
      Top             =   7180
      Width           =   405
   End
   Begin VB.Image cmdLoad 
      Height          =   200
      Left            =   200
      Top             =   7180
      Width           =   400
   End
   Begin VB.Image cmdPower 
      Height          =   255
      Left            =   3900
      ToolTipText     =   "Power"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdStop 
      Height          =   255
      Left            =   3420
      ToolTipText     =   "Stop"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdPause 
      Height          =   255
      Left            =   2940
      ToolTipText     =   "Pause"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdForwTrack 
      Height          =   255
      Left            =   2460
      ToolTipText     =   "Next track"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdForw 
      Height          =   255
      Left            =   1980
      ToolTipText     =   "Next 5 seconds"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdPlay 
      Height          =   255
      Left            =   1500
      ToolTipText     =   "Play"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdPrev 
      Height          =   255
      Left            =   1020
      ToolTipText     =   "Back 5 seconds"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image cmdPrevTrack 
      Height          =   255
      Left            =   540
      ToolTipText     =   "Back track"
      Top             =   2100
      Width           =   495
   End
   Begin VB.Image imgMixer 
      Height          =   300
      Left            =   4350
      Picture         =   "frmMedia.frx":000C
      Stretch         =   -1  'True
      Top             =   1080
      Width           =   300
   End
   Begin VB.Shape shpMonoR 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   105
      Left            =   3860
      Shape           =   4  'Rounded Rectangle
      Top             =   1250
      Width           =   135
   End
   Begin VB.Shape shpMonoL 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   105
      Left            =   3020
      Shape           =   4  'Rounded Rectangle
      Top             =   1250
      Width           =   135
   End
   Begin VB.Label lblTitleBar 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MediaPlayerPHT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   1800
      TabIndex        =   37
      Top             =   75
      Width           =   1425
   End
   Begin VB.Label lblDuration 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3550
      TabIndex        =   12
      Top             =   800
      Width           =   375
   End
   Begin VB.Label lblPosition 
      BackStyle       =   0  'Transparent
      Caption         =   "0:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   800
      Width           =   600
   End
   Begin VB.Label lblVol 
      Caption         =   "100"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgVol2 
      Height          =   105
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   1500
      Width           =   1000
   End
   Begin VB.Label lblVolInc 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4373
      TabIndex        =   6
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblVolDec 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4080
      TabIndex        =   5
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   4350
      TabIndex        =   4
      Top             =   45
      Width           =   255
   End
   Begin VB.Label lblMini 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   45
      Width           =   255
   End
   Begin VB.Image imgMedia 
      Height          =   5070
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   4965
   End
   Begin VB.Image imgTitle 
      Height          =   375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4965
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   8400
      Visible         =   0   'False
      Width           =   1095
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Image imgVisual 
      Height          =   1695
      Left            =   0
      Top             =   360
      Width           =   4965
   End
   Begin VB.Image imgControl 
      Height          =   375
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   4965
   End
   Begin VB.Menu mnuMedia 
      Caption         =   "Media"
      Visible         =   0   'False
      Begin VB.Menu mnuInfor 
         Caption         =   "A        About......."
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShowPL 
         Caption         =   "Playlist"
      End
      Begin VB.Menu mnuPL 
         Caption         =   "Playback"
         Begin VB.Menu mnuVolume 
            Caption         =   "Volume"
            Begin VB.Menu mnuVol 
               Caption         =   "+        Inc"
               Index           =   0
            End
            Begin VB.Menu mnuVol 
               Caption         =   " -         Dec"
               Index           =   1
            End
         End
         Begin VB.Menu mnuBar5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "D        Top"
            Index           =   0
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "B        Previous track"
            Index           =   1
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "<--     Prev 5 seconds"
            Index           =   2
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "N        Next track"
            Index           =   3
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "-->     Next 5 secconds"
            Index           =   4
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "C        Bottom"
            Index           =   5
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "P        Pause"
            Index           =   6
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   """ ""      Stop"
            Index           =   7
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "-"
            Index           =   8
         End
         Begin VB.Menu mnuMediaControl 
            Caption         =   "Jump to  ..."
            Index           =   9
         End
      End
      Begin VB.Menu mnuType 
         Caption         =   "Type"
         Begin VB.Menu mnuTypeC 
            Caption         =   "S             Shuffe"
            Index           =   0
         End
         Begin VB.Menu mnuTypeC 
            Caption         =   "R             Repeat One"
            Index           =   1
         End
         Begin VB.Menu mnuTypeC 
            Caption         =   "Speed ..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add"
         Begin VB.Menu mnuAddChild 
            Caption         =   "Add File ..."
            Index           =   0
         End
         Begin VB.Menu mnuAddChild 
            Caption         =   "Add folder ..."
            Index           =   1
         End
         Begin VB.Menu mnuAddChild 
            Caption         =   "Add Playlist ..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Sub"
         Begin VB.Menu mnuSubChild 
            Caption         =   "Remove selected file"
            Index           =   0
         End
         Begin VB.Menu mnuSubChild 
            Caption         =   "Remove All"
            Index           =   1
         End
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "Display ..."
         Begin VB.Menu mnuSkin 
            Caption         =   "Skin"
            Begin VB.Menu mnuColor 
               Caption         =   "Theme"
               Begin VB.Menu mnuColorStyle 
                  Caption         =   "Default"
                  Index           =   0
               End
               Begin VB.Menu mnuColorStyle 
                  Caption         =   "Black_White"
                  Index           =   1
               End
               Begin VB.Menu mnuColorStyle 
                  Caption         =   "Sliver_Black"
                  Index           =   2
               End
               Begin VB.Menu mnuColorStyle 
                  Caption         =   "Sliver_Brown"
                  Index           =   3
               End
               Begin VB.Menu mnuColorStyle 
                  Caption         =   "Magenta_Blue"
                  Index           =   4
               End
               Begin VB.Menu mnuColorStyle 
                  Caption         =   "Orange_Blue"
                  Index           =   5
               End
            End
            Begin VB.Menu mnuTransparent 
               Caption         =   "Opacity"
               Begin VB.Menu mnuAlpha 
                  Caption         =   "100 %"
                  Index           =   0
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "90 %"
                  Index           =   1
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "80 %"
                  Index           =   2
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "70 %"
                  Index           =   3
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "60 %"
                  Index           =   4
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "50 %"
                  Index           =   5
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "40 %"
                  Index           =   6
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "30 %"
                  Index           =   7
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "20 %"
                  Index           =   8
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "10 %"
                  Index           =   9
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "-"
                  Index           =   10
               End
               Begin VB.Menu mnuAlpha 
                  Caption         =   "Custom ..."
                  Index           =   11
               End
            End
            Begin VB.Menu mnuTitle 
               Caption         =   "Scroll Title"
            End
            Begin VB.Menu mnuNewSkin 
               Caption         =   "More Skin ..."
               Begin VB.Menu mnuMoreSkin 
                  Caption         =   "0 - Default"
                  Index           =   0
               End
               Begin VB.Menu mnuMoreSkin 
                  Caption         =   "1 - Black"
                  Index           =   1
               End
               Begin VB.Menu mnuMoreSkin 
                  Caption         =   "2 - Plastic"
                  Index           =   2
               End
               Begin VB.Menu mnuMoreSkin 
                  Caption         =   "3 - Rubi"
                  Index           =   3
               End
            End
         End
         Begin VB.Menu mnuSpectrum 
            Caption         =   "Spectrum"
            Begin VB.Menu mnuSpectrumC 
               Caption         =   "0-None"
               Index           =   0
            End
            Begin VB.Menu mnuSpectrumC 
               Caption         =   "1-Col"
               Index           =   1
            End
            Begin VB.Menu mnuSpectrumC 
               Caption         =   "2-Row"
               Index           =   2
            End
            Begin VB.Menu mnuSpectrumC 
               Caption         =   "3-Visual Extend"
               Index           =   3
            End
         End
         Begin VB.Menu mnuConfig 
            Caption         =   "Config ..."
         End
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "X        Exit"
      End
   End
   Begin VB.Menu mnuCmdAdd 
      Caption         =   "Add"
      Visible         =   0   'False
      Begin VB.Menu mnuCmdAddC 
         Caption         =   "Add file ..."
         Index           =   0
      End
      Begin VB.Menu mnuCmdAddC 
         Caption         =   "Add folder ..."
         Index           =   1
      End
      Begin VB.Menu mnuCmdAddC 
         Caption         =   "Add playlist ... "
         Index           =   2
      End
   End
   Begin VB.Menu mnuCmdSub 
      Caption         =   "Sub"
      Visible         =   0   'False
      Begin VB.Menu mnuSubMedia 
         Caption         =   "Remove selected file"
         Index           =   0
      End
      Begin VB.Menu mnuSubMedia 
         Caption         =   "Remove All"
         Index           =   1
      End
      Begin VB.Menu mnuSubMedia 
         Caption         =   "Delete selected file"
         Index           =   2
      End
   End
   Begin VB.Menu mnuForm 
      Caption         =   "System"
      Visible         =   0   'False
      Begin VB.Menu mnuFormChild 
         Caption         =   "         Minimize"
         Index           =   0
      End
      Begin VB.Menu mnuFormChild 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFormChild 
         Caption         =   "X       Close"
         Index           =   2
      End
   End
   Begin VB.Menu mnuTrayMedia 
      Caption         =   "Media"
      Visible         =   0   'False
      Begin VB.Menu mnuTrayType 
         Caption         =   "Type"
         Begin VB.Menu mnuPlaytype 
            Caption         =   "Shuffe"
            Index           =   0
         End
         Begin VB.Menu mnuPlaytype 
            Caption         =   "Repeat One"
            Index           =   1
         End
         Begin VB.Menu mnuPlaytype 
            Caption         =   "Speed ..."
            Index           =   2
         End
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Top"
         Index           =   0
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Previous track"
         Index           =   1
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Prev 5 seconds"
         Index           =   2
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Next track"
         Index           =   3
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Next 5 seconds"
         Index           =   4
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Bottom"
         Index           =   5
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Pause"
         Index           =   6
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Stop"
         Index           =   7
      End
      Begin VB.Menu mnuTray 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Return"
         Index           =   9
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Exit"
         Index           =   10
      End
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu mnuEditPL 
         Caption         =   "View file information ..."
         Index           =   0
      End
      Begin VB.Menu mnuEditPL 
         Caption         =   "Play"
         Index           =   1
      End
      Begin VB.Menu mnuEditPL 
         Caption         =   "Remove selected file"
         Index           =   2
      End
      Begin VB.Menu mnuEditPL 
         Caption         =   "Remove All"
         Index           =   3
      End
      Begin VB.Menu mnuEditPL 
         Caption         =   "Delete selected file"
         Index           =   4
      End
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public allowplay As String
Public allowpause As Boolean
Public paused As Boolean
Dim nPosition As Long
Dim X, i As Integer
Dim bRDown As Boolean
Dim ctr As Control
Sub KillFile(strFileName As String)
    If lstList.ListCount = 0 Or lstList.ListIndex = -1 Then
        Exit Sub
    End If
    ReadTag (strFileName)
    If MsgBox("Are you sure you want to deleted " & " '' " & Trim(Tag1.Title) & " '' ", vbYesNo, "Thong bao") = vbYes Then
        Kill (strFileName)
        Call SubFile
    End If
End Sub
Public Sub VisualNone()
            For X = 0 To prow.Count - 1
                prow(X).Visible = False
                prowM(X).Visible = False
            Next X
            For X = 0 To pM.Count - 1
                p(X).Visible = False
                pM(X).Visible = False
            Next X
            mnuSpectrumC(3).Enabled = False
            If frmVisual.bShow = True Then Unload frmVisual
End Sub
Public Sub VisualRow()
Randomize MediaPlayer1.AudioStream
            For X = 0 To prow.Count - 1
                prow(X).Visible = True
                prowM(X).Visible = True
                prow(X).Width = Int(2200 * Rnd)
            Next X
End Sub
Public Sub VisualRain()
Randomize MediaPlayer1.AudioStream
            For X = 0 To pM.Count - 1
                p(X).Visible = True
                pM(X).Visible = True
                pM(X).Height = 600 * Rnd
            Next X
End Sub

Private Sub cmdForw_Click()
    Call Mp3.Forw
End Sub

Private Sub cmdForw_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdForw.Picture = LoadPicture(strPathSkin & "\Button\Next1Down.jpg")
End Sub

Private Sub cmdForw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdForw.Picture = LoadPicture(strPathSkin & "\Button\Next1.jpg")
End Sub

Private Sub cmdForwTrack_Click()
    Call Mp3.NextTrack
End Sub

Private Sub cmdForwTrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdForwTrack.Picture = LoadPicture(strPathSkin & "\Button\NextDown.jpg")
End Sub

Private Sub cmdForwTrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdForwTrack.Picture = LoadPicture(strPathSkin & "\Button\Next.jpg")
End Sub

Private Sub cmdLoad_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdLoad.Picture = LoadPicture(strPathSkin & "\Button\AddDown.jpg")
End Sub

Private Sub cmdLoad_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        PopupMenu mnuCmdAdd, vbPopupMenuLeftAlign
    End If
    cmdLoad.Picture = LoadPicture(strPathSkin & "\Button\Add.jpg")
End Sub

Private Sub cmdPause_Click()
On Error Resume Next
If allowpause = True Then
    On Error Resume Next
    If paused = False Then
        hscPosition.Value = MediaPlayer1.CurrentPosition
        MediaPlayer1.Pause
        paused = True
        allowplay = "no"
        mnuMediaControl(6).Checked = True
        mnuTray(6).Checked = True
        Exit Sub
    End If
    If paused = True Then
        hscPosition.Value = MediaPlayer1.CurrentPosition
        MediaPlayer1.Play
        paused = False
        allowplay = "yes"
        mnuMediaControl(6).Checked = False
        mnuTray(6).Checked = False
    End If
End If
End Sub

Private Sub cmdPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPause.Picture = LoadPicture(strPathSkin & "\Button\PauseDown.jpg")
End Sub

Private Sub cmdPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPause.Picture = LoadPicture(strPathSkin & "\Button\Pause.jpg")
End Sub

Private Sub cmdPlay_Click()
On Error Resume Next
    If lstList.ListIndex = -1 Or lstList.ListCount = 0 Then
        MsgBox "Khong co bai hat nao duoc chon", vbOKOnly, "Thong bao"
    Else
        If paused = True Then
            hscPosition.Value = MediaPlayer1.CurrentPosition
            MediaPlayer1.Play
            lblDuration.Caption = MediaPlayer1.Duration / 60
            allowplay = "yes"
            allowpause = True
            paused = False
            mnuMediaControl(6).Checked = False
            mnuTray(6).Checked = False
            Exit Sub
        End If
        If paused = False Then
            hscPosition.Max = MediaPlayer1.Duration
            txtFileName.Text = lstList1.List(lstList1.ListIndex)
            Call Mp3.Play(txtFileName.Text)
            allowplay = "yes"
            allowpause = False
            paused = True
            mnuMediaControl(6).Checked = True
            mnuTray(6).Checked = True
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdPlay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPlay.Picture = LoadPicture(strPathSkin & "\Button\PlayDown.jpg")
End Sub

Private Sub cmdPlay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPlay.Picture = LoadPicture(strPathSkin & "\Button\Play.jpg")
End Sub

Private Sub cmdPower_Click()
        If allowplay = "yes" Or paused = False Then
            MediaPlayer1.Stop
        End If
            Unload Me
End Sub

Private Sub cmdPower_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPower.Picture = LoadPicture(strPathSkin & "\Button\PowerDown.jpg")
End Sub

Private Sub cmdPower_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPower.Picture = LoadPicture(strPathSkin & "\Button\Power.jpg")
End Sub

Private Sub cmdPrev_Click()
    Call Mp3.Prev
End Sub

Private Sub cmdPrev_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrev.Picture = LoadPicture(strPathSkin & "\Button\Pre1Down.jpg")
End Sub

Private Sub cmdPrev_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrev.Picture = LoadPicture(strPathSkin & "\Button\Pre1.jpg")
End Sub

Private Sub cmdPrevTrack_Click()
    Call Mp3.BackTrack
End Sub

Private Sub cmdPrevTrack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrevTrack.Picture = LoadPicture(strPathSkin & "\Button\PreviousDown.jpg")
End Sub

Private Sub cmdPrevTrack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdPrevTrack.Picture = LoadPicture(strPathSkin & "\Button\Previous.jpg")
End Sub

Private Sub cmdSave_Click()
    Call SavePlaylist
End Sub

Private Sub cmdSave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        cmdSave.Picture = LoadPicture(strPathSkin & "\Button\SaveDown.jpg")
End Sub

Private Sub cmdSave_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        cmdSave.Picture = LoadPicture(strPathSkin & "\Button\Save.jpg")
End Sub

Private Sub cmdStop_Click()
    Call StopPlayer
End Sub

Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdStop.Picture = LoadPicture(strPathSkin & "\Button\StopDown.jpg")
End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdStop.Picture = LoadPicture(strPathSkin & "\Button\Stop.jpg")
End Sub

Private Sub cmdSub_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdSub.Picture = LoadPicture(strPathSkin & "\Button\SubDown.jpg")
End Sub

Private Sub cmdSub_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
                PopupMenu mnuCmdSub, vbPopupMenuLeftAlign
    End If
    cmdSub.Picture = LoadPicture(strPathSkin & "\Button\Sub.jpg")
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 65 Then Call mnuInfor_Click 'A
    If KeyCode = 68 Then Call GoTop 'D
    If KeyCode = 66 Then Call Mp3.BackTrack   'B
    If KeyCode = 37 Then Call Mp3.Prev   'Left <--
    If KeyCode = 39 Then Call Mp3.Forw   'Right -->
    If KeyCode = 78 Then Call Mp3.NextTrack   'N
    If KeyCode = 67 Then Call GoBottom 'C
    If KeyCode = 80 Then Call cmdPause_Click 'P
    If KeyCode = 32 Then Call StopPlayer 'Spacebar
    If KeyCode = 107 Then Call Mp3.VolInc   '+
    If KeyCode = 109 Then Call Mp3.VolDec   '-
    If KeyCode = 83 Then Call mnuTypeC_Click(1) 'S
    If KeyCode = 82 Then Call mnuTypeC_Click(2)  'R
    If KeyCode = 84 Then Call mnuTypeC_Click(0)  'T
    If KeyCode = 88 Then Call mnuExit_Click 'X
End Sub

Private Sub Form_Load()
On Error Resume Next
    Call mdlBoot.LoadSkin(strPathSkin)
    For i = 0 To 3
        If Mid(mnuMoreSkin(i).Caption, 5) = Right(strPathSkin, Len(Mid(mnuMoreSkin(i).Caption, 5))) Then
            mnuMoreSkin(i).Checked = True
        End If
    Next i
    Call mdlBoot.Make_Transparent(frmMedia.hWnd, nAlpha)
    For i = 0 To mnuAlpha.Count - 1
        If Val(Trim(Mid(mnuAlpha(i).Caption, 1, 3))) = nAlpha Then
            mnuAlpha(i).Checked = True
        End If
    Next i
    mnuColorStyle(nColorIndex).Checked = True
    For Each ctr In Me.Controls
        If TypeOf ctr Is Label Or TypeOf ctr Is ListBox Then
            If ctr.Name = "lblMini" Or ctr.Name = "lblClose" Or ctr.Name = "lstList" Then
                ctr.BackColor = strRGBbg
                ctr.ForeColor = strRGBfg
            End If
        End If
    Next ctr
    '
    mnuShowPL.Checked = Not bHidePL
    If Not bHidePL Then
        pctPlaylist.Picture = LoadPicture(strPathSkin & "\Button\PLactive.jpg")
    Else
        pctPlaylist.Picture = LoadPicture(strPathSkin & "\Button\PL.jpg")
    End If
    Me.Icon = LoadPicture(App.Path & "\Skins\Default\CD.ico")
    allowpause = False
    paused = False
    allowplay = "no"
    'Visuallization
    timerVisual.Enabled = True
    mnuSpectrumC(nVisual).Checked = True
    If nVisual = 3 Then frmVisual.Show
    'Gia tri mac nhien
    mnuTypeC(0).Checked = True
    mnuPlaytype(0).Checked = True
    pctShuffe.Picture = LoadPicture(App.Path & "\Graphics\Nen\Shuffe.jpg")
    '
    If bShowList = True Then
        'Lay playlist cuoi cung
        Call LoadPlaylistM3U(strPathP & "\Lastlist.m3u")
        lstList.ListIndex = 0
        lstList1.ListIndex = lstList.ListIndex
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Err.Clear
Static bBusy As Boolean
    If bBusy = False Then
        bBusy = True
        Select Case CLng(X / 15)
            Case WM_LBUTTONDBLCLK
                frmMedia.WindowState = vbNormal
                Call mdlSys.RemoveIcon
                DoEvents
            Case WM_LBUTTONDOWN
            Case WM_LBUTTONUP
            Case WM_RBUTTONDBLCLK
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
                'Right mouse button released: display popup menu
                PopupMenu mnuTrayMedia, 2, , , mnuTray(10)
        End Select
        bBusy = False
    End If
End Sub

Private Sub Form_Resize()
Dim nTop, nLeft As Integer
On Error Resume Next
    Me.Width = 4965
    If bHidePL = False Then Me.Height = 7500
    If bHidePL = True Then Me.Height = 7500 - Me.imgMedia.Height
    If Me.WindowState = vbMinimized Then
        Me.Caption = lblTitle.Caption
    Else
        Me.Caption = ""
        Call mdlSys.RemoveIcon
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo loi
Dim strFilePath As String
Dim intRecord As Integer
Dim fn As Integer
fn = FreeFile
Me.Visible = False
strFilePath = strPathP & "\Lastlist.m3u"
    If lstList1.ListCount > 0 Then
        If Mid(lstList1.List(0), 2, 2) = ":\" Then
            Open strFilePath For Output As #fn
                Print #fn, "#EXTM3U"
                For intRecord = 0 To lstList1.ListCount - 1
                        Call Mp3.ReadTag(lstList1.List(intRecord))
                        Print #fn, "#EXTINF:" & "," & Trim(Tag1.Artist) & "_" & Trim(Tag1.Title)
                        Print #fn, lstList1.List(intRecord)
                Next intRecord
            Close #fn
        End If
    End If
    Call mdlBoot.SaveConfig
    If frmVisual.Visible = True Then
        Unload frmVisual
    End If
    If frmRate.Visible = True Then
        Unload frmRate
    End If
    End
loi:
    Exit Sub
End Sub
Private Sub imgMedia_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton Then
    Exit Sub
        Else
    If Button = vbMiddleButton Then
    Exit Sub
        Else
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
    End If
    End If
End Sub

Private Sub imgMedia_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMedia, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub imgMixer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
          Shell "sndvol32.exe", vbNormalFocus
    End If
End Sub

Private Sub imgTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbRightButton Then
    Exit Sub
        Else
    If Button = vbMiddleButton Then
    Exit Sub
        Else
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
    End If
    End If
End Sub

Private Sub imgTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuForm, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub imgVisual_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = vbRightButton Then
        Exit Sub
    Else
            If Button = vbMiddleButton Then
                Exit Sub
            Else
                ReleaseCapture
                SendMessage Me.hWnd, &HA1, 2, 0&
            End If
    End If
End Sub

Private Sub imgVisual_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMedia, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub lblArtist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Exit Sub
    Else
        If Button = vbMiddleButton Then
            Exit Sub
        Else
            ReleaseCapture
            SendMessage Me.hWnd, &HA1, 2, 0&
        End If
    End If
End Sub

Private Sub lblArtist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMedia, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub lblClose_Click()
        If allowplay = "yes" Or paused = False Then
            MediaPlayer1.Stop
        End If
    Unload Me
End Sub

Private Sub lblMini_Click()
On Error Resume Next
    Me.WindowState = vbMinimized
    Call mdlSys.AddIcon(frmMedia, mnuTrayMedia)
    Call mdlSys.SysTip("MediaPHT 1.0 _ " & lblTitle.Caption)
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    Exit Sub
    Else
        If Button = vbMiddleButton Then
            Exit Sub
        Else
            ReleaseCapture
            SendMessage Me.hWnd, &HA1, 2, 0&
        End If
    End If
End Sub

Private Sub lblTitle_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMedia, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub lblVolDec_Click()
    Call Mp3.VolDec
End Sub

Private Sub lblVolInc_Click()
    Call Mp3.VolInc
End Sub

Private Sub lblPosition_Click()
    If optTimer.Value = False Then
        optTimer.Value = True
    Else
        optTimer.Value = False
    End If
End Sub
Private Sub lstList_Click()
    If bRDown = False Then
        If bLClickChoice = True Then
            lstList1.ListIndex = lstList.ListIndex
            txtFileName.Text = lstList1.Text
        End If
    End If
End Sub

Private Sub lstList_DblClick()
On Error GoTo loi
    allowplay = "yes"
    Timer1.Enabled = True
    lstList1.ListIndex = lstList.ListIndex
    txtFileName.Text = lstList1.Text
    Call Mp3.Play(txtFileName.Text)
    allowpause = True
    paused = False
    mnuMediaControl(6).Checked = False
    mnuTray(6).Checked = False
loi:
    If Err.Number = 380 Then
        MsgBox "Error !!!Can't find filename _ Please remove this file  ", vbOKOnly, "Error"
        Exit Sub
    End If
End Sub
Private Sub lstList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    On Error GoTo loi
        allowplay = "yes"
        Timer1.Enabled = True
        lstList1.ListIndex = lstList.ListIndex
        txtFileName.Text = lstList1.Text
        Call Mp3.Play(txtFileName.Text)
        allowpause = True
        paused = False
        mnuMediaControl(6).Checked = False
        mnuTray(6).Checked = False
loi:
    If Err.Number = 380 Then
        MsgBox "Error !!!Can't find filename _ Please remove this file  ", vbOKOnly, "Error"
        Exit Sub
    End If
End If
    If KeyAscii = 32 Then Call StopPlayer
End Sub

Private Sub lstList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        bRDown = True
    Else
        bRDown = False
    End If
End Sub

Private Sub lstList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        If bRDown = True Then
            If X > 0 And X < lstList.Width Then
                If Y > 0 And Y < lstList.Height Then
                ' 247 = lstList / Total file view
                    lstList.ListIndex = lstList.TopIndex + (Y \ 247)
                    PopupMenu mnuList, vbPopupMenuLeftAlign
                End If
            End If
        End If
    End If
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)
On Error Resume Next
    allowplay = "no"
    allowpause = False
    If mnuTypeC(1).Checked = False Then
         If mnuTypeC(0).Checked = True Then
                allowpause = True
                Randomize lstList1.ListCount
                    lstList1.ListIndex = Int(lstList1.ListCount * Rnd)
                    txtFileName.Text = lstList1.Text
                    lstList.ListIndex = lstList1.ListIndex
                    Call Mp3.Play(txtFileName.Text)
            Else
                allowpause = True
                If lstList1.ListIndex < lstList1.ListCount Then
                    lstList1.ListIndex = lstList1.ListIndex + 1
                    txtFileName.Text = lstList1.Text
                    lstList.ListIndex = lstList1.ListIndex
                    Call Mp3.Play(txtFileName.Text)
                    Exit Sub
                Else
                    lstList1.ListIndex = 0
                    txtFileName.Text = lstList1.Text
                    lstList.ListIndex = lstList1.ListIndex
                    Call Mp3.Play(txtFileName.Text)
                    Exit Sub
                End If
            End If
    Else
        allowpause = True
        lstList1.ListIndex = lstList1.ListIndex
        txtFileName.Text = lstList1.Text
        lstList.ListIndex = lstList1.ListIndex
        Call Mp3.Play(txtFileName.Text)
    End If
End Sub
Private Sub MediaPlayer1_NewStream()
    allowplay = "yes"
    allowpause = True
    lblTitle.Left = lblArtist.Left
End Sub
Private Sub mnuAddChild_Click(Index As Integer)
    Select Case Index
        Case 0
            Call AddFile
            Exit Sub
        Case 1
            frmDir.Show
            Exit Sub
        Case 2
            Call AddPl
            Exit Sub
        End Select
End Sub

Private Sub mnuAlpha_Click(Index As Integer)
On Error Resume Next
If Index <> 11 Then
    nAlpha = Val(Trim(Mid(mnuAlpha(Index).Caption, 1, 3)))
    Call mdlBoot.Make_Transparent(frmMedia.hWnd, nAlpha)
    mnuAlpha(Index).Checked = True
Else
    frmAlphaCustom.Show
End If
For i = 0 To 11
    If i <> Index Then
        mnuAlpha(i).Checked = False
    End If
Next i
End Sub

Private Sub mnuCmdAddC_Click(Index As Integer)
    Select Case Index
        Case 0
            Call AddFile
            Exit Sub
        Case 1
            frmDir.Show
            Exit Sub
        Case 2
            Call AddPl
            Exit Sub
        End Select
End Sub

Private Sub mnuColorStyle_Click(Index As Integer)
Dim strRGBbc, strRGBfc As String
    mnuColorStyle(Index).Checked = True
    For i = 0 To 5
        If i <> Index Then
            mnuColorStyle(i).Checked = False
        End If
    Next i
    Select Case Index
        Case Is = 0
            strRGBbc = RGB(0, 192, 192)
            strRGBfc = RGB(0, 0, 64)
        Case Is = 1
            strRGBbc = RGB(0, 0, 0)
            strRGBfc = RGB(255, 255, 255)
        Case Is = 2
            strRGBbc = RGB(204, 213, 211)
            strRGBfc = RGB(0, 0, 0)
        Case Is = 3
            strRGBbc = RGB(223, 236, 236)
            strRGBfc = RGB(111, 7, 7)
        Case Is = 4
            strRGBbc = RGB(220, 52, 255)
            strRGBfc = RGB(52, 105, 255)
        Case Is = 5
            strRGBbc = RGB(255, 152, 52)
            strRGBfc = RGB(55, 252, 224)
    End Select
    For Each ctr In Me.Controls
        If TypeOf ctr Is Label Or TypeOf ctr Is ListBox Then
            If ctr.Name = "lblMini" Or ctr.Name = "lblClose" Or ctr.Name = "lstList" Then
                ctr.BackColor = strRGBbc
                ctr.ForeColor = strRGBfc
            End If
        End If
    Next ctr
    nColorIndex = Index
    strRGBbg = strRGBbc
    strRGBfg = strRGBfc
End Sub

Private Sub mnuConfig_Click()
    frmOption.Show
End Sub
Private Sub mnuEditPL_Click(Index As Integer)
    Select Case Index
        Case 0
            frmTagID.Show
            Exit Sub
        Case 1
            Call Mp3.Play(lstList1.List(lstList.ListIndex))
            allowplay = "yes"
            Exit Sub
        Case 2
            Call SubFile
            Exit Sub
        Case 3
            Call SubFolder
            Exit Sub
        Case 4
            Call KillFile(lstList1.List(lstList.ListIndex))
            Exit Sub
    End Select
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub


Private Sub mnuFormChild_Click(Index As Integer)
    Select Case Index
        Case 0
            On Error Resume Next
            Me.WindowState = vbMinimized
            Call mdlSys.AddIcon(frmMedia, mnuTrayMedia)
            Call mdlSys.SysTip("MediaPHT 1.0 _ " & lblTitle.Caption)
            Exit Sub
        Case 1
            Exit Sub
        Case 2
            Unload Me
    End Select
End Sub

Private Sub mnuInfor_Click()
On Error Resume Next
    frmAbout.Show
End Sub


Private Sub mnuMediaControl_Click(Index As Integer)
    Select Case Index
        Case 0
            Call GoTop
            Exit Sub
        Case 1
            Call Mp3.BackTrack
            Exit Sub
        Case 2
            Call Mp3.Prev
            Exit Sub
        Case 3
            Call Mp3.NextTrack
            Exit Sub
        Case 4
            Call Mp3.Forw
            Exit Sub
        Case 5
            Call GoBottom
            Exit Sub
        Case 6
            Call cmdPause_Click
            Exit Sub
        Case 7
            Call StopPlayer
            Exit Sub
        Case 8
        Case 9
            frmInput.Show
            Exit Sub
    End Select
End Sub


Private Sub mnuMoreSkin_Click(Index As Integer)
    mnuMoreSkin(Index).Checked = True
    For i = 0 To 3
        If i <> Index Then mnuMoreSkin(i).Checked = False
    Next i
    Select Case Index
        Case 0
            strPathSkin = App.Path & "\Skins\Default"
        Case 1
            strPathSkin = App.Path & "\Skins\Black"
        Case 2
            strPathSkin = App.Path & "\Skins\Plastic"
        Case 3
            strPathSkin = App.Path & "\Skins\Rubi"
    End Select
    Call LoadSkin(strPathSkin)
    Call RefreshSkin
End Sub

Private Sub mnuPlaytype_Click(Index As Integer)
    If Index <> 2 Then
        mnuPlaytype(Index).Checked = True
        Select Case Index
            Case 0
                    Call mnuTypeC_Click(0)
            Case 1
                    Call mnuTypeC_Click(1)
        End Select
    Else
        frmRate.Show
    End If
End Sub
Private Sub mnuShowPL_Click()
    If mnuShowPL.Checked = False Then
        mnuShowPL.Checked = True
        bHidePL = False
        pctPlaylist.Picture = LoadPicture(strPathSkin & "\Button\PLactive.jpg")
        Me.Height = 9000
    Else
        mnuShowPL.Checked = False
        bHidePL = True
        pctPlaylist.Picture = LoadPicture(strPathSkin & "\Button\PL.jpg")
        Me.Height = 9000 - Me.imgMedia.Height
    End If
End Sub

Private Sub mnuSpectrumC_Click(Index As Integer)
    Select Case Index
        Case 0
            If mnuSpectrumC(0).Checked = False Then
                mnuSpectrumC(0).Checked = True
                mnuSpectrumC(1).Checked = False
                mnuSpectrumC(2).Checked = False
                mnuSpectrumC(3).Enabled = False
                   For X = 0 To p.Count - 1
                        p(X).Visible = False
                        pM(X).Visible = False
                    Next X
                    For X = 0 To prow.Count - 1
                        prow(X).Visible = False
                        prowM(X).Visible = False
                    Next X
                    If frmVisual.Timer1.Enabled = True Then Unload frmVisual
            Else
                Exit Sub
            End If
            timerVisual.Enabled = True
        Case 1
            mnuSpectrumC(0).Checked = False
            mnuSpectrumC(1).Checked = True
            mnuSpectrumC(2).Checked = False
            mnuSpectrumC(3).Checked = False
            mnuSpectrumC(3).Enabled = True
            For X = 0 To p.Count - 1
                p(X).Visible = True
                pM(X).Visible = True
            Next X
            For X = 0 To prow.Count - 1
                prow(X).Visible = False
                prowM(X).Visible = False
            Next X
            timerVisual.Enabled = True
        Case 2
            mnuSpectrumC(1).Checked = False
            mnuSpectrumC(0).Checked = False
            mnuSpectrumC(2).Checked = True
            mnuSpectrumC(3).Checked = False
            mnuSpectrumC(3).Enabled = True
            For X = 0 To p.Count - 1
                p(X).Visible = False
                pM(X).Visible = False
            Next X
            For X = 0 To prow.Count - 1
                prow(X).Visible = True
                prowM(X).Visible = True
            Next X
            timerVisual.Enabled = True
        Case 3
            If mnuSpectrumC(3).Checked = False Then
                    If mnuSpectrumC(1).Checked = True Then
                        For X = 0 To pM.Count - 1
                            p(X).Visible = False
                            pM(X).Visible = False
                        Next X
                    End If
                    If mnuSpectrumC(2).Checked = True Then
                        For X = 0 To prow.Count - 1
                            prow(X).Visible = False
                            prowM(X).Visible = False
                        Next X
                    End If
                    frmVisual.Show
                    mnuSpectrumC(3).Checked = True
                    mnuSpectrumC(1).Enabled = False
                    mnuSpectrumC(2).Enabled = False
                    timerVisual.Enabled = False
            Else
                Unload frmVisual
            End If
    End Select
    nVisual = Index
End Sub

Private Sub mnuSubChild_Click(Index As Integer)
    Select Case Index
        Case 0
            Call SubFile
            Exit Sub
        Case 1
            Call SubFolder
            Exit Sub
    End Select
End Sub

Private Sub mnuSubMedia_Click(Index As Integer)
    Select Case Index
        Case 2
            Call KillFile(lstList1.List(lstList.ListIndex))
        Case 1
            Call SubFolder
        Case 0
            Call SubFile
    End Select
End Sub

Private Sub mnuTitle_Click()
    frmTitle.Show
End Sub

Private Sub mnuTray_Click(Index As Integer)
Select Case Index
    Case 0
        Call GoTop
        Exit Sub
    Case 1
        Call Mp3.BackTrack
        Exit Sub
    Case 2
        Call Mp3.Prev
        Exit Sub
    Case 3
        Call Mp3.NextTrack
        Exit Sub
    Case 4
        Call Mp3.Forw
        Exit Sub
    Case 5
        Call GoBottom
        Exit Sub
    Case 6
        Call cmdPause_Click
        Exit Sub
    Case 7
        Call StopPlayer
        Exit Sub
    Case 8
        Exit Sub
    Case 9
        frmMedia.WindowState = vbNormal
        Exit Sub
    Case 10
        Unload Me
    End Select
End Sub



Private Sub mnuTypeC_Click(Index As Integer)
    Select Case Index
        Case 0
            If mnuTypeC(0).Checked = False Then
                mnuTypeC(0).Checked = True
                mnuPlaytype(0).Checked = True
                pctShuffe.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\Shuffe.jpg")
            Else
                mnuTypeC(0).Checked = False
                mnuPlaytype(0).Checked = False
                pctShuffe.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\Cont.jpg")
            End If
        Case 1
            If mnuTypeC(1).Checked = False Then
                mnuTypeC(1).Checked = True
                mnuPlaytype(1).Checked = True
                pctRepeat.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\ableRepeat.jpg")
            Else
                pctRepeat.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\Repeat.jpg")
                mnuPlaytype(1).Checked = False
                mnuTypeC(1).Checked = False
            End If
        Case 2
            frmRate.Show
    End Select
End Sub

Private Sub mnuVol_Click(Index As Integer)
    Select Case Index
        Case 0
            Call Mp3.VolInc
        Case 1
            Call Mp3.VolDec
    End Select
End Sub


Private Sub pctInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = vbRightButton Then
        Exit Sub
    Else
            If Button = vbMiddleButton Then
                Exit Sub
            Else
                ReleaseCapture
                SendMessage Me.hWnd, &HA1, 2, 0&
            End If
    End If
End Sub

Private Sub pctInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMedia, vbPopupMenuLeftAlign
    End If
End Sub

Private Sub pctListDown_DblClick()
        If lstList.ListCount > 18 Then
            If lstList.TopIndex < lstList.ListCount - 18 Then
                lstList.TopIndex = lstList.ListCount - 18
            End If
        End If
End Sub

Private Sub pctListDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If lstList.ListCount > 18 Then
            If lstList.TopIndex = lstList.ListCount - 18 Then
                lstList.TopIndex = lstList.ListCount - 18
            Else
                TimerListDown.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub pctListDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If lstList.ListCount > 18 Then
            If lstList.TopIndex = lstList.ListCount - 18 Then
                lstList.TopIndex = lstList.ListCount - 18
            Else
                lstList.TopIndex = lstList.TopIndex + 1
            End If
        End If
    TimerListDown.Enabled = False
    End If
End Sub

Private Sub pctListUp_DblClick()
        If lstList.TopIndex > 0 Then
            lstList.TopIndex = 0
        End If
End Sub

Private Sub pctListUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If lstList.TopIndex = 0 Then
            lstList.TopIndex = 0
        Else
            TimerListUp.Enabled = True
        End If
    End If
End Sub

Private Sub pctListUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If lstList.TopIndex = 0 Then
            lstList.TopIndex = 0
        Else
            lstList.TopIndex = lstList.TopIndex - 1
        End If
        TimerListUp.Enabled = False
    End If
End Sub

Private Sub hscVol_Scroll()
On Error GoTo Err
   If hscVol.Value > -500 And hscVol.Value < 500 Then
      shpMonoL.BorderColor = RGB(255, 255, 255)
      shpMonoR.BorderColor = RGB(255, 255, 255)
   End If
   If hscVol.Value > 500 Then
      shpMonoR.BorderColor = RGB(50, 255, 50)
    End If
    If hscVol.Value < -500 Then
        shpMonoL.BorderColor = RGB(50, 255, 50)
    End If
   MediaPlayer1.Balance = hscVol.Value
   Exit Sub
Err:
    MsgBox "Error !!!"
End Sub
Private Sub hscPosition_Scroll()
    MediaPlayer1.CurrentPosition = hscPosition.Value
End Sub

Private Sub pctPlaylist_Click()
    Call mnuShowPL_Click
End Sub

Private Sub pctRepeat_Click()
    Call mnuTypeC_Click(1)
End Sub

Private Sub pctShuffe_Click()
    If mnuTypeC(0).Checked = False Then
        mnuTypeC(0).Checked = True
        mnuPlaytype(0).Checked = True
        pctShuffe.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\Shuffe.jpg")
    Else
        mnuTypeC(0).Checked = False
        mnuPlaytype(0).Checked = False
        pctShuffe.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\Cont.jpg")
    End If
End Sub

Private Sub Timer1_Timer()
Dim TimeCurrent As Integer
On Error Resume Next
    If allowplay = "yes" Then
        hscPosition.Value = MediaPlayer1.CurrentPosition
        TimeCurrent = MediaPlayer1.CurrentPosition
        Dim min, sec, MinValue, SecValue As Integer
        min = TimeCurrent \ 60
        sec = TimeCurrent - (min * 60)
        If sec = "-1" Then sec = "0"
        MinValue = (MediaPlayer1.Duration - MediaPlayer1.CurrentPosition) \ 60
        SecValue = (MediaPlayer1.Duration - MediaPlayer1.CurrentPosition) - MinValue * 60
        If optTimer.Value = False Then
                lblPosition.Caption = min & ":" & sec \ 1
        Else
            lblPosition.Caption = "-" & MinValue & ":" & SecValue \ 1
        End If
    End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
If allowplay = "yes" Then
    Dim scroll%
        scroll = lblTitle.Left - 50
        lblTitle.Left = scroll
        If lblTitle.Left < -lblTitle.Width Then
            lblTitle.Left = pctInfo.Width
        End If
    If Me.WindowState = vbMinimized Then
        Call mdlSys.SysTip("MediaPHT 1.0 _" & lblTitle.Caption)
        Me.Caption = lblTitle.Caption
    Else
        Me.Caption = ""
    End If
End If
End Sub
Private Sub Timer5_Timer()
On Error Resume Next
    If allowpause = True Then
        If paused = True Then
            Me.cmdPause.Picture = LoadPicture(App.Path & "\Skins\Default\Button\PauseActive.jpg")
        Else
            Me.cmdPause.Picture = LoadPicture(App.Path & "\Skins\Default\Button\Pause.jpg")
        End If
    Else
            Me.cmdPause.Picture = LoadPicture(App.Path & "\Skins\Default\Button\Pause.jpg")
    End If
    If allowplay = "yes" Then
        Me.cmdPlay.Picture = LoadPicture(App.Path & "\Skins\Default\Button\PlayActive.jpg")
    Else
        Me.cmdPlay.Picture = LoadPicture(App.Path & "\Skins\Default\Button\Play.jpg")
    End If
End Sub
Private Sub TimerListDown_Timer()
    If lstList.TopIndex < lstList.ListCount - 18 Then
        lstList.TopIndex = lstList.TopIndex + 1
    End If
End Sub

Private Sub TimerListUp_Timer()
    If lstList.TopIndex > 0 Then
        lstList.TopIndex = lstList.TopIndex - 1
    End If
End Sub

Private Sub timerVisual_Timer()
    If allowplay = "yes" Then
    Randomize MediaPlayer1.AudioStream
        If mnuSpectrumC(1).Checked = True Then Call VisualRain
        If mnuSpectrumC(2).Checked = True Then Call VisualRow
        If mnuSpectrumC(0).Checked = True Then mnuSpectrumC(3).Enabled = False
    End If
End Sub

