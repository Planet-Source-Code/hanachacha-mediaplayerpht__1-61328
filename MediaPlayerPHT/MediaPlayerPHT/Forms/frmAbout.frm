VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3090
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "VNI-Times"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   360
      Top             =   2520
   End
   Begin VB.Label lblWriter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please send email phtmouse84@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   1680
      Width           =   4380
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgAboutM 
      Height          =   1290
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4680
   End
   Begin VB.Label lblWriter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is a basic media player ^o^"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   1
      Top             =   1320
      Width           =   4020
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   285
      Left            =   3960
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2760
      Width           =   660
   End
   Begin VB.Image imgAbout 
      Height          =   3090
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    imgAbout.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\nenAbout.jpg")
    imgAboutM.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\nenAboutTop.jpg")
    frmMedia.Enabled = False
    If frmMedia.timerVisual.Enabled = True Then frmMedia.timerVisual.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    frmMedia.Enabled = True
    frmMedia.timerVisual.Enabled = True
End Sub

Private Sub imgAbout_DblClick()
On Error Resume Next
    Unload Me
End Sub

Private Sub imgAboutM_DblClick()
On Error Resume Next
    Unload Me
End Sub

Private Sub lblOk_Click()
On Error Resume Next
    Unload Me
End Sub

Private Sub Timer1_Timer()
    lblWriter(0).Top = lblWriter(0).Top - 20
    If lblWriter(0).Top < Me.imgAboutM.Height Then lblWriter(0).Top = Me.Height
    lblWriter(1).Top = lblWriter(0).Top + 200
End Sub
