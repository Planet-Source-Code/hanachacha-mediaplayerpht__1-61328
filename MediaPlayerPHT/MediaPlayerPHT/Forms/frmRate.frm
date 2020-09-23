VERSION 5.00
Begin VB.Form frmRate 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   840
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H0080FF80&
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2700
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2415
      LargeChange     =   10
      Left            =   120
      Max             =   10
      Min             =   200
      SmallChange     =   10
      TabIndex        =   3
      Top             =   120
      Value           =   100
      Width           =   135
   End
   Begin VB.Label lblPercent2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 %"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   2340
      Width           =   345
   End
   Begin VB.Label lblPercent1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200 %"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblPercent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1260
      Width           =   435
   End
   Begin VB.Image imgRate 
      Height          =   3255
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   855
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReset_Click()
    VScroll1.Value = 100
End Sub

Private Sub Form_Load()
Me.imgRate.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
If frmMedia.WindowState = vbNormal Then
    Me.Left = frmMedia.Left - Me.Width
    Me.Top = frmMedia.Top
Else
    Me.Left = Screen.Width - Me.Width - 500
    Me.Top = Screen.Height - Me.Height - 500
End If
cmdReset.MousePointer = 10
VScroll1.MousePointer = 10
End Sub

Private Sub imgRate_DblClick()
    Unload Me
End Sub

Private Sub imgRate_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub VScroll1_Change()
frmMedia.MediaPlayer1.Rate = VScroll1.Value / 100
frmMedia.timerVisual.Interval = 100 / frmMedia.MediaPlayer1.Rate
If frmVisual.bShow = True Then frmVisual.Timer1.Interval = 100 / frmMedia.MediaPlayer1.Rate
End Sub

Private Sub VScroll1_Scroll()
frmMedia.MediaPlayer1.Rate = VScroll1.Value / 100
frmMedia.timerVisual.Interval = 100 / frmMedia.MediaPlayer1.Rate
If frmVisual.bShow = True Then frmVisual.Timer1.Interval = 100 / frmMedia.MediaPlayer1.Rate
End Sub
