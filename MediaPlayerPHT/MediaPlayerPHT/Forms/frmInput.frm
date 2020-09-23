VERSION 5.00
Begin VB.Form frmInput 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4005
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   4005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FF80&
      Caption         =   "&X"
      Height          =   255
      Left            =   3240
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   220
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0080FF80&
      Caption         =   "&Ok"
      Height          =   255
      Left            =   2640
      MaskColor       =   &H00FF80FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   220
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.TextBox txtVal 
      Height          =   375
      Left            =   1058
      TabIndex        =   0
      Top             =   160
      Width           =   1455
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jump to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   300
      TabIndex        =   3
      Top             =   227
      Width           =   675
   End
   Begin VB.Image imgInput 
      Height          =   720
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4000
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public j As Integer
Private Sub cmdCancel_Click()
    j = 0
    Unload Me
End Sub

Private Sub cmdOk_Click()
    j = CInt(Val(txtVal.Text))
    If j > 0 And j <= frmMedia.lstList.ListCount Then
        frmMedia.lstList.ListIndex = j - 1
        frmMedia.lstList1.ListIndex = frmMedia.lstList.ListIndex
    End If
    frmMedia.MediaPlayer1.Filename = frmMedia.lstList1.Text
    Call Mp3.ReadTag(frmMedia.lstList1.Text)
    frmMedia.lblArtist.Caption = Trim(Tag1.Artist)
    frmMedia.lblTitle.Caption = Trim(Tag1.Title)
    frmMedia.lblDuration.Caption = frmMedia.MediaPlayer1.Duration / 60
    frmMedia.hscPosition.Max = frmMedia.MediaPlayer1.Duration
    frmMedia.MediaPlayer1.Play
    frmMedia.allowpause = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdOk_Click
End Sub

Private Sub Form_Load()
    Me.imgInput.Picture = LoadPicture(strPathSkin & "\Frame\Control.jpg")
    Me.Top = frmMedia.Top + frmMedia.imgTitle.Height + frmMedia.imgVisual.Height
    Me.Left = frmMedia.Left + frmMedia.Width
End Sub

Private Sub imgInput_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
