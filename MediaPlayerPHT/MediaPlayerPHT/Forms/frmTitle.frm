VERSION 5.00
Begin VB.Form frmTitle 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   ClientHeight    =   510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   7800
      Top             =   360
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   90
   End
   Begin VB.Image imgTitle 
      Height          =   495
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10215
   End
End
Attribute VB_Name = "frmTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DblClick()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Width = Screen.Width
    Me.Top = 0
    Me.Left = 0
    Me.imgTitle.Width = Me.Width
    Me.imgTitle.Height = Me.Height
    Me.imgTitle.Picture = LoadPicture(strPathSkin & "\Frame\FormTitle.jpg")
    Call mdlBoot.Make_Transparent(Me.hWnd, 70)
End Sub

Private Sub imgTitle_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
    If frmMedia.allowplay = "yes" Then
        Dim scroll%
            scroll = Me.lblTitle.Left - 100
            Me.lblTitle.Caption = frmMedia.lblArtist.Caption & "_" & frmMedia.lblTitle.Caption
            Me.lblTitle.Left = scroll
            If Me.lblTitle.Left < 0 - lblTitle.Width Then
                Me.lblTitle.Left = Me.Width
            End If
    End If
End Sub
