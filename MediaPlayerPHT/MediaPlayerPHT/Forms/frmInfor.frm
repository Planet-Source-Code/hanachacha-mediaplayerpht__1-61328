VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2205
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5175
   ControlBox      =   0   'False
   Icon            =   "frmInfor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   5400
      Top             =   3240
   End
   Begin VB.Label lblCon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading Configuration"
      BeginProperty Font 
         Name            =   "TobysHand"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   2490
   End
   Begin VB.Image imgNen 
      Height          =   2205
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Byte
Private Sub Form_Load()
    i = 0
    imgNen.Picture = LoadPicture(App.Path & "\Skins\Default\Frame\Infor.jpg")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMedia.Show
End Sub

Private Sub Timer1_Timer()
Randomize
    i = i + 1
    lblCon.Caption = lblCon.Caption & "."
    If i > 5 Then Unload Me
End Sub
