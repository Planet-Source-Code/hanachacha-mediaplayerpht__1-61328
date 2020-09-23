VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00800080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Option"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5055
   ForeColor       =   &H0000C0C0&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMyMusic 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MyMusic"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   170
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Chon thu muc Music"
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton cmdMyDoc 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MyPlaylist"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   170
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Chon thu muc MyDoccuments"
      Top             =   1640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   170
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Bo qua"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   170
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Dong y"
      Top             =   2320
      Width           =   1095
   End
   Begin VB.ListBox lstText 
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   1800
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lblPath 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   0
      TabIndex        =   9
      Top             =   90
      Width           =   45
   End
   Begin VB.Shape shpCaption 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   5100
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00400040&
      BackStyle       =   1  'Opaque
      Height          =   2775
      Left            =   45
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblDoc 
      Height          =   135
      Left            =   2880
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblMusic 
      Height          =   15
      Left            =   2040
      TabIndex        =   3
      Top             =   3600
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strPathMusic, strOldPathMusic As String
Dim strPathDoc, strOldPathDoc  As String

Private Sub cmdCancel_Click()
    Me.lblDoc.Caption = strOldPathDoc
    Me.lblMusic.Caption = strOldPathMusic
    Unload Me
End Sub

Private Sub cmdMyDoc_Click()
    strOldPathDoc = strPathDoc
    strPathDoc = Me.Dir1.Path
    Me.lblDoc.Caption = strPathDoc
    Me.lblMusic.Caption = strPathMusic
End Sub

Private Sub cmdMyMusic_Click()
    strOldPathMusic = strPathMusic
    strPathMusic = Me.Dir1.Path
    Me.lblMusic.Caption = strPathMusic
    Me.lblDoc.Caption = strPathDoc
End Sub

Private Sub cmdOk_Click()
Dim strPath As String
    On Error Resume Next
    strPath = App.Path & "\Config.txt"
    Open strPath For Output As #2
        Print #2, "EXTConfig"
        Print #2, lblMusic.Caption
        Print #2, lblDoc.Caption
    Close #2
    Unload Me
End Sub

Private Sub Dir1_Change()
    lblPath.Caption = Me.Dir1.Path
End Sub

Private Sub Drive1_Change()
    Me.Dir1.Path = Me.Drive1.Drive
    lblPath.Caption = Me.Drive1.Drive
End Sub

Private Sub Form_Load()
Dim strPath, listPath As String
    Me.Left = frmMedia.Left + frmMedia.Width
    Me.Top = frmMedia.Top
    On Error Resume Next
    strPath = App.Path & "\Config.txt"
    Open strPath For Input As #2
    Do Until EOF(2)
        Line Input #2, listPath
        lstText.AddItem listPath
    Loop
    Close #2
    strPathMusic = lstText.List(1)
    strPathDoc = lstText.List(2)
    Me.Dir1.Path = strPathMusic
    Me.Drive1.Drive = Me.Dir1.Path
    Me.lblDoc.Caption = strPathDoc
    Me.lblMusic.Caption = strPathMusic
    strOldPathMusic = strPathMusic
    strOldPathDoc = strPathDoc
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.lblDoc.Caption = strPathDoc
    Me.lblMusic.Caption = strPathMusic
End Sub
