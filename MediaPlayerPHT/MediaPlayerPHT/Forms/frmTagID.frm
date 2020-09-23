VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTagID 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "TagIDv3_1"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFileName 
      Height          =   375
      Left            =   0
      TabIndex        =   17
      Top             =   3480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H0080FF80&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2663
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080FF80&
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1583
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox pctShape 
      BackColor       =   &H00C0C000&
      Height          =   2895
      Left            =   250
      ScaleHeight     =   2835
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4700
      Begin VB.ComboBox cboGenre 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1920
         Width           =   2175
      End
      Begin MSForms.TextBox txtComment 
         Height          =   375
         Left            =   960
         TabIndex        =   16
         Top             =   2400
         Width           =   3615
         VariousPropertyBits=   746604571
         Size            =   "6376;661"
         SpecialEffect   =   6
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtYear 
         Height          =   375
         Left            =   960
         TabIndex        =   15
         Top             =   1920
         Width           =   855
         VariousPropertyBits=   746604571
         Size            =   "1508;661"
         SpecialEffect   =   6
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtAlbum 
         Height          =   375
         Left            =   960
         TabIndex        =   14
         Top             =   1440
         Width           =   3615
         VariousPropertyBits=   746604571
         Size            =   "6376;661"
         SpecialEffect   =   6
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtTitle 
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   960
         Width           =   3615
         VariousPropertyBits=   746604571
         Size            =   "6376;661"
         SpecialEffect   =   6
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtArtist 
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   3615
         VariousPropertyBits=   746604571
         Size            =   "6376;661"
         SpecialEffect   =   6
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "VNI-Bodon-Poster"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   405
         Left            =   1200
         TabIndex        =   11
         Top             =   0
         Width           =   1950
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comment :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   20
         TabIndex        =   10
         Top             =   2400
         Width           =   960
      End
      Begin VB.Label lblGenre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Genre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   1875
         TabIndex        =   9
         Top             =   1920
         Width           =   585
      End
      Begin VB.Label lblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   20
         TabIndex        =   8
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label lblAlbum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   20
         TabIndex        =   7
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   20
         TabIndex        =   6
         Top             =   960
         Width           =   510
      End
      Begin VB.Label lblArtist 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Artist :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   20
         TabIndex        =   5
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         BeginProperty Font 
            Name            =   "VNI-Bodon-Poster"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   405
         Left            =   1230
         TabIndex        =   12
         Top             =   0
         Width           =   1950
      End
   End
   Begin VB.Image imgTagID 
      Height          =   3900
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5220
   End
End
Attribute VB_Name = "frmTagID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bShow As Boolean
Dim i As Integer
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
    If txtFileName.Text <> frmMedia.MediaPlayer1.Filename Then
        Tag1.Artist = ""
        Tag1.Title = ""
        Tag1.Album = ""
        Tag1.Year = ""
        Tag1.Comment = ""
        Tag1.Genre = ""
'
        Tag1.Artist = txtArtist.Text
        Tag1.Title = txtTitle.Text
        Tag1.Album = txtAlbum.Text
        Tag1.Year = txtYear.Text
        Tag1.Comment = txtComment.Text
        Tag1.Genre = cboGenre.ListIndex
        Call Mp3.WriteTag(txtFileName.Text, Tag1)
        frmMedia.lstList.List(frmMedia.lstList.ListIndex) = Mid(frmMedia.lstList.List(frmMedia.lstList.ListIndex), 1, InStrRev(frmMedia.lstList.List(frmMedia.lstList.ListIndex), "-", -1, vbBinaryCompare)) & Trim(Tag1.Artist) & "_" & Trim(Tag1.Title)
    Else
        MsgBox "Khong the cap nhat bai hat dang choi", vbInformation, "Thong bao"
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    bShow = True
    Me.imgTagID.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
    Me.Left = frmMedia.Left + frmMedia.Width
    Me.Top = frmMedia.Top
    txtFileName.Text = frmMedia.lstList1.List(frmMedia.lstList.ListIndex)
    If UCase(Right(txtFileName.Text, 3)) = "MP3" Then
            Call Mp3.ReadTag(txtFileName)
            txtArtist.Text = Trim(Tag1.Artist)
            txtTitle.Text = Trim(Tag1.Title)
            txtAlbum.Text = Trim(Tag1.Album)
            txtYear.Text = Trim(Tag1.Year)
            txtComment.Text = Trim(Tag1.Comment)
            GenreArray = Split(strGenreMatrix, "|")
            For i = LBound(GenreArray) To UBound(GenreArray)
                cboGenre.AddItem GenreArray(i)
            Next i
            cboGenre.ListIndex = Tag1.Genre
    Else
        MsgBox "This is not mp3 file !!!Sorry , can't read information", vbOKOnly + vbInformation, "Error"
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bShow = False
End Sub

Private Sub imgTagID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
