VERSION 5.00
Begin VB.Form frmOption 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Option"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ControlBox      =   0   'False
   FontTransparent =   0   'False
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmOption 
      BackColor       =   &H00000000&
      Caption         =   "Config"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   1575
      Left            =   200
      TabIndex        =   8
      Top             =   3240
      Width           =   4300
      Begin VB.CheckBox chkLClick 
         BackColor       =   &H00000000&
         Caption         =   "Left click for next track"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Dat bai ke tiep"
         Top             =   960
         Width           =   3375
      End
      Begin VB.CheckBox chkShowNumber 
         BackColor       =   &H00000000&
         Caption         =   "Show number in playlist"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox chkShowPL 
         BackColor       =   &H00000000&
         Caption         =   "Auto load last playlist"
         BeginProperty Font 
            Name            =   "VNI-Times"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame frmPath 
      BackColor       =   &H00000000&
      Caption         =   "Home path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   3015
      Left            =   200
      TabIndex        =   3
      Top             =   120
      Width           =   4300
      Begin VB.CommandButton cmdSetPathP 
         BackColor       =   &H0080FF80&
         Caption         =   "My Playlist"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Set default path for playlist"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdSetPathM 
         BackColor       =   &H0080FF80&
         Caption         =   "My Music"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Set default path for my music"
         Top             =   840
         Width           =   1095
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00C0C000&
         ForeColor       =   &H00400000&
         Height          =   2115
         Left            =   1320
         TabIndex        =   5
         Top             =   760
         Width           =   2900
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00C0C000&
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4100
      End
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H0080FF80&
      Caption         =   "&Apply"
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
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5080
      Width           =   615
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FF80&
      Caption         =   "&X"
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
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5080
      Width           =   615
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H0080FF80&
      Caption         =   "&Ok"
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
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5080
      Width           =   615
   End
   Begin VB.Image imgOption 
      Height          =   5415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkLClick_Click()
    cmdApply.Enabled = True
    If chkLClick.Value = 1 Then
        bLClickChoice = True
    Else
        bLClickChoice = False
    End If
End Sub

Private Sub chkShowNumber_Click()
    cmdApply.Enabled = True
    If chkShowNumber.Value = 1 Then
        bShowNumber = True
    Else
        bShowNumber = False
    End If
End Sub

Private Sub chkShowPL_Click()
    cmdApply.Enabled = True
    If chkShowPL.Value = 1 Then
        bShowList = True
    Else
        bShowList = False
    End If
End Sub

Private Sub cmdApply_Click()
    strOldPathM = strPathM
    strOldPathP = strPathP
    bOldShowPl = bShowList
    bOldShowNumber = bShowNumber
    bOldLClickChoice = bLClickChoice
    Call mdlBoot.SaveConfig
    cmdApply.Enabled = False
End Sub

Private Sub cmdCancel_Click()
    strPathM = strOldPathM
    strPathP = strOldPathP
    bShowList = bOldShowPl
    bShowNumber = bOldShowNumber
    bLClickChoice = bOldLClickChoice
    Unload Me
End Sub

Private Sub cmdOk_Click()
        strOldPathM = strPathM
        strOldPathP = strPathP
        bOldShowPl = bShowList
        bOldShowNumber = bShowNumber
        bOldLClickChoice = bLClickChoice
        Call mdlBoot.SaveConfig
        MsgBox "Chuong trinh se cap nhat thong tin o lan sau", vbOKCancel, "Thong bao"
        Unload Me
End Sub
Private Sub cmdSetPathM_Click()
    strPathM = Dir1.List(Dir1.ListIndex)
    cmdApply.Enabled = True
End Sub
Private Sub cmdSetPathP_Click()
    strPathP = Dir1.List(Dir1.ListIndex)
    cmdApply.Enabled = True
End Sub
Private Sub Dir1_Change()
    cmdApply.Enabled = True
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Me.imgOption.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
    frmMedia.Enabled = False
    strOldPathM = strPathM
    strOldPathP = strPathP
    bOldShowPl = bShowList
    bOldShowNumber = bShowNumber
    bOldLClickChoice = bLClickChoice
    If bShowList = True Then
        chkShowPL.Value = 1
    Else
        chkShowPL.Value = 0
    End If
    If bShowNumber = True Then
        chkShowNumber.Value = 1
    Else
        chkShowNumber.Value = 0
    End If
    If bLClickChoice = True Then
        chkLClick.Value = 1
    Else
        chkLClick.Value = 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMedia.Enabled = True
    frmMedia.Refresh
End Sub

Private Sub imgOption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
