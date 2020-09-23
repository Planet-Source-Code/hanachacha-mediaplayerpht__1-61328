VERSION 5.00
Begin VB.Form frmDir 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4575
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   FontTransparent =   0   'False
   ForeColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkInsub 
      BackColor       =   &H00000000&
      Caption         =   "Include sub folder"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   200
      TabIndex        =   5
      Top             =   220
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FF80&
      Caption         =   "&X"
      Height          =   250
      Left            =   3840
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080FF80&
      Caption         =   "&Ok"
      Height          =   250
      Left            =   360
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4680
      UseMaskColor    =   -1  'True
      Width           =   400
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00C0C000&
      ForeColor       =   &H00400000&
      Height          =   3690
      Left            =   200
      TabIndex        =   1
      Top             =   900
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00C0C000&
      ForeColor       =   &H00400000&
      Height          =   315
      Left            =   200
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image imgDir 
      Height          =   5000
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4600
   End
End
Attribute VB_Name = "frmDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i, X As Integer
Private Function RepairPath(ByVal Path As String, File As String) As String
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    RepairPath = Path & File
End Function
Private Sub AddSingleFolder(strFolder As String)
Dim ExtName As String
    File1.Path = strFolder
    For i = 0 To File1.ListCount - 1
        File1.ListIndex = i
        frmMedia.lstList1.AddItem (strFolder & "\" & File1.Filename)
        If UCase(Right(File1.Filename, 3)) = "MP3" Then
            Call Mp3.ReadTag(strFolder & "\" & File1.Filename)
            frmMedia.lstList.AddItem (Trim(Tag1.Artist) & "_" & Trim(Tag1.Title))
        Else
             frmMedia.lstList.AddItem (GetShortName(File1.Filename))
       End If
    Next i
    If bShowNumber = True Then
                For i = 1 To frmMedia.lstList.ListCount
                    If Asc(Mid(frmMedia.lstList.List(i - 1), 1, 1)) >= 65 And Asc(Mid(frmMedia.lstList.List(i - 1), 1, 1)) <= 122 Then
                        frmMedia.lstList.List(i - 1) = i & "-" & frmMedia.lstList.List(i - 1)
                    Else
                        frmMedia.lstList.List(i - 1) = frmMedia.lstList.List(i - 1)
                    End If
                Next i
    End If
End Sub
Private Sub AddSubFolder(strTopFolder As String)
    Dim DirList As New Collection, Temp As String, ExtName() As String
    On Error Resume Next
    DirList.Add RepairPath(strTopFolder, "")
    Do While DirList.Count
        Temp = Dir$(DirList(1), vbDirectory)
        Do Until Temp = ""
            If Temp = "." Or Temp = ".." Then
            ElseIf (GetAttr(RepairPath(DirList(1), Temp)) And vbDirectory) = vbDirectory Then
                DirList.Add RepairPath(DirList(1), Temp) & "\"
            ElseIf InStr(Temp, ".") Then
                ExtName = Split(Temp, ".")
                Select Case UCase(ExtName(UBound(ExtName)))
                    Case "MP3", "WMA", "WAV", "MID"
                        frmMedia.lstList1.AddItem (RepairPath(DirList(1), Temp))
                        If UCase(Right(RepairPath(DirList(1), Temp), 3)) = "MP3" Then
                            Call Mp3.ReadTag(RepairPath(DirList(1), Temp))
                            frmMedia.lstList.AddItem (Trim(Tag1.Artist) & "_" & Trim(Tag1.Title))
                        Else
                            frmMedia.lstList.AddItem (GetShortName(RepairPath(DirList(1), Temp)))
                        End If
                End Select
            End If
            Temp = Dir$
        Loop
        DirList.Remove 1
    Loop
    If bShowNumber = True Then
            For i = 1 To frmMedia.lstList.ListCount
                If Asc(Mid(frmMedia.lstList.List(i - 1), 1, 1)) >= 65 And Asc(Mid(frmMedia.lstList.List(i - 1), 1, 1)) <= 122 Then
                    frmMedia.lstList.List(i - 1) = i & "-" & frmMedia.lstList.List(i - 1)
                Else
                    frmMedia.lstList.List(i - 1) = frmMedia.lstList.List(i - 1)
                End If
            Next i
    End If
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error Resume Next
If chkInsub.Value = Unchecked Then
    Call AddSingleFolder(Dir1.List(Dir1.ListIndex))
Else
    Call AddSubFolder(Dir1.List(Dir1.ListIndex))
End If
Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Dir1_Click()
    Me.File1.Path = Dir1.List(Dir1.ListIndex)
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
'variable to hold main region
    Me.imgDir.Picture = LoadPicture(strPathSkin & "\Frame\PL.jpg")
    frmMedia.Enabled = False
    Me.Dir1.Path = strPathM
    ChDir Me.Dir1.Path
    Me.Drive1.Drive = Dir1.Path
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMedia.Enabled = True
End Sub

Private Sub imgDir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
