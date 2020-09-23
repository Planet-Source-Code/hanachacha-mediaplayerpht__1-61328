VERSION 5.00
Begin VB.Form frmAlphaCustom 
   BorderStyle     =   0  'None
   Caption         =   "Dieu chinh do sang"
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hscAlpha 
      Height          =   135
      Left            =   473
      Max             =   100
      Min             =   10
      TabIndex        =   0
      Top             =   320
      Value           =   100
      Width           =   3735
   End
   Begin VB.Label lblApha2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   3675
      TabIndex        =   2
      Top             =   100
      Width           =   540
   End
   Begin VB.Label lblAlpha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 %"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   480
      TabIndex        =   1
      Top             =   100
      Width           =   435
   End
   Begin VB.Image imgAlpha 
      Height          =   600
      Left            =   0
      Stretch         =   -1  'True
      ToolTipText     =   "Double Click to exit"
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmAlphaCustom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    imgAlpha.Picture = LoadPicture(strPathSkin & "\Frame\Control.jpg")
    hscAlpha.Value = nAlpha
End Sub

Private Sub hscAlpha_Change()
    nAlpha = hscAlpha.Value
    Call mdlBoot.Make_Transparent(frmMedia.hWnd, nAlpha)
End Sub

Private Sub hscAlpha_Scroll()
    nAlpha = hscAlpha.Value
    Call mdlBoot.Make_Transparent(frmMedia.hWnd, nAlpha)
End Sub

Private Sub imgAlpha_DblClick()
    Unload Me
End Sub

Private Sub imgAlpha_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
