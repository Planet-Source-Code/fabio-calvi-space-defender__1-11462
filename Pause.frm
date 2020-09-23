VERSION 5.00
Begin VB.Form frmPause 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   293
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDC 
      AutoRedraw      =   -1  'True
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   0
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Press Pause to resume"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "GAME PAUSED"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Width           =   3735
   End
End
Attribute VB_Name = "frmPause"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = Form1.PicMain.Left + 1000
    Me.Top = Form1.PicMain.Top + 2500
    
    picDC.Width = Me.ScaleWidth + 10
    picDC.Height = Me.ScaleHeight + 10
    picDC.Left = 0
    picDC.Top = 0
    DeskHdc = GetDC(0)
    ret = BitBlt(picDC.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight, DeskHdc, Me.Left / Screen.TwipsPerPixelX, Me.Top / Screen.TwipsPerPixelY, vbSrcCopy)
    ret = ReleaseDC(0&, DeskHdc)
    Blend Me, picDC, 200, 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = &H13 Then
   sVel = 0: Firing = 0
   sUp = 0: sDown = 0: sLeft = 0: sRight = 0
   Form1.Timer2.Enabled = True
   Unload Me
End If
End Sub

