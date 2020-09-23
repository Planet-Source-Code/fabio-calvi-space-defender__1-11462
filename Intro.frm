VERSION 5.00
Begin VB.Form Frmintro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Defender - Intro"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   ForeColor       =   &H00000000&
   Icon            =   "Intro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      FillColor       =   &H00C0C0C0&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7935
      Begin VB.Timer Tmrsound 
         Interval        =   100
         Left            =   360
         Top             =   1680
      End
      Begin VB.Timer Timer2 
         Interval        =   300
         Left            =   360
         Top             =   960
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   360
         Top             =   240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   2160
         TabIndex        =   1
         Top             =   2640
         Width           =   3645
      End
   End
End
Attribute VB_Name = "Frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Dim fIn As Boolean
Dim flag As Integer


Private Sub Form_Load()
Dim L_nRun As Integer
Dim lngReturnResult As Long
On Error Resume Next
RememberScreenRes
ChangeScreenSettings 800, 600, 32

''Sound Call:
Music
If Soundcard = True Then
   lngReturnResult = mciSendString("close all", 0&, 0, 0)
   lngReturnResult = mciSendString("open " + App.Path + "\intro.mid type sequencer alias backplay", 0&, 0, 0)
   lngReturnResult = mciSendString("play backplay", 0&, 0, 0)
End If
mission = 1
counter = 0
flag = 0
roll = False
fIn = True
        For L_nRun = 0 To 599
            With G_dStar(L_nRun)
                
                ' Set position
                .nX = Int(Rnd * Frmintro.Picture1.ScaleWidth) + 10
                .nY = Int(Rnd * Frmintro.Picture1.ScaleHeight) + 10
                                
                ' Set speed and color (the further "back", the slower and darker)
                Select Case Int(Rnd * 9) + 1
                
                    Case 1
                        .nSpeed = 30
                        .nColor = &HFFFFFF
                    Case 2, 3, 4
                        .nSpeed = 15
                        .nColor = &H808080
                    Case 5, 6, 7, 8, 9
                        .nSpeed = 10
                        .nColor = &H404040 '&HC0C0C0
                End Select
            End With
        Next

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Then
   Timer1.Enabled = False
   Timer2.Enabled = False
   CurLevel.Damage = 1
   CurLevel.NumOfBadGuys = 0
   CurLevel.BulletSpeed = 10
   CurLevel.OddsOfFiring = 15
   CurLevel.Velocity = 5
   CurLevel.Damagelimit = 50
   ScreenWidth = Screen.Width / Screen.TwipsPerPixelX
   ScreenHeight = Screen.Height / Screen.TwipsPerPixelY
   BufferWidth = Form1.PicScreenBuffer.ScaleWidth
   BufferHeight = Form1.PicScreenBuffer.ScaleHeight
   MainHDC = Form1.PicMain.hdc
   BufferHDC = Form1.PicScreenBuffer.hdc
   InitializeGameEngine
   ShowCursor 0
   Form1.Show
   Form4.Show
   Unload Me
End If
End Sub

Private Sub Timer1_Timer()
For L_nRun = 0 To 599
    With G_dStar(L_nRun)
       Frmintro.Picture1.PSet (.nX, .nY), &H0&
       .nX = .nX - .nSpeed
       If .nX < 5 Then .nX = Frmintro.Picture1.ScaleWidth
       Frmintro.Picture1.PSet (.nX, .nY), .nColor
    End With
Next
End Sub

Private Sub Timer2_Timer()
Label1.Refresh
If fIn = True Then
    If counter <= 170 Then
        colVal = RGB(counter, 0, 0)
        Label1.ForeColor = colVal
        counter = counter + 6
        If flag = 0 Then Label1.Caption = "SPACE  DEFENDER"
    Else
        fIn = False
        counter = 0
        If flag < 3 Then flag = flag + 1
    End If
Else
    If counter <= 170 Then
        colVal = RGB(170 - counter, 0, 0)
        Label1.ForeColor = colVal
        counter = counter + 6
    Else
        fIn = True
        counter = 0
        If flag = 1 Then Label1.Caption = "Coding by Fabio Calvi"
        If flag = 2 Then Label1.Caption = "Press Spacebar to Play"
   
    End If
End If
End Sub

Private Sub Tmrsound_Timer()
Dim strReturnString As String * 255 'dimensions a 255 character length string to hold the return string sent by MCI
Dim lngReturnResult As Long ' dimensions a long to hold the return result sent by the API call
If Soundcard = True Then
    lngReturnResult = mciSendString("status backplay mode", ByVal strReturnString, 255, 0)
    'this sends the command "status" which checks the status of the MCI device, and it looks for the place where the
    'alias "background" is, and checks the "mode" of the MCI device, to see whether it is stopped, or if it is playing.
    'It returns a string that describes the mode and stores it in the fixed-length string strReturnString. We also tell
    'tell the MCI device to only return the first 255 characters of the mode string. As usual, the callback parameter is
    'zero.
    
    If Left$(strReturnString, 7) = "playing" Then 'if the first 7 characters of the string are "playing", then we
                                                  'know that the MCI device is busy playing a MIDI file, and we
        Exit Sub                                  'exit the subroutine.
    
    ElseIf Left$(strReturnString, 6) = "paused" Then 'if the first 6 characters of the string are "paused", then we
                                                     'know the MCI device is paused, and we
        Exit Sub                                     'exit the subroutine
    
    Else                                             'otherwise, the MIDI has stopped playing, and we restart it again
        
        'These API calls are the exact same as clicking the play button.
        lngReturnResult = mciSendString("close all", 0&, 0, 0)
        lngReturnResult = mciSendString("open " + App.Path + "\intro.mid type sequencer alias backplay", 0&, 0, 0)
        lngReturnResult = mciSendString("play backplay", 0&, 0, 0)
    
    End If
End If
End Sub

