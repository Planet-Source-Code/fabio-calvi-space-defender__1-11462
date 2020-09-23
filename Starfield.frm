VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Picture         =   "Starfield.frx":0000
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox rad3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   120
      Left            =   8760
      Picture         =   "Starfield.frx":15F942
      ScaleHeight     =   4
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   4
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox Pic_radarv 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      Height          =   2325
      Left            =   9240
      Picture         =   "Starfield.frx":15F9B4
      ScaleHeight     =   155
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   16
      Top             =   450
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.PictureBox rad2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   8760
      Picture         =   "Starfield.frx":170BA2
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox rad 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   8760
      Picture         =   "Starfield.frx":170C5C
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Piclive 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   9000
      Picture         =   "Starfield.frx":170D16
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   11
      Top             =   5880
      Width           =   1935
   End
   Begin VB.PictureBox Picled2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   9240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   73
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Timer ClockTimer 
      Interval        =   1000
      Left            =   2760
      Top             =   360
   End
   Begin VB.PictureBox picLED1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   8880
      Picture         =   "Starfield.frx":17333C
      ScaleHeight     =   14
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   108
      TabIndex        =   9
      Top             =   5400
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox Pic_radar 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      Height          =   2325
      Left            =   9240
      Picture         =   "Starfield.frx":1745DE
      ScaleHeight     =   155
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   8
      Top             =   480
      Width           =   2250
   End
   Begin VB.PictureBox bar2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2010
      Left            =   9840
      Picture         =   "Starfield.frx":1857CC
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   2
      TabIndex        =   7
      Top             =   6360
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox bar1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   90
      Left            =   9480
      Picture         =   "Starfield.frx":185C1E
      ScaleHeight     =   2
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   6
      Top             =   6600
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   9600
      Picture         =   "Starfield.frx":185F90
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   5
      Top             =   6480
      Width           =   1800
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   960
      Top             =   360
   End
   Begin VB.Timer Tmr_flgm 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1560
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   360
   End
   Begin VB.PictureBox PicHealth 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   11160
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
      Top             =   4920
      Width           =   135
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2160
      Top             =   360
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1635
      Left            =   8760
      Picture         =   "Starfield.frx":1909FA
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   19
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.PictureBox PicMain 
      BackColor       =   &H00000000&
      Height          =   8775
      Left            =   120
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   571
      TabIndex        =   1
      Top             =   120
      Width           =   8625
      Begin VB.Timer Tmrsound 
         Interval        =   100
         Left            =   3840
         Top             =   240
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   3240
         Top             =   240
      End
   End
   Begin VB.PictureBox PicScreenBuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   8775
      Left            =   120
      ScaleHeight     =   581
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   571
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   8625
   End
   Begin VB.Label lblhiscore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   15
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblscore 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9720
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   9960
      TabIndex        =   3
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim result As Long, pressed As Boolean
Dim ClockFlashState As Boolean
Private Const SRCCOPY = &HCC0020

Private Sub Form_Load()
Dim lngReturnResult As Long

NUM = 1
BackYPos = 600

lblscore.Caption = score

Open App.Path + "\hiscore" For Input As #1
Input #1, hiscore
Close #1

lblhiscore.Caption = hiscore

missions
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyLeft Then
    sLeft = 1
    Form2.Picture1.Picture = Form2.Pic2l.Picture
    Form2.Picture2.Picture = Form2.Pic2lm.Picture
End If
If KeyCode = vbKeyRight Then
    sRight = 1
    Form2.Picture1.Picture = Form2.Pic2r.Picture
    Form2.Picture2.Picture = Form2.Pic2rm.Picture
End If
If KeyCode = vbKeyUp Then sUp = 1
If KeyCode = vbKeyDown Then sDown = 1
If KeyCode = vbKeyControl Then Timer1.Enabled = True
If KeyCode = &H13 Then
   Timer2.Enabled = False
   frmPause.Show
End If
If KeyCode = vbKeyEscape Then
   GameActive = 0
   Fade PicMain, 100
   ShowCursor 1
   Form1.Visible = False
''Sound Call
   Music
   If Soundcard = True Then
      lngReturnResult = mciSendString("close all", 0&, 0, 0)
      lngReturnResult = mciSendString("open " + App.Path + "\end.mid type sequencer alias backplay", 0&, 0, 0)
      lngReturnResult = mciSendString("play backplay", 0&, 0, 0)
   End If
   Form3.Show
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyLeft Then sLeft = 0
If KeyCode = vbKeyRight Then sRight = 0
If KeyCode = vbKeyUp Then sUp = 0
If KeyCode = vbKeyDown Then sDown = 0
If KeyCode = vbKeyControl Then
   Firing = 0
   Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()

OldShipX = ShipX
OldShipY = ShipY

'Do all of the game engine stuff...

MoveAndDrawBack
If stars = True Then DrawStars
VelocityCode
If stage = 1 Then badguy1
If stage = 2 Then badguy2
If stage = 3 And flgr < 2 Then badguy3
If stage = 4 Then badguy4
If stage = 5 Then badguy5
If stage = 6 Then astero
If stage = 7 Then badguy6
If stage = 8 And flgr < 3 Then badguy7
If stage = 9 Then badguy8
If stage = 10 Then badguy9
If stage = 11 Then badguy10
If stage = 12 Then badguy11
If stage = 13 Then badguy12
If stage = 14 Then badguy13
If stage = 15 Then badguy14
If stage = 16 Then badguy15
If stage = 17 Then badguy16
If stage = 18 Then badguy17
If stage = 19 Then badguy18
If stage = 20 And flgr < 2 Then badguy19
If stage = 21 Then badguy20
If stage = 22 Then badguy22
If stage = 23 Then badguy21

FireBullets
GoodGuyStuff
If roll = True Then RollIT

lblscore.Caption = score
If score > hiscore Then lblhiscore.Caption = score

If score > bonus Then
   If lives < 5 Then
      lives = lives + 1
      If lives = 5 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live5.bmp")
      If lives = 4 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live4.bmp")
      If lives = 3 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live3.bmp")
      If lives = 2 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live2.bmp")
      If lives = 1 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live1.bmp")
   Else
      Health = 100
      DrawHealthBar
   End If
   bonus = bonus + 1500
End If
End Sub
Private Sub Timer1_Timer()
Timer1.Interval = Timer1.Interval + 1
Firing = 1
If Timer1.Interval > 3 Then
   Firing = 0
   Timer1.Interval = 1
End If
End Sub

Private Sub Timer3_Timer()
Timer3.Interval = Timer3.Interval + 1
If Timer3.Interval = 50 Then
   If mission = 1 Then
      stage = 1
      BuildBadGuys
   End If
   If mission = 2 Then
      stage = 21
      BuildBadGuys
   End If
End If
If Timer3.Interval = 100 Then
   If mission = 1 Then
      If stage = 1 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 2
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 21 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 12
         BuildBadGuys
      End If
   End If

End If
If Timer3.Interval = 150 Then
   If mission = 1 Then
      If stage = 2 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 3
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 12 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 13
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 160 Then
   If mission = 1 Then
       If stage = 3 Then Timer3.Interval = Timer3.Interval - 1
       If stage = 0 Then
          stage = 4
          BuildBadGuys
       End If
    End If
   If mission = 2 Then
      If stage = 13 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 14
         BuildBadGuys
      End If
   End If

End If
If Timer3.Interval = 170 Then
   If mission = 1 Then
      If stage = 4 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 5
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 14 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 15
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 180 Then
   If mission = 1 Then
      If stage = 5 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 6
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 15 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 16
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 190 Then
   If mission = 1 Then
      If stage = 6 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 7
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 16 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 17
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 200 Then
   If mission = 1 Then
      If stage = 7 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 8
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 17 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 18
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 210 Then
   If mission = 1 Then
      If stage = 8 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 9
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 18 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 19
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 220 Then
   If mission = 1 Then
      If stage = 9 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 10
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 19 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 20
         BuildBadGuys
      End If
   End If

End If
If Timer3.Interval = 230 Then
   If mission = 1 Then
      If stage = 10 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 11
         BuildBadGuys
      End If
   End If
   If mission = 2 Then
      If stage = 20 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 2
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 240 Then
   If mission = 1 Then
      If stage = 11 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         Timer2.Enabled = False
         Pic_radarv.Visible = False
         Fade PicMain, 100
         For x = 0 To NumOfBullets
             Bulletl(x).Activated = 0
             Bulletr(x).Activated = 0
         Next
         Timer1.Enabled = False
         Firing = 0
         flgs = 1
         flgm = 0
         sVel = 0
         sUp = 0: sDown = 0: sLeft = 0: sRight = 0
         ShipX = Form1.PicMain.ScaleWidth / 2 - 25
         ShipY = Form1.PicMain.ScaleHeight - 44
         BulletsActivated = 0
         lives = 5
         Health = 100
         DrawHealthBar
         Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live5.bmp")
         Timer3.Interval = 1
         stage = 12
         mission = 2
         Form4.Show
      End If
   End If
   If mission = 2 Then
      If stage = 2 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         stage = 22
         BuildBadGuys
      End If
   End If
End If
If Timer3.Interval = 250 Then
   If mission = 2 Then
      If stage = 22 Then Timer3.Interval = Timer3.Interval - 1
      If stage = 0 Then
         Timer2.Enabled = False
         Pic_radarv.Visible = False
         Fade PicMain, 100
         For x = 0 To NumOfBullets
             Bulletl(x).Activated = 0
             Bulletr(x).Activated = 0
         Next
         Timer1.Enabled = False
         Firing = 0
         flgs = 1
         flgm = 0
         sVel = 0
         sUp = 0: sDown = 0: sLeft = 0: sRight = 0
         ShipX = Form1.PicMain.ScaleWidth / 2 - 25
         ShipY = Form1.PicMain.ScaleHeight - 44
         BulletsActivated = 0
         lives = 5
         Health = 100
         DrawHealthBar
         Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live5.bmp")
         Timer3.Interval = 1
         stage = 23
         mission = 3
         Form4.Show
      End If
   End If
End If
End Sub

Private Sub Tmr_flgm_Timer()
Tmr_flgm.Interval = Tmr_flgm.Interval + 1
If Tmr_flgm.Interval > 200 Then
   flgm = 2
   Tmr_flgm.Enabled = False
   Tmr_flgm.Interval = 1
End If
End Sub

Private Sub ClockTimer_Timer()
    Dim tempTime As String
    tempTime = Format(Time, "hh:mm:ss")
    BitBlt Picled2.hdc, 0, 1, 10, 14, picLED1.hdc, 10 * Int(Left$(tempTime, 1)), 0, SRCCOPY 'h
    BitBlt Picled2.hdc, 10, 1, 10, 14, picLED1.hdc, 10 * Int(Mid$(tempTime, 2, 1)), 0, SRCCOPY 'h
    BitBlt Picled2.hdc, 25, 1, 10, 14, picLED1.hdc, 10 * Int(Mid$(tempTime, 4, 1)), 0, SRCCOPY 'm
    BitBlt Picled2.hdc, 35, 1, 10, 14, picLED1.hdc, 10 * Int(Mid$(tempTime, 5, 1)), 0, SRCCOPY 'm
    BitBlt Picled2.hdc, 50, 1, 10, 14, picLED1.hdc, 10 * Int(Mid$(tempTime, 7, 1)), 0, SRCCOPY 's
    BitBlt Picled2.hdc, 60, 1, 10, 14, picLED1.hdc, 10 * Int(Right$(tempTime, 1)), 0, SRCCOPY 's
End Sub

Private Sub UserControl_Terminate()
ClockTimer.Enabled = False
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
        lngReturnResult = mciSendString("open " + App.Path + "\music.mid type sequencer alias backplay", 0&, 0, 0)
        lngReturnResult = mciSendString("play backplay", 0&, 0, 0)
    
    End If
End If
End Sub


