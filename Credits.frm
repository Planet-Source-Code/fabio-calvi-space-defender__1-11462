VERSION 5.00
Begin VB.Form FrmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Defender - Credits"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "Credits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Tmrsound 
      Interval        =   100
      Left            =   360
      Top             =   2880
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   360
      Top             =   2160
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   1440
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "------- CREDITS -------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   3000
      Width           =   3495
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Dim fIn As Boolean
Dim flag As Integer
Public Sub Form_Load()
Form3.Visible = False
counter = 0
flag = 0
fIn = True
End Sub
Private Sub credits()
rec = 0
CreditText(0) = " "
CreditText(1) = " "
CreditText(2) = "Author: Tanner Helland"
CreditText(3) = "Code: Return of the Avenger"
CreditText(4) = " "
CreditText(5) = " "
Colo.R = 255
Colo.G = 255
Colo.b = 255
' Set background to black
FrmCredits.BackColor = RGB(0, 0, 0)

lblCredits(0).Top = FrmCredits.ScaleHeight 'Set first label just out of sight
lblCredits(0).Left = (FrmCredits.ScaleWidth / 2) - (lblCredits(0).Width / 2) 'Centre first label

For i = 1 To NumLines
    Load lblCredits(i) 'Create new labels
    lblCredits(i).Top = lblCredits(i - 1).Top + lblCredits(i - 1).Height        'Set labels vertical position
    lblCredits(i).Left = (FrmCredits.ScaleWidth / 2) - (lblCredits(i).Width / 2)   'Centre label
    lblCredits(i).Visible = True    'Show label
Next i

For i = 0 To NumLines
lblCredits(i).Caption = CreditText(i)   'Set labels Text
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
cleanup
ChangeScreenSettings iWidth, iHeight, 32
If Soundcard = True Then result = mciSendString("close all", 0&, 0, 0)
Unload Me
End
End Sub

Public Sub Timer1_Timer()
For i = 0 To NumLines
    'Check if label is out of the top
    If lblCredits(i).Top <= -lblCredits(i).Height Then
       lblCredits(i).Top = FrmCredits.ScaleHeight
       If i = 2 Then rec = rec + 1
       If rec > 11 Then rec = 0
       If rec = 0 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Tanner Helland"
          lblCredits(3).Caption = "Code: Return of the Avenger"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 1 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Authors: Soren Christensen , Burt Abreu"
          lblCredits(3).Caption = "Code: Windwalker"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 2 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Jonathan Roach"
          lblCredits(3).Caption = "Code: Fading Text"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 3 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Derek Hall"
          lblCredits(3).Caption = "Code: Digital Clock"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 4 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Richard Lowe"
          lblCredits(3).Caption = "Code: Pixel Collision Detection"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
          End If
       If rec = 5 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Duane Odom"
          lblCredits(3).Caption = "Code: Fade"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 6 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Rocky Clark"
          lblCredits(3).Caption = "Code: Basic Geometry"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 7 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Kevin Tupper"
          lblCredits(3).Caption = "Code: Semi - Transparent Form"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 8 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Carl Warwick"
          lblCredits(3).Caption = "Code: Scrolling Credits"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 9 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Author: Mountain King Studios"
          lblCredits(3).Caption = "Graphics: Demonstar v3.21 Shareware"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 10 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "Special Thanks to"
          lblCredits(3).Caption = "www.Planet-Source-Code.com"
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
       If rec = 11 Then
          lblCredits(0).Caption = " "
          lblCredits(1).Caption = " "
          lblCredits(2).Caption = "------- CREDITS -------"
          lblCredits(3).Caption = " "
          lblCredits(4).Caption = " "
          lblCredits(5).Caption = " "
       End If
    
    End If
    'Sets the color for the text
    If lblCredits(i).Top <= FrmCredits.ScaleHeight And lblCredits(i).Top >= FrmCredits.ScaleHeight / 2 Then
        lblCredits(i).ForeColor = RGB(((FrmCredits.ScaleHeight - lblCredits(i).Top) / (FrmCredits.ScaleHeight / 2)) * Colo.R, ((FrmCredits.ScaleHeight - lblCredits(i).Top) / (FrmCredits.ScaleHeight / 2)) * Colo.G, ((FrmCredits.ScaleHeight - lblCredits(i).Top) / (FrmCredits.ScaleHeight / 2)) * Colo.b)
    End If
    If lblCredits(i).Top <= FrmCredits.ScaleHeight / 2 And lblCredits(i).Top >= 0 Then
        lblCredits(i).ForeColor = RGB((lblCredits(i).Top) / (FrmCredits.ScaleHeight / 2) * Colo.R, (lblCredits(i).Top) / (FrmCredits.ScaleHeight / 2) * Colo.G, (lblCredits(i).Top) / (FrmCredits.ScaleHeight / 2) * Colo.b)
    End If
    
    'Moves the label up
    lblCredits(i).Top = lblCredits(i).Top - 2
Next i

End Sub

Private Sub Timer2_Timer()
Label2.Refresh
If fIn = True Then
    If counter <= 255 Then
        colVal = RGB(counter, counter, counter)
        Label2.ForeColor = colVal
        counter = counter + 6
        If flag = 0 Then Label2.Caption = "------- CREDITS -------"
    Else
        fIn = False
        counter = 0
        flag = flag + 1
    End If
Else
    If counter <= 255 Then
        colVal = RGB(255 - counter, 255 - counter, 255 - counter)
        Label2.ForeColor = colVal
        counter = counter + 6
    Else
        fIn = True
        counter = 0
        If flag = 1 Then
           Timer2.Enabled = False
           Timer1.Enabled = True
           lblCredits(0).Visible = True
           Label2.Move 0, 0
           credits
        End If
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
        lngReturnResult = mciSendString("open " + App.Path + "\end.mid type sequencer alias backplay", 0&, 0, 0)
        lngReturnResult = mciSendString("play backplay", 0&, 0, 0)

    End If
End If
End Sub
