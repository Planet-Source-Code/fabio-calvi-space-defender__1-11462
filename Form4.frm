VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   8775
   ClientLeft      =   90
   ClientTop       =   105
   ClientWidth     =   8640
   LinkTopic       =   "Form4"
   ScaleHeight     =   585
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mission 1: Outspace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2640
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   3570
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Dim fIn As Boolean

Private Sub Form_Load()
Form4.ScaleTop = Form1.PicMain.Top
Form4.ScaleLeft = Form1.PicMain.Left
If mission = 1 Then
   Form4.Picture = LoadPicture(App.Path + "\ship1.jpg")
   stars = True
End If
If mission = 2 Then
   Form4.Picture = LoadPicture(App.Path + "\ship2.jpg")
   missions
   stars = False
End If
If mission = 3 Then
   Form4.Picture = LoadPicture(App.Path + "\ship3.jpg")
   missions
   stars = True
End If
fIn = True
counter = 0
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Label1.Refresh
If fIn = True Then
    If counter <= 255 Then
        colVal = RGB(counter, counter, counter)
        Label1.Visible = True
        Label1.ForeColor = colVal
        counter = counter + 6
        If mission = 1 Then Label1.Caption = "Mission 1: Outspace"
        If mission = 2 Then Label1.Caption = "Mission 2: First Planet"
        If mission = 3 Then Label1.Caption = "Mission 3: Your Mission"
    Else
        fIn = False
        counter = 0
    End If
Else
    If counter <= 255 Then
        colVal = RGB(255 - counter, 255 - counter, 255 - counter)
        Label1.ForeColor = colVal
        counter = counter + 6
    Else
        fIn = True
        counter = 0
        Timer1.Enabled = False
        Label1.Visible = False
        Unload Me
''try to un-comment these two lines below: good radar effect
''but the speed of program decrease (but if you have a fast pc)
'        Form1.Pic_radarv.Visible = True
'        roll = True
'
''Sound Call:
       Music
       If Soundcard = True Then
          lngReturnResult = mciSendString("close all", 0&, 0, 0)
          lngReturnResult = mciSendString("open " + App.Path + "\music.mid type sequencer alias backplay", 0&, 0, 0)
          lngReturnResult = mciSendString("play backplay", 0&, 0, 0)
       End If
       Form1.Timer2.Enabled = True
    End If
End If
End Sub

