Attribute VB_Name = "Game_Logic"
'ALL VARIABLES:
Public iWidth As Integer
Public iHeight As Integer
Public Bad1 As BadGuy
Public ShipX, ShipY As Integer
Public OldShipX, OldShipY
Public KeyState As Byte
Public sLeft As Byte
Public sRight As Byte
Public sUp As Byte
Public sDown As Byte
Public sVelocity As Single
Public sVel As Single
Public Firing As Byte
Public BulletsActivated As Byte
Public CurLevel As Level
Public Exploding As Byte
Public ExplodingFrame As Byte
Public Explosions(0 To 9) As PointXY
Public BadGuys() As BadGuy
Public BadGuyNum As Byte
Public Bulletl(0 To NumOfBullets) As bullet
Public Bulletr(0 To NumOfBullets) As bullet
Public StarArray(0 To NumOfStars) As Star
Public flgs As Integer
Public fboom As Integer
Public flgm As Integer
Public flgr As Integer
Public lives As Integer
Public stage As Integer
Public mission As Integer
Public TempCalc As Byte
Public Tempclac As Byte
Public score As Integer
Public hiscore As Integer
Public numsec As Integer
Public bullets As Boolean
Public oldX As Single
Public oldY As Single
Public frame As Integer
Public stars As Boolean
Public roll As Boolean
Public seq As Byte
Public appo As Boolean
Public bonus As Integer
Public Angle  As Double
Public curX As Integer, curY As Integer
Public deltax As Double, deltay As Double
Public diffwidth As Integer, diffheight As Integer
Public difwidth As Integer, difheight As Integer
Public bulletframe As Integer
Public CreditText(50) As String, i As Integer, rec As Integer
Public Colo As RGB, INIFound As Boolean
Public Const NumLines = 24


Public Sub InitializeGameEngine()

    Form1.WindowState = 2
    Randomize
    
    For x = 0 To NumOfStars
        BuildNewStar1 (x)
    Next x
    
    diffwidth = Form1.PicMain.ScaleWidth / Form1.Pic_radar.ScaleWidth
    diffheight = Form1.PicMain.ScaleHeight / Form1.Pic_radar.ScaleHeight
    difwidth = Form1.PicMain.ScaleWidth / Form1.Picture4.ScaleWidth
    difheight = Form1.PicMain.ScaleHeight / Form1.Picture4.ScaleHeight
    
    TempCalc = 0
    Tempclac = 0

    score = 0
    lives = 5
    flgs = 1
    flgm = 0
    bonus = 1500
    bulletframe = 0
    fboom = 1
    OldShipX = 0
    OldShipY = 0
    ShipX = Form1.PicMain.ScaleWidth / 2 - 25
    ShipY = Form1.PicMain.ScaleHeight - 44
    sVel = 1
    sVelocity = 10
    BulletsActivated = 0
    Health = 100
    BuildBadGuys
    DrawHealthBar
    Form1.Show
End Sub

Public Sub DrawStars()

'Form1.PicScreenBuffer.Cls
'Draw the stars to their buffer
For x = 0 To NumOfStars
    StarArray(x).y = StarArray(x).y + StarArray(x).SPEED
    If StarArray(x).y > Form1.PicMain.ScaleHeight Then BuildNewStar (x)
    SetPixelV Form1.PicScreenBuffer.hdc, StarArray(x).x, StarArray(x).y, RGB(StarArray(x).bright, StarArray(x).bright, StarArray(x).bright)
Next x

End Sub

Public Sub FireBullets()
Dim i As Integer

If Firing = 1 Then
    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 14, ShipY + 4, 7, 12, Form2.Picshotm.hdc, 9 * bulletframe, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 14, ShipY + 4, 7, 12, Form2.Picshot.hdc, 9 * bulletframe, 0, vbSrcAnd
    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 31, ShipY + 4, 4, 12, Form2.Picshotm.hdc, 9 * bulletframe, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, ShipX + 31, ShipY + 4, 4, 12, Form2.Picshot.hdc, 9 * bulletframe, 0, vbSrcAnd
    bulletframe = bulletframe + 1
    If bulletframe > 2 Then bulletframe = 0
End If
'Bullets
For x = 0 To NumOfBullets
If Firing = 1 And BulletsActivated <= 1 And Bulletl(x).Activated = 0 And Bulletr(x).Activated = 0 Then
   Bulletl(x).Activated = 1
   Bulletl(x).x = ShipX + 16
   Bulletl(x).y = ShipY + 25
   Bulletr(x).Activated = 1
   Bulletr(x).x = ShipX + 35
   Bulletr(x).y = ShipY + 25
   BulletsActivated = BulletsActivated + 1
End If

If Bulletl(x).Activated = 1 Then
    Bulletl(x).y = Bulletl(x).y - BulletSpeed
    If Bulletl(x).y < -7 Then Bulletl(x).Activated = 0
    BitBlt Form1.PicScreenBuffer.hdc, Bulletl(x).x, Bulletl(x).y, 1, 8, Form2.PicBulletM.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, Bulletl(x).x, Bulletl(x).y, 1, 8, Form2.PicBullet.hdc, 0, 0, vbSrcAnd

    For y = 0 To CurLevel.NumOfBadGuys
      If BadGuys(y).Activated = 1 And BadGuys(y).Exploding = 0 And Bulletl(x).Activated = 1 Then
         If Abs((BadGuys(y).x + (BadGuys(y).xsize / 2)) - (Bulletl(x).x + 0.5)) < (BadGuys(y).xsize / 2) And Abs((BadGuys(y).y + (BadGuys(y).ysize / 2)) - (Bulletl(x).y + 2)) < (BadGuys(y).ysize / 2) Then
            BadGuys(y).Damage = BadGuys(y).Damage + 1
            Bulletl(x).Activated = 0
            BitBlt Form1.PicScreenBuffer.hdc, Bulletl(x).x, Bulletl(x).y - 12, 12, 12, Form2.PicHITM.hdc, 12 * BadGuys(y).frame, 0, vbPatInvert
            BitBlt Form1.PicScreenBuffer.hdc, Bulletl(x).x, Bulletl(x).y - 12, 12, 12, Form2.PicHIT.hdc, 12 * BadGuys(y).frame, 0, vbSrcPaint
            BadGuys(y).frame = BadGuys(y).frame + 1
            If BadGuys(y).frame > 6 Then BadGuys(y).frame = 0
         End If
      End If
    Next y
    
End If
If Bulletr(x).Activated = 1 Then
    Bulletr(x).y = Bulletr(x).y - BulletSpeed
    If Bulletr(x).y < -7 Then Bulletr(x).Activated = 0
    BitBlt Form1.PicScreenBuffer.hdc, Bulletr(x).x, Bulletr(x).y, 1, 8, Form2.PicBulletM.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, Bulletr(x).x, Bulletr(x).y, 1, 8, Form2.PicBullet.hdc, 0, 0, vbSrcAnd

    For y = 0 To CurLevel.NumOfBadGuys
       If BadGuys(y).Activated = 1 And BadGuys(y).Exploding = 0 And Bulletr(x).Activated = 1 Then
         If Abs((BadGuys(y).x + (BadGuys(y).xsize / 2)) - (Bulletr(x).x + 0.5)) < (BadGuys(y).xsize / 2) And Abs((BadGuys(y).y + (BadGuys(y).ysize / 2)) - (Bulletr(x).y + 2)) < (BadGuys(y).ysize / 2) Then
            BadGuys(y).Damage = BadGuys(y).Damage + 1
            Bulletr(x).Activated = 0
            BitBlt Form1.PicScreenBuffer.hdc, Bulletr(x).x, Bulletr(x).y - 12, 12, 12, Form2.PicHITM.hdc, 12 * BadGuys(y).frame, 0, vbPatInvert
            BitBlt Form1.PicScreenBuffer.hdc, Bulletr(x).x, Bulletr(x).y - 12, 12, 12, Form2.PicHIT.hdc, 12 * BadGuys(y).frame, 0, vbSrcPaint
            BadGuys(y).frame = BadGuys(y).frame + 1
            If BadGuys(y).frame > 6 Then BadGuys(y).frame = 0
         End If
       End If
    Next y
    
End If

11 Next x

'Don't allow any more bullets to be created
BulletsActivated = 0

End Sub

Public Sub ScrollShip()
    'Scrolling across edges of screen
    If ShipX > PicMain.ScaleWidth Then ShipX = 0
    If ShipX < -51 Then ShipX = PicMain.ScaleWidth
    If ShipY > PicMain.ScaleHeight Then ShipY = 0
    If ShipY < -44 Then ShipY = PicMain.ScaleHeight
End Sub

Public Sub VelocityCode()

Form1.Picture4.Cls

'Movement Up
If sUp = 1 Then
sVel = sVel - VAcc
If sVel < -10 Then sVel = -10
ShipY = ShipY + sVel
If ShipY < 0 Then
   ShipY = 0
   sVel = 0
End If
End If

'Movement Down
If sDown = 1 Then
sVel = sVel + VAcc
If sVel > 10 Then sVel = 10
ShipY = ShipY + sVel
If ShipY > Form1.PicMain.ScaleHeight - 44 Then
   ShipY = Form1.PicMain.ScaleHeight - 44
   sVel = 0
End If
End If

'Vertical Deceleration
If sUp = 0 And sDown = 0 And sVel <> 0 Then
If sVel > 0 Then
sVel = sVel - VDel
If sVel <= 0 Then sVel = 0
Else
sVel = sVel + VDel
If sVel >= 0 Then sVel = 0
End If
If ShipY > Form1.PicMain.ScaleHeight - 44 Then
   ShipY = Form1.PicMain.ScaleHeight - 44
   sVel = 0
ElseIf ShipY < 0 Then
   ShipY = 0
   sVel = 0
Else
   ShipY = ShipY + sVel
End If
End If

'Movement Left
If sLeft = 1 Then
    sVelocity = sVelocity - HAcc
    If sVelocity < -10 Then sVelocity = -10
    ShipX = ShipX + sVelocity
    flgs = 2
End If

'Movement Right
If sRight = 1 Then
    sVelocity = sVelocity + HAcc
    If sVelocity > 10 Then sVelocity = 10
    ShipX = ShipX + sVelocity
    flgs = 1
End If

If ShipX >= Form1.PicMain.ScaleWidth - 51 Then
    ShipX = Form1.PicMain.ScaleWidth - 51
    sVelocity = 0
End If

If ShipX <= 0 Then
    ShipX = 0
    sVelocity = 0
End If

'Horizontal Deceleration
If ShipX > 0 And ShipX < Form1.PicMain.ScaleWidth - 51 Then
 If sRight = 0 And sLeft = 0 And sVelocity <> 0 Then
   If flgs = 1 Then
      Form2.Picture1.Picture = Form2.Pic1R.Picture
      Form2.Picture2.Picture = Form2.Pic1RM.Picture
   End If
   If flgs = 2 Then
      Form2.Picture1.Picture = Form2.Pic1l.Picture
      Form2.Picture2.Picture = Form2.Pic1lm.Picture
   End If
   flgs = 0
   If sVelocity > 0 Then
        sVelocity = sVelocity - HDel
        If sVelocity <= 0 Then sVelocity = 0
    Else
        sVelocity = sVelocity + HDel
        If sVelocity >= 0 Then sVelocity = 0
    End If
    ShipX = ShipX + sVelocity
 End If
End If

If sVelocity = 0 And sLeft = 0 And sRight = 0 Then
    Form2.Picture1.Picture = Form2.PicT.Picture
    Form2.Picture2.Picture = Form2.Pictm.Picture
    flgs = 0
End If

BitBlt Form1.Picture4.hdc, 0, ShipY / (difheight - 0.5), 115, 2, Form1.bar1.hdc, Form1.Picture4.ScaleLeft - 5, 0, vbSrcCopy
BitBlt Form1.Picture4.hdc, ShipX / (difwidth - 0.5), 0, 2, 116, Form1.bar2.hdc, 0, Form1.Picture4.ScaleTop - 5, vbSrcCopy

End Sub

Public Sub BuildNewStar1(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).x = Rnd * Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).y = Rnd * Form1.PicMain.ScaleHeight
    StarArray(ArrayVal).bright = Rnd * 255
    StarArray(ArrayVal).SPEED = Rnd * 5 + 2
End Sub

Public Sub BuildNewStar(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).x = Rnd * Form1.PicMain.ScaleWidth
    StarArray(ArrayVal).y = 0
    StarArray(ArrayVal).bright = Rnd * 200 + 55
    StarArray(ArrayVal).SPEED = Rnd * 5 + 1
End Sub

Public Sub GoodGuyStuff()

    If Health <= 0 Then Exploding = 1
    If Exploding = 1 Then
       Firing = 0
       Form1.Timer1.Enabled = False
    End If
    If Exploding = 1 And ExplodingFrame = 0 Then
        For q = 0 To 5
            Explosions(q).x = ShipX + Int(Rnd * 51)
            Explosions(q).y = ShipY + Int(Rnd * 44)
        Next q
        ExplodingFrame = 1
    End If
    If Exploding = 1 And ExplodingFrame <> 0 Then
        For q = 0 To 5
          If q <= 1 Then
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 60, 60, Form2.PicExplode1m.hdc, 60 * ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 60, 60, Form2.PicExplode1.hdc, 60 * ExplodingFrame, 0, vbSrcPaint
          ElseIf q <= 3 Then
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 75, 64, Form2.PicExplodeM.hdc, 77 * ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 75, 64, Form2.PicExplode.hdc, 77 * ExplodingFrame, 0, vbSrcPaint
          Else
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 80, 70, Form2.PicExplode2m.hdc, 80 * ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 80, 70, Form2.PicExplode2.hdc, 80 * ExplodingFrame, 0, vbSrcPaint
          End If
'          BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 75, 64, Form2.PicExplodeM.hdc, 77 * ExplodingFrame, 0, vbPatInvert
'          BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).X, Explosions(q).Y, 75, 64, Form2.PicExplode.hdc, 77 * ExplodingFrame, 0, vbSrcPaint
        Next q
        ExplodingFrame = ExplodingFrame + 1
    End If

    If ExplodingFrame >= 14 Then
         Exploding = 0
         Health = 100
         lives = lives - 1
         If lives = 4 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live4.bmp")
         If lives = 3 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live3.bmp")
         If lives = 2 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live2.bmp")
         If lives = 1 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live1.bmp")
         If lives = 0 Then Set Form1.Piclive.Picture = LoadPicture(App.Path & "\live0.bmp")
         DrawHealthBar
         ExplodingFrame = 0
         Form2.Picture1.Picture = Form2.PicFlash.Picture
         If lives < 0 Then
            GameActive = 0
            Fade Form1.PicMain, 100
            Form1.Visible = False
            ShowCursor 1
''Sound Call
            Music
            If Soundcard = True Then
               lngReturnResult = mciSendString("close all", 0&, 0, 0)
               lngReturnResult = mciSendString("open " + App.Path + "\end.mid type sequencer alias backplay", 0&, 0, 0)
               lngReturnResult = mciSendString("play backplay", 0&, 0, 0)
            End If
            Form3.Show
         End If
    End If

    If Exploding = 0 Then
        BitBlt Form1.PicScreenBuffer.hdc, ShipX, ShipY, 51, 44, Form2.Picture2.hdc, 0, 0, vbMergePaint
        BitBlt Form1.PicScreenBuffer.hdc, ShipX, ShipY, 51, 44, Form2.Picture1.hdc, 0, 0, vbSrcAnd
    End If
    
BitBlt Form1.PicMain.hdc, 0, 0, BufferWidth, BufferHeight, Form1.PicScreenBuffer.hdc, 0, 0, vbSrcCopy

End Sub

Public Sub BuildBadGuys()

If stage = 1 Then
        CurLevel.Damage = 5
        CurLevel.NumOfBadGuys = 6
        CurLevel.Velocity = 8
        CurLevel.Damagelimit = 30
'        CurLevel.BulletSpeed = 10
'        CurLevel.OddsOfFiring = 40
End If
If stage = 2 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 14
        CurLevel.Velocity = 6
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 3 Then
        CurLevel.Damage = 5
        CurLevel.NumOfBadGuys = 6
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 30
        CurLevel.BulletSpeed = 15
        CurLevel.OddsOfFiring = 2
End If
If stage = 4 Then
        CurLevel.Damage = 3
        CurLevel.NumOfBadGuys = 9
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 5 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 1
        CurLevel.Velocity = 2
        CurLevel.Damagelimit = 60
        CurLevel.BulletSpeed = 15
        CurLevel.OddsOfFiring = 10
End If
If stage = 6 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 30
        CurLevel.Velocity = 8
        CurLevel.Damagelimit = 40
'        CurLevel.BulletSpeed = 20
'        CurLevel.OddsOfFiring = 2
End If
If stage = 7 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 1
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 200
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 8 Then
        CurLevel.Damage = 5
        CurLevel.NumOfBadGuys = 5
        CurLevel.Velocity = 6
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 10
        CurLevel.OddsOfFiring = 2
End If
If stage = 9 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 6
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 50
        CurLevel.BulletSpeed = 15
        CurLevel.OddsOfFiring = 2
End If
If stage = 10 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 6
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 15
        CurLevel.OddsOfFiring = 2
End If
If stage = 11 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 0
        CurLevel.Velocity = 1
        CurLevel.Damagelimit = 400
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 20
End If
If stage = 12 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 11
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 13 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 30
        CurLevel.Velocity = 35
        CurLevel.Damagelimit = 10
End If
If stage = 14 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 20
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 50
        CurLevel.BulletSpeed = 30
        CurLevel.OddsOfFiring = 6
End If
If stage = 15 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 14
        CurLevel.Velocity = 10
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 30
        CurLevel.OddsOfFiring = 5
End If
If stage = 16 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 14
        CurLevel.Velocity = 6
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 17 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 62
        CurLevel.Velocity = 2
        CurLevel.Damagelimit = 10
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 18 Then
        CurLevel.Damage = 3
        CurLevel.NumOfBadGuys = 9
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 30
        CurLevel.OddsOfFiring = 2
End If
If stage = 19 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 20
        CurLevel.Velocity = 6
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 6
End If
If stage = 20 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 21
        CurLevel.Velocity = 6
        CurLevel.Damagelimit = 60
        CurLevel.BulletSpeed = 30
        CurLevel.OddsOfFiring = 2
End If
If stage = 21 Then
        CurLevel.Damage = 2
        CurLevel.NumOfBadGuys = 6
        CurLevel.Velocity = 5
        CurLevel.Damagelimit = 40
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 2
End If
If stage = 22 Then
        CurLevel.Damage = 1
        CurLevel.NumOfBadGuys = 0
        CurLevel.Velocity = 1
        CurLevel.Damagelimit = 800
        CurLevel.BulletSpeed = 20
        CurLevel.OddsOfFiring = 20
End If
TempCalc = 0
Tempclac = 0
flgm = 0
Form1.Tmr_flgm.Interval = 1
Form1.Tmr_flgm.Enabled = False

ReDim BadGuys(0 To CurLevel.NumOfBadGuys) As BadGuy

For x = 0 To CurLevel.NumOfBadGuys
Randomize
     BadGuys(x).Activated = 1
     BadGuys(x).x = Rnd * ((Form1.PicMain.ScaleWidth - 70) - Form1.PicMain.ScaleLeft)
     BadGuys(x).y = Form1.PicMain.ScaleTop - 100
     BadGuys(x).DstX = Rnd * ((Form1.PicMain.ScaleWidth - 70) - Form1.PicMain.ScaleLeft)
'     BadGuys(X).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
     BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
     BadGuys(x).Velocity = Rnd * CurLevel.Velocity
     If BadGuys(x).Velocity <= 2 Then BadGuys(x).Velocity = BadGuys(x).Velocity + 4
     BadGuys(x).Damage = CurLevel.Damage
     BadGuys(x).Exploding = 0
     BadGuys(x).frame = 0
     BadGuys(x).ExplodingFrame = 0
     bullets = False
     numsec = 0
     frame = 0
     appo = False
Next x

End Sub

Public Sub MoveAndDrawBack()

'Get the exact Yposition relative to the background tiles
If BackYPos >= 600 And BackYPos < 800 Then
    
    BitBlt Form1.PicScreenBuffer.hdc, 0, 0, ViewportWidth, ViewportHeight, BackTile1, 0, FirstTileHeight - BackYPos, vbSrcCopy
        
'ElseIf BackYPos >= 800 And BackYPos < 1200 Then
'
'    OverlapTop = BackYPos - FirstTileHeight
'    OverlapBottom = ViewportHeight - OverlapTop
'
'    'draw the top first
'    BitBlt Form1.PicScreenBuffer.HDC, 0, 0, ViewportWidth, OverlapTop, BackTile2, 0, FirstTileHeight - OverlapTop, vbSrcCopy
'    BitBlt Form1.PicScreenBuffer.HDC, 0, OverlapTop, ViewportWidth, OverlapBottom, BackTile1, 0, 0, vbSrcCopy
'
'ElseIf BackYPos >= 1200 And BackYPos < 1600 Then
'
'    BitBlt HDC, 0, 0, ViewportWidth, ViewportHeight, BackTile2, 0, SecondTileHeight - BackYPos, vbSrcCopy
'
'ElseIf BackYPos >= 1600 And BackYPos < 2000 Then
'
'    OverlapTop = BackYPos - SecondTileHeight
'    OverlapBottom = ViewportHeight - OverlapTop
'
'    'draw the top first
'    BitBlt Form1.PicScreenBuffer.HDC, 0, 0, ViewportWidth, OverlapTop, BackTile3, 0, FirstTileHeight - OverlapTop, vbSrcCopy
'    'then the reminder of the previous back tile
'    BitBlt Form1.PicScreenBuffer.HDC, 0, OverlapTop, ViewportWidth, OverlapBottom, BackTile2, 0, 0, vbSrcCopy
'
'ElseIf BackYPos >= 2000 And BackYPos < 2400 Then
'
'    BitBlt Form1.PicScreenBuffer.HDC, 0, 0, ViewportWidth, ViewportHeight, BackTile3, 0, ThirdTileHeight - BackYPos, vbSrcCopy
'
'ElseIf BackYPos >= 2400 And BackYPos < 2800 Then
'
'    OverlapTop = BackYPos - ThirdTileHeight
'    OverlapBottom = ViewportHeight - OverlapTop
'
'    'draw the top first
'    BitBlt Form1.PicScreenBuffer.HDC, 0, 0, ViewportWidth, OverlapTop, BackTile4, 0, FirstTileHeight - OverlapTop, vbSrcCopy
'    'then the reminder of the previous back tile
'    BitBlt Form1.PicScreenBuffer.HDC, 0, OverlapTop, ViewportWidth, OverlapBottom, BackTile3, 0, 0, vbSrcCopy
'
'ElseIf BackYPos >= 2800 And BackYPos < 3200 Then
'
'    BitBlt Form1.PicScreenBuffer.HDC, 0, 0, ViewportWidth, ViewportHeight, BackTile4, 0, FourthTileHeight - BackYPos, vbSrcCopy
'
ElseIf BackYPos >= 800 Then
    
    BackYPos = 0

ElseIf BackYPos >= 0 And BackYPos < 600 Then

    
    OverlapTop = BackYPos
    OverlapBottom = ViewportHeight - OverlapTop
    
    'draw the top first
    BitBlt Form1.PicScreenBuffer.hdc, 0, 0, ViewportWidth, OverlapTop, BackTile1, 0, FirstTileHeight - OverlapTop, vbSrcCopy
    'then the reminder of the previous back tile
    BitBlt Form1.PicScreenBuffer.hdc, 0, OverlapTop, ViewportWidth, OverlapBottom, BackTile4, 0, 0, vbSrcCopy

End If

BackYPos = BackYPos + 1

End Sub
Public Sub cleanup()
  DeleteGeneratedDC BackTile1
  DeleteGeneratedDC BackTile2
  DeleteGeneratedDC BackTile3
  DeleteGeneratedDC BackTile4
End Sub
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(ByVal 0&)

If DC < 1 Then
    GenerateDC = 0
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_DEFAULTSIZE Or LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Public Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function
Public Sub AngleRadians()

    deltax = (ShipX + 25) - curX
    deltay = (ShipY - 22) - curY
    
    If deltax = 0 Then      'Vertical
        If deltay < 0 Then
            Angle = Pi / 2
        Else
'            Angle = Pi * 2
            Angle = Pi * 1.5
        End If
    
    ElseIf deltay = 0 Then  'Horizontal
        If deltax >= 0 Then
            Angle = 0
        Else
            Angle = Pi
        End If
    
    Else
        'Note: ++ = positive X, positive Y; +- = positive X, negative Y; etc.
        'On a true coordinate plane, Y increases as it move upward.
        'In VB coordinates, Y is reversed. It increases as it moves downward.
        
        'Calc for true Upper Right Quadrant (++) (For VB this is +-)
        Angle = Atn(Abs(deltay / deltax))        'VB Upper Right (+-)
        
        'Correct for other 3 quadrants in VB coordinates (Reversed Y)
        If deltax >= 0 And deltay >= 0 Then       'VB Lower Right (++)
            Angle = (Pi * 2) - Angle
            
        ElseIf deltax < 0 And deltay >= 0 Then    'VB Lower Left (-+)
            Angle = Pi + Angle
            
        ElseIf deltax < 0 And deltay < 0 Then     'VB Upper Left (--)
            Angle = Pi - Angle
            
        End If
        
    End If
    Angle = Angle * (180# / Pi)
End Sub
Public Sub Fade(Pic As PictureBox, Blocks As Integer)
    
    Dim width_section_size As Integer
    Dim height_section_size As Integer
    Dim i As Integer, j As Integer
    Dim save_color As Long
    
    'Saves the picbox's current forecolor
    save_color = Pic.ForeColor

    'Set Pics forecolor to its backcolor
    Pic.ForeColor = Pic.BackColor

    'Corrects the Blocks if needed
    If Blocks < 5 Then Blocks = 5
    If Blocks > 100 Then Blocks = 100

    'Sets the size of each width section
    width_section_size = Pic.ScaleWidth / Blocks

    'Sets the size of each height section
    height_section_size = Pic.ScaleHeight / Blocks

    For i = (Blocks / 2) To 0 Step -1
        Sleep (20)
        Pic.Line (i * width_section_size, i * height_section_size)-(((Blocks - i) + 1) * width_section_size, ((Blocks - i) + 1) * height_section_size), , BF
    Next

    'Restores the picbox's original forecolor
    Pic.ForeColor = save_color
        
End Sub
Public Sub missions()
If mission = 1 Then
   BackTile1 = GenerateDC(App.Path & "\space1.bmp")
   BackTile2 = GenerateDC(App.Path & "\space1.bmp")
   BackTile3 = GenerateDC(App.Path & "\space1.bmp")
   BackTile4 = GenerateDC(App.Path & "\space1.bmp")
End If
If mission = 2 Then
   BackTile1 = GenerateDC(App.Path & "\land1.bmp")
   BackTile2 = GenerateDC(App.Path & "\land1.bmp")
   BackTile3 = GenerateDC(App.Path & "\land1.bmp")
   BackTile4 = GenerateDC(App.Path & "\land1.bmp")
   BackYPos = 600
End If
If mission = 3 Then
   BackTile1 = GenerateDC(App.Path & "\space1.bmp")
   BackTile2 = GenerateDC(App.Path & "\space1.bmp")
   BackTile3 = GenerateDC(App.Path & "\space1.bmp")
   BackTile4 = GenerateDC(App.Path & "\space1.bmp")
   BackYPos = 600
End If
End Sub
Public Sub Blend(Destination As Object, Source As Object, Amount As Integer, x, y, X2, Y2)
AlphaBlending Destination.hdc, x, y, X2, Y2, Source.hdc, x, y, X2, Y2, Amount
End Sub

Public Function RollIT()
    Dim nRet As Long
    Const RD As Single = 350
    Const ANG_SIZE As Integer = 50
    Const SPEED As Integer = 20
    Static Rolling As Single
    nRet = SetWindowRgn(Form1.Pic_radarv.hwnd, CreateRegion(Rolling, ANG_SIZE, 110, -34, -34), True)
    nRet = SetWindowRgn(Form1.Pic_radar.hwnd, CreateFrame(150, 0, 0), True)
    Rolling = Rolling + SPEED
    If (Rolling) > 360 Then Rolling = 0
End Function

Public Sub TriAngl(ByVal vAng1 As Single, vAng2 As Single, _
                    ByVal vOffsetX As Single, ByVal vOffestY As Single, _
                    ByVal vR As Single)
    Dim KerenAng As Single
    Dim yeter As Single
    Dim b2 As Single
    Dim q As Single
    Dim n As Single
    Dim m As Single
        
    PolyPoints(0).x = vOffsetX
    PolyPoints(0).y = vOffestY
    
    PolyPoints(0).x = vOffsetX + vR
    PolyPoints(0).y = vOffestY + vR
    
    PolyPoints(1).x = (Sin(vAng1 * (Pi / 180)) * vR) + PolyPoints(0).x
    PolyPoints(1).y = (Cos(vAng1 * (Pi / 180)) * (-1) * vR) + PolyPoints(0).y
    
    PolyPoints(4).x = (Sin((vAng1 + vAng2) * (Pi / 180)) * vR) + PolyPoints(0).x
    PolyPoints(4).y = (Cos((vAng1 + vAng2) * (Pi / 180)) * (-1) * vR) + PolyPoints(0).y
    
    
    KerenAng = vAng1 + (vAng2 / 2)
    PolyPoints(3).x = (Sin((KerenAng) * (Pi / 180)) * vR) + PolyPoints(4).x
    PolyPoints(3).y = (Cos((KerenAng) * (Pi / 180)) * (-1) * vR) + PolyPoints(4).y
    
    PolyPoints(2).x = (Sin((KerenAng) * (Pi / 180)) * vR) + PolyPoints(1).x
    PolyPoints(2).y = (Cos((KerenAng) * (Pi / 180)) * (-1) * vR) + PolyPoints(1).y
End Sub

Public Function CreateRegion(Ang As Single, angOffsett As Single, Radius As Single, _
                            OffsetX As Integer, OffsetY As Integer) As Long
    Dim Corraction As Integer
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long
    Dim i As Integer
    
    ResultRegion = CreateRectRgn(10, 10, 10, 10)
    
    'set the rgn
    Call TriAngl(Ang, angOffsett, OffsetX, OffsetY, Radius)
    
    
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 5, 1)

    HolderRegion = CreateEllipticRgn(PolyPoints(0).x - Radius, PolyPoints(0).y - Radius, _
                                    PolyPoints(0).x + Radius, PolyPoints(0).y + Radius)
            
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, RGN_AND)
    
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateRegion = ResultRegion
End Function

Public Function CreateFrame(ByVal Radius As Single, _
                ByVal OffsetX As Single, ByVal OffsetY As Single) As Long
    Dim Corraction As Integer
    Dim HolderRegion As Long, nRet As Long
    Dim i As Integer
    ResultRegion = CreateRectRgn(10, 10, 10, 10)
    
    
    HolderRegion = CreateEllipticRgn(PolyPoints(0).x - Radius + OffsetX, _
                                    PolyPoints(0).y - Radius + OffsetY, _
                                    PolyPoints(0).x + Radius + OffsetX, _
                                    PolyPoints(0).y + Radius + OffsetY)
    nRet = CombineRgn(ResultRegion, HolderRegion, HolderRegion, RGN_AND)
    
    DeleteObject HolderRegion
    CreateFrame = ResultRegion
End Function
