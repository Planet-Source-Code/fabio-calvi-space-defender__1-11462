Attribute VB_Name = "Enemy9"
Public Sub badguy9()

Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

For x = 0 To CurLevel.NumOfBadGuys
 
If BadGuys(x).Activated = 0 Then GoTo 10
    Set BadGuys(x).PicT = Form2.PicB
    Set BadGuys(x).mask = Form2.PicBM
    BadGuys(x).bulletlxpos = 15
    BadGuys(x).bulletlypos = 40
    BadGuys(x).bulletrxpos = 55
    BadGuys(x).bulletrypos = 40
    BadGuys(x).bulletcxpos = 35
    BadGuys(x).bulletcypos = 20
    BadGuys(x).xsize = 76
    BadGuys(x).ysize = 60
    BadGuys(x).oldX = BadGuys(x).x
    BadGuys(x).oldY = BadGuys(x).y
    If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
    If BadGuys(x).DstX < BadGuys(x).x Then
        BadGuys(x).x = BadGuys(x).x - BadGuys(x).Velocity
    Else
        BadGuys(x).x = BadGuys(x).x + BadGuys(x).Velocity
    End If
    If BadGuys(x).y >= (Form1.PicMain.ScaleTop - 20) And flgm = 0 Then
       flgm = 1
       Form1.Tmr_flgm.Enabled = True
       BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight)
    End If
    If flgm = 1 Then
       If BadGuys(x).DstY < BadGuys(x).y Then
          BadGuys(x).y = BadGuys(x).y - BadGuys(x).Velocity
       Else
          BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
       End If
     End If
 
    If flgm = 0 Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
    If flgm = 2 Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity + 1
    
    If Abs(BadGuys(x).x - BadGuys(x).DstX) < CurLevel.Velocity + 1 Then BadGuys(x).DstX = Rnd * ((Form1.PicMain.ScaleWidth - 70) - Form1.PicMain.ScaleLeft)
    If Abs(BadGuys(x).y - BadGuys(x).DstY) < CurLevel.Velocity + 1 Then BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
    
    BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, 0, 0, vbSrcAnd
    BitBlt Form1.Pic_radar.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad.hdc, 0, 0, vbSrcCopy
    If roll = True Then BitBlt Form1.Pic_radarv.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad2.hdc, 0, 0, vbSrcCopy
    
    End If
    If BadGuys(x).Damage > CurLevel.Damagelimit Then BadGuys(x).Exploding = 1

If BadGuys(x).Activated = 1 Then
   BadGuys(x).Firing = Int(Rnd * CurLevel.OddsOfFiring)
Else
   BadGuys(x).Firing = 0
End If

'Firing bullets
If BadGuys(x).y > Form1.PicMain.ScaleTop + 20 Then
For y = 0 To 1
If BadGuys(x).Firing = 1 And BadGuys(x).BulletsActivated <= 1 Then
If BadGuys(x).Activated > 0 Then
If BadGuys(x).Bulletl(y).Activated = 0 Then
   BadGuys(x).Bulletl(y).Activated = 1
   BadGuys(x).Bulletl(y).x = BadGuys(x).x + BadGuys(x).bulletlxpos
   BadGuys(x).Bulletl(y).y = BadGuys(x).y + BadGuys(x).bulletlypos
End If
If BadGuys(x).Bulletr(y).Activated = 0 Then
   BadGuys(x).Bulletr(y).Activated = 1
   BadGuys(x).Bulletr(y).x = BadGuys(x).x + BadGuys(x).bulletrxpos
   BadGuys(x).Bulletr(y).y = BadGuys(x).y + BadGuys(x).bulletrypos
End If
If BadGuys(x).Bulletc(y).Activated = 0 Then
   BadGuys(x).Bulletc(y).Activated = 1
   BadGuys(x).Bulletc(y).x = BadGuys(x).x + BadGuys(x).bulletcxpos
   BadGuys(x).Bulletc(y).y = BadGuys(x).y + BadGuys(x).bulletcypos
   curX = BadGuys(x).Bulletc(y).x
   curY = BadGuys(x).Bulletc(y).y
   AngleRadians
   BadGuys(x).Bulletca(y) = Angle
End If
BadGuys(x).BulletsActivated = BadGuys(x).BulletsActivated + 1
End If
End If

11 Next y
End If
'Don't allow any more bullets to be created
'BadGuys(X).BulletsActivated = 0
If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(x).x, BadGuys(x).y, BadGuys(x).mask, Form2.PicTemp) Then
   Exploding = 1
   BadGuys(x).Exploding = 1
'   BadGuys(X).Activated = 0
End If

10 Next x

'**********************************************************

For x = 0 To CurLevel.NumOfBadGuys
For y = 0 To 1
If BadGuys(x).Bulletl(y).Activated = 1 Then
bullets = True
'BadGuys(X).Bulletl(Y).X = BadGuys(X).Bulletl(Y).X + Cos(5 * Radians)
BadGuys(x).Bulletl(y).y = BadGuys(x).Bulletl(y).y + CurLevel.BulletSpeed
If BadGuys(x).Bulletl(y).y > Form1.PicMain.ScaleHeight Then
   BadGuys(x).Bulletl(y).Activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 6, 6, Form2.PicBBulletM.hdc, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 6, 6, Form2.PicBBullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletl(y).x + Form2.PicBBullet.ScaleWidth > ShipX And BadGuys(x).Bulletl(y).x < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(x).Bulletl(y).y - 25) - ShipY) < 18 Then
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 10
   UpdateHealth
   BadGuys(x).Bulletl(y).Activated = 0
   bullets = False
End If
End If
 
If BadGuys(x).Bulletr(y).Activated = 1 Then
bullets = True
BadGuys(x).Bulletr(y).y = BadGuys(x).Bulletr(y).y + CurLevel.BulletSpeed
If BadGuys(x).Bulletr(y).y > Form1.PicMain.ScaleHeight Then
   BadGuys(x).Bulletr(y).Activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 6, 6, Form2.PicBBulletM.hdc, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 6, 6, Form2.PicBBullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletr(y).x + Form2.PicBBullet.ScaleWidth > ShipX And BadGuys(x).Bulletr(y).x < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(x).Bulletr(y).y - 25) - ShipY) < 18 Then
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 10
   UpdateHealth
   BadGuys(x).Bulletr(y).Activated = 0
   bullets = False
End If
End If

If BadGuys(x).Bulletc(y).Activated = 1 Then
bullets = True
BadGuys(x).Bulletc(y).x = BadGuys(x).Bulletc(y).x + 10 * Cos(BadGuys(x).Bulletca(y) * Radians)
BadGuys(x).Bulletc(y).y = BadGuys(x).Bulletc(y).y - (CurLevel.BulletSpeed + 5) * Sin(BadGuys(x).Bulletca(y) * Radians)  'CurLevel.BulletSpeed + 2
If (BadGuys(x).Bulletc(y).x > Form1.PicMain.ScaleWidth) Or _
   (BadGuys(x).Bulletc(y).x < Form1.PicMain.ScaleLeft) Or _
   (BadGuys(x).Bulletc(y).y > Form1.PicMain.ScaleHeight) Or _
   (BadGuys(x).Bulletc(y).y < Form1.PicMain.ScaleTop) Then
    BadGuys(x).Bulletc(y).Activated = 0
    bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 8, Form2.PicABullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletc(y).x + Form2.PicABullet.ScaleWidth > ShipX And BadGuys(x).Bulletc(y).x < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(x).Bulletc(y).y - 25) - ShipY) < 18 Then
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 10
   UpdateHealth
   BadGuys(x).Bulletc(y).Activated = 0
   bullets = False
End If
End If
Next y
'if comment BadGuys shoot once only
BadGuys(x).BulletsActivated = 0
Next x

For x = 0 To CurLevel.NumOfBadGuys
    If BadGuys(x).Activated = 1 And BadGuys(x).y > Form1.PicMain.ScaleHeight Then
       Tempclac = Tempclac + 1
       BadGuys(x).Activated = 0
    End If
    If BadGuys(x).Exploding = 1 Then
       BadGuys(x).Activated = 0
       If fboom = 1 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 75, 64, Form2.PicExplodeM.hdc, 77 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 75, 64, Form2.PicExplode.hdc, 77 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
       ElseIf fboom = 2 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 60, 60, Form2.PicExplode1m.hdc, 60 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 60, 60, Form2.PicExplode1.hdc, 60 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
       Else
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 80, 70, Form2.PicExplode2m.hdc, 80 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 80, 70, Form2.PicExplode2.hdc, 80 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
       End If
       BadGuys(x).ExplodingFrame = BadGuys(x).ExplodingFrame + 1
       If BadGuys(x).ExplodingFrame = 13 Then
          BadGuys(x).Exploding = 0
          score = score + 80
          TempCalc = TempCalc + 1
          fboom = fboom + 1
          If fboom > 3 Then fboom = 1
       End If
    End If
Next x
   
If Tempclac + TempCalc >= CurLevel.NumOfBadGuys + 1 And bullets = False Then
   flgm = 0
   Form1.Tmr_flgm.Interval = 1
   Form1.Tmr_flgm.Enabled = False
   stage = 0
End If

End Sub



