Attribute VB_Name = "Enemy10"
Public Sub badguy10()

Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

Dim t As Integer
For x = 0 To CurLevel.NumOfBadGuys
 
If BadGuys(x).Activated = 0 Then GoTo 10
    Set BadGuys(x).PicT = Form2.PicBoss1
    Set BadGuys(x).mask = Form2.PicBoss1m
    BadGuys(x).bulletlxpos = 35
    BadGuys(x).bulletlypos = 101
    BadGuys(x).bulletrxpos = 70
    BadGuys(x).bulletrypos = 101
    BadGuys(x).bulletcxpos = 40
    BadGuys(x).bulletcypos = 50
    BadGuys(x).xsize = 111
    BadGuys(x).ysize = 103
    BadGuys(x).oldX = BadGuys(x).x
    BadGuys(x).oldY = BadGuys(x).y
    If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
    If BadGuys(x).DstX < BadGuys(x).x Then
        BadGuys(x).x = (BadGuys(x).x - BadGuys(x).Velocity) + 2
    Else
        BadGuys(x).x = BadGuys(x).x + BadGuys(x).Velocity
    End If
    If BadGuys(x).y >= (Form1.PicMain.ScaleTop - 20) And flgm = 0 Then
       flgm = 1
       BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
    End If
    If flgm = 1 Then
       If BadGuys(x).DstY < BadGuys(x).y Then
          BadGuys(x).y = (BadGuys(x).y - BadGuys(x).Velocity) + 2
       Else
          BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
       End If
     End If
    If flgm = 0 Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
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
If BadGuys(x).y > Form1.PicMain.ScaleTop + 10 Then
For y = 0 To 9
If BadGuys(x).Activated > 0 Then
If BadGuys(x).Firing = 1 Then
If BadGuys(x).Bulletl(y).Activated = 0 Then
   BadGuys(x).Bulletl(y).Activated = 1
   If y > 0 Then
      BadGuys(x).bulletlxpos = 22
      BadGuys(x).bulletlypos = 67
   End If
   BadGuys(x).Bulletl(y).x = BadGuys(x).x + BadGuys(x).bulletlxpos + Int(Rnd * (y * 10))
   BadGuys(x).Bulletl(y).y = BadGuys(x).y + BadGuys(x).bulletlypos + Int(Rnd * (y * 10))
End If
If BadGuys(x).Bulletr(y).Activated = 0 Then
   BadGuys(x).Bulletr(y).Activated = 1
      If y > 0 Then
      BadGuys(x).bulletrxpos = 84
      BadGuys(x).bulletrypos = 67
   End If
   BadGuys(x).Bulletr(y).x = BadGuys(x).x + BadGuys(x).bulletrxpos + Int(Rnd * (y * 10))
   BadGuys(x).Bulletr(y).y = BadGuys(x).y + BadGuys(x).bulletrypos + Int(Rnd * (y * 10))
End If
If BadGuys(x).Bulletc(y).Activated = 0 Then
   BadGuys(x).Bulletc(y).Activated = 1
   BadGuys(x).Bulletc(y).x = BadGuys(x).x + BadGuys(x).bulletcxpos + Int(Rnd * (y * 10))
   BadGuys(x).Bulletc(y).y = BadGuys(x).y + BadGuys(x).bulletcypos + Int(Rnd * (y * 10))
   curX = BadGuys(x).Bulletc(y).x
   curY = BadGuys(x).Bulletc(y).y
   AngleRadians
   BadGuys(x).Bulletca(y) = Angle
End If
End If
End If

11 Next y
End If

If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(x).x, BadGuys(x).y, BadGuys(x).mask, Form2.PicTemp) Then
   Exploding = 1
   BadGuys(x).Damage = BadGuys(x).Damage + 10
End If

10 Next x

'**********************************************************

For x = 0 To CurLevel.NumOfBadGuys
For y = 0 To 9
If y = 0 Then
 If BadGuys(x).Bulletl(y).Activated = 1 Then
    BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 9, Form1.PicMain.ScaleHeight - BadGuys(x).Bulletl(y).y, Form2.Piclaser2m.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 9, Form1.PicMain.ScaleHeight - BadGuys(x).Bulletl(y).y, Form2.Piclaser2.hdc, 0, 0, vbSrcAnd
    If BadGuys(x).Bulletl(y).x > ShipX And BadGuys(x).Bulletl(y).x < ShipX + Form2.Picture1.ScaleWidth And _
       ShipY > Form1.PicMain.ScaleTop + BadGuys(x).Bulletl(y).y Then
       Exploding = 1
    End If
    BadGuys(x).Bulletl(y).Activated = 0
 End If
 
 If BadGuys(x).Bulletr(y).Activated = 1 Then
    BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 9, Form1.PicMain.ScaleHeight - BadGuys(x).Bulletr(y).y, Form2.Piclaser2m.hdc, 0, 0, vbMergePaint
    BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 9, Form1.PicMain.ScaleHeight - BadGuys(x).Bulletr(y).y, Form2.Piclaser2.hdc, 0, 0, vbSrcAnd
    If BadGuys(x).Bulletr(y).x > ShipX And BadGuys(x).Bulletr(y).x < ShipX + Form2.Picture1.ScaleWidth And _
       ShipY > Form1.PicMain.ScaleTop + BadGuys(x).Bulletr(y).y Then
       Exploding = 1
    End If
    BadGuys(x).Bulletr(y).Activated = 0
 End If
If BadGuys(x).Bulletc(y).Activated = 1 Then
 BadGuys(x).Bulletc(y).x = BadGuys(x).Bulletc(y).x + 10 * Cos(BadGuys(x).Bulletca(y) * Radians)
 BadGuys(x).Bulletc(y).y = BadGuys(x).Bulletc(y).y - (20) * Sin(BadGuys(x).Bulletca(y) * Radians)
 If (BadGuys(x).Bulletc(y).x > Form1.PicMain.ScaleWidth) Or _
    (BadGuys(x).Bulletc(y).x < Form1.PicMain.ScaleLeft) Or _
    (BadGuys(x).Bulletc(y).y > Form1.PicMain.ScaleHeight) Or _
    (BadGuys(x).Bulletc(y).y < Form1.PicMain.ScaleTop) Then
     BadGuys(x).Bulletc(y).Activated = 0
 End If
 BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
 BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 8, Form2.PicA4Bullet.hdc, 0, 0, vbSrcAnd
 If Abs((BadGuys(x).Bulletc(y).x - 32) - ShipX) < 25 And Abs((BadGuys(x).Bulletc(y).y - 32) - ShipY) < 25 Then
  If BadGuys(x).Bulletc(y).x + Form2.PicABullet.ScaleWidth > ShipX And BadGuys(x).Bulletc(y).x < ShipX + Form2.Picture1.ScaleWidth And _
     Abs((BadGuys(x).Bulletc(y).y - 25) - ShipY) < 18 Then
     Form2.Picture1.Picture = Form2.PicFlash.Picture
     Health = Health - 10
     UpdateHealth
     BadGuys(x).Bulletc(y).Activated = 0
  End If
 End If
End If
End If

If y > 0 Then
 If BadGuys(x).Bulletl(y).Activated = 1 Then
  BadGuys(x).Bulletl(y).x = (BadGuys(x).Bulletl(y).x - Cos(45 * Radians)) - 10
  BadGuys(x).Bulletl(y).y = BadGuys(x).Bulletl(y).y + 20
 If (BadGuys(x).Bulletl(y).x >= Form1.PicMain.ScaleWidth) Or _
    (BadGuys(x).Bulletl(y).x <= Form1.PicMain.ScaleLeft) Or _
    (BadGuys(x).Bulletl(y).y >= Form1.PicMain.ScaleHeight) Or _
    (BadGuys(x).Bulletl(y).y <= Form1.PicMain.ScaleTop) Then
     BadGuys(x).Bulletl(y).Activated = 0
  End If
  BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
  BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 8, 8, Form2.PicA2Bullet.hdc, 0, 0, vbSrcAnd
  If BadGuys(x).Bulletl(y).x + Form2.PicBBullet.ScaleWidth > ShipX And BadGuys(x).Bulletl(y).x < ShipX + Form2.Picture1.ScaleWidth And _
     Abs((BadGuys(x).Bulletl(y).y - 25) - ShipY) < 18 Then
     Form2.Picture1.Picture = Form2.PicFlash.Picture
     Health = Health - 10
     UpdateHealth
     BadGuys(x).Bulletl(y).Activated = 0
  End If
 End If
 If BadGuys(x).Bulletr(y).Activated = 1 Then
  BadGuys(x).Bulletr(y).x = (BadGuys(x).Bulletr(y).x + Cos(45 * Radians)) + 10
  BadGuys(x).Bulletr(y).y = BadGuys(x).Bulletr(y).y + 20
 If (BadGuys(x).Bulletr(y).x >= Form1.PicMain.ScaleWidth) Or _
    (BadGuys(x).Bulletr(y).x <= Form1.PicMain.ScaleLeft) Or _
    (BadGuys(x).Bulletr(y).y >= Form1.PicMain.ScaleHeight) Or _
    (BadGuys(x).Bulletr(y).y <= Form1.PicMain.ScaleTop) Then
     BadGuys(x).Bulletr(y).Activated = 0
 End If
 BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
 BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 8, 8, Form2.PicA2Bullet.hdc, 0, 0, vbSrcAnd
 If BadGuys(x).Bulletr(y).x + Form2.PicBBullet.ScaleWidth > ShipX And BadGuys(x).Bulletr(y).x < ShipX + Form2.Picture1.ScaleWidth And _
     Abs((BadGuys(x).Bulletr(y).y - 25) - ShipY) < 18 Then
     Form2.Picture1.Picture = Form2.PicFlash.Picture
     Health = Health - 10
     UpdateHealth
     BadGuys(x).Bulletr(y).Activated = 0
  End If
 End If
End If
If BadGuys(x).Bulletc(y).Activated = 1 Then
 BadGuys(x).Bulletc(y).x = BadGuys(x).Bulletc(y).x + 10 * Cos(BadGuys(x).Bulletca(y) * Radians)
 BadGuys(x).Bulletc(y).y = BadGuys(x).Bulletc(y).y - (20) * Sin(BadGuys(x).Bulletca(y) * Radians)
 If (BadGuys(x).Bulletc(y).x > Form1.PicMain.ScaleWidth) Or _
    (BadGuys(x).Bulletc(y).x < Form1.PicMain.ScaleLeft) Or _
    (BadGuys(x).Bulletc(y).y > Form1.PicMain.ScaleHeight) Or _
    (BadGuys(x).Bulletc(y).y < Form1.PicMain.ScaleTop) Then
     BadGuys(x).Bulletc(y).Activated = 0
 End If
 BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
 BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 8, Form2.PicABullet.hdc, 0, 0, vbSrcAnd
 If Abs((BadGuys(x).Bulletc(y).x - 32) - ShipX) < 25 And Abs((BadGuys(x).Bulletc(y).y - 32) - ShipY) < 25 Then
  If BadGuys(x).Bulletc(y).x + Form2.PicABullet.ScaleWidth > ShipX And BadGuys(x).Bulletc(y).x < ShipX + Form2.Picture1.ScaleWidth And _
     Abs((BadGuys(x).Bulletc(y).y - 25) - ShipY) < 18 Then
     Form2.Picture1.Picture = Form2.PicFlash.Picture
     Health = Health - 10
     UpdateHealth
     BadGuys(x).Bulletc(y).Activated = 0
  End If
 End If
End If

Next y
'If comment BadGuys shoot once only
'BadGuys(X).BulletsActivated = 0
Next x

For x = 0 To CurLevel.NumOfBadGuys
      
    If BadGuys(x).Exploding = 1 And BadGuys(x).ExplodingFrame = 0 Then
       BadGuys(x).Activated = 0
       For q = 0 To 9
           Explosions(q).x = BadGuys(x).x + Int(Rnd * 111)
           Explosions(q).y = BadGuys(x).y + Int(Rnd * 103)
       Next q
       BadGuys(x).ExplodingFrame = 1
    End If
    If BadGuys(x).Exploding = 1 And BadGuys(x).ExplodingFrame <> 0 Then
        For q = 0 To 9
          If q <= 3 Then
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 60, 60, Form2.PicExplode1m.hdc, 60 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 60, 60, Form2.PicExplode1.hdc, 60 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
          ElseIf q <= 6 Then
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 75, 64, Form2.PicExplodeM.hdc, 77 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 75, 64, Form2.PicExplode.hdc, 77 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
          Else
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 80, 70, Form2.PicExplode2m.hdc, 80 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
             BitBlt Form1.PicScreenBuffer.hdc, Explosions(q).x, Explosions(q).y, 80, 70, Form2.PicExplode2.hdc, 80 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
          End If
        Next q
        BadGuys(x).ExplodingFrame = BadGuys(x).ExplodingFrame + 1
    End If

    If BadGuys(x).ExplodingFrame >= 14 Then
       BadGuys(x).Exploding = 0
       flgm = 0
       score = score + 500
       stage = 0
    End If
  
Next x

End Sub




