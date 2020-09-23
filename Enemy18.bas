Attribute VB_Name = "Enemy18"
Public Sub badguy18()

Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

For x = 0 To CurLevel.NumOfBadGuys
   If BadGuys(x).Activated = 0 Then GoTo 10
   If x = 20 Then
      Set BadGuys(x).PicT = Form2.PicR
      Set BadGuys(x).mask = Form2.PicRm
      BadGuys(x).bulletlxpos = 7
      BadGuys(x).bulletlypos = 54
      BadGuys(x).bulletrxpos = 67
      BadGuys(x).bulletrypos = 54
      BadGuys(x).bulletcxpos = 34
      BadGuys(x).bulletcypos = 77
      BadGuys(x).xsize = 76
      BadGuys(x).ysize = 79
   End If
   If x < 20 Then
      Set BadGuys(x).PicT = Form2.PicZ
      Set BadGuys(x).mask = Form2.PicZm
      BadGuys(x).bulletlxpos = 21
      BadGuys(x).bulletlypos = 35
      BadGuys(x).bulletrxpos = 52
      BadGuys(x).bulletrypos = 35
      BadGuys(x).xsize = 74
      BadGuys(x).ysize = 50
   End If
   
   If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
      If x > 1 And numsec < 50 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 3 And numsec < 100 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 5 And numsec < 150 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 7 And numsec < 200 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 9 And numsec < 250 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 11 And numsec < 300 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 13 And numsec < 350 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 15 And numsec < 400 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x > 17 And numsec < 450 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If x = 20 And numsec < 500 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
      If numsec = 600 Then flgm = 0
      
      If x < 20 Then
         If BadGuys(x).y > (Form1.PicMain.ScaleTop - 20) And flgm = 0 Then
            flgm = 1
            Form1.Tmr_flgm.Enabled = True
         End If
         If flgm = 1 Or flgm = 2 Then
            If BadGuys(x).DstX < BadGuys(x).x Then
               BadGuys(x).x = BadGuys(x).x - BadGuys(x).Velocity
            Else
               BadGuys(x).x = BadGuys(x).x + BadGuys(x).Velocity
            End If
            BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
         End If
         If flgm = 0 Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
         If Abs(BadGuys(x).x - BadGuys(x).DstX) < CurLevel.Velocity + 1 Then BadGuys(x).DstX = Rnd * ((Form1.PicMain.ScaleWidth - 70) - Form1.PicMain.ScaleLeft)
         If Abs(BadGuys(x).y - BadGuys(x).DstY) < CurLevel.Velocity + 1 Then BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
      End If
      
      If x = 20 Then
         If BadGuys(x).DstX < BadGuys(x).x Then
            BadGuys(x).x = (BadGuys(x).x - BadGuys(x).Velocity) + 2
         Else
            BadGuys(x).x = BadGuys(x).x + BadGuys(x).Velocity
         End If
         If BadGuys(x).y >= (Form1.PicMain.ScaleTop - 20) And flgm = 0 Then
            flgm = 1
            Form1.Tmr_flgm.Interval = 1
            Form1.Tmr_flgm.Enabled = True
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
         If Form1.Tmr_flgm.Interval > 80 Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
         If Abs(BadGuys(x).x - BadGuys(x).DstX) < CurLevel.Velocity + 1 Then BadGuys(x).DstX = Rnd * ((Form1.PicMain.ScaleWidth - 70) - Form1.PicMain.ScaleLeft)
         If Abs(BadGuys(x).y - BadGuys(x).DstY) < CurLevel.Velocity + 1 Then BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
      End If
      
      BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, 0, 0, vbMergePaint
      BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, 0, 0, vbSrcAnd
      BitBlt Form1.Pic_radar.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad.hdc, 0, 0, vbSrcCopy
      If roll = True Then BitBlt Form1.Pic_radarv.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad2.hdc, 0, 0, vbSrcCopy
   
   End If
 
   If x = 20 And BadGuys(x).Damage > 200 Then BadGuys(x).Exploding = 1
   If x < 20 And BadGuys(x).Damage > CurLevel.Damagelimit Then BadGuys(x).Exploding = 1
      
   If BadGuys(x).Activated = 1 Then
      BadGuys(x).Firing = Int(Rnd * CurLevel.OddsOfFiring)
   Else
     BadGuys(x).Firing = 0
   End If

'Firing bullets
If BadGuys(x).y > Form1.PicMain.ScaleTop + 20 Then
For y = 0 To 0
If BadGuys(x).Activated > 0 And BadGuys(x).Firing = 1 Then
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
If x = 20 Then
If BadGuys(x).Bulletc(y).Activated = 0 Then
   BadGuys(x).Bulletc(y).Activated = 1
   BadGuys(x).Bulletc(y).x = BadGuys(x).x + BadGuys(x).bulletcxpos
   BadGuys(x).Bulletc(y).y = BadGuys(x).y + BadGuys(x).bulletcypos
End If
End If
BadGuys(x).BulletsActivated = BadGuys(x).BulletsActivated + 1
End If
Next y
End If

If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(x).x, BadGuys(x).y, BadGuys(x).mask, Form2.PicTemp) Then
   Exploding = 1
   BadGuys(x).Exploding = 1
End If
      
10 Next x

For x = 0 To CurLevel.NumOfBadGuys
For y = 0 To 0
If x < 20 Then
If BadGuys(x).Bulletl(y).Activated = 1 Then
bullets = True
BadGuys(x).Bulletl(y).y = BadGuys(x).Bulletl(y).y + CurLevel.BulletSpeed
If BadGuys(x).Bulletl(y).y > Form1.PicMain.ScaleHeight Then
   BadGuys(x).Bulletl(y).Activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 6, 6, Form2.PicBBulletM.hdc, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 6, 6, Form2.PicGBullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletl(y).x + Form2.PicCBullet.ScaleWidth > ShipX And BadGuys(x).Bulletl(y).x < ShipX + Form2.Picture1.ScaleWidth And _
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
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 6, 6, Form2.PicGBullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletr(y).x + Form2.PicCBullet.ScaleWidth > ShipX And BadGuys(x).Bulletr(y).x < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(x).Bulletr(y).y - 25) - ShipY) < 18 Then
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 10
   UpdateHealth
   BadGuys(x).Bulletr(y).Activated = 0
   bullets = False
End If
End If
End If

If x = 20 Then
If BadGuys(x).Bulletl(y).Activated = 1 Then
bullets = True
BadGuys(x).Bulletl(y).x = BadGuys(x).Bulletl(y).x - Cos(10 * Radians)
BadGuys(x).Bulletl(y).y = BadGuys(x).Bulletl(y).y + CurLevel.BulletSpeed
If BadGuys(x).Bulletl(y).y > Form1.PicMain.ScaleHeight Then
   BadGuys(x).Bulletl(y).Activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletl(y).x, BadGuys(x).Bulletl(y).y, 8, 8, Form2.PicABullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletl(y).x + Form2.PicCBullet.ScaleWidth > ShipX And BadGuys(x).Bulletl(y).x < ShipX + Form2.Picture1.ScaleWidth And _
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
BadGuys(x).Bulletr(y).x = BadGuys(x).Bulletr(y).x + Cos(10 * Radians)
BadGuys(x).Bulletr(y).y = BadGuys(x).Bulletr(y).y + CurLevel.BulletSpeed
If BadGuys(x).Bulletr(y).y > Form1.PicMain.ScaleHeight Then
   BadGuys(x).Bulletr(y).Activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 8, 8, Form2.PicABulletM.hdc, 0, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletr(y).x, BadGuys(x).Bulletr(y).y, 8, 8, Form2.PicABullet.hdc, 0, 0, vbSrcAnd
If BadGuys(x).Bulletr(y).x + Form2.PicCBullet.ScaleWidth > ShipX And BadGuys(x).Bulletr(y).x < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(x).Bulletr(y).y - 25) - ShipY) < 18 Then
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 10
   UpdateHealth
   BadGuys(x).Bulletr(y).Activated = 0
   bullets = False
End If
End If
If BadGuys(x).Bulletc(y).Activated = 1 Then
   BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 6, Form1.PicMain.ScaleHeight - BadGuys(x).Bulletc(y).y, Form2.Piclaser1m.hdc, 0, 0, vbMergePaint
   BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 6, Form1.PicMain.ScaleHeight - BadGuys(x).Bulletc(y).y, Form2.Piclaser1.hdc, 0, 0, vbSrcAnd
   If BadGuys(x).Bulletc(y).x > ShipX And BadGuys(x).Bulletc(y).x < ShipX + Form2.Picture1.ScaleWidth And _
      ShipY > Form1.PicMain.ScaleTop + BadGuys(x).Bulletc(y).y Then
      Exploding = 1
   End If
   BadGuys(x).Bulletc(y).Activated = 0
End If
End If
Next y

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
          score = score + 50
          If x = 20 Then score = score + 50
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
numsec = numsec + 1

End Sub

