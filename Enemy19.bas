Attribute VB_Name = "Enemy19"
Public Sub badguy19()

Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

For x = 0 To CurLevel.NumOfBadGuys
If BadGuys(x).Activated = 0 Then GoTo 10
    Set BadGuys(x).PicT = Form2.PicP
    Set BadGuys(x).mask = Form2.PicPm
    BadGuys(x).bulletcxpos = 36
    BadGuys(x).bulletcypos = 26
    BadGuys(x).xsize = 70
    BadGuys(x).ysize = 58
   
    BadGuys(x).oldX = BadGuys(x).x
    BadGuys(x).oldY = BadGuys(x).y
    If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
      If x > 3 And numsec < 40 Then
         BadGuys(x).y = Form1.PicMain.ScaleTop - 150
         BadGuys(x).x = BadGuys(x).x - 0.5
      End If
      If x > 7 And numsec < 80 Then
         BadGuys(x).y = Form1.PicMain.ScaleTop - 150
         BadGuys(x).x = BadGuys(x).x - 0.5
      End If
      If x > 11 And numsec < 120 Then
         BadGuys(x).y = Form1.PicMain.ScaleTop - 150
         BadGuys(x).x = BadGuys(x).x - 0.75
      End If
      If x > 15 And numsec < 160 Then
         BadGuys(x).y = Form1.PicMain.ScaleTop - 150
         BadGuys(x).x = BadGuys(x).x - 1
      End If
      If x > 19 And numsec < 200 Then
         BadGuys(x).y = Form1.PicMain.ScaleTop - 150
         BadGuys(x).x = BadGuys(x).x - 1
      End If
      
      BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
      BadGuys(x).x = BadGuys(x).x + Sin(BadGuys(x).y / 25) * 5
    
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
            If BadGuys(x).Bulletc(y).Activated = 0 Then
               BadGuys(x).Bulletc(y).Activated = 1
               BadGuys(x).Bulletc(y).x = BadGuys(x).x + BadGuys(x).bulletcxpos
               BadGuys(x).Bulletc(y).y = BadGuys(x).y + BadGuys(x).bulletcypos
            End If
            BadGuys(x).BulletsActivated = BadGuys(x).BulletsActivated + 1
         End If
      End If

      Next y
   End If

If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(x).x, BadGuys(x).y, BadGuys(x).mask, Form2.PicTemp) Then
   Exploding = 1
   BadGuys(x).Exploding = 1
End If

10 Next x

'**********************************************************

For x = 0 To CurLevel.NumOfBadGuys
For y = 0 To 0
If BadGuys(x).Bulletc(y).Activated = 1 Then
bullets = True
BadGuys(x).Bulletc(y).y = BadGuys(x).Bulletc(y).y + CurLevel.BulletSpeed
If BadGuys(x).Bulletc(y).y > Form1.PicMain.ScaleHeight Then
   BadGuys(x).Bulletc(y).Activated = 0
   bullets = False
End If
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 44, Form2.PicFBulletm.hdc, 8 * frame, 0, vbMergePaint
BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).Bulletc(y).x, BadGuys(x).Bulletc(y).y, 8, 44, Form2.PicFBullet.hdc, 8 * frame, 0, vbSrcAnd
frame = frame + 1
If frame > 7 Then frame = 0
If BadGuys(x).Bulletc(y).x + 8 > ShipX And BadGuys(x).Bulletc(y).x < ShipX + Form2.Picture1.ScaleWidth And _
   Abs((BadGuys(x).Bulletc(y).y - 25) - ShipY) < 18 Then
   Form2.Picture1.Picture = Form2.PicFlash.Picture
   Health = Health - 10
   UpdateHealth
   BadGuys(x).Bulletc(y).Activated = 0
   bullets = False
End If
End If
Next y
'If comment BadGuys shoot once only
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
          TempCalc = TempCalc + 1
          fboom = fboom + 1
          If fboom > 3 Then fboom = 1
       End If
    End If
Next x

If Tempclac + TempCalc >= CurLevel.NumOfBadGuys + 1 And bullets = False Then
   flgm = 0
   flgr = flgr + 1
   Form1.Tmr_flgm.Interval = 1
   Form1.Tmr_flgm.Enabled = False
   If flgr >= 2 Then
      flgr = 0
      stage = 0
    Else
      BuildBadGuys
    End If
End If
numsec = numsec + 1
End Sub

