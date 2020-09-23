Attribute VB_Name = "Enemy16"
Public Sub badguy16()
Dim res As Integer
Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

For x = 0 To CurLevel.NumOfBadGuys
    If BadGuys(x).Activated = 0 Then GoTo 10
    res = x Mod 2
    If res > 0 Then
       Set BadGuys(x).PicT = Form2.PicMines1
       Set BadGuys(x).mask = Form2.PicMines1m
    End If
    If res = 0 Then
       Set BadGuys(x).PicT = Form2.PicMines2
       Set BadGuys(x).mask = Form2.PicMines2m
    End If
    BadGuys(x).xsize = 14
    BadGuys(x).ysize = 13
    
    If x = 30 Then
       Set BadGuys(x).PicT = Form2.PicQ
       Set BadGuys(x).mask = Form2.PicQm
       BadGuys(x).xsize = 68
       BadGuys(x).ysize = 69
    End If
    If x = 40 Or x = 50 Or x = 60 Then
       Set BadGuys(x).PicT = Form2.PicV
       Set BadGuys(x).mask = Form2.PicVm
       BadGuys(x).bulletlxpos = 22
       BadGuys(x).bulletlypos = 34
       BadGuys(x).bulletrxpos = 53
       BadGuys(x).bulletrypos = 34
       BadGuys(x).xsize = 74
       BadGuys(x).ysize = 51
    End If

       BadGuys(x).oldX = BadGuys(x).x
       BadGuys(x).oldY = BadGuys(x).y
       If x > 3 And numsec < 20 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 6 And numsec < 30 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 10 And numsec < 40 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 15 And numsec < 60 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 20 And numsec < 70 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 23 And numsec < 90 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x = 30 And numsec < 200 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 31 And numsec < 120 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 36 And numsec < 180 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 40 And numsec < 200 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 45 And numsec < 210 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 50 And numsec < 220 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 53 And numsec < 240 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x = 40 And numsec < 90 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x = 50 And numsec < 120 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x = 60 And numsec < 160 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       
       If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
       
       If x <> 40 And x <> 50 And x <> 60 Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
       
       If x = 30 Then BadGuys(x).y = BadGuys(x).y + 3
       
       If x = 40 Or x = 50 Or x = 60 Then
          If BadGuys(x).DstX < BadGuys(x).x Then
             BadGuys(x).x = BadGuys(x).x - BadGuys(x).Velocity + 2
          Else
             BadGuys(x).x = BadGuys(x).x + BadGuys(x).Velocity
          End If
          If BadGuys(x).y >= (Form1.PicMain.ScaleTop - 20) And flgm = 0 Then
             flgm = 1
             Form1.Tmr_flgm.Enabled = True
             BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
          End If
          If flgm = 1 Then
             If BadGuys(x).DstY < BadGuys(x).y Then
                BadGuys(x).y = BadGuys(x).y - BadGuys(x).Velocity + 2
             Else
                BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
             End If
          End If
 
          If flgm = 0 Or flgm = 2 Then BadGuys(x).y = BadGuys(x).y + 5
          If Abs(BadGuys(x).x - BadGuys(x).DstX) < CurLevel.Velocity + 1 Then BadGuys(x).DstX = Rnd * ((Form1.PicMain.ScaleWidth - 70) - Form1.PicMain.ScaleLeft)
          If Abs(BadGuys(x).y - BadGuys(x).DstY) < CurLevel.Velocity + 1 Then BadGuys(x).DstY = Rnd * (Form1.PicMain.ScaleHeight - 200)
    
       End If
       
       If x = 30 Or x = 40 Or x = 50 Or x = 60 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, 0, 0, vbMergePaint
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, 0, 0, vbSrcAnd
          BitBlt Form1.Pic_radar.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad.hdc, 0, 0, vbSrcCopy
          If roll = True Then BitBlt Form1.Pic_radarv.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad2.hdc, 0, 0, vbSrcCopy
       Else
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, 14 * frame, 0, vbMergePaint
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, 14 * frame, 0, vbSrcAnd
          BitBlt Form1.Pic_radar.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 4, 4, Form1.rad3.hdc, 0, 0, vbSrcCopy
          frame = frame + 1
          If frame > 5 Then frame = 0
       End If
               
    End If
   
   If (x = 30 Or x = 40 Or x = 50 Or x = 60) And BadGuys(x).Damage > 90 Then BadGuys(x).Exploding = 1
   If (x <> 30 And x <> 40 And x <> 50 And x <> 60) And BadGuys(x).Damage > CurLevel.Damagelimit Then BadGuys(x).Exploding = 1


If x = 40 Or x = 50 Or x = 60 Then
If BadGuys(x).y > Form1.PicMain.ScaleTop + 20 Then
   If BadGuys(x).Activated = 1 Then
      BadGuys(x).Firing = 1
   Else
      BadGuys(x).Firing = 0
   End If
End If
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
BadGuys(x).BulletsActivated = BadGuys(x).BulletsActivated + 1
End If
End If
Next y
End If
If x = 30 Or x = 40 Or x = 50 Or x = 60 Then
   If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(x).x, BadGuys(x).y, BadGuys(x).mask, Form2.PicTemp) Then
      Exploding = 1
      BadGuys(x).Exploding = 1
   End If
End If
If x <> 30 And x <> 40 And x <> 50 And x <> 60 Then
   If BadGuys(x).x + 14 > ShipX And BadGuys(x).x < ShipX + Form2.Picture1.ScaleWidth And _
      Abs((BadGuys(x).y - 25) - ShipY) < 20 Then
          Form2.Picture1.Picture = Form2.PicFlash.Picture
          Health = Health - 5
          UpdateHealth
          BadGuys(x).Exploding = 1
   End If
End If
10 Next x

'***********************************
For x = 0 To CurLevel.NumOfBadGuys
For y = 0 To 1
If BadGuys(x).Bulletl(y).Activated = 1 Then
bullets = True
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
Next y
'If comment BadGuys shoot once only
BadGuys(x).BulletsActivated = 0
Next x

'***********************************

For x = 0 To CurLevel.NumOfBadGuys
    If BadGuys(x).Activated = 1 And BadGuys(x).y > Form1.PicMain.ScaleHeight Then
       Tempclac = Tempclac + 1
       BadGuys(x).Activated = 0
    End If
    If BadGuys(x).Exploding = 1 Then
       BadGuys(x).Activated = 0
       If (x = 30 Or x = 40 Or x = 50 Or x = 60) Then
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
       If (x <> 30 And x <> 40 And x <> 50 And x <> 60) Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 30, 30, Form2.PicExplode3m.hdc, 30 * BadGuys(x).ExplodingFrame, 0, vbPatInvert
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, 30, 30, Form2.PicExplode3.hdc, 30 * BadGuys(x).ExplodingFrame, 0, vbSrcPaint
          BadGuys(x).ExplodingFrame = BadGuys(x).ExplodingFrame + 1
          If BadGuys(x).ExplodingFrame = 8 Then
             BadGuys(x).Exploding = 0
             BadGuys(x).Activated = 0
             score = score + 10
             TempCalc = TempCalc + 1
          End If
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
