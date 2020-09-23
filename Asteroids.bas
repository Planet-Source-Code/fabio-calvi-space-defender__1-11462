Attribute VB_Name = "asteroids"
Public Sub astero()
Dim res As Integer

Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

For x = 0 To CurLevel.NumOfBadGuys
    If BadGuys(x).Activated = 0 Then GoTo 10
    res = x Mod 2
    If res > 0 Then
       Set BadGuys(x).PicT = Form2.Picaster2
       Set BadGuys(x).mask = Form2.Picaster2m
       BadGuys(x).xsize = 37
       BadGuys(x).ysize = 32
    End If
    If res = 0 Then
       Set BadGuys(x).PicT = Form2.Picaster1
       Set BadGuys(x).mask = Form2.Picaster1m
       BadGuys(x).xsize = 32
       BadGuys(x).ysize = 37
    End If
    BadGuys(x).oldX = BadGuys(x).x
    BadGuys(x).oldY = BadGuys(x).y
    If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
       If x > 3 And numsec < 40 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 7 And numsec < 80 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 11 And numsec < 120 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 15 And numsec < 160 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 19 And numsec < 200 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 21 And numsec < 240 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       If x > 25 And numsec < 280 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 150
       
       BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
   
       If res > 0 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, BadGuys(x).xsize * BadGuys(x).frame, 0, vbMergePaint
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, BadGuys(x).xsize * BadGuys(x).frame, 0, vbSrcAnd
       End If
       If res = 0 Then
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, 0, BadGuys(x).ysize * BadGuys(x).frame, vbMergePaint
          BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, 0, BadGuys(x).ysize * BadGuys(x).frame, vbSrcAnd
       End If
         
       BitBlt Form1.Pic_radar.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad.hdc, 0, 0, vbSrcCopy
       If roll = True Then BitBlt Form1.Pic_radarv.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad2.hdc, 0, 0, vbSrcCopy
               
       BadGuys(x).frame = BadGuys(x).frame + 1
       If BadGuys(x).frame >= 29 Then BadGuys(x).frame = 0
    End If
    If BadGuys(x).Damage > CurLevel.Damagelimit Then BadGuys(x).Exploding = 1

If CollisionDetect(ShipX, ShipY, Form2.Pictm, BadGuys(x).x, BadGuys(x).y, BadGuys(x).mask, Form2.PicTemp) Then
   Exploding = 1
   BadGuys(x).Exploding = 1
End If
      
10 Next x

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
          score = score + 10
          TempCalc = TempCalc + 1
          fboom = fboom + 1
          If fboom > 3 Then fboom = 1
       End If
    End If
Next x
   
If Tempclac + TempCalc >= CurLevel.NumOfBadGuys + 1 Then
   flgm = 0
   Form1.Tmr_flgm.Interval = 1
   Form1.Tmr_flgm.Enabled = False
   stage = 0
End If
numsec = numsec + 1
End Sub

