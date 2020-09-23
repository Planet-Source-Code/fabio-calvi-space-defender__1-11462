Attribute VB_Name = "Enemy12"
Public Sub badguy12()

Form1.Pic_radar.Cls
Form1.Pic_radarv.Cls

For x = 0 To CurLevel.NumOfBadGuys
If BadGuys(x).Activated = 0 Then GoTo 10
    Set BadGuys(x).PicT = Form2.PicI
    Set BadGuys(x).mask = Form2.PicIm
    BadGuys(x).xsize = 59
    BadGuys(x).ysize = 50
    BadGuys(x).Velocity = 35
    BadGuys(x).oldX = BadGuys(x).x
    BadGuys(x).oldY = BadGuys(x).y
    If BadGuys(x).Activated = 1 And BadGuys(x).Exploding = 0 Then
       If x = 1 And numsec < 5 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 2 And numsec < 10 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 3 And numsec < 15 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 4 And numsec < 20 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 5 And numsec < 25 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 6 And numsec < 30 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 7 And numsec < 35 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 8 And numsec < 40 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 9 And numsec < 45 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 10 And numsec < 50 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 11 And numsec < 55 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 12 And numsec < 60 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 13 And numsec < 65 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 14 And numsec < 70 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 15 And numsec < 75 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 16 And numsec < 80 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 17 And numsec < 85 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 18 And numsec < 90 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 19 And numsec < 95 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 20 And numsec < 100 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 21 And numsec < 105 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 22 And numsec < 110 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 23 And numsec < 115 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 24 And numsec < 120 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 25 And numsec < 125 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 25 And numsec < 130 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 26 And numsec < 135 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 27 And numsec < 140 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 28 And numsec < 145 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 29 And numsec < 150 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       If x = 30 And numsec < 155 Then BadGuys(x).y = Form1.PicMain.ScaleTop - 200
       
       BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
       
       BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).mask.hdc, 0, 0, vbMergePaint
       BitBlt Form1.PicScreenBuffer.hdc, BadGuys(x).x, BadGuys(x).y, BadGuys(x).xsize, BadGuys(x).ysize, BadGuys(x).PicT.hdc, 0, 0, vbSrcAnd
       BitBlt Form1.Pic_radar.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad.hdc, 0, 0, vbSrcCopy
       If roll = True Then BitBlt Form1.Pic_radarv.hdc, BadGuys(x).x / (diffwidth - 0.7), (BadGuys(x).y + 100) / diffheight, 6, 6, Form1.rad2.hdc, 0, 0, vbSrcCopy
    
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
          score = score + 50
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


