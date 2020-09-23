Attribute VB_Name = "collide"
Option Explicit
Dim R As Integer, c As Integer

Dim hNewBMP As Long
Dim hPrevBMP As Long
Dim tmpObj As Long

Dim hMemDC As Long
Dim nRet As Long

Dim blnCollision As Boolean
            
Dim iMaskWidth As Integer
Dim iMaskHeight As Integer

Dim iM1SrcX As Integer
Dim iM1SrcY As Integer

Dim iM2SrcX As Integer
Dim iM2SrcY As Integer

Dim iDestX As Integer
Dim iDestY As Integer

Dim iStartBlankWidth As Integer
Dim iStartBlankHeight As Integer

Function CollisionDetect(ByVal X1 As Integer, ByVal Y1 As Integer, picMask As PictureBox, ByVal X2 As Integer, ByVal Y2 As Integer, picMask1 As PictureBox, picBlank As PictureBox) As Boolean
    
    If X1 <= X2 Then
        iMaskWidth = X1 + picMask.ScaleWidth - X2
        iM1SrcX = picMask.ScaleWidth - iMaskWidth
        iM2SrcX = 0
        iDestX = 0
        iStartBlankWidth = iMaskWidth
    Else
        iMaskWidth = X2 + picMask.ScaleWidth - X1
        iM1SrcX = 0
        iM2SrcX = picMask.ScaleWidth - iMaskWidth
        iDestX = 0
        iStartBlankWidth = iMaskWidth
    End If
    
    If Y1 <= Y2 Then
        iMaskHeight = Y1 + picMask.ScaleHeight - Y2
        iM1SrcY = picMask.ScaleHeight - iMaskHeight
        iM2SrcY = 0
        iDestX = 0
        iStartBlankHeight = iMaskHeight
    Else
        iMaskHeight = Y2 + picMask.ScaleHeight - Y1
        iM1SrcY = 0
        iM2SrcY = picMask.ScaleHeight - iMaskHeight
        iDestX = 0
        iStartBlankHeight = iMaskHeight
    End If
    
   
    blnCollision = False
    For c = 0 To iMaskHeight - 1
        For R = 0 To iMaskWidth - 1
            If GetPixel(Form1.PicScreenBuffer.hdc, R, c) <> 16777215 Then
                blnCollision = True
                Exit For
            Else
            End If
        Next
        
        If blnCollision = True Then
            Exit For
        End If
        
    Next

'------------------------------------------------------------

    CollisionDetect = blnCollision
    
End Function


