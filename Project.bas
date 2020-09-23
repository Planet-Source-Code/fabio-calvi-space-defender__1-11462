Attribute VB_Name = "Miscellaneous"
Public MainHDC As Long
Public BufferHDC As Long
Public Const Pi As Double = 3.14159265358979
Public Const Radians As Double = (2 * Pi) / 360

'Constants for the GenerateDC function
'**LoadImage Constants**
Public Const IMAGE_BITMAP As Long = 0
Public Const LR_LOADFROMFILE As Long = &H10
Public Const LR_CREATEDIBSECTION As Long = &H2000
Public Const LR_DEFAULTSIZE As Long = &H40
'ALL API CALLS:
'Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public NUM As Integer
Public Declare Function StretchBlt Lib "gdi32" _
    (ByVal hdc As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, _
    ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Declare Function AlphaBlending Lib "Alphablending.dll" _
       (ByVal destHDC As Long, ByVal XDest As Long, ByVal YDest As Long, _
        ByVal destWidth As Long, ByVal destHeight As Long, ByVal srcHDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal srcWidth As Long, ByVal srcHeight As Long, ByVal AlphaSource As Long) As Long

Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public ResultRegion As Long

Public Type POINTAPI
   x As Long
   y As Long
End Type

Public PolyPoints(4) As POINTAPI

Public Const RGN_AND = 1
Public Const RGN_OR = 2
Public Const RGN_XOR = 3
Public Const RGN_DIFF = 4
Public Const RGN_COPY = 5
Public Const HWND_TOPMOST = -1

Public Const FADE_T_TO_B = 0
Public Const FADE_B_TO_T = 1
Public Const FADE_L_TO_R = 2
Public Const FADE_R_TO_L = 3
Public Const FADE_RANDOM = 4
Public Const FADE_OUTWARD = 5

Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXFULLSCREEN = 16

Public BackTile1 As Long
Public BackTile2 As Long
Public BackTile3 As Long
Public BackTile4 As Long
Public Const FirstTileHeight As Long = 800
Public Const SecondTileHeight As Long = 1600
Public Const ThirdTileHeight As Long = 2400
Public Const FourthTileHeight As Long = 3200
'General position variables
Public BackYPos As Long
Public OverlapBottom As Long
Public OverlapTop As Long

Public yrisoluz As Single
Public xrisoluz As Single

Public Const NumOfStars = 50
'public Const BulletSpeed = 20 '1°
'Public Const BulletSpeed = 40 '2°
Public Const BulletSpeed = 30

Public Const HAcc As Single = 2
Public Const HDel As Single = 2
Public Const VAcc As Single = 2
Public Const VDel As Single = 2
Public Const BadBulletSpeed = 5
Public Const Damagelimit = 5
Public Const KeySpeed = 10
'public Const NumOfBullets = 60 '1°
'Public Const NumOfBullets = 30 '2°
Public Const NumOfBullets = 50

Public Const ViewportHeight As Long = 600
Public Const ViewportWidth As Long = 600
Public Const BackTileHeight As Long = 600
Public Const BackTileWidth As Long = 600
'ALL TYPES:
Public Type Star
    x As Integer
    y As Integer
    bright As Byte
    SPEED As Byte
End Type

Public Type bullet
    x As Integer
    y As Integer
    Velocity As Integer
    Activated As Byte
End Type

Public Type BadBullet
    x As Integer
    y As Integer
    Velocity As Integer
    Activated As Integer
End Type

Public Type BadGuy
    PicT As Object
    mask As Object
    xsize As Integer
    ysize As Integer
    x As Integer
    y As Integer
    oldX As Integer
    oldY As Integer
    Exploding As Integer
    ExplodingFrame As Double
    frame As Long
    Activated As Integer
    DstX As Integer
    DstY As Integer
    Damage As Integer
    Velocity As Integer
    Bulletl(0 To 10) As BadBullet
    Bulletc(0 To 10) As BadBullet
    Bulletr(0 To 10) As BadBullet
    Bulletla(0 To 10) As Double
    Bulletra(0 To 10) As Double
    Bulletca(0 To 10) As Double
    BulletsActivated As Byte
    bulletlxpos As Integer
    bulletlypos As Integer
    bulletrxpos As Integer
    bulletrypos As Integer
    bulletcxpos As Integer
    bulletcypos As Integer
    Firing As Byte
  
End Type

Public Type Level
    NumOfBadGuys As Integer
    Damage As Integer
    Damagelimit As Integer
    Velocity As Integer
    OddsOfFiring As Integer
    BulletSpeed As Integer
End Type

Public Type PointXY
    x As Integer
    y As Integer
End Type

Public Type RGB
    R As Integer
    G As Integer
    b As Integer
End Type

'ALL public VARIABLES:
Public Health As Single
Public BufferWidth As Long
Public BufferHeight As Long

Public Sub UpdateHealth()
BitBlt Form1.PicHealth.hdc, 0, 0, Form1.PicHealth.ScaleWidth, ((100 - Health) / 100) * Form1.PicHealth.ScaleHeight, Form1.Picture3.hdc, 0, 0, vbSrcCopy
Form1.PicHealth.Refresh
End Sub

Public Sub DrawHealthBar()
'calculation variables for r,g,b gradiency
Dim vR, VG, VB As Single
'colors of the picture boxes
Dim Color1, Color2 As Long
'r,g,b variables for each picture box
Dim R, G, b, R2, G2, b2 As Integer
'calculation variable for extracting the rgb values
Dim temp As Long

Color1 = RGB(0, 255, 0)
Color2 = RGB(255, 0, 0)

'extract the r,g,b values from the first picture box
temp = (Color1 And 255)
R = temp And 255
temp = Int(Color1 / 256)
G = temp And 255
temp = Int(Color1 / 65536)
b = temp And 255
temp = (Color2 And 255)
R2 = temp And 255
temp = Int(Color2 / 256)
G2 = temp And 255
temp = Int(Color2 / 65536)
b2 = temp And 255

'create a calculation variable for determining the step between
'each level of the gradient; this also allows the user to create
'a perfect gradient regardless of the form size
vR = Abs(R - R2) / Form1.PicHealth.ScaleHeight
VG = Abs(G - G2) / Form1.PicHealth.ScaleHeight
VB = Abs(b - b2) / Form1.PicHealth.ScaleHeight
'if the second value is lower then the first value, make the step
'negative
If R2 < R Then vR = -vR
If G2 < G Then VG = -VG
If b2 < b Then VB = -VB
'run a loop through the form height, incrementing the gradient color
'according to the height of the line being drawn
For y = 0 To Form1.PicHealth.ScaleHeight
R2 = R + vR * y
G2 = G + VG * y
b2 = b + VB * y
'draw the line and continue through the loop
Form1.PicHealth.Line (0, y)-(Form1.PicHealth.ScaleWidth, y), RGB(R2, G2, b2)
Next y

End Sub


