VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form2"
   ScaleHeight     =   8235
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox PicBoss2A 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1155
      Left            =   8160
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   104
      Top             =   4200
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox PicBoss2Am 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1155
      Left            =   8280
      Picture         =   "Form2.frx":4146
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   103
      Top             =   4200
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox PicBoss2B 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1125
      Left            =   8160
      Picture         =   "Form2.frx":44FC
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   101
      Top             =   3120
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox PicBoss2C 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1125
      Left            =   8160
      Picture         =   "Form2.frx":8B22
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   99
      Top             =   2040
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox PicHBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1920
      Picture         =   "Form2.frx":DC60
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   98
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox PicGBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1800
      Picture         =   "Form2.frx":DD1A
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   97
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox PicExplode3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   480
      Picture         =   "Form2.frx":DDD4
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   95
      Top             =   5520
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.PictureBox PicExplode3m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   510
      Left            =   480
      Picture         =   "Form2.frx":13276
      ScaleHeight     =   30
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   96
      Top             =   5760
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.PictureBox Picshotm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   1200
      Picture         =   "Form2.frx":13680
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   93
      Top             =   1920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Picshot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   1200
      Picture         =   "Form2.frx":136FA
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   21
      TabIndex        =   92
      Top             =   1680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicFBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   8040
      Picture         =   "Form2.frx":13A3C
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   90
      Top             =   720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicMines2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   6600
      Picture         =   "Form2.frx":15B7E
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   88
      Top             =   720
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox PicMines1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   5040
      Picture         =   "Form2.frx":1688C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   86
      Top             =   720
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox PicEBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   1800
      Picture         =   "Form2.frx":1759A
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   85
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox PicK 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Left            =   6840
      Picture         =   "Form2.frx":176EC
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   83
      Top             =   3360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicZ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   6840
      Picture         =   "Form2.frx":1A3CE
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   81
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicV 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Left            =   6840
      Picture         =   "Form2.frx":1CFD0
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   79
      Top             =   1680
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicDBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   1560
      Picture         =   "Form2.frx":1FCB2
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   78
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox PicFlash 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   720
      Left            =   4200
      Picture         =   "Form2.frx":1FE04
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   77
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox PicMissile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1200
      Left            =   9240
      Picture         =   "Form2.frx":21916
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   75
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PicU 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   885
      Left            =   5520
      Picture         =   "Form2.frx":24EC8
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   73
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox PicS 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   120
      Picture         =   "Form2.frx":27A02
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   71
      Top             =   1680
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicR 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1245
      Left            =   5400
      Picture         =   "Form2.frx":29A94
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   69
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicQ 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   4320
      Picture         =   "Form2.frx":2E132
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   67
      Top             =   2640
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox PicP 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   930
      Left            =   4200
      Picture         =   "Form2.frx":31870
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   65
      Top             =   1680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox Piclaser2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   11160
      Picture         =   "Form2.frx":348BA
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox PicBoss1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1605
      Left            =   4560
      Picture         =   "Form2.frx":38A9C
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   61
      Top             =   3840
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.PictureBox PicO 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   120
      Picture         =   "Form2.frx":4120E
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   59
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picaster2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5040
      Picture         =   "Form2.frx":43B50
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1076
      TabIndex        =   57
      Top             =   120
      Visible         =   0   'False
      Width           =   16200
   End
   Begin VB.PictureBox Picaster1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   16200
      Left            =   10320
      Picture         =   "Form2.frx":5CF12
      ScaleHeight     =   1076
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   55
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox PicI 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   3120
      Picture         =   "Form2.frx":762D4
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   54
      Top             =   2640
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox PicIm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   3360
      Picture         =   "Form2.frx":7863E
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   59
      TabIndex        =   53
      Top             =   2640
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox PicCBulletm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   1440
      Picture         =   "Form2.frx":78818
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   52
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox PicCBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   315
      Left            =   1680
      Picture         =   "Form2.frx":7896A
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   51
      Top             =   1200
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox PicA4Bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   600
      Picture         =   "Form2.frx":78ABC
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   50
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicA3Bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1080
      Picture         =   "Form2.frx":78BBE
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   49
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   3120
      Picture         =   "Form2.frx":78CC0
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   47
      Top             =   3480
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox PicNm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   3360
      Picture         =   "Form2.frx":7AE22
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   48
      Top             =   3480
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox PicL 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   3000
      Picture         =   "Form2.frx":7B04C
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   45
      Top             =   4200
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox PicLm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   585
      Left            =   3120
      Picture         =   "Form2.frx":7D302
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   46
      Top             =   4200
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox Piclaser1m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   11640
      Picture         =   "Form2.frx":7D4F0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Piclaser1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   11520
      Picture         =   "Form2.frx":80412
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox PicA2Bullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   840
      Picture         =   "Form2.frx":83334
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   42
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   3000
      Picture         =   "Form2.frx":83436
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   40
      Top             =   4800
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox PicMm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   3240
      Picture         =   "Form2.frx":85598
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   41
      Top             =   4800
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox PicH 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Left            =   1440
      Picture         =   "Form2.frx":857C2
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   38
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicHm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   765
      Left            =   1680
      Picture         =   "Form2.frx":88704
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicG 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   1680
      Picture         =   "Form2.frx":88982
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   36
      Top             =   3840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox PicGm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   1920
      Picture         =   "Form2.frx":8B3F4
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   37
      Top             =   3840
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox PicF 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   2760
      Picture         =   "Form2.frx":8B696
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicE 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1170
      Left            =   1680
      Picture         =   "Form2.frx":8ED2C
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   30
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicD 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   1680
      Picture         =   "Form2.frx":92F56
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox PicC 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   120
      Picture         =   "Form2.frx":951E8
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicExplode1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   -480
      Picture         =   "Form2.frx":997A2
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   780
      TabIndex        =   34
      Top             =   5520
      Visible         =   0   'False
      Width           =   11760
   End
   Begin VB.PictureBox PicExplode1m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   -480
      Picture         =   "Form2.frx":BBC54
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   780
      TabIndex        =   35
      Top             =   5880
      Visible         =   0   'False
      Width           =   11760
   End
   Begin VB.PictureBox PicFm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   3000
      Picture         =   "Form2.frx":BD40E
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   33
      Top             =   1680
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicEm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1170
      Left            =   1920
      Picture         =   "Form2.frx":BD734
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   31
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicDm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   975
      Left            =   1920
      Picture         =   "Form2.frx":BDAF6
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   29
      Top             =   1680
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox PicTemp 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   780
      Left            =   2160
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   27
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox PicExplode 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   0
      Picture         =   "Form2.frx":BDD28
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1067
      TabIndex        =   24
      Top             =   6120
      Visible         =   0   'False
      Width           =   16065
   End
   Begin VB.PictureBox PicExplodeM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1020
      Left            =   0
      Picture         =   "Form2.frx":CEC6A
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1067
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
      Width           =   16065
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   2040
      Picture         =   "Form2.frx":DFBAE
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox PicBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   120
      Picture         =   "Form2.frx":E16C0
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   22
      Top             =   960
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox PicBBullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1680
      Picture         =   "Form2.frx":E1722
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   21
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox PicBBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   150
      Left            =   1560
      Picture         =   "Form2.frx":E17DC
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox Pic2r 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   1080
      Picture         =   "Form2.frx":E18AE
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Pic2rm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   1200
      Picture         =   "Form2.frx":E2FA0
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   43
      TabIndex        =   18
      Top             =   120
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox Pic1l 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   3120
      Picture         =   "Form2.frx":E314A
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Pic1lm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   735
      Left            =   3240
      Picture         =   "Form2.frx":E4ADC
      ScaleHeight     =   45
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.PictureBox Pic2l 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   4200
      Picture         =   "Form2.frx":E6CDE
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox Pic2lm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   4320
      Picture         =   "Form2.frx":E81C0
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox PicT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   3120
      Picture         =   "Form2.frx":E836A
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox Pictm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   3240
      Picture         =   "Form2.frx":E9E7C
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   12
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox PicHIT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   120
      Picture         =   "Form2.frx":EC1CE
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox PicABullet 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      Picture         =   "Form2.frx":ECDE0
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicABulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1320
      Picture         =   "Form2.frx":ECEE2
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   8
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicCM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1230
      Left            =   360
      Picture         =   "Form2.frx":ECF4C
      ScaleHeight     =   78
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   2160
      Picture         =   "Form2.frx":ED33E
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   51
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.PictureBox PicBulletM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      Picture         =   "Form2.frx":EF690
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox PicB 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   0
      Picture         =   "Form2.frx":EF6F2
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicBM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   960
      Left            =   120
      Picture         =   "Form2.frx":F2CA4
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox Pic1R 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   705
      Left            =   0
      Picture         =   "Form2.frx":F2FBE
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox Pic1RM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   705
      Left            =   120
      Picture         =   "Form2.frx":F48DC
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.PictureBox PicHITM 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   240
      Left            =   120
      Picture         =   "Form2.frx":F4A7E
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox Picaster1m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   16200
      Left            =   10440
      Picture         =   "Form2.frx":F4B58
      ScaleHeight     =   1076
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   56
      Top             =   120
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox Picaster2m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   5040
      Picture         =   "Form2.frx":F5C72
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1076
      TabIndex        =   58
      Top             =   0
      Visible         =   0   'False
      Width           =   16200
   End
   Begin VB.PictureBox PicOm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   675
      Left            =   240
      Picture         =   "Form2.frx":F6DBC
      ScaleHeight     =   41
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   60
      Top             =   2400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PicBoss1m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1605
      Left            =   4680
      Picture         =   "Form2.frx":F6FF2
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   111
      TabIndex        =   62
      Top             =   3840
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.PictureBox Piclaser2m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   9060
      Left            =   11280
      Picture         =   "Form2.frx":F76AC
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   64
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.PictureBox PicPm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   930
      Left            =   4320
      Picture         =   "Form2.frx":F8056
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   70
      TabIndex        =   66
      Top             =   1680
      Visible         =   0   'False
      Width           =   1110
   End
   Begin VB.PictureBox PicQm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1095
      Left            =   4440
      Picture         =   "Form2.frx":F8358
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   68
      TabIndex        =   68
      Top             =   2640
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox PicRm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1245
      Left            =   5520
      Picture         =   "Form2.frx":F86DE
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   76
      TabIndex        =   70
      Top             =   1560
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox PicSm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   240
      Picture         =   "Form2.frx":F8ADC
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   62
      TabIndex        =   72
      Top             =   1680
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.PictureBox PicUm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   885
      Left            =   5640
      Picture         =   "Form2.frx":F8C86
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   74
      Top             =   2880
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox PicMissilem 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1200
      Left            =   9360
      Picture         =   "Form2.frx":F8F64
      ScaleHeight     =   76
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   76
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox PicVm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Left            =   6960
      Picture         =   "Form2.frx":F920E
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   80
      Top             =   1680
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicZm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   6960
      Picture         =   "Form2.frx":F94BC
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   82
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicKm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   825
      Left            =   6960
      Picture         =   "Form2.frx":F975E
      ScaleHeight     =   51
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   74
      TabIndex        =   84
      Top             =   3360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox PicMines1m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   5040
      Picture         =   "Form2.frx":F9A0C
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   87
      Top             =   840
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox PicMines2m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   255
      Left            =   6600
      Picture         =   "Form2.frx":F9AF2
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   84
      TabIndex        =   89
      Top             =   840
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.PictureBox PicFBulletm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   720
      Left            =   8160
      Picture         =   "Form2.frx":F9BD8
      ScaleHeight     =   44
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   91
      Top             =   720
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.PictureBox PicExplode2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1110
      Left            =   0
      Picture         =   "Form2.frx":F9D82
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   94
      Top             =   6720
      Visible         =   0   'False
      Width           =   14460
   End
   Begin VB.PictureBox PicExplode2m 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1110
      Left            =   0
      Picture         =   "Form2.frx":12B144
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   960
      TabIndex        =   23
      Top             =   7080
      Visible         =   0   'False
      Width           =   14460
   End
   Begin VB.PictureBox PicBoss2Cm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1125
      Left            =   8280
      Picture         =   "Form2.frx":12D25E
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   100
      Top             =   2040
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox PicBoss2Bm 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1125
      Left            =   8280
      Picture         =   "Form2.frx":12D718
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   83
      TabIndex        =   102
      Top             =   3120
      Visible         =   0   'False
      Width           =   1305
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
