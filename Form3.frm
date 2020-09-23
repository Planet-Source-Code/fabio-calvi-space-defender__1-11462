VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Space Defender - Game Over"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form3.frx":104A
   ScaleHeight     =   6600
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(Cancel As Integer)
If score >= hiscore Then
   Open App.Path + "\hiscore" For Output As #1
   Write #1, score
   Close #1
End If
FrmCredits.Show
Unload Me
Unload Form1
Unload Form2
End Sub
