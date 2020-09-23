Attribute VB_Name = "Sound"
'Function and constants used to play sounds.
'These are for .wav files
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function sndStopSound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszNull As Long, ByVal FLAGS As Integer) As Integer
'This one is for .mid files
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_MEMORY = &H4
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Global Const SND_NODEFAULT = &H2
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Const SPI_SCREENSAVERRUNNING = 97
Global Soundcard As Boolean

Public Function NoiseGet(ByVal FileName) As String
'Loads a sound file into a string variable
Dim buffer As String
Dim f As Integer
Dim SoundBuffer As String

    On Error GoTo NoiseGet_Error
    
    buffer = Space$(1024)
    SoundBuffer = ""
    f = FreeFile
    Open FileName For Binary As f
    Do While Not EOF(f)
        Get #f, , buffer    'Load in 1K chunks
        SoundBuffer = SoundBuffer & buffer
    Loop
    Close f
    NoiseGet = Trim$(SoundBuffer)
Exit Function

NoiseGet_Error:
    SoundBuffer = ""
    Exit Function
End Function
Public Sub NoisePlay(SoundBuffer As String, ByVal PlayMode As Integer)
Dim retcode As Integer
    If SoundBuffer = "" Then Exit Sub
    
    'stop any sound that may currently be playing
    retcode = sndStopSound(0, SND_ASYNC)
    
    'PlayMode should be SND_SYNC or SND_ASYNC
    retcode = sndPlaySound(ByVal SoundBuffer, PlayMode Or SND_MEMORY)
End Sub
Public Function Music() As Boolean
Dim lng As Long
lng = waveOutGetNumDevs()
If lng > 0 Then
   Soundcard = True
   Exit Function
Else
   Soundcard = False
   Exit Function
End If
End Function

