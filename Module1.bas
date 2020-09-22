Attribute VB_Name = "Module1"
Public Kof1
Public dtime
Public Sound
Option Explicit


Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
        ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
        ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal SoundName As String, ByVal Flags As Long) As Long
Global Const SND_ASYNC = &H1
Global Const SND_SYNC = &H0
Global Const SND_NOSTOP = &H10
Global Const SND_LOOP = &H8

Sub delay(ByVal DelayTime As Long)
    Dim tm1 As Long
    tm1 = timeGetTime
   Do
        If (timeGetTime - tm1) > DelayTime Then Exit Do
        DoEvents
   Loop
End Sub
