Attribute VB_Name = "Module6"
Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_MEMORY = &H4
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

Public Sub WAVStop()
    Call WAVPlay(" ")
End Sub

Public Sub WAVLoop(File)
    Dim SoundName As String
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_LOOP
    X = sndPlaySound(SoundName$, wFlags%)
End Sub
Public Sub WAVPlay(File)
Dim wFlags%
Dim X
    Dim SoundName As String
    SoundName$ = File
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X = sndPlaySound(SoundName$, wFlags%)
End Sub
