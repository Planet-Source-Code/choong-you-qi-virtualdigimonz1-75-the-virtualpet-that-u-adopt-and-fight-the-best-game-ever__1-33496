VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMidiPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public m_hWnd As Long

Public Sub Create(hWnd As Long)
    m_hWnd = hWnd
End Sub

Public Function PlayMusic(Filename As String) As Boolean
    If mciSendString("close all", ByVal 0&, 0, ByVal 0&) <> 0 Then PlayMusic = False: Exit Function
    If mciSendString("open " & Filename & " type sequencer alias MUSIC", ByVal 0&, 0, ByVal 0&) <> 0 Then PlayMusic = False: Exit Function
    If mciSendString("play MUSIC from 0 notify", ByVal 0&, 0, m_hWnd) <> 0 Then PlayMusic = False: Exit Function
    PlayMusic = True
End Function

Public Sub StopMusic()
    mciSendString "close all", ByVal 0&, 0, ByVal 0&
End Sub

Public Function PauseMusic() As Boolean
    If mciSendString("stop MUSIC", ByVal 0&, 0, ByVal 0&) <> 0 Then PauseMusic = False: Exit Function
    PauseMusic = True
End Function

Public Function ResumeMusic() As Boolean
    If mciSendString("play MUSIC notify", ByVal 0&, 0, ByVal 0&) <> 0 Then ResumeMusic = False: Exit Function
    ResumeMusic = True
End Function

Public Function RestartMusic() As Boolean
    If mciSendString("play MUSIC from 0 notify", ByVal 0&, 0, m_hWnd) <> 0 Then RestartMusic = False: Exit Function
    RestartMusic = True
End Function

Private Sub Class_Terminate()
    StopMusic
End Sub

