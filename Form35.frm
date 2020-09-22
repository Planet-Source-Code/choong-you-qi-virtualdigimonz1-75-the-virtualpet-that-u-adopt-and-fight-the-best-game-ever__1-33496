VERSION 5.00
Begin VB.Form Form35 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Online Ranking"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form35"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4770
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   4200
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1560
      TabIndex        =   25
      Top             =   4200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VirtualDigimonz.FlatButton FlatButton3 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Tag             =   "http://www2.domaindlx.com/choongyouqi/score/view_vb.asp"
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Get Score"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Tag             =   "0"
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Send Score"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton2 
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Online Ranking - High Scores"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   960
         Width           =   90
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   1200
         Width           =   90
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "6"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   1680
         Width           =   90
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "7"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   90
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "8"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   90
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2400
         Width           =   90
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   180
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   2520
         TabIndex        =   27
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1080
         TabIndex        =   26
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   23
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score1"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   22
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score2"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   20
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   960
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score3"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   17
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score4"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   16
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player5"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   1440
         Width           =   525
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score5"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   14
         Top             =   1440
         Width           =   510
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player6"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   13
         Top             =   1680
         Width           =   525
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score6"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   12
         Top             =   1680
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player7"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score7"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   10
         Top             =   1920
         Width           =   510
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player8"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   9
         Top             =   2160
         Width           =   525
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score8"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   8
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player9"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   2400
         Width           =   525
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score9"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   6
         Top             =   2400
         Width           =   510
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Player10"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   5
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Score10"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   2640
         Width           =   600
      End
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "YOUR SCORE: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   24
      Top             =   3240
      Width           =   1770
   End
End
Attribute VB_Name = "Form35"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub FlatButton1_Click()
If FlatButton1.Tag = "0" Then
MsgBox "You must get score 1st."
Exit Sub
End If
Dim FinalScore As Long
Dim abcmyrank As Integer
WinValue = GetSetting("Digimon", "Digimon", "Win")
Losevalue = GetSetting("Digimon", "Digimon", "Lose")
Drawvalue = GetSetting("Digimon", "Digimon", "Draw")
WinValue = WinValue * Val(3)
Losevalue = Losevalue * Val(1)
Drawvalue = Drawvalue * Val(2)
FinalScore = Val(WinValue) + Val(Losevalue) + Val(Drawvalue)
If FinalScore = GetSetting("Digimon", "Profile", "OnlineScore") Then
MsgBox "Sorry, You already submit your score this season." & vbCrLf & "Try again after you have a new score."
Exit Sub
End If
abcmyrank = WhatsMyRank(FinalScore)
If abcmyrank = 88 Then
MsgBox "You Never Hit The High Score."
Exit Sub
End If
If Not CheckDatabaseActive(Me) = "Activation" Then Exit Sub
Me.Hide
Form16.Show
MousePointer = vbHourglass
Isnothingornot = GetUrlSource("http://www2.domaindlx.com/choongyouqi/score/modify_vb.asp?name=" & GetSetting("Digimon", "Profile", "Name") & "&score=" & FinalScore & "&rank=" & abcmyrank)
Unload Form16
Me.Show
If Isnothingornot = "" Then Exit Sub
MousePointer = vbDefault
SaveSetting "Digimon", "Profile", "OnlineScore", FinalScore
MsgBox "Score Sent."
End Sub

Private Sub FlatButton2_Click()
SaveSetting "Digimon", "Online", "Player1", Label1.Caption
SaveSetting "Digimon", "Online", "Player2", Label3.Caption
SaveSetting "Digimon", "Online", "Player3", Label5.Caption
SaveSetting "Digimon", "Online", "Player4", Label7.Caption
SaveSetting "Digimon", "Online", "Player5", Label9.Caption
SaveSetting "Digimon", "Online", "Player6", Label11.Caption
SaveSetting "Digimon", "Online", "Player7", Label13.Caption
SaveSetting "Digimon", "Online", "Player8", Label15.Caption
SaveSetting "Digimon", "Online", "Player9", Label17.Caption
SaveSetting "Digimon", "Online", "Player10", Label19.Caption
SaveSetting "Digimon", "Online", "Score1", Label2.Caption
SaveSetting "Digimon", "Online", "Score2", Label4.Caption
SaveSetting "Digimon", "Online", "Score3", Label6.Caption
SaveSetting "Digimon", "Online", "Score4", Label8.Caption
SaveSetting "Digimon", "Online", "Score5", Label10.Caption
SaveSetting "Digimon", "Online", "Score6", Label12.Caption
SaveSetting "Digimon", "Online", "Score7", Label14.Caption
SaveSetting "Digimon", "Online", "Score8", Label16.Caption
SaveSetting "Digimon", "Online", "Score9", Label18.Caption
SaveSetting "Digimon", "Online", "Score10", Label20.Caption

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub FlatButton3_Click()
Me.Hide
Form16.Show
Write_Player = 1
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
Label10.Caption = ""
Label11.Caption = ""
Label12.Caption = ""
Label13.Caption = ""
Label14.Caption = ""
Label15.Caption = ""
Label16.Caption = ""
Label17.Caption = ""
Label18.Caption = ""
Label19.Caption = ""
Label20.Caption = ""


Dim AllPage
MousePointer = vbHourglass
AllPage = "c:\"
AllPage = GetUrlSource(FlatButton3.Tag)
MousePointer = vbDefault
'MsgBox AllPage
'MsgBox Len(AllPage)
Me.Show
Unload Form16
If AllPage = "" Then Exit Sub
If Not CheckDatabaseActive(Me) = "Activation" Then Exit Sub
Text1.Text = AllPage
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Frame1.Caption = "Online Ranking - High Scores @ " & GetSetting("Digimon", "Online", "Time")
Write_Player = 1
Dim FinalScore As Long
WinValue = GetSetting("Digimon", "Digimon", "Win")
Drawvalue = GetSetting("Digimon", "Digimon", "Draw")
Losevalue = GetSetting("Digimon", "Digimon", "Lose")
WinValue = WinValue * Val(3)
Drawvalue = Drawvalue * Val(2)
Losevalue = Losevalue * Val(1)
FinalScore = Val(WinValue) + Val(Losevalue) + Val(Drawvalue)
Label21.Caption = Label21.Caption & FinalScore


Label1.Caption = GetSetting("Digimon", "Online", "Player1")
Label3.Caption = GetSetting("Digimon", "Online", "Player2")
Label5.Caption = GetSetting("Digimon", "Online", "Player3")
Label7.Caption = GetSetting("Digimon", "Online", "Player4")
Label9.Caption = GetSetting("Digimon", "Online", "Player5")
Label11.Caption = GetSetting("Digimon", "Online", "Player6")
Label13.Caption = GetSetting("Digimon", "Online", "Player7")
Label15.Caption = GetSetting("Digimon", "Online", "Player8")
Label17.Caption = GetSetting("Digimon", "Online", "Player9")
Label19.Caption = GetSetting("Digimon", "Online", "Player10")
Label2.Caption = GetSetting("Digimon", "Online", "Score1")
Label4.Caption = GetSetting("Digimon", "Online", "Score2")
Label6.Caption = GetSetting("Digimon", "Online", "Score3")
Label8.Caption = GetSetting("Digimon", "Online", "Score4")
Label10.Caption = GetSetting("Digimon", "Online", "Score5")
Label12.Caption = GetSetting("Digimon", "Online", "Score6")
Label14.Caption = GetSetting("Digimon", "Online", "Score7")
Label16.Caption = GetSetting("Digimon", "Online", "Score8")
Label18.Caption = GetSetting("Digimon", "Online", "Score9")
Label20.Caption = GetSetting("Digimon", "Online", "Score10")

End Sub

Private Sub Timer1_Timer()
Do Until Len(Text1) = 0

Select Case left(Text1, 5)
Case "<rnk>"
Text1 = Mid(Text1, 6)
Rank_Value = left(Text1, 1)

Case "<nam>"
Text1 = Mid(Text1, 6)
WriteName = 1
WriteScore = 0
WriteRank = 0
Exit Sub

Case "<scr>"
Text1 = Mid(Text1, 6)
WriteName = 0
WriteScore = 1
WriteRank = 0
Exit Sub

Case "<nex>"
Rank_Value = 0
WriteName = 0
WriteScore = 0
WriteRank = 0
Text1 = Mid(Text1, 6)
Write_Player = Write_Player + Val(1)
If Write_Player = 11 Then
Timer1.Enabled = False
FlatButton1.Tag = "1"
End If
Exit Sub
End Select

If WriteName = "1" Then WriteNameModule
If WriteScore = "1" Then WriteScoreModule

Text1 = Mid(Text1, 2)
Loop

Timer1.Enabled = False
FlatButton1.Tag = "1"
SaveSetting "Digimon", "Online", "Time", time & " - " & DateTime.Date
Frame1.Caption = "Online Ranking - High Scores @ " & GetSetting("Digimon", "Online", "Time")
End Sub
