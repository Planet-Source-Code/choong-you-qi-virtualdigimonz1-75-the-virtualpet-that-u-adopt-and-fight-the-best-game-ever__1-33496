VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form13 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ranking Battle Arena"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form13"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "You now are the Campion, No need to fight already!"
      Height          =   1335
      Left            =   4920
      TabIndex        =   13
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2880
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1320
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Start"
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   7
      Tag             =   "0"
      Text            =   "0"
      Top             =   1920
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5640
      Tag             =   "0"
      Top             =   840
   End
   Begin VB.Timer Update 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   120
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   855
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   855
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   25
      Scrolling       =   1
   End
   Begin VB.Label HEALTH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label DEFENCE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label POWER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4320
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press: Alt+S to start Battle!"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Tag             =   "0"
      Top             =   1920
      Width           =   1875
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   3360
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   840
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   2520
      Picture         =   "Form13.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   1560
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   5160
      Picture         =   "Form13.frx":0CCA
      Tag             =   "Greymon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   5160
      Picture         =   "Form13.frx":1594
      Tag             =   "Angemon"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image8 
      Height          =   480
      Left            =   5160
      Picture         =   "Form13.frx":1E5E
      Tag             =   "Kunemon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   5160
      Picture         =   "Form13.frx":2168
      Tag             =   "Bakemon"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   4680
      Picture         =   "Form13.frx":2472
      Tag             =   "Tyrano"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   4680
      Picture         =   "Form13.frx":277C
      Tag             =   "Punimon"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   4680
      Picture         =   "Form13.frx":3446
      Tag             =   "Tanemon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   4680
      Picture         =   "Form13.frx":4110
      Tag             =   "Dijitama"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3960
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Timer3.Enabled = True
Image18.Visible = True
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Timer2.Enabled = True
Image17.Visible = True
Command2.Enabled = False
End Sub

Private Sub Command3_Click()
Form2.Show
Unload Me
End Sub

Private Sub Command4_Click()
If Label4.Tag = "1" Then Exit Sub
Label4.Tag = "1"
FormBattleShow = "1"
Text1.BackColor = vbYellow
Text1.ForeColor = vbBlack
Timer1.Tag = "1"
Timer1.Interval = GenerateRandom(1000, 3200)
Timer1.Enabled = True
Text1.SetFocus
End Sub

Private Sub Form_Load()
If Not GetSetting("Digimon", "Digimon", "Setting1-1") = "" Then Me.Picture = LoadPicture(GetSetting("Digimon", "Digimon", "Setting1-1"))
AutoDetectTop Me
Image17.Picture = Form7.Image39.Picture
Image18.Picture = Form7.Image38.Picture
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Image1.Picture = Form2.Screen_Mon.Picture
Dim Ibadge As Long
Ibadge = Val(GetSetting("Digimon", "Profile", "Badge")) + Val(1)
Form13.Tag = Ibadge
If Not TournamentBotSkill = "0" Then Form13.Tag = TournamentBotSkill
Select Case Form13.Tag
Case "1"
Image12.Picture = Form6.Image1.Picture
Image2.Picture = Image3.Picture
Label3.Caption = Image3.Tag
POWER.Caption = "15"
DEFENCE.Caption = "15"
HEALTH.Caption = "15"
Case "2"
Image12.Picture = Form6.Image2.Picture
Image2.Picture = Image4.Picture
Label3.Caption = Image4.Tag
POWER.Caption = "20"
DEFENCE.Caption = "20"
HEALTH.Caption = "20"
Case "3"
Image12.Picture = Form6.Image3.Picture
Image2.Picture = Image5.Picture
Label3.Caption = Image5.Tag
POWER.Caption = "30"
DEFENCE.Caption = "30"
HEALTH.Caption = "30"
Case "4"
Image12.Picture = Form6.Image4.Picture
Image2.Picture = Image6.Picture
Label3.Caption = Image6.Tag
POWER.Caption = "50"
DEFENCE.Caption = "50"
HEALTH.Caption = "50"
Case "5"
Image12.Picture = Form6.Image5.Picture
Image2.Picture = Image7.Picture
Label3.Caption = Image7.Tag
POWER.Caption = "75"
DEFENCE.Caption = "75"
HEALTH.Caption = "75"
Case "6"
Image12.Picture = Form6.Image6.Picture
Image2.Picture = Image8.Picture
Label3.Caption = Image8.Tag
POWER.Caption = "150"
DEFENCE.Caption = "150"
HEALTH.Caption = "150"
Case "7"
Image12.Picture = Form6.Image7.Picture
Image2.Picture = Image9.Picture
Label3.Caption = Image9.Tag
POWER.Caption = "350"
DEFENCE.Caption = "350"
HEALTH.Caption = "350"
Case "8"
Image12.Picture = Form6.Image8.Picture
Image2.Picture = Image10.Picture
Label3.Caption = Image10.Tag
POWER.Caption = "500"
DEFENCE.Caption = "500"
HEALTH.Caption = "500"
Case "9"
Dim AfterBattle As Long
AfterBattle = GetSetting("Digimon", "Digimon", "AfterBattle")
Image12.Picture = Form6.Image8.Picture
Image2.Picture = Form7.Image32.Picture
Label3.Caption = "Ultimate Creature"
POWER.Caption = AfterBattle
DEFENCE.Caption = AfterBattle
HEALTH.Caption = AfterBattle
'Command3.Top = "0"
'Command3.Left = "0"
'Command3.Width = Me.Width - 50
'Command3.Height = Me.Height - 50
'Exit Sub
Case Else
Image12.Picture = Form6.Image8.Picture
Image2.Picture = Form7.Image32.Picture
Label3.Caption = TournamentBotName

POWER.Caption = TournamentBotSkill
DEFENCE.Caption = TournamentBotSkill
HEALTH.Caption = TournamentBotSkill


End Select
Select Case GetSetting("Digimon", "Profile", "Picture")
Case "1"
Image11.Picture = Form6.Image1.Picture
Case "2"
Image11.Picture = Form6.Image2.Picture
Case "3"
Image11.Picture = Form6.Image3.Picture
Case "4"
Image11.Picture = Form6.Image4.Picture
Case "5"
Image11.Picture = Form6.Image5.Picture
Case "6"
Image11.Picture = Form6.Image6.Picture
Case "7"
Image11.Picture = Form6.Image7.Picture
Case "8"
Image11.Picture = Form6.Image8.Picture
End Select
Image11.left = Label2.left - Image11.Width - Val(100)
Image12.left = Label2.left + Label2.Width + Val(100)
Label1.Caption = GetSetting("Digimon", "Profile", "name")
ProgressBar1.Max = GetSetting("Digimon", "Digimon", "health")
ProgressBar1.Value = GetSetting("Digimon", "Digimon", "CurrentHealth")
ProgressBar1.ToolTipText = ProgressBar1.Value & "/" & GetSetting("Digimon", "Digimon", "health")

'Start COM
ProgressBar2.Max = HEALTH.Caption
ProgressBar2.Value = ProgressBar2.Max
ProgressBar2.ToolTipText = ProgressBar2.Value & "/" & "25"
'END COM

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form14
End Sub

Private Sub Image1_Click()
Form14.Show
Form14.Image1.Picture = Image1.Picture
Form14.Label1.Caption = left(Form14.Label1.Caption, 6) & GetSetting("Digimon", "Digimon", "name")
Form14.Label2.Caption = left(Form14.Label2.Caption, 8) & GetSetting("Digimon", "Digimon", "health")
Form14.Label3.Caption = left(Form14.Label3.Caption, 7) & GetSetting("Digimon", "Digimon", "Power")
Form14.Label4.Caption = left(Form14.Label4.Caption, 9) & GetSetting("Digimon", "Digimon", "defence")
End Sub

Private Sub Image2_Click()
Form14.Show
Form14.Image1.Picture = Image2.Picture
Form14.Label1.Caption = left(Form14.Label1.Caption, 6) & Label3.Caption
Form14.Label2.Caption = left(Form14.Label2.Caption, 8) & ProgressBar2.Max
Form14.Label3.Caption = left(Form14.Label3.Caption, 7) & POWER.Caption
Form14.Label4.Caption = left(Form14.Label4.Caption, 9) & DEFENCE.Caption
End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Tag = "0" Then Exit Sub
If KeyCode = vbKeyLeft Then
If Text1.Tag = "0" Then Exit Sub
Text1.Tag = "0"
Text1.Text = Val(Text1.Text) + Val(5)
End If
If KeyCode = vbKeyRight Then
If Text1.Tag = "1" Then Exit Sub
Text1.Tag = "1"
End If
End Sub

Private Sub Timer1_Timer()
FormBattleShow = "0"
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
Image17.Tag = Text1.Text
'''''''COM'''''''
Image18.Tag = GenerateRandom(5, 75)
'''''END COM'''''
Command2.Value = True
Command1.Value = True
Text1.Text = "0"
Timer1.Tag = "0"
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If Image17.left = "3360" Then
Dim PlayerAttackDamage As Long
Dim MinusComBlood As Long
PlayerAttackDamage = Val(Image17.Tag) / Val(100) * Val(GetSetting("Digimon", "Digimon", "Power"))
Dim DefenceCalculationPlayer As Long
comreduce = GenerateRandom(5, 10)
DefenceCalculationPlayer = Val(DEFENCE.Caption * Val(1) / Val(comreduce))
MinusComBlood = ProgressBar2.Value - Val(PlayerAttackDamage - DefenceCalculationPlayer)
If Not DefenceCalculationPlayer > PlayerAttackDamage Then
If left(MinusComBlood, 1) = "-" Then MinusComBlood = "0"
ProgressBar2.Value = MinusComBlood
End If
Update.Enabled = True
Image17.Visible = False
Image17.left = "840"
Timer2.Enabled = False
Command2.Enabled = True
End If
Image17.left = Image17.left + 120
End Sub

Private Sub Timer3_Timer()
If Image18.left = "840" Then
Dim ComAttackDamage As Long
Dim MinusPlayerBlood As Long
ComAttackDamage = Val(Image18.Tag) / Val(100) * Val(POWER.Caption)
Dim DefenceCalculationCom As Long
playerreduce = GenerateRandom(5, 10)
DefenceCalculationCom = Val(GetSetting("Digimon", "Digimon", "defence") * Val(1) / Val(playerreduce))
MinusPlayerBlood = ProgressBar1.Value - Val(ComAttackDamage - DefenceCalculationCom)
If Not DefenceCalculationCom > ComAttackDamage Then
If left(MinusPlayerBlood, 1) = "-" Then MinusPlayerBlood = "0"
ProgressBar1.Value = MinusPlayerBlood
End If
Image18.Visible = False
Image18.left = "3360"
Timer3.Enabled = False
Command1.Enabled = True
End If
Image18.left = Image18.left - 120
End Sub
Private Sub Update_Timer()
SaveSetting "Digimon", "Digimon", "CurrentHealth", ProgressBar1.Value
Label4.Tag = "0"
Form2_Update
Dim PLUSbadge As Long
Dim PLUSLose As Long
Dim PLUSDraw As Long
Dim ValueAfterBattle As Long

If Form13.Tag = "9" Then
ProgressBar1.ToolTipText = ProgressBar1.Value & "/" & GetSetting("Digimon", "Digimon", "health")
ProgressBar2.ToolTipText = ProgressBar2.Value & "/" & ProgressBar2.Max
If ProgressBar2.Value = "0" Then
If ProgressBar1.Value = "0" Then
WillPoint "PLUS", 10
PLUSDraw = GetSetting("Digimon", "Digimon", "Draw")
PLUSDraw = Val(PLUSDraw) + Val(1)
SaveSetting "Digimon", "Digimon", "Draw", PLUSDraw

MsgBox "Draw"

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Exit Sub
End If
Dim Iwinhowmany As Long
Iwinhowmany = GenerateRandom(0, 1000)
winmultiple = Val(GetSetting("Digimon", "Profile", "Badge")) + Val(1)
WillMoney "PLUS", Val(Iwinhowmany) * Val(winmultiple)
WillPoint "PLUS", 15

ValueAfterBattle = GetSetting("Digimon", "Digimon", "AfterBattle")
ValueAfterBattle = Val(ValueAfterBattle) + Val(50)
SaveSetting "Digimon", "Digimon", "AfterBattle", ValueAfterBattle

PLUSwin = GetSetting("Digimon", "Digimon", "Win")
PLUSwin = Val(PLUSwin) + Val(1)
SaveSetting "Digimon", "Digimon", "Win", PLUSwin

MsgBox "You Win"
MsgBox "Winning Money: $" & Val(Val(Iwinhowmany) * Val(winmultiple))

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Exit Sub
End If
If ProgressBar1.Value = "0" Then
If ProgressBar2.Value = "0" Then
PLUSDraw = GetSetting("Digimon", "Digimon", "Draw")
PLUSDraw = Val(PLUSDraw) + Val(1)
SaveSetting "Digimon", "Digimon", "Draw", PLUSDraw
WillPoint "PLUS", 10

MsgBox "Draw"

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Exit Sub
End If
WillPoint "PLUS", 5
PLUSbadge = GetSetting("Digimon", "Profile", "Badge")
PLUSbadge = Val(PLUSbadge) - Val(1)
SaveSetting "Digimon", "Profile", "Badge", PLUSbadge

ValueAfterBattle = GetSetting("Digimon", "Digimon", "AfterBattle")
ValueAfterBattle = ValueAfterBattle - Val(50)
SaveSetting "Digimon", "Digimon", "AfterBattle", ValueAfterBattle

PLUSLose = GetSetting("Digimon", "Digimon", "Lose")
PLUSLose = PLUSLose + Val(1)
SaveSetting "Digimon", "Digimon", "Lose", PLUSLose

MsgBox "You Lose"

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Exit Sub
End If
Update.Enabled = False
Exit Sub
End If





'''''''''''For Lower Than 9'''''''''''''''





ProgressBar1.ToolTipText = ProgressBar1.Value & "/" & GetSetting("Digimon", "Digimon", "health")
ProgressBar2.ToolTipText = ProgressBar2.Value & "/" & ProgressBar2.Max
If ProgressBar2.Value = "0" Then
If ProgressBar1.Value = "0" Then
WillPoint "PLUS", 2
PLUSDraw = GetSetting("Digimon", "Digimon", "Draw")
PLUSDraw = Val(PLUSDraw) + Val(1)
SaveSetting "Digimon", "Digimon", "Draw", PLUSDraw

MsgBox "Draw"

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Select Case TournamentAvailable
Case "0"
Skin Form2, m_cN
Form2.Show
Case "1"
Skin Form36, m_cN
Form36.Show
Top8OR1 = Top8OR1 + 1
Form36.Winsock1.SendData "X" & Top8OR1 & "CoDe_DRAW_Code"
End Select
Unload Me

Exit Sub
End If
Iwinhowmany = GenerateRandom(0, 250)
winmultiple = Val(GetSetting("Digimon", "Profile", "Badge")) + Val(1)
WillMoney "PLUS", Val(Iwinhowmany) * Val(winmultiple)
WillPoint "PLUS", 3

PLUSbadge = GetSetting("Digimon", "Profile", "Badge")
PLUSbadge = Val(PLUSbadge) + Val(1)
If TournamentAvailable = 0 Then SaveSetting "Digimon", "Profile", "Badge", PLUSbadge

PLUSwin = GetSetting("Digimon", "Digimon", "Win")
PLUSwin = Val(PLUSwin) + Val(1)
SaveSetting "Digimon", "Digimon", "Win", PLUSwin

MsgBox "You Win"
MsgBox "Winning Money: $" & Val(Val(Iwinhowmany) * Val(winmultiple))

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Select Case TournamentAvailable
Case "0"
Skin Form2, m_cN
Form2.Show
Case "1"
Skin Form36, m_cN
Form36.Show
Top8OR1 = Top8OR1 + 1
Form36.Winsock1.SendData "X" & Top8OR1 & "CoDe_WIN_Code"
End Select
Unload Me

Exit Sub
End If
If ProgressBar1.Value = "0" Then
If ProgressBar2.Value = "0" Then
PLUSDraw = GetSetting("Digimon", "Digimon", "Draw")
PLUSDraw = Val(PLUSDraw) + Val(1)
SaveSetting "Digimon", "Digimon", "Draw", PLUSDraw
WillPoint "PLUS", 2

MsgBox "Draw"

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Select Case TournamentAvailable
Case "0"
Skin Form2, m_cN
Form2.Show
Case "1"
Skin Form36, m_cN
Form36.Show
Top8OR1 = Top8OR1 + 1
Form36.Winsock1.SendData "X" & Top8OR1 & "CoDe_DRAW_Code"
End Select
Unload Me

Exit Sub
End If
WillPoint "PLUS", 1
PLUSbadge = GetSetting("Digimon", "Profile", "Badge")
PLUSbadge = Val(PLUSbadge) - Val(1)

If left(PLUSbadge, 1) = "-" Then PLUSbadge = "0"
If TournamentAvailable = 0 Then SaveSetting "Digimon", "Profile", "Badge", PLUSbadge

PLUSLose = GetSetting("Digimon", "Digimon", "Lose")
PLUSLose = Val(PLUSLose) + Val(1)
SaveSetting "Digimon", "Digimon", "Lose", PLUSLose

MsgBox "You Lose"

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Select Case TournamentAvailable
Case "0"
Skin Form2, m_cN
Form2.Show
Case "1"
Skin Form36, m_cN
Form36.Show
Top8OR1 = Top8OR1 + 1
Form36.Winsock1.SendData "X" & Top8OR1 & "CoDe_LOSE_Code"
End Select
Unload Me

Exit Sub
End If
Update.Enabled = False


End Sub
