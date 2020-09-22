VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual Digimonz"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5130
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "DIGIMON"
      ForeColor       =   &H80000009&
      Height          =   1455
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   800
         Left            =   120
         Top             =   240
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   360
         TabIndex        =   9
         Top             =   240
         Width           =   495
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   800
         Left            =   2160
         Tag             =   "0"
         Top             =   480
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   240
         Width           =   495
         Begin VB.Image sleep_Image 
            Height          =   480
            Left            =   0
            Picture         =   "Form2.frx":0000
            Top             =   0
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2160
         Top             =   120
      End
      Begin VB.Image Image37 
         Height          =   480
         Left            =   360
         Top             =   720
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Food_Image 
         Height          =   480
         Left            =   1680
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Screen_Mon 
         Height          =   480
         Left            =   960
         Tag             =   "0"
         Top             =   600
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Current Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   3120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      Begin VB.Timer Anim_Heal 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1080
         Top             =   2160
      End
      Begin VB.Timer Anim_EnergyFull 
         Enabled         =   0   'False
         Interval        =   10
         Left            =   1080
         Top             =   1680
      End
      Begin VB.Timer StockMarket_Timer 
         Interval        =   30000
         Left            =   1080
         Top             =   1200
      End
      Begin VB.Timer Timer7 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   1080
         Top             =   720
      End
      Begin VB.Timer Timer2 
         Interval        =   60000
         Left            =   1080
         Top             =   240
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Owner: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defence: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Age: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Health: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   555
      End
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   300
      Left            =   1920
      TabIndex        =   14
      Top             =   2300
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "enough"
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
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "/100"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2535
      TabIndex        =   13
      Top             =   1800
      Width           =   345
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Money: $"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Energy:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   540
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu minimizemenu 
         Caption         =   "&Minimize"
      End
      Begin VB.Menu menuformfiledash1 
         Caption         =   "-"
      End
      Begin VB.Menu reset 
         Caption         =   "&Reset"
      End
      Begin VB.Menu resettogether 
         Caption         =   "&Reset Together"
      End
      Begin VB.Menu linemenu11121 
         Caption         =   "-"
      End
      Begin VB.Menu Burger 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu Option_menu 
      Caption         =   "&Option"
      Begin VB.Menu OwnerProfile 
         Caption         =   "&Owner Profile"
      End
      Begin VB.Menu BattleStatistic 
         Caption         =   "&Battle Statistic"
      End
      Begin VB.Menu onlineranking 
         Caption         =   "Online &Ranking"
      End
      Begin VB.Menu OptioMinusBar12 
         Caption         =   "-"
      End
      Begin VB.Menu setting 
         Caption         =   "&Setting"
      End
   End
   Begin VB.Menu Games 
      Caption         =   "&City"
      Begin VB.Menu Casino 
         Caption         =   "&Casino"
         Begin VB.Menu RandomWin 
            Caption         =   "&Open Mystery"
         End
         Begin VB.Menu BarofExcitement 
            Caption         =   "Bar of &Excitement"
         End
         Begin VB.Menu Lottery 
            Caption         =   "&Lottery"
         End
         Begin VB.Menu slot 
            Caption         =   "&Slot Machine"
         End
         Begin VB.Menu BlackJack 
            Caption         =   "Black&Jack"
            Enabled         =   0   'False
         End
         Begin VB.Menu bingo 
            Caption         =   "&Bingo"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu WizardShop 
         Caption         =   "&Wizard Shop"
      End
      Begin VB.Menu ItemStore 
         Caption         =   "&Item Store"
      End
      Begin VB.Menu bank 
         Caption         =   "&Bank"
      End
      Begin VB.Menu ProfessionalHire 
         Caption         =   "&Professional Hire"
         Begin VB.Menu Robber 
            Caption         =   "&Robber"
         End
         Begin VB.Menu Theif 
            Caption         =   "&Theif"
         End
         Begin VB.Menu Police 
            Caption         =   "&Police"
         End
      End
      Begin VB.Menu SecurityCompany 
         Caption         =   "&Security Company"
      End
      Begin VB.Menu TraningCenter 
         Caption         =   "T&raning Center"
      End
      Begin VB.Menu NewsCenter 
         Caption         =   "&News Center"
      End
      Begin VB.Menu ShareMarket 
         Caption         =   "Stock &Market"
      End
      Begin VB.Menu gilalinecity 
         Caption         =   "-"
      End
      Begin VB.Menu CasinoDouble 
         Caption         =   "Casino &Double!"
      End
   End
   Begin VB.Menu Battle 
      Caption         =   "&Battle"
      Begin VB.Menu TCPIP 
         Caption         =   "&TCP/IP"
      End
      Begin VB.Menu LevelGainer 
         Caption         =   "&Battle Arena"
      End
      Begin VB.Menu dashforonlinetournament 
         Caption         =   "-"
      End
      Begin VB.Menu tournament 
         Caption         =   "&Tournament"
         Begin VB.Menu OnlineTournament 
            Caption         =   "&Online Tournament"
         End
         Begin VB.Menu BotMatchTournament 
            Caption         =   "&BotMatch Tournament"
            Enabled         =   0   'False
         End
         Begin VB.Menu tournamentsubdash 
            Caption         =   "-"
         End
         Begin VB.Menu CheckTournament 
            Caption         =   "&Check Tournament"
         End
      End
   End
   Begin VB.Menu Events 
      Caption         =   "&Events"
      Begin VB.Menu Carry 
         Caption         =   "C&arry"
         Begin VB.Menu Item 
            Caption         =   "&Item"
         End
      End
      Begin VB.Menu feed 
         Caption         =   "&Feed"
      End
      Begin VB.Menu clean 
         Caption         =   "&Clean"
      End
      Begin VB.Menu rest 
         Caption         =   "&Sleep"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "&Help"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu helptopic 
         Caption         =   "&Help Topic"
         Shortcut        =   {F1}
      End
      Begin VB.Menu errorfixed 
         Caption         =   "&Error Fixed / Update"
      End
      Begin VB.Menu lineabdjksbakdjas 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "&About"
         Shortcut        =   {F5}
      End
      Begin VB.Menu CheckUpdate 
         Caption         =   "&Check for Update"
      End
      Begin VB.Menu Register 
         Caption         =   "&Register"
         Shortcut        =   {F12}
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Maxfood

Private Sub About_Click()
UnloadSubForm
Form11.Show
End Sub

Private Sub Anim_EnergyFull_Timer()
ProgressBar1.Value = ProgressBar1.Value + Val(1)
Form2.ProgressBar1.ToolTipText = Form2.ProgressBar1.Value & "/100"
Form2.Label9.Caption = Form2.ProgressBar1.Value & "/100"
If ProgressBar1.Value = "100" Then Anim_EnergyFull.Enabled = False
End Sub

Private Sub Anim_Heal_Timer()
If AnimHealStart = GetSetting("Digimon", "Digimon", "CurrentHealth") Then Anim_Heal.Enabled = False
Form2.Label1.Caption = left(Form2.Label1.Caption, 8) & AnimHealStart & "/" & GetSetting("Digimon", "Digimon", "health")
AnimHealStart = AnimHealStart + Val(1)
End Sub

Private Sub bank_Click()
UnloadSubForm
Form27.Show
End Sub

Private Sub BarofExcitement_Click()
UnloadSubForm
Form33.Show
End Sub

Private Sub BattleStatistic_Click()
UnloadSubForm
Form15.Show
Form15.Tag = "Main"
End Sub

Private Sub Burger_Click()
Unload Form39
UnloadSubForm
m_cN.Detach
Unload Me
End Sub

Private Sub CasinoDouble_Click()
SaveSetting "Digimon", "Profile", "CasinoDouble", "0"
CasinoDouble.Enabled = False
UnloadSubForm
Form26.Show
Form26.Tag = GetSetting("Digimon", "Profile", "DoubleType")
Form26.Timer1.Enabled = True
End Sub

Private Sub CheckTournament_Click()
If CheckRegister = "ExitSub" Then Exit Sub
CheckTournamentValue = GetUrlSource("http://www.angelfire.com/co/choongyouqi/VirtualDigimonz/TournamentIP.dat")
If CheckTournamentValue = "" Then
MsgBox "Sorry, currently no any online tournament available."
Exit Sub
End If
UnloadSubForm
Form36.Show
Form36.Text1 = CheckTournamentValue
End Sub

Private Sub CheckUpdate_Click()
MousePointer = vbHourglass
CheckUpdate.Tag = GetUrlSource("http://www.angelfire.com/co/choongyouqi/VirtualDigimonz/Update.dat")
checkupdatesize = GetUrlSource("http://www.angelfire.com/co/choongyouqi/VirtualDigimonz/Filesize.dat")
MousePointer = vbDefault
If CheckUpdate.Tag = "" Then
MsgBox "Please Connect to Internet 1st."
Exit Sub
End If
If left(CheckUpdate.Tag, 4) = App.Major & "." & App.Minor & App.Revision Then
MsgBox "Sorry, currently no new version available on the internet."
Exit Sub
End If
AskUpdate = MsgBox("Your Version: " & App.Major & "." & App.Minor & App.Revision & vbCrLf & "Newest Version: " & left(CheckUpdate.Tag, 4) & vbCrLf & "Download: " & Mid(CheckUpdate.Tag, 5), vbYesNo, "File Size: " & checkupdatesize)
If AskUpdate = vbYes Then
OpenNotepad = ShellExecute(hWnd, "open", Mid(CheckUpdate.Tag, 5), vbNull, vbNull, SW_SHOWNORMAL)
End If
End Sub

Private Sub clean_Click()
Dim moneyclean
moneyclean = GetSetting("Digimon", "Profile", "Money")
moneyclean = Val(moneyclean) - Val(50)
If left(moneyclean, 1) = "-" Then
MsgBox "Sorry, You don't have enough money!"
Exit Sub
End If

cleaningquestion = MsgBox("This cost $50 for cleaning, Are you sure?", vbYesNo)
If cleaningquestion = vbYes Then
SaveSetting "Digimon", "Profile", "Money", moneyclean
Timer5.Enabled = True
End If
End Sub


Private Sub FlatButton1_Click()
sleep_Image.Visible = False
Timer7.Enabled = False
FlatButton1.Visible = False
Label8.Visible = False
AllMenu (True)
If GetSetting("Digimon", "Profile", "CasinoDouble") = "1" Then CasinoDouble.Enabled = True
If GetSetting("Digimon", "Application", "Register") = "0" Then Register.Enabled = True
End Sub

Private Sub errorfixed_Click()
UnloadSubForm
Form20.Show
End Sub

Private Sub feed_Click()
Unload Form28
Form8.Show
End Sub

Private Sub Form_Load()
Form39.Show
CheckMidiPlay
AutoDetectTop Me

Set m_cN = New cNeoCaption
Skin Me, m_cN
If GetSetting("Digimon", "Application", "Register") = "1" Then
Register.Enabled = False
Me.Caption = "Virtual Digimonz " & App.Major & "." & App.Minor & App.Revision
Else
Me.Caption = "Virtual Digimonz " & App.Major & "." & App.Minor & App.Revision & " (Un)"
End If
Image37.Picture = Form7.Image37.Picture
ProgressBar1.Value = GetSetting("Digimon", "Digimon", "Energy")
ProgressBar1.ToolTipText = ProgressBar1.Value & "/100"
Label9.Caption = ProgressBar1.Value & "/100"
Maxfood = GetSetting("Digimon", "Profile", "MaxFood")
DigimonType
Frame2.Caption = Frame2.Caption & " -- " & GetSetting("Digimon", "Digimon", "Name")
Label1.Caption = left(Label1.Caption, 8) & GetSetting("Digimon", "Digimon", "CurrentHealth") & "/" & GetSetting("Digimon", "Digimon", "health")
Label2.Caption = Label2.Caption & GetSetting("Digimon", "Profile", "Played") & "min"
Label3.Caption = Label3.Caption & GetSetting("Digimon", "Digimon", "Power")
Label4.Caption = Label4.Caption & GetSetting("Digimon", "Digimon", "defence")
Label5.Tag = Label5.Tag & GetSetting("Digimon", "Digimon", "Fooded")
Label6.Caption = Label6.Caption & GetSetting("Digimon", "Profile", "Name")
Timer2.Tag = GetSetting("Digimon", "Profile", "Played")
Food_Event_Caption
Label5.ToolTipText = Label5.Tag & "/" & Maxfood
If GetSetting("Digimon", "Digimon", "Shit") = "1" Then Image37.Visible = True
If GetSetting("Digimon", "Profile", "CasinoDouble") = "0" Then CasinoDouble.Enabled = False
End Sub

Private Sub helptopic_Click()
UnloadSubForm
Form25.Show
End Sub

Private Sub Item_Click()
Unload Form8
Form28.Show
End Sub

Private Sub ItemStore_Click()
UnloadSubForm
Form32.Show
End Sub

Private Sub LevelGainer_Click()
UnloadSubForm
Form12.Show
End Sub

Private Sub Lottery_Click()
UnloadSubForm
Form37.Show
End Sub

Private Sub minimizemenu_Click()
Me.WindowState = 1
End Sub

Private Sub NewsCenter_Click()
UnloadSubForm
Form30.Show
End Sub

Private Sub onlineranking_Click()
If CheckRegister = "ExitSub" Then Exit Sub
UnloadSubForm
Form35.Show
End Sub

Private Sub OnlineTournament_Click()
If CheckRegister = "ExitSub" Then Exit Sub
UnloadSubForm
Form36.Show
End Sub

Private Sub OwnerProfile_Click()
UnloadSubForm
Form3.Show
End Sub

Private Sub RandomWin_Click()
UnloadSubForm
Form9.Show
End Sub

Private Sub Register_Click()
UnloadSubForm
Form34.Show
End Sub

Private Sub Reset_Click()
a = MsgBox("Are You Sure YOU want to RESET??", vbYesNo, "RESET")
If a = vbYes Then
Reset_Program
MsgBox ("Please Open Again This Program")
Unload Form39
UnloadSubForm
m_cN.Detach
Unload Me
End If
End Sub

Private Sub resettogether_Click()
UnloadSubForm
Form5.Show
Form5.Tag = "reset"
Form5_Event
End Sub

Private Sub rest_Click()
resmsgboxask1 = MsgBox("This will cost you $2 per sec. OK?", vbYesNo)
If resmsgboxask1 = vbYes Then
Unload Form8
Unload Form5
Unload Form28
FlatButton1.Visible = True
Timer7.Enabled = True
Label8.Visible = True
Label8.Caption = "Money: $" & GetSetting("Digimon", "Profile", "Money")
AllMenu (False)
CasinoDouble.Enabled = False
Register.Enabled = False
sleep_Image.Visible = True
End If
End Sub

Private Sub setting_Click()
UnloadSubForm
Form24.Show
End Sub

Private Sub ShareMarket_Click()
UnloadSubForm
Form31.Show
End Sub

Private Sub slot_Click()
UnloadSubForm
Form38.Show
End Sub

Private Sub StockMarket_Timer_Timer()
ShareMarketValue
StockMarket_Timer.Interval = GenerateRandom(1, 30000)
End Sub

Private Sub TCPIP_Click()
UnloadSubForm
Form5.Show
Form5.Tag = "battle"
Form5_Event
End Sub

Private Sub Timer1_Timer()
Select Case Screen_Mon.Tag
Case "0"
Screen_Mon.top = Screen_Mon.top - 100
Screen_Mon.Tag = "1"
Case "1"
Screen_Mon.top = Screen_Mon.top + 100
Screen_Mon.Tag = "0"
End Select
End Sub

Private Sub Timer2_Timer()
Timer2.Tag = Val(Timer2.Tag) + Val(1)
SaveSetting "Digimon", "Profile", "Played", Timer2.Tag
Label2.Caption = left(Label2.Caption, 5) & GetSetting("Digimon", "Profile", "Played") & "min"
Select Case GenerateRandom(1, 10)
Case 1
WillFood "MINUS", 1
Case 3
WillFood "MINUS", 2
Case 5
WillFood "MINUS", 3
End Select

LotteryTimer
WillEnergy "MINUS", 1
Personality
ShitTimer
BankInterest
End Sub

Private Sub Timer3_Timer()
Frame3.Width = Food_Image.Width
Frame3.Height = Food_Image.Height
If Frame3.top = 720 Then
Frame3.top = Frame3.top - 480
Food_Image.Picture = Form7.Image2.Picture
WillFood "PLUS", Timer3.Tag
Food_Event_Caption
Timer3.Tag = "0"
Timer3.Enabled = False
Exit Sub
End If
Frame3.top = Frame3.top + 120
End Sub

Private Sub Timer5_Timer()
If Frame4.top = 720 Then
Frame4.top = Frame4.top - 480
Image37.Visible = False
SaveSetting "Digimon", "Digimon", "Shit", "0"
Timer5.Enabled = False
Exit Sub
End If
Frame4.top = Frame4.top + 120
End Sub

Private Sub Timer7_Timer()
If WillMoney("MINUS", 2) = "ExitSub" Then
Timer7.Enabled = False
FlatButton1.Visible = False
Label8.Visible = False
AllMenu (True)
If GetSetting("Digimon", "Profile", "CasinoDouble") = "1" Then CasinoDouble.Enabled = True
If GetSetting("Digimon", "Application", "Register") = "0" Then Register.Enabled = True
sleep_Image.Visible = False
Exit Sub
End If
Label8.Caption = "Money: $" & GetSetting("Digimon", "Profile", "Money")
WillEnergy "PLUS", 1
End Sub

Private Sub TraningCenter_Click()
UnloadSubForm
Form29.Show
End Sub
Private Sub WizardShop_Click()
UnloadSubForm
Form10.Show
End Sub
