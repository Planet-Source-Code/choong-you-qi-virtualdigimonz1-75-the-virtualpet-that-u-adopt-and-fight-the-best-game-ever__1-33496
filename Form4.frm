VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup Your Profile"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4305
   StartUpPosition =   2  'CenterScreen
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&OK"
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
      Left            =   1560
      TabIndex        =   17
      Top             =   4800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Cancel"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Your Profile"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         MaxLength       =   9
         TabIndex        =   18
         Top             =   480
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   240
         Top             =   1080
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00400000&
         Height          =   550
         Left            =   240
         MousePointer    =   99  'Custom
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         Tag             =   "0"
         Top             =   600
         Width           =   550
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIC:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Your Name: (Max Lengh: 9)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   7
         Top             =   240
         Width           =   1950
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Email:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Your Digimon"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "Form4.frx":0000
         Left            =   1320
         List            =   "Form4.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   23
         Text            =   "13"
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   22
         Text            =   "7"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   840
         TabIndex        =   21
         Text            =   "5"
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   2160
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   15
         SelStart        =   13
         Value           =   13
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   1800
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   15
         SelStart        =   7
         Value           =   7
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   1440
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   15
         SelStart        =   5
         Value           =   5
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   2640
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digimon Type:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "25"
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
         Left            =   720
         TabIndex        =   12
         Top             =   2520
         Width           =   225
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Digimon Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "25 Points for average"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1845
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Defence:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Power:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Health: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   555
      End
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".............."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   540
      Index           =   0
      Left            =   0
      TabIndex        =   27
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".............."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   540
      Index           =   2
      Left            =   2520
      TabIndex        =   26
      Top             =   3600
      Width           =   1680
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H000080FF&
      Height          =   945
      Index           =   4
      Left            =   120
      TabIndex        =   25
      Top             =   2040
      Width           =   1245
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1680
      TabIndex        =   24
      Top             =   360
      Width           =   2325
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Select Case Combo1.Text
Case "Botamon"
Image1.Picture = Form7.Image11.Picture
Case "Tanemon"
Image1.Picture = Form7.Image29.Picture
Case "Tsunomon"
Image1.Picture = Form7.Image31.Picture
Case "Punimon"
Image1.Picture = Form7.Image28.Picture
End Select
End Sub

Private Sub FlatButton1_Click()
If Picture1.Tag = "0" Then
MsgBox ("Please Select A Picture For Yourself.")
Form6.Show
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "Please type Your Name."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please type Your Email Address."
Text2.SetFocus
Exit Sub
End If
If Text7.Text = "" Then
MsgBox "Please type Your Digimon Name"
Text7.SetFocus
Exit Sub
End If
If Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text) > 25 Then
MsgBox "Can't More than 25"
Exit Sub
End If
SaveSetting "Digimon", "Profile", "Name", Text1.Text
SaveSetting "Digimon", "Profile", "Email", Text2.Text
SaveSetting "Digimon", "Profile", "Played", "0"
SaveSetting "Digimon", "Profile", "Score", "0"
SaveSetting "Digimon", "Profile", "Money", "100"
SaveSetting "Digimon", "Profile", "Ranked", "Junior"
SaveSetting "Digimon", "Profile", "Picture", Picture1.Tag
SaveSetting "Digimon", "Profile", "Badge", "0"
SaveSetting "Digimon", "Profile", "Crate", "0"
SaveSetting "Digimon", "Profile", "Version", App.Major & "." & App.Minor & App.Revision
SaveSetting "Digimon", "Profile", "Hour", "0"
SaveSetting "Digimon", "Profile", "MaxFood", "20"
SaveSetting "Digimon", "Profile", "Bank", "0"
SaveSetting "Digimon", "Profile", "CasinoDouble", "0"
SaveSetting "Digimon", "Profile", "DoubleType", "none"
SaveSetting "Digimon", "Profile", "LastTimeBank", "0"
SaveSetting "Digimon", "Profile", "Bonus", "0"
SaveSetting "Digimon", "Profile", "OnlineScore", "0"
SaveSetting "Digimon", "Profile", "BankInterest", "0"

SaveSetting "Digimon", "Digimon", "Name", Text7.Text
SaveSetting "Digimon", "Digimon", "Type", Combo1.Text
SaveSetting "Digimon", "Digimon", "Power", Text3.Text
SaveSetting "Digimon", "Digimon", "Defence", Text4.Text
SaveSetting "Digimon", "Digimon", "Health", Text5.Text
SaveSetting "Digimon", "Digimon", "Fooded", "8"
SaveSetting "Digimon", "Digimon", "Win", "0"
SaveSetting "Digimon", "Digimon", "Lose", "0"
SaveSetting "Digimon", "Digimon", "Draw", "0"
SaveSetting "Digimon", "Digimon", "AfterBattle", "550"
SaveSetting "Digimon", "Digimon", "Shit", "0"
SaveSetting "Digimon", "Digimon", "Energy", "100"
SaveSetting "Digimon", "Digimon", "CurrentHealth", Text5.Text

SaveSetting "Digimon", "Item", "1", "0"
SaveSetting "Digimon", "Item", "2", "0"
SaveSetting "Digimon", "Item", "3", "0"
SaveSetting "Digimon", "Item", "4", "0"
SaveSetting "Digimon", "Item", "5", "0"
SaveSetting "Digimon", "Item", "deposit1", "0"
SaveSetting "Digimon", "Item", "deposit2", "0"
SaveSetting "Digimon", "Item", "deposit3", "0"
SaveSetting "Digimon", "Item", "deposit4", "0"
SaveSetting "Digimon", "Item", "deposit5", "0"
SaveSetting "Digimon", "Item", "space", "2"
SaveSetting "Digimon", "Item", "allsub", "0"

SaveSetting "Digimon", "Share", "Own1", "0"
SaveSetting "Digimon", "Share", "Own2", "0"
SaveSetting "Digimon", "Share", "Own3", "0"
SaveSetting "Digimon", "Share", "Price1", GenerateRandom(10, 100)
SaveSetting "Digimon", "Share", "Price2", GenerateRandom(100, 300)
SaveSetting "Digimon", "Share", "Price3", GenerateRandom(50, 200)

SaveSetting "Digimon", "Application", "Path", App.Path
SaveSetting "Digimon", "Application", "PathAndFile", App.Path & "\" & "Virtual Digimonz.exe"
SaveSetting "Digimon", "Application", "Register", "0"

SaveSetting "Digimon", "Online", "Player1", "none"
SaveSetting "Digimon", "Online", "Player2", "none"
SaveSetting "Digimon", "Online", "Player3", "none"
SaveSetting "Digimon", "Online", "Player4", "none"
SaveSetting "Digimon", "Online", "Player5", "none"
SaveSetting "Digimon", "Online", "Player6", "none"
SaveSetting "Digimon", "Online", "Player7", "none"
SaveSetting "Digimon", "Online", "Player8", "none"
SaveSetting "Digimon", "Online", "Player9", "none"
SaveSetting "Digimon", "Online", "Player10", "none"
SaveSetting "Digimon", "Online", "Score1", "0"
SaveSetting "Digimon", "Online", "Score2", "0"
SaveSetting "Digimon", "Online", "Score3", "0"
SaveSetting "Digimon", "Online", "Score4", "0"
SaveSetting "Digimon", "Online", "Score5", "0"
SaveSetting "Digimon", "Online", "Score6", "0"
SaveSetting "Digimon", "Online", "Score7", "0"
SaveSetting "Digimon", "Online", "Score8", "0"
SaveSetting "Digimon", "Online", "Score9", "0"
SaveSetting "Digimon", "Online", "Score10", "0"
SaveSetting "Digimon", "Online", "Time", time & " - " & DateTime.Date

SaveSetting "Digimon", "Digimon", "Setting1", "0"
SaveSetting "Digimon", "Digimon", "Setting1-1", ""
SaveSetting "Digimon", "Digimon", "Setting2", "1"
SaveSetting "Digimon", "Digimon", "Setting2-1", ""
SaveSetting "Digimon", "Digimon", "Setting3", "0"
SaveSetting "Digimon", "Digimon", "Setting4", "1"

SaveSetting "Digimon", "Lottery", "LotteryTime", "10"
SaveSetting "Digimon", "Lottery", "Lottery1", ""
SaveSetting "Digimon", "Lottery", "Lottery2", ""
SaveSetting "Digimon", "Lottery", "Lottery3", ""
SaveSetting "Digimon", "Lottery", "Lottery4", ""
SaveSetting "Digimon", "Lottery", "Lottery5", ""
SaveSetting "Digimon", "Lottery", "LotteryLastNum", "0"

SaveSetting "Digimon", "Setting", "IP1", "123.456.789.012"
SaveSetting "Digimon", "Setting", "IP2", ""
SaveSetting "Digimon", "Setting", "IP3", ""
SaveSetting "Digimon", "Setting", "IP4", ""

Me.Hide
m_cN.Detach
Form2.Show
Unload Me
End Sub

Private Sub FlatButton2_Click()
Me.Hide
m_cN.Detach
DeleteSetting "Digimon", "FirstTime", "FirstTime"
End
End Sub

Private Sub Form_Load()
Combo1.ListIndex = 0
AutoDetectTop Me
'm_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
Picture1.MouseIcon = Form11.LabelCopyRight.MouseIcon
Unload Form11
Set m_cN = New cNeoCaption
Skin Me, m_cN
Label10.Caption = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
End Sub

Private Sub Label10_Change()
If Label10.Caption > 25 Then
Label10.ForeColor = vbRed
Else
Label10.ForeColor = vbWhite
End If
End Sub

Private Sub Picture1_Click()
Form6.Show
End Sub

Private Sub Slider1_Click()
Text3.Text = Slider1.Value
End Sub

Private Sub Slider1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text3.Text = Slider1.Value
End Sub

Private Sub Slider2_Click()
Text4.Text = Slider2.Value
End Sub

Private Sub Slider2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text4.Text = Slider2.Value
End Sub

Private Sub Slider3_Click()
Text5.Text = Slider3.Value
End Sub

Private Sub Slider3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Text5.Text = Slider3.Value
End Sub

Private Sub Text3_Change()
Label10.Caption = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
End Sub

Private Sub Text4_Change()
Label10.Caption = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
End Sub

Private Sub Text5_Change()
Label10.Caption = Val(Text3.Text) + Val(Text4.Text) + Val(Text5.Text)
End Sub

Private Sub Timer1_Timer()
Text1.SetFocus
Timer1.Enabled = False
End Sub
