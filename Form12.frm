VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battle Arena"
   ClientHeight    =   3705
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   3795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   3795
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton8 
      Height          =   735
      Left            =   2040
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 8"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton2 
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   735
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton3 
      Height          =   735
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   735
      Left            =   360
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton5 
      Height          =   735
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton6 
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VirtualDigimonz.FlatButton FlatButton7 
      Height          =   735
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Opponent 7"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColor      =   16777215
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
   Begin VB.Menu BattleStatistic 
      Caption         =   "&Battle Statistic"
      Begin VB.Menu nothing1 
         Caption         =   "nothing1"
      End
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BattleStatistic_Click()
Form15.Show
Form15.Tag = "Out"
Unload Me
End Sub

Private Sub close_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub FlatButton1_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(5), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 5
Me.Hide
Form13.Show
Unload Me
End Sub

Private Sub FlatButton2_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(5), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 5
Me.Hide
Form13.Show
Unload Me
End Sub

Private Sub FlatButton3_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(10), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 10
Me.Hide
Form13.Show
Unload Me
End Sub

Private Sub Flatbutton4_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(15), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 15
Me.Hide
Form13.Show
Unload Me
End Sub
Private Sub Flatbutton5_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(15), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 15
Me.Hide
Form13.Show
Unload Me
End Sub
Private Sub FlatButton6_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(25), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 25
Me.Hide
Form13.Show
Unload Me
End Sub
Private Sub FlatButton7_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
If left(Val(Form2.ProgressBar1.Value) - Val(35), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 35
Me.Hide
Form13.Show
Unload Me
End Sub
Private Sub FlatButton8_Click()
If GetSetting("Digimon", "Digimon", "CurrentHealth") = "0" Then
MsgBox "Your digimon fainted. Please heal it 1st."
Exit Sub
End If
Select Case FlatButton8.Tag
Case ""
If left(Val(Form2.ProgressBar1.Value) - Val(45), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 45
Case "Ultimate"
If left(Val(Form2.ProgressBar1.Value) - Val(80), 1) = "-" Then
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Exit Sub
End If
WillEnergy "MINUS", 80
End Select
Me.Hide
Form13.Show
Unload Me
End Sub

Private Sub Form_Load()
AutoDetectTop Me
TournamentBotSkill = "0"
TournamentAvailable = "0"
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Select Case GetSetting("Digimon", "Profile", "Badge")
Case "0"
FlatButton1.Enabled = True
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = False
FlatButton5.Enabled = False
FlatButton6.Enabled = False
FlatButton7.Enabled = False
FlatButton8.Enabled = False
Case "1"
FlatButton1.Enabled = False
FlatButton2.Enabled = True
FlatButton3.Enabled = False
FlatButton4.Enabled = False
FlatButton5.Enabled = False
FlatButton6.Enabled = False
FlatButton7.Enabled = False
FlatButton8.Enabled = False
Case "2"
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = True
FlatButton4.Enabled = False
FlatButton5.Enabled = False
FlatButton6.Enabled = False
FlatButton7.Enabled = False
FlatButton8.Enabled = False
Case "3"
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = True
FlatButton5.Enabled = False
FlatButton6.Enabled = False
FlatButton7.Enabled = False
FlatButton8.Enabled = False
Case "4"
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = False
FlatButton5.Enabled = True
FlatButton6.Enabled = False
FlatButton7.Enabled = False
FlatButton8.Enabled = False
Case "5"
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = False
FlatButton5.Enabled = False
FlatButton6.Enabled = True
FlatButton7.Enabled = False
FlatButton8.Enabled = False
Case "6"
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = False
FlatButton5.Enabled = False
FlatButton6.Enabled = False
FlatButton7.Enabled = True
FlatButton8.Enabled = False
Case "7"
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = False
FlatButton5.Enabled = False
FlatButton6.Enabled = False
FlatButton7.Enabled = False
FlatButton8.Enabled = True
Case "8"
FlatButton5.Visible = False
FlatButton6.Visible = False
FlatButton7.Visible = False
thebottomhor = FlatButton8.left + FlatButton8.Width - FlatButton1.left - "200"
thebottomver = FlatButton8.top + FlatButton8.Height - "100"
FlatButton8.top = FlatButton1.top
FlatButton8.left = FlatButton1.left
FlatButton8.Width = thebottomhor
FlatButton8.Height = thebottomver
FlatButton8.Caption = "Shadow Knight" & vbCrLf & vbCrLf & "Digimon: Ultimate Creature" & vbCrLf & "POWER: EXTREME" & vbCrLf & "DEFENCE: EXTREME" & vbCrLf & "HEALTH: EXTREME" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "...The Legend Hero in the Digimon World..."
FlatButton8.Tag = "Ultimate"
End Select
End Sub
