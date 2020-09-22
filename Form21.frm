VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "The Crate"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   4800
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton2 
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Back"
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
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Open a Mystery-Crate now..."
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
      Caption         =   "The Crate "
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "crate"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "You Have:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "Form21.frx":0000
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form21.frx":1272
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
If Label7.Caption = "0" Then
MsgBox "You don't have any crate."
Exit Sub
End If
Label7.Caption = Label7.Caption - Val(1)
SaveSetting "Digimon", "Profile", "Crate", Label7.Caption
Select Case GenerateRandom(1, 5)
Case 1
WillMoney "PLUS", 100000
MsgBox "You Win $100,000"
Case 2
MsgBox "You Have 150 point to balance for your pet."
Form22.Show
Unload Me
Form22.Tag = "150"
Form22.Timer1.Enabled = True
Case 3
WillPoint "PLUS", 1000
MsgBox "You Win 1000 exp. point!"
Case 4
Dim decreaseafterbattle As Long
decreaseafterbattle = GetSetting("Digimon", "Digimon", "AfterBattle")
decreaseafterbattle = decreaseafterbattle - Val(500)
If Val(decreaseafterbattle) < "1" Then decreaseafterbattle = "1"
SaveSetting "Digimon", "Digimon", "AfterBattle", decreaseafterbattle
MsgBox "The Ultimate creature power, denfence, and health will decrease 500 point!"
Case 5
MsgBox "EMPTY BOX! BAD LUCK!"
End Select
End Sub

Private Sub Flatbutton2_Click()
Form10.Show
Unload Me
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Label7.Caption = GetSetting("Digimon", "Profile", "Crate")
End Sub
