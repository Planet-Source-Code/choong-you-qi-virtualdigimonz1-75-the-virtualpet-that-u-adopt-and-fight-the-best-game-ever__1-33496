VERSION 5.00
Begin VB.Form Form34 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form34"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   3570
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Register Now"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "Female"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      Top             =   2160
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Male"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   1920
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Top             =   1200
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "Form34.frx":0000
      Left            =   1320
      List            =   "Form34.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".........................."
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
      Left            =   360
      TabIndex        =   17
      Top             =   840
      Width           =   3120
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00004080&
      BorderWidth     =   8
      Height          =   435
      Index           =   0
      Left            =   2760
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gendle: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Age:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   330
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   1560
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Digimon Name: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R.Name: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Top             =   480
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nickname: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   ".................."
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
      Height          =   495
      Index           =   3
      Left            =   -120
      TabIndex        =   19
      Top             =   1920
      Width           =   2115
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00400000&
      BorderWidth     =   15
      Height          =   615
      Index           =   1
      Left            =   240
      Top             =   240
      Width           =   615
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   1200
      Top             =   2160
      Width           =   1755
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H00400040&
      Height          =   225
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   1560
      Width           =   3045
   End
End
Attribute VB_Name = "Form34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
If Text2.Text = "" Then
Text2.SetFocus
Exit Sub
End If
If Text3.Text = "" Then
Text3.SetFocus
Exit Sub
End If
If Text5.Text = "" Then
Text5.SetFocus
Exit Sub
End If
Select Case Option1.Value
Case "1"
Gendle = "Male"
Case "0"
Gendle = "Female"
End Select
MousePointer = vbHourglass
Form16.Show
Me.Hide
Me.Tag = GetUrlSource( _
"http://www2.domaindlx.com/choongyouqi/register/add_vb.asp?" & _
"Nickname=" & Text1.Text & "&" & _
"Realname=" & Text2.Text & "&" & _
"Email=" & Text3.Text & "&" & _
"Age=" & Text5.Text & "&" & _
"Country=" & Combo1.Text & "&" & _
"Gendle=" & Gendle & "&" & _
"Digimon=" & Text4.Text & "&" & _
"Version=" & App.Major & "." & App.Minor & App.Revision & "&" & _
"Played=" & GetSetting("Digimon", "Profile", "Played"))
MousePointer = vbDefault
Unload Form16
Me.Show
If Not Me.Tag = "" Then
If Not CheckDatabaseActive(Me) = "Activation" Then Exit Sub

MsgBox "Register Success. Thanks for register this program."
MsgBox "For some award you register, we award you $3,000"
WillMoney "PLUS", 3000
SaveSetting "Digimon", "Application", "Register", "1"
Form2.Register.Enabled = False
RegisterForm2Caption = "Registered"
Form2.Caption = "Virtual Digimonz " & App.Major & "." & App.Minor & App.Revision
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Exit Sub
End If
MsgBox "Can't Register. Try again after you connect to internet."
End Sub

Private Sub FlatButton2_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Combo1.ListIndex = 0
Text1.Text = GetSetting("Digimon", "Profile", "Name")
Text3.Text = GetSetting("Digimon", "Profile", "Email")
Text4.Text = GetSetting("Digimon", "Digimon", "Name")
End Sub
