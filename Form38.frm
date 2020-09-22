VERSION 5.00
Begin VB.Form Form38 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slot Machine"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   3945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form38"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3945
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2160
      Top             =   240
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   240
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   720
      Top             =   240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   1080
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   2160
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   1440
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   720
      Top             =   1080
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   2160
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Tag             =   "0"
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   1440
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   1
      Tag             =   "4"
      Top             =   600
      Width           =   510
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   720
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Tag             =   "2"
      Top             =   600
      Width           =   510
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   12582912
      Caption         =   "Insert Money"
      HasFocusRect    =   0   'False
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
      Left            =   1920
      TabIndex        =   7
      Tag             =   "0"
      Top             =   1800
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   12582912
      Caption         =   "Pull"
      Enabled         =   0   'False
      HasFocusRect    =   0   'False
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
   Begin VB.Shape Shape4 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©VD casino."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Money:"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   630
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   1335
      Left            =   3120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Slot Machine"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   345
      Left            =   720
      TabIndex        =   3
      Top             =   1320
      Width           =   1920
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BackStyle       =   1  'Opaque
      Height          =   2535
      Left            =   480
      Shape           =   5  'Rounded Square
      Top             =   240
      Width           =   2535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00800000&
      BackStyle       =   1  'Opaque
      Height          =   255
      Left            =   1680
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Menu leave 
      Caption         =   "Leave"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Abccan

Private Sub FlatButton1_Click()
If Abccan = "1" Then Exit Sub
Abccan = "1"
If WillMoney("MINUS", 50) = "ExitSub" Then Exit Sub
Picture1.Tag = GenerateRandom(0, 4)
Picture2.Tag = GenerateRandom(2, 6)
Picture3.Tag = GenerateRandom(0, 4)
Timer1.Interval = GenerateRandom(1, 50)
Timer2.Interval = GenerateRandom(1, 50)
Timer3.Interval = GenerateRandom(1, 50)
'Timer4.Interval = GenerateRandom(1, 2000)
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
'Timer4.Enabled = True

Timer5.Interval = GenerateRandom(1000, 2000)
Timer6.Interval = GenerateRandom(1000, 2000)
Timer7.Interval = GenerateRandom(1000, 2000)
'Timer5.Enabled = True
FlatButton2.Tag = "1"
FlatButton1.Enabled = False
FlatButton2.Enabled = True
Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
End Sub

Private Sub FlatButton2_Click()
If FlatButton2.Tag = "0" Then Exit Sub
Timer5.Enabled = True
FlatButton2.Tag = "0"
FlatButton2.Enabled = False
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
End Sub

Private Sub leave_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Timer1_Timer()
Picture1.Tag = Picture1.Tag + Val(1)
Select Case Picture1.Tag
Case 1
Picture1.Picture = Form7.Image13.Picture
Case 2
Picture1.Picture = Form7.Image14.Picture
Case 3
Picture1.Picture = Form7.Image15.Picture
Case 4
Picture1.Picture = Form7.Image16.Picture
Case 5
Picture1.Picture = Form7.Image17.Picture
Picture1.Tag = 0
Case Else
Picture1.Tag = 0
End Select
End Sub

Private Sub Timer2_Timer()
Picture2.Tag = Picture2.Tag - Val(1)
Select Case Picture2.Tag
Case 1
Picture2.Picture = Form7.Image13.Picture
Picture2.Tag = 6
Case 2
Picture2.Picture = Form7.Image14.Picture
Case 3
Picture2.Picture = Form7.Image15.Picture
Case 4
Picture2.Picture = Form7.Image16.Picture
Case 5
Picture2.Picture = Form7.Image17.Picture
Case Else
Picture2.Tag = 6
End Select
End Sub

Private Sub Timer3_Timer()
Picture3.Tag = Picture3.Tag + Val(1)
Select Case Picture3.Tag
Case 1
Picture3.Picture = Form7.Image15.Picture
Case 2
Picture3.Picture = Form7.Image13.Picture
Case 3
Picture3.Picture = Form7.Image17.Picture
Case 4
Picture3.Picture = Form7.Image14.Picture
Case 5
Picture3.Picture = Form7.Image16.Picture
Picture3.Tag = 0
Case Else
Picture3.Tag = 0
End Select
End Sub

Private Sub Timer4_Timer()
Timer1.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
FlatButton1.Enabled = True

If Picture1.Picture = Picture2.Picture Then
If Picture2.Picture = Picture3.Picture Then
WillMoney "PLUS", 2000
MsgBox "JACKPOT! $2,000"
Abccan = "0"
Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
Exit Sub
End If
End If

If Picture1.Picture = Picture2.Picture Then
WillMoney "PLUS", 350
MsgBox "NOTBAD! $350"
Abccan = "0"
Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
Exit Sub
End If

If Picture2.Picture = Picture3.Picture Then
WillMoney "PLUS", 250
MsgBox "NOTBAD! $250"
Abccan = "0"
Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
Exit Sub
End If

If Picture1.Picture = Picture3.Picture Then
WillMoney "PLUS", 100
MsgBox "NOTBAD! $100"
Abccan = "0"
Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
Exit Sub
End If

Label3.Caption = "Your Money: $" & GetSetting("Digimon", "Profile", "Money")
MsgBox "Sorry!"
Abccan = "0"
End Sub

Private Sub Timer5_Timer()
Timer1.Enabled = False
Timer5.Enabled = False
Timer6.Enabled = True
End Sub

Private Sub Timer6_Timer()
Timer2.Enabled = False
Timer6.Enabled = False
Timer7.Enabled = True
End Sub

Private Sub Timer7_Timer()
Timer3.Enabled = False
Timer7.Enabled = False
Timer4.Enabled = True
End Sub
