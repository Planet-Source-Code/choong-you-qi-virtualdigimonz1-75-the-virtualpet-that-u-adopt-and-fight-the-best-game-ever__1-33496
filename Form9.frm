VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Casino > Open Mystery"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer crate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   840
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   960
      Top             =   1680
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   1680
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Status:"
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1455
      Begin VB.Label Label4 
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
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
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
         TabIndex        =   7
         Top             =   1680
         Width           =   120
      End
      Begin VB.Label Label8 
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
         TabIndex        =   6
         Top             =   1080
         Width           =   75
      End
      Begin VB.Label Label7 
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
         TabIndex        =   5
         Top             =   600
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Crates Collected: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp.Points: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2655
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1920
         Picture         =   "Form9.frx":0000
         ToolTipText     =   "$240 Per Play"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   1080
         Picture         =   "Form9.frx":08CA
         ToolTipText     =   "$80 Per Play"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "Form9.frx":1194
         ToolTipText     =   "$8 Per Play"
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   3360
      Picture         =   "Form9.frx":1A5E
      Top             =   1560
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose a game to play."
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
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Image Image12 
      Height          =   495
      Left            =   2040
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   1200
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   360
      Top             =   600
      Width           =   495
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   3360
      Picture         =   "Form9.frx":2328
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3360
      Picture         =   "Form9.frx":2BF2
      Top             =   1080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   3720
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   3720
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   3720
      Top             =   1200
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu leave 
      Caption         =   "&Leave"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MoneyRandom As Long
Private Sub crate_Timer()
Select Case GenerateRandom(1, 200)
Case 100
MsgBox "Congratulations! You found a crate."
acra = GetSetting("Digimon", "Profile", "Crate")
acra = acra + Val(1)
SaveSetting "Digimon", "Profile", "Crate", acra
Label7.Caption = GetSetting("Digimon", "Profile", "Crate")
End Select
crate.Enabled = False
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Image5.Picture = Image1.Picture
Image6.Picture = Image2.Picture
Image7.Picture = Image3.Picture
Image8.Picture = Image9.Picture
Image11.Picture = Image9.Picture
Image12.Picture = Image9.Picture
Label7.Caption = GetSetting("Digimon", "Profile", "Crate")
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image5.Picture
Image2.Picture = Image6.Picture
Image3.Picture = Image7.Picture
End Sub

Private Sub Image1_Click()
If Not Image8.Picture = Image9.Picture Then Exit Sub
If Not Image11.Picture = Image9.Picture Then Exit Sub
If Not Image12.Picture = Image9.Picture Then Exit Sub
If WillMoney("MINUS", 8) = "ExitSub" Then Exit Sub
Image8.Picture = Image10.Picture
Image11.Picture = Image9.Picture
Image12.Picture = Image9.Picture
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
Timer1.Enabled = True
End Sub

Private Sub Image2_Click()
If Not Image8.Picture = Image9.Picture Then Exit Sub
If Not Image11.Picture = Image9.Picture Then Exit Sub
If Not Image12.Picture = Image9.Picture Then Exit Sub
If WillMoney("MINUS", 80) = "ExitSub" Then Exit Sub
Image8.Picture = Image9.Picture
Image11.Picture = Image10.Picture
Image12.Picture = Image9.Picture
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
Timer3.Enabled = True
End Sub

Private Sub Image3_Click()
If Not Image8.Picture = Image9.Picture Then Exit Sub
If Not Image11.Picture = Image9.Picture Then Exit Sub
If Not Image12.Picture = Image9.Picture Then Exit Sub
If WillMoney("MINUS", 240) = "ExitSub" Then Exit Sub
Image8.Picture = Image9.Picture
Image11.Picture = Image9.Picture
Image12.Picture = Image10.Picture
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
Timer4.Enabled = True
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image4.Picture
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = Image4.Picture
End Sub
Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = Image4.Picture
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = Image5.Picture
End Sub
Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = Image6.Picture
End Sub
Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = Image7.Picture
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
Select Case GenerateRandom(1, 50)
Case 1 To 8
WillPoint "PLUS", 1
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "You Gain 1 Exp.Points"
Case 17 To 21
WillPoint "PLUS", 5
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "Not Bad! You Gain 5 Exp.Points"
Case 25 To 28
MoneyRandom = GenerateRandom(5, 100)
WillMoney "PLUS", MoneyRandom
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "You Found $" & MoneyRandom & " on the Floor."
Case 37 To 45
MoneyRandom = GenerateRandom(1, 100)
MinusMoney MoneyRandom
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "A Ghost come out and stole $" & MoneyRandom & " from you"
Case 49
MinusMoney 500
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "YOU MEET VAMPIRE! He grab $500 from you"
Case 50
WillPoint "PLUS", 15
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "JACKPOT! You Gain 15 Exp.Points"
End Select
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Image8.Picture = Image9.Picture
Image11.Picture = Image9.Picture
Image12.Picture = Image9.Picture
crate.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Select Case GenerateRandom(1, 20)
Case 1 To 2
WillPoint "PLUS", 2
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "You Gain 2 Exp.Points"
Case 4 To 5
MoneyRandom = GenerateRandom(80, 1000)
WillMoney "PLUS", MoneyRandom
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "You Found $" & MoneyRandom & " on the Floor."
Case 10
WillPoint "PLUS", 12
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "Not Bad! You Gain 12 Exp.Points"
Case 16 To 17
MoneyRandom = GenerateRandom(50, 200)
MinusMoney MoneyRandom
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "A Ghost come out and stole $" & MoneyRandom & " from you"
Case 19
MinusMoney 1000
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "YOU MEET VAMPIRE! He grab $1000 from you"
Case 20
WillPoint "PLUS", 35
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "JACKPOT! You Gain 35 Exp.Points"
End Select
Timer2.Enabled = True
Timer3.Enabled = False
End Sub

Private Sub Timer4_Timer()
Select Case GenerateRandom(1, 5)
Case 1
MoneyRandom = GenerateRandom(150, 350)
MinusMoney MoneyRandom
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "A Ghost come out and stole $" & MoneyRandom & " from you"
Case 2
MoneyRandom = GenerateRandom(240, 1500)
WillMoney "PLUS", MoneyRandom
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "You Found $" & MoneyRandom & " on the Floor."
Case 3
MinusMoney 1500
Label2.Caption = "$" & GetSetting("Digimon", "Profile", "Money")
MsgBox "YOU MEET VAMPIRE! He grab $1500 from you"
Case 5
WillPoint "PLUS", 60
Label8.Caption = GetSetting("Digimon", "Profile", "Score")
MsgBox "JACKPOT! You Gain 60 Exp.Points"
End Select
Timer2.Enabled = True
Timer4.Enabled = False
End Sub
