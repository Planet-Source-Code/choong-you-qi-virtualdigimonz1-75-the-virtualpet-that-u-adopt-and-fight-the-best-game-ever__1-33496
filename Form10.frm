VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shop/Market Stall"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4980
   StartUpPosition =   2  'CenterScreen
   Begin VirtualDigimonz.FlatButton FlatButton1 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Gain Defence"
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
      CausesValidation=   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Gain Attack"
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
   Begin VirtualDigimonz.FlatButton FlatButton3 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Gain Health"
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
   Begin VirtualDigimonz.FlatButton FlatButton4 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Exp. Points > Money"
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
   Begin VirtualDigimonz.FlatButton FlatButton5 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   1560
      TabIndex        =   7
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Crate Reward..."
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
   Begin VirtualDigimonz.FlatButton FlatButton6 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Evolve"
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
   Begin VirtualDigimonz.FlatButton FlatButton7 
      CausesValidation=   0   'False
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Leave"
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Have: "
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
      TabIndex        =   2
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Have: $"
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
      TabIndex        =   1
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  Welcome, what can I do for you?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1425
      Left            =   3120
      Picture         =   "Form10.frx":0000
      ToolTipText     =   " ShopKeeperJedi"
      Top             =   1320
      Width           =   1455
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
MsgboxAnswer = MsgBox("Gain Defence" & vbCrLf & "This will cost you $50 + 15 Exp.Points." & vbCrLf & "Are You Sure You Want To Continue?", vbYesNo)
If MsgboxAnswer = vbYes Then
If GetSetting("Digimon", "Profile", "Score") < Val(15) Then
MsgBox "You don't have enough Exp.Points!"
Exit Sub
End If
If WillMoney("MINUS", 50) = "ExitSub" Then Exit Sub
If WillPoint("MINUS", 15) = "ExitSub" Then Exit Sub
Dim GainDefenceRandom As Long
GainDefenceRandom = GenerateRandom(0, 5)
WillDefence "PLUS", GainDefenceRandom
Label2.Caption = left(Label2.Caption, 15) & GetSetting("Digimon", "Profile", "Money")
Label3.Caption = left(Label3.Caption, 14) & GetSetting("Digimon", "Profile", "Score") & " Exp.Points"
MsgBox "Success!" & vbCrLf & GetSetting("Digimon", "Digimon", "name") & " have gain " & GainDefenceRandom & " defence point."
Form2_Update
End If
End Sub

Private Sub Flatbutton2_Click()
MsgboxAnswer = MsgBox("Gain Attack" & vbCrLf & "This will cost you $75 + 20 Exp.Points." & vbCrLf & "Are You Sure You Want To Continue?", vbYesNo)
If MsgboxAnswer = vbYes Then
If GetSetting("Digimon", "Profile", "Score") < Val(20) Then
MsgBox "You don't have enough Exp.Points!"
Exit Sub
End If
If WillMoney("MINUS", 75) = "ExitSub" Then Exit Sub
If WillPoint("MINUS", 20) = "ExitSub" Then Exit Sub
Dim GainPowerRandom As Long
GainPowerRandom = GenerateRandom(0, 5)
WillAttack "PLUS", GainPowerRandom
Label2.Caption = left(Label2.Caption, 15) & GetSetting("Digimon", "Profile", "Money")
Label3.Caption = left(Label3.Caption, 14) & GetSetting("Digimon", "Profile", "Score") & " Exp.Points"
MsgBox "Success!" & vbCrLf & GetSetting("Digimon", "Digimon", "name") & " have gain " & GainPowerRandom & " Attack point."
Form2_Update
End If
End Sub

Private Sub FlatButton3_Click()
MsgboxAnswer = MsgBox("Gain Health" & vbCrLf & "This will cost you $40 + 10 Exp.Points." & vbCrLf & "Are You Sure You Want To Continue?", vbYesNo)
If MsgboxAnswer = vbYes Then
If GetSetting("Digimon", "Profile", "Score") < Val(10) Then
MsgBox "You don't have enough Exp.Points!"
Exit Sub
End If
If WillMoney("MINUS", 40) = "ExitSub" Then Exit Sub
If WillPoint("MINUS", 10) = "ExitSub" Then Exit Sub
Dim GainHealthRandom As Long
GainHealthRandom = GenerateRandom(0, 5)
WillHealth "PLUS", GainHealthRandom
Label2.Caption = left(Label2.Caption, 15) & GetSetting("Digimon", "Profile", "Money")
Label3.Caption = left(Label3.Caption, 14) & GetSetting("Digimon", "Profile", "Score") & " Exp.Points"
MsgBox "Success!" & vbCrLf & GetSetting("Digimon", "Digimon", "name") & " have gain " & GainHealthRandom & " Health point."
Form2_Update
End If
End Sub

Private Sub Flatbutton4_Click()
MsgboxAnswer = MsgBox("Are You Sure You Want To Change Exp.Points > Money?", vbYesNo)
If MsgboxAnswer = vbYes Then
If WillPoint("MINUS", 1) = "ExitSub" Then Exit Sub
WillMoney "PLUS", GenerateRandom(0, 200)
Label2.Caption = left(Label2.Caption, 15) & GetSetting("Digimon", "Profile", "Money")
Label3.Caption = left(Label3.Caption, 14) & GetSetting("Digimon", "Profile", "Score") & " Exp.Points"
End If
End Sub

Private Sub Flatbutton5_Click()
Form21.Show
Unload Me
End Sub

Private Sub FlatButton6_Click()
Form18.Show
Unload Me
End Sub

Private Sub FlatButton7_Click()
Form2_Update
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

Label2.Caption = left(Label2.Caption, 15) & GetSetting("Digimon", "Profile", "Money")
Label3.Caption = left(Label3.Caption, 14) & GetSetting("Digimon", "Profile", "Score") & " Exp.Points"
End Sub
