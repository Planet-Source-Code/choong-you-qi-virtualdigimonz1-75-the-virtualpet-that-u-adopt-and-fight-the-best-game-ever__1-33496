VERSION 5.00
Begin VB.Form Form27 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bank"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form27"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Bank"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.Timer Interest 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1440
         Top             =   2280
      End
      Begin VB.Timer Update 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   960
         Top             =   2280
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Your Account"
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
         Begin VirtualDigimonz.FlatButton FlatButton2 
            Height          =   495
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            ForeColor       =   12632256
            BackColor       =   0
            Caption         =   "&Deposit"
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
            Height          =   495
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   873
            ForeColor       =   12632256
            BackColor       =   0
            Caption         =   "&Withdraw"
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
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Your Info"
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   3255
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your hand cash: "
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank cash:"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   480
            TabIndex        =   3
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*Bank will give you an 1% interest per minute."
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   2400
         Width           =   3195
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000040&
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   120
      Width           =   1575
   End
   Begin VB.Menu Leave 
      Caption         =   "&Leave"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
   Begin VB.Menu Cheque 
      Caption         =   "&Cheque"
      Begin VB.Menu nothing1 
         Caption         =   "nothing1"
      End
   End
End
Attribute VB_Name = "Form27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cheque_Click()
If CheckRegister = "ExitSub" Then Exit Sub
Form27Show = 0
Me.Hide
Form23.Show
Unload Me
End Sub

Private Sub Flatbutton2_Click()
Deposit1st = InputBox("You currently have $" & Label2.Tag & " dollar in your hand." & vbCrLf & "How much do you want to deposit?")
If Deposit1st = "" Then Exit Sub
If left(Val(Label2.Tag) - Val(Deposit1st), 1) = "-" Then
MsgBox "Sorry, you don't have enough money."
Exit Sub
End If
Deposit2st = MsgBox("Deposit $" & Deposit1st & ", are you sure?", vbYesNo)
If Deposit2st = vbYes Then
Label1.Tag = Val(Label1.Tag) + Val(Deposit1st)
Label2.Tag = Val(Label2.Tag) - Val(Deposit1st)
SaveSetting "Digimon", "Profile", "Bank", Label1.Tag
SaveSetting "Digimon", "Profile", "Money", Label2.Tag
MsgBox "Deposit Success."
Update.Enabled = True
End If
End Sub
 
Private Sub FlatButton3_Click()
Withdraw1st = InputBox("You currently have $" & Label1.Tag & " dollar in you bank." & vbCrLf & "How much do you want to withdraw?")
If Withdraw1st = "" Then Exit Sub
If left(Val(Label1.Tag) - Val(Withdraw1st), 1) = "-" Then
MsgBox "Sorry, you don't have enough money."
Exit Sub
End If
Withdraw2st = MsgBox("Withdraw $" & Withdraw1st & ", are you sure?", vbYesNo)
If Withdraw2st = vbYes Then
Label1.Tag = Val(Label1.Tag) - Val(Withdraw1st)
Label2.Tag = Val(Label2.Tag) + Val(Withdraw1st)
SaveSetting "Digimon", "Profile", "Bank", Label1.Tag
SaveSetting "Digimon", "Profile", "Money", Label2.Tag
MsgBox "Withdraw Success."
Update.Enabled = True
End If
End Sub

Private Sub Form_Load()
Form27Show = 1
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Label1.Tag = GetSetting("Digimon", "Profile", "Bank")
Label2.Tag = GetSetting("Digimon", "Profile", "Money")
Label1.Caption = "Bank cash: " & Label1.Tag
Label2.Caption = "Your hand cash: " & Label2.Tag
Interest.Enabled = True
End Sub

Private Sub Interest_Timer()
BankInterestvalue = GetSetting("Digimon", "Profile", "BankInterest")
SaveSetting "Digimon", "Profile", "BankInterest", 0
If BankInterestvalue = "0" Then Exit Sub
MsgBox "The Bank Give You An Interest of: " & BankInterestvalue
Update.Enabled = True
End Sub

Private Sub leave_Click()
Form27Show = 0
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Update_Timer()
Label1.Tag = GetSetting("Digimon", "Profile", "Bank")
Label2.Tag = GetSetting("Digimon", "Profile", "Money")
Label1.Caption = "Bank cash: " & Label1.Tag
Label2.Caption = "Your hand cash: " & Label2.Tag

BankInterestvalue = GetSetting("Digimon", "Profile", "BankInterest")
SaveSetting "Digimon", "Profile", "BankInterest", 0
If BankInterestvalue = "0" Then Exit Sub
MsgBox "The Bank Give You An Interest of: " & BankInterestvalue
Update.Enabled = False
End Sub
