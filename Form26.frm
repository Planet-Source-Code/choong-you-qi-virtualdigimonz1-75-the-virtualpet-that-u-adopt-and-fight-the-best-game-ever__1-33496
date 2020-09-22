VERSION 5.00
Begin VB.Form Form26 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casino DOUBLE!"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form26"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4260
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3600
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "--------------------"
      Top             =   360
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3600
      Top             =   480
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1800
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Play"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form26.frx":0000
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Answer:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Currently have: "
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WHAT?"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1080
      TabIndex        =   2
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2X "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   600
   End
   Begin VB.Menu Leave 
      Caption         =   "&Leave"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form26"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
leave.Enabled = False
leave.Tag = "1"
If FlatButton1.Tag = "1" Then Exit Sub
FlatButton1.Tag = "1"
FlatButton1.Enabled = False
Select Case GenerateRandom(1, 2)
Case "1" 'jackpot

Select Case Label2.Caption
Case "Money"
doublemoney = GetSetting("Digimon", "Profile", "Money")
doublemoney = doublemoney * Val(2)
SaveSetting "Digimon", "Profile", "Money", doublemoney

bankdouble = GetSetting("Digimon", "Profile", "Bank")
bankdouble = bankdouble * Val(2)
SaveSetting "Digimon", "Profile", "Bank", bankdouble

Timer2.Tag = "winmoney"
Case "Exp. Point"
doublepoint = GetSetting("Digimon", "Profile", "Score")
doublepoint = Val(doublepoint) + Val(doublepoint)
SaveSetting "Digimon", "Profile", "Score", doublepoint
Timer2.Tag = "winpoint"
End Select

Case "2" 'Lose
Select Case Label2.Caption
Case "Money"
SaveSetting "Digimon", "Profile", "Money", "0"
SaveSetting "Digimon", "Profile", "Bank", "0"
Timer2.Tag = "losemoney"
Case "Exp. Point"
SaveSetting "Digimon", "Profile", "Score", "0"
Timer2.Tag = "losepoint"
End Select
End Select
Timer2.Enabled = True
End Sub

Private Sub Command3_Click()
If Command3.Tag = "0" Then Exit Sub
Command3.Visible = False

End Sub

Private Sub leave_Click()
If leave.Tag = "1" Then Exit Sub
SaveSetting "Digimon", "Profile", "DoubleType", "none"
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
End Sub

Private Sub Timer1_Timer()
Select Case Me.Tag
Case "money"
Label2.Caption = "Money"
Label3.Caption = "Currently have: $" & GetSetting("Digimon", "Profile", "Money") & ", Bank: $" & GetSetting("Digimon", "Profile", "Bank")
Case "point"
Label2.Caption = "Exp. Point"
Label3.Caption = "Currently have: " & GetSetting("Digimon", "Profile", "Score") & " Exp. points."
Case Else
MsgBox "If You Can Enter Here Now, This is a bug." & vbCrLf & "please contact the author for the bug."
Unload Me
Form2.Show
End Select
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Text1.Text = Mid(Text1.Text, 2)
If Text1.Text = "" Then
Select Case Timer2.Tag
Case "winmoney"
Text1.Text = "MONEY X 2"
Label3.Caption = "Currently have: $" & GetSetting("Digimon", "Profile", "Money") & ", Bank: $" & GetSetting("Digimon", "Profile", "Bank")
Case "losemoney"
Text1.Text = "O.. my money!..."
Label3.Caption = "Currently have: $" & GetSetting("Digimon", "Profile", "Money") & ", Bank: $" & GetSetting("Digimon", "Profile", "Bank")
Case "winpoint"
Text1.Text = "Exp. Point X 2"
Label3.Caption = "Currently have: " & GetSetting("Digimon", "Profile", "Score") & " Exp. points."
Case "losepoint"
Text1.Text = "Oh.. no!..."
Label3.Caption = "Currently have: " & GetSetting("Digimon", "Profile", "Score") & " Exp. points."
End Select
Timer1.Enabled = True
Timer2.Enabled = False
leave.Enabled = True
leave.Tag = "0"
SaveSetting "Digimon", "Profile", "DoubleType", "none"
End If
End Sub
