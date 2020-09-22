VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form29 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Training Center"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form29"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4095
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Tag             =   "0"
      Top             =   1680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Training for: "
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
      Begin VirtualDigimonz.FlatButton FlatButton1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Attack"
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
         TabIndex        =   11
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Defence"
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
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Health"
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   1935
      Left            =   3600
      TabIndex        =   9
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   3413
      _Version        =   393216
      Appearance      =   1
      Max             =   95
      Orientation     =   1
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Training: "
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   1560
      TabIndex        =   4
      Top             =   840
      Width           =   1935
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "Training Center"
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
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
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
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
         Height          =   495
         Left            =   720
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
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
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Start"
      Enabled         =   0   'False
      Height          =   375
      Left            =   -24000
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Status: "
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1935
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   570
      End
   End
   Begin VB.Menu leave 
      Caption         =   "&Leave"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WhichTime As Integer
Dim TrainingWhat As String

Private Sub FlatButton1_Click()
If Not WhichTime = "0" Then Exit Sub
Confirmvalue = MsgBox("This training course cost $2200, Are you sure?", vbYesNo)
If Confirmvalue = vbYes Then
If WillMoney("MINUS", 2200) = "ExitSub" Then Exit Sub
Label1.Caption = "Money: " & GetSetting("Digimon", "Profile", "Money")
WhichTime = "1"
TrainingWhat = "Attack"
Text1.Text = "Alt+S to Start!"
Text1.SetFocus
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
Command5.Enabled = True
End If
End Sub

Private Sub FlatButton2_Click()
If Not WhichTime = "0" Then Exit Sub
Confirmvalue = MsgBox("This training course cost $2000, Are you sure?", vbYesNo)
If Confirmvalue = vbYes Then
If WillMoney("MINUS", 2000) = "ExitSub" Then Exit Sub
Label1.Caption = "Money: " & GetSetting("Digimon", "Profile", "Money")
WhichTime = "1"
TrainingWhat = "Defence"
Text1.Text = "Alt+S to Start!"
Text1.SetFocus
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
Command5.Enabled = True
End If
End Sub

Private Sub FlatButton3_Click()
If Not WhichTime = "0" Then Exit Sub
Confirmvalue = MsgBox("This training course cost $1800, Are you sure?", vbYesNo)
If Confirmvalue = vbYes Then
If WillMoney("MINUS", 1800) = "ExitSub" Then Exit Sub
Label1.Caption = "Money: " & GetSetting("Digimon", "Profile", "Money")
WhichTime = "1"
TrainingWhat = "Health"
Text1.Text = "Alt+S to Start!"
Text1.SetFocus
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
Command5.Enabled = True
End If
End Sub


Private Sub Command5_Click()
If Timer1.Tag = "1" Then Exit Sub
Text1.BackColor = vbYellow
Text1.ForeColor = vbBlack
Timer1.Tag = "1"
Timer1.Interval = GenerateRandom(1000, 3200)
Timer1.Enabled = True
Text1.Text = "0"
Text1.SetFocus
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Label1.Caption = "Money: " & GetSetting("Digimon", "Profile", "Money")
WhichTime = "0"
End Sub

Private Sub leave_Click()
If WhichTime = "0" Then
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Else
ConfirmLeave = MsgBox("You are attending on a course." & vbCrLf & "Leaving will end the training. Are you sure?", vbYesNo)
If ConfirmLeave = vbNo Then Exit Sub
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Val(Text1.Text) + Val(5) > 95 Then Exit Sub
If Timer1.Tag = "0" Then Exit Sub
If KeyCode = vbKeyLeft Then
If Text1.Tag = "0" Then Exit Sub
Text1.Tag = "0"
Text1.Text = Val(Text1.Text) + Val(5)
ProgressBar1.Value = Text1.Text
End If
If KeyCode = vbKeyRight Then
If Text1.Tag = "1" Then Exit Sub
Text1.Tag = "1"
End If
End Sub

Private Sub Timer1_Timer()
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
Select Case WhichTime
Case "1"
WhichTime = "2"
Label3.Caption = Text1.Text
ProgressBar1.Value = 0
Case "2"
WhichTime = "3"
Label4.Caption = Text1.Text
ProgressBar1.Value = 0
Case "3"
Label5.Caption = Text1.Text
ProgressBar1.Value = 0
FlatButton1.Enabled = True
FlatButton2.Enabled = True
FlatButton3.Enabled = True
Command5.Enabled = False
Timer2.Enabled = True
End Select
Text1.Text = "0"
Timer1.Tag = "0"
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Dim WinValue As Long
sub3value = Val(Label3.Caption) + Val(Label4.Caption) + Val(Label5.Caption)
Select Case Val(sub3value)
Case 0 To 140
WinValue = 0
Case 145 To 185
WinValue = 1
Case 190 To 205
WinValue = 2
Case 210 To 225
WinValue = 3
Case 230 To 245
WinValue = 4
Case 250 To 260
WinValue = 5
Case 265 To 280
WinValue = 10
Case 285
MsgBox "Genius!"
WinValue = 50
End Select

Select Case TrainingWhat
Case "Attack"
WillAttack "PLUS", WinValue
Case "Defence"
WillDefence "PLUS", WinValue
Case "Health"
WillHealth "PLUS", WinValue
End Select

If Not WinValue = "0" Then
MsgBox "Training Result!" & vbCrLf & "Training Success." & vbCrLf & "Your pet gain " & WinValue & " " & TrainingWhat & " point."
Else
MsgBox "Training Result!" & vbCrLf & "Training Failed." & vbCrLf & "0o0o0o0o0..."
End If

Form2_Update
Text1.Text = "Training Center"
Timer2.Enabled = False
WhichTime = "0"
Label3.Caption = "?"
Label4.Caption = "?"
Label5.Caption = "?"
End Sub
