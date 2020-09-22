VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form19 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "NETplay Battle Arena"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton2 
      Height          =   1060
      Left            =   -10000
      TabIndex        =   15
      Top             =   1440
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   1879
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Back"
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
      Height          =   375
      Left            =   3600
      TabIndex        =   14
      Top             =   1560
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Exit"
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
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      Height          =   855
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer Update 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3720
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   4200
      Tag             =   "0"
      Top             =   1560
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Start"
      Height          =   375
      Left            =   -3120
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3240
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   2
      Tag             =   "0"
      Text            =   "0"
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      Height          =   855
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3240
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Max             =   25
      Scrolling       =   1
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Transfering and Recieving Data and Information..."
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
      TabIndex        =   13
      Top             =   2040
      Width           =   4290
   End
   Begin VB.Label HEALTH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Label DEFENCE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defence"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   10
      Top             =   1080
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label POWER 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Power"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press: Alt+S to start Battle!"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   1680
      Width           =   1875
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   3360
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   840
      Tag             =   "0"
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   2520
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image11 
      Height          =   495
      Left            =   1560
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   3960
      Top             =   840
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDamage As Long
Dim OpponentDamage As Long
Dim Win As Boolean
Private Sub Command1_Click()
Timer3.Enabled = True
Image18.Visible = True
End Sub

Private Sub Command2_Click()
Timer2.Enabled = True
Image17.Visible = True
End Sub

Private Sub FlatButton2_Click()
Winsock1.close
Set m_cN = New cNeoCaption
Skin Form2, m_cN
Form2.Show
Unload Me
End Sub

Private Sub Command4_Click()
If Not Label5.Caption = "Status: Your Turn" Then Exit Sub
FormBattleShow = "1"
Text1.BackColor = vbYellow
Text1.ForeColor = vbBlack
Timer1.Tag = "1"
Timer1.Interval = GenerateRandom(1000, 3200)
Timer1.Enabled = True
Text1.SetFocus
End Sub

Private Sub FlatButton1_Click()
Winsock1.close
Set m_cN = New cNeoCaption
Skin Form2, m_cN
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
If Not GetSetting("Digimon", "Digimon", "Setting1-1") = "" Then Me.Picture = LoadPicture(GetSetting("Digimon", "Digimon", "Setting1-1"))
AutoDetectTop Me
Image17.Picture = Form7.Image39.Picture
Image18.Picture = Form7.Image38.Picture

Set m_cN3 = New cNeoCaption
Skin Me, m_cN3
'#####START YOURSELF SETTING#####'
Image1.Picture = Form2.Screen_Mon.Picture
Select Case GetSetting("Digimon", "Profile", "Picture")
Case "1"
Image11.Picture = Form6.Image1.Picture
Case "2"
Image11.Picture = Form6.Image2.Picture
Case "3"
Image11.Picture = Form6.Image3.Picture
Case "4"
Image11.Picture = Form6.Image4.Picture
Case "5"
Image11.Picture = Form6.Image5.Picture
Case "6"
Image11.Picture = Form6.Image6.Picture
Case "7"
Image11.Picture = Form6.Image7.Picture
Case "8"
Image11.Picture = Form6.Image8.Picture
End Select
'#####END YOURSELF SETTING#####'
Label1.Caption = GetSetting("Digimon", "Profile", "Name")
ProgressBar1.Max = GetSetting("Digimon", "Digimon", "health")
ProgressBar1.Value = GetSetting("Digimon", "Digimon", "Currenthealth")
Win = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Form14
End Sub

Private Sub Image1_Click()
Form14.Show
Form14.Image1.Picture = Image1.Picture
Form14.Label1.Caption = left(Form14.Label1.Caption, 6) & GetSetting("Digimon", "Digimon", "name")
Form14.Label2.Caption = left(Form14.Label2.Caption, 8) & GetSetting("Digimon", "Digimon", "health")
Form14.Label3.Caption = left(Form14.Label3.Caption, 7) & GetSetting("Digimon", "Digimon", "Power")
Form14.Label4.Caption = left(Form14.Label4.Caption, 9) & GetSetting("Digimon", "Digimon", "defence")
Form13.Enabled = False
End Sub

Private Sub Image2_Click()
Form14.Show
Form14.Image1.Picture = Image2.Picture
Form14.Label1.Caption = left(Form14.Label1.Caption, 6) & Image2.Tag
Form14.Label2.Caption = left(Form14.Label2.Caption, 8) & ProgressBar2.Max
Form14.Label3.Caption = left(Form14.Label3.Caption, 7) & POWER.Caption
Form14.Label4.Caption = left(Form14.Label4.Caption, 9) & DEFENCE.Caption
Form13.Enabled = False
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If Timer1.Tag = "0" Then Exit Sub
If KeyCode = vbKeyLeft Then
If Text1.Tag = "0" Then Exit Sub
Text1.Tag = "0"
Text1.Text = Val(Text1.Text) + Val(5)
End If
If KeyCode = vbKeyRight Then
If Text1.Tag = "1" Then Exit Sub
Text1.Tag = "1"
End If
End Sub
Private Sub Timer1_Timer()
FormBattleShow = "0"
Text1.BackColor = vbBlack
Text1.ForeColor = vbWhite
Image17.Tag = Text1.Text
firstdamage = Val(Image17.Tag) / Val(100) * Val(GetSetting("Digimon", "Digimon", "Power"))

comreduce = GenerateRandom(5, 10)
DC = Val(DEFENCE.Caption * Val(1) / Val(comreduce))
MyDamage = Val(firstdamage - DC)

'MsgBox "firstdamage = " & firstdamage
'MsgBox "comreduce = " & comreduce
'MsgBox "DC = " & DC
'MsgBox "MyDamage = " & MyDamage


If Not DC > firstdamage Then
If left(MinusComBlood, 1) = "-" Then MyDamage = "0"
Else
'MsgBox "MyDamage = " & MyDamage
MyDamage = "0"
End If
Winsock1.SendData "H" & MyDamage
Label5.Caption = "Status: Battling"
Command2.Value = True
Command1.Value = True
Text1.Text = "0"
Timer1.Tag = "0"
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
If Image17.left = "3360" Then
Dim MinusComBlood As Long
MinusComBlood = Val(ProgressBar2.Value) - MyDamage
If left(MinusComBlood, 1) = "-" Then MinusComBlood = "0"
ProgressBar2.Value = MinusComBlood
Image17.Visible = False
Image17.left = "840"
Timer2.Enabled = False
Update.Enabled = True
End If
Image17.left = Image17.left + 120
End Sub

Private Sub Timer3_Timer()
If Image18.left = "840" Then
Dim MinusPlayerBlood As Long
MinusPlayerBlood = Val(ProgressBar1.Value) - OpponentDamage
If left(MinusPlayerBlood, 1) = "-" Then MinusPlayerBlood = "0"
ProgressBar1.Value = MinusPlayerBlood
Image18.Visible = False
Image18.left = "3360"
Timer3.Enabled = False
End If
Image18.left = Image18.left - 120
End Sub

Private Sub Update_Timer()
SaveSetting "Digimon", "Digimon", "CurrentHealth", ProgressBar1.Value
Dim PLUSLose As Long
Dim PLUSDraw As Long
Dim PLUSwin As Long
ProgressBar1.ToolTipText = ProgressBar1.Value & "/" & GetSetting("Digimon", "Digimon", "health")
ProgressBar2.ToolTipText = ProgressBar2.Value & "/" & ProgressBar2.Max
If ProgressBar2.Value = "0" Then
If ProgressBar1.Value = "0" Then
WillPoint "PLUS", 2
PLUSDraw = GetSetting("Digimon", "Digimon", "Draw")
PLUSDraw = PLUSDraw + Val(1)
SaveSetting "Digimon", "Digimon", "Draw", PLUSDraw
'Form2.Show
'Unload Me
FlatButton2.left = "0"
Win = True
Update.Enabled = False
WinLoseCode = "Draw"
MsgBox "Draw"
Exit Sub
End If
WillPoint "PLUS", 3
PLUSwin = GetSetting("Digimon", "Digimon", "Win")
PLUSwin = PLUSwin + Val(1)
SaveSetting "Digimon", "Digimon", "Win", PLUSwin
'Form2.Show
'Unload Me
FlatButton2.left = "0"
Win = True
Update.Enabled = False
WinLoseCode = "Win"
MsgBox "You Win"
Exit Sub
End If

If ProgressBar1.Value = "0" Then
If ProgressBar2.Value = "0" Then
PLUSDraw = GetSetting("Digimon", "Digimon", "Draw")
PLUSDraw = PLUSDraw + Val(1)
SaveSetting "Digimon", "Digimon", "Draw", PLUSDraw
WillPoint "PLUS", 2
'Form2.Show
'Unload Me
FlatButton2.left = "0"
Win = True
Update.Enabled = False
WinLoseCode = "Draw"
MsgBox "Draw"
Exit Sub
End If
WillPoint "PLUS", 1
PLUSLose = GetSetting("Digimon", "Digimon", "Lose")
PLUSLose = PLUSLose + Val(1)
SaveSetting "Digimon", "Digimon", "Lose", PLUSLose
'Form2.Show
'Unload Me
FlatButton2.left = "0"
Win = True
Update.Enabled = False
WinLoseCode = "Lose"
MsgBox "You Lose"
Exit Sub
End If
Update.Enabled = False
Label5.Caption = "Opponent Turn."
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData, strData2 As String
Call Winsock1.GetData(strData, vbString)
strData2 = left(strData, 1)
strData = Mid(strData, 2)

Select Case strData2
Case "A"
Unload Form5
Me.Show

Select Case Me.Tag
Case "Form2"
Form2.Hide
Case "Form36"
Form2.Hide
Form36.Hide
End Select





Form2.Hide
Label3.Caption = strData

If left(Val(Form2.ProgressBar1.Value) - Val(10), 1) = "-" Then
Winsock1.SendData "I"
Win = True
MsgBox "You don't have enough energy to battle." & vbCrLf & "Try again after your pet have a rest."
Winsock1.close
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN
Form2.Show
Unload Me
Exit Sub
End If
WillEnergy "MINUS", 10

Winsock1.SendData "A" & GetSetting("Digimon", "Profile", "Name")
Case "B"
Select Case strData
Case "1"
Image12.Picture = Form6.Image1.Picture
Case "2"
Image12.Picture = Form6.Image2.Picture
Case "3"
Image12.Picture = Form6.Image3.Picture
Case "4"
Image12.Picture = Form6.Image4.Picture
Case "5"
Image12.Picture = Form6.Image5.Picture
Case "6"
Image12.Picture = Form6.Image6.Picture
Case "7"
Image12.Picture = Form6.Image7.Picture
Case "8"
Image12.Picture = Form6.Image8.Picture
End Select
Winsock1.SendData "B" & GetSetting("Digimon", "Profile", "Picture")
Case "C"
Image2.Tag = strData
Winsock1.SendData "C" & GetSetting("Digimon", "Digimon", "name")
Case "D"
HEALTH.Caption = strData
ProgressBar2.Max = HEALTH.Caption
ProgressBar2.Value = ProgressBar2.Max
Winsock1.SendData "D" & GetSetting("Digimon", "Digimon", "health")
Case "E"
POWER.Caption = strData
Winsock1.SendData "E" & GetSetting("Digimon", "Digimon", "Power")
Case "F"
DEFENCE.Caption = strData
Winsock1.SendData "F" & GetSetting("Digimon", "Digimon", "defence")
Case "G"
DigimonType3 strData
Winsock1.SendData "G" & GetSetting("Digimon", "Digimon", "Type")
Case "H"
OpponentDamage = strData
Label5.Caption = "Status: Your Turn"
Case "I"
Win = True
MsgBox "Opponent Don't have enough energy to battle."
Unload Form5
Winsock1.close
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN
Form2.Show
Unload Me
Case "J"
ProgressBar2.Value = strData
Winsock1.SendData "J" & GetSetting("Digimon", "Digimon", "Currenthealth")
Label5.Caption = "Status: Opponent Turn."
FlatButton1.Visible = False
Case "K"
Case "L"
Case "M"
End Select
End Sub
Private Sub Winsock1_Close()
If Win = True Then Exit Sub
MsgBox "Connection Lost"
Form2.Show
Unload Me
End Sub
