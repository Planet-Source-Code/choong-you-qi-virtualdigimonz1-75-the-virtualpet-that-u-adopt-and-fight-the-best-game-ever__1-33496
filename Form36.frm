VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form36 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Online Tournament"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form36"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2400
      Top             =   3600
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   1815
      Left            =   240
      TabIndex        =   22
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   3201
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Start Battle"
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
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1800
      Top             =   3600
   End
   Begin VB.Timer Checking_Timer 
      Left            =   480
      Top             =   3600
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   840
      Left            =   1800
      TabIndex        =   20
      Top             =   2640
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   2760
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   1815
      ItemData        =   "Form36.frx":0000
      Left            =   240
      List            =   "Form36.frx":0002
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1080
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VirtualDigimonz.FlatButton FlatButton2 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&Connect"
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
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0/8"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1320
      TabIndex        =   21
      Top             =   2160
      Width           =   255
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   4680
      X2              =   4800
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   1320
      Y2              =   2040
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   3720
      X2              =   3840
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2760
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   3720
      X2              =   3720
      Y1              =   720
      Y2              =   1920
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   480
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   3600
      X2              =   3720
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   3600
      X2              =   3720
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   2
      X1              =   2520
      X2              =   2640
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHAMPION"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      TabIndex        =   19
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      TabIndex        =   18
      Top             =   1320
      Width           =   525
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3960
      TabIndex        =   17
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player4"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   16
      Top             =   1920
      Width           =   525
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player3"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   15
      Top             =   1680
      Width           =   525
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   14
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2880
      TabIndex        =   13
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player8"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   12
      Top             =   2280
      Width           =   525
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player7"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   11
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player6"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   10
      Top             =   1680
      Width           =   525
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player5"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player4"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player3"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player2"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player1"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1800
      TabIndex        =   5
      Top             =   240
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Server IP:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Player Connected:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1305
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form36"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public TEMPOpponentIP
'Public TEMPOpponentName

Private Sub Checking_Timer_Timer()
Label18.Caption = List1.ListCount & "/8"
End Sub

Private Sub close_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN
Form2.Show
Unload Me

'End
End Sub

Private Sub FlatButton1_Click()
Select Case FlatButton1.Tag
Case "Form13"
Form13.Show
Me.Hide
FlatButton1.Visible = False
Case "Form17"
Case "Form19"
Form19.Winsock1.close
Form19.Winsock1.Connect OpponentIP, 333
FlatButton1.Enabled = False
Timer2.Enabled = True
Case Else
MsgBox "If you see this message, then this is an error." & vbCrLf & "Please contact the author for the error."
End Select
End Sub

Private Sub FlatButton2_Click()
If Text1.Text = "" Then MsgBox "Please Type The Server IP": Exit Sub
List1.Clear
Winsock1.close
Winsock1.Connect Text1.Text, 334
FlatButton2.Enabled = False
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
AutoDetectTop Me
Top8OR1 = 0
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
List2.AddItem "Status:"
End Sub

Private Sub Text1_Click()
Text1.Text = "127.0.0.1"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim KeySetLock As String
Dim ch As String * 1
KeySetLock = "1234567890." & vbBack
ch = Chr$(KeyAscii)

If InStr(KeySetLock, ch) Then
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Else
    KeyAscii = 0
    Exit Sub
End If
End Sub

Private Sub Timer1_Timer()
If List2.ListCount = 1 Then
Winsock1.close
MsgBox "Connection Time Out."
FlatButton2.Enabled = True
End If
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
FlatButton1.Enabled = True
Timer2.Enabled = False
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData, strData2 As String
Call Winsock1.GetData(strData, vbString)
strData2 = left(strData, 1)
strData = Mid(strData, 2)

Select Case strData2

Case "A"
AddEvent "Connected."
Winsock1.SendData "A" & Winsock1.LocalIP
Case "B"
AddEvent "Getting Information."
abc = GenerateRandom(1, 67676)
Winsock1.SendData "B" & GetSetting("Digimon", "Profile", "Name") 'abc
Me.Caption = Me.Caption & " -- " & abc
Case "C"
Winsock1.SendData "C"
Case "D"
PlayerNameTournament = strData
Do Until PlayerNameTournament = ""
If left(PlayerNameTournament, 1) = Chr(1) Then
List1.AddItem TempNameTournament
TempNameTournament = ""
PlayerNameTournament = Mid(PlayerNameTournament, 2)
End If
TempNameTournament = TempNameTournament & left(PlayerNameTournament, 1)
PlayerNameTournament = Mid(PlayerNameTournament, 2)
Loop
AddEvent "Done."
Label18.Caption = List1.ListCount & "/8"
Case "E"
For Y = 0 To List1.ListCount - 1
If List1.List(Y) = strData Then
List1.RemoveItem Y 'Error Place
AddEvent strData & " Leaved."
Label18.Caption = List1.ListCount & "/8"
End If
Next Y
Case "F"
List1.AddItem strData
AddEvent strData & " Joined."
Label18.Caption = List1.ListCount & "/8"





'YourPOST
'1= 1,3,5,7  |  2= 2,4,6,8
'''''''''''''START YOUR OPPONENT DATA'''''''''''''''''''
Case "G"
ArrangeOpponentData (strData)

Select Case YourPOST
Case "1"
If CheckVersusBot = "True" Then Exit Sub
Form17.Tag = "Form36"
Form17.Winsock1.close
Form17.Winsock1.LocalPort = CLng(333)
Form17.Winsock1.Listen
AddEvent "Game Start(You Are Server)."
AddEvent "OpponentName:" & OpponentName
AddEvent "OpponentIP:" & OpponentIP
MsgBox "Tournament Begin!" & vbCrLf & "OpponentName:" & OpponentName & vbCrLf & "OpponentIP:" & OpponentIP & vbCrLf & "Please wait for you opponent join you."
FlatButton1.Tag = "Form17"
FlatButton1.Visible = True
FlatButton1.Enabled = False

Case "2"
If CheckVersusBot = "True" Then Exit Sub

MsgBox "Tournament Begin!" & vbCrLf & "OpponentName:" & OpponentName & vbCrLf & "OpponentIP:" & OpponentIP & vbCrLf & "Press OK to join your opponent."
Form19.Tag = "Form36"
Form19.Winsock1.close
Form19.Winsock1.Connect OpponentIP, 333
FlatButton1.Tag = "Form19"
FlatButton1.Visible = True
FlatButton1.Enabled = True
End Select

'''''''''''''END YOUR OPPONENT DATA'''''''''''''''''''


Case "H"
Case "M"
Winsock1.close
WillMoney "PLUS", 100000
MsgBox "Your are the Tournament Champion." & "You recieved cash $100,000"
close_Click
Case "V"

TEMPstrdata = Mid(strData, 2)
TEMPdata = ""
Select Case left(strData, 1)
Case 8
GraphUpdate = 1
Case 4
GraphUpdate = 9
Case 2
GraphUpdate = 13
Case 1
GraphUpdate = 15
End Select

Do Until TEMPstrdata = ""
If left(TEMPstrdata, 1) = Chr(1) Then
MakeGraph (TEMPdata)
TEMPdata = ""
TEMPstrdata = Mid(TEMPstrdata, 2)
End If
TEMPdata = TEMPdata & left(TEMPstrdata, 1)
TEMPstrdata = Mid(TEMPstrdata, 2)
Loop


Case "W"
Select Case left(strData, 1)
Case "W"
AddEvent Mid(strData, 2) & " Win."
Case "L"
AddEvent Mid(strData, 2) & " Lose."
Case "D"
AddEvent Mid(strData, 2) & " Draw."
End Select
Winsock1.SendData "Y"
Case "Y"
Winsock1.close
MsgBox "You're been Kicked by Server."
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN
Form2.Show
Unload Me

Case "Z"
Winsock1.close
Timer1.Enabled = False
MsgBox "Sorry. Server Max Connection." & vbCrLf & "Try again later or connect to other server."
FlatButton2.Enabled = True
End Select
End Sub
