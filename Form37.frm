VERSION 5.00
Begin VB.Form Form37 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lottery"
   ClientHeight    =   3780
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form37"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Selling @ $50 each"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   240
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "Your Number..."
         Top             =   720
         Width           =   1575
      End
      Begin VirtualDigimonz.FlatButton FlatButton1 
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Buy"
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
         BackStyle       =   0  'Transparent
         Caption         =   "Please choose your lucky number:(1-200)"
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
         Height          =   435
         Left            =   80
         TabIndex        =   7
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Your Lott.Ticket"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1620
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Last Lottery Result: The Number is..."
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "XXX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   3135
      End
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "You Have:"
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
      Left            =   1680
      TabIndex        =   9
      Top             =   2760
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Next OpenTime:"
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
      Left            =   1680
      TabIndex        =   6
      Top             =   3000
      Width           =   1380
   End
   Begin VB.Menu leave 
      Caption         =   "&Leave"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form37"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
If Text1 = "Your Number..." Then Exit Sub
If List1.ListCount = 5 Then
MsgBox "Your already reach the maximum amout of Lottery Ticket.(max:5)" & vbCrLf & "Please try again next time."
Exit Sub
End If
If Val(Text1) < 1 Then
MsgBox "Please choose a legal number."
Exit Sub
End If
If Val(Text1) > 200 Then
MsgBox "Please choose a legal number."
Exit Sub
End If
areyousurebuyticket = MsgBox("Are you sure you want to buy number " & Val(Text1) & "?", vbYesNo)
If areyousurebuyticket = vbNo Then Exit Sub
WillMoney "MINUS", 50
Select Case List1.ListCount
Case "0"
SaveSetting "Digimon", "Lottery", "Lottery1", Val(Text1)
Case "1"
SaveSetting "Digimon", "Lottery", "Lottery2", Val(Text1)
Case "2"
SaveSetting "Digimon", "Lottery", "Lottery3", Val(Text1)
Case "3"
SaveSetting "Digimon", "Lottery", "Lottery4", Val(Text1)
Case "4"
SaveSetting "Digimon", "Lottery", "Lottery5", Val(Text1)
End Select
LotteryTicketList
Label4.Caption = "You have: $" & GetSetting("Digimon", "Profile", "Money")
Text1.Text = ""
Text1.SetFocus
End Sub

Private Sub Form_Load()
Form37Show = 1
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
Label2.Caption = "Next OpenTime: " & GetSetting("Digimon", "Lottery", "LotteryTime") & "min"
Label1.Caption = GetSetting("Digimon", "Lottery", "LotteryLastNum")
Label4.Caption = "You have: $" & GetSetting("Digimon", "Profile", "Money")
LotteryTicketList
End Sub

Private Sub leave_Click()
Form37Show = 0
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub


Private Sub Text1_Click()
If Text1.Tag = "" Then
Text1.Text = ""
Text1.Tag = "1"
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim KeySetLock As String
Dim ch As String * 1
KeySetLock = "1234567890" & vbBack
ch = Chr$(KeyAscii)

If InStr(KeySetLock, ch) Then
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Else
    KeyAscii = 0
    Exit Sub
End If
End Sub

