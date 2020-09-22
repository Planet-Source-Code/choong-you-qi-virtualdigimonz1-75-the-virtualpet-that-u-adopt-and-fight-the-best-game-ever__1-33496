VERSION 5.00
Begin VB.Form Form23 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque"
   ClientHeight    =   2670
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5625
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form23"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Cheque: "
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         MaxLength       =   12
         TabIndex        =   6
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   480
         MaxLength       =   8
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
      Begin VirtualDigimonz.FlatButton FlatButton1 
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "Send"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BANK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   4080
         TabIndex        =   10
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual Digimonz"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   480
         Width           =   1380
      End
      Begin VB.Shape shpRect 
         BorderColor     =   &H00400000&
         BorderWidth     =   16
         Height          =   615
         Index           =   1
         Left            =   3840
         Shape           =   2  'Oval
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Example: (A-543B-234-C)"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2160
         TabIndex        =   8
         Top             =   1440
         Width           =   1770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "$"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   840
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amout: (Your Bank only have:)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2160
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque Number: (May content: 1234567890-ABC)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   3570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
   Begin VB.Menu recieve 
      Caption         =   "&Recieve"
      Begin VB.Menu nothing1 
         Caption         =   "nothing1"
      End
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub FlatButton1_Click()
If Text1.Text = "" Then Exit Sub
If Text2.Text = "" Then Exit Sub
Dim AfterBankTax As Long
AfterBankTax = Val(Text1) / Val(2)
confirmsending = MsgBox("Bank will charge 50% tax." & vbCrLf & "Your money need to transfer: $" & Text1 & vbCrLf & "After charging 50% tax, The money will be tranfer is: $" & AfterBankTax & vbCrLf & "Are you sure you want to continue?", vbYesNo)
If confirmsending = vbNo Then Exit Sub

If WillBank("MINUS", Text1) = "ExitSub" Then Exit Sub
Me.Hide
Form16.Show
AllPage = GetUrlSource("http://www2.domaindlx.com/choongyouqi/cheque/add_vb.asp?Amout=" & Text1.Text & "&Number=" & Text2.Text)
If Not AllPage = "" Then
If Not CheckDatabaseActive(Me) = "Activation" Then GoTo skiptonoactiveform23
MsgBox "Cheque Sent"
Text1 = ""
Text2 = ""
Else
MsgBox "Please Connect To Internet 1st."

skiptonoactiveform23:

WillBank "PLUS", Text1
Text1 = ""
End If
Unload Form16
Me.Show
Label3.Caption = "Amout: (Your Bank have: $" & GetSetting("Digimon", "Profile", "Bank") & ")"
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
Label1.Caption = "From: " & GetSetting("Digimon", "Profile", "Name")
Label3.Caption = "Amout: (Your Bank have: $" & GetSetting("Digimon", "Profile", "Bank") & ")"
End Sub

Private Sub recieve_Click()
'Recieve Cheque Code
ConfirmNumber = InputBox("Please Enter Your Cheque Number:(Case Sensitive)")
If ConfirmNumber = "" Then Exit Sub
Me.Hide
Form16.Show
AllPage = GetUrlSource("http://www2.domaindlx.com/choongyouqi/cheque/recieve_vb.asp?Number=" & ConfirmNumber)
If Not AllPage = "" Then
Do Until Len(AllPage) = 0
If left(AllPage, 5) = "<Con>" Then
AllPage = Mid(AllPage, 6)
Text1.Tag = ""
Do Until left(AllPage, 6) = "</Con>"
Text1.Tag = Text1.Tag & left(AllPage, 1)
AllPage = Mid(AllPage, 2)
Loop
Unload Form16
Me.Show
WillMoney "PLUS", Text1.Tag
MsgBox "Recieved Cash: $" & Text1.Tag
Text1.Tag = ""
Exit Sub
End If
AllPage = Mid(AllPage, 2)
Loop
End If
Unload Form16
Me.Show
MsgBox "No Cheque In This Number"
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

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim KeySetLock As String
Dim ch As String * 1
KeySetLock = "1234567890-ABC" & vbBack
ch = Chr$(KeyAscii)

If InStr(KeySetLock, ch) Then
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Else
    KeyAscii = 0
    Exit Sub
End If
End Sub
