VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battle"
   ClientHeight    =   2145
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   5475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Connection"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2400
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H0000FF00&
         Height          =   315
         ItemData        =   "Form5.frx":0000
         Left            =   120
         List            =   "Form5.frx":0007
         TabIndex        =   11
         Text            =   "IP Address Here!"
         Top             =   480
         Width           =   1575
      End
      Begin VirtualDigimonz.FlatButton FlatButton4 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "C&onnect"
         Enabled         =   0   'False
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
      Begin VirtualDigimonz.FlatButton FlatButton1 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&JOIN"
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
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&HOST"
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Challenge:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status: (None)"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Profile"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VirtualDigimonz.FlatButton FlatButton3 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "Copy IP &address"
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
         Caption         =   "IP: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Click Here to copy the IP address."
         Top             =   840
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   510
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   120
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   480
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   840
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub close_Click()
Me.Hide
If Form5.Tag = "battle" Then Form17.Winsock1.close
Form2.Show
Unload Me
End Sub

Private Sub Combo1_GotFocus()
If Not Combo1.Tag = "1" Then
Combo1 = ""
Combo1.Tag = "1"
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
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

Private Sub FlatButton1_Click()
Combo1.Enabled = True
Combo1.SetFocus
FlatButton4.Enabled = True
FlatButton1.Enabled = False
FlatButton2.Enabled = False
Label5.Caption = "Status: Join."
End Sub

Private Sub FlatButton2_Click()
Select Case Form5.Tag
Case "battle"
Form17.Tag = "Form2"
Form17.Winsock1.close
Form17.Winsock1.LocalPort = CLng(333)
Form17.Winsock1.Listen
Case "reset"
Form5.Winsock1.close
Form5.Winsock1.LocalPort = CLng(335)
Form5.Winsock1.Listen
End Select
FlatButton1.Enabled = False
FlatButton2.Enabled = False
Label5.Caption = "Status: Host."
End Sub

Private Sub FlatButton3_Click()
Clipboard.SetText Winsock1.LocalIP
End Sub

Private Sub Flatbutton4_Click()
If Len(Combo1) > 15 Then Exit Sub
If Combo1 = "" Then Exit Sub
If Combo1.Text = "127.0.0.1" Then
MsgBox "You Can't Connect To The Same Computer."
Combo1.SetFocus
Combo1.SelStart = 0
Combo1.SelLength = Len(Combo1)
Exit Sub
End If

ManageIP "SAVE", Combo1
MemoryCombo = Combo1
ManageIP "GET", Nothing
Combo1 = MemoryCombo

Select Case Form5.Tag
Case "battle"
Form19.Tag = "Form2"
Form19.Winsock1.close
Form19.Winsock1.Connect Combo1.Text, 333
Case "reset"
Form5.Winsock1.close
Form5.Winsock1.Connect Combo1.Text, 335
End Select

End Sub

Private Sub Form_Load()
ManageIP "GET", Nothing

AutoDetectTop Me
Set m_cN2 = New cNeoCaption
Skin Me, m_cN2

Select Case GetSetting("Digimon", "Profile", "Picture")
Case "1"
Image1.Picture = Form6.Image1.Picture
Case "2"
Image1.Picture = Form6.Image2.Picture
Case "3"
Image1.Picture = Form6.Image3.Picture
Case "4"
Image1.Picture = Form6.Image4.Picture
Case "5"
Image1.Picture = Form6.Image5.Picture
Case "6"
Image1.Picture = Form6.Image6.Picture
Case "7"
Image1.Picture = Form6.Image7.Picture
Case "8"
Image1.Picture = Form6.Image8.Picture
End Select
Label1.Caption = Label1.Caption & GetSetting("Digimon", "Profile", "Name")
Label2.Caption = Label2.Caption & GetSetting("Digimon", "Profile", "Ranked")
Label3.Caption = Label3.Caption & Winsock1.LocalIP
Form5.left = Form2.left + Form2.Width
Form5.top = Form2.top
End Sub

Private Sub Label3_Click()
Clipboard.SetText Winsock1.LocalIP
End Sub

Private Sub Winsock1_Close()
If Winsock1.Tag = "1" Then Exit Sub
MsgBox "Connection Lose"
End Sub

Private Sub winsock1_ConnectionRequest(ByVal requestID As Long)
If Winsock1.State <> sckClosed Then Winsock1.close
Winsock1.Accept requestID
Winsock1.SendData "A"
Winsock1.Tag = "1"
Reset_Program
MsgBox "Reseted Together."
Me.Hide
Unload Form39
UnloadSubForm
m_cN.Detach
Unload Form2
Unload Me
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strData, strData2 As String
Call Winsock1.GetData(strData, vbString)
strData2 = left(strData, 1)
strData = Mid(strData, 2)
If strData2 = "A" Then
Winsock1.Tag = "1"
Reset_Program
MsgBox "Reseted Together."
Me.Hide
Unload Form39
UnloadSubForm
m_cN.Detach
Unload Form2
Unload Me
End If
End Sub
