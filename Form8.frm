VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Foods"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Close"
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
      Interval        =   1
      Left            =   960
      Top             =   2400
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   840
      Picture         =   "Form8.frx":0000
      ToolTipText     =   "$75"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "Form8.frx":08CA
      ToolTipText     =   "$50"
      Top             =   1800
      Width           =   480
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   840
      Picture         =   "Form8.frx":1194
      ToolTipText     =   "$35"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "Form8.frx":1A5E
      ToolTipText     =   "$25"
      Top             =   1200
      Width           =   480
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   840
      Picture         =   "Form8.frx":2328
      ToolTipText     =   "$15"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "Form8.frx":2BF2
      ToolTipText     =   "$5"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   840
      Picture         =   "Form8.frx":34BC
      ToolTipText     =   "$2"
      Top             =   0
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   120
      Picture         =   "Form8.frx":3D86
      ToolTipText     =   "$1"
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
Unload Me
Form2.SetFocus
End Sub

Private Sub Form_Load()
AutoDetectTop Me
Set m_cN2 = New cNeoCaption
Skin2 Me, m_cN2
Me.left = Form2.left - Me.Width
Me.top = Form2.top
Me.Height = Form2.Height
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image5.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 1
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.BorderStyle = 1
End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.BorderStyle = 1
End Sub
Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.BorderStyle = 1
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 1
End Sub
Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.BorderStyle = 1
End Sub
Private Sub Image1_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 1) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image1.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "1"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image2_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 2) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image2.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "2"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image3_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 5) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image3.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "3"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image4_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 15) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image4.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "5"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image5_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 25) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image5.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "10"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image6_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 35) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image6.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "15"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image7_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 50) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image7.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "20"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub
Private Sub Image8_Click()
If Not Form2.Timer3.Tag = "0" Then
MsgBox "Please wait your digimon finish its food first."
Exit Sub
End If
If WillMoney("MINUS", 75) = "ExitSub" Then Exit Sub
Form2.Food_Image.Picture = Image8.Picture
Form2.Timer3.Enabled = True
Form2.Timer3.Tag = "30"
FlatButton1.ToolTipText = "$" & GetSetting("Digimon", "Profile", "Money")
Unload Me
End Sub

Private Sub Timer1_Timer()
Form8.left = Form2.left - Form8.Width
Form8.top = Form2.top
End Sub
