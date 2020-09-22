VERSION 5.00
Begin VB.Form Form15 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Battle Statistic"
   ClientHeight    =   3060
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Statistic: "
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Win: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lose: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   3
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Draw: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1080
         TabIndex        =   2
         Top             =   720
         Width           =   465
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   240
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Badge: "
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
      Begin VB.Image Image9 
         Height          =   480
         Left            =   3000
         Picture         =   "Form15.frx":0000
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   2040
         Picture         =   "Form15.frx":0CCA
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   1080
         Picture         =   "Form15.frx":1994
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   120
         Picture         =   "Form15.frx":265E
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   3000
         Picture         =   "Form15.frx":3328
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   2040
         Picture         =   "Form15.frx":3FF2
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1080
         Picture         =   "Form15.frx":4CBC
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "Form15.frx":5986
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   3840
      Picture         =   "Form15.frx":6650
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
   Begin VB.Menu OnlineStatistic 
      Caption         =   "&Online Ranking"
      Begin VB.Menu nothing1 
         Caption         =   "nothing1"
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub close_Click()
Select Case Form15.Tag
Case "Out"
Form12.Show
Unload Me

Case "Main"
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
Case ""
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Select
End Sub


Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Image1.Picture = Form2.Screen_Mon.Picture
Label1.Caption = left(Label1.Caption, 5) & GetSetting("Digimon", "Digimon", "Win")
Label2.Caption = left(Label2.Caption, 6) & GetSetting("Digimon", "Digimon", "Lose")
Label3.Caption = left(Label3.Caption, 6) & GetSetting("Digimon", "Digimon", "Draw")

Select Case GetSetting("Digimon", "Profile", "Badge")
Case "0"
Image2.Picture = Image14.Picture
Image3.Picture = Image14.Picture
Image4.Picture = Image14.Picture
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "1"
Image3.Picture = Image14.Picture
Image4.Picture = Image14.Picture
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "2"
Image4.Picture = Image14.Picture
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "3"
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "4"
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "5"
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "6"
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "7"
Image9.Picture = Image14.Picture
Case "8"
End Select
End Sub

Private Sub OnlineStatistic_Click()
If CheckRegister = "ExitSub" Then Exit Sub
Me.Hide
Form35.Show
Unload Me
End Sub
