VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Owner Profile"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Badges"
      ForeColor       =   &H00FFFFFF&
      Height          =   1600
      Left            =   195
      TabIndex        =   6
      Top             =   2280
      Width           =   3495
      Begin VB.Image Image9 
         Height          =   480
         Left            =   2760
         Picture         =   "Form3.frx":0000
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   1920
         Picture         =   "Form3.frx":0CCA
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   1080
         Picture         =   "Form3.frx":1994
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   240
         Picture         =   "Form3.frx":265E
         Top             =   960
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   2760
         Picture         =   "Form3.frx":3328
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1920
         Picture         =   "Form3.frx":3FF2
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1080
         Picture         =   "Form3.frx":4CBC
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Picture         =   "Form3.frx":5986
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Profile"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   195
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Image Image13 
         Height          =   1125
         Left            =   120
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money: $"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exp.Points: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rank: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   3
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Played: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   840
         X2              =   840
         Y1              =   360
         Y2              =   2040
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   1680
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00000000&
      Height          =   1890
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   525
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".........................."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   540
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3120
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00000080&
      BorderWidth     =   8
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   480
      Width           =   495
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   1080
      Picture         =   "Form3.frx":6650
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image12 
      Height          =   1125
      Left            =   4560
      Picture         =   "Form3.frx":731A
      Top             =   1560
      Width           =   600
   End
   Begin VB.Image Image11 
      Height          =   1125
      Left            =   4560
      Picture         =   "Form3.frx":814E
      Top             =   2640
      Width           =   600
   End
   Begin VB.Image Image10 
      Height          =   1125
      Left            =   4560
      Picture         =   "Form3.frx":9094
      Top             =   480
      Width           =   600
   End
   Begin VB.Menu close 
      Caption         =   "Close"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form3"
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

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Select Case GetSetting("Digimon", "Profile", "Badge")
Case "0"
SaveSetting "Digimon", "Profile", "Ranked", "Junior"
Image2.Picture = Image14.Picture
Image3.Picture = Image14.Picture
Image4.Picture = Image14.Picture
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "1"
SaveSetting "Digimon", "Profile", "Ranked", "Junior"
Image3.Picture = Image14.Picture
Image4.Picture = Image14.Picture
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "2"
SaveSetting "Digimon", "Profile", "Ranked", "Junior"
Image4.Picture = Image14.Picture
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "3"
SaveSetting "Digimon", "Profile", "Ranked", "Senior"
Image5.Picture = Image14.Picture
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "4"
SaveSetting "Digimon", "Profile", "Ranked", "Senior"
Image6.Picture = Image14.Picture
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "5"
SaveSetting "Digimon", "Profile", "Ranked", "Senior"
Image7.Picture = Image14.Picture
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "6"
SaveSetting "Digimon", "Profile", "Ranked", "Master"
Image8.Picture = Image14.Picture
Image9.Picture = Image14.Picture
Case "7"
SaveSetting "Digimon", "Profile", "Ranked", "Master"
Image9.Picture = Image14.Picture
Case "8"
SaveSetting "Digimon", "Profile", "Ranked", "Master"
End Select

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
Label2.Caption = Label2.Caption & GetSetting("Digimon", "Profile", "Played") & "min"
Label3.Caption = Label3.Caption & GetSetting("Digimon", "Profile", "Ranked")
Label4.Caption = Label4.Caption & GetSetting("Digimon", "Profile", "Score")
Label5.Caption = Label5.Caption & GetSetting("Digimon", "Profile", "Money")
Label5.Tag = GetSetting("Digimon", "Profile", "Money")
Label4.Tag = GetSetting("Digimon", "Profile", "Score")
Select Case right(Label3.Caption, 6)
Case "Junior"
Image13.Picture = Image10.Picture
Case "Senior"
Image13.Picture = Image12.Picture
Case "Master"
Image13.Picture = Image11.Picture
End Select

End Sub
