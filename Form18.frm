VERSION 5.00
Begin VB.Form Form18 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Evolve"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form18"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   2040
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   2760
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1085
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "Close"
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "25000"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Image Image35 
      Height          =   480
      Left            =   2160
      Tag             =   "Angelwoman"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image33 
      Height          =   480
      Left            =   1680
      Tag             =   "Tyrano"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image32 
      Height          =   480
      Left            =   1200
      Tag             =   "Tuskmon"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image31 
      Height          =   480
      Left            =   720
      Tag             =   "Tsunomon"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image30 
      Height          =   480
      Left            =   240
      Tag             =   "Tokomon"
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image Image29 
      Height          =   480
      Left            =   3120
      Tag             =   "Tanemon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image28 
      Height          =   480
      Left            =   2640
      Tag             =   "Punimon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image27 
      Height          =   480
      Left            =   2160
      Tag             =   "Poyomon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image26 
      Height          =   480
      Left            =   1680
      Tag             =   "Piyomon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image25 
      Height          =   480
      Left            =   1200
      Tag             =   "Patamon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image24 
      Height          =   480
      Left            =   720
      Tag             =   "Orgemon"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image23 
      Height          =   480
      Left            =   240
      Tag             =   "Megadra"
      Top             =   1560
      Width           =   480
   End
   Begin VB.Image Image22 
      Height          =   480
      Left            =   3120
      Tag             =   "M_tyrano"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image21 
      Height          =   480
      Left            =   2640
      Tag             =   "Kuwagamo"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image20 
      Height          =   480
      Left            =   2160
      Tag             =   "Kunemon"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image19 
      Height          =   480
      Left            =   1680
      Tag             =   "Greymon"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image18 
      Height          =   480
      Left            =   1200
      Tag             =   "Garurumo"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image17 
      Height          =   480
      Left            =   720
      Tag             =   "Etemon"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image16 
      Height          =   480
      Left            =   240
      Tag             =   "Elecmon"
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Image15 
      Height          =   480
      Left            =   3120
      Tag             =   "Dijitama"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image14 
      Height          =   480
      Left            =   2640
      Tag             =   "Deltamon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image13 
      Height          =   480
      Left            =   2160
      Tag             =   "Coelamon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image12 
      Height          =   480
      Left            =   1680
      Tag             =   "Bukamon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image11 
      Height          =   480
      Left            =   1200
      Tag             =   "Botamon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image10 
      Height          =   480
      Left            =   720
      Tag             =   "Bakemon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Image Image9 
      Height          =   480
      Left            =   240
      Tag             =   "Angemon"
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Have: $88888888"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2205
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2_Update
Form2.Show
Unload Me
End Sub

Private Sub Form_Load()
AutoDetectTop Me
DigimonType4
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Me.Tag = GetSetting("Digimon", "Profile", "Money")
Label1.Caption = "You Have: $" & Me.Tag
End Sub

Private Sub Image10_Click()
Timer1.Tag = Image10.Tag
Timer1.Enabled = True
End Sub

Private Sub Image11_Click()
Timer1.Tag = Image11.Tag
Timer1.Enabled = True
End Sub

Private Sub Image12_Click()
Timer1.Tag = Image12.Tag
Timer1.Enabled = True
End Sub

Private Sub Image13_Click()
Timer1.Tag = Image13.Tag
Timer1.Enabled = True
End Sub

Private Sub Image14_Click()
Timer1.Tag = Image14.Tag
Timer1.Enabled = True
End Sub

Private Sub Image15_Click()
Timer1.Tag = Image15.Tag
Timer1.Enabled = True
End Sub

Private Sub Image16_Click()
Timer1.Tag = Image16.Tag
Timer1.Enabled = True
End Sub

Private Sub Image17_Click()
Timer1.Tag = Image17.Tag
Timer1.Enabled = True
End Sub

Private Sub Image18_Click()
Timer1.Tag = Image18.Tag
Timer1.Enabled = True
End Sub

Private Sub Image19_Click()
Timer1.Tag = Image19.Tag
Timer1.Enabled = True
End Sub

Private Sub Image20_Click()
Timer1.Tag = Image20.Tag
Timer1.Enabled = True
End Sub

Private Sub Image21_Click()
Timer1.Tag = Image21.Tag
Timer1.Enabled = True
End Sub

Private Sub Image22_Click()
Timer1.Tag = Image22.Tag
Timer1.Enabled = True
End Sub

Private Sub Image23_Click()
Timer1.Tag = Image23.Tag
Timer1.Enabled = True
End Sub

Private Sub Image24_Click()
Timer1.Tag = Image24.Tag
Timer1.Enabled = True
End Sub

Private Sub Image25_Click()
Timer1.Tag = Image25.Tag
Timer1.Enabled = True
End Sub

Private Sub Image26_Click()
Timer1.Tag = Image26.Tag
Timer1.Enabled = True
End Sub

Private Sub Image27_Click()
Timer1.Tag = Image27.Tag
Timer1.Enabled = True
End Sub

Private Sub Image28_Click()
Timer1.Tag = Image28.Tag
Timer1.Enabled = True
End Sub

Private Sub Image29_Click()
Timer1.Tag = Image29.Tag
Timer1.Enabled = True
End Sub

Private Sub Image30_Click()
Timer1.Tag = Image30.Tag
Timer1.Enabled = True
End Sub

Private Sub Image31_Click()
Timer1.Tag = Image31.Tag
Timer1.Enabled = True
End Sub

Private Sub Image32_Click()
Timer1.Tag = Image32.Tag
Timer1.Enabled = True
End Sub

Private Sub Image33_Click()
Timer1.Tag = Image33.Tag
Timer1.Enabled = True
End Sub

Private Sub Image35_Click()
Timer1.Tag = Image35.Tag
Timer1.Enabled = True
End Sub

Private Sub Image9_Click()
Timer1.Tag = Image9.Tag
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
If WillMoney("MINUS", Text1.Text) = "ExitSub" Then
Timer1.Enabled = False
Exit Sub
End If
SaveSetting "Digimon", "Digimon", "Type", Timer1.Tag
Form2_Update
Timer1.Enabled = False
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub
