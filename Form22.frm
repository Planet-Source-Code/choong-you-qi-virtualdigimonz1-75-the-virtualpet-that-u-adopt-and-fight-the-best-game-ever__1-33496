VERSION 5.00
Begin VB.Form Form22 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance the Point!"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2505
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form22"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   2505
   StartUpPosition =   1  'CenterOwner
   Tag             =   "0"
   Begin VirtualDigimonz.FlatButton FlatButton3 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   1440
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "+"
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
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "+"
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
      Left            =   1800
      TabIndex        =   8
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "+"
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
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Health"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defence"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Another X point left"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "Form22"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
Me.Tag = Me.Tag - 1
Text1.Text = Text1.Text + 1
SaveSetting "Digimon", "Digimon", "Power", Text1.Text
Timer1.Enabled = True
End Sub

Private Sub Flatbutton2_Click()
Me.Tag = Me.Tag - 1
Text2.Text = Text2.Text + 1
SaveSetting "Digimon", "Digimon", "Defence", Text1.Text
Timer1.Enabled = True
End Sub

Private Sub FlatButton3_Click()
Me.Tag = Me.Tag - 1
Text3.Text = Text3.Text + 1
SaveSetting "Digimon", "Digimon", "Health", Text1.Text
Timer1.Enabled = True
End Sub

Private Sub Flatbutton4_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Form2_Update
Unload Me
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Text1.Text = GetSetting("Digimon", "Digimon", "Power")
Text2.Text = GetSetting("Digimon", "Digimon", "Defence")
Text3.Text = GetSetting("Digimon", "Digimon", "Health")
End Sub

Private Sub Timer1_Timer()
Label1.Caption = "Another " & Me.Tag & "-point left"
If Me.Tag = "0" Then
FlatButton1.Enabled = False
FlatButton2.Enabled = False
FlatButton3.Enabled = False
FlatButton4.Enabled = True
End If
End Sub
