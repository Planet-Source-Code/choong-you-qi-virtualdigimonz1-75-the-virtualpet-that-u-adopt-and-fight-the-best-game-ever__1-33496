VERSION 5.00
Begin VB.Form Form30 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "News Center"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form30"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4035
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Download"
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
   Begin VirtualDigimonz.FlatButton FlatButton3 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Bonus"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Leave"
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
      Left            =   360
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Get Info"
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
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2325
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   3375
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00004080&
      BorderWidth     =   8
      Height          =   255
      Index           =   0
      Left            =   3480
      Top             =   2900
      Width           =   255
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
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3120
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00000000&
      Height          =   1890
      Left            =   3240
      TabIndex        =   7
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "no date available."
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
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Document Done."
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   1440
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form30"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub FlatButton1_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Flatbutton2_Click()
Label1.Caption = "Getting Information..."
MousePointer = vbHourglass
Label2.Caption = GetUrlSource("http://www.angelfire.com/co/choongyouqi/VirtualDigimonz/Date.dat")
MousePointer = vbDefault
Label1.Caption = "Document Done."

Label1.Caption = "Getting Information..."
MousePointer = vbHourglass
Text1.Text = GetUrlSource("http://www.angelfire.com/co/choongyouqi/VirtualDigimonz/News.dat")
MousePointer = vbDefault
Label1.Caption = "Document Done."
If right(Text1.Text, 6) = "RcD4Fg" Then FlatButton3.Visible = True
End Sub

Private Sub FlatButton3_Click()
MsgBox "Comming Soon"
Me.Tag = GetSetting("Digimon", "Profile", "Money")
SaveSetting "Digimon", "Profile", "Money", Me.Tag
End Sub
Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
End Sub

