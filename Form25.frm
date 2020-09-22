VERSION 5.00
Begin VB.Form Form25 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Help"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form25"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   6855
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   495
      Left            =   4440
      TabIndex        =   18
      Top             =   3240
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&OK"
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
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Events"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   2160
      TabIndex        =   12
      Top             =   1440
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form25.frx":0000
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Battle"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   2160
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form25.frx":0093
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "City"
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   2160
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form25.frx":01C9
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Option"
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nothing special here. Just viewing your profile, your pet status, online ranking. and the Virtual Digimonz Setting."
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Reset"
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   2160
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form25.frx":02E8
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
   End
   Begin VirtualDigimonz.FlatButton FlatButton2 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&Getting Started"
      HasFocusRect    =   0   'False
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&Reset"
      HasFocusRect    =   0   'False
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
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&Option"
      HasFocusRect    =   0   'False
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
   Begin VirtualDigimonz.FlatButton FlatButton5 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&City"
      HasFocusRect    =   0   'False
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
   Begin VirtualDigimonz.FlatButton FlatButton6 
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&Battle"
      HasFocusRect    =   0   'False
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
   Begin VirtualDigimonz.FlatButton FlatButton7 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      ForeColor       =   12632256
      HasBorder       =   -1  'True
      BackColor       =   0
      Caption         =   "&Events"
      HasFocusRect    =   0   'False
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "&Getting Started"
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   4095
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Very easy to get start, Press the left hand side menu for help."
         ForeColor       =   &H00FFFFFF&
         Height          =   675
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   2235
      End
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub FlatButton1_Click()
Unload Me
Form2.Show
End Sub

Private Sub Flatbutton2_Click()
Frame1.Visible = True
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub FlatButton3_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub Flatbutton4_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = True
End Sub

Private Sub Flatbutton5_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = False
Frame5.Visible = True
Frame6.Visible = False
End Sub

Private Sub FlatButton6_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = True
Frame4.Visible = False
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub FlatButton7_Click()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
Frame4.Visible = True
Frame5.Visible = False
Frame6.Visible = False
End Sub

Private Sub Form_Load()
AutoDetectTop Me
Set m_cN2 = New cNeoCaption
Skin2 Me, m_cN2
End Sub
