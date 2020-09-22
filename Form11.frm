VERSION 5.00
Begin VB.Form Form11 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About -- Virtual Digimonz"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form11.frx":0000
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   3840
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "OK"
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
   Begin VB.TextBox txtSpecialCopyright 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "Form11.frx":08CA
      Top             =   3300
      Width           =   4455
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   360
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "choongyouqi@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4320
      Width           =   1920
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://vb.onweb.cx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   3000
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   4560
      Width           =   1380
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.choongyouqi.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4560
      Width           =   2130
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "32349038 -- Shivan Dragon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   720
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4800
      Width           =   1965
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.75(Beta)"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   735
      TabIndex        =   12
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   120
      X2              =   5669
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label LabelCopyRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2001-2002 Choong You Qi."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   735
      MouseIcon       =   "Form11.frx":0C00
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Tag             =   "http://vbaccelerator.com/j-index.htm?url=cright.htm"
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label lblProduct 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://VirtualDigimonz.ChoongYouQi.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   735
      MouseIcon       =   "Form11.frx":0F0A
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Tag             =   "http://vbaccelerator.com/"
      Top             =   2820
      Width           =   2940
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   ".................."
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
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1020
      Width           =   2115
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Choong You Qi"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   585
      Index           =   6
      Left            =   1440
      TabIndex        =   8
      Top             =   780
      Width           =   3525
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual Digimonz(Beta Version). Hope you all have fun with it. Any comment, email me."
      ForeColor       =   &H00E0E0E0&
      Height          =   705
      Index           =   5
      Left            =   1200
      TabIndex        =   7
      Top             =   1560
      Width           =   3750
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   420
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Virtual Digimonz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1020
      TabIndex        =   4
      Top             =   420
      Width           =   2865
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00000080&
      BorderWidth     =   16
      Height          =   855
      Index           =   1
      Left            =   240
      Top             =   300
      Width           =   855
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   2
      Left            =   4920
      Top             =   240
      Width           =   555
   End
   Begin VB.Label lblInfo 
      BackColor       =   &H000080FF&
      Height          =   225
      Index           =   4
      Left            =   1140
      TabIndex        =   6
      Top             =   900
      Width           =   3885
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H000040C0&
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1140
      TabIndex        =   5
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Shape shpRect 
      BorderColor     =   &H00004080&
      BorderWidth     =   8
      Height          =   675
      Index           =   0
      Left            =   4560
      Top             =   660
      Width           =   615
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   ".............."
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
      Height          =   495
      Index           =   2
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   2475
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "NeoCaption"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Top             =   780
      Width           =   3375
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "A serious window-style modification framework for Visual Basic programmers. "
      ForeColor       =   &H00E0E0E0&
      Height          =   705
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   1500
      Width           =   3750
   End
   Begin VB.Shape shpRect 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00004040&
      FillStyle       =   0  'Solid
      Height          =   735
      Index           =   3
      Left            =   4440
      Top             =   1680
      Width           =   795
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FlatButton1_Click()
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

Label8.MouseIcon = LabelCopyRight.MouseIcon
Label7.MouseIcon = LabelCopyRight.MouseIcon
Label5.MouseIcon = LabelCopyRight.MouseIcon
Label9.MouseIcon = LabelCopyRight.MouseIcon
lblProduct.MouseIcon = LabelCopyRight.MouseIcon

Me.Caption = Me.Caption & " Ver" & App.Major & "." & App.Minor & App.Revision
Image2.Picture = Me.Icon
Select Case GenerateRandom(1, 10)
Case "1"
Image5.Picture = Form7.Image9.Picture
Case "2"
Image5.Picture = Form7.Image13.Picture
Case "3"
Image5.Picture = Form7.Image14.Picture
Case "4"
Image5.Picture = Form7.Image18.Picture
Case "5"
Image5.Picture = Form7.Image19.Picture
Case "6"
Image5.Picture = Form7.Image21.Picture
Case "7"
Image5.Picture = Form7.Image22.Picture
Case "8"
Image5.Picture = Form7.Image23.Picture
Case "9"
Image5.Picture = Form7.Image33.Picture
Case "10"
Image5.Picture = Form7.Image35.Picture
End Select
txtSpecialCopyright.Text = left(txtSpecialCopyright, Len(txtSpecialCopyright) - Val(2))
End Sub

Private Sub Image5_Click()
Select Case GenerateRandom(1, 10)
Case "1"
Image5.Picture = Form7.Image9.Picture
Case "2"
Image5.Picture = Form7.Image13.Picture
Case "3"
Image5.Picture = Form7.Image14.Picture
Case "4"
Image5.Picture = Form7.Image18.Picture
Case "5"
Image5.Picture = Form7.Image19.Picture
Case "6"
Image5.Picture = Form7.Image21.Picture
Case "7"
Image5.Picture = Form7.Image22.Picture
Case "8"
Image5.Picture = Form7.Image23.Picture
Case "9"
Image5.Picture = Form7.Image33.Picture
Case "10"
Image5.Picture = Form7.Image35.Picture
End Select
End Sub

Private Sub Label5_Click()
email = ShellExecute(hWnd, "open", "mailto:choongyouqi@hotmail.com", vbNull, vbNull, SW_SHOWNORMAL)
End Sub

Private Sub Label7_Click()
web1 = ShellExecute(hWnd, "open", "http://vb.onweb.cx", vbNull, vbNull, SW_SHOWNORMAL)
End Sub

Private Sub Label8_Click()
web2 = ShellExecute(hWnd, "open", "http://www.choongyouqi.com", vbNull, vbNull, SW_SHOWNORMAL)
End Sub

Private Sub Label9_Click()
ICQadd = ShellExecute(hWnd, "open", "http://web.icq.com/whitepages/add_me/1,,,00.icq?uin=32349038&action=add", vbNull, vbNull, SW_SHOWNORMAL)
End Sub

Private Sub LabelCopyRight_Click()
MsgBox "Refer to CopyRight Information above."
End Sub

Private Sub lblProduct_Click()
web3 = ShellExecute(hWnd, "open", "http://VirtualDigimonz.ChoongYouQi.com", vbNull, vbNull, SW_SHOWNORMAL)
End Sub
