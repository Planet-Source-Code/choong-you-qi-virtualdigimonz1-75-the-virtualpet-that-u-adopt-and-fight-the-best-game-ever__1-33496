VERSION 5.00
Begin VB.Form Form28 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   1680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form28"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   1680
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   480
      Top             =   2400
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Driver"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   660
      TabIndex        =   5
      Top             =   1560
      Width           =   735
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   495
         Begin VB.Image Image6 
            Height          =   495
            Left            =   0
            Top             =   0
            Width           =   495
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   2400
   End
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   480
      TabIndex        =   4
      Top             =   360
      Width           =   120
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Width           =   120
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1200
      Width           =   120
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   120
   End
   Begin VB.Image Image1 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   120
      Tag             =   "Heal20"
      ToolTipText     =   "This disket will heal 20 point."
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   840
      Tag             =   "Heal50"
      ToolTipText     =   "This disket will heal 50 point."
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   120
      Tag             =   "Heal100"
      ToolTipText     =   "This disket will heal 100 point."
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image4 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   840
      Tag             =   "Heal200"
      ToolTipText     =   "This disket will heal 200 point."
      Top             =   960
      Width           =   480
   End
   Begin VB.Image Image5 
      DragMode        =   1  'Automatic
      Height          =   480
      Left            =   120
      Tag             =   "Heal500"
      ToolTipText     =   "This Batery will full up your energy."
      Top             =   1800
      Width           =   480
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub close_Click()
Unload Me
End Sub

Private Sub FlatButton1_Click()
Unload Me
End Sub

Private Sub Form_Load()
AutoDetectTop Me
Set m_cN2 = New cNeoCaption
Skin2 Me, m_cN2

Me.left = Form2.left - Me.Width
Me.top = Form2.top
Image1.Picture = Form7.Image7.Picture
Image2.Picture = Form7.Image8.Picture
Image3.Picture = Form7.Image5.Picture
Image4.Picture = Form7.Image4.Picture
Image5.Picture = Form7.Image6.Picture
Image6.Picture = Form7.Image3.Picture
Label1.Caption = GetSetting("Digimon", "Item", "1")
Label2.Caption = GetSetting("Digimon", "Item", "2")
Label3.Caption = GetSetting("Digimon", "Item", "3")
Label4.Caption = GetSetting("Digimon", "Item", "4")
Label5.Caption = GetSetting("Digimon", "Item", "5")
Image1.DragIcon = Image1.Picture
Image2.DragIcon = Image2.Picture
Image3.DragIcon = Image3.Picture
Image4.DragIcon = Image4.Picture
Image5.DragIcon = Image5.Picture
End Sub

Private Sub Image6_DragDrop(Source As Control, X As Single, Y As Single)
maxhealthvalue = GetSetting("Digimon", "Digimon", "Health")
currenthealthvalue = GetSetting("Digimon", "Digimon", "CurrentHealth")

If Not Source.Tag = "Heal500" Then

If maxhealthvalue = currenthealthvalue Then
MsgBox "Your pet is fully heal."
Exit Sub
End If

Else

If Form2.ProgressBar1.Value = "100" Then
MsgBox "Your pet is fully energy."
Exit Sub
End If

End If

If Not left(Source.Tag, 4) = "Heal" Then Exit Sub
Select Case Mid(Source.Tag, 5)
Case "20"
Label1.Caption = Val(Label1.Caption) - Val(1)
If left(Label1.Caption, 1) = "-" Then
Label1.Caption = "0"
MsgBox "You don't have this Item."
Exit Sub
End If
SaveSetting "Digimon", "Item", "1", Label1.Caption
Case "50"
Label2.Caption = Val(Label2.Caption) - Val(1)
If left(Label2.Caption, 1) = "-" Then
Label2.Caption = "0"
MsgBox "You don't have this Item."
Exit Sub
End If
SaveSetting "Digimon", "Item", "2", Label2.Caption
Case "100"
Label3.Caption = Val(Label3.Caption) - Val(1)
If left(Label3.Caption, 1) = "-" Then
Label3.Caption = "0"
MsgBox "You don't have this Item."
Exit Sub
End If
SaveSetting "Digimon", "Item", "3", Label3.Caption
Case "200"
Label4.Caption = Val(Label4.Caption) - Val(1)
If left(Label4.Caption, 1) = "-" Then
Label4.Caption = "0"
MsgBox "You don't have this Item."
Exit Sub
End If
SaveSetting "Digimon", "Item", "4", Label4.Caption
Case "500"
Label5.Caption = Val(Label5.Caption) - Val(1)
If left(Label5.Caption, 1) = "-" Then
Label5.Caption = "0"
MsgBox "You don't have this Item."
Exit Sub
End If
SaveSetting "Digimon", "Item", "5", Label5.Caption
End Select
Image6.Picture = Source.Picture
Timer2.Tag = Mid(Source.Tag, 5)
Timer2.Enabled = True
End Sub

Private Sub Timer1_Timer()
Form28.left = Form2.left - Form28.Width
Form28.top = Form2.top
End Sub

Private Sub Timer2_Timer()
ItemSub
If Image6.top = 600 Then
Image6.top = 0


''''''''ENERGY FULL CODE''''''''
If Timer2.Tag = "500" Then
SaveSetting "Digimon", "Digimon", "Energy", "100"
Timer2.Enabled = False
Image6.Picture = Form7.Image3.Picture
Form2.Anim_EnergyFull.Enabled = True
Exit Sub
End If
''''''''ENERGY FULL CODE''''''''


maxhealthvalue = GetSetting("Digimon", "Digimon", "Health")
currenthealthvalue = GetSetting("Digimon", "Digimon", "CurrentHealth")
Animation_Heal (currenthealthvalue)
currenthealthvalue = Val(currenthealthvalue) + Val(Timer2.Tag)
If Val(currenthealthvalue) > Val(maxhealthvalue) Then currenthealthvalue = maxhealthvalue
SaveSetting "Digimon", "Digimon", "CurrentHealth", currenthealthvalue
Form2.Anim_Heal.Enabled = True
Timer2.Enabled = False
Image6.Picture = Form7.Image3.Picture
Exit Sub
End If
Image6.top = Image6.top + 120
End Sub
