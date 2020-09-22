VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form24 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setting"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form24"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4815
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3840
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VirtualDigimonz.FlatButton FlatButton3 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
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
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Cancel"
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
      Caption         =   "Setting"
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CheckBox Check4 
         BackColor       =   &H00000000&
         Caption         =   "Make Virtual Digimonz Program Always On Top"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Value           =   1  'Checked
         Width           =   3855
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Record Battle TCP/IP LOG:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   3495
      End
      Begin VirtualDigimonz.FlatButton FlatButton6 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Clear"
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
      Begin VirtualDigimonz.FlatButton FlatButton5 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Clear"
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
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Browse"
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
      Begin VirtualDigimonz.FlatButton FlatButton2 
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Browse"
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Sound (Midi) :"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "C:\"
         Top             =   1320
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00000000&
         Caption         =   "Battle Arena Background:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Text            =   "C:\"
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version Number:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1170
   End
End
Attribute VB_Name = "Form24"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
Select Case Check1.Value
Case 1
Text1.Enabled = True
FlatButton1.Enabled = True
FlatButton5.Enabled = True
Case 0
Text1.Enabled = False
FlatButton1.Enabled = False
FlatButton5.Enabled = False
End Select
End Sub

Private Sub Check2_Click()
'Select Case Check2.Value
'Case 1
'Text2.Enabled = True
'FlatButton2.Enabled = True
'FlatButton6.Enabled = True
'Case 0
'Text2.Enabled = False
'FlatButton2.Enabled = False
'FlatButton6.Enabled = False
'End Select
End Sub

Private Sub FlatButton1_Click()
CommonDialog1.DialogTitle = "Open Battle Arena Background Picture File"
CommonDialog1.Filter = _
"JPEG/BMP/GIF (*.bmp, *.jpg, *.gif)|*.bmp;*.jpg;*.gif;" _
& "|Bitmap (*.bmp)|*.bmp" _
& "|JPEG (*.jpg)|*.jpg;*.jpeg;" _
& "|GIF (*.gif)|*.gif"

CommonDialog1.ShowOpen
If CommonDialog1.Filename = "" Then Exit Sub
Select Case right(CommonDialog1.Filename, 4)
Case ".gif"
Case ".jpg"
Case "jpeg"
Case ".bmp"
Case Else
MsgBox "Please specify a real filename"
Exit Sub
End Select
Text1.Text = CommonDialog1.Filename
End Sub

Private Sub Flatbutton2_Click()
CommonDialog1.DialogTitle = "Open Virtual Digimonz Sound File"
CommonDialog1.Filter = "WAV (*.wav)|*.wav;"
CommonDialog1.ShowOpen
If Not right(CommonDialog1.Filename, 4) = ".wav" Then
MsgBox "Please specify a real filename"
Exit Sub
End If
Text2.Text = CommonDialog1.Filename
End Sub

Private Sub FlatButton3_Click()
If Check1.Value = "1" Then
If Text1.Text = "" Then
MsgBox "Please choose a file 1st."
Exit Sub
End If
End If
'If Check2.Value = "1" Then
'If Text2.Text = "" Then
'MsgBox "Please choose a file 1st."
'Exit Sub
'End If
'End If
If Check1.Value = "0" Then Text1.Text = ""
If Check2.Value = "0" Then Text2.Text = ""
SaveSetting "Digimon", "Digimon", "Setting1", Check1.Value
SaveSetting "Digimon", "Digimon", "Setting1-1", Text1.Text
SaveSetting "Digimon", "Digimon", "Setting2", Check2.Value
SaveSetting "Digimon", "Digimon", "Setting2-1", Text2.Text
SaveSetting "Digimon", "Digimon", "Setting3", Check3.Value
SaveSetting "Digimon", "Digimon", "Setting4", Check4.Value

AutoDetectTop Form2
CheckMidiPlay

Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Flatbutton4_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Flatbutton5_Click()
Text1.Text = ""
Check1.Value = 0
Text1.Enabled = False
FlatButton1.Enabled = False
FlatButton5.Enabled = False
End Sub

Private Sub FlatButton6_Click()
Text2.Text = ""
Check2.Value = 0
Text2.Enabled = False
FlatButton2.Enabled = False
FlatButton6.Enabled = False
End Sub

Private Sub Form_Load()
AutoDetectTop Me
Check1.Value = GetSetting("Digimon", "Digimon", "Setting1")
Text1.Text = GetSetting("Digimon", "Digimon", "Setting1-1")
Check2.Value = GetSetting("Digimon", "Digimon", "Setting2")
Text2.Text = GetSetting("Digimon", "Digimon", "Setting2-1")
Check3.Value = GetSetting("Digimon", "Digimon", "Setting3")
Check4.Value = GetSetting("Digimon", "Digimon", "Setting4")

If Check1.Value = "1" Then Text1.Enabled = True
If Check2.Value = "1" Then Text2.Enabled = True

m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Label1.Caption = Label1.Caption & App.Major & "." & App.Minor & App.Revision
End Sub
