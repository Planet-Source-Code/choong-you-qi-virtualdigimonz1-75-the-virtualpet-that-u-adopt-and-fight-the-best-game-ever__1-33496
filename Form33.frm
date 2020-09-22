VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form33 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bar of Excitement"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form33"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4260
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Status: "
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   3615
      Begin VirtualDigimonz.FlatButton FlatButton2 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Play"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
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
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   240
         Top             =   1440
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Left            =   720
         Top             =   1440
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Number:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   960
         TabIndex        =   5
         Top             =   840
         Width           =   165
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money: $ "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$50 per play"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1050
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      MousePointer    =   1
      Enabled         =   0   'False
      LargeChange     =   0
      Max             =   12
      TickStyle       =   3
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "88"
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
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   225
   End
End
Attribute VB_Name = "Form33"
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

Private Sub Flatbutton2_Click()
Confirmplay = MsgBox("$50 per play. Are you sure?", vbYesNo)
If Confirmplay = vbYes Then
If WillMoney("MINUS", 50) = "ExitSub" Then Exit Sub
Label4.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
Timer2.Interval = GenerateRandom(1, 10000)
Timer2.Enabled = True
FlatButton1.Enabled = False
FlatButton2.Enabled = False
Label2.Caption = "Playing..."
End If
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

Label4.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
End Sub


Private Sub Timer1_Timer()
If Slider1.Value = "12" Then Slider1.Tag = "1"
If Slider1.Value = "0" Then Slider1.Tag = "0"
Select Case Slider1.Tag
Case "0"
Slider1.Value = Slider1.Value + Val(1)
Case "1"
Slider1.Value = Slider1.Value - Val(1)
End Select
Label1.Caption = Slider1.Value
End Sub

Private Sub Timer2_Timer()
Dim PrizeGot
Timer1.Enabled = False
Label3.Caption = Label1.Caption
Timer1.Enabled = True
FlatButton1.Enabled = True
FlatButton2.Enabled = True
Timer2.Enabled = False
Label2.Caption = "$50 per play"
PrizeGot = "0"

Select Case Label3.Caption
Case "0"
WillEnergy "PLUS", Val(Form2.ProgressBar1.Max) - Val(Form2.ProgressBar1.Value)
MsgBox "Energy Full"
PrizeGot = "1"
Case "1"
WillMoney "PLUS", GenerateRandom(50, 300)
Label4.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Win Some Money..."
PrizeGot = "1"
Case "3"
Select Case Val(GenerateRandom(1, 41))
Case 1 To 10
howmany = GetSetting("Digimon", "Item", "1")
howmany = Val(howmany) + Val(1)
SaveSetting "Digimon", "Item", "1", howmany
Case 11 To 20
howmany = GetSetting("Digimon", "Item", "2")
howmany = Val(howmany) + Val(1)
SaveSetting "Digimon", "Item", "2", howmany
Case 21 To 30
howmany = GetSetting("Digimon", "Item", "3")
howmany = Val(howmany) + Val(1)
SaveSetting "Digimon", "Item", "3", howmany
Case 31 To 40
howmany = GetSetting("Digimon", "Item", "4")
howmany = Val(howmany) + Val(1)
SaveSetting "Digimon", "Item", "4", howmany
Case 41
howmany = GetSetting("Digimon", "Item", "5")
howmany = Val(howmany) + Val(1)
SaveSetting "Digimon", "Item", "5", howmany
End Select
MsgBox "Win Random Diskett..."
PrizeGot = "1"
Case "7"
Select Case GenerateRandom(1, 3)
Case 1
WillHealth "PLUS", 2
Case 2
WillAttack "PLUS", 2
Case 3
WillDefence "PLUS", 2
End Select
Form2_Update
MsgBox "Random gain 2 attack or defence or health point..."
PrizeGot = "1"
Case "12"
SaveSetting "Digimon", "Digimon", "Fooded", GetSetting("Digimon", "Profile", "MaxFood")
Form2.Label5.Tag = GetSetting("Digimon", "Digimon", "Fooded")
Food_Event_Caption
MsgBox "Full(Not Hungry)"
PrizeGot = "1"
End Select
If Not PrizeGot = "1" Then MsgBox Label3.Caption & ", Bad Luck!"
End Sub
