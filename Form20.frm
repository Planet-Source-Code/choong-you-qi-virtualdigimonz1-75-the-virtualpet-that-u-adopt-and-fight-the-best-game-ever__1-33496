VERSION 5.00
Begin VB.Form Form20 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Error Fixed / Updated."
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form20"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   5280
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Error Fixed / Updated. Version 1.75"
      ForeColor       =   &H00FFFFFF&
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New Function: Minimize."
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   7680
         Width           =   1725
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Added Item Drag mouse Icon changed to the Item."
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   7440
         Width           =   3585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Added Heal Animation."
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   7200
         Width           =   1620
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Some changed in about form."
         ForeColor       =   &H0080C0FF&
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   6960
         Width           =   2085
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Added New: Check Online Tournament IP."
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   6720
         Width           =   3030
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form20.frx":0000
         ForeColor       =   &H00FFFF00&
         Height          =   555
         Left            =   240
         TabIndex        =   19
         Top             =   6120
         Width           =   4440
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Added Energy Recharge Animation."
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   5880
         Width           =   2535
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NEW: Energy Recharger. (5 Diskett Heal Item Deleted.)"
         ForeColor       =   &H000080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   5640
         Width           =   3945
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "All icon changed to 256:32x32 style. Found any smaller than that, contact the author."
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   240
         TabIndex        =   16
         Top             =   5160
         Width           =   4320
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "I think the small icon(bug) in the casino/open mystery fixed. Tester of this game please check with it."
         ForeColor       =   &H00FFFF80&
         Height          =   435
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   4170
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form20.frx":009B
         ForeColor       =   &H0000FFFF&
         Height          =   1155
         Left            =   240
         TabIndex        =   14
         Top             =   3480
         Width           =   4365
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Online Battle Now Will decrease your health. (Bug Fixed.)"
         ForeColor       =   &H00FFFF80&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   3240
         Width           =   4035
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Training -- ""Gaining"" Level Reduce. Min now is 140."
         ForeColor       =   &H00FF80FF&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   3690
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Casino Double! bug.(Money playing will +Bank money.)"
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   4290
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Connection Page Added IP History."
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   2505
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank cheque added tax% function. (50% CHARGE)"
         ForeColor       =   &H008080FF&
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   3645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPENDS ON YOUR OPPONENT."
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   2565
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BATTLE ARENA WINNING MONEY NOW WILL MULTIPLE "
         ForeColor       =   &H00FFFF00&
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   1800
         Width           =   4395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MIDI AVAILABLE."
         ForeColor       =   &H00FFC0C0&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lottery Result Now will only annouce if you have buy Lot.Tickt."
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "discover before in all visual basic website.) using in our game."
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "The BIGGEST DISCOVER -- INTERNAL SOUND (Never been"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   4455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Lottery Bug."
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fixed Slot Machine Bug."
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1725
      End
   End
   Begin VB.Menu close 
      Caption         =   "&Close"
      Begin VB.Menu nothing 
         Caption         =   "nothing"
      End
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

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
End Sub
