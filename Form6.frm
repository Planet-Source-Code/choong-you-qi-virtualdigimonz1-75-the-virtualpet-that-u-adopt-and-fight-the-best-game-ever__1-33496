VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Your Charactor"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image8 
      Height          =   480
      Left            =   3000
      Picture         =   "Form6.frx":0000
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   2160
      Picture         =   "Form6.frx":0CCA
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image6 
      Height          =   480
      Left            =   1320
      Picture         =   "Form6.frx":1994
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   480
      Picture         =   "Form6.frx":265E
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   3000
      Picture         =   "Form6.frx":3328
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   2160
      Picture         =   "Form6.frx":3FF2
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1320
      Picture         =   "Form6.frx":4CBC
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "Form6.frx":5986
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
AutoDetectTop Me
Set m_cN3 = New cNeoCaption
Skin2 Me, m_cN3
End Sub

Private Sub Image1_Click()
Form4.Picture1.Picture = Image1.Picture
Form4.Picture1.Tag = "1"
Unload Me
End Sub
Private Sub Image2_Click()
Form4.Picture1.Picture = Image2.Picture
Form4.Picture1.Tag = "2"
Unload Me
End Sub
Private Sub Image3_Click()
Form4.Picture1.Picture = Image3.Picture
Form4.Picture1.Tag = "3"
Unload Me
End Sub
Private Sub Image4_Click()
Form4.Picture1.Picture = Image4.Picture
Form4.Picture1.Tag = "4"
Unload Me
End Sub
Private Sub Image5_Click()
Form4.Picture1.Picture = Image5.Picture
Form4.Picture1.Tag = "5"
Unload Me
End Sub
Private Sub Image6_Click()
Form4.Picture1.Picture = Image6.Picture
Form4.Picture1.Tag = "6"
Unload Me
End Sub
Private Sub Image7_Click()
Form4.Picture1.Picture = Image7.Picture
Form4.Picture1.Tag = "7"
Unload Me
End Sub
Private Sub Image8_Click()
Form4.Picture1.Picture = Image8.Picture
Form4.Picture1.Tag = "8"
Unload Me
End Sub

