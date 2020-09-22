VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4665
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0FF&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3360
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3840
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   2295
      Left            =   240
      Picture         =   "Form1.frx":000C
      ScaleHeight     =   2235
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   240
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   2160
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scanning: "
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
      TabIndex        =   0
      Tag             =   "1234"
      Top             =   2640
      Width           =   930
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
i = 1
End Sub

Private Sub Timer1_Timer()
Text1.Text = GetSetting("Digimon", "FirstTime", "FirstTime")
Select Case Text1.Text
Case ""
SaveSetting "Digimon", "FirstTime", "FirstTime", "1"
Timer1.Enabled = False
Form4.Show
Unload Me
Exit Sub
Case "1"
If Not GetSetting("Digimon", "Profile", "Version") = App.Major & "." & App.Minor & App.Revision Then
Reset_Program
MsgBox "Sorry, We Detected You Are Using An Old Version Of Virtual Digimonz." & vbCrLf & "This version didn't support for older version." & vbCrLf & "Please setup your profile again, thank you."
Form4.Show
Unload Me
End If
Timer2.Enabled = True
End Select
End Sub

Private Sub Timer2_Timer()
If Label1.Caption = "Scanning: .sfx" Then
Timer2.Enabled = False
Form2.Show
Unload Me
Exit Sub
End If
Label1.Caption = left(Label1.Caption, 10) & left(Mid(Label1.Tag, i), 4) & ".sfx"
i = i + 1
End Sub
