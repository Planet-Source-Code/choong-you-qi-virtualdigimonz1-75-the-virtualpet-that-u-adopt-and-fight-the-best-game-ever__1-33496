VERSION 5.00
Begin VB.Form Form32 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Store"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form32"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   3990
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
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
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Upgrade"
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
      Height          =   375
      Left            =   1320
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Future"
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Counter"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3495
      Begin VirtualDigimonz.FlatButton FlatButton2 
         Height          =   375
         Left            =   2280
         TabIndex        =   15
         Top             =   720
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Buy"
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
      Begin VirtualDigimonz.FlatButton FlatButton3 
         Height          =   375
         Left            =   2280
         TabIndex        =   16
         Top             =   1200
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         ForeColor       =   12632256
         BackColor       =   0
         Caption         =   "&Sell"
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
      Begin VB.Timer CheckDetail 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   2280
         Top             =   120
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sell Price:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   840
         TabIndex        =   11
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Buy Price:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Own:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   840
         TabIndex        =   9
         Top             =   1320
         Width           =   375
      End
      Begin VB.Image Image6 
         Height          =   495
         Left            =   120
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item/Space: "
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Selling"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$1000"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2760
         TabIndex        =   5
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$500"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2040
         TabIndex        =   4
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$250"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1440
         TabIndex        =   3
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$100"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "$50"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   2040
         ToolTipText     =   "Increase Defence Value by 10%"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1440
         ToolTipText     =   "Heal"
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2760
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   240
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   840
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckDetail_Timer()
Label1.Caption = "Money: $" & GetSetting("Digimon", "Profile", "Money")
Label2.Caption = "Item/Space: " & GetSetting("Digimon", "Item", "allsub") & "/" & GetSetting("Digimon", "Item", "space")
FlatButton2.Enabled = True
Select Case CheckDetail.Tag
Case 1
Label10.Caption = "Sell Price: $25"
Label10.Tag = "25"
Label8.Tag = GetSetting("Digimon", "Item", "1")
Label8.Caption = "Own: " & Label8.Tag
Case 2
Label10.Caption = "Sell Price: $50"
Label10.Tag = "50"
Label8.Tag = GetSetting("Digimon", "Item", "2")
Label8.Caption = "Own: " & Label8.Tag
Case 3
Label10.Caption = "Sell Price: $125"
Label10.Tag = "125"
Label8.Tag = GetSetting("Digimon", "Item", "3")
Label8.Caption = "Own: " & Label8.Tag
Case 4
Label10.Caption = "Sell Price: $250"
Label10.Tag = "250"
Label8.Tag = GetSetting("Digimon", "Item", "4")
Label8.Caption = "Own: " & Label8.Tag
Case 5
Label10.Caption = "Sell Price: $500"
Label10.Tag = "500"
Label8.Tag = GetSetting("Digimon", "Item", "5")
Label8.Caption = "Own: " & Label8.Tag
End Select
If Label8.Tag = "0" Then
FlatButton3.Enabled = False
Else
FlatButton3.Enabled = True
End If
CheckDetail.Enabled = False
End Sub

Private Sub FlatButton1_Click()
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Flatbutton2_Click()
ConfirmSell = MsgBox("Are You Sure You Want To Buy This Item?", vbYesNo)
If ConfirmSell = vbYes Then

spacevalue222 = GetSetting("Digimon", "Item", "space")
ItemSubvalue222 = GetSetting("Digimon", "Item", "allsub")
ItemSubvalue222 = Val(ItemSubvalue222) + Val(1)
If Val(ItemSubvalue222) > Val(spacevalue222) Then
MsgBox "You Don't have enough space to store your Item." & vbCrLf & "Please upgrade or sell something 1st before you try to buy again."
Exit Sub
End If

Select Case CheckDetail.Tag
Case 1
If WillMoney("MINUS", Label9.Tag) = "ExitSub" Then Exit Sub
Label8.Tag = Val(Label8.Tag) + Val(1)
SaveSetting "Digimon", "Item", "1", Label8.Tag
Case 2
If WillMoney("MINUS", Label9.Tag) = "ExitSub" Then Exit Sub
Label8.Tag = Val(Label8.Tag) + Val(1)
SaveSetting "Digimon", "Item", "2", Label8.Tag
Case 3
If WillMoney("MINUS", Label9.Tag) = "ExitSub" Then Exit Sub
Label8.Tag = Val(Label8.Tag) + Val(1)
SaveSetting "Digimon", "Item", "3", Label8.Tag
Case 4
If WillMoney("MINUS", Label9.Tag) = "ExitSub" Then Exit Sub
Label8.Tag = Val(Label8.Tag) + Val(1)
SaveSetting "Digimon", "Item", "4", Label8.Tag
Case 5
If WillMoney("MINUS", Label9.Tag) = "ExitSub" Then Exit Sub
Label8.Tag = Val(Label8.Tag) + Val(1)
SaveSetting "Digimon", "Item", "5", Label8.Tag
End Select
ItemSub
MsgBox "Thank you."
End If
CheckDetail.Enabled = True
End Sub

Private Sub FlatButton3_Click()
ConfirmSell = MsgBox("Are You Sure You Want To Sell This Item?", vbYesNo)
If ConfirmSell = vbYes Then
Select Case CheckDetail.Tag
Case 1
WillMoney "PLUS", Label10.Tag
Label8.Tag = Val(Label8.Tag) - Val(1)
SaveSetting "Digimon", "Item", "1", Label8.Tag
Case 2
WillMoney "PLUS", Label10.Tag
Label8.Tag = Val(Label8.Tag) - Val(1)
SaveSetting "Digimon", "Item", "2", Label8.Tag
Case 3
WillMoney "PLUS", Label10.Tag
Label8.Tag = Val(Label8.Tag) - Val(1)
SaveSetting "Digimon", "Item", "3", Label8.Tag
Case 4
WillMoney "PLUS", Label10.Tag
Label8.Tag = Val(Label8.Tag) - Val(1)
SaveSetting "Digimon", "Item", "4", Label8.Tag
Case 5
WillMoney "PLUS", Label10.Tag
Label8.Tag = Val(Label8.Tag) - Val(1)
SaveSetting "Digimon", "Item", "5", Label8.Tag
End Select
ItemSub
MsgBox "Item Sold."
End If
CheckDetail.Enabled = True
End Sub

Private Sub Flatbutton4_Click()
askingupgrade = MsgBox("This will upgrade 2 spaces for your Item storage." & vbCrLf & "Cost: $500. Are you sure?", vbYesNo)
If askingupgrade = vbYes Then
If WillMoney("MINUS", 500) = "ExitSub" Then Exit Sub
spacevalue = GetSetting("Digimon", "Item", "space")
spacevalue = Val(spacevalue) + Val(2)
SaveSetting "Digimon", "Item", "space", spacevalue
MsgBox "Item Storage Space Successfull Upgraded."
CheckDetail.Enabled = True
End If
End Sub

Private Sub Flatbutton5_Click()
MsgBox "For Future Use"
End Sub

Private Sub Form_Load()
AutoDetectTop Me
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN
Image2.Picture = Form7.Image7.Picture
Image3.Picture = Form7.Image8.Picture
Image5.Picture = Form7.Image5.Picture
Image4.Picture = Form7.Image4.Picture
Image1.Picture = Form7.Image6.Picture

Label1.Caption = "Money: $" & GetSetting("Digimon", "Profile", "Money")
Label2.Caption = "Item/Space: " & GetSetting("Digimon", "Item", "allsub") & "/" & GetSetting("Digimon", "Item", "space")
End Sub

Private Sub Image1_Click()
Image6.Picture = Image1.Picture
Label9.Caption = "Buy Price: " & Label7.Caption
Label9.Tag = "1000"
CheckDetail.Enabled = True
CheckDetail.Tag = "5"
End Sub

Private Sub Image2_Click()
Image6.Picture = Image2.Picture
Label9.Caption = "Buy Price: " & Label3.Caption
Label9.Tag = "50"
CheckDetail.Tag = "1"
CheckDetail.Enabled = True
End Sub

Private Sub Image3_Click()
Image6.Picture = Image3.Picture
Label9.Caption = "Buy Price: " & Label4.Caption
Label9.Tag = "100"
CheckDetail.Tag = "2"
CheckDetail.Enabled = True
End Sub

Private Sub Image4_Click()
Image6.Picture = Image4.Picture
Label9.Caption = "Buy Price: " & Label6.Caption
Label9.Tag = "500"
CheckDetail.Tag = "4"
CheckDetail.Enabled = True
End Sub

Private Sub Image5_Click()
Image6.Picture = Image5.Picture
Label9.Caption = "Buy Price: " & Label5.Caption
Label9.Tag = "250"
CheckDetail.Enabled = True
CheckDetail.Tag = "3"
End Sub
