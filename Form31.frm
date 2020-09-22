VERSION 5.00
Begin VB.Form Form31 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Market"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form31"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VirtualDigimonz.FlatButton FlatButton1 
      Height          =   375
      Left            =   2400
      TabIndex        =   19
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Buy"
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
      Left            =   2400
      TabIndex        =   20
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Sell"
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
      Left            =   3720
      TabIndex        =   21
      Top             =   2040
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
   Begin VirtualDigimonz.FlatButton FlatButton4 
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ForeColor       =   12632256
      BackColor       =   0
      Caption         =   "&Internet"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "&Your Money"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   2055
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1440
         Top             =   240
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Money: $"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Stock &Market"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Own:"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   240
         Width           =   450
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   15
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Height          =   255
         Left            =   4320
         TabIndex        =   13
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot:"
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
         Left            =   3720
         TabIndex        =   12
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value:"
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
         Left            =   2880
         TabIndex        =   11
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Company:"
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
         TabIndex        =   10
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "200"
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
         Left            =   3720
         TabIndex        =   9
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "200"
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
         Left            =   3720
         TabIndex        =   8
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "200"
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
         Left            =   3720
         TabIndex        =   7
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1320
         Width           =   330
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "3) Truso Private Hospital"
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
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "2) Monster Equitment Corp."
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "1) Digi Drink Co. Ltd."
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
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form31"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LotValue As Integer
Dim TotalMoney As Long
Private Sub FlatButton1_Click()
Dim BuyShare2 As Long
Select Case InputBox(Label1.Caption & vbCrLf & Label2.Caption & vbCrLf & Label3.Caption & vbCrLf & vbCrLf & "Please Select A Number.")
Case "1"

On Error Resume Next
BuyShare2 = InputBox(Label1.Caption & vbCrLf & vbCrLf & "Current Price: " & vbTab & "$" & Label4.Caption & vbCrLf & "Current Own: " & vbTab & Label13.Caption & vbCrLf & "Lot Left: " & vbTab & Label7.Caption & vbCrLf & vbCrLf & "How many Lot you want to buy?")
If BuyShare2 > Val(Label7.Caption) Then
MsgBox "Sorry, We don't have so many Lot for sell."
Exit Sub
End If
If BuyShare2 = "0" Then Exit Sub
TotalMoney = Val(Label4.Caption) * Val(BuyShare2)
ConfirmBuyShare = MsgBox("Buy " & BuyShare2 & " Lot on " & Mid(Label1.Caption, 3) & "?" & vbCrLf & "There will be $" & TotalMoney & ". Are you sure?", vbYesNo)
If ConfirmBuyShare = vbNo Then Exit Sub
If ConfirmBuyShare = vbYes Then
If WillMoney("MINUS", TotalMoney) = "ExitSub" Then Exit Sub
Label13.Caption = Val(Label13.Caption) + Val(BuyShare2)
Label7.Caption = Val(LotValue) - Val(Label13.Caption)
SaveSetting "Digimon", "Share", "Own1", Label13.Caption
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Purchased"
Exit Sub
End If

Case "2"

On Error Resume Next
BuyShare2 = InputBox(Label2.Caption & vbCrLf & vbCrLf & "Current Price: " & vbTab & "$" & Label5.Caption & vbCrLf & "Current Own: " & vbTab & Label14.Caption & vbCrLf & "Lot Left: " & vbTab & Label8.Caption & vbCrLf & vbCrLf & "How many Lot you want to buy?")
If BuyShare2 > Val(Label8.Caption) Then
MsgBox "Sorry, We don't have so many Lot for sell."
Exit Sub
End If
If BuyShare2 = "0" Then Exit Sub

TotalMoney = Val(Label5.Caption) * Val(BuyShare2)
ConfirmBuyShare = MsgBox("Buy " & BuyShare2 & " Lot on " & Mid(Label2.Caption, 3) & "?" & vbCrLf & "There will be $" & TotalMoney & ". Are you sure?", vbYesNo)
If ConfirmBuyShare = vbNo Then Exit Sub
If ConfirmBuyShare = vbYes Then
If WillMoney("MINUS", TotalMoney) = "ExitSub" Then Exit Sub
Label14.Caption = Val(Label14.Caption) + Val(BuyShare2)
Label8.Caption = Val(LotValue) - Val(Label14.Caption)
SaveSetting "Digimon", "Share", "Own2", Label14.Caption
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Purchased"
Exit Sub
End If

Case "3"

On Error Resume Next
BuyShare2 = InputBox(Label3.Caption & vbCrLf & vbCrLf & "Current Price: " & vbTab & "$" & Label6.Caption & vbCrLf & "Current Own: " & vbTab & Label15.Caption & vbCrLf & "Lot Left: " & vbTab & Label9.Caption & vbCrLf & vbCrLf & "How many Lot you want to buy?")
If BuyShare2 > Val(Label9.Caption) Then
MsgBox "Sorry, We don't have so many Lot for sell."
Exit Sub
End If
If BuyShare2 = "0" Then Exit Sub

TotalMoney = Val(Label6.Caption) * Val(BuyShare2)
ConfirmBuyShare = MsgBox("Buy " & BuyShare2 & " Lot on " & Mid(Label3.Caption, 3) & "?" & vbCrLf & "There will be $" & TotalMoney & ". Are you sure?", vbYesNo)
If ConfirmBuyShare = vbNo Then Exit Sub
If ConfirmBuyShare = vbYes Then
If WillMoney("MINUS", TotalMoney) = "ExitSub" Then Exit Sub
Label15.Caption = Val(Label15.Caption) + Val(BuyShare2)
Label9.Caption = Val(LotValue) - Val(Label15.Caption)
SaveSetting "Digimon", "Share", "Own3", Label15.Caption
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Purchased"
Exit Sub
End If

Case ""
Exit Sub
End Select
MsgBox "Please select a legal number."
End Sub

Private Sub Flatbutton2_Click()
Dim SellShare2 As Long
Select Case InputBox(Label1.Caption & vbCrLf & Label2.Caption & vbCrLf & Label3.Caption & vbCrLf & vbCrLf & "Please Select A Number.")
Case "1"

On Error Resume Next
SellShare2 = InputBox(Label1.Caption & vbCrLf & vbCrLf & "Current Price: " & vbTab & "$" & Label4.Caption & vbCrLf & "Lot Left: " & vbTab & Label7.Caption & vbCrLf & "Current Own: " & vbTab & Label13.Caption & vbCrLf & vbCrLf & "How many Lot you want to sell?")
If SellShare2 > Val(Label13.Caption) Then
MsgBox "Sorry, You don't have so many Lot for sell."
Exit Sub
End If
If SellShare2 = "0" Then Exit Sub

TotalMoney = Val(Label4.Caption) * Val(SellShare2)
ConfirmSellShare = MsgBox("Sell " & SellShare2 & " Lot on " & Mid(Label1.Caption, 3) & "?" & vbCrLf & "Selling Price: $" & TotalMoney & ". Are you sure?", vbYesNo)
If ConfirmSellShare = vbNo Then Exit Sub
If ConfirmSellShare = vbYes Then
WillMoney "PLUS", TotalMoney
Label13.Caption = Val(Label13.Caption) - Val(SellShare2)
Label7.Caption = Val(LotValue) - Val(Label13.Caption)
SaveSetting "Digimon", "Share", "Own1", Label13.Caption
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Lot Sell."
Exit Sub
End If

Case "2"

On Error Resume Next
SellShare2 = InputBox(Label2.Caption & vbCrLf & vbCrLf & "Current Price: " & vbTab & "$" & Label5.Caption & vbCrLf & "Lot Left: " & vbTab & Label8.Caption & vbCrLf & "Current Own: " & vbTab & Label14.Caption & vbCrLf & vbCrLf & "How many Lot you want to sell?")
If SellShare2 > Val(Label14.Caption) Then
MsgBox "Sorry, You don't have so many Lot for sell."
Exit Sub
End If
If SellShare2 = "0" Then Exit Sub

TotalMoney = Val(Label5.Caption) * Val(SellShare2)
ConfirmSellShare = MsgBox("Sell " & SellShare2 & " Lot on " & Mid(Label2.Caption, 3) & "?" & vbCrLf & "Selling Price: $" & TotalMoney & ". Are you sure?", vbYesNo)
If ConfirmSellShare = vbNo Then Exit Sub
If ConfirmSellShare = vbYes Then
WillMoney "PLUS", TotalMoney
Label14.Caption = Val(Label14.Caption) - Val(SellShare2)
Label8.Caption = Val(LotValue) - Val(Label14.Caption)
SaveSetting "Digimon", "Share", "Own2", Label14.Caption
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Lot Sell."
Exit Sub
End If

Case "3"

On Error Resume Next
SellShare2 = InputBox(Label3.Caption & vbCrLf & vbCrLf & "Current Price: " & vbTab & "$" & Label6.Caption & vbCrLf & "Lot Left: " & vbTab & Label9.Caption & vbCrLf & "Current Own: " & vbTab & Label15.Caption & vbCrLf & vbCrLf & "How many Lot you want to sell?")
If SellShare2 > Val(Label15.Caption) Then
MsgBox "Sorry, You don't have so many Lot for sell."
Exit Sub
End If
If SellShare2 = "0" Then Exit Sub

TotalMoney = Val(Label6.Caption) * Val(SellShare2)
ConfirmSellShare = MsgBox("Sell " & SellShare2 & " Lot on " & Mid(Label3.Caption, 3) & "?" & vbCrLf & "Selling Price: $" & TotalMoney & ". Are you sure?", vbYesNo)
If ConfirmSellShare = vbNo Then Exit Sub
If ConfirmSellShare = vbYes Then
WillMoney "PLUS", TotalMoney
Label15.Caption = Val(Label15.Caption) - Val(SellShare2)
Label9.Caption = Val(LotValue) - Val(Label15.Caption)
SaveSetting "Digimon", "Share", "Own3", Label15.Caption
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")
MsgBox "Lot Sell."
Exit Sub
End If

Case ""
Exit Sub
End Select
MsgBox "Please select a legal number."

End Sub

Private Sub FlatButton3_Click()
Form31Show = 0
Me.Hide
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Form2, m_cN

Form2.Show
Unload Me
End Sub

Private Sub Flatbutton4_Click()
FlatButton4.Caption = "Soon!"
End Sub

Private Sub Form_Load()
AutoDetectTop Me
Form31Show = 1
m_cN.Detach
Set m_cN = New cNeoCaption
Skin Me, m_cN

LotValue = 500
Label4.Caption = GetSetting("Digimon", "Share", "Price1")
Label5.Caption = GetSetting("Digimon", "Share", "Price2")
Label6.Caption = GetSetting("Digimon", "Share", "Price3")
Label13.Caption = GetSetting("Digimon", "Share", "Own1")
Label14.Caption = GetSetting("Digimon", "Share", "Own2")
Label15.Caption = GetSetting("Digimon", "Share", "Own3")
Label7.Caption = Val(LotValue) - Val(GetSetting("Digimon", "Share", "Own1"))
Label8.Caption = Val(LotValue) - Val(GetSetting("Digimon", "Share", "Own2"))
Label9.Caption = Val(LotValue) - Val(GetSetting("Digimon", "Share", "Own3"))
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")

Select Case Val(Label4.Caption)
Case Is > 100
Label1.ForeColor = vbGreen
Label4.ForeColor = vbGreen
Label7.ForeColor = vbGreen
Label13.ForeColor = vbGreen
Case 20 To 100
Label1.ForeColor = vbYellow
Label4.ForeColor = vbYellow
Label7.ForeColor = vbYellow
Label13.ForeColor = vbYellow
Case Is < 20
Label1.ForeColor = vbRed
Label4.ForeColor = vbRed
Label7.ForeColor = vbRed
Label13.ForeColor = vbRed
End Select

Select Case Val(Label5.Caption)
Case Is > 100
Label2.ForeColor = vbGreen
Label5.ForeColor = vbGreen
Label8.ForeColor = vbGreen
Label14.ForeColor = vbGreen
Case 20 To 100
Label2.ForeColor = vbYellow
Label5.ForeColor = vbYellow
Label8.ForeColor = vbYellow
Label14.ForeColor = vbYellow
Case Is < 20
Label2.ForeColor = vbRed
Label5.ForeColor = vbRed
Label8.ForeColor = vbRed
Label14.ForeColor = vbRed
End Select

Select Case Val(Label6.Caption)
Case Is > 100
Label3.ForeColor = vbGreen
Label6.ForeColor = vbGreen
Label9.ForeColor = vbGreen
Label15.ForeColor = vbGreen
Case 20 To 100
Label3.ForeColor = vbYellow
Label6.ForeColor = vbYellow
Label9.ForeColor = vbYellow
Label15.ForeColor = vbYellow
Case Is < 20
Label3.ForeColor = vbRed
Label6.ForeColor = vbRed
Label9.ForeColor = vbRed
Label15.ForeColor = vbRed
End Select
End Sub

Private Sub Timer1_Timer()
Label4.Caption = GetSetting("Digimon", "Share", "Price1")
Label5.Caption = GetSetting("Digimon", "Share", "Price2")
Label6.Caption = GetSetting("Digimon", "Share", "Price3")
Label13.Caption = GetSetting("Digimon", "Share", "Own1")
Label14.Caption = GetSetting("Digimon", "Share", "Own2")
Label15.Caption = GetSetting("Digimon", "Share", "Own3")
Label7.Caption = Val(LotValue) - Val(GetSetting("Digimon", "Share", "Own1"))
Label8.Caption = Val(LotValue) - Val(GetSetting("Digimon", "Share", "Own2"))
Label9.Caption = Val(LotValue) - Val(GetSetting("Digimon", "Share", "Own3"))
Label17.Caption = "Money: $ " & GetSetting("Digimon", "Profile", "Money")

Select Case Val(Label4.Caption)
Case Is > 100
Label1.ForeColor = vbGreen
Label4.ForeColor = vbGreen
Label7.ForeColor = vbGreen
Label13.ForeColor = vbGreen
Case 20 To 100
Label1.ForeColor = vbYellow
Label4.ForeColor = vbYellow
Label7.ForeColor = vbYellow
Label13.ForeColor = vbYellow
Case Is < 20
Label1.ForeColor = vbRed
Label4.ForeColor = vbRed
Label7.ForeColor = vbRed
Label13.ForeColor = vbRed
End Select

Select Case Val(Label5.Caption)
Case Is > 100
Label2.ForeColor = vbGreen
Label5.ForeColor = vbGreen
Label8.ForeColor = vbGreen
Label14.ForeColor = vbGreen
Case 20 To 100
Label2.ForeColor = vbYellow
Label5.ForeColor = vbYellow
Label8.ForeColor = vbYellow
Label14.ForeColor = vbYellow
Case Is < 20
Label2.ForeColor = vbRed
Label5.ForeColor = vbRed
Label8.ForeColor = vbRed
Label14.ForeColor = vbRed
End Select

Select Case Val(Label6.Caption)
Case Is > 100
Label3.ForeColor = vbGreen
Label6.ForeColor = vbGreen
Label9.ForeColor = vbGreen
Label15.ForeColor = vbGreen
Case 20 To 100
Label3.ForeColor = vbYellow
Label6.ForeColor = vbYellow
Label9.ForeColor = vbYellow
Label15.ForeColor = vbYellow
Case Is < 20
Label3.ForeColor = vbRed
Label6.ForeColor = vbRed
Label9.ForeColor = vbRed
Label15.ForeColor = vbRed
End Select

End Sub
