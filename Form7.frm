VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICON"
   ClientHeight    =   4965
   ClientLeft      =   -7515
   ClientTop       =   1680
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Picture"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   4215
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         Picture         =   "Form7.frx":08CA
         ScaleHeight     =   735
         ScaleWidth      =   2415
         TabIndex        =   4
         Top             =   1320
         Width           =   2415
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         Picture         =   "Form7.frx":108D
         ScaleHeight     =   255
         ScaleWidth      =   1215
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   240
         Picture         =   "Form7.frx":14C6
         ScaleHeight     =   735
         ScaleWidth      =   3615
         TabIndex        =   2
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "ICONS"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Image Image48 
         Height          =   480
         Left            =   3480
         Picture         =   "Form7.frx":235C
         Tag             =   "Scyther"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image47 
         Height          =   480
         Left            =   3000
         Picture         =   "Form7.frx":2C26
         Tag             =   "Farfetch'd"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image46 
         Height          =   480
         Left            =   2520
         Picture         =   "Form7.frx":34F0
         Tag             =   "Lapras"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image45 
         Height          =   480
         Left            =   2040
         Picture         =   "Form7.frx":3DBA
         Tag             =   "Haunter"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image44 
         Height          =   480
         Left            =   1560
         Picture         =   "Form7.frx":497C
         Tag             =   "Growlithe"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image43 
         Height          =   480
         Left            =   1080
         Picture         =   "Form7.frx":5246
         Tag             =   "Golduck"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image42 
         Height          =   480
         Left            =   600
         Picture         =   "Form7.frx":5B10
         Tag             =   "Sandslash"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image41 
         Height          =   480
         Left            =   120
         Picture         =   "Form7.frx":63DA
         Tag             =   "Hakuryu"
         Top             =   2760
         Width           =   480
      End
      Begin VB.Image Image40 
         Height          =   480
         Left            =   3480
         Picture         =   "Form7.frx":6F9C
         Tag             =   "Nidoran"
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image39 
         Height          =   480
         Left            =   3000
         Picture         =   "Form7.frx":7866
         Tag             =   "0"
         Top             =   2280
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image38 
         Height          =   480
         Left            =   2520
         Picture         =   "Form7.frx":7B70
         Tag             =   "0"
         Top             =   2280
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   600
         Picture         =   "Form7.frx":7E7A
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   120
         Picture         =   "Form7.frx":8744
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   2040
         Picture         =   "Form7.frx":900E
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   1080
         Picture         =   "Form7.frx":98D8
         ToolTipText     =   "Heal"
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1560
         Picture         =   "Form7.frx":A1A2
         ToolTipText     =   "Increase Defence Value by 10%"
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   3000
         Picture         =   "Form7.frx":AA6C
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   3480
         Top             =   1800
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2520
         Picture         =   "Form7.frx":B336
         Tag             =   "Unknown Person"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image37 
         Height          =   480
         Left            =   2040
         Picture         =   "Form7.frx":B488
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   120
         Picture         =   "Form7.frx":C152
         Tag             =   "Angemon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   600
         Picture         =   "Form7.frx":CA1C
         Tag             =   "Bakemon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   1080
         Picture         =   "Form7.frx":CD26
         Tag             =   "Botamon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   1560
         Picture         =   "Form7.frx":D9F0
         Tag             =   "Bukamon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   2040
         Picture         =   "Form7.frx":E6BA
         Tag             =   "Coelamon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   2520
         Picture         =   "Form7.frx":EF84
         Tag             =   "Deltamon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image15 
         Height          =   480
         Left            =   3000
         Picture         =   "Form7.frx":F84E
         Tag             =   "Dijitama"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   3480
         Picture         =   "Form7.frx":FB58
         Tag             =   "Elecmon"
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image17 
         Height          =   480
         Left            =   120
         Picture         =   "Form7.frx":FE62
         Tag             =   "Etemon"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image18 
         Height          =   480
         Left            =   600
         Picture         =   "Form7.frx":1072C
         Tag             =   "Garurumo"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image19 
         Height          =   480
         Left            =   1080
         Picture         =   "Form7.frx":10FF6
         Tag             =   "Greymon"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image20 
         Height          =   480
         Left            =   1560
         Picture         =   "Form7.frx":118C0
         Tag             =   "Kunemon"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image21 
         Height          =   480
         Left            =   2040
         Picture         =   "Form7.frx":11BCA
         Tag             =   "Kuwagamo"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image22 
         Height          =   480
         Left            =   2520
         Picture         =   "Form7.frx":11ED4
         Tag             =   "M_tyrano"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Left            =   3000
         Picture         =   "Form7.frx":121DE
         Tag             =   "Megadra"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image24 
         Height          =   480
         Left            =   3480
         Picture         =   "Form7.frx":124E8
         Tag             =   "Orgemon"
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image25 
         Height          =   480
         Left            =   120
         Picture         =   "Form7.frx":12DB2
         Tag             =   "Patamon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image26 
         Height          =   480
         Left            =   600
         Picture         =   "Form7.frx":130BC
         Tag             =   "Piyomon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image27 
         Height          =   480
         Left            =   1080
         Picture         =   "Form7.frx":133C6
         Tag             =   "Poyomon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image28 
         Height          =   480
         Left            =   1560
         Picture         =   "Form7.frx":14090
         Tag             =   "Punimon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image29 
         Height          =   480
         Left            =   2040
         Picture         =   "Form7.frx":14D5A
         Tag             =   "Tanemon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image30 
         Height          =   480
         Left            =   2520
         Picture         =   "Form7.frx":15A24
         Tag             =   "Tokomon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image31 
         Height          =   480
         Left            =   3000
         Picture         =   "Form7.frx":166EE
         Tag             =   "Tsunomon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image32 
         Height          =   480
         Left            =   3480
         Picture         =   "Form7.frx":173B8
         Tag             =   "Tuskmon"
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image33 
         Height          =   480
         Left            =   120
         Picture         =   "Form7.frx":17C82
         Tag             =   "Tyrano"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image34 
         Height          =   480
         Left            =   600
         Picture         =   "Form7.frx":17F8C
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image35 
         Height          =   480
         Left            =   1080
         Picture         =   "Form7.frx":18C56
         Tag             =   "Angelwoman"
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image36 
         Height          =   480
         Left            =   1560
         Picture         =   "Form7.frx":19920
         Top             =   1800
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
