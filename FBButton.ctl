VERSION 5.00
Begin VB.UserControl FlatButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   KeyPreview      =   -1  'True
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   91
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picImage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer tmrHighlight 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   840
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblCaption"
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "FlatButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================
'FlatButton Usage :
'
'The use of this control is fairly straightforward, but the use of images
'warrants some explanation.
'
'To use a picture you need to set the PicturehDC property of the FlatButton
'to the hDC property of a PictureBox (or other compatible hDC).
'
'Example...
'Public Sub Form_Load()
'Me.FlatButton1.PicturehDC = Me.Picture1.hdc
'End Sub
'
'The following picture box (Picture1) properties should be set...
'Appearance : 0 - Flat
'AutoRedraw : True
'AutoSize : True
'BackColor : Button Face
'BorderStyle : 0 - None
'Picture : Set to the picture you want to display on the FlatButton
'
'
'To show the full picture on the FlatButton set the PictureHeight and PictureWidth
'properties to the height and width of the picture in pixels. The default 16x16
'supports small icons.
'
'*** Any questions : please feel free to email wrhartel@camtech.net.au ***
'
'========================================================================

'========================================================================
'Notes : 6th August, 2000
'
'The UserControl AccessKeyPress event does not fire if the access key of one
'of the UserControl's constituent control's access keys is pressed. As such,
'when the lblCaption access key is pressed, the UserControl only receives an
'EnterFocus event (as lblCaption takes the focus), and the UserControl does
'not receive an AccessKeyPress event.
'
'To get the FlatButton user control to emulate CommandButton behaviour w.r.t.
'access keys (i.e., generate a click event when the user presses 'ALT' plus
'one of the control's access keys) it would be necessary to disable the
'UseMnemonic property of lblCaption, and in code REMOVE any '&' characters
'from the lblCaption caption property and assign them as UserControl Access
'keys. Next a line would need to be drawn underneath the appropriate characters
'in lblCaption by some other means (so the control user would still know what
'the access keys were). Now if the user presses 'ALT' plus an access key the
'UserControl would receive an AccessKeyPress event since lblCaption no longer
'has 'real' access keys.
'
'Without this code, and leaving the UseMnemonic property of lblCaption set to
'True, a user pressing 'ALT' plus an access key will tab to the UserControl
'and give it focus, but will not generate a click event. Since this is not
'a major concern (just a small detail) this code has not yet been written.
'========================================================================

'========================================================================
'DLL Base Address
'========================================================================
'&H6880000, based on 16M (16,777,216 or &H1000000) plus 1416 * 16K

'========================================================================
'Windows API Types
'========================================================================
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'========================================================================
'Windows API Declarations
'========================================================================
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint _
As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) _
As Long
Private Declare Function DrawEdge Lib "user32.dll" (ByVal hdc As Long, _
qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, _
lpRect As RECT) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, _
ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As _
Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, _
ByVal dwRop As Long) As Long
    
'========================================================================
'Enumerations
'========================================================================
Public Enum fbAlignment
    fbLeft = 0
    fbRight
    fbCenter
End Enum
    
'========================================================================
'Constants
'========================================================================
Private Const BDR_RAISEDINNER As Long = &H4
Private Const BDR_SUNKENOUTER As Long = &H2
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA
Private Const BDR_MOUSEOVER As Long = BDR_RAISEDINNER
Private Const BDR_MOUSEDOWN As Long = BDR_SUNKENOUTER
Private Const BDR_MOUSEOVER_HB As Long = BDR_RAISED
Private Const BDR_MOUSEDOWN_HB As Long = BDR_SUNKEN
Private Const BF_BOTTOM As Long = &H8
Private Const BF_LEFT As Long = &H1
Private Const BF_RIGHT As Long = &H4
Private Const BF_TOP As Long = &H2
'Bitwise comparison
Private Const BF_RECT As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const DUD_VALUE As Integer = -1
Private Const NOT_APPLY_ALL As Integer = 0
Private Const APPLY_ALL As Integer = 1
Private Const INIT_PROP_FLAG As Integer = 0
Private Const READ_PROP_FLAG As Integer = 1
Private Const FORCE_FLATTEN As Integer = 1
Private Const FOCUS_RECT_OFFSET As Integer = 4
    
Private Const mDef_lngForeColor As Long = vbBlack
Private Const mDef_lngBackColor As Long = vbButtonFace
Private Const mDef_lngHoverColor As Long = vbHighlight
Private Const mDef_fbAlignment As Integer = fbAlignment.fbCenter
Private Const mDef_booHasBorder As Boolean = False
Private Const mDef_strCaption As String = "FlatButton"
Private Const mDef_booEnabled As Boolean = True
Private Const mDef_booHasFocusRect As Boolean = True
Private Const mDef_booAlignPicLeft As Boolean = True
Private Const mDef_intPictureWidth As Integer = 16
Private Const mDef_intPictureHeight As Integer = mDef_intPictureWidth
Private Const mDef_lngPicturehDC As Long = 0
Private Const mDef_booHasPicture As Boolean = False
Private Const mDef_booHasCaption As Boolean = True
    
Private Const FORECOLOR_PROPERTY_NAME As String = "ForeColor"
Private Const ALIGNMENT_PROPERTY_NAME As String = "Alignment"
Private Const HOVERCOLOR_PROPERTY_NAME As String = "HoverColor"
Private Const ENABLED_PROPERTY_NAME As String = "Enabled"
Private Const FONT_PROPERTY_NAME As String = "Font"
Private Const HASFOCUSRECT_PROPERTY_NAME As String = "HasFocusRect"
Private Const CAPTION_PROPERTY_NAME As String = "Caption"
Private Const BACKCOLOR_PROPERTY_NAME As String = "BackColor"
Private Const HASBORDER_PROPERTY_NAME As String = "HasBorder"
Private Const ALIGNPICLEFT_PROPERTY_NAME As String = "AlignPicLeft"
Private Const PICTUREWIDTH_PROPERTY_NAME As String = "PictureWidth"
Private Const PICTUREHEIGHT_PROPERTY_NAME As String = "PictureHeight"
Private Const HASPICTURE_PROPERTY_NAME As String = "HasPicture"
Private Const HASCAPTION_PROPERTY_NAME As String = "HasCaption"
Private Const PICTUREHDC_PROPERTY_NAME As String = "PicturehDC"
    
'========================================================================
'Variables
'========================================================================
Private mprop_lngForeColor As Long
Private mProp_lngHoverColor As Long
Private mProp_lngBackColor As Long
Private mProp_fbAlignment As fbAlignment
Private mProp_booHasBorder As Boolean
Private mProp_strCaption As String
Private mProp_booEnabled As Boolean
Private mProp_booHasFocusRect As Boolean
Private mProp_fntFont As StdFont
Private mProp_booAlignPicLeft As Boolean
Private mProp_intPictureHeight As Integer
Private mProp_intPictureWidth As Integer
Private mProp_lngPicturehDC As Long
Private mProp_booHasPicture As Boolean
Private mProp_booHasCaption As Boolean

Private mbooHasCapture As Boolean
Private mpntLabelPos As POINTAPI
Private mpntOldSize As POINTAPI
Private mpntPicPos As POINTAPI
Private intPropertiesKnown As Integer

Event Click()

'========================================================================
'UserControl Enter/Exit Focus
'========================================================================
Public Sub UserControl_EnterFocus()

Dim rctFocus As RECT

If Not mProp_booHasFocusRect Then Exit Sub

'Draw a focus rectangle
rctFocus.Left = FOCUS_RECT_OFFSET
rctFocus.Top = FOCUS_RECT_OFFSET
rctFocus.Right = ScaleWidth - FOCUS_RECT_OFFSET
rctFocus.Bottom = ScaleHeight - FOCUS_RECT_OFFSET
DrawFocusRect hdc, rctFocus

End Sub

Public Sub UserControl_ExitFocus()
'Remove the focus rectangle
If mProp_booHasFocusRect Then Line (FOCUS_RECT_OFFSET, FOCUS_RECT_OFFSET)- _
(ScaleWidth - FOCUS_RECT_OFFSET - 1, ScaleHeight - FOCUS_RECT_OFFSET - 1), _
mProp_lngBackColor, B
End Sub

'========================================================================
'UserControl Initialize/InitProprties
'========================================================================
Public Sub UserControl_Initialize()
tmrHighlight.Enabled = False
tmrHighlight.Interval = 100
End Sub

Public Sub UserControl_InitProperties()

UserControl.Width = 1095
UserControl.Height = 390

mprop_lngForeColor = mDef_lngForeColor
mProp_fbAlignment = mDef_fbAlignment
mProp_booAlignPicLeft = mDef_booAlignPicLeft
mProp_intPictureWidth = mDef_intPictureWidth
mProp_intPictureHeight = mDef_intPictureHeight
mProp_lngPicturehDC = mDef_lngPicturehDC
mProp_booHasCaption = mDef_booHasCaption
mProp_booHasPicture = mDef_booHasPicture
mProp_booHasBorder = mDef_booHasBorder
mProp_lngBackColor = mDef_lngBackColor
mProp_strCaption = mDef_strCaption
mProp_booEnabled = mDef_booEnabled
mProp_booHasFocusRect = mDef_booHasFocusRect
mProp_lngHoverColor = mDef_lngHoverColor

Set mProp_fntFont = Ambient.Font
intPropertiesKnown = 1

ApplyAllProperties INIT_PROP_FLAG

End Sub

'========================================================================
'UserControl Property Bag Stuff
'========================================================================
Public Sub UserControl_ReadProperties(PropBag As PropertyBag)

With PropBag
    mprop_lngForeColor = .ReadProperty(FORECOLOR_PROPERTY_NAME, mDef_lngForeColor)
    mProp_fbAlignment = .ReadProperty(ALIGNMENT_PROPERTY_NAME, mDef_fbAlignment)
    mProp_booAlignPicLeft = .ReadProperty(ALIGNPICLEFT_PROPERTY_NAME, mDef_booAlignPicLeft)
    mProp_intPictureWidth = .ReadProperty(PICTUREWIDTH_PROPERTY_NAME, mDef_intPictureWidth)
    mProp_intPictureHeight = .ReadProperty(PICTUREHEIGHT_PROPERTY_NAME, mDef_intPictureHeight)
    mProp_lngPicturehDC = .ReadProperty(PICTUREHDC_PROPERTY_NAME, mDef_lngPicturehDC)
    mProp_booHasPicture = .ReadProperty(HASPICTURE_PROPERTY_NAME, mDef_booHasPicture)
    mProp_booHasCaption = .ReadProperty(HASCAPTION_PROPERTY_NAME, mDef_booHasCaption)
    mProp_booHasBorder = .ReadProperty(HASBORDER_PROPERTY_NAME, mDef_booHasBorder)
    mProp_lngBackColor = .ReadProperty(BACKCOLOR_PROPERTY_NAME, mDef_lngBackColor)
    mProp_strCaption = .ReadProperty(CAPTION_PROPERTY_NAME, mDef_strCaption)
    mProp_booEnabled = .ReadProperty(ENABLED_PROPERTY_NAME, mDef_booEnabled)
    mProp_booHasFocusRect = .ReadProperty(HASFOCUSRECT_PROPERTY_NAME, mDef_booHasFocusRect)
    Set mProp_fntFont = .ReadProperty(FONT_PROPERTY_NAME, Ambient.Font)
    mProp_lngHoverColor = .ReadProperty(HOVERCOLOR_PROPERTY_NAME, mDef_lngHoverColor)
End With

intPropertiesKnown = 1

ApplyAllProperties READ_PROP_FLAG

If Ambient.UserMode Then 'Runtime only
    If mProp_booHasBorder Then
        ApplyBorder FORCE_FLATTEN
    End If
    tmrHighlight.Enabled = True
End If

End Sub

Public Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty FORECOLOR_PROPERTY_NAME, mprop_lngForeColor, mDef_lngForeColor
    .WriteProperty ALIGNMENT_PROPERTY_NAME, mProp_fbAlignment, mDef_fbAlignment
    .WriteProperty ALIGNPICLEFT_PROPERTY_NAME, mProp_booAlignPicLeft, mDef_booAlignPicLeft
    .WriteProperty PICTUREWIDTH_PROPERTY_NAME, mProp_intPictureWidth, mDef_intPictureWidth
    .WriteProperty PICTUREHEIGHT_PROPERTY_NAME, mProp_intPictureHeight, mDef_intPictureHeight
    .WriteProperty PICTUREHDC_PROPERTY_NAME, mProp_lngPicturehDC, mDef_lngPicturehDC
    .WriteProperty HASPICTURE_PROPERTY_NAME, mProp_booHasPicture, mDef_booHasPicture
    .WriteProperty HASCAPTION_PROPERTY_NAME, mProp_booHasCaption, mDef_booHasCaption
    .WriteProperty HASBORDER_PROPERTY_NAME, mProp_booHasBorder, mDef_booHasBorder
    .WriteProperty BACKCOLOR_PROPERTY_NAME, mProp_lngBackColor, mDef_lngBackColor
    .WriteProperty CAPTION_PROPERTY_NAME, mProp_strCaption, mDef_strCaption
    .WriteProperty ENABLED_PROPERTY_NAME, mProp_booEnabled, mDef_booEnabled
    .WriteProperty HASFOCUSRECT_PROPERTY_NAME, mProp_booHasFocusRect, mDef_booHasFocusRect
    .WriteProperty FONT_PROPERTY_NAME, mProp_fntFont, Ambient.Font
    .WriteProperty HOVERCOLOR_PROPERTY_NAME, mProp_lngHoverColor, mDef_lngHoverColor
End With
End Sub

'========================================================================
'Key Events
'========================================================================
Public Sub UserControl_AccessKeyPress(KeyAscii As Integer)
RaiseEvent Click
End Sub

Public Sub UserControl_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Or KeyAscii = vbKeyReturn Then
    RaiseEvent Click
End If
End Sub

Public Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
    UserControl_MouseDown vbLeftButton, DUD_VALUE, DUD_VALUE, DUD_VALUE
End If
End Sub

Public Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
    UserControl_MouseUp DUD_VALUE, DUD_VALUE, DUD_VALUE, DUD_VALUE
End If
End Sub

'========================================================================
'MouseDown Events
'========================================================================
Public Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Const OFFSET As Integer = 1
Dim rctBtn As RECT

If Button = vbLeftButton Then
    tmrHighlight.Enabled = False
    lblCaption.Left = mpntLabelPos.x + OFFSET
    lblCaption.Top = mpntLabelPos.y + OFFSET
    picImage.Move mpntPicPos.x + OFFSET, mpntPicPos.y + OFFSET, _
    picImage.Width, picImage.Height
    Line (0, 0)-(Width, Height), mProp_lngBackColor, B
    rctBtn.Left = 0
    rctBtn.Top = 0
    rctBtn.Right = ScaleWidth
    rctBtn.Bottom = ScaleHeight
    If mProp_booHasBorder = True Then
        DrawEdge hdc, rctBtn, BDR_MOUSEDOWN_HB, BF_RECT
    Else
        DrawEdge hdc, rctBtn, BDR_MOUSEDOWN, BF_RECT
    End If
End If

End Sub

Public Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseDown Button, Shift, x, y
End Sub

Public Sub picImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseDown Button, Shift, x, y
End Sub

'========================================================================
'MouseUp Events
'========================================================================
Public Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Dim pntCursor As POINTAPI

lblCaption.Left = mpntLabelPos.x
lblCaption.Top = mpntLabelPos.y
picImage.Move mpntPicPos.x, mpntPicPos.y, picImage.Width, picImage.Height
GetCursorPos pntCursor
If WindowFromPoint(pntCursor.x, pntCursor.y) = hwnd Or _
WindowFromPoint(pntCursor.x, pntCursor.y) = picImage.hwnd Or _
mProp_booHasBorder Then
    ApplyBorder
    mbooHasCapture = True
Else
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mProp_lngBackColor, B
    mbooHasCapture = False
End If

tmrHighlight.Enabled = True

RaiseEvent Click

End Sub

Public Sub lblCaption_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseUp Button, Shift, x, y
End Sub

Public Sub picImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl_MouseUp Button, Shift, x, y
End Sub

'========================================================================
'Other UserControl Events
'========================================================================
Public Sub UserControl_Resize()

If intPropertiesKnown = 0 Then Exit Sub

Cls
RedrawControl

'Design-time or Has Border
If Not Ambient.UserMode Or mProp_booHasBorder Then
    ApplyBorder FORCE_FLATTEN
End If

End Sub

Public Sub UserControl_AmbientChanged(PropertyName As String)
If UCase$(PropertyName) = "BACKCOLOR" Then
    BackColor = Ambient.BackColor
End If
End Sub

'========================================================================
'Other Object Events
'========================================================================
Public Sub picImage_Paint()

If mProp_booEnabled = True Then
    'Draw picture
    BitBlt picImage.hdc, 0, 0, mProp_intPictureWidth, _
    mProp_intPictureHeight, mProp_lngPicturehDC, 0, 0, vbSrcCopy
Else
    'Draw picture (incase button begins its life disabled)
    BitBlt picImage.hdc, 0, 0, mProp_intPictureWidth, _
    mProp_intPictureHeight, mProp_lngPicturehDC, 0, 0, vbSrcCopy
    'Draw dimmed (darker) picture
    BitBlt picImage.hdc, 0, 0, mProp_intPictureWidth, _
    mProp_intPictureHeight, picBuffer.hdc, 0, 0, vbSrcAnd
End If

End Sub

Public Sub tmrHighlight_Timer()

Dim pntCursor As POINTAPI

GetCursorPos pntCursor

'If mouse is over this control
If WindowFromPoint(pntCursor.x, pntCursor.y) = hwnd Or _
WindowFromPoint(pntCursor.x, pntCursor.y) = picImage.hwnd Then
    If Not mbooHasCapture Then
        ApplyBorder
        lblCaption.ForeColor = mProp_lngHoverColor
        mbooHasCapture = True
    End If
Else
    If mbooHasCapture Then
        'Remove thick edge
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mProp_lngBackColor, B
        lblCaption.ForeColor = mprop_lngForeColor
        mbooHasCapture = False
    End If
End If

End Sub

'========================================================================
'Properties requiring Apply**** to be called
'========================================================================
Public Property Get HasBorder() As Boolean
Attribute HasBorder.VB_Description = "Sets/returns whether the FlatButton is drawn with a border at runtime."
    HasBorder = mProp_booHasBorder
End Property

Public Property Let HasBorder(ByVal booNewValue As Boolean)
If Ambient.UserMode Then 'Design-time only
    Err.Raise 383
Else
    mProp_booHasBorder = booNewValue
    PropertyChanged HASBORDER_PROPERTY_NAME
End If
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
BackColor = mProp_lngBackColor
End Property

Public Property Let BackColor(ByVal oleNewValue As OLE_COLOR)
mProp_lngBackColor = oleNewValue
ApplyBackColor
ApplyBorder
PropertyChanged BACKCOLOR_PROPERTY_NAME
End Property

Public Property Get Alignment() As fbAlignment
Attribute Alignment.VB_Description = "Returns/sets the FlatButton control's caption alignment."
Alignment = mProp_fbAlignment
End Property

Public Property Let Alignment(ByVal fbNewValue As fbAlignment)
mProp_fbAlignment = fbNewValue
ApplyCaption
PropertyChanged ALIGNMENT_PROPERTY_NAME
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Caption = mProp_strCaption
End Property

Public Property Let Caption(ByVal strNewValue As String)
mProp_strCaption = strNewValue
ApplyCaption
PropertyChanged CAPTION_PROPERTY_NAME
End Property

Public Property Get HasFocusRect() As Boolean
Attribute HasFocusRect.VB_Description = "Read-only at runtime. Set/returns whether a focus rectangle is drawn on the FlatButton when it has focus."
HasFocusRect = mProp_booHasFocusRect
End Property

Public Property Let HasFocusRect(ByVal booNewValue As Boolean)
If Ambient.UserMode Then 'Design-time only
    Err.Raise 383
Else
    mProp_booHasFocusRect = booNewValue
    PropertyChanged HASFOCUSRECT_PROPERTY_NAME
End If
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."
Set Font = mProp_fntFont
End Property

Public Property Set Font(ByVal fntNewValue As StdFont)
Set mProp_fntFont = fntNewValue
ApplyFont
PropertyChanged FONT_PROPERTY_NAME
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_UserMemId = -514
Enabled = mProp_booEnabled
End Property

Public Property Let Enabled(ByVal booNewValue As Boolean)
mProp_booEnabled = booNewValue
ApplyEnabled
PropertyChanged ENABLED_PROPERTY_NAME
End Property

Public Property Get HoverColor() As OLE_COLOR
Attribute HoverColor.VB_Description = "Returns/sets the color of the FlatButton caption text when the mouse pointer is over the control."
HoverColor = mProp_lngHoverColor
End Property

Public Property Let HoverColor(ByVal oleNewValue As OLE_COLOR)
mProp_lngHoverColor = oleNewValue
PropertyChanged HOVERCOLOR_PROPERTY_NAME
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the FlatButtons foreground color which is used to display the button caption."
ForeColor = mprop_lngForeColor
End Property

Public Property Let ForeColor(ByVal oleNewValue As OLE_COLOR)
mprop_lngForeColor = oleNewValue
ApplyCaption
PropertyChanged FORECOLOR_PROPERTY_NAME
End Property

'========================================================================
'Properties requiring UserControl_Resize to be called
'========================================================================
Public Property Get HasPicture() As Boolean
Attribute HasPicture.VB_Description = "Returns/sets whether a picture is used on the FlatButton."
HasPicture = mProp_booHasPicture
End Property

Public Property Let HasPicture(ByVal booNewValue As Boolean)
mProp_booHasPicture = booNewValue
PropertyChanged HASPICTURE_PROPERTY_NAME
UserControl_Resize
End Property

Public Property Get HasCaption() As Boolean
Attribute HasCaption.VB_Description = "Returns/sets whether a text caption is used on the FlatButton."
HasCaption = mProp_booHasCaption
End Property

Public Property Let HasCaption(ByVal booNewValue As Boolean)
mProp_booHasCaption = booNewValue
PropertyChanged HASCAPTION_PROPERTY_NAME
UserControl_Resize
End Property

Public Property Get AlignPicLeft() As Boolean
Attribute AlignPicLeft.VB_Description = "Specifies whether to align the FlatButton's picture to the left hand side of the button."
AlignPicLeft = mProp_booAlignPicLeft
End Property

Public Property Let AlignPicLeft(ByVal booNewValue As Boolean)
mProp_booAlignPicLeft = booNewValue
PropertyChanged ALIGNPICLEFT_PROPERTY_NAME
UserControl_Resize
End Property

Public Property Get PicturehDC() As Long
Attribute PicturehDC.VB_Description = "Returns/sets the handle to a device context used as the source device context for the FlatButton picture."
PicturehDC = mProp_lngPicturehDC
End Property

Public Property Let PicturehDC(ByVal lngNewValue As Long)
mProp_lngPicturehDC = lngNewValue
PropertyChanged PICTUREHDC_PROPERTY_NAME
UserControl_Resize
End Property

Public Property Get PictureHeight() As Integer
Attribute PictureHeight.VB_Description = "Returns/sets the height in pixels of the source device context."
PictureHeight = mProp_intPictureHeight
End Property

Public Property Let PictureHeight(ByVal intNewValue As Integer)
mProp_intPictureHeight = intNewValue
PropertyChanged PICTUREHEIGHT_PROPERTY_NAME
UserControl_Resize
End Property

Public Property Get PictureWidth() As Integer
Attribute PictureWidth.VB_Description = "Returns/sets the width in pixels of the source device context."
PictureWidth = mProp_intPictureWidth
End Property

Public Property Let PictureWidth(ByVal intNewValue As Integer)
mProp_intPictureWidth = intNewValue
PropertyChanged PICTUREWIDTH_PROPERTY_NAME
UserControl_Resize
End Property

'========================================================================
'Public Subroutines
'========================================================================
Public Sub RedrawControl()

Dim intX As Integer
Dim intLabelTop As Integer
Dim intBadPicSizeFlag As Integer

'Check that the picture has a valid size
If mProp_booHasPicture = True Then
    If mProp_intPictureWidth = 0 Or mProp_intPictureHeight = 0 Then
        intBadPicSizeFlag = 1
        mProp_booHasPicture = False 'Temporarily disable
    Else
        picImage.Width = mProp_intPictureWidth
        picImage.Height = mProp_intPictureHeight
        picBuffer.Width = mProp_intPictureWidth
        picBuffer.Height = mProp_intPictureWidth
    End If
End If

lblCaption.AutoSize = True
intLabelTop = (ScaleHeight / 2) - (lblCaption.Height / 2)

picImage.Top = (ScaleHeight / 2) - (picImage.Height / 2)

intX = (ScaleWidth - picImage.Width - lblCaption.Width) / 3

If mProp_booHasPicture = True And mProp_booHasCaption = True Then

    lblCaption.Top = intLabelTop

    If mProp_booAlignPicLeft = False Then
        picImage.Left = intX
        lblCaption.Visible = True
        lblCaption.Left = 2 * intX + picImage.Width
    Else
        picImage.Left = 6
        lblCaption.Visible = True
        lblCaption.Left = 6 + picImage.Width + _
        (ScaleWidth - 6 - picImage.Width - lblCaption.Width) / 2
    End If
    
    picImage.Visible = True
    picImage_Paint

ElseIf mProp_booHasPicture = False And mProp_booHasCaption = True Then
    
    lblCaption.AutoSize = False
    lblCaption.Move 5, intLabelTop, ScaleWidth - 10, ScaleHeight
    lblCaption.Visible = True
    
    picImage.Visible = False

ElseIf mProp_booHasPicture = True And mProp_booHasCaption = False Then
    
    lblCaption.Visible = False
    picImage.Left = (ScaleWidth / 2) - (picImage.Width / 2)
    
    picImage.Visible = True
    picImage_Paint

Else
    
    picImage.Visible = False
    lblCaption.Visible = False

End If

'Restore the HasPicture property if required
If intBadPicSizeFlag = 1 Then mProp_booHasPicture = True

mpntLabelPos.x = lblCaption.Left
mpntLabelPos.y = lblCaption.Top
mpntPicPos.x = picImage.Left
mpntPicPos.y = picImage.Top
mpntOldSize.x = ScaleWidth
mpntOldSize.y = ScaleHeight

End Sub

Public Sub ApplyAllProperties(ByVal intCallFlag As Integer)
ApplyBackColor
ApplyCaption APPLY_ALL
ApplyFont APPLY_ALL
ApplyEnabled APPLY_ALL
If intCallFlag = READ_PROP_FLAG Then UserControl_Resize
End Sub

Public Sub ApplyBackColor()
UserControl.BackColor = mProp_lngBackColor
End Sub

Public Sub ApplyCaption(Optional ByVal intApplyAll As Integer = NOT_APPLY_ALL)

Dim lngA As Long

AccessKeys = ""

For lngA = Len(mProp_strCaption) To 1 Step -1
    If Mid$(mProp_strCaption, lngA, 1) = "&" Then
        If lngA = 1 Then
            AccessKeys = Mid$(mProp_strCaption, lngA + 1, 1)
        ElseIf Not Mid$(mProp_strCaption, lngA - 1, 1) = "&" Then
            AccessKeys = Mid$(mProp_strCaption, lngA + 1, 1)
            Exit For
        Else
            lngA = lngA - 1
        End If
    End If
Next

With lblCaption
    .Caption = mProp_strCaption
    .Alignment = mProp_fbAlignment
    .ForeColor = mprop_lngForeColor
End With

If intApplyAll = NOT_APPLY_ALL Then UserControl_Resize

End Sub

Public Sub ApplyFont(Optional ByVal intApplyAll As Integer = NOT_APPLY_ALL)
Set UserControl.Font = mProp_fntFont
Set lblCaption.Font = mProp_fntFont
If intApplyAll = NOT_APPLY_ALL Then UserControl_Resize
End Sub

Public Sub ApplyEnabled(Optional ByVal intApplyAll As Integer = NOT_APPLY_ALL)
lblCaption.Enabled = mProp_booEnabled
UserControl.Enabled = mProp_booEnabled
If mProp_booHasPicture = True Then
    If intApplyAll = NOT_APPLY_ALL Then UserControl_Resize
End If
End Sub

Public Sub ApplyBorder(Optional ByVal intFirstApply As Integer = 0)
Dim rctBtn As RECT
Line (0, 0)-(Width, Height), mProp_lngBackColor, B
rctBtn.Left = 0
rctBtn.Top = 0
rctBtn.Right = ScaleWidth
rctBtn.Bottom = ScaleHeight
If mProp_booHasBorder = True Then
    DrawEdge hdc, rctBtn, BDR_MOUSEOVER_HB, BF_RECT
    If intFirstApply = FORCE_FLATTEN Or Not Ambient.UserMode Then
        'Remove thick edge
        Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), mProp_lngBackColor, B
    End If
Else
    DrawEdge hdc, rctBtn, BDR_MOUSEOVER, BF_RECT
End If
End Sub



