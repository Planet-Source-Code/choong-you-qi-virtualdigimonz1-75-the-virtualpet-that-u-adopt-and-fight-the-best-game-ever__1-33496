VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNeoCaption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' APIs
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Const CLR_INVALID = -1

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_CALCRECT = &H400
Private Const DT_WORDBREAK = &H10
Private Const DT_VCENTER = &H4
Private Const DT_TOP = &H0
Private Const DT_TABSTOP = &H80
Private Const DT_SINGLELINE = &H20
Private Const DT_RIGHT = &H2
Private Const DT_NOCLIP = &H100
Private Const DT_INTERNAL = &H1000
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_EXPANDTABS = &H40
Private Const DT_CHARSTREAM = 4
Private Const DT_NOPREFIX = &H800
Private Const DT_EDITCONTROL = &H2000&
Private Const DT_PATH_ELLIPSIS = &H4000&
Private Const DT_END_ELLIPSIS = &H8000&
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_RTLREADING = &H20000
Private Const DT_WORD_ELLIPSIS = &H40000

' Font:
Private Const LF_FACESIZE = 32
Private Type LOGFONT
   lfHeight As Long
   lfWidth As Long
   lfEscapement As Long
   lfOrientation As Long
   lfWeight As Long
   lfItalic As Byte
   lfUnderline As Byte
   lfStrikeOut As Byte
   lfCharSet As Byte
   lfOutPrecision As Byte
   lfClipPrecision As Byte
   lfQuality As Byte
   lfPitchAndFamily As Byte
   lfFaceName(LF_FACESIZE) As Byte
End Type
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const FF_DONTCARE = 0
Private Const DEFAULT_QUALITY = 0
Private Const DEFAULT_PITCH = 0
Private Const DEFAULT_CHARSET = 1
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const LOGPIXELSY = 90

Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112


' Implementation
Implements INCAreaModifier

Private Enum ECNCButtonStates
   up
   Down
End Enum

Private m_cNCS As cNCCalcSize
Private m_hWnd As Long

' MemDCs for storing GFX
Private m_cBorder As cMemDC
Private m_cCaption As cMemDC

' MemDC for building caption:
Private m_cFF As cMemDC
' and l/r borders
Private m_cFFB As cMemDC
' Menu bar:
Private m_cMenu As cMenuBar

Private m_oActiveCaptionColor As OLE_COLOR
Private m_oInActiveCaptionColor As OLE_COLOR
Private m_fnt As IFont

Private m_oActiveMenuColor As OLE_COLOR
Private m_oActiveMenuColorOver As OLE_COLOR
Private m_oInActiveMenuColor As OLE_COLOR
Private m_oMenuBackgroundColor As OLE_COLOR
Private m_fntMenu As IFont

Private m_lButtonWidth As Long
Private m_lButtonHeight As Long
Private m_lActiveLeftEnd As Long
Private m_lActiveRightStart As Long
Private m_lActiveRightEnd As Long
Private m_lInactiveOffset As Long

Private m_tBtn(0 To 2) As RECT
Private m_bMaximise As Boolean
Private m_bMinimise As Boolean
Private m_bClose As Boolean
Private m_bMouseDownMinimise As Boolean
Private m_bMouseDownMaximise As Boolean
Private m_bMouseDownClose As Boolean



Public Sub Detach()
Dim lMenu As Long
   If Not m_cNCS Is Nothing Then
      m_cNCS.Detach
   End If
   If Not m_cMenu Is Nothing Then
      lMenu = m_cMenu.hMenu
      m_cMenu.Detach
   End If
   If Not (lMenu = 0) Then
      SetMenu m_hWnd, lMenu
   End If
   
End Sub

Public Sub Attach( _
      f As Object, _
      picCaption As StdPicture, _
      picBorder As StdPicture, _
      lButtonWidth As Long, _
      lButtonHeight As Long, _
      lActiveLeftEnd As Long, _
      lActiveRightStart As Long, _
      lActiveRightEnd As Long, _
      lInactiveOffset As Long _
   )
   
   Detach
   
   ' Store the pictures:
   Set m_cCaption = New cMemDC
   m_cCaption.CreateFromPicture picCaption
   Set m_cBorder = New cMemDC
   m_cBorder.CreateFromPicture picBorder
   
   ' FF drawing
   Set m_cFF = New cMemDC
   Set m_cFFB = New cMemDC
   
   ' Store passed in vars:
   m_lButtonWidth = lButtonWidth
   m_lButtonHeight = lButtonHeight

   m_lActiveLeftEnd = lActiveLeftEnd
   m_lActiveRightStart = lActiveRightStart
   m_lActiveRightEnd = lActiveRightEnd
   m_lInactiveOffset = lInactiveOffset
         
   ' Store hWNd:
   m_hWnd = f.hwnd
         
   ' Menu:
   Set m_cMenu = New cMenuBar
   m_cMenu.Attach m_hWnd
   m_cMenu.Font = m_fntMenu
   m_cMenu.SetColors m_oActiveMenuColor, m_oActiveMenuColorOver, m_oInActiveMenuColor, m_oMenuBackgroundColor
   m_cMenu.CaptionHeight = m_cCaption.Height

         
   ' Start non-client modification:
   Set m_cNCS = New cNCCalcSize
   m_cNCS.Attach Me
   m_cNCS.Display f
   
   If IsWindowVisible(m_hWnd) <> 0 Then
      SetForegroundWindow m_hWnd
      SetFocusAPI m_hWnd
      SendMessageLong m_hWnd, WM_NCACTIVATE, 1, 0
   End If
   
End Sub
Public Property Get MenuBackgroundColor() As OLE_COLOR
   MenuBackgroundColor = m_oMenuBackgroundColor
End Property
Public Property Let MenuBackgroundColor(ByVal oColor As OLE_COLOR)
   m_oMenuBackgroundColor = oColor
End Property
Public Property Get ActiveCaptionColor() As OLE_COLOR
   ActiveCaptionColor = m_oActiveCaptionColor
End Property
Public Property Let ActiveCaptionColor(ByVal oColor As OLE_COLOR)
   m_oActiveCaptionColor = oColor
End Property
Public Property Get InActiveCaptionColor() As OLE_COLOR
   InActiveCaptionColor = m_oInActiveCaptionColor
End Property
Public Property Let InActiveCaptionColor(ByVal oColor As OLE_COLOR)
   m_oInActiveCaptionColor = oColor
End Property
Public Property Get CaptionFont() As IFont
   Set CaptionFont = m_fnt
End Property
Public Property Let CaptionFont(iFnt As IFont)
   Set m_fnt = iFnt
End Property
Public Property Get MenuFont() As IFont
   Set MenuFont = m_fntMenu
End Property
Public Property Let MenuFont(iFnt As IFont)
   Set m_fntMenu = iFnt
End Property
Public Property Get ActiveMenuColor() As OLE_COLOR
   ActiveMenuColor = m_oActiveMenuColor
End Property
Public Property Get ActiveMenuColorOver() As OLE_COLOR
   ActiveMenuColorOver = m_oActiveMenuColorOver
End Property
Public Property Get InActiveMenuColor() As OLE_COLOR
   InActiveMenuColor = m_oInActiveMenuColor
End Property
Public Property Let ActiveMenuColor(oColor As OLE_COLOR)
   m_oActiveMenuColor = oColor
End Property
Public Property Let ActiveMenuColorOver(oColor As OLE_COLOR)
   m_oActiveMenuColorOver = oColor
End Property
Public Property Let InActiveMenuColor(oColor As OLE_COLOR)
   m_oInActiveMenuColor = oColor
End Property
Private Sub Class_Initialize()
   m_oActiveCaptionColor = &HCCCCCC
   m_oInActiveCaptionColor = &H999999
   m_oActiveMenuColor = &H0&
   m_oActiveMenuColorOver = &H0&
   m_oInActiveMenuColor = &H808080
   m_oMenuBackgroundColor = &HFFFFFF
   Set m_fnt = New StdFont
   m_fnt.Name = "MS Sans Serif"
   Set m_fntMenu = New StdFont
   m_fntMenu.Name = "MS Sans Serif"
End Sub

Private Sub Class_Terminate()
   '
End Sub

Private Function INCAreaModifier_AltKeyAccelerator(ByVal vKey As KeyCodeConstants) As Long
    INCAreaModifier_AltKeyAccelerator = m_cMenu.AltKeyAccelerator(vKey)
End Function

Private Sub INCAreaModifier_ExitMenuLoop()
   m_cMenu.pRestoreList
End Sub

Private Sub INCAreaModifier_HitTest(ByVal x As Long, ByVal y As Long, eHitTest As VirtualDigimonz.ECNCHitTestConstants)
Dim bMouseOverClose As Boolean
Dim bMouseOverMaximise As Boolean
Dim bMouseOverMinimise As Boolean
Dim bBtnMouseDown As Boolean
Dim hDC As Long

   '
   Dim tR As RECT
   tR.left = 12: tR.top = 11: tR.right = 42: tR.bottom = 43
   If PtInRect(tR, x, y) <> 0 Then
      eHitTest = HTSYSMENU
      Exit Sub
   End If

   ' Code for working out whether in the buttons or not:
   If m_bClose Then
      If PtInRect(m_tBtn(0), x, y) <> 0 Then
         eHitTest = HTSYSMENU
         bMouseOverClose = True
      Else
         bMouseOverClose = False
      End If
   End If
   If m_bMaximise Then
      If PtInRect(m_tBtn(1), x, y) <> 0 Then
         eHitTest = HTSYSMENU
         bMouseOverMaximise = True
      Else
         bMouseOverMaximise = False
      End If
   End If
   If m_bMinimise Then
      If PtInRect(m_tBtn(2), x, y) <> 0 Then
         eHitTest = HTSYSMENU
         bMouseOverMinimise = True
      Else
         bMouseOverMinimise = False
      End If
   End If
   
   hDC = GetWindowDC(m_hWnd)
   
   bBtnMouseDown = GetAsyncKeyState(vbLeftButton)
   If m_bClose Then
      If Not (m_bMouseDownClose = bMouseOverClose) Then
         If bMouseOverClose And bBtnMouseDown And m_bMouseDownClose Then
            DrawButton hDC, 0, Down
         Else
            DrawButton hDC, 0, up
         End If
      End If
   End If
   If m_bMaximise Then
      If Not (m_bMouseDownMaximise = bMouseOverMaximise) Then
         If bMouseOverMaximise And bBtnMouseDown And m_bMouseDownMaximise Then
            DrawButton hDC, 1, Down
         Else
            DrawButton hDC, 1, up
         End If
      End If
   End If
   If m_bMinimise Then
      If Not (m_bMouseDownMinimise = bMouseOverMinimise) Then
         If bMouseOverMinimise And bBtnMouseDown And m_bMouseDownMinimise Then
            DrawButton hDC, 2, Down
         Else
            DrawButton hDC, 2, up
         End If
      End If
   End If
   ReleaseDC m_hWnd, hDC
   
End Sub

Private Property Get INCAreaModifier_hWnd() As Long
   INCAreaModifier_hWnd = m_hWnd
End Property


Private Sub INCAreaModifier_InitMenuPopup(ByVal wParam As Long, ByVal lParam As Long)
   ' Set all the menu items to Owner-Draw:
   ' wParam = hMenu
   m_cMenu.OwnerDrawMenu wParam
End Sub

Private Sub INCAreaModifier_NCMouseDown(ByVal x As Long, ByVal y As Long, bHandled As Boolean, ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
   If m_bClose Then
      If PtInRect(m_tBtn(0), x, y) <> 0 Then
         ' Redraw close button pressed:
         DrawButton hDC, 0, Down
         m_bMouseDownClose = True
         bHandled = True
      End If
   End If
   If m_bMaximise Then
      If PtInRect(m_tBtn(1), x, y) <> 0 Then
         ' Redraw maximise button pressed:
         DrawButton hDC, 1, Down
         m_bMouseDownMaximise = True
         bHandled = True
      End If
   End If
   If m_bMinimise Then
      If PtInRect(m_tBtn(2), x, y) <> 0 Then
         ' Redraw minimise button pressed:
         DrawButton hDC, 2, Down
         m_bMouseDownMinimise = True
         bHandled = True
      End If
   End If

End Sub

Private Sub INCAreaModifier_NCMouseUp(ByVal x As Long, ByVal y As Long, ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Dim lStyle As Long
   If m_bClose Then
      If PtInRect(m_tBtn(0), x, y) <> 0 Then
         If m_bMouseDownClose Then
            m_cNCS.SysCommand SC_CLOSE
         End If
      End If
   End If
   If m_bMaximise Then
      If PtInRect(m_tBtn(1), x, y) <> 0 Then
         If m_bMouseDownMaximise Then
            ' Redraw maximise button pressed:
            lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
            If ((lStyle And WS_MAXIMIZE) = WS_MAXIMIZE) Then
               m_cNCS.SysCommand SC_RESTORE
            Else
               m_cNCS.SysCommand SC_MAXIMIZE
            End If
         End If
      End If
   End If
   If m_bMinimise Then
      If PtInRect(m_tBtn(2), x, y) <> 0 Then
         If m_bMouseDownMinimise Then
            m_cNCS.SysCommand SC_MINIMIZE
         End If
      End If
   End If
   DrawButton hDC, 0, up
   DrawButton hDC, 1, up
   DrawButton hDC, 2, up
   
   m_bMouseDownMinimise = False
   m_bMouseDownMaximise = False
   m_bMouseDownClose = False
   
End Sub
Private Sub DrawButton(ByVal hDC As Long, ByVal iIndex As Long, ByVal eState As ECNCButtonStates)
Dim lY As Long
Dim lStyle As Long
   If eState = Down Then
      lY = m_lButtonHeight
   Else
      lY = 0
   End If
   Select Case iIndex
   Case 0
      If m_bClose Then
         BitBlt hDC, m_tBtn(0).left, m_tBtn(0).top, m_lButtonWidth, m_lButtonHeight, m_cCaption.hDC, 241, lY, vbSrcCopy
      End If
   Case 1
      If m_bMaximise Then
         lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
         If ((lStyle And WS_MAXIMIZE) = WS_MAXIMIZE) Then
            BitBlt hDC, m_tBtn(1).left, m_tBtn(1).top, m_lButtonWidth, m_lButtonHeight, m_cCaption.hDC, 240 + m_lButtonWidth, lY, vbSrcCopy
         Else
            BitBlt hDC, m_tBtn(1).left, m_tBtn(1).top, m_lButtonWidth, m_lButtonHeight, m_cCaption.hDC, 240 + m_lButtonWidth * 2, lY, vbSrcCopy
         End If
      End If
   Case 2
      If m_bMinimise Then
         BitBlt hDC, m_tBtn(2).left, m_tBtn(2).top, m_lButtonWidth, m_lButtonHeight, m_cCaption.hDC, 240 + m_lButtonWidth * 3, lY, vbSrcCopy
      End If
   End Select
End Sub

Private Sub INCAreaModifier_NCPaint(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Dim lX As Long, lXE As Long
Dim lY As Long
Dim lW As Long, lH As Long, lRW As Long
Dim lT As Long
Dim lSrcDC As Long, lSrcX As Long, lSrcY As Long
Dim lOrgX As Long
Dim bNoMiddle As Boolean
Dim tR As RECT
Dim sCaption As String
Dim lLen As Long
Dim tLF As LOGFONT
Dim hFnt As Long
Dim hFntOld As Long
Dim lStyle As Long
Dim lhDC As Long, lhDCB As Long
Dim hFntMenu As Long

   '
   ' Here we do the work!
   tR.left = lLeft
   tR.top = lTop
   tR.right = lRight
   tR.bottom = lBottom
   
   ' Ensure mem DCs are big enough to draw into:
   m_cFF.Width = tR.right - tR.left + 1
   m_cFF.Height = m_cCaption.Height
   lhDC = m_cFF.hDC
   
   m_cFFB.Width = m_cBorder.Width * 2
   m_cFFB.Height = tR.bottom - tR.top + 1
   lhDCB = m_cFFB.hDC
      
         
   pOLEFontToLogFont m_fnt, hDC, tLF
   If m_cNCS.WindowActive Then
      tLF.lfWeight = FW_BOLD
   End If
   hFnt = CreateFontIndirect(tLF)
   hFntOld = SelectObject(lhDC, hFnt)
  
   If m_cNCS.WindowActive Then
      lOrgX = 0
   Else
      lOrgX = m_lInactiveOffset
   End If
   ' Draw the caption
   BitBlt lhDC, lLeft, lTop, lLeft + m_lActiveLeftEnd, m_cCaption.Height, m_cCaption.hDC, lOrgX, 0, vbSrcCopy
   lRW = (m_lActiveRightEnd - m_lActiveRightStart + 1)
   lXE = lRight - lRW + 1
   If lXE < lLeft + lRW Then
      lXE = lLeft + lRW
      bNoMiddle = True
   End If
   BitBlt lhDC, lXE, lTop, lRW, m_cCaption.Height, m_cCaption.hDC, lOrgX + m_lActiveRightStart, 0, vbSrcCopy
   
   ' Buttons:
   lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
   m_bMaximise = ((lStyle And WS_MAXIMIZEBOX) = WS_MAXIMIZEBOX)
   m_bMinimise = ((lStyle And WS_MINIMIZEBOX) = WS_MINIMIZEBOX)
   m_bClose = ((lStyle And WS_SYSMENU) = WS_SYSMENU)
   m_tBtn(0).left = lXE + lRW - m_cBorder.Height + 4
   If m_bClose Then
      m_tBtn(0).left = m_tBtn(0).left - (m_lButtonWidth + 1)
      m_tBtn(0).top = lTop + 5
      m_tBtn(0).right = m_tBtn(0).left + m_lButtonWidth + 1
      m_tBtn(0).bottom = m_tBtn(0).top + m_lButtonHeight
      DrawButton lhDC, 0, up
   End If
   If m_bMaximise Then
      m_tBtn(1).left = m_tBtn(0).left - (m_lButtonWidth + 1)
      m_tBtn(1).top = lTop + 5
      m_tBtn(1).right = m_tBtn(1).left + m_lButtonWidth + 1
      m_tBtn(1).bottom = m_tBtn(1).top + m_lButtonHeight
      DrawButton lhDC, 1, up
   Else
      m_tBtn(1).left = m_tBtn(0).left
   End If
   If m_bMinimise Then
      m_tBtn(2).left = m_tBtn(1).left - (m_lButtonWidth + 1)
      m_tBtn(2).top = lTop + 5
      m_tBtn(2).right = m_tBtn(2).left + (m_lButtonWidth + 1)
      m_tBtn(2).bottom = m_tBtn(2).top + m_lButtonHeight
      DrawButton lhDC, 2, up
   End If
            
   ' Fill in:
   lX = lLeft + 90
   Do
      lW = 52
      If lX + 52 > lXE Then
         lW = lXE - lX
      End If
      BitBlt lhDC, lX, 0, lW, m_cCaption.Height, m_cCaption.hDC, lOrgX + m_lActiveLeftEnd + 1, 0, vbSrcCopy
      lX = lX + 52
   Loop While lX < lXE
      
   If Not bNoMiddle Then
      
      ' Draw the caption:
      SetBkMode lhDC, TRANSPARENT
      If m_cNCS.WindowActive Then
         SetTextColor lhDC, TranslateColor(m_oActiveCaptionColor)
      Else
         SetTextColor lhDC, TranslateColor(m_oInActiveCaptionColor)
      End If
      lLen = GetWindowTextLength(m_hWnd)
      If lLen > 0 Then
         tR.left = lLeft + 92
         tR.right = lRight - 96
         tR.top = m_cBorder.Height + 1
         tR.bottom = tR.top + (m_cCaption.Height - m_cBorder.Height - 2) \ 2
         sCaption = String$(lLen + 1, 0)
         GetWindowText m_hWnd, sCaption, lLen + 1
         DrawText lhDC, sCaption, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_END_ELLIPSIS Or DT_NOPREFIX
      End If
      
   End If
   
   ' Menu:
   m_cMenu.hMenu = m_cNCS.hMenu
   lW = lXE - m_lActiveLeftEnd
   tLF.lfWeight = FW_NORMAL
   hFntMenu = CreateFontIndirect(tLF)
   m_cMenu.Render hFntMenu, lhDC, m_lActiveLeftEnd, m_cCaption.Height \ 2, lW, m_cCaption.Height \ 2, -m_cCaption.Height \ 2 + 2
   DeleteObject hFntMenu
   
   BitBlt hDC, 0, 0, m_cFF.Width, m_cFF.Height, lhDC, 0, 0, vbSrcCopy
   
   
   ' Draw the border:
   lY = m_cCaption.Height
   lH = m_cBorder.Height
   lW = lH
   lSrcDC = m_cBorder.hDC
   lSrcX = lW * 4
   lSrcY = 0
   ' We draw double the amount each time for a quick finish:
   Do
      ' Draw to lhs:
      BitBlt lhDCB, 0, lY + lTop, lW, lH, lSrcDC, 0, lSrcY, vbSrcCopy
      ' Draw to right:
      BitBlt lhDCB, lW, lY + lTop, lW, lH, lSrcDC, lSrcX, lSrcY, vbSrcCopy
      'Exit Do
      If lSrcY = 0 Then
         lSrcDC = lhDCB
         lSrcY = lY + lTop
         lSrcX = lW
         lY = lY + lH
      Else
         lY = lY + lH
         lH = lH * 2
      End If
   Loop While lY < lBottom - lW
   lT = m_cCaption.Height + lTop
   lH = lBottom - lT
   BitBlt hDC, lLeft, lT, lW, lH, lhDCB, 0, lT, vbSrcCopy
   BitBlt hDC, lRight - lW, lT, lW, lH, lhDCB, lW, lT, vbSrcCopy
   
   lT = lBottom - lW
   If lT < m_cCaption.Height Then
      lT = m_cCaption.Height
   End If
   
   ' Bottom - we draw into the caption mem dc for flicker free
   lX = lLeft + lW
   lH = m_cBorder.Height
   lSrcDC = m_cBorder.hDC
   lSrcX = lW * 3
   lSrcY = 0
   ' We draw double the amount each time for a quick finish:
   Do
      BitBlt lhDC, lX, 0, lW, lH, lSrcDC, lSrcX, lSrcY, vbSrcCopy
      If lSrcY = 0 Then
         lSrcDC = lhDC
         lSrcX = lX
         lX = lX + lW
      Else
         lX = lX + lW
         lW = lW * 2
      End If
   Loop While lX < lRight - lH
   ' Bottom corners
   BitBlt lhDC, lLeft, 0, lH, lH, m_cBorder.hDC, lH * 2, 0, vbSrcCopy
   BitBlt lhDC, lRight - lH, 0, lH, lH, m_cBorder.hDC, lH * 6, 0, vbSrcCopy
   
   ' Swap out to display:
   BitBlt hDC, lLeft, lT, m_cFF.Width, lH, lhDC, 0, 0, vbSrcCopy
  
   SelectObject lhDC, hFntOld
   DeleteObject hFnt

End Sub

Private Sub INCAreaModifier_GetBottomMarginHeight(cy As Long)
  '
  cy = m_cBorder.Height
End Sub

Private Sub INCAreaModifier_GetLeftMarginWidth(cx As Long)
   '
   cx = m_cBorder.Height
End Sub

Private Sub INCAreaModifier_GetRightMarginWidth(cx As Long)
   '
   cx = m_cBorder.Height
End Sub

Private Sub INCAreaModifier_GetTopMarginHeight(cy As Long)
   '
   cy = m_cCaption.Height
End Sub

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


Private Sub pOLEFontToLogFont(fntThis As StdFont, ByVal hDC As Long, tLF As LOGFONT)
Dim sFont As String
Dim iChar As Integer
Dim b() As Byte

   ' Convert an OLE StdFont to a LOGFONT structure:
   With tLF
     sFont = fntThis.Name
     b = StrConv(sFont, vbFromUnicode)
     For iChar = 1 To Len(sFont)
       .lfFaceName(iChar - 1) = b(iChar - 1)
     Next iChar
     ' Based on the Win32SDK documentation:
     .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
     .lfItalic = fntThis.Italic
     If (fntThis.Bold) Then
       .lfWeight = FW_BOLD
     Else
       .lfWeight = FW_NORMAL
     End If
     .lfUnderline = fntThis.Underline
     .lfStrikeOut = fntThis.Strikethrough
     .lfCharSet = fntThis.Charset
   End With

End Sub




