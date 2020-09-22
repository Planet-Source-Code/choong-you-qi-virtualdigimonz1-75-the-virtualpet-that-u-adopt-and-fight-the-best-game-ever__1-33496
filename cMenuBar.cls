VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMenuBar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =======================================================================
' MENU private declares:
' =======================================================================

' Menu flag constants:
Private Const MF_APPEND = &H100&
Private Const MF_BITMAP = &H4&
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_CALLBACKS = &H8000000
Private Const MF_CHANGE = &H80&
Private Const MF_CHECKED = &H8&
Private Const MF_CONV = &H40000000
Private Const MF_DELETE = &H200&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_END = &H80
Private Const MF_ERRORS = &H10000000
Private Const MF_GRAYED = &H1&
Private Const MF_HELP = &H4000&
Private Const MF_HILITE = &H80&
Private Const MF_HSZ_INFO = &H1000000
Private Const MF_INSERT = &H0&
Private Const MF_LINKS = &H20000000
Private Const MF_MASK = &HFF000000
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_MOUSESELECT = &H8000&
Private Const MF_OWNERDRAW = &H100&
Private Const MF_POPUP = &H10&
Private Const MF_POSTMSGS = &H4000000
Private Const MF_REMOVE = &H1000&
Private Const MF_SENDMSGS = &H2000000
Private Const MF_SEPARATOR = &H800&
Private Const MF_STRING = &H0&
Private Const MF_SYSMENU = &H2000&
Private Const MF_UNCHECKED = &H0&
Private Const MF_UNHILITE = &H0&
Private Const MF_USECHECKBITMAPS = &H200&
Private Const MF_DEFAULT = &H1000&

Private Const MFT_STRING = MF_STRING
Private Const MFT_BITMAP = MF_BITMAP
Private Const MFT_MENUBARBREAK = MF_MENUBARBREAK
Private Const MFT_MENUBREAK = MF_MENUBREAK
Private Const MFT_OWNERDRAW = MF_OWNERDRAW
Private Const MFT_RADIOCHECK = &H200&
Private Const MFT_SEPARATOR = MF_SEPARATOR
Private Const MFT_RIGHTORDER = &H2000&

' New versions of the names...
Private Const MFS_GRAYED = &H3&
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MFS_CHECKED = MF_CHECKED
Private Const MFS_HILITE = MF_HILITE
Private Const MFS_ENABLED = MF_ENABLED
Private Const MFS_UNCHECKED = MF_UNCHECKED
Private Const MFS_UNHILITE = MF_UNHILITE
Private Const MFS_DEFAULT = MF_DEFAULT

' MenuItemInfo Mask constants
Private Const MIIM_STATE = &H1&
Private Const MIIM_ID = &H2&
Private Const MIIM_SUBMENU = &H4&
Private Const MIIM_CHECKMARKS = &H8&
Private Const MIIM_TYPE = &H10&
Private Const MIIM_DATA = &H20&

Private Const SC_RESTORE = &HF120&
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Const SC_MAXIMIZE = &HF030&
Private Const SC_MINIMIZE = &HF020&
Private Const SC_CLOSE = &HF060&
     
Private Const SC_ARRANGE = &HF110&
Private Const SC_HOTKEY = &HF150&
Private Const SC_HSCROLL = &HF080&
Private Const SC_KEYMENU = &HF100&
Private Const SC_MOUSEMENU = &HF090&
Private Const SC_NEXTWINDOW = &HF040&
Private Const SC_PREVWINDOW = &HF050&
Private Const SC_SCREENSAVE = &HF140&
Private Const SC_TASKLIST = &HF130&
Private Const SC_VSCROLL = &HF070&
Private Const SC_ZOOM = SC_MAXIMIZE
Private Const SC_ICON = SC_MINIMIZE

' Owner draw information:
Private Const ODS_CHECKED = &H8
Private Const ODS_DISABLED = &H4
Private Const ODS_FOCUS = &H10
Private Const ODS_GRAYED = &H2
Private Const ODS_SELECTED = &H1
Private Const ODT_BUTTON = 4
Private Const ODT_COMBOBOX = 3
Private Const ODT_LISTBOX = 2
Private Const ODT_MENU = 1

Private Type MEASUREITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemWidth As Long
   itemHeight As Long
   ItemData As Long
End Type

Private Type DRAWITEMSTRUCT
   CtlType As Long
   CtlID As Long
   itemID As Long
   itemAction As Long
   itemState As Long
   hwndItem As Long
   hdc As Long
   rcItem As RECT
   ItemData As Long
End Type

Private Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As Long
   cch As Long
End Type
Private Type MENUITEMINFO_STRINGDATA
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type

Private Type MENUITEMTEMPLATE
   mtOption As Integer
   mtID As Integer
   mtString As Byte
End Type
Private Type MENUITEMTEMPLATEHEADER
   versionNumber As Integer
   Offset As Integer
End Type

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long

Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Private Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuItemInfoStr Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Boolean, lpMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfoStr Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO_STRINGDATA) As Long
Private Declare Function GetMenuItemRect Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Private Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateMenu Lib "user32" () As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function CreatePopupMenu Lib "user32" () As Long

Private Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function AppendMenuByString Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function InsertMenuByString Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As String) As Long
Private Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function CheckMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function EnableMenuItem Lib "user32" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function HiliteMenuItem Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long

Private Declare Function MenuItemFromPoint Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal ptScreen As POINTAPI) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

' =======================================================================
' GDI private declares:
' =======================================================================

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

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

Private Const OPAQUE = 2
Private Const TRANSPARENT = 1

' DrawEdge:
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_SUNKENINNER = &H8

Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC
Private Const BDR_RAISED = &H5
Private Const BDR_SUNKEN = &HA

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)

Private Const BF_LEFT = &H1
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_BOTTOM = &H8

Private Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Private Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Private Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Private Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const CLR_INVALID = -1


' =======================================================================
' General Win private declares:
' =======================================================================
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const HWND_DESKTOP = 0


' =======================================================================
' IMPLEMENTATION
' =======================================================================

Private m_cMemDC As cMemDC
Private m_cToolbarMenu As cToolbarMenu
Private m_hMenu As Long
Private m_hWnd As Long

Private m_tR() As RECT
Private m_hSubMenu() As Long
Private m_iCount As Long
Private m_iDownOn As Long
Private m_iOver As Long

Private m_oActiveMenuColor As OLE_COLOR
Private m_oActiveMenuColorOver As OLE_COLOR
Private m_oInActiveMenuColor As OLE_COLOR

Private m_oMenuBackgroundColor As OLE_COLOR

Private m_lCaptionHeight As Long

Private m_iRestore As Long
Private m_hMenuRestore() As Long
Private m_iMenuPosition() As Long
Private m_tMIIS() As MENUITEMINFO_STRINGDATA
Private m_sCaption() As String
Private m_sShortCut() As String
Private m_sAccelerator() As String
Private m_lMenuTextSize() As Long
Private m_lMenuShortCutSize() As Long

Private m_iHaveSeenCount As Long
Private m_hMenuSeen() As Long

Private m_fnt As StdFont
Private m_fntSymbol As StdFont

Private m_lMenuItemHeight As Long

Private WithEvents m_cTmr As CTimer
Attribute m_cTmr.VB_VarHelpID = -1

Implements ISubclass

Friend Property Let Font( _
      fntThis As StdFont _
   )
   Set m_fnt = fntThis
End Property
Friend Property Set Font( _
      fntThis As StdFont _
   )
   Set m_fnt = fntThis
   m_fntSymbol.Name = "Marlett"
   m_fntSymbol.Size = m_fnt.Size * 1.2
End Property
Friend Property Get Font() As StdFont
   Set Font = m_fnt
End Property
Friend Sub SetColors( _
      ByVal oActiveMenuColor As OLE_COLOR, _
      ByVal oActiveMenuColorOver As OLE_COLOR, _
      ByVal oInActiveMenuColor As OLE_COLOR, _
      ByVal oMenuBackgroundColor As OLE_COLOR _
   )
   m_oActiveMenuColor = oActiveMenuColor
   m_oActiveMenuColorOver = oActiveMenuColorOver
   m_oInActiveMenuColor = oInActiveMenuColor
   m_oMenuBackgroundColor = oMenuBackgroundColor
End Sub
Private Property Get hFont() As Long
Dim iFn As IFont
   Set iFn = m_fnt
   hFont = iFn.hFont
End Property
Private Property Get hFontSymbol() As Long
Dim iFn As IFont
   Set iFn = m_fntSymbol
   hFontSymbol = iFn.hFont
End Property

Public Property Let hMenu(ByVal hTheMenu As Long)
   m_hMenu = hTheMenu
End Property
Public Property Get hMenu() As Long
   hMenu = m_hMenu
End Property
Public Sub Attach(ByVal lhWnd As Long)
   Detach
   m_hWnd = lhWnd
   Set m_cToolbarMenu = New cToolbarMenu
   m_cToolbarMenu.CoolMenuAttach m_hWnd, Me
   AttachMessage Me, m_hWnd, WM_LBUTTONDOWN
   AttachMessage Me, m_hWnd, WM_MOUSEMOVE
   AttachMessage Me, m_hWnd, WM_DRAWITEM
   AttachMessage Me, m_hWnd, WM_MEASUREITEM
   AttachMessage Me, m_hWnd, WM_MENUCHAR
End Sub
Public Sub Detach()
   If Not m_hWnd = 0 Then
      DetachMessage Me, m_hWnd, WM_LBUTTONDOWN
      DetachMessage Me, m_hWnd, WM_MOUSEMOVE
      DetachMessage Me, m_hWnd, WM_DRAWITEM
      DetachMessage Me, m_hWnd, WM_MEASUREITEM
      DetachMessage Me, m_hWnd, WM_MENUCHAR
   End If
   If Not m_cToolbarMenu Is Nothing Then
      m_cToolbarMenu.CoolMenuDetach
      Set m_cToolbarMenu = Nothing
   End If
End Sub
Public Property Let CaptionHeight(ByVal lHeight As Long)
   m_lCaptionHeight = lHeight
End Property

Public Sub Render( _
      ByVal hFnt As Long, _
      ByVal lhDC As Long, _
      ByVal lLeft As Long, _
      ByVal lTop As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByVal lYoffset As Long _
   )
Dim iIdx As Long
Dim lC As Long
Dim lhDCC As Long
Dim tMII As MENUITEMINFO_STRINGDATA
Dim sCap As String
Dim hFntOld As Long
Dim tTR As RECT, tBR As RECT
Dim lX As Long
Dim lR As Long
Dim bPress As Boolean
Dim lID As Long

   If Not (m_hMenu = 0) Then
      m_cMemDC.Width = lWidth
      m_cMemDC.Height = lHeight
      lhDCC = m_cMemDC.hdc

      hFntOld = SelectObject(lhDCC, hFnt)
      m_iCount = 0
      Erase m_tR

      lC = GetMenuItemCount(m_hMenu)
      If lC > 0 Then
         lX = 8
         lTop = lTop + 2
         BitBlt lhDCC, 0, 0, lWidth, lHeight, lhDC, lLeft, lTop, vbSrcCopy
         SetBkMode lhDCC, TRANSPARENT
         For iIdx = 0 To lC - 1
            lID = GetMenuItemID(m_hMenu, iIdx)
            If lID = -1 Then
               tMII.fMask = MIIM_TYPE
               tMII.cch = 127
               tMII.dwTypeData = String$(128, 0)
               tMII.cbSize = LenB(tMII)
               lR = GetMenuItemInfoStr(m_hMenu, iIdx, True, tMII)
               If (tMII.fType And MFT_STRING) = MFT_STRING Then
                  If tMII.cch > 0 Then
                     sCap = left$(tMII.dwTypeData, tMII.cch)
                  Else
                     sCap = ""
                  End If
                  tTR.top = 0
                  tTR.bottom = lHeight
                  tTR.left = 0: tTR.right = 0
                  DrawText lhDCC, sCap, -1, tTR, DT_CALCRECT
                  OffsetRect tTR, lX, 2
                  LSet tBR = tTR
                  InflateRect tBR, 2, 2
                  tBR.right = tBR.right + 7
                  m_iCount = m_iCount + 1
                  bPress = False
                  If m_iCount = m_iDownOn Then
                     ' This is the item that was clicked:
                     If m_iDownOn = m_iOver Then
                        ' Draw Pressed
                        'Debug.Print "DrawPressed"
                        bPress = True
                        SetTextColor lhDCC, TranslateColor(m_oActiveMenuColorOver)
                        DrawEdge lhDCC, tBR, BDR_SUNKENOUTER, BF_RECT
                     Else
                        ' Draw Raised
                        'Debug.Print "DrawRaised"
                        SetTextColor lhDCC, TranslateColor(m_oActiveMenuColorOver)
                        DrawEdge lhDCC, tBR, BDR_RAISEDINNER, BF_RECT
                     End If
                  Else
                     ' Not down on, may be over:
                     If m_iCount = m_iOver Then
                        ' Draw Raised
                        'Debug.Print "DrawRaised"
                        SetTextColor lhDCC, TranslateColor(m_oActiveMenuColorOver)
                        DrawEdge lhDCC, tBR, BDR_RAISEDINNER, BF_RECT
                     Else
                        ' Draw None
                        SetTextColor lhDCC, TranslateColor(m_oActiveMenuColor)
                     End If
                  End If
                  If bPress Then
                     OffsetRect tTR, 1, 1
                  End If
                  DrawText lhDCC, sCap, -1, tTR, DT_LEFT Or DT_SINGLELINE
                  If bPress Then
                     OffsetRect tTR, -1, -1
                  End If
                  ReDim Preserve m_tR(1 To m_iCount) As RECT
                  ReDim Preserve m_hSubMenu(1 To m_iCount) As Long
                  OffsetRect tBR, lLeft, lYoffset
                  LSet m_tR(m_iCount) = tBR
                  m_hSubMenu(m_iCount) = GetSubMenu(m_hMenu, iIdx)
                  lX = lX + tTR.right - tTR.left + 1 + 10
               End If
            End If
         Next iIdx

         BitBlt lhDC, lLeft, lTop, lWidth, lHeight, lhDCC, 0, 0, vbSrcCopy

      End If
   
      SelectObject lhDCC, hFntOld
   End If
End Sub
Friend Function AltKeyAccelerator(ByVal vKey As KeyCodeConstants) As Boolean
Dim lC As Long
Dim iIdx As Long
Dim tMII As MENUITEMINFO_STRINGDATA
Dim lR As Long
Dim sCap As String
Dim iPos As Long
Dim sAccel As String

   lC = GetMenuItemCount(m_hMenu)
   If lC > 0 Then
      For iIdx = 0 To lC - 1
         tMII.fMask = MIIM_TYPE Or MIIM_DATA
         tMII.cch = 127
         tMII.dwTypeData = String$(128, 0)
         tMII.cbSize = LenB(tMII)
         lR = GetMenuItemInfoStr(m_hMenu, iIdx, True, tMII)
         If tMII.cch > 0 Then
            sCap = left$(tMII.dwTypeData, tMII.cch)
            iPos = InStr(sCap, "&")
            If iPos > 0 And iPos < Len(sCap) Then
               sAccel = UCase$(Mid$(sCap, iPos + 1, 1))
               If sAccel = Chr$(vKey) Then
                  PressButton iIdx + 1, True
                  If Not m_cTmr Is Nothing Then
                     m_cTmr.Interval = 0
                  End If
                  lR = m_cToolbarMenu.TrackPopup(m_iDownOn)
                  pRestoreList
                  AltKeyAccelerator = True
               End If
            End If
         End If
      Next iIdx
   End If
End Function
Private Function MenuHitTest() As Long

   If m_iCount > 0 Then
      Dim tP As POINTAPI
      GetCursorPos tP
      MenuHitTest = HitTest(tP)
   End If
   
End Function
Friend Function HitTest(tP As POINTAPI) As Long

   ' Is tP within a top level menu button? tP
   ' is in screen coords
   '
Dim iMenu As Long

   ScreenToClient m_hWnd, tP
   For iMenu = 1 To m_iCount
      'Debug.Print m_tR(iMenu).left, m_tR(iMenu).top, m_tR(iMenu).right, m_tR(iMenu).bottom, tP.x, tP.y
      If PtInRect(m_tR(iMenu), tP.x, tP.y) <> 0 Then
         HitTest = iMenu
         Exit Function
      End If
   Next iMenu
End Function
Friend Property Get Count() As Long
   
   ' Number of top level menu items:?
   '
   Count = m_iCount
   
End Property
Friend Function GetMenuHandle(ByVal iNewPopup As Long) As Long
   
   ' Returns the popup menu handle for a given top level
   ' menu item (1 based index)
   '
   If iNewPopup > 0 And iNewPopup <= m_iCount Then
      GetMenuHandle = m_hSubMenu(iNewPopup)
   End If
End Function
Friend Sub PressButton(ByVal iButton As Long, ByVal bState As Boolean)
   '
   If bState Then
      m_iDownOn = iButton
   Else
      If m_iDownOn = iButton Then
         m_iDownOn = -1
      End If
   End If
   SendMessageLong m_hWnd, WM_NCPAINT, 0, 0
   
End Sub
Friend Sub GetRect(ByVal iButton As Long, ByRef tR As RECT)
Dim tRW As RECT
   If iButton > 0 And iButton <= m_iCount Then
      LSet tR = m_tR(iButton)
      GetWindowRect m_hWnd, tRW
      OffsetRect tR, tRW.left, tRW.top + m_lCaptionHeight
   End If
End Sub
Friend Property Get HotItem() As Long
   '
   HotItem = m_iDownOn
End Property
Friend Property Let HotItem(ByVal iHotItem As Long)
   ' Set the hotitem
   m_iOver = iHotItem
   ' Repaint:
   SendMessageLong m_hWnd, WM_NCPAINT, 0, 0
End Property

Friend Sub OwnerDrawMenu(ByVal hMenu As Long)
Dim lC As Long
Dim tMIIS As MENUITEMINFO_STRINGDATA
Dim tMII As MENUITEMINFO
Dim iMenu As Long
Dim sCap As String
Dim sShortCut As String
Dim tR As RECT
Dim iPos As Long
Dim lID As Long
Dim bHaveSeen As Boolean
Dim hFntOld As Long
Dim lMenuTextSize As Long
Dim lMenuShortCutSize As Long
Dim i As Long
                  
   ' Set OD flag on the fly...
   bHaveSeen = pbHaveSeen(hMenu)

   hFntOld = SelectObject(m_cMemDC.hdc, hFont)
   lC = GetMenuItemCount(hMenu)
   For iMenu = 0 To lC - 1
      
      If Not bHaveSeen Then
               
         tMIIS.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
         tMIIS.cch = 127
         tMIIS.dwTypeData = String$(128, 0)
         tMIIS.cbSize = LenB(tMIIS)
         GetMenuItemInfoStr hMenu, iMenu, True, tMIIS
         'Debug.Print "New Item", tMIIS.dwTypeData
         
         lID = plAddToRestoreList(hMenu, iMenu, tMIIS)
      
         If Not (tMIIS.fType And MFT_OWNERDRAW) = MFT_OWNERDRAW Then
            ' Setting this flag causes tMIIS.dwTypeData to be
            ' overwritten with our own app-defined value:
            tMII.fType = tMIIS.fType Or MFT_OWNERDRAW
            tMII.dwItemData = lID
            tMII.cbSize = LenB(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_DATA
            SetMenuItemInfo hMenu, iMenu, True, tMII
         End If
      
      Else
         
         tMII.fMask = MIIM_TYPE Or MIIM_DATA
         tMII.cbSize = Len(tMII)
         GetMenuItemInfo hMenu, iMenu, True, tMII
         lID = tMII.dwItemData
         
         If Not ((tMII.fType And MFT_OWNERDRAW) = MFT_OWNERDRAW) Then
            
            lID = plReplaceIndex(hMenu, iMenu)
         
            'Debug.Print "VB has done something to it!", lID
            tMIIS.fMask = MIIM_TYPE Or MIIM_DATA Or MIIM_ID
            tMIIS.cch = 127
            tMIIS.dwTypeData = String$(128, 0)
            tMIIS.cbSize = LenB(tMIIS)
            GetMenuItemInfoStr hMenu, iMenu, True, tMIIS
            
            pReplaceRestoreList lID, hMenu, iMenu, tMIIS
            
            ' Setting this flag causes tMIIS.dwTypeData to be
            ' overwritten with our own app-defined value:
            tMII.fType = tMIIS.fType Or MFT_OWNERDRAW
            tMII.dwItemData = lID
            tMII.cbSize = LenB(tMII)
            tMII.fMask = MIIM_TYPE Or MIIM_DATA
            SetMenuItemInfo hMenu, iMenu, True, tMII
            
         End If
         
      End If
                              
      If lID > 0 And lID <= m_iRestore Then
         sCap = m_sCaption(lID)
         sShortCut = m_sShortCut(lID)
         
         'Debug.Print m_sCaption(lID), m_sShortCut(lID)
         
         DrawText m_cMemDC.hdc, sCap, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
         If tR.right - tR.left + 1 > lMenuTextSize Then
            lMenuTextSize = tR.right - tR.left + 1
         End If
         If Len(sShortCut) > 0 Then
            DrawText m_cMemDC.hdc, sShortCut, -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
            If tR.right - tR.left + 1 > lMenuShortCutSize Then
               lMenuShortCutSize = tR.right - tR.left + 1
            End If
         End If
         m_lMenuItemHeight = tR.bottom - tR.top + 1
         
      Else
         'Debug.Print "ERROR! ERROR! ERROR!"
      End If
      
   Next iMenu
   
   For i = 1 To m_iRestore
      If m_hMenuRestore(i) = hMenu Then
         m_lMenuTextSize(i) = lMenuTextSize
         m_lMenuShortCutSize(i) = lMenuShortCutSize
      End If
   Next i
   
   SelectObject m_cMemDC.hdc, hFntOld
   
End Sub
Private Function pbHaveSeen(ByVal hMenu As Long) As Boolean
   
   ' When WM_INITMENUPOPUP fires, this may or not be
   ' a new menu.  We use an array to store which menus
   ' we've already worked on:

Dim i As Long
   
   For i = 1 To m_iHaveSeenCount
      If hMenu = m_hMenuSeen(i) Then
         pbHaveSeen = True
         Exit Function
      End If
   Next i
   m_iHaveSeenCount = m_iHaveSeenCount + 1
   ReDim Preserve m_hMenuSeen(1 To m_iHaveSeenCount) As Long
   m_hMenuSeen(m_iHaveSeenCount) = hMenu

End Function
Private Function plReplaceIndex(ByVal hMenu As Long, ByVal iMenu As Long)
Dim i As Long
   For i = 1 To m_iRestore
      If m_hMenuRestore(i) = hMenu Then
         If m_iMenuPosition(i) = iMenu Then
            plReplaceIndex = i
            Exit Function
         End If
      End If
   Next i
End Function
Private Function plAddToRestoreList(ByVal hMenu As Long, ByVal iMenu As Long, tMIIS As MENUITEMINFO_STRINGDATA) As Long
   
   ' Here we store information about a menu item.  When the
   ' menus are closed again we can reset things back to the
   ' way they were using this struct.

   m_iRestore = m_iRestore + 1
   ReDim Preserve m_hMenuRestore(1 To m_iRestore) As Long
   ReDim Preserve m_iMenuPosition(1 To m_iRestore) As Long
   ReDim Preserve m_tMIIS(1 To m_iRestore) As MENUITEMINFO_STRINGDATA
   ReDim Preserve m_sCaption(1 To m_iRestore) As String
   ReDim Preserve m_sShortCut(1 To m_iRestore) As String
   ReDim Preserve m_sAccelerator(1 To m_iRestore) As String
   ReDim Preserve m_lMenuTextSize(1 To m_iRestore) As Long
   ReDim Preserve m_lMenuShortCutSize(1 To m_iRestore) As Long
   pReplaceRestoreList m_iRestore, hMenu, iMenu, tMIIS
   plAddToRestoreList = m_iRestore

End Function
Private Sub pReplaceRestoreList(ByVal lIdx As Long, hMenu As Long, iMenu As Long, tMIIS As MENUITEMINFO_STRINGDATA)
Dim sCap As String
Dim sShortCut As String
Dim iPos As Long

   m_hMenuRestore(lIdx) = hMenu
   m_iMenuPosition(lIdx) = iMenu
   LSet m_tMIIS(lIdx) = tMIIS
   If tMIIS.cch > 0 Then
      sCap = left$(tMIIS.dwTypeData, tMIIS.cch)
   Else
      sCap = ""
   End If
   iPos = InStr(sCap, vbTab)
   If iPos > 0 Then
      m_sShortCut(lIdx) = Mid$(sCap, iPos + 1)
      m_sCaption(lIdx) = left$(sCap, iPos - 1)
   Else
      m_sCaption(lIdx) = sCap
      m_sShortCut(lIdx) = ""
   End If
   iPos = InStr(m_sCaption(lIdx), "&")
   If iPos > 0 And iPos < Len(m_sCaption(lIdx)) Then
      m_sAccelerator(lIdx) = UCase$(Mid$(m_sCaption(lIdx), iPos + 1, 1))
   End If
End Sub
Private Function InternalIDForWindowsID(ByVal wID As Long) As Long
Dim i As Long
   ' linear search I'm afraid, but it is only called once
   ' per menu item shown (when WM_MEASUREITEM is fired)
   For i = 1 To m_iRestore
      If m_tMIIS(i).wID = wID Then
         InternalIDForWindowsID = i
         Exit Function
      End If
   Next i
End Function
Friend Sub pRestoreList()
Dim i As Long
   'Debug.Print "RESTORELIST"
   ' erase the lot:
   For i = 1 To m_iRestore
      SetMenuItemInfoStr m_hMenuRestore(i), m_iMenuPosition(i), True, m_tMIIS(i)
   Next i
   m_iRestore = 0
   Erase m_hMenuRestore
   Erase m_iMenuPosition
   Erase m_tMIIS
   Erase m_sCaption()
   Erase m_sShortCut()
   Erase m_sAccelerator()
   m_iHaveSeenCount = 0
   Erase m_hMenuSeen()
End Sub

Private Sub Class_Initialize()
   Set m_cMemDC = New cMemDC
   Set m_fnt = New StdFont
   m_fnt.Name = "MS Sans Serif"
   Set m_fntSymbol = New StdFont
   m_fntSymbol.Name = "Marlett"
   m_fntSymbol.Size = m_fnt.Size * 1.2
End Sub

Private Sub Class_Terminate()
   Set m_cMemDC = Nothing
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ISubclass_MsgResponse = emrConsume
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim iMenu As Long
Dim iLastDownOn As Long
Dim iLastOver As Long
Dim lR As Long
Dim lFlag As Long
Dim hMenu As Long
Dim iChar As Long

   Select Case iMsg
   Case WM_LBUTTONDOWN
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      ' If in range, then...
      iMenu = MenuHitTest()
      iLastDownOn = m_iDownOn
      m_iDownOn = iMenu
      If m_iDownOn <> iLastDownOn Then
         ' !Repaint!
         'Debug.Print "Repaint"
         SendMessageLong m_hWnd, WM_NCPAINT, 0, 0
      End If
      
      If m_iDownOn > 0 Then
         m_cTmr.Interval = 0
         lR = m_cToolbarMenu.TrackPopup(m_iDownOn)
         pRestoreList
      End If
      
   Case WM_MOUSEMOVE
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      pMouseMove
      
   Case WM_MEASUREITEM
      ISubclass_WindowProc = MeasureItem(wParam, lParam)
   
   Case WM_DRAWITEM
      DrawItem wParam, lParam
      
   Case WM_MENUCHAR
      ' Check that this is my menu:
      lFlag = wParam \ &H10000
      If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
         hMenu = lParam
         iChar = (wParam And &HFFFF&)
         ' See if this corresponds to an accelerator on the menu:
         lR = ParseMenuChar(hMenu, iChar)
         If lR > 0 Then
            ISubclass_WindowProc = lR
            Exit Function
         End If
      End If
      ISubclass_WindowProc = CallOldWindowProc(m_hWnd, WM_MENUCHAR, wParam, lParam)
   
   End Select
   
End Function
Private Function ParseMenuChar( _
        ByVal hMenu As Long, _
        ByVal iChar As Integer _
    ) As Long
Dim sChar As String
Dim l As Long
Dim lH() As Long
Dim sItems() As String

   'Debug.Print "WM_MENUCHAR"
   sChar = UCase$(Chr$(iChar))
   For l = 1 To m_iRestore
      If (m_hMenuRestore(l) = hMenu) Then
         If (m_sAccelerator(l) = sChar) Then
            ParseMenuChar = &H20000 Or m_iMenuPosition(l)
            ' Debug.Print "Found Menu Char"
            Exit Function
         End If
      End If
   Next l

End Function

Private Function MeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tMIS As MEASUREITEMSTRUCT
Dim lID As Long
   CopyMemory tMIS, ByVal lParam, LenB(tMIS)
   If tMIS.CtlType = ODT_MENU Then
                  
      ' because we don't get the popup menu handle
      ' in the tMIS structure, we have to do an internal
      ' lookup to find info about this menu item.
      ' poor implementation of MEASUREITEMSTRUCT - it
      ' should have a .hWndItem field like DRAWITEMSTRUCT
      ' - spm
      lID = InternalIDForWindowsID(tMIS.itemID)
            
      ' Width:
      tMIS.itemWidth = 4 + 22 + m_lMenuTextSize(lID) + 4
      If m_lMenuShortCutSize(lID) > 0 Then
         tMIS.itemWidth = tMIS.itemWidth + 4 + m_lMenuShortCutSize(lID) + 4
      End If
      
      ' Height:
      If lID > 0 And lID <= m_iRestore Then
         If (m_tMIIS(lID).fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
            tMIS.itemHeight = 6
         Else
            ' menu item height is always the same
            tMIS.itemHeight = m_lMenuItemHeight + 8
         End If
      Else
         ' problem.
      End If
      
      CopyMemory ByVal lParam, tMIS, LenB(tMIS)
      
   Else
      MeasureItem = CallOldWindowProc(m_hWnd, WM_MEASUREITEM, wParam, lParam)
   End If
End Function
Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tDIS As DRAWITEMSTRUCT
Dim hBr As Long
Dim tR As RECT, tTR As RECT, tWR As RECT
Dim lhDC As Long
Dim hFntOld As Long
Dim tMII As MENUITEMINFO
Dim bRadioCheck As Boolean, bDisabled As Boolean, bChecked As Boolean, bHighlighted As Boolean
Dim lID As Long
Dim hFntS As Long, hFntSOld As Long

   CopyMemory tDIS, ByVal lParam, LenB(tDIS)
   If tDIS.CtlType = ODT_MENU Then
      ' Todo
      ' tDIS.hWndItem is the menu containing the item, tDIS.itemID is the wID
      
      m_cMemDC.Width = tDIS.rcItem.right - tDIS.rcItem.left + 1
      m_cMemDC.Height = tDIS.rcItem.bottom - tDIS.rcItem.top + 1
      lhDC = m_cMemDC.hdc
      hFntOld = SelectObject(lhDC, hFont)
      
      LSet tR = tDIS.rcItem
      OffsetRect tR, -tR.left, -tR.top
      
      ' Fill background:
      tTR.right = m_cMemDC.Width
      tTR.bottom = m_cMemDC.Height
      hBr = CreateSolidBrush(TranslateColor(m_oMenuBackgroundColor))
      FillRect lhDC, tTR, hBr
      DeleteObject hBr
      
      SetBkMode lhDC, TRANSPARENT
      
      ' Draw the text:
      tMII.cbSize = LenB(tMII)
      tMII.fMask = MIIM_TYPE Or MIIM_STATE Or MIIM_DATA
      GetMenuItemInfo tDIS.hwndItem, tDIS.itemID, False, tMII
      
      If (tMII.fType And MFT_SEPARATOR) = MFT_SEPARATOR Then
         ' Separator:
         LSet tWR = tR
         tWR.top = (tWR.bottom - tWR.top - 2) \ 2 + tWR.top
         tWR.bottom = tWR.top + 2
         InflateRect tWR, -8, 0
         DrawEdge lhDC, tWR, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM
      Else
         ' Text item:
         bRadioCheck = ((tMII.fType And MFT_RADIOCHECK) = MFT_RADIOCHECK)
         bDisabled = ((tMII.fState And MFS_DISABLED) = MFS_DISABLED)
         bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
         bHighlighted = ((tMII.fState And MFS_HILITE) = MFS_HILITE)
         If bHighlighted Then
            SetTextColor lhDC, TranslateColor(m_oActiveMenuColorOver)
         Else
            SetTextColor lhDC, TranslateColor(m_oActiveMenuColor)
         End If
         
         ' Check:
         If bChecked Then
            LSet tWR = tR
            InflateRect tWR, -4, -4
            tWR.left = tWR.left + 2
            tWR.right = tWR.left + (tWR.bottom - tWR.top + 1)
            DrawEdge lhDC, tWR, BDR_SUNKENOUTER, BF_RECT
            
            SelectObject lhDC, hFntOld
            hFntSOld = SelectObject(lhDC, hFontSymbol)
            If bRadioCheck Then
               pDrawItem lhDC, "h", tWR, bDisabled, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
            Else
               pDrawItem lhDC, "b", tWR, bDisabled, DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
            End If
            SelectObject lhDC, hFntSOld
            hFntOld = SelectObject(lhDC, hFont)
            
         End If
         
         ' Draw text:
         LSet tWR = tR
         tWR.left = 20 + 4
         lID = tMII.dwItemData
         If lID > 0 And lID <= m_iRestore Then
            pDrawItem lhDC, m_sCaption(lID), tWR, bDisabled, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
            If Len(m_sShortCut(lID)) > 0 Then
               tWR.left = tWR.left + m_lMenuTextSize(lID) + 4 + 4
               pDrawItem lhDC, m_sShortCut(lID), tWR, bDisabled, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER
            End If
         End If
         
         ' Highlighted:
         If bHighlighted And Not (bDisabled) Then
            LSet tWR = tR
            InflateRect tWR, -2, 0
            DrawEdge lhDC, tWR, BDR_RAISEDINNER, BF_RECT
         End If
         
      End If
      
      SelectObject lhDC, hFntOld
      
      BitBlt tDIS.hdc, tDIS.rcItem.left, tDIS.rcItem.top, tDIS.rcItem.right - tDIS.rcItem.left + 1, tDIS.rcItem.bottom - tDIS.rcItem.top + 1, lhDC, 0, 0, vbSrcCopy
      
   Else
      DrawItem = CallOldWindowProc(m_hWnd, WM_DRAWITEM, wParam, lParam)
   End If
End Function
Private Sub pDrawItem( _
      ByVal lhDC As Long, _
      ByVal sText As String, _
      ByRef tR As RECT, _
      ByVal bDisabled As Boolean, _
      ByVal dtFlags As Long _
   )
Dim tWR As RECT
   LSet tWR = tR
   If bDisabled Then
      SetTextColor lhDC, TranslateColor(vb3DHighlight)
      OffsetRect tWR, 1, 1
      DrawText lhDC, sText, -1, tWR, dtFlags
      SetTextColor lhDC, TranslateColor(vbButtonShadow)
      OffsetRect tWR, -1, -1
      DrawText lhDC, sText, -1, tWR, dtFlags
   Else
      DrawText lhDC, sText, -1, tWR, dtFlags
   End If
End Sub
Private Sub pMouseMove()
Dim iMenu As Long
Dim iLastOver As Long
   iMenu = MenuHitTest()
   iLastOver = m_iOver
   m_iOver = iMenu
   'Debug.Print "Over:", m_iOver, iLastOver
   If m_iOver <> iLastOver Then
      ' !Repaint!
      'Debug.Print "Repaint"
      SendMessageLong m_hWnd, WM_NCPAINT, 0, 0
   End If
   If m_cTmr Is Nothing Then
      Set m_cTmr = New CTimer
   End If
   If m_iOver < 1 And m_iDownOn = 0 Then
      m_cTmr.Interval = 0
   Else
      If m_iDownOn > 0 Then
         If GetAsyncKeyState(vbLeftButton) = 0 Then
            m_iDownOn = 0
            SendMessageLong m_hWnd, WM_NCPAINT, 0, 0
         End If
      End If
      m_cTmr.Interval = 50
   End If
End Sub

Private Sub m_cTmr_ThatTime()
   pMouseMove
End Sub
' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function



