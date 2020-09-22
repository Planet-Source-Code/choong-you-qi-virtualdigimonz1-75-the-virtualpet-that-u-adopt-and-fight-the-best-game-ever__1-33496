VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNCCalcSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' cNCCalcSize
'
' Copyright © 2000 Steve McMahon (steve@vbaccelerator.com)
'
' Allows you to significantly modify the title and
' borders for a window.
'
' -------------------------------------------------------------------------
' Visit vbAccelerator at http://vbaccelerator.com
' =========================================================================

Private Type POINTS
   x  As Integer
   y  As Integer
End Type
Private Type WINDOWPOS
   hwnd As Long
   hWndInsertAfter As Long
   x As Long
   y As Long
   cx As Long
   cy As Long
   flags As Long
End Type
Private Type NCCALCSIZE_PARAMS
   rgrc(0 To 2) As RECT
   lppos As Long 'WINDOWPOS
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, ByVal hwndTo As Long, lppt As Any, ByVal cPoints As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, lpsz2 As Any) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Any) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function DrawFrameControl Lib "user32" (ByVal lhDC As Long, tR As RECT, ByVal eFlag As Long, ByVal eStyle As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawCaptionAPI Lib "user32" Alias "DrawCaption" (ByVal hwnd As Long, ByVal hDC As Long, pcRect As RECT, ByVal un As Long) As Long

Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

' mouseevent
Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
Private Const MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
Private Const MOUSEEVENTF_MOVE = &H1 '  mouse move
Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up

' SysMetrics
Private Const SM_CXBORDER = 5
Private Const SM_CXDLGFRAME = 7
Private Const SM_CXFIXEDFRAME = SM_CXDLGFRAME
Private Const SM_CXFRAME = 32
Private Const SM_CXHSCROLL = 21
Private Const SM_CXVSCROLL = 2
Private Const SM_CYCAPTION = 4
Private Const SM_CYDLGFRAME = 8
Private Const SM_CYFIXEDFRAME = SM_CYDLGFRAME
Private Const SM_CYFRAME = 33
Private Const SM_CYHSCROLL = 3
Private Const SM_CYMENU = 15
Private Const SM_CYSMSIZE = 31
Private Const SM_CXSMSIZE = 30

' DrawFrameControl:
Private Const DFC_CAPTION = 1
Private Const DFC_MENU = 2
Private Const DFC_SCROLL = 3
Private Const DFC_BUTTON = 4
'#if(WINVER >= =&H0500)
Private Const DFC_POPUPMENU = 5
'#endif /* WINVER >= =&H0500 */

Private Const DFCS_CAPTIONCLOSE = &H0
Private Const DFCS_CAPTIONMIN = &H1
Private Const DFCS_CAPTIONMAX = &H2
Private Const DFCS_CAPTIONRESTORE = &H3
Private Const DFCS_CAPTIONHELP = &H4

Private Const DFCS_INACTIVE = &H100
Private Const DFCS_PUSHED = &H200
Private Const DFCS_CHECKED = &H400

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

' Map WIndow Points
Private Const HWND_DESKTOP = 0

' Redraw window:
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_ERASE = &H4
Private Const RDW_ERASENOW = &H200
Private Const RDW_FRAME = &H400
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_INVALIDATE = &H1
Private Const RDW_NOCHILDREN = &H40
Private Const RDW_NOERASE = &H20
Private Const RDW_NOFRAME = &H800
Private Const RDW_NOINTERNALPAINT = &H10
Private Const RDW_UPDATENOW = &H100
Private Const RDW_VALIDATE = &H8

' Sys colours:
Private Const COLOR_WINDOWFRAME = 6
Private Const COLOR_BTNFACE = 15
Private Const COLOR_BTNTEXT = 18
Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_ACTIVEBORDER = 10
Private Const COLOR_ACTIVECAPTION = 2
Private Const COLOR_INACTIVEBORDER = 11

' Window MEssages
Private Const WM_DESTROY = &H2
Private Const WM_SETTEXT = &HC
Private Const WM_ACTIVATEAPP = &H1C
Private Const WM_SETCURSOR = &H20
Private Const WM_CHILDACTIVATE = &H22
Private Const WM_STYLECHANGING = &H7C
Private Const WM_STYLECHANGED = &H7D
Private Const WM_NCCALCSIZE = &H83
Private Const WM_NCPAINT = &H85
Private Const WM_NCHITTEST = &H84
Private Const WM_NCACTIVATE = &H86
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCLBUTTONUP = &HA2
Private Const WM_NCLBUTTONDBLCLK = &HA3
Private Const WM_SYSCOMMAND = &H112
Private Const WM_INITMENU = &H116
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_MDIGETACTIVE = &H229

' flags for DrawCaption
Private Const DC_ACTIVE = &H1
Private Const DC_SMALLCAP = &H2
Private Const DC_ICON = &H4
Private Const DC_TEXT = &H8
Private Const DC_INBUTTON = &H10
Private Const DC_GRADIENT = &H20

' WM_NCCALCSIZE return values;
Private Const WVR_ALIGNBOTTOM = &H40
Private Const WVR_ALIGNLEFT = &H20
Private Const WVR_ALIGNRIGHT = &H80
Private Const WVR_ALIGNTOP = &H10
Private Const WVR_HREDRAW = &H100
Private Const WVR_VALIDRECTS = &H400
Private Const WVR_VREDRAW = &H200
Private Const WVR_REDRAW = (WVR_HREDRAW Or WVR_VREDRAW)

' Window Long:
Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_USERDATA = (-21)
Private Const GWL_WNDPROC = -4
Private Const GWL_HWNDPARENT = (-8)

'Window Styles:
Private Const WS_THICKFRAME = &H40000
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_CHILD = &H40000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_BORDER = &H800000
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TOOLWINDOW = &H80&
Private Const CW_USEDEFAULT = &H80000000

' SetWIndowPos
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREDRAW = &H8
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOCOPYBITS = &H100
Private Const SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SWP_NOZORDER = &H4

Implements ISubclass

Public Enum ECNCSysCommandConstants
   SC_ARRANGE = &HF110&
   SC_CLOSE = &HF060&
   SC_MAXIMIZE = &HF030&
   SC_MINIMIZE = &HF020&
   SC_MOVE = &HF010&
   SC_NEXTWINDOW = &HF040&
   SC_PREVWINDOW = &HF050&
   SC_RESTORE = &HF120&
   SC_SIZE = &HF000&
End Enum

Public Enum ECNCHitTestConstants
   HTBORDER = 18
   HTBOTTOM = 15
   HTBOTTOMLEFT = 16
   HTBOTTOMRIGHT = 17
   HTCAPTION = 2
   HTCLIENT = 1
   HTGROWBOX = 4
   HTHSCROLL = 6
   HTLEFT = 10
   HTMAXBUTTON = 9
   HTMENU = 5
   HTMINBUTTON = 8
   HTNOWHERE = 0
   HTRIGHT = 11
   HTSYSMENU = 3
   HTTOP = 12
   HTTOPLEFT = 13
   HTTOPRIGHT = 14
   HTVSCROLL = 7
End Enum


' Window handles:
Private m_hWnd As Long
Private m_hWndMDIClient As Long
Private m_bIsMDIChild As Boolean

' Menu handle
Private m_hMenu As Long
' App activate & window activation state:
Private m_bActive As Boolean
Private m_bAppActive As Boolean
' Is our MDI Child zoomed in or not?
Private m_bZoomedMDIChild As Boolean
' MemDC for title bar drawing:
Private m_hDC As Long
Private m_hBmp As Long
Private m_hBmpOld As Long
' Maximized MDI Child?
Private m_bState As Boolean
' Borders:
Private m_lLeft As Long, m_lTop As Long
Private m_lRight As Long, m_lBottom As Long
' Last HitTest result
Private m_eLastHT As ECNCHitTestConstants

Public Sub Redraw(hwnd As Long)
   RedrawWindow hwnd, ByVal 0&, 0, RDW_UPDATENOW Or RDW_INVALIDATE Or RDW_INTERNALPAINT Or RDW_ALLCHILDREN
End Sub
Public Sub Display(f As Object)
   'f.Show
   On Error Resume Next
   f.Refresh
   SetWindowPos m_hWnd, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED
End Sub
Public Property Get WindowActive() As Boolean
   WindowActive = m_bActive
End Property
Public Property Get AppActive() As Boolean
   AppActive = m_bAppActive
End Property

Public Sub TitleBarMouseDown()
Dim tPS As POINTS
Dim tP As POINTAPI
   GetCursorPos tP
   tPS.x = tP.x: tPS.y = tP.y
   ReleaseCapture
   SendMessage m_hWnd, WM_NCLBUTTONDOWN, HTCAPTION, tPS
End Sub
Public Sub SysCommand(ByVal eCmd As ECNCSysCommandConstants)
   PostMessage m_hWnd, WM_SYSCOMMAND, eCmd, 0
End Sub

Public Sub Attach(ByVal iTo As INCAreaModifier)
Dim lhDC As Long
   
   Detach
   
   m_hWnd = iTo.hwnd
   m_hMenu = GetMenu(m_hWnd)
   
   m_bIsMDIChild = IsMDIChildForm(m_hWnd)
   
   ' Allows us to remove menu bar, caption etc:
   AttachMessage Me, m_hWnd, WM_NCCALCSIZE
   ' Handle drawing borders, caption etc ourselves:
   AttachMessage Me, m_hWnd, WM_NCPAINT
   ' Win redraws caption during NCACTIVATE:
   AttachMessage Me, m_hWnd, WM_NCACTIVATE
   ' On NC Button Down, Win redraws the min/max/close buttons:
   AttachMessage Me, m_hWnd, WM_NCLBUTTONDOWN
   ' Check for button up so we can notify client:
   AttachMessage Me, m_hWnd, WM_NCLBUTTONUP
   ' on NC double click, Win redraws the min/max/close buttons:
   AttachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
   ' Allows us to use the default implementations
   ' for hittest events:
   AttachMessage Me, m_hWnd, WM_NCHITTEST
   ' Hack:
   AttachMessage Me, m_hWnd, WM_SETCURSOR
   ' On SysMenu Show, Win redraws the min/max/close buttons:
   AttachMessage Me, m_hWnd, WM_INITMENU
   AttachMessage Me, m_hWnd, WM_INITMENUPOPUP
   ' On ChangeStyle, Win redraws the entire caption:
   AttachMessage Me, m_hWnd, WM_STYLECHANGED
   ' On SetText, Win redraws the entire caption:
   AttachMessage Me, m_hWnd, WM_SETTEXT
   ' Checking for activateapp:
   AttachMessage Me, m_hWnd, WM_ACTIVATEAPP
   ' EnterMenuLoop
   AttachMessage Me, m_hWnd, WM_ENTERMENULOOP
   ' ExitMenuLoop
   AttachMessage Me, m_hWnd, WM_EXITMENULOOP
   
   If m_bIsMDIChild Then
      AttachMessage Me, m_hWnd, WM_SIZE
   End If
   
   ' So we can automatically detach ourselves when the parent closes:
   AttachMessage Me, m_hWnd, WM_DESTROY
   
   
   lhDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   m_hDC = CreateCompatibleDC(lhDC)
   m_hBmp = CreateCompatibleBitmap(lhDC, Screen.Width \ Screen.TwipsPerPixelX, GetSystemMetrics(SM_CYCAPTION) * 4)
   DeleteDC lhDC
   m_hBmpOld = SelectObject(m_hDC, m_hBmp)
   
   m_hWndMDIClient = FindWindowEx(m_hWnd, 0, "MDIClient", ByVal 0&)
      
   SetProp m_hWnd, "vbalCNCImplementation", ObjPtr(iTo)
   
   AttachKeyboardHook Me
End Sub
Public Property Get hMenu() As Long
   hMenu = m_hMenu
End Property
Public Sub Detach()
   DetachKeyboardHook Me
   If m_hWnd <> 0 Then
      DetachMessage Me, m_hWnd, WM_NCCALCSIZE
      DetachMessage Me, m_hWnd, WM_NCPAINT
      DetachMessage Me, m_hWnd, WM_NCACTIVATE
      DetachMessage Me, m_hWnd, WM_NCLBUTTONDOWN
      DetachMessage Me, m_hWnd, WM_NCLBUTTONUP
      DetachMessage Me, m_hWnd, WM_NCLBUTTONDBLCLK
      DetachMessage Me, m_hWnd, WM_NCHITTEST
      
      DetachMessage Me, m_hWnd, WM_SETCURSOR
      
      DetachMessage Me, m_hWnd, WM_INITMENU
      DetachMessage Me, m_hWnd, WM_INITMENUPOPUP
      
      DetachMessage Me, m_hWnd, WM_STYLECHANGED
      DetachMessage Me, m_hWnd, WM_SETTEXT
      
      DetachMessage Me, m_hWnd, WM_ACTIVATEAPP
      
      DetachMessage Me, m_hWnd, WM_ENTERMENULOOP
      DetachMessage Me, m_hWnd, WM_EXITMENULOOP
      
      If m_bIsMDIChild Then
         DetachMessage Me, m_hWnd, WM_SIZE
         m_bIsMDIChild = False
      End If
      
      DetachMessage Me, m_hWnd, WM_DESTROY
   End If
   If m_hDC <> 0 Then
      If m_hBmpOld <> 0 Then
         SelectObject m_hDC, m_hBmp
         m_hBmpOld = 0
      End If
      If m_hBmp <> 0 Then
         DeleteObject m_hBmp
         m_hBmp = 0
      End If
      If m_hDC <> 0 Then
         DeleteDC m_hDC
         m_hDC = 0
      End If
   End If
   RemoveProp m_hWnd, "vbalCNCImplementation"
   m_hWnd = 0
   m_hWndMDIClient = 0
   m_hMenu = 0
End Sub

Friend Function AltKeyAccelerator(ByVal vKey As KeyCodeConstants) As Long
Dim Implementation As INCAreaModifier
   If GetImplementation(Implementation) Then
      AltKeyAccelerator = Implementation.AltKeyAccelerator(vKey)
   End If
End Function

Private Sub pShowMDIButtons(ByVal hwnd As Long, ByVal bState As Boolean)
   m_bState = bState
End Sub

Private Sub MyMoveWindow()
Dim tPInit As POINTAPI
Dim tPLast As POINTAPI
Dim tP As POINTAPI
Dim tR As RECT
Dim hWndParent As Long
Dim tWRInit As RECT
Dim dx As Long, dy As Long
      
   GetWindowRect m_hWnd, tR
   hWndParent = GetParent(m_hWnd)
   If Not hWndParent = 0 Then
      MapWindowPoints HWND_DESKTOP, hWndParent, tR, 2
   End If
   GetCursorPos tPInit
   LSet tPLast = tPInit
   Do While Not (GetAsyncKeyState(vbLeftButton) = 0) And m_bActive
      GetCursorPos tP
      If tP.x <> tPLast.x Or tP.y <> tPLast.y Then
         ' Moved:
         dx = tP.x - tPLast.x
         dy = tP.y - tPLast.y
         SetWindowPos m_hWnd, 0, tR.left + dx, tR.top + dy, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
         LSet tPLast = tP
         GetWindowRect m_hWnd, tR
         If Not hWndParent = 0 Then
            MapWindowPoints HWND_DESKTOP, hWndParent, tR, 2
         End If
      End If
      DoEvents
      Sleep 1
   Loop
   
End Sub

Private Sub Class_Terminate()
   Detach
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
   '
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   
   Select Case CurrentMessage
   Case WM_NCPAINT, WM_NCLBUTTONDOWN, _
         WM_NCLBUTTONDBLCLK, _
         WM_INITMENUPOPUP, WM_INITMENU, _
         WM_SETCURSOR, WM_CHILDACTIVATE, _
         WM_STYLECHANGED, WM_SETTEXT, _
         WM_NCHITTEST, WM_SIZE, _
         WM_ENTERMENULOOP, WM_EXITMENULOOP
      ISubclass_MsgResponse = emrConsume
   Case Else
      ' ActiveApp, Destroy:
      ISubclass_MsgResponse = emrPreprocess
   End Select
   
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim tNCR As NCCALCSIZE_PARAMS
Dim tWP As WINDOWPOS
Dim tP As POINTAPI
Dim tR As RECT
Dim lhWnd As Long
Dim lpfMaximised As Long
Dim lPtr As Long
Dim hDC As Long
Dim lStyle As Long
Dim eHt As ECNCHitTestConstants
Static s_dx As Long
Static s_dy As Long
Dim bCanSize As Boolean
Dim Implementation As INCAreaModifier
Dim bHandled As Boolean
Static s_bNoStyleChangeProcessing As Boolean
Static s_bChildActivate As Boolean

   Select Case iMsg
   
   Case WM_DESTROY
      ' Goodbye!
      Detach
   
   Case WM_NCPAINT
      
      ' Due to processing elsewhere in this subclass, we
      ' might inadvertently be drawing when the window
      ' is being closed or invisible.  Check before
      ' drawing:
      If Not (IsWindowVisible(hwnd) = 0) Then
         m_bZoomedMDIChild = (IsMDIChildForm(hwnd) And (IsZoomed(hwnd) <> 0))
         If m_bZoomedMDIChild Then
            ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         Else
            ' Get the non-client DC to draw in:
            hDC = GetWindowDC(m_hWnd)
                        
            GetWindowRect m_hWnd, tR
            OffsetRect tR, -tR.left, -tR.top
                     
            If GetImplementation(Implementation) Then
               Implementation.NCPaint hDC, tR.left, tR.top, tR.right, tR.bottom
            Else
               DefaultNCPaint hDC, tR.left, tR.top, tR.right, tR.bottom
            End If
            
            ReleaseDC m_hWnd, hDC
         End If
      End If
         
   Case WM_NCHITTEST
      
      If GetImplementation(Implementation) Then
         eHt = pGetHitTestCode()
         m_eLastHT = eHt
         If eHt = HTMENU Then
            ' Cannot allow windows to have this; if you
            ' mouse down on menu or caption then windows
            ' redraws the caption on top...
            ISubclass_WindowProc = HTCLIENT
         Else
            ISubclass_WindowProc = eHt
         End If
         
      Else
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
      
   Case WM_NCLBUTTONDOWN
      '
      ' a hack.
      '
      ' Win suspends when we do a NC Button down.  It also
      ' redraws the min/max/close buttons. We can force them
      ' to go away by moving the mouse
      '
      If s_dx = 0 Then s_dx = 1
      If s_dy = 0 Then s_dy = 1
      s_dx = -1 * s_dx: s_dy = -1 * s_dy
      mouse_event MOUSEEVENTF_MOVE, s_dx, s_dy, 0, 0
      
      ' We cannot allow Windows to do the default HTCAPTION action,
      ' because it redraws the caption during the move.  THerefore
      ' swallow HTCAPTION events and reimplement window moving
      ' ourselves:
      wParam = pGetHitTestCode()
      If GetImplementation(Implementation) Then
         If m_bActive Then
            If m_eLastHT = HTCAPTION Then
               MyMoveWindow
               Exit Function
            End If
         Else
            If m_eLastHT = HTCAPTION Then
               SetForegroundWindow m_hWnd
               MyMoveWindow
               Exit Function
            End If
         End If
         
         GetCursorPos tP
         GetWindowRect m_hWnd, tR
         tP.x = tP.x - tR.left: tP.y = tP.y - tR.top
         OffsetRect tR, -tR.left, -tR.top
         hDC = GetWindowDC(m_hWnd)
         Implementation.NCMouseDown tP.x, tP.y, bHandled, hDC, tR.left, tR.top, tR.right, tR.bottom
         ReleaseDC m_hWnd, hDC
         If bHandled Then
            Exit Function
         End If
                  
      End If
      
      ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
   
   Case WM_NCLBUTTONUP
      If GetImplementation(Implementation) Then
         GetCursorPos tP
         GetWindowRect m_hWnd, tR
         tP.x = tP.x - tR.left: tP.y = tP.y - tR.top
         OffsetRect tR, -tR.left, -tR.top
         hDC = GetWindowDC(m_hWnd)
         Implementation.NCMouseDown tP.x, tP.y, bHandled, hDC, tR.left, tR.top, tR.right, tR.bottom
         ReleaseDC m_hWnd, hDC
         Implementation.NCMouseUp tP.x, tP.y, hDC, tR.left, tR.top, tR.right, tR.bottom
      End If
         
   Case WM_SETCURSOR
      '
      ' a Very Nasty Hack :)
      ' discovered by watching NeoPlanet and MSOffice
      ' in Spy++
      '
      ' Without this, Win will redraw caption areas and
      ' min/max/close buttons whenever the mouse is released
      ' following a NC mouse down.
      '
      s_bNoStyleChangeProcessing = True
      lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
      SetWindowLong m_hWnd, GWL_STYLE, lStyle And Not WS_VISIBLE
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      If GetMenu(m_hWnd) <> 0 Then
         SetMenu m_hWnd, 0
      End If
      SetWindowLong m_hWnd, GWL_STYLE, lStyle
      s_bNoStyleChangeProcessing = False
   
   Case WM_INITMENU
      ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
      
   Case WM_CHILDACTIVATE
      If Not s_bChildActivate Then
         s_bChildActivate = True
         ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
         s_bChildActivate = False
      End If
      
   Case WM_SIZE
      '
      ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
   
   Case WM_INITMENUPOPUP
      '
      ' During a WM_INITMENUPOPUP, the system redraws the
      ' min/max/close buttons.
      
      
      ' Check HiWord of lParam to see whether this is
      ' a SysMenu:
      If Not (lParam And &HFFFF0000) = 0 Then
         ' Sys Menu:
         ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
      Else
         ' App Menu:
         ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
         If GetImplementation(Implementation) Then
            Implementation.InitMenuPopup wParam, lParam
         End If
      End If
      
   Case WM_ENTERMENULOOP, WM_EXITMENULOOP
      ISubclass_MsgResponse = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
      If iMsg = WM_EXITMENULOOP Then
         If GetImplementation(Implementation) Then
            Implementation.ExitMenuLoop
         End If
      End If
      
   Case WM_SETTEXT, WM_STYLECHANGED, WM_NCLBUTTONDBLCLK
      '
      ' The whole title bar is repainted by the defwindowproc.
      ' Therefore redraw once complete:
      If Not s_bNoStyleChangeProcessing Then
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
         ISubclass_WindowProc hwnd, WM_NCPAINT, 0, 0
      Else
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
      
   Case WM_NCCALCSIZE
      '
      ' No Hacks!
      '
      ' This simply tells windows to modify the client
      ' area to the appropriate size:
      '
      
      ' First set the zoomed MDI Child flag:
      m_bZoomedMDIChild = (IsMDIChildForm(hwnd) And (IsZoomed(hwnd) <> 0))
      If wParam <> 0 Then
      
         ' Get the structure pointed to by lParam:
         CopyMemory tNCR, ByVal lParam, Len(tNCR)
         CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)
         
         'pDebugCalcSize tNCR
         With tNCR.rgrc(0)
            ' Set these
            .left = tWP.x
            .top = tWP.y
            .right = tWP.x + tWP.cx
            .bottom = tWP.y + tWP.cy
                        
            ' Defaults
            m_lLeft = GetSystemMetrics(SM_CXFRAME)
            m_lRight = m_lLeft
            m_lTop = GetSystemMetrics(SM_CYCAPTION) + GetSystemMetrics(SM_CYFRAME)
            m_lBottom = GetSystemMetrics(SM_CYFRAME)
            
            ' If the window in question is an MDI child, then we
            ' ant to ensure that the standard settings get sent
            ' back to windows: to prevent drawing additional borders,
            ' which aren't required:
            If Not m_bZoomedMDIChild Then
               ' If the implementation is valid then request the
               ' physical size of the title bar and borders:
               If GetImplementation(Implementation) Then
                  Implementation.GetLeftMarginWidth m_lLeft
                  Implementation.GetTopMarginHeight m_lTop
                  Implementation.GetRightMarginWidth m_lRight
                  Implementation.GetBottomMarginHeight m_lBottom
               End If
            End If
            
            ' Set our physical left/top/right/bottom values:
            .left = .left + m_lLeft
            .top = .top + m_lTop
            .right = .right - m_lRight
            .bottom = .bottom - m_lBottom
         End With
         
         ' Return the new client area size to windows:
         LSet tNCR.rgrc(1) = tNCR.rgrc(0)
         CopyMemory ByVal lParam, tNCR, Len(tNCR)
         ISubclass_WindowProc = WVR_VALIDRECTS
         
      Else
         ' lParam points to a rectangle
         ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      End If
               
      ' Check for the active window:
      'lPtr = VarPtr(lpfMaximised)
      'If Not m_hWndMDIClient = 0 Then
      '   lhWnd = SendMessageLong(m_hWndMDIClient, WM_MDIGETACTIVE, 0, lPtr)
      '   pShowMDIButtons lhWnd, (lpfMaximised <> 0)
      'End If
      
   Case WM_NCACTIVATE
      '
      ' When we get a NC Activate The title bar is
      ' being redrawn to show active or inactive states.
      '
      ' This processing ensures the title bar is updated
      ' correctly following state change:
      '
      
      ' We must call the defwindowproc otherwise VB goes
      ' funny.  This draws a full titlebar:
      m_bActive = Not (wParam = 0)
      ISubclass_WindowProc = CallOldWindowProc(hwnd, iMsg, wParam, lParam)
      
      ' Now fix it:
      ISubclass_WindowProc m_hWnd, WM_NCPAINT, 0, 0
   
   Case WM_ACTIVATEAPP
      '
      ' This is for detecting which app is active
      '
      m_bAppActive = Not (wParam = 0)
   
   End Select
            
End Function
Private Function IsMDIChildForm(ByVal hwnd As Long) As Boolean
Dim hWndP As Long
Dim sBuf As String
Dim iPos As Long
   hWndP = GetParent(hwnd)
   sBuf = String$(260, 0)
   GetClassName hWndP, sBuf, 259
   iPos = InStr(sBuf, vbNullChar)
   If iPos > 1 Then
      If left$(sBuf, iPos - 1) = "MDIClient" Then
         IsMDIChildForm = True
      End If
   End If
End Function
Private Function pGetHitTestCode() As ECNCHitTestConstants
Dim lStyle As Long
Dim bCanSize As Boolean
Dim Implementation As INCAreaModifier
Dim eHt As ECNCHitTestConstants
Dim tP As POINTAPI
Dim tR As RECT
   
   If GetImplementation(Implementation) Then
      lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
      bCanSize = ((lStyle And WS_SIZEBOX) = WS_SIZEBOX)
      eHt = HTCLIENT
      GetCursorPos tP
      
      GetWindowRect m_hWnd, tR
      tP.x = tP.x - tR.left: tP.y = tP.y - tR.top
      OffsetRect tR, -tR.left, -tR.top
      eHt = HTCLIENT
      If Not (PtInRect(tR, tP.x, tP.y) = 0) Then
         ' Left
         If tP.x <= m_lLeft Then
            If tP.y <= m_lBottom Then
               If bCanSize Then
                  eHt = HTTOPLEFT
               End If
            ElseIf tP.y >= tR.bottom - m_lBottom Then
               If bCanSize Then
                  eHt = HTBOTTOMLEFT
               End If
            Else
               If bCanSize Then
                  eHt = HTLEFT
               End If
            End If
         ' Right
         ElseIf tP.x >= tR.right - m_lRight Then
            If tP.y <= m_lBottom Then
               If bCanSize Then
                  eHt = HTTOPRIGHT
               End If
            ElseIf tP.y >= tR.bottom - m_lBottom Then
               If bCanSize Then
                  eHt = HTBOTTOMRIGHT
               End If
            Else
               If bCanSize Then
                  eHt = HTRIGHT
               End If
            End If
         ' Top/Bottom?
         ElseIf tP.y <= m_lBottom Then
            If bCanSize Then
               eHt = HTTOP
            End If
         ElseIf tP.y >= tR.bottom - m_lBottom Then
            If bCanSize Then
               eHt = HTBOTTOM
            End If
         ' Caption/Menu
         ElseIf tP.y <= m_lTop Then
            ' We assume for default that the caption
            ' is the same as the system caption etc:
            If tP.y <= m_lBottom + GetSystemMetrics(SM_CYCAPTION) Then
               eHt = HTCAPTION
               If tP.x <= GetSystemMetrics(SM_CYCAPTION) Then
                  eHt = HTSYSMENU
               Else
                  ' todo min/max/close btns
               End If
            ElseIf tP.y > m_lBottom + GetSystemMetrics(SM_CYCAPTION) Then
               eHt = HTCLIENT
            End If
         End If
      End If
      Implementation.HitTest tP.x, tP.y, eHt
   End If
   pGetHitTestCode = eHt
   
End Function
Public Sub DefaultNCPaint(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
Dim tR As RECT, tTR As RECT, tSR As RECT, tBR As RECT
Dim lFlag As Long
Dim hBr As Long, hBrButton As Long

   tR.left = lLeft
   tR.top = lTop
   tR.right = lRight
   tR.bottom = lBottom
   LSet tBR = tR

   If m_bActive Then
      lFlag = DC_ACTIVE
      hBrButton = GetSysColorBrush(COLOR_ACTIVECAPTION)
      hBr = GetSysColorBrush(COLOR_ACTIVEBORDER)
   Else
      hBrButton = GetSysColorBrush(COLOR_INACTIVECAPTION)
      hBr = GetSysColorBrush(COLOR_INACTIVEBORDER)
   End If

   ' Titlebar area:
   ' Draw the part between the edge & the client:
   LSet tTR = tR
   ' left edge
   tTR.top = GetSystemMetrics(SM_CYFRAME)
   tTR.bottom = tTR.bottom - GetSystemMetrics(SM_CYFRAME)
   tTR.right = GetSystemMetrics(SM_CXFRAME)
   FillRect hDC, tTR, hBr
   ' top
   LSet tTR = tR
   tTR.bottom = GetSystemMetrics(SM_CYFRAME)
   FillRect hDC, tTR, hBr
   ' right
   LSet tTR = tR
   tTR.top = GetSystemMetrics(SM_CYFRAME)
   tTR.bottom = tTR.bottom - GetSystemMetrics(SM_CYFRAME)
   tTR.left = tTR.right - GetSystemMetrics(SM_CXFRAME)
   FillRect hDC, tTR, hBr
   ' bottom
   LSet tTR = tR
   tTR.top = tTR.bottom - GetSystemMetrics(SM_CYFRAME)
   FillRect hDC, tTR, hBr

   ' Draw the caption into the caption area:

   ' top bit under titlebar:
   LSet tTR = tR
   tTR.top = GetSystemMetrics(SM_CXFRAME) + GetSystemMetrics(SM_CYCAPTION) - 1
   tTR.bottom = tTR.top + 1
   FillRect hDC, tTR, hBr
   DeleteObject hBr

   ' Draw the titlebar into a work DC to prevent flicker:
   lFlag = lFlag Or DC_ICON Or DC_TEXT
   LSet tTR = tR
   tTR.left = tTR.left + GetSystemMetrics(SM_CXFRAME)
   tTR.right = tTR.right - GetSystemMetrics(SM_CXFRAME)
   tTR.top = tTR.top + GetSystemMetrics(SM_CYFRAME)
   tTR.bottom = tTR.top + GetSystemMetrics(SM_CYCAPTION) - 1
   LSet tR = tTR
   OffsetRect tR, -tR.left, -tR.top
   LSet tSR = tR
   tSR.right = tSR.right - (tR.bottom - tR.top) * 3 - 2
   DrawCaptionAPI m_hWnd, m_hDC, tSR, lFlag

   ' Draw the titlebar buttons:
   tSR.left = tSR.right
   tSR.right = tR.right
   FillRect m_hDC, tSR, hBrButton
   DeleteObject hBrButton

   InflateRect tR, 0, -2
   tR.right = tR.right - 2
   tR.left = tR.right - (tR.bottom - tR.top) - 2
   DrawFrameControl m_hDC, tR, DFC_CAPTION, DFCS_CAPTIONCLOSE
   OffsetRect tR, -(tR.right - tR.left + 2), 0
   If IsZoomed(m_hWnd) Then
      DrawFrameControl m_hDC, tR, DFC_CAPTION, DFCS_CAPTIONRESTORE
   Else
      DrawFrameControl m_hDC, tR, DFC_CAPTION, DFCS_CAPTIONMAX
   End If
   OffsetRect tR, -(tR.right - tR.left), 0
   DrawFrameControl m_hDC, tR, DFC_CAPTION, DFCS_CAPTIONMIN

   ' Finished drawing the NC area:
   BitBlt hDC, tTR.left, tTR.top, tTR.right - tTR.left, tTR.bottom - tTR.top, m_hDC, 0, 0, vbSrcCopy

   ' Edge 3d
   DrawEdge hDC, tBR, EDGE_RAISED, BF_RECT

End Sub

Public Function GetImplementation(iTo As INCAreaModifier) As Boolean
Dim lPtr As Long
   lPtr = GetProp(m_hWnd, "vbalCNCImplementation")
   If Not lPtr = 0 Then
      Dim iToTemp As INCAreaModifier
      CopyMemory iToTemp, lPtr, 4
      Set iTo = iToTemp
      CopyMemory iToTemp, 0&, 4
      GetImplementation = True
   End If
End Function


#If 0 = 1 Then
Private Sub pDebugCalcSize(ByRef tNCR As NCCALCSIZE_PARAMS)
Dim i As Long
Dim tWP As WINDOWPOS
   ' Use to show what is happening:
   With tNCR
      For i = 1 To 3
         With .rgrc(i - 1)
            Debug.Print .left, .top, .right, .bottom
         End With
      Next i
      CopyMemory tWP, ByVal .lppos, Len(tWP)
      With tWP
         Debug.Print .x, .y, .x + .cx, .y + .cy
      End With
      
   End With
End Sub
#End If


