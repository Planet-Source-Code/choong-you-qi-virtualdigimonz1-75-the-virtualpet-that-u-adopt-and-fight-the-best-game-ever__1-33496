Attribute VB_Name = "mDeclares"
Option Explicit

Public Type POINTAPI
   x As Long
   y As Long
End Type
Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Public Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Public Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Public Const TPM_CENTERALIGN = &H4&
Public Const TPM_LEFTALIGN = &H0&
Public Const TPM_LEFTBUTTON = &H0&
Public Const TPM_RIGHTALIGN = &H8&
Public Const TPM_RIGHTBUTTON = &H2&

Public Const TPM_NONOTIFY = &H80&           '/* Don't send any notification msgs */
Public Const TPM_RETURNCMD = &H100
Public Const TPM_HORIZONTAL = &H0          '/* Horz alignment matters more */
Public Const TPM_VERTICAL = &H40           '/* Vert alignment matters more */

   ' Win98/2000 menu animation and menu within menu options:
Public Const TPM_RECURSE = &H1&
Public Const TPM_HORPOSANIMATION = &H400&
Public Const TPM_HORNEGANIMATION = &H800&
Public Const TPM_VERPOSANIMATION = &H1000&
Public Const TPM_VERNEGANIMATION = &H2000&
   ' Win2000 only:
Public Const TPM_NOANIMATION = &H4000&

Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Public Declare Function TrackPopupMenuByLong Lib "user32" Alias "TrackPopupMenu" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByVal lprc As Long) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As TPMPARAMS) As Long

' Window MEssages
Public Const WM_DESTROY = &H2
Public Const WM_SIZE = &H5
Public Const WM_SETTEXT = &HC
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_CANCELMODE = &H1F
Public Const WM_SETCURSOR = &H20
Public Const WM_MEASUREITEM = &H2C
Public Const WM_DRAWITEM = &H2B
Public Const WM_STYLECHANGING = &H7C
Public Const WM_STYLECHANGED = &H7D
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCHITTEST = &H84
Public Const WM_NCPAINT = &H85
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_KEYDOWN = &H100
Public Const WM_COMMAND = &H111
Public Const WM_SYSCOMMAND = &H112
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_MENUSELECT = &H11F
Public Const WM_MENUCHAR = &H120
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_RBUTTONUP = &H205
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_EXITMENULOOP = &H212

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SetProp Lib "user32" Alias "SetPropA" (ByVal hwnd As Long, ByVal lpString As String, ByVal hData As Long) As Long
Public Declare Function RemoveProp Lib "user32" Alias "RemovePropA" (ByVal hwnd As Long, ByVal lpString As String) As Long

Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hMod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Const WH_KEYBOARD As Long = 2
Private Const WH_MSGFILTER As Long = (-1)
Private Const MSGF_MENU = 2
Private Const HC_ACTION = 0

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

' Message filter hook:
Private m_hMsgHook As Long
Private m_lMsgHookPtr As Long

' Keyboard Hook:
Private m_hKeyHook As Long
Private m_lKeyHookPtr() As Long
Private m_lKeyHookCount As Long

Public Sub AttachKeyboardHook(cN As cNCCalcSize)

Dim lpFn As Long
Dim lPtr As Long
Dim i As Long
   If m_hKeyHook = 0 Then
      lpFn = HookAddress(AddressOf KeyboardFilter)
      m_hKeyHook = SetWindowsHookEx(WH_KEYBOARD, lpFn, 0&, GetCurrentThreadId())
      Debug.Assert (m_hKeyHook <> 0)
   End If
   
   lPtr = ObjPtr(cN)
   If GetKeyHookPtrIndex(lPtr) = 0 Then
      m_lKeyHookCount = m_lKeyHookCount + 1
      ReDim Preserve m_lKeyHookPtr(1 To m_lKeyHookCount) As Long
      m_lKeyHookPtr(m_lKeyHookCount) = lPtr
   End If
   
End Sub
Private Function GetKeyHookPtrIndex(ByVal lPtr As Long) As Long
Dim i As Long
   For i = 1 To m_lKeyHookCount
      If m_lKeyHookPtr(i) = lPtr Then
         GetKeyHookPtrIndex = i
         Exit For
      End If
   Next i
End Function
Public Sub DetachKeyboardHook(cN As cNCCalcSize)
Dim lPtr As Long
Dim i As Long
Dim lIdx As Long
      
   lPtr = ObjPtr(cN)
   lIdx = GetKeyHookPtrIndex(lPtr)
   
   If lIdx > 0 Then
      If m_lKeyHookCount > 1 Then
         For i = lIdx To m_lKeyHookCount - 1
            m_lKeyHookPtr(i) = m_lKeyHookPtr(i + 1)
         Next i
         m_lKeyHookCount = m_lKeyHookCount - 1
         ReDim Preserve m_lKeyHookPtr(1 To m_lKeyHookCount) As Long
      Else
         m_lKeyHookCount = 0
         Erase m_lKeyHookPtr
      End If
   End If
   
   If m_lKeyHookCount <= 0 Then
      If (m_hKeyHook <> 0) Then
         UnhookWindowsHookEx m_hKeyHook
         m_hKeyHook = 0
      End If
   End If
   
End Sub
Private Function GetActiveConsumer(ByRef cM As cNCCalcSize) As Boolean
Dim i As Long
   For i = 1 To m_lKeyHookCount
      If Not m_lKeyHookPtr(i) = 0 Then
         Set cM = ObjectFromPtr(m_lKeyHookPtr(i))
         If cM.WindowActive Then
            GetActiveConsumer = True
            Exit Function
         End If
      End If
   Next i
End Function
Private Function KeyboardFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim bKeyUp As Boolean
Dim bAlt As Boolean, bCtrl As Boolean, bShift As Boolean
Dim bFKey As Boolean, bEscape As Boolean, bDelete As Boolean
Dim wMask As KeyCodeConstants
Dim i As Long
Dim lPtr As Long
Dim cM As cNCCalcSize

On Error GoTo ErrorHandler

   If nCode = HC_ACTION And m_hKeyHook > 0 Then
      ' Key up or down:
      bAlt = ((lParam And &H20000000) = &H20000000)
      If bAlt And (wParam > 0) And (wParam <> vbKeyMenu) Then
         bKeyUp = ((lParam And &H80000000) = &H80000000)
         If Not bKeyUp Then
            bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
            bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
            bFKey = ((wParam >= vbKeyF1) And (wParam <= vbKeyF12))
            bEscape = (wParam = vbKeyEscape)
            bDelete = (wParam = vbKeyDelete)
            If Not (bCtrl Or bFKey Or bEscape Or bDelete) Then
               If GetActiveConsumer(cM) Then
                  If cM.AltKeyAccelerator(wParam) Then
                     ' Don't pass accelerator on...
                     KeyboardFilter = 1
                     Exit Function
                  End If
               End If
            End If
         End If
      End If
   End If
   KeyboardFilter = CallNextHookEx(m_hKeyHook, nCode, wParam, lParam)

   Exit Function
   
ErrorHandler:
   Debug.Print "Keyboard Hook Error!"
   Exit Function
   Resume 0
End Function

Public Sub AttachMsgHook(cThis As cToolbarMenu)
Dim lpFn As Long
   DetachMsgHook
   m_lMsgHookPtr = ObjPtr(cThis)
   lpFn = HookAddress(AddressOf MenuInputFilter)
   m_hMsgHook = SetWindowsHookEx(WH_MSGFILTER, lpFn, 0&, GetCurrentThreadId())
   Debug.Assert (m_hMsgHook <> 0)
End Sub
Public Sub DetachMsgHook()
   If (m_hMsgHook <> 0) Then
      UnhookWindowsHookEx m_hMsgHook
      m_hMsgHook = 0
   End If
End Sub

'////////////////
'// Menu filter hook just passes to virtual CMenuBar function
'//
Private Function MenuInputFilter(ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim cM As cToolbarMenu
Dim lpMsg As Msg
   If nCode = MSGF_MENU Then
      If Not m_lMsgHookPtr = 0 Then
         Set cM = ObjectFromPtr(m_lMsgHookPtr)
         CopyMemory lpMsg, ByVal lParam, Len(lpMsg)
         If (cM.MenuInput(lpMsg)) Then
            MenuInputFilter = 1
            Exit Function
         End If
      End If
   End If
   MenuInputFilter = CallNextHookEx(m_hMsgHook, nCode, wParam, lParam)
End Function


Private Function HookAddress(ByVal lPtr As Long) As Long
   HookAddress = lPtr
End Function

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Property


