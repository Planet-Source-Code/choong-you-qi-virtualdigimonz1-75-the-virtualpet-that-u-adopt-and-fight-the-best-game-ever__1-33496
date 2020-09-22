VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INCAreaModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get hwnd() As Long

End Property
Public Sub GetTopMarginHeight(cy As Long)

End Sub
Public Sub GetLeftMarginWidth(cx As Long)

End Sub
Public Sub GetRightMarginWidth(cx As Long)

End Sub
Public Sub GetBottomMarginHeight(cy As Long)

End Sub
Public Sub NCPaint(ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)

End Sub
Public Sub HitTest(ByVal x As Long, ByVal y As Long, ByRef eHitTest As ECNCHitTestConstants)

End Sub
Public Sub NCMouseDown(ByVal x As Long, ByVal y As Long, ByRef bHandled As Boolean, ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)

End Sub
Public Sub NCMouseUp(ByVal x As Long, ByVal y As Long, ByVal hDC As Long, ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)

End Sub
Public Sub InitMenuPopup(ByVal wParam As Long, ByVal lParam As Long)

End Sub
Public Sub ExitMenuLoop()

End Sub
Public Function AltKeyAccelerator(ByVal vKey As KeyCodeConstants) As Long

End Function

