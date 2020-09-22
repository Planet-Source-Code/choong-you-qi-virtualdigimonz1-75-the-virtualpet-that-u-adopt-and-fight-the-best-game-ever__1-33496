Attribute VB_Name = "Module5"
Option Explicit
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) _
   As Long

   If Topmost = True Then 'Make the window topmost
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function

Public Function AutoDetectTop(f As Form)
Dim lR As Long
Select Case GetSetting("Digimon", "Digimon", "Setting4")
Case "1"
lR = SetTopMostWindow(f.hWnd, True)
Case "0"
lR = SetTopMostWindow(f.hWnd, False)
End Select
f.Icon = Form7.Icon
End Function
