Attribute VB_Name = "Module1"
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Public Declare Function SetWindowPos Lib "user32" _
      (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
      ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, _
      ByVal wFlags As Long) As Long


Public Declare Function CreateEllipticRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, _
ByVal Y2 As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" _
(ByVal hWnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long



Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
   If (Topmost) Then
      SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function


