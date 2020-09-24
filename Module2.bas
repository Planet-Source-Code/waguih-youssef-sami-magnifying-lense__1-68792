Attribute VB_Name = "Module2"
Private Declare Function SetWindowRgn Lib "user32" _
(ByVal hwnd As Long, ByVal hRgn As Long, _
ByVal bRedraw As Boolean) As Long

Private Declare Function DeleteObject Lib "gdi32" _
(ByVal hObject As Long) As Long

Private Declare Function SelectObject Lib "gdi32" _
(ByVal hDC As Long, ByVal hObject As Long) As Long

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function CreateRectRgn Lib "gdi32" _
(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" _
(ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, _
ByVal nCombineMode As Long) As Long



Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" _
(ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" _
(ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private hRgn As Long

Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_PATHMUSTEXIST = &H800
Private Const CC_FULLOPEN = &H2
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_RGBINIT = &H1
Private Const CC_ANYCOLOR = &H100
Public Sub MakeMeTransparent()

If hRgn Then Call DeleteObject(hRgn)
'You can change the color that wants to be transparant by changing the rgb color above.
hRgn = GetBitmapRegion(frmMagnify.Picture, RGB(255, 0, 255))
Call SetWindowRgn(frmMagnify.hwnd, hRgn, True)
    
End Sub
Public Function GetBitmapRegion(cPicture As StdPicture, cTransparent As Long)

Dim hRgn As Long
Dim tRgn As Long
Dim X As Integer
Dim Y As Integer
Dim X0 As Integer
Dim hDC As Long
Dim BM As BITMAP

hDC = CreateCompatibleDC(0)
If hDC Then
    Call SelectObject(hDC, cPicture)
    Call GetObject(cPicture, Len(BM), BM)
    hRgn = CreateRectRgn(0, 0, BM.bmWidth, BM.bmHeight)
    For Y = 0 To BM.bmHeight
        For X = 0 To BM.bmWidth
            While X <= BM.bmWidth And GetPixel(hDC, X, Y) <> cTransparent
                X = X + 1
            Wend
            X0 = X
            While X <= BM.bmWidth And GetPixel(hDC, X, Y) = cTransparent
                X = X + 1
            Wend
            If X0 < X Then
                tRgn = CreateRectRgn(X0, Y, X, Y + 1)
                Call CombineRgn(hRgn, hRgn, tRgn, 4)
                Call DeleteObject(tRgn)
            End If
        Next X
    Next Y
    GetBitmapRegion = hRgn
    Call DeleteObject(SelectObject(hDC, cPicture))
End If
Call DeleteDC(hDC)

End Function




