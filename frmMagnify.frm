VERSION 5.00
Begin VB.Form frmMagnify 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3465
   Icon            =   "frmMagnify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMagnify.frx":09CA
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   231
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1395
      Top             =   510
   End
   Begin VB.PictureBox picMag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   75
      ScaleHeight     =   228
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   212
      TabIndex        =   0
      Top             =   75
      Width           =   3210
   End
End
Attribute VB_Name = "frmMagnify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" _
(ByVal hwnd As Long) As Long

Private Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, _
ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function SetRectRgn Lib "gdi32" _
(ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, _
ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function GetCursorPos Lib "user32" _
(lpPoint As POINTAPI) As Long



Const SRCCOPY = &HCC0020
'Dim r As Long
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Dim MyPoint As POINTAPI
Dim X1 As Integer
Dim Y1 As Integer
Dim X2 As Integer
Dim Y2 As Integer

Dim tex As Long

Dim W1 As Long
Dim H1 As Long
        
        Dim hWndDesk As Long
        Dim hDCDesk As Long
'        Dim LeftDesk As Long
'        Dim TopDesk As Long
'        Dim WidthDesk As Long
'        Dim HeightDesk As Long
       
       
Dim InitP As POINTAPI
Dim FinP As POINTAPI


Private Sub Form_Load()
'frmMagnify.Picture = LoadPicture(App.Path & "\lense.bmp")
'frmMagnify.Icon = LoadPicture(App.Path & "\lense.ico")

SetTopMostWindow Me.hwnd, True
'**********************
SetWindowRgn hwnd, CreateEllipticRgn(0, 0, frmMagnify.ScaleWidth, frmMagnify.ScaleHeight), True

'******************
Call MakeMeTransparent

W1 = picMag.Width
H1 = picMag.Height


hWndDesk = GetDesktopWindow()
hDCDesk = GetWindowDC(hWndDesk)
End Sub

Private Sub Screen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
GetCursorPos InitP
X1 = InitP.X
Y1 = InitP.Y
X = X1
Y = Y1
End Sub


Private Sub Timer1_Timer()
GetCursorPos FinP
X2 = FinP.X
Y2 = FinP.Y

StretchBlt frmMagnify.picMag.hDC, 0, 0, W1, H1, hDCDesk, X2 - 20, Y2 - 20, 0.35 * W1, 0.35 * H1, SRCCOPY

frmMagnify.picMag.Refresh
frmMagnify.Left = X2 * Screen.TwipsPerPixelX
frmMagnify.Top = Y2 * Screen.TwipsPerPixelY + frmMagnify.ScaleHeight * 1.25
ReleaseCapture

Dim K As KeyCodeConstants
If K = vbKeyEscape Then
Timer1.Enabled = False
Unload frmMagnify
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Timer1.Enabled = False
Unload Me
End If

End Sub
Private Sub Screen_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
Timer1.Enabled = False
Unload Me
End If

End Sub

Private Sub Command1_Click()
        
       'setup the screen coordinates (upper corner (0,0) and lower
       '     corner (Width,Height)
       
       '     'copy the desktop to the picture box
       'r = BitBlt(frmMagnify.picBack.hDC, 0, 0, WidthDesk, HeightDesk, hDCDesk, LeftDesk, TopDesk, vbSrcCopy)

'tex = StretchBlt(frmMagnify.picMag.hdc, 0, 0, W1, H1, hDCDesk, X2, Y2, 100, 100, SRCCOPY)

End Sub

