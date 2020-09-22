Attribute VB_Name = "Module1"
Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
 
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

 
Const WINDING = 2
Const ALTERNATE = 1
Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Type POINTAPI
  X As Long
  Y As Long
End Type

Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Global Const SRCCOPY = &HCC0020
Global Const SRCERASE = &H440328
Global Const SRCINVERT = &H660046
Global Const SRCAND = &H8800C6

Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long


Private hRegion As Long
Sub Main()

If App.PrevInstance Then End

' use the following if this is to be used a a screen saver.

'If Left(Command$, 2) = "/s" Then
'    Form1.Show
'ElseIf Left$(Command$, 2) = "/c" Then
'    Form2.Show
'End If


End Sub
Function PrintScreen()
       Dim r As Long
       Dim hWndDesk As Long
       Dim hDCDesk As Long
       Dim LeftDesk As Long
       Dim TopDesk As Long
       Dim WidthDesk As Long
       Dim HeightDesk As Long
       'setup the screen coordinates (upper corner (0,0) and lower
       '     corner (Width,Height)
       LeftDesk = 0
       TopDesk = 0
       WidthDesk = Screen.Width \ Screen.TwipsPerPixelX
       HeightDesk = Screen.Height \ Screen.TwipsPerPixelY
       '     'get the desktop handle and display context
       hWndDesk = GetDesktopWindow()
       hDCDesk = GetWindowDC(hWndDesk)
       '     'copy the desktop to the picture box
       r = BitBlt(Form1.picBack.hdc, 0, 0, WidthDesk, HeightDesk, hDCDesk, LeftDesk, TopDesk, vbSrcCopy)
       r = ReleaseDC(hWndDesk, hDCDesk)
End Function

