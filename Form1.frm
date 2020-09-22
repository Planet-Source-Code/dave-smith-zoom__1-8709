VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11205
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   371
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   747
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMag 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   600
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   245
      TabIndex        =   2
      Top             =   1500
      Width           =   3675
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   420
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   300
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   1
      Top             =   300
      Width           =   2055
   End
   Begin VB.PictureBox picCopy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   7020
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   141
      TabIndex        =   0
      Top             =   780
      Width           =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X, Y, newx, newy, hig, wid, incN, mm As Integer
Private Sub Form_Activate()

Form2.Hide

Me.WindowState = 2

picBack.Left = 0
picBack.Top = 0
picBack.Width = Screen.Width \ Screen.TwipsPerPixelX
picBack.Height = Screen.Height \ Screen.TwipsPerPixelY

picMag.Height = Form2.HScroll2.Value
picMag.Width = Form2.HScroll2.Value

PrintScreen

newx = 100
newy = 100

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Then
    Unload Form1
    Form2.Show
End If

End Sub

Private Sub Form_Load()

picMag.Height = Form2.HScroll2.Value
picMag.Width = Form2.HScroll2.Value

If Form2.chkOval.Value = 1 Then
    picMag.ScaleMode = vbPixels
    hRegion = CreateEllipticRgn(0, 0, picMag.Width, picMag.Height)
    X = SetWindowRgn(picMag.hwnd, hRegion, True)
End If

wid = picMag.Width
hig = picMag.Height

incN = 2

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'tex = StretchBlt(Form1.picMag.hdc, 0, 0, wid, hig, Form1.picBack.hdc, X, Y, 100, 100, SRCCOPY)

End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

mm = mm + 1
Debug.Print mm

'If mm = 10 Then End
ReleaseCapture
tex = StretchBlt(Form1.picMag.hdc, 0, 0, wid, hig, Form1.picBack.hdc, X, Y, 100, 100, SRCCOPY)
picMag.Refresh
'newx = X
'newy = Y

End Sub

Private Sub picMag_KeyPress(KeyAscii As Integer)

If KeyAscii <> 27 Then End

End Sub


Private Sub picMag_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

ReleaseCapture
SendMessage picMag.hwnd, &HA1, 2, 0&

tex = StretchBlt(Form1.picMag.hdc, 0, 0, wid, hig, Form1.picBack.hdc, X, Y, 100, 100, SRCCOPY)

picMag.Refresh

End Sub

Private Sub picMag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ReleaseCapture
SendMessage picMag.hwnd, &HA1, 2, 0&

tex = StretchBlt(Form1.picMag.hdc, 0, 0, wid, hig, Form1.picBack.hdc, X, Y, 100, 100, SRCCOPY)

picMag.Refresh

End Sub

Private Sub picMag_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

tex = StretchBlt(Form1.picMag.hdc, 0, 0, wid, hig, Form1.picBack.hdc, X, Y, 100, 100, SRCCOPY)

End Sub

Private Sub Timer1_Timer()

incN = incN + Form2.HScroll1.Value / 1000   '0.005

newx = Form1.ScaleHeight / 2 + Sin(incN * 1.23) * (Form1.ScaleWidth / 2) '- (picMag.Width * 1.5)
newy = Form1.ScaleWidth / 2 + Cos(incN * 1.75) * (Form1.ScaleHeight / 2) - (picMag.Height * 1.5)

'newx = 100
'newy = 100

' Blit the area from picBack (Background Picturebox) to the picMag.hdc and stretch.
tex = StretchBlt(Form1.picMag.hdc, 0, 0, wid, hig, Form1.picBack.hdc, newx + (wid / 5), newy + (hig / 5), 100, 100, SRCCOPY)

picMag.Refresh
        
Y = newy
X = newx

picMag.Top = newy
picMag.Left = newx

End Sub
