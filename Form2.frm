VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   ScaleHeight     =   2730
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Foatie Magnifier"
      Height          =   2595
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3615
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   435
         Left            =   180
         TabIndex        =   7
         Top             =   1980
         Width           =   1575
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "&Ok"
         Height          =   435
         Left            =   1860
         TabIndex        =   6
         Top             =   1980
         Width           =   1575
      End
      Begin VB.CheckBox chkOval 
         Caption         =   "Clip to Oval Area"
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   1260
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   255
         LargeChange     =   5
         Left            =   1080
         Max             =   270
         Min             =   100
         TabIndex        =   4
         Top             =   960
         Value           =   250
         Width           =   2355
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         LargeChange     =   2
         Left            =   1080
         Max             =   9
         Min             =   1
         TabIndex        =   2
         Top             =   480
         Value           =   7
         Width           =   2355
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         Height          =   315
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Speed:"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   480
         Width           =   675
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdStart_Click()

Me.Hide

Form1.Show

End Sub

