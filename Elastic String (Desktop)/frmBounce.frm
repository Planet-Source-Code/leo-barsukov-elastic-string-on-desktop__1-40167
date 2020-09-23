VERSION 5.00
Begin VB.Form frmBounce 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animated Mouse's Tail"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   Icon            =   "frmBounce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   343
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   636
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   7
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   7
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   6
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   5
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   4
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   3
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   2
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   1
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox img 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Index           =   0
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   3240
   End
   Begin VB.Image ImgBall 
      Height          =   165
      Index           =   0
      Left            =   2040
      Picture         =   "frmBounce.frx":1782
      Top             =   480
      Width           =   165
   End
End
Attribute VB_Name = "frmBounce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Shape As New clsShaped
Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Sub Command1_Click()
frmAbout.Show vbModal
End Sub
Private Sub Form_DblClick()
Dim i As Integer
For i = ImgBall.LBound + 1 To ImgBall.UBound
Unload ImgBall(i)
Next i
Unload Me
End
End Sub
Private Sub Form_Load()
Dim i As Integer
For i = ImgBall.UBound + 1 To 7
Load ImgBall(i)
ImgBall(i).Visible = True
ImgBall(i).Top = ImgBall(i - 1).Top + 11
Next i
ImgBall(0).Visible = False
Call InitVal
Call InitBall
Timer1.Interval = 20
Timer1.Enabled = True

For i = 1 To 7
SetParent img(i).hwnd, FindWindow(vbNullString, "Program Manager")
img(i) = ImgBall(0)
img(0) = ImgBall(0)
Shape.Window img(i).hwnd, LoadPicture(App.Path & "\bullet.bmp"), vbMagenta
Next

Me.Height = Screen.Height
Me.Width = Screen.Width

End Sub
Private Sub Timer1_Timer()
Dim Pos As POINTAPI
GetCursorPos Pos
MoveHandler Pos.X, Pos.Y
Animate
For i = 1 To 7

img(i).Left = ImgBall(i).Left * 15
img(i).Top = ImgBall(i).Top * 15
Next
End Sub
