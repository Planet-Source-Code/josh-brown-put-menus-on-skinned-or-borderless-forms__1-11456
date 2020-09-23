VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000011&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2775
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Button"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.Image Image10 
      Height          =   255
      Left            =   960
      Picture         =   "joshsskin.frx":0000
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Image9 
      Height          =   255
      Left            =   360
      Picture         =   "joshsskin.frx":120BA
      Stretch         =   -1  'True
      Top             =   2160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   4320
      TabIndex        =   1
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   4080
      Picture         =   "joshsskin.frx":24174
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Josh's Skins"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image Image7 
      Height          =   240
      Left            =   360
      Picture         =   "joshsskin.frx":3622E
      Stretch         =   -1  'True
      Top             =   1680
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   360
      Picture         =   "joshsskin.frx":39F32
      Stretch         =   -1  'True
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   360
      Picture         =   "joshsskin.frx":3DC36
      Stretch         =   -1  'True
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   360
      Picture         =   "joshsskin.frx":4193A
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   240
      Picture         =   "joshsskin.frx":4563E
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   2535
      Left            =   5160
      Picture         =   "joshsskin.frx":91582
      Stretch         =   -1  'True
      Top             =   240
      Width           =   495
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   0
      Picture         =   "joshsskin.frx":1752C6
      Stretch         =   -1  'True
      Top             =   240
      Width           =   255
   End
   Begin VB.Image imgclose 
      Height          =   240
      Left            =   5160
      Picture         =   "joshsskin.frx":25900A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Imgmini 
      Height          =   240
      Left            =   4920
      Picture         =   "joshsskin.frx":25CD0E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   240
   End
   Begin VB.Image imgheader 
      Height          =   285
      Left            =   0
      Picture         =   "joshsskin.frx":260A12
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5730
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Left = (Screen.Width - Width) \ 2
  Top = (Screen.Height - Height) \ 2
End Sub

Private Sub Image8_Click()
  frmabout.Show
  frmabout.Left = Me.Left + Me.Width
  frmabout.Top = Me.Top
End Sub

Private Sub Image8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image8.Picture = Image9.Picture
  Label2.ForeColor = vbBlue
End Sub

Private Sub Image8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image8.Picture = Image10.Picture
  Label2.ForeColor = vbWhite
  frmabout.Left = Me.Left + Me.Width
  frmabout.Top = Me.Top
End Sub

Private Sub imgclose_Click()
  Unload Me
End Sub

Private Sub imgclose_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgclose.Picture = Image4.Picture
End Sub

Private Sub imgclose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgclose.Picture = Image5.Picture
End Sub

Private Sub imgheader_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub Imgmini_Click()
  Me.WindowState = vbMinimized
End Sub

Private Sub Imgmini_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Imgmini.Picture = Image6.Picture
End Sub

Private Sub Imgmini_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Imgmini.Picture = Image7.Picture
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
End If
End Sub

Private Sub Label2_Click()
  frmabout.Show
  frmabout.Left = Me.Left + Me.Width
  frmabout.Top = Me.Top
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image8.Picture = Image9.Picture
  Label2.ForeColor = vbBlue
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image8.Picture = Image10.Picture
  Label2.ForeColor = vbWhite
End Sub

Private Sub Label3_Click()
  frmmnu.PopupMenu frmmnu.mnufile, 1
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.ForeColor = vbBlue
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label3.ForeColor = vbWhite
End Sub

Private Sub Label4_Click()
  frmabout.Show
  frmabout.Left = Me.Left + Me.Width
  frmabout.Top = Me.Top
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label4.ForeColor = vbBlue
End Sub

Private Sub Label4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label4.ForeColor = vbWhite
End Sub

Private Sub Label5_Click()
  frmmnu.PopupMenu frmmnu.mnubutton, 1
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ForeColor = vbBlue
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Label5.ForeColor = vbWhite
End Sub
