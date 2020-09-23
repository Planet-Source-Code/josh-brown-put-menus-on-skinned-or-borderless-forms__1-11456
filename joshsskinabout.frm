VERSION 5.00
Begin VB.Form frmabout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4320
      Top             =   2040
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000011&
      Height          =   3135
      Left            =   240
      ScaleHeight     =   3075
      ScaleWidth      =   4035
      TabIndex        =   0
      Top             =   240
      Width           =   4095
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   0
         ScaleHeight     =   3135
         ScaleWidth      =   4095
         TabIndex        =   1
         Top             =   960
         Width           =   4095
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Josh's Skins v 1.1"
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   6
            Top             =   600
            Width           =   2535
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Enjoy the program"
            BeginProperty Font 
               Name            =   "One Stroke Script LET"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   5
            Top             =   2280
            Width           =   2175
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Bbal1223@aol.com"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1920
            TabIndex        =   4
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "Email feedback to: "
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFF00&
            BackStyle       =   0  'Transparent
            Caption         =   "This program was made by Josh Brown"
            Height          =   255
            Left            =   480
            TabIndex        =   2
            Top             =   1560
            Width           =   2895
         End
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "About "
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
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   4320
      Picture         =   "joshsskinabout.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image7 
      Height          =   255
      Left            =   0
      Picture         =   "joshsskinabout.frx":3D04
      Stretch         =   -1  'True
      Top             =   1800
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   0
      Picture         =   "joshsskinabout.frx":7A08
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image Image4 
      Height          =   3975
      Left            =   0
      Picture         =   "joshsskinabout.frx":B70C
      Stretch         =   -1  'True
      Top             =   -600
      Width           =   345
   End
   Begin VB.Image Image3 
      Height          =   660
      Left            =   0
      Picture         =   "joshsskinabout.frx":EF450
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   5100
   End
   Begin VB.Image Image2 
      Height          =   3135
      Left            =   4320
      Picture         =   "joshsskinabout.frx":13B394
      Stretch         =   -1  'True
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   0
      Picture         =   "joshsskinabout.frx":21F0D8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
  Left = (Screen.Width - Width) \ 2
  Top = (Screen.Height - Height) \ 2
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    Call DragForm(Me)
  End If
End Sub


Private Sub Image5_Click()
  Unload Me
End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Image5.Picture = Image7.Picture
End Sub


Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = Image6.Picture
End Sub



Private Sub Label2_Click()
  Timer1.Enabled = True
  Label2.Visible = False
End Sub
  


Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    Call DragForm(Me)
  End If
End Sub

Private Sub Timer1_Timer()
  Picture2.Top = Picture2.Top - 15
End Sub
