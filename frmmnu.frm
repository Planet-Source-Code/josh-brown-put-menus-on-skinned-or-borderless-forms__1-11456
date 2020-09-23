VERSION 5.00
Begin VB.Form frmmnu 
   Caption         =   "Form2"
   ClientHeight    =   -75
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   ScaleHeight     =   -75
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnubutton 
      Caption         =   "Button"
      Begin VB.Menu mnublue 
         Caption         =   "Blue Caption"
      End
      Begin VB.Menu mnured 
         Caption         =   "Red Caption"
      End
      Begin VB.Menu mnugreen 
         Caption         =   "Green Caption"
      End
   End
End
Attribute VB_Name = "frmmnu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnublue_Click()
  Form1.Label2.ForeColor = vbBlue
End Sub

Private Sub mnuexit_Click()
  End
End Sub

Private Sub mnugreen_Click()
  Form1.Label2.ForeColor = vbGreen
End Sub

Private Sub mnured_Click()
  Form1.Label2.ForeColor = vbRed
End Sub
