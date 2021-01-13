VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3645
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   240
      ScaleHeight     =   1755
      ScaleWidth      =   2955
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      Picture1.BackColor = &HC0C0C0
End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Picture1.BackColor = &HFFFFFF
End Sub

