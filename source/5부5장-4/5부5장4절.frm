VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3960
   ClientTop       =   2550
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "5ºÎ5Àå4Àý.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
       Source.Move X - Source.Width / 2, Y - Source.Height / 2
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    Source.Move Image1.Left + X - Source.Width / 2, Image1.Top + Y - Source.Height / 2
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image1.Drag 1
End Sub

