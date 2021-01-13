VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Line 연습하기"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "이   동"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   1200
      X2              =   3360
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Line1.X1 = 2700
   Line1.Y1 = 3000
   Line1.X2 = Int(Form1.Width * Rnd)
   Line1.Y2 = Int(Form1.Height * Rnd)
   Line1.BorderColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
   Line1.BorderWidth = Int(50 * Rnd)
End Sub

