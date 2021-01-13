VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtSecond 
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.TextBox txtFirst 
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtFirst_Change()
    Beep
    MsgBox "txtFirst 컨트롤에 Click 이벤트가 발생했습니다."
End Sub

Private Sub txtFirst_Click()
    Beep
End Sub

Private Sub txtFirst_GotFocus()
    MsgBox "txtFirst 컨트롤에 GotFocus 이벤트가 발생했습니다."
End Sub

Private Sub txtFirst_LostFocus()
    MsgBox "txtFirst 컨트롤에 LostFocus 이벤트가 발생했습니다."
End Sub

Private Sub txtSecond_Change()
    Beep
End Sub

Private Sub txtSecond_DblClick()
    Beep
End Sub
