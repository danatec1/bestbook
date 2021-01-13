VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3855
   ClientTop       =   2850
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Label Label1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ShiftTest As Integer
   ShiftTest = Shift And 7
   Select Case ShiftTest
   Case 1
       Label1.Caption = "Shift 키를 눌렀습니다."
   Case 2
       Label1.Caption = "Ctrl 키를 눌렀습니다."
   Case 3
       Label1.Caption = "Shift 와 Ctrl 키를 눌렀습니다."
   Case 4
       Label1.Caption = "Alt 키를 눌렀습니다."
   Case 5
       Label1.Caption = "Shift 와 Alt 키를 눌렀습니다."
   Case 6
       Label1.Caption = "Ctrl과 Alt 키를 눌렀습니다."
   Case 7
       Label1.Caption = "Shift, Alt, 그리고 Ctrl 키를 눌렀습니다."
   End Select
End Sub


