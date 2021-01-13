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
   Begin VB.Label Label1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim ButtonTest As Integer
     ButtonTest = Button And 7
     Select Case ButtonTest
Case 1
  Label1.Caption = "왼쪽 버튼을 눌렀습니다."
Case 2
  Label1.Caption = "오른쪽 버튼을 눌렀습니다."
Case 3
  Label1.Caption = "왼쪽과 오른쪽 버튼을 눌렀습니다."
Case 4
  Label1.Caption = "가운데 버튼을 눌렀습니다."
Case 5
  Label1.Caption = "왼쪽과 가운데 버튼을 눌렀습니다."
Case 6
  Label1.Caption = "오른쪽과 가운데 버튼을 눌렀습니다."
Case 7
  Label1.Caption = "모든 버튼을 다 눌렀습니다."
    End Select
End Sub

