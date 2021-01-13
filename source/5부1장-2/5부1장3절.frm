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
   Begin VB.Menu mnuPopup 
      Caption         =   "팝업메뉴"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupMsg 
         Caption         =   "메시지 보이기"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "팝업메뉴 종료"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbRightButton) = vbRightButton Then
    '버튼이 눌리고 그것이 오른쪽 버튼이라면 아래의 코드를 실행합니다
        Me.PopupMenu mnuPopup, , X, Y
    '실행 폼(Me)에서 마우스의 클릭 위치에 팝업 메뉴를 호출합니다
    End If
End Sub

Private Sub mnuPopupExit_Click()
    End
    '프로그램을 종료합니다
End Sub

Private Sub mnuPopupMsg_Click()
    MsgBox "팝업 메뉴를 호출하셨습니다"
End Sub

