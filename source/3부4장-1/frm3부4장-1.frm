VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdQuit 
      Caption         =   "종료"
      Height          =   660
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "프린트"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
    Dim Result As Integer '메세지 박스의 결과값을 받을 정수
    Result = MsgBox("프린트 작업을 하시겠습니까?", vbOKCancel Or _
                        vbInformation, "프린트 확인") '프린트 작업 진행 확인
    If Result = vbOK Then 'OK버튼을 눌렀을 경우
        MsgBox "프린트 작업을 성공적으로 수행하였습니다.", , "수행완료"
    End If
End Sub

Private Sub cmdQuit_Click()
    Dim Result As Integer '메세지 박스의 결과값을 받을 정수
    Result = MsgBox("프로그램을 종료하시겠습니까?", vbYesNo _
            Or vbCritical Or vbApplicationModal, "종료 확인")
            '프로그램 종료 확인
    If Result = vbYes Then 'Yes버튼을 눌렀을 경우
        End
    Else
        MsgBox "종료를 취소합니다.", , "취소"
    End If
End Sub



