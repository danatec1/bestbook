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
   Begin VB.CommandButton cmdGetValue 
      Caption         =   "구구단을 외자!"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblResult 
      BorderStyle     =   1  '단일 고정
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGetValue_Click()
    Dim StrDataForInputBox As String 'InputBox 함수로 받을 스트링
    Dim StrData As String '구구단 결과값을 가질 스트링
    Dim IntDataForCast As Integer 'InputBox 함수로 받은 값을 숫자로 변환
    Dim i As Integer 'For ~ Next 문에서 사용될 변수 선언
    
    lblResult.Caption = "" '레이블을 초기화한다
    StrDataForInputBox = InputBox("구구단을 출력할 단을 입력하세요.", _
                            "단수 입력", "2") '디폴트로 2를 입력
    If StrDataForInputBox = "" Then 'InputBox 에서 < Esc > 키를 눌렀을 경우
        MsgBox "구구단 출력을 취소합니다.", vbInformation, "출력 취소"
        Exit Sub '구구단 출력을 취소합니다
    End If
    
    For i = 1 To 9
        IntDataForCast = CInt(StrDataForInputBox) 'InputBox 문자를 숫자로 변환
        StrData = CStr(IntDataForCast) & " * " & _
                    CStr(i) & " = " & CStr(IntDataForCast * i) & _
                    Chr$(10) & Chr$(13)     '계산값을 변수에 저장
        lblResult.Caption = lblResult.Caption & StrData '레이블로 출력
    Next i
End Sub


