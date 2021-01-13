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
      Caption         =   "구구단을 외자"
      Height          =   525
      Left            =   2820
      TabIndex        =   1
      Top             =   1080
      Width           =   1635
   End
   Begin VB.TextBox txtStart 
      Height          =   525
      Left            =   2790
      TabIndex        =   0
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label lblResult 
      BorderStyle     =   1  '단일 고정
      Height          =   2655
      Left            =   210
      TabIndex        =   2
      Top             =   240
      Width           =   2325
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGetValue_Click()
    Dim StrData As String '구구단 결과값을 가질 스트링
    Dim IntDataForCast As Integer '텍스트박스의 값을 숫자로 변환
    Dim i As Integer 'For ~ Next 문에서 사용될 변수 선언
    
    lblResult.Caption = "" '레이블을 초기화한다
    For i = 1 To 9
        If txtStart.Text = "" Then txtStart.Text = "1"
                '텍스트박스에 아무것도 입력하지 않을시에 1을 넣음
        IntDataForCast = CInt(txtStart.Text) '텍스트박스 문자를 숫자로 변환
        
    StrData = CStr(IntDataForCast) & " * " & _
                    CStr(i) & " = " & CStr(IntDataForCast * i) & _
                    Chr$(10) & Chr$(13)     '계산값을 변수에 저장
        lblResult.Caption = lblResult.Caption & StrData '레이블로 출력
    Next i
End Sub



