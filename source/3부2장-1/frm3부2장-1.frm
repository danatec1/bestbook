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
   Begin VB.CommandButton cmdGetValue2 
      Caption         =   "Get 10!(Long)"
      Height          =   525
      Left            =   2400
      TabIndex        =   6
      Top             =   2220
      Width           =   1635
   End
   Begin VB.TextBox txtResult2 
      Height          =   525
      Left            =   2580
      TabIndex        =   4
      Top             =   1410
      Width           =   1245
   End
   Begin VB.CommandButton cmdGetValue 
      Caption         =   "Get 10!(Integer)"
      Height          =   525
      Left            =   480
      TabIndex        =   3
      Top             =   2220
      Width           =   1605
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   210
      TabIndex        =   2
      Top             =   1140
      Width           =   4305
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "Result"
      Height          =   525
      Left            =   690
      TabIndex        =   1
      Top             =   390
      Width           =   1245
   End
   Begin VB.TextBox txtInput 
      Height          =   525
      Left            =   2550
      TabIndex        =   0
      Top             =   390
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Caption         =   "Result"
      Height          =   525
      Left            =   720
      TabIndex        =   5
      Top             =   1410
      Width           =   1245
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGetValue2_Click()
    Dim intResult As Long       'Integer타입을 Long타입으로 수정
    Dim intL As Integer
    
    intL = 1
    intResult = 1
    
    While intL < 10
        intResult = intResult * intL
        intL = intL + 1
    Wend
    
    txtResult2.Text = intResult

End Sub

Private Sub cmdResult_Click()
    Dim strInput As String
    Dim intInput As Integer
    
    strInput = txtInput.Text
    intInput = Val(strInput)
    
    If intInput > 0 Then
        MsgBox "양의 정수를 입력했습니다"
    End If
    
    If intInput < 0 Then
        MsgBox "음의 정수를 입력했습니다"
    End If
    
    If intInput = 0 Then
        MsgBox "0을 입력했습니다"
    End If
    
End Sub


Private Sub cmdGetValue_Click()
    Dim intResult As Integer    'Integer타입의 intResult변수 선언
    Dim intL As Integer 'Integer타입의 intL변수 선언
    
    intL = 1            '초기값 할당
    intResult = 1
    
    While intL <= 10        '10보다 작거나 같은 값일때만 반복문 수행
        intResult = intResult * intL    '1*2*3*4*5.... 가 수행되는 부분
        intL = intL + 1     'intL값을 1씩 증가시킨다
    Wend
    
    txtResult2.Text = intResult      '결과를 텍스트박스에 보여준다

End Sub



