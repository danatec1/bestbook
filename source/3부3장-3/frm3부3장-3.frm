VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  '단일 고정
   Caption         =   "초간단 계산기"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdDivide 
      Caption         =   "나누기"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "곱하기"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "빼기"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "더하기"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtValue2 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtValue1 
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblResult 
      Caption         =   "결과"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblValue2 
      Caption         =   "값2"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblValue1 
      Caption         =   "값1"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Plus = 1
Const Minus = 2
Const Multifle = 3
Const Divide = 4

Public Function Calculator(Value1 As Integer, Value2 As Integer, _
                                    Method As Integer) As Integer

    Select Case Method '계산 방법에 따라 Case문을 사용
        Case Plus '더하기
            Calculator = Value1 + Value2
        
        Case Minus '빼기
            Calculator = Value1 - Value2
            
        Case Multifle '곱하기
            Calculator = Value1 * Value2
            
        Case Divide '나누기
            Calculator = Value1 / Value2
            
    End Select
    
End Function

Private Sub cmdPlus_Click()
    Dim Result As Integer '결과값을 받는 변수
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Plus) '함수 호출
    txtResult.Text = CStr(Result)
End Sub

Private Sub cmdMinus_Click()
    Dim Result As Integer '결과값을 받는 변수
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Minus) '함수 호출
    txtResult.Text = CStr(Result)
End Sub

Private Sub cmdMultiply_Click()
    Dim Result As Integer '결과값을 받는 변수
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Multifle) '함수 호출
    txtResult.Text = CStr(Result)
End Sub

Private Sub cmdDivide_Click()
    Dim Result As Integer '결과값을 받는 변수
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Divide) '함수 호출
    txtResult.Text = CStr(Result)
End Sub



