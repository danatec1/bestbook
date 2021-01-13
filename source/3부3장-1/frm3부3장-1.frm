VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtValue1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtValue2 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "더하기"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblEqual 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "="
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblPlus 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCalculate_Click()
        Dim IntValue As Integer '계산값을 받기 위한 정수
        IntValue = Plus(CInt(txtValue1.Text), CInt(txtValue2.Text))
                        'Function함수 실행
        txtResult.Text = CStr(IntValue) '받은 값을 출력
    End Sub
    
    Private Function Plus(Value1 As Integer, Value2 As Integer) As Integer
        Plus = Value1 + Value2 '매개변수로 받은 값을 더함
    End Function

