VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "계산기"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4635
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton cmdEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLS"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdPoint 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   16
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdSign 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdDivision 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtScreen 
      Alignment       =   1  '오른쪽 맞춤
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'이 프로그램은 간단한 산술연산을 할 수 있는계산기프로그램입니다.

'Flag는 계속 숫자를 입력받을 것인지 아닌지를 판단하는 변수입니다.
'Continue는 "="로 마지막 계산을 하는 것인지 산술연산자로 계속 연산을
          '할 것인지를 판단하는 변수입니다.
'Number는 프로그램 실행 시에 텍스트박스에서 입력받은 숫자들을 저장
          '하는 변수입니다.
'Ispoint는 초기값이 '0.'이므로 이것이 숫자로 입력한 '0.'인지 여부를 판단
          '하는 변수입니다
          
'연산자가 입력되었는지의 상태정보를 저장하는 변수
Dim Flag As Boolean
'숫자값에 대한 연산이 제대로 이루어졌는지에 대한 상태정보를 저장하는 변수
Dim Continue As Boolean
'사용자가 .를 입력했는지의 상태정보를 저장하는 변수
Dim Ispoint As Boolean
'숫자값을 입력받는 것이 가능한지에 대한 상태정보를 저장하는 변수
Dim opStatus As Boolean
'첫번째 숫자값에 대한 문자열을 저장하는 변수
Dim First As String
'둘째 숫자값에 대한 문자열을 저장하는 변수
Dim Second As String
'매번 사용자가 누른 숫자를 처리하기위한 변수
Dim Number As String
'연산자를 저장하는 변수
Dim Opt As String
'음수의 처리를 위한 변수
Dim Sign As String
'이전값을 저장하기 위한 변수
Dim Previous As String
'결과값 저장을위한 변수
Dim Result As String

Private Sub cmdArray_Click(Index As Integer)
    
    '여러숫자 버튼의 처리를 간단히 하기위해서 컨트롤배열을 사용했습니다.
    '숫자버튼을 누르면 그 버튼의 인덱스숫자 값을 문자로 변환시켜서
    '사용자가 누른 버튼의 값을 구해옵니다.
    Number = CStr(cmdArray(Index).Index)
    
    '산술 연산자를 누른 후 Flag는 True로 바뀌고 다음 계산될 숫자를
    '입력받는 것입니다
    If Flag = True Then
    
        txtScreen.Text = ""
        txtScreen.Text = Number
        
    Else
        
        If txtScreen.Text = "0." Then
        
            If (Ispoint) Then
                '초기값으로 "0."이 아니라 사용자가 입력한 '0.'인 경우의 처리입니다.
                txtScreen.Text = txtScreen.Text + Number
            Else
                '초기값으로 "0."인 경우의 처리입니다.
                txtScreen.Text = Number
            End If
            
        Else
            If (CStr(txtScreen.Text) = "0" And Number = "0") Then
                '0000...이 입력되는 경우 에러메세지 발생시킵니다.
                MsgBox "소수점이나 연산자를 입력하세요"
            Else
                '숫자버튼을 누른 후 텍스트박스의 내용은 문자열로 바뀌어
                '있으므로 그 뒤에 또 누른 숫자를 이미 저장된 문자열 뒤에 붙입니다
                txtScreen.Text = txtScreen.Text + Number
            End If
        End If
    
    End If
    
    Flag = False
    
    '계속 숫자를 입력받을 준비를 합니다
    opStatus = True
    First = txtScreen.Text
    
End Sub

Private Sub cmdClear_Click()
    '텍스트박스의 내용을 초기화시킵니다
    Flag = False
    txtScreen.Text = "0."
    
    '계속 산술연산자를 이용한 계산이 아니라는 뜻입니다
    Continue = False
    opStatus = False
    
    '"0."이 초기값에 의한 것임을 알립니다
    Ispoint = False
End Sub

Private Sub cmdDivision_Click()
    
    If opStatus = True Then
        Opt = "/"
        '산술 연산자를 눌렀으므로 다음 숫자를 받을 준비를 합니다
        Flag = True
        '계속 이전에 선택한 산술연산자가 있다면 계속 산술연산자를 이용한
        '연산이므로 Preoperation()를 이용해서 이전의 산술연산자에 의한
        '계산을 먼저합니다
        If (Continue) Then
            Call Preoperation
        End If
    
        '계산된 결과를 Second에 저장하고 다음에 계산할 숫자는 First에 저장하게 됩니다
        Second = txtScreen.Text
    
        '다음에 또 산술연산자에 의한 연산이 계속될 것이므로 지금 선택된
        '산술연산자를 이전 연산자로 저장합니다
        Continue = True
        Previous = "/"
    Else
        '연산자가 두번 이상 눌렸을 때
        MsgBox "숫자를 입력하세요"
    End If
    
    opStatus = False
    
End Sub

Private Sub cmdEqual_Click()
    
    If opStatus = True Then
        If Opt = "+" Then
            '문자로 저장되어 있는 수를 Double형태의 숫자로 바꾸어 연산합니다
            txtScreen.Text = CDbl(Second) + CDbl(First)
            ElseIf Opt = "-" Then
                txtScreen.Text = CDbl(Second) - CDbl(First)
            ElseIf Opt = "*" Then
                txtScreen.Text = CDbl(Second) * CDbl(First)
            ElseIf Opt = "/" Then
                If First = "0" Then
                '0으로 나누는 경우에 에러메세지를 띄웁니다
                    Beep
                    MsgBox "0으로 나눌 수 없습니다.", vbOKOnly, "Error"
                    txtScreen.Text = "0."
                Else
                    txtScreen.Text = CDbl(Second) / CDbl(First)
                End If
            End If
    
            '계산 결과값이 -1과 1사이의 값이면 텍스트박스 화면에 "0"을 붙여줍니다
            '왜냐하면 만약 계산결과값이 0.6이라면 텍스트박스화면에는 .6이라고
            '보이기 때문입니다
            If (CDbl(txtScreen.Text) > -1 And CDbl(txtScreen.Text) < 1) Then
                txtScreen.Text = "0" + txtScreen.Text
            End If

            '"="연산을 수행한 후이므로 더 이상 계속 산술연산자에 의한 연산을
            '하지 않습니다
            Flag = True
            Second = txtScreen.Text
            Continue = False
            Ispoint = False
    Else
        '연산자를 입력하고 다시 =연산자를 입력한 경우이거나
        '=연산자를 두번 이상 클릭한 경우
        If txtScreen.Text = "0." Then
            '수를 입력받기 위한 초기값이 '0.'일 때 =연산자를 클릭한 경우
            MsgBox "숫자를 입력하세요"
        Else
            MsgBox "'=' 연산 이후에는 새로운 계산을 시작하십시오"
        End If
    End If
    opStatus = False

End Sub

Private Sub cmdMinus_Click()

    If opStatus = True Then
        Opt = "-"
        Flag = True
        
        If (Continue) Then
            Call Preoperation
        End If
    
        Second = txtScreen.Text
    
        Continue = True
        Previous = "-"
    Else
        '연산자가 두 번 이상 눌렸을때
        MsgBox "숫자를 입력하세요"
    End If
    opStatus = False

End Sub

Private Sub cmdMulti_Click()

    If opStatus = True Then
        Opt = "*"
        Flag = True

        If (Continue) Then Call Preoperation
            
        Second = txtScreen.Text
    
        Continue = True
        Previous = "*"
    Else
        '연산자가 두 번 이상 눌렸을때
        MsgBox "숫자를 입력하세요"
    End If
    opStatus = False

End Sub

Private Sub cmdOff_Click()

    'OFF버튼을 눌렀을 때 계산기 프로그램을 종료합니다
    Result = MsgBox("계산기를 종료하시겠습니까?", _
                     vbOKCancel + vbQuestion, "종료")
    'Result에는 "계산기를 종료하시겠습니까?"에 대한 "확인"또는 "취소"의
    '결과값이 저장됩니다
    If Result = vbOK Then  '"확인"을 누를 경우에 프로그램을 종료합니다
        End
    End If
    '"취소"를 누른 경우에는 무시합니다
    
End Sub

Private Sub cmdPlus_Click()

    If opStatus = True Then
        Opt = "+"
        Flag = True

        If (Continue) Then Call Preoperation
        
        Second = txtScreen.Text
       
        Continue = True
        Previous = "+"
    Else
        '연산자가 두 번 이상 눌렸을때
        MsgBox "숫자를 입력하세요"
    End If
    opStatus = False

End Sub


Private Sub cmdPoint_Click()

    '소수점을 문자로 저장된 숫자에 추가합니다
    If opStatus = True Then
        If txtScreen.Text = "0." Or txtScreen.Text = "" Then
            txtScreen.Text = "0."
        Else
            txtScreen.Text = txtScreen.Text + "."
        End If
    
        '소수점을 누른 뒤에도 계속 숫자를 입력받을 준비를 합니다
        Flag = False
        
        First = txtScreen.Text
        ''0.'이 초기값에 의한 수가 아니라 사용자에 의해 입력받은 수임을
        '알립니다
        Ispoint = True
        
    Else
        '소수점이 두 번 이상 눌렸을때
        MsgBox "숫자를 먼저 입력하세요"
    End If
    opStatus = False

End Sub

Private Sub cmdSign_Click()

    '"-"부호가 선택된 경우 이를 문자로 저장된 숫자의 맨 앞에 추가합니다
    If Not (txtScreen.Text = "0.") Then
        If (CDbl(txtScreen.Text) > 0) Then
           txtScreen.Text = "-" + txtScreen.Text
            First = txtScreen
        Else
            MsgBox "연산자를 입력하세요" ''-'부호가 두 번 이상 붙을 때
        End If
    End If

End Sub

Private Sub Form_Initialize()

    '폼이 로드되었을 때는 아무런 입력을 받지 않았으므로 숫자를 입력받을 준
    '비를 하고 각 변수들을 초기화합니다.
    Flag = False
    txtScreen.Text = "0."
    Continue = False
    Ispoint = False
    opStatus = False

End Sub

Private Function Preoperation()

    '계속 산술연산자를 눌러 계산을 할 경우를 처리하는 함수입니다
    '먼저 연산자를 누를 때마다 저장된 산술연산자를 이용해서 연산을 한 후
    '그 결과값을 텍스트박스에 보여줍니다
    Select Case Previous
        Case "+"
            txtScreen.Text = CDbl(Second) + CDbl(First)
        Case "-"
            txtScreen.Text = CDbl(Second) - CDbl(First)
        Case "*"
            txtScreen.Text = CDbl(Second) * CDbl(First)
        Case "/"
            If (First = "0") Then
                Beep
                MsgBox "0으로 나눌 수 없습니다.", vbOKOnly, "Error"
                txtScreen.Text = "0."
            Else
                txtScreen.Text = CDbl(Second) / CDbl(First)
            End If

    End Select
    
    '계산 결과값이 -1과 1사이의 값이면 텍스트박스 화면에 "0"을 붙여줍니다
    '왜냐하면 만약 계산결과값이 0.6이라면 텍스트박스화면에는 .6이라고
    '보이기 때문입니다
    If (CDbl(txtScreen.Text) > -1 And CDbl(txtScreen.Text) < 1) Then
       txtScreen.Text = "0" + txtScreen.Text
    End If
    
End Function
