VERSION 5.00
Begin VB.Form frmPostSerch 
   Caption         =   "우편번호 검색기 Ver 1.0"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdSearch 
      Caption         =   "찾아주세요"
      BeginProperty Font 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   2340
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   2625
      Begin VB.Label Label2 
         Caption         =   "동명에 찾고자 하는 동의 이름만 입력하고 '찾기'버튼을 누르세요.예) 동명='압구정'"
         Height          =   555
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.TextBox txtDong 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   1545
   End
   Begin VB.ListBox QueryResult 
      Height          =   2760
      Left            =   2820
      TabIndex        =   0
      Top             =   120
      Width           =   5025
   End
   Begin VB.Label Result 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '단일 고정
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1230
      TabIndex        =   7
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "우편번호"
      Height          =   285
      Left            =   210
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "동(리)명"
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   510
      Width           =   705
   End
End
Attribute VB_Name = "frmPostSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Rs As Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub cmdSearch_Click()

  Dim db As Connection
  Set db = New Connection
  '클라이언트 커서 사용
  db.CursorLocation = adUseClient
  '데이터베이스 연다
  db.Open "dsn=우편번호;uid=;pwd=;"

  Dim YesNO As Integer

  If txtDong.Text = "" Then
    YesNO = MsgBox("동명에 값이 입력되지 않았습니다. 전체를 검색하시겠습니까?", vbYesNo + vbQuestion, "값이 입력되지 않았습니다.")
    If YesNO = vbNo Then
      Exit Sub
    End If
  End If

  'SQL 문을 이용하여 쿼리를 연다
  Set Rs = New Recordset
  Rs.Open "select 우편번호,동명,전체주소 from 우편번호 where 동명 LIKE '" + txtDong.Text + "%'", db, adOpenKeyset, adLockOptimistic

  QueryResult.Clear
  
  '데이터베이스의 크기가 0이하이면 검색된 레코드 없슴
  If Rs.RecordCount > 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
      '리스트박스 컨트롤에 [우편번호] 전체 주소의 형태로 레코드를 추가
      QueryResult.AddItem "[" + Mid(Rs!우편번호, 2, 3) + "-" + Mid(Rs!우편번호, 4, 3) + "] " + Rs!전체주소
      Rs.MoveNext
    Loop
  Else
    MsgBox "검색된 레코드가 없습니다."
  End If
    
End Sub

Private Sub QueryResult_Click()
    Result.Caption = Mid(QueryResult.Text, 2, 3) + "-" + Mid(QueryResult.Text, 6, 3)
End Sub

Private Sub QueryResult_Scroll()
    QueryResult_Click
End Sub
