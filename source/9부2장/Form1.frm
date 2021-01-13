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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "삭제"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "추가"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "이동"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtMove 
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "끝으로"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "처음으로"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "다음"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "이전"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtSect 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "SECT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   1200
      Width           =   1600
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "DATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   1600
   End
   Begin VB.TextBox txtItem 
      DataField       =   "ITEM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   3585
   End
   Begin VB.TextBox txtID 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1600
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "AMOUNT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1600
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\vb60\가계부\가계부.mdb"
      DefaultCursorType=   0  '기본 커서
      DefaultType     =   2  'ODBC사용
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '다이너셋
      RecordSource    =   "가계부"
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "분류"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "비고"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "번호"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "날짜"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "금액"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNew_Click()
  
  '새로운 레코드를 추가한다
  Data1.Recordset.AddNew

End Sub

Private Sub cmdDelete_Click()
  
  '현재 레코드를 삭제합니다.
  Data1.Recordset.Delete
  '파일 끝이 아니면다음 레코드로 이동
  If Not Data1.Recordset.EOF Then
    Data1.Recordset.MoveNext
  '파일 끝일 경우 마지막 레코드로 이동
  Else
    Data1.Recordset.MoveLast
  End If
  
End Sub

Private Sub cmdMove_Click()

  '텍스트박스에 입력된 숫자만큼 이동
  Data1.Recordset.Move Val(txtMove.Text)
  
  '파일끝일 경우 각각 첫번째 레코드 혹은
  '마지막 레코드로 이동시킨다.
  If Data1.Recordset.BOF Then
    Data1.Recordset.MoveFirst
  ElseIf Data1.Recordset.EOF Then
    Data1.Recordset.MoveLast
  End If
  
End Sub

Private Sub cmdMoveFirst_Click()

  '첫 번째 레코드로 이동
  Data1.Recordset.MoveFirst
  
End Sub

Private Sub cmdMoveLast_Click()
  
  '마지막 레코드로 이동
  Data1.Recordset.MoveLast
  
End Sub

Private Sub cmdMoveNext_Click()
  
  '다음 레코드로 이동
  Data1.Recordset.MoveNext
  
  '파일의 맨 끝이면
  If Data1.Recordset.EOF Then
    '마지막 레코드로 이동
    Data1.Recordset.MoveLast
  End If
  
End Sub

Private Sub cmdMovePrevious_Click()
  
  '이전 레코드로 이동
  Data1.Recordset.MovePrevious

  If Data1.Recordset.BOF Then
    '첫 번째 레코드로 이동
    Data1.Recordset.MoveFirst
  End If
  
End Sub

Private Sub Data1_Reposition()

  '데이터 컨트롤에 레코드의 위치를 표시
  Data1.Caption = "현재 위치:" + Str(Data1.Recordset.AbsolutePosition)
  
End Sub

