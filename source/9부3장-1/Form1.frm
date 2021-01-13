VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DAO를 이용한 데이터베이스"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdCancel 
      Caption         =   "취소"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "편집"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "저장"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "AMOUNT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1600
   End
   Begin VB.TextBox txtID 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1600
   End
   Begin VB.TextBox txtItem 
      DataField       =   "ITEM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   3585
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "DATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1600
   End
   Begin VB.TextBox txtSect 
      Alignment       =   1  '오른쪽 맞춤
      DataField       =   "SECT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1600
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "이전"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "다음"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "처음으로"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "끝으로"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtMove 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "이동"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "추가"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "삭제"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "금액"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "날짜"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "번호"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "비고"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "분류"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDB As Database
Dim MySet As Recordset

Private Sub cmdCancel_Click()

  '저장을 취소
  MySet.CancelUpdate
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable
  
End Sub

Private Sub cmdEdit_Click()
  
  '레코드를 수정
  MySet.Edit
  '저장, 취소 버튼의 활성화
  SaveCancel_Enable
  
End Sub

Private Sub cmdSave_Click()
  
  '현재 레코드를 데이터베이스에 저장
  MySet.Update
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable
  
End Sub

Private Sub Form_Load()
    
    '데이터베이스 파일 열기
    Set MyDB = DBEngine.Workspaces(0).OpenDatabase("D:\vb60\가계부\가계부.MDB")
    '테이블 열기
    Set MySet = MyDB.OpenRecordset("가계부", dbOpenTable)
    
    '첫 번째 레코드로 이동
    MySet.MoveFirst
    '레코드를 화면에 표시
    ShowRecord

End Sub

Private Sub cmdAddNew_Click()
  
  '새로운 레코드를 추가한다
  MySet.AddNew
  '저장, 취소 버튼의 활성화
  SaveCancel_Enable

End Sub

Private Sub cmdDelete_Click()
  
  '현재 레코드를 삭제합니다.
  MySet.Delete
  '파일 끝이 아니면다음 레코드로 이동
  If Not MySet.EOF Then
    MySet.MoveNext
  '파일 끝일 경우 마지막 레코드로 이동
  Else
    MySet.MoveLast
  End If
  
End Sub

Private Sub cmdMove_Click()

  '텍스트박스에 입력된 숫자만큼 이동
  MySet.Move Val(txtMove.Text)
  
  '파일끝일 경우 각각 첫번째 레코드 혹은
  '마지막 레코드로 이동시킨다.
  If MySet.BOF Then
    MySet.MoveFirst
  ElseIf MySet.EOF Then
    MySet.MoveLast
  End If
  
  '화면에 현재 레코드를 표시
  ShowRecord
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable

End Sub

Private Sub cmdMoveFirst_Click()

  '첫 번째 레코드로 이동
  MySet.MoveFirst
  '화면에 현재 레코드를 표시
  ShowRecord
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable
  
End Sub

Private Sub cmdMoveLast_Click()
  
  '마지막 레코드로 이동
  MySet.MoveLast
  '화면에 현재 레코드를 표시
  ShowRecord
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable
  
End Sub

Private Sub cmdMoveNext_Click()
  
  '다음 레코드로 이동
  MySet.MoveNext
  
  '파일의 맨 끝이면
  If MySet.EOF Then
    '마지막 레코드로 이동
    MySet.MoveLast
  End If
  
  '화면에 현재 레코드를 표시
  ShowRecord
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable
  
End Sub

Private Sub cmdMovePrevious_Click()
  
  '이전 레코드로 이동
  MySet.MovePrevious

  If MySet.BOF Then
    '첫 번째 레코드로 이동
    MySet.MoveFirst
  End If
  
  '화면에 현재 레코드를 표시
  ShowRecord
  '저장, 취소 버튼의 비활성화
  SaveCancel_Disable

End Sub

Private Sub ShowRecord()

  'txtID.Text = MySet.Fields("ID")
  'txtDate.Text = MySet.Fields("DATE")
  'txtSect.Text = MySet.Fields("SECT")
  'txtAmount.Text = MySet.Fields("AMOUNT")
  'txtItem.Text = MySet.Fields("ITEM")

  txtID.Text = MySet!ID
  txtDate.Text = MySet!Date
  txtSect.Text = MySet!SECT
  txtAmount.Text = MySet!AMOUNT
  txtItem.Text = MySet!Item

End Sub

Private Sub SaveCancel_Enable()

  '저장, 취소 버튼을 활성화한다
  cmdSave.Enabled = True
  cmdCancel.Enabled = True

End Sub

Private Sub SaveCancel_Disable()

  '저장, 취소 버튼을 비활성화한다
  cmdSave.Enabled = False
  cmdCancel.Enabled = False

End Sub

