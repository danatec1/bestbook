VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPost 
   BorderStyle     =   3  '크기 고정 대화 상자
   Caption         =   "우편번호"
   ClientHeight    =   2910
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5775
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫기"
      Height          =   300
      Left            =   4616
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "새로 고침"
      Height          =   300
      Left            =   3462
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "삭제"
      Height          =   300
      Left            =   2308
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "업데이트"
      Height          =   300
      Left            =   1154
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "추가"
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtAddr 
      DataField       =   "전체주소"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtNum 
      DataField       =   "우편번호"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtDong 
      DataField       =   "동명"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  '아래 맞춤
      Height          =   330
      Left            =   0
      Top             =   2580
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=우편번호"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "우편번호"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select 동명,우편번호,전체주소 from 우편번호"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "전체주소"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "우편번호"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '오른쪽 맞춤
      Caption         =   "동명"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  '오류 처리 코드를 넣는 위치입니다.
  '오류를 무시하려면 다음 줄을 주석으로 처리하십시오.
  '오류를 잡으려면 여기에 오류를 처리하는 코드를 추가하십시오.
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '이 레코드 집합의 현재 레코드 위치를 표시합니다.
  datPrimaryRS.Caption = "Record: " + Str(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub cmdAddNew_Click()
  '에러처리 구문
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  
  '에러처리 구문
  On Error GoTo DeleteErr
  
  '레코드 삭제하기
  atPrimaryRS.Recordset.Delete
  atPrimaryRS.Recordset.MoveNext
  '파일끝일 경우 마지막 레코드로 이동
  If atPrimaryRS.Recordset.EOF Then
    atPrimaryRS.MoveLast
  End If
  
  Exit Sub
DeleteErr:
  '에러 메시지 출력
  MsgBox Err.Description

End Sub

Private Sub cmdRefresh_Click()
  
  '에러처리 구문
  On Error GoTo RefreshErr
  
  '데이터 컨트롤을 새로 갱신
  datPrimaryRS.Refresh
  Exit Sub

RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  
  '에러처리 구문
  On Error GoTo UpdateErr

  '레코드 저장하기
  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
