VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "랜덤 억세스 파일"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4890
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command4 
      Caption         =   "종       료"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "레코드 추가"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "다음 레코드"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "이전 레코드"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "학 과 :"
      Height          =   180
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "학 번 :"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "성 명 :"
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Record
  Name As String * 20
  SeNum As String * 15
  Depart As String * 30
End Type
Dim SeNum As Record
Dim CurRec As Integer
Dim NumRec As Integer

Private Sub ShowRec()
  Form1.Caption = "데이터 레코드 번호" + Str(CurRec)
  Get #1, CurRec, SeNum
  Text1.Text = SeNum.Name
  Text2.Text = SeNum.SeNum
  Text3.Text = SeNum.Depart
End Sub

Private Sub SaveRec()
SeNum.Name = Text1.Text
SeNum.SeNum = Text2.Text
SeNum.Depart = Text3.Text
Put #1, CurRec, SeNum
End Sub

Private Sub Form_Load()
  CurRec = 1
  Open "SeNum.db" For Random As #1 Len = Len(SeNum)
  NumRec = LOF(1) / Len(SeNum)
  If NumRec = 0 Then
  NumRec = 1
  End If
  ShowRec
End Sub

Private Sub Command1_Click()
   SaveRec
   If CurRec > 1 Then
   CurRec = CurRec - 1
   Else
   MsgBox "첫 번째 레코드입니다."
   End If
   ShowRec
End Sub

Private Sub Command2_Click()
  SaveRec
  If CurRec < NumRec Then
  CurRec = CurRec + 1
  Else
  MsgBox "마지막 레코드입니다."
  End If
   ShowRec
End Sub

Private Sub Command3_Click()
 SaveRec
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
 NumRec = NumRec + 1
 CurRec = NumRec
 Text1.SetFocus
End Sub

Private Sub Command4_Click()
SaveRec
Close #1
End
End Sub

