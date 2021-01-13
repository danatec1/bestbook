VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command2 
      Caption         =   "삭    제"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "입    력"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2040
      ItemData        =   "4부2장2절.frx":0000
      Left            =   240
      List            =   "4부2장2절.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Dim Inse, Univ, Department As String
   Inse = InputBox("신입생 이름은 무엇입니까?", "신입생 이름")
   Univ = InputBox("신입생 학교는 어디입니까?", "학교")
   Department = InputBox("신입생 학과는 무엇입니까?", "학과")
      '리스트박스에 inputbox를 통해 입력된 사항들을 나타냅니다.
   List1.AddItem Inse & "  " & Univ & "  " & Department
End Sub

Private Sub Command2_Click()
    a = MsgBox("정말 삭제하시겠습니까?", vbOKCancel + vbQuestion, "경고")
    If a = vbOK Then
    '현재 선택되어진 항목을 삭제합니다.
     List1.RemoveItem (List1.ListIndex)
    End If
End Sub

Private Sub List1_Click()
    '선택된 항목을 레이블박스에 나타냅니다.
  Label1.Caption = List1.List(List1.ListIndex)
End Sub

