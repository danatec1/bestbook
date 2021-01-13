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
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "계산시작"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtValueC 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtValueB 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtValueA 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
    Dim AplusB As Integer 'A와 B를 더한 값을 받을 변수
    Dim BplusC As Integer 'B와 C를 더한 값을 받을 변수
    
    AplusB = CInt(txtValueA.Text) + CInt(txtValueB.Text)
    MsgBox AplusB, , "A+B" 'A+B를 출력하는 메세지 박스
    BplusC = CInt(txtValueB.Text) + CInt(txtValueC.Text)
    MsgBox BplusC, , "B+C" 'B+C를 출력하는 메세지 박스
    
    If AplusB > BplusC Then 'A+B > B+C 인 경우
        MsgBox " A + B 가 B + C 보다 크거나 같습니다.", vbInformation, "결과"
    Else
        MsgBox " B + C 가 A + B 보다 큽니다.", vbInformation, "결과"
    End If
    
End Sub


