VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "순차 억세스 파일"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command4 
      Caption         =   "화면 삭제"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "추 가 용"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "입 력 용"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "출 력 용 "
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim OutPt As String
  Open "text.txt" For Output As #1
  Print #1, Text1.Text
  Close #1
End Sub

Private Sub Command2_Click()
     Dim OutPt, FFile As String
     Open "text.txt" For Input As #1
     Do Until EOF(1)
         Input #1, OutPt
         FFile = FFile + OutPt + Chr(13) + Chr(10)
     Loop
     Text1.Text = FFile
     Close #1
End Sub

Private Sub Command4_Click()
     Text1.Text = ""
End Sub

Private Sub Command3_Click()
    Open "text.txt" For Append As #1
    Print #1, Text1.Text
    Close #1
End Sub

