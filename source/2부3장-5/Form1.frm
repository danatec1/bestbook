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
   Begin VB.CommandButton cmdTest2 
      Caption         =   "Test2"
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "Test1"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSecond 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtFirst 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdTest1_Click()
    MsgBox txtFirst.Text
End Sub

Private Sub cmdTest2_Click()
    mstbox txtSecond.Text
End Sub

Private Sub txtFirst_Change()

End Sub

Private Sub txtFirst_Click()
    MsgBox "텍스트박스 컨트롤에 Click 이벤트가 발생했습니다."
End Sub
