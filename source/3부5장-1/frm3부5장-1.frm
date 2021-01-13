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
   Begin VB.CommandButton cmdOperation2 
      Caption         =   "정수 더하기"
      Height          =   495
      Left            =   2520
      TabIndex        =   4
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdOperation 
      Caption         =   "문자열 더하기"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox txtSecond 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox txtFirst 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOperation_Click()
    txtResult = txtFirst.Text + txtSecond.Text
End Sub

Private Sub cmdOperation2_Click()
    txtResult = CInt(txtFirst.Text) + CInt(txtSecond.Text)
End Sub
