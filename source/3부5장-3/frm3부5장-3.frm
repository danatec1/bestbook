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
   Begin VB.CommandButton cmdLong 
      Caption         =   "Long으로 변환"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdInteger 
      Caption         =   "Integer로 변환"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtExpression 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyInteger, MyLong

Private Sub cmdInteger_Click()
    MyInteger = CInt(txtExpression.Text)
    txtExpression.Text = MyInteger
End Sub

Private Sub cmdLong_Click()
    MyLong = CLng(txtExpression.Text)
    txtExpression.Text = MyLong
End Sub


