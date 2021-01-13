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
   Begin VB.CommandButton cmdSingle 
      Caption         =   "Single로 변환"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton cmdDouble 
      Caption         =   "Double로 변환"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txtExpression 
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDouble, MySingle

Private Sub cmdDouble_Click()
    MyDouble = CDbl(txtExpression.Text)
    txtExpression.Text = MyDouble
End Sub
Private Sub cmdSingle_Click()
    MySingle = CSng(txtExpression.Text)
    txtExpression.Text = MySingle
End Sub


