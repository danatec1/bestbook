VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6585
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.TextBox txtValue1 
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtValue2 
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "���ϱ�"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label lblEqual 
      Alignment       =   2  '��� ����
      Caption         =   "="
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblPlus 
      Alignment       =   2  '��� ����
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCalculate_Click()
        Dim IntValue As Integer '��갪�� �ޱ� ���� ����
        IntValue = Plus(CInt(txtValue1.Text), CInt(txtValue2.Text))
                        'Function�Լ� ����
        txtResult.Text = CStr(IntValue) '���� ���� ���
    End Sub
    
    Private Function Plus(Value1 As Integer, Value2 As Integer) As Integer
        Plus = Value1 + Value2 '�Ű������� ���� ���� ����
    End Function

