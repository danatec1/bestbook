VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   1  '���� ����
   Caption         =   "�ʰ��� ����"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5310
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdDivide 
      Caption         =   "������"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdMultiply 
      Caption         =   "���ϱ�"
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "����"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "���ϱ�"
      Height          =   495
      Left            =   3000
      TabIndex        =   6
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtResult 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtValue2 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtValue1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblResult 
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblValue2 
      Caption         =   "��2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblValue1 
      Caption         =   "��1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Plus = 1
Const Minus = 2
Const Multifle = 3
Const Divide = 4

Public Function Calculator(Value1 As Integer, Value2 As Integer, _
                                    Method As Integer) As Integer

    Select Case Method '��� ����� ���� Case���� ���
        Case Plus '���ϱ�
            Calculator = Value1 + Value2
        
        Case Minus '����
            Calculator = Value1 - Value2
            
        Case Multifle '���ϱ�
            Calculator = Value1 * Value2
            
        Case Divide '������
            Calculator = Value1 / Value2
            
    End Select
    
End Function

Private Sub cmdPlus_Click()
    Dim Result As Integer '������� �޴� ����
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Plus) '�Լ� ȣ��
    txtResult.Text = CStr(Result)
End Sub

Private Sub cmdMinus_Click()
    Dim Result As Integer '������� �޴� ����
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Minus) '�Լ� ȣ��
    txtResult.Text = CStr(Result)
End Sub

Private Sub cmdMultiply_Click()
    Dim Result As Integer '������� �޴� ����
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Multifle) '�Լ� ȣ��
    txtResult.Text = CStr(Result)
End Sub

Private Sub cmdDivide_Click()
    Dim Result As Integer '������� �޴� ����
    Result = Calculator(CInt(txtValue1.Text), CInt(txtValue2.Text), Divide) '�Լ� ȣ��
    txtResult.Text = CStr(Result)
End Sub



