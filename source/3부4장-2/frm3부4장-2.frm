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
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "������"
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
    Dim AplusB As Integer 'A�� B�� ���� ���� ���� ����
    Dim BplusC As Integer 'B�� C�� ���� ���� ���� ����
    
    AplusB = CInt(txtValueA.Text) + CInt(txtValueB.Text)
    MsgBox AplusB, , "A+B" 'A+B�� ����ϴ� �޼��� �ڽ�
    BplusC = CInt(txtValueB.Text) + CInt(txtValueC.Text)
    MsgBox BplusC, , "B+C" 'B+C�� ����ϴ� �޼��� �ڽ�
    
    If AplusB > BplusC Then 'A+B > B+C �� ���
        MsgBox " A + B �� B + C ���� ũ�ų� �����ϴ�.", vbInformation, "���"
    Else
        MsgBox " B + C �� A + B ���� Ů�ϴ�.", vbInformation, "���"
    End If
    
End Sub


