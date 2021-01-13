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
   Begin VB.CommandButton cmdText 
      Caption         =   "Text���� �׽�Ʈ"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBorderStyle 
      Caption         =   "BorderStyle �׽�Ʈ"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdVisible 
      Caption         =   "Visible �׽�Ʈ"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnabled 
      Caption         =   "Enabled �׽�Ʈ"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblResult 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const EnabledTest = 0
Const VisibleTest = 1
Const BorderStyleTest = 2
Const TextTest = 3

Public Sub ChangeStatus(Task As Integer) '�۾��� ������ �Ű������� ����
    
    Select Case Task
        Case EnabledTest 'Enabled �� ���¸� �ٲ۴�
            
            If lblResult.Enabled = False Then '���̺��� ���� ���¸� ����
                lblResult.Enabled = True
            Else
                lblResult.Enabled = False
            End If
            
            If txtResult.Enabled = False Then '�ؽ�Ʈ�ڽ��� ���� ���¸� ����
                txtResult.Enabled = True
            Else
                txtResult.Enabled = False
            End If
            
        Case VisibleTest 'Visible �� ���¸� �ٲ۴�
            
            If lblResult.Visible = True Then '���̺��� ���� ���¸� ����
                lblResult.Visible = False
            Else
                lblResult.Visible = True
            End If
            
            If txtResult.Visible = True Then '�ؽ�Ʈ�ڽ��� ���� ���¸� ����
                txtResult.Visible = False
            Else
                txtResult.Visible = True
            End If
        
        Case BorderStyleTest 'BorderStyle �� �ٲ۴�
        
            If lblResult.BorderStyle = 0 Then '���̺��� ���� ���¸� ����
                lblResult.BorderStyle = 1 '1-���� ����
            Else
                lblResult.BorderStyle = 0 '0-����
            End If
            
            If txtResult.BorderStyle = 0 Then '�ؽ�Ʈ�ڽ��� ���� ���¸� ����
                txtResult.BorderStyle = 1 '1-���� ����
            Else
                txtResult.BorderStyle = 0 '0-����
            End If
        
        Case TextTest '���̺��� ���� Caption��, �ؽ�Ʈ�ڽ��� ���� Text�� �ٲ۴�
            
            If lblResult.Caption = "�׽�Ʈ ���̺�" Then '���̺��� ���� ���¸� ����
                lblResult.Caption = ""
            Else
                lblResult.Caption = "�׽�Ʈ ���̺�"
            End If
            
            If txtResult.Text = "�׽�Ʈ ����Ʈ" Then '�ؽ�Ʈ�ڽ��� ���� ���¸� ����
                txtResult.Text = ""
            Else
                txtResult.Text = "�׽�Ʈ ����Ʈ"
            End If
            
    End Select
End Sub


Private Sub cmdEnabled_Click()
    ChangeStatus EnabledTest 'Enabled�Ӽ��� �ٲ۴�
End Sub

Private Sub cmdVisible_Click()
    Call ChangeStatus(VisibleTest)  'Visible�Ӽ��� �ٲ۴�
End Sub

Private Sub cmdBorderStyle_Click()
    ChangeStatus BorderStyleTest 'BorderStyle�� �ٲ۴�
End Sub

Private Sub cmdText_Click()
    Call ChangeStatus(TextTest)  'Text�� �ٲ۴�
End Sub


