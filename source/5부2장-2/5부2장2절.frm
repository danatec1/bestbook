VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmTestCoolbar 
   BorderStyle     =   1  '���� ����
   Caption         =   "��� �׽�Ʈ"
   ClientHeight    =   2580
   ClientLeft      =   2835
   ClientTop       =   2940
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6585
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "��   ��"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1588
      BandCount       =   4
      _CBWidth        =   6495
      _CBHeight       =   900
      _Version        =   "6.0.8169"
      Caption1        =   "Border ����"
      Child1          =   "cmdChgBdr"
      MinHeight1      =   405
      Width1          =   975
      NewRow1         =   0   'False
      Caption2        =   "Enabled ����"
      Child2          =   "cmdChgEnabled"
      MinHeight2      =   405
      Width2          =   1170
      NewRow2         =   0   'False
      Caption3        =   "Visible ����"
      Child3          =   "cmdChgVisible"
      MinHeight3      =   405
      Width3          =   1815
      NewRow3         =   -1  'True
      Caption4        =   "Caption ����"
      Child4          =   "txtChgCmdCaption"
      MinHeight4      =   405
      Width4          =   1335
      NewRow4         =   0   'False
      Begin VB.TextBox txtChgCmdCaption 
         Height          =   405
         Left            =   3135
         TabIndex        =   5
         Top             =   465
         Width           =   3270
      End
      Begin VB.CommandButton cmdChgVisible 
         Caption         =   "Visible �Ӽ� �ٲٱ�"
         Height          =   405
         Left            =   1215
         TabIndex        =   3
         Top             =   465
         Width           =   570
      End
      Begin VB.CommandButton cmdChgEnabled 
         Caption         =   "�ؽ�Ʈ�ڽ� Enabled �ٲٱ�"
         Height          =   405
         Left            =   6465
         TabIndex        =   2
         Top             =   30
         Width           =   90
      End
      Begin VB.CommandButton cmdChgBdr 
         Caption         =   "�ؽ�Ʈ�ڽ� BorderStyle �ٲٱ�"
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   30
         Width           =   3870
      End
   End
End
Attribute VB_Name = "frmTestCoolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChgBdr_Click()
    If txtResult.BorderStyle = 1 Then '��� �ؽ�Ʈ�ڽ� ���� ���� ���̸�
        txtResult.BorderStyle = 0 '�������� ����
    Else '�����̸�
        txtResult.BorderStyle = 1 '���ϰ������� ����
    End If
End Sub

Private Sub cmdChgEnabled_Click()
    
    If txtResult.Enabled = True Then '�ؽ�Ʈ�ڽ��� Enabled�̸�
        txtResult.Enabled = False 'False��
    Else 'False�̸�
        txtResult.Enabled = True 'False��
    End If
    
    If cmdResult.Enabled = True Then 'Ŀ�ǵ��ư�� Enabled�̸�
        cmdResult.Enabled = False 'False��
    Else 'False�̸�
        cmdResult.Enabled = True 'True��
    End If
    
End Sub

Private Sub cmdChgVisible_Click()

    If txtResult.Visible = True Then  '�ؽ�Ʈ�ڽ��� Visible�̸�
        txtResult.Visible = False 'False��
    Else 'False�̸�
        txtResult.Visible = True 'False��
    End If
    
    If cmdResult.Visible = True Then 'Ŀ�ǵ��ư�� Visible�̸�
        cmdResult.Visible = False 'False��
    Else 'False�̸�
        cmdResult.Visible = True 'True��
    End If

End Sub

Private Sub cmdResult_Click()
    Dim i As Integer '�޼��� �ڽ��� ���ϰ��� ���� ����
    i = MsgBox("���� �����Ͻðڽ��ϱ�?", vbInformation Or vbOKCancel, "���� Ȯ��")
    If i = vbOK Then 'OK��ư�� ������
        End '����
    End If
End Sub

Private Sub coolTest_Resize()
    '��� ��Ʈ���� ũ�Ⱑ ����� ��
    frmTestCoolbar.Height = coolTest.Height + 1100
                    '���� ũ�⸦ ��ٺ��� 1100ũ��
    '��� �ؽ�Ʈ�ڽ��� Ŀ�ǵ��ư�� ��ġ ����
    txtResult.Top = coolTest.Top + coolTest.Height + 200
    cmdResult.Top = coolTest.Top + coolTest.Height + 200
End Sub

Private Sub txtChgCmdCaption_Change()
    cmdResult.Caption = txtChgCmdCaption.Text
        '�ؽ�Ʈ�ڽ��� ���ڿ��� Ŀ�ǵ�ڽ��� Caption����
End Sub

