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
   Begin VB.CommandButton cmdGetValue 
      Caption         =   "�������� ����"
      Height          =   525
      Left            =   2820
      TabIndex        =   1
      Top             =   1080
      Width           =   1635
   End
   Begin VB.TextBox txtStart 
      Height          =   525
      Left            =   2790
      TabIndex        =   0
      Top             =   240
      Width           =   1635
   End
   Begin VB.Label lblResult 
      BorderStyle     =   1  '���� ����
      Height          =   2655
      Left            =   210
      TabIndex        =   2
      Top             =   240
      Width           =   2325
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGetValue_Click()
    Dim StrData As String '������ ������� ���� ��Ʈ��
    Dim IntDataForCast As Integer '�ؽ�Ʈ�ڽ��� ���� ���ڷ� ��ȯ
    Dim i As Integer 'For ~ Next ������ ���� ���� ����
    
    lblResult.Caption = "" '���̺��� �ʱ�ȭ�Ѵ�
    For i = 1 To 9
        If txtStart.Text = "" Then txtStart.Text = "1"
                '�ؽ�Ʈ�ڽ��� �ƹ��͵� �Է����� �����ÿ� 1�� ����
        IntDataForCast = CInt(txtStart.Text) '�ؽ�Ʈ�ڽ� ���ڸ� ���ڷ� ��ȯ
        
    StrData = CStr(IntDataForCast) & " * " & _
                    CStr(i) & " = " & CStr(IntDataForCast * i) & _
                    Chr$(10) & Chr$(13)     '��갪�� ������ ����
        lblResult.Caption = lblResult.Caption & StrData '���̺�� ���
    Next i
End Sub



