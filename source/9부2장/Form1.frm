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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "����"
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "�߰�"
      Height          =   375
      Left            =   2520
      TabIndex        =   16
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "�̵�"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtMove 
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "������"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "ó������"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "����"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "����"
      Height          =   375
      Left            =   2520
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtSect 
      Alignment       =   1  '������ ����
      DataField       =   "SECT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   9
      Top             =   1200
      Width           =   1600
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  '������ ����
      DataField       =   "DATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   1600
   End
   Begin VB.TextBox txtItem 
      DataField       =   "ITEM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   3585
   End
   Begin VB.TextBox txtID 
      Alignment       =   1  '������ ����
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   1600
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  '������ ����
      DataField       =   "AMOUNT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   1600
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "D:\vb60\�����\�����.mdb"
      DefaultCursorType=   0  '�⺻ Ŀ��
      DefaultType     =   2  'ODBC���
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  '���̳ʼ�
      RecordSource    =   "�����"
      Top             =   2760
      Width           =   4455
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "�з�"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "���"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "��ȣ"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "��¥"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "�ݾ�"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddNew_Click()
  
  '���ο� ���ڵ带 �߰��Ѵ�
  Data1.Recordset.AddNew

End Sub

Private Sub cmdDelete_Click()
  
  '���� ���ڵ带 �����մϴ�.
  Data1.Recordset.Delete
  '���� ���� �ƴϸ���� ���ڵ�� �̵�
  If Not Data1.Recordset.EOF Then
    Data1.Recordset.MoveNext
  '���� ���� ��� ������ ���ڵ�� �̵�
  Else
    Data1.Recordset.MoveLast
  End If
  
End Sub

Private Sub cmdMove_Click()

  '�ؽ�Ʈ�ڽ��� �Էµ� ���ڸ�ŭ �̵�
  Data1.Recordset.Move Val(txtMove.Text)
  
  '���ϳ��� ��� ���� ù��° ���ڵ� Ȥ��
  '������ ���ڵ�� �̵���Ų��.
  If Data1.Recordset.BOF Then
    Data1.Recordset.MoveFirst
  ElseIf Data1.Recordset.EOF Then
    Data1.Recordset.MoveLast
  End If
  
End Sub

Private Sub cmdMoveFirst_Click()

  'ù ��° ���ڵ�� �̵�
  Data1.Recordset.MoveFirst
  
End Sub

Private Sub cmdMoveLast_Click()
  
  '������ ���ڵ�� �̵�
  Data1.Recordset.MoveLast
  
End Sub

Private Sub cmdMoveNext_Click()
  
  '���� ���ڵ�� �̵�
  Data1.Recordset.MoveNext
  
  '������ �� ���̸�
  If Data1.Recordset.EOF Then
    '������ ���ڵ�� �̵�
    Data1.Recordset.MoveLast
  End If
  
End Sub

Private Sub cmdMovePrevious_Click()
  
  '���� ���ڵ�� �̵�
  Data1.Recordset.MovePrevious

  If Data1.Recordset.BOF Then
    'ù ��° ���ڵ�� �̵�
    Data1.Recordset.MoveFirst
  End If
  
End Sub

Private Sub Data1_Reposition()

  '������ ��Ʈ�ѿ� ���ڵ��� ��ġ�� ǥ��
  Data1.Caption = "���� ��ġ:" + Str(Data1.Recordset.AbsolutePosition)
  
End Sub

