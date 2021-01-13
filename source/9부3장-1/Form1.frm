VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "DAO�� �̿��� �����ͺ��̽�"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdCancel 
      Caption         =   "���"
      Height          =   375
      Left            =   3600
      TabIndex        =   15
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "����"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "����"
      Height          =   375
      Left            =   2520
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  '������ ����
      DataField       =   "AMOUNT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   1600
   End
   Begin VB.TextBox txtID 
      Alignment       =   1  '������ ����
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1600
   End
   Begin VB.TextBox txtItem 
      DataField       =   "ITEM"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   2280
      Width           =   3585
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  '������ ����
      DataField       =   "DATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1600
   End
   Begin VB.TextBox txtSect 
      Alignment       =   1  '������ ����
      DataField       =   "SECT"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1600
   End
   Begin VB.CommandButton cmdMovePrevious 
      Caption         =   "����"
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveNext 
      Caption         =   "����"
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveFirst 
      Caption         =   "ó������"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdMoveLast 
      Caption         =   "������"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox txtMove 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "�̵�"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "�߰�"
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "����"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "�ݾ�"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "��¥"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "��ȣ"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   360
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "���"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "�з�"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1320
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDB As Database
Dim MySet As Recordset

Private Sub cmdCancel_Click()

  '������ ���
  MySet.CancelUpdate
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable
  
End Sub

Private Sub cmdEdit_Click()
  
  '���ڵ带 ����
  MySet.Edit
  '����, ��� ��ư�� Ȱ��ȭ
  SaveCancel_Enable
  
End Sub

Private Sub cmdSave_Click()
  
  '���� ���ڵ带 �����ͺ��̽��� ����
  MySet.Update
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable
  
End Sub

Private Sub Form_Load()
    
    '�����ͺ��̽� ���� ����
    Set MyDB = DBEngine.Workspaces(0).OpenDatabase("D:\vb60\�����\�����.MDB")
    '���̺� ����
    Set MySet = MyDB.OpenRecordset("�����", dbOpenTable)
    
    'ù ��° ���ڵ�� �̵�
    MySet.MoveFirst
    '���ڵ带 ȭ�鿡 ǥ��
    ShowRecord

End Sub

Private Sub cmdAddNew_Click()
  
  '���ο� ���ڵ带 �߰��Ѵ�
  MySet.AddNew
  '����, ��� ��ư�� Ȱ��ȭ
  SaveCancel_Enable

End Sub

Private Sub cmdDelete_Click()
  
  '���� ���ڵ带 �����մϴ�.
  MySet.Delete
  '���� ���� �ƴϸ���� ���ڵ�� �̵�
  If Not MySet.EOF Then
    MySet.MoveNext
  '���� ���� ��� ������ ���ڵ�� �̵�
  Else
    MySet.MoveLast
  End If
  
End Sub

Private Sub cmdMove_Click()

  '�ؽ�Ʈ�ڽ��� �Էµ� ���ڸ�ŭ �̵�
  MySet.Move Val(txtMove.Text)
  
  '���ϳ��� ��� ���� ù��° ���ڵ� Ȥ��
  '������ ���ڵ�� �̵���Ų��.
  If MySet.BOF Then
    MySet.MoveFirst
  ElseIf MySet.EOF Then
    MySet.MoveLast
  End If
  
  'ȭ�鿡 ���� ���ڵ带 ǥ��
  ShowRecord
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable

End Sub

Private Sub cmdMoveFirst_Click()

  'ù ��° ���ڵ�� �̵�
  MySet.MoveFirst
  'ȭ�鿡 ���� ���ڵ带 ǥ��
  ShowRecord
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable
  
End Sub

Private Sub cmdMoveLast_Click()
  
  '������ ���ڵ�� �̵�
  MySet.MoveLast
  'ȭ�鿡 ���� ���ڵ带 ǥ��
  ShowRecord
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable
  
End Sub

Private Sub cmdMoveNext_Click()
  
  '���� ���ڵ�� �̵�
  MySet.MoveNext
  
  '������ �� ���̸�
  If MySet.EOF Then
    '������ ���ڵ�� �̵�
    MySet.MoveLast
  End If
  
  'ȭ�鿡 ���� ���ڵ带 ǥ��
  ShowRecord
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable
  
End Sub

Private Sub cmdMovePrevious_Click()
  
  '���� ���ڵ�� �̵�
  MySet.MovePrevious

  If MySet.BOF Then
    'ù ��° ���ڵ�� �̵�
    MySet.MoveFirst
  End If
  
  'ȭ�鿡 ���� ���ڵ带 ǥ��
  ShowRecord
  '����, ��� ��ư�� ��Ȱ��ȭ
  SaveCancel_Disable

End Sub

Private Sub ShowRecord()

  'txtID.Text = MySet.Fields("ID")
  'txtDate.Text = MySet.Fields("DATE")
  'txtSect.Text = MySet.Fields("SECT")
  'txtAmount.Text = MySet.Fields("AMOUNT")
  'txtItem.Text = MySet.Fields("ITEM")

  txtID.Text = MySet!ID
  txtDate.Text = MySet!Date
  txtSect.Text = MySet!SECT
  txtAmount.Text = MySet!AMOUNT
  txtItem.Text = MySet!Item

End Sub

Private Sub SaveCancel_Enable()

  '����, ��� ��ư�� Ȱ��ȭ�Ѵ�
  cmdSave.Enabled = True
  cmdCancel.Enabled = True

End Sub

Private Sub SaveCancel_Disable()

  '����, ��� ��ư�� ��Ȱ��ȭ�Ѵ�
  cmdSave.Enabled = False
  cmdCancel.Enabled = False

End Sub

