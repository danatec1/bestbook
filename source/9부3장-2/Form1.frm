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
   Begin VB.CommandButton cmdCreate 
      Caption         =   "�����ͺ��̽� �����"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    
    '�����ͺ��̽�, ���̺�, �ʵ带 ������ ������ �����Ѵ�
    Dim MyDB As Database
    Dim MyTable As TableDef
    Dim MyField As Field

    '���ο� �����ͺ��̽� ������ �����Ѵ�
    Set MyDB = DBEngine.Workspaces(0).CreateDatabase("D:\�����.MDB", dbLangKorean, dbEncrypt)
    '���ο� ���̺��� �����Ѵ�
    Set MyTable = MyDB.CreateTableDef("�����")

    '���̺� �ʵ带 �߰��Ѵ�
    Set MyField = MyTable.CreateField("ID", dbLong)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("DATE", dbText, 10)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("SECT", dbText, 20)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("ITEM", dbText, 50)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("AMOUNT", dbLong)
    MyTable.Fields.Append MyField
    
    'TableDefs ��ü�� ���̺��� �߰��Ѵ�
    MyDB.TableDefs.Append MyTable
    MyDB.Close
    DBEngine.Workspaces(0).Close

End Sub


