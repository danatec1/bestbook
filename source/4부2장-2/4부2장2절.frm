VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   5550
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command2 
      Caption         =   "��    ��"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��    ��"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   2040
      ItemData        =   "4��2��2��.frx":0000
      Left            =   240
      List            =   "4��2��2��.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '���� ����
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Dim Inse, Univ, Department As String
   Inse = InputBox("���Ի� �̸��� �����Դϱ�?", "���Ի� �̸�")
   Univ = InputBox("���Ի� �б��� ����Դϱ�?", "�б�")
   Department = InputBox("���Ի� �а��� �����Դϱ�?", "�а�")
      '����Ʈ�ڽ��� inputbox�� ���� �Էµ� ���׵��� ��Ÿ���ϴ�.
   List1.AddItem Inse & "  " & Univ & "  " & Department
End Sub

Private Sub Command2_Click()
    a = MsgBox("���� �����Ͻðڽ��ϱ�?", vbOKCancel + vbQuestion, "���")
    If a = vbOK Then
    '���� ���õǾ��� �׸��� �����մϴ�.
     List1.RemoveItem (List1.ListIndex)
    End If
End Sub

Private Sub List1_Click()
    '���õ� �׸��� ���̺�ڽ��� ��Ÿ���ϴ�.
  Label1.Caption = List1.List(List1.ListIndex)
End Sub

