VERSION 5.00
Begin VB.Form frmMemo 
   BorderStyle     =   4  '���� ���� â
   Caption         =   "����Ʈ ��"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4485
   LinkTopic       =   "frmMemo"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdMemo 
      Caption         =   "�� �޸�"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox txtMemo 
      Height          =   2775
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmMemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  
  '�� �����츦 �ݴ´�
  Unload Me
  
End Sub

Private Sub cmdMemo_Click()
  
  '������ �� ������ ��ü�� Clone�� ����
  Dim Clone As Form
 
  '����Ʈ �� ���� Clone�� ����
  Set Clone = New frmMemo
  Clone.Show

End Sub

