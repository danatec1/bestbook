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
   Begin VB.ListBox List1 
      Height          =   1860
      ItemData        =   "5��5��1-1��.frx":0000
      Left            =   240
      List            =   "5��5��1-1��.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    List1.AddItem "�� �� �� �ٽ� - �� â�� "
    List1.AddItem "�׳���� �̺� - �� ���� "
    List1.AddItem "�����ʴ� ���� - �� ���� "
    List1.AddItem "�ƻԽ� - �� ��� "
End Sub

Private Sub List1_DblClick()
   Label1.Caption = List1.Text
End Sub

