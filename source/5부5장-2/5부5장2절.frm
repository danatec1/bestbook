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
   Begin VB.Label Label1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Dim ButtonTest As Integer
     ButtonTest = Button And 7
     Select Case ButtonTest
Case 1
  Label1.Caption = "���� ��ư�� �������ϴ�."
Case 2
  Label1.Caption = "������ ��ư�� �������ϴ�."
Case 3
  Label1.Caption = "���ʰ� ������ ��ư�� �������ϴ�."
Case 4
  Label1.Caption = "��� ��ư�� �������ϴ�."
Case 5
  Label1.Caption = "���ʰ� ��� ��ư�� �������ϴ�."
Case 6
  Label1.Caption = "�����ʰ� ��� ��ư�� �������ϴ�."
Case 7
  Label1.Caption = "��� ��ư�� �� �������ϴ�."
    End Select
End Sub

