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
   Begin VB.Menu mnuPopup 
      Caption         =   "�˾��޴�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupMsg 
         Caption         =   "�޽��� ���̱�"
      End
      Begin VB.Menu mnuPopupExit 
         Caption         =   "�˾��޴� ����"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button And vbRightButton) = vbRightButton Then
    '��ư�� ������ �װ��� ������ ��ư�̶�� �Ʒ��� �ڵ带 �����մϴ�
        Me.PopupMenu mnuPopup, , X, Y
    '���� ��(Me)���� ���콺�� Ŭ�� ��ġ�� �˾� �޴��� ȣ���մϴ�
    End If
End Sub

Private Sub mnuPopupExit_Click()
    End
    '���α׷��� �����մϴ�
End Sub

Private Sub mnuPopupMsg_Click()
    MsgBox "�˾� �޴��� ȣ���ϼ̽��ϴ�"
End Sub

