VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   990
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   ScaleHeight     =   990
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdQuit 
      Caption         =   "����"
      Height          =   660
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "����Ʈ"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
    Dim Result As Integer '�޼��� �ڽ��� ������� ���� ����
    Result = MsgBox("����Ʈ �۾��� �Ͻðڽ��ϱ�?", vbOKCancel Or _
                        vbInformation, "����Ʈ Ȯ��") '����Ʈ �۾� ���� Ȯ��
    If Result = vbOK Then 'OK��ư�� ������ ���
        MsgBox "����Ʈ �۾��� ���������� �����Ͽ����ϴ�.", , "����Ϸ�"
    End If
End Sub

Private Sub cmdQuit_Click()
    Dim Result As Integer '�޼��� �ڽ��� ������� ���� ����
    Result = MsgBox("���α׷��� �����Ͻðڽ��ϱ�?", vbYesNo _
            Or vbCritical Or vbApplicationModal, "���� Ȯ��")
            '���α׷� ���� Ȯ��
    If Result = vbYes Then 'Yes��ư�� ������ ���
        End
    Else
        MsgBox "���Ḧ ����մϴ�.", , "���"
    End If
End Sub



