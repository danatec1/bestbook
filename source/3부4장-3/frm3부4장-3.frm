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
      Caption         =   "�������� ����!"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblResult 
      BorderStyle     =   1  '���� ����
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdGetValue_Click()
    Dim StrDataForInputBox As String 'InputBox �Լ��� ���� ��Ʈ��
    Dim StrData As String '������ ������� ���� ��Ʈ��
    Dim IntDataForCast As Integer 'InputBox �Լ��� ���� ���� ���ڷ� ��ȯ
    Dim i As Integer 'For ~ Next ������ ���� ���� ����
    
    lblResult.Caption = "" '���̺��� �ʱ�ȭ�Ѵ�
    StrDataForInputBox = InputBox("�������� ����� ���� �Է��ϼ���.", _
                            "�ܼ� �Է�", "2") '����Ʈ�� 2�� �Է�
    If StrDataForInputBox = "" Then 'InputBox ���� < Esc > Ű�� ������ ���
        MsgBox "������ ����� ����մϴ�.", vbInformation, "��� ���"
        Exit Sub '������ ����� ����մϴ�
    End If
    
    For i = 1 To 9
        IntDataForCast = CInt(StrDataForInputBox) 'InputBox ���ڸ� ���ڷ� ��ȯ
        StrData = CStr(IntDataForCast) & " * " & _
                    CStr(i) & " = " & CStr(IntDataForCast * i) & _
                    Chr$(10) & Chr$(13)     '��갪�� ������ ����
        lblResult.Caption = lblResult.Caption & StrData '���̺�� ���
    Next i
End Sub


