VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ   ��"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "�� ��ȿ"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "�� �ؿ�"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "�� â��"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "�� �¿�"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "�� ����"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "�� ����"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "���� �� ����� �����ϴ� �ŷ�Ʈ�� ��� �����Ͻʽÿ�."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 
   If Check1.Value = Checked Then
       A = "������ "
   Else
       A = ""
   End If
   If Check2.Value = Checked Then
       B = "������ "
   Else
       B = ""
   End If
   If Check3.Value = Checked Then
       C = "�̽¿� "
   Else
       C = ""
   End If
   If Check4.Value = Checked Then
       D = "��â�� "
   Else
       D = ""
   End If
   If Check5.Value = Checked Then
       E = "���ؿ� "
   Else
       E = ""
   End If
   If Check6.Value = Checked Then
       F = "���ȿ "
   Else
       F = ""
   End If
        '������ ���� �ʾ��� ��� �޽��� ����
   If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 Then
      MsgBox "������ ���� �ʾҽ��ϴ�. ������ ���ּ���!!", vbOKOnly, "�˸�"
   Else
        '���û����� ��Ÿ��.
   G = A + B + C + D + E + F
   MsgBox (G + "���� �����Ͻô±���."), vbInformation, "�˸�"
   End If
End Sub

