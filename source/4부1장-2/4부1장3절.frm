VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3750
   ClientTop       =   2850
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ   ��"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "��  ��"
      Height          =   2775
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      Begin VB.OptionButton Option5 
         Caption         =   "�� �� ��"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "�� �� ��"
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "����л�"
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��  ��"
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
      Begin VB.OptionButton Option2 
         Caption         =   "   ��"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "   ��"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
  Dim A, B, C As String
    If Option1.Value = True Then
    A = "����"
  ElseIf Option2.Value = True Then
    A = "����"
  End If
  If Option3.Value = True Then
     B = "����л�"
  ElseIf Option4.Value = True Then
     B = "���л�"
  ElseIf Option5.Value = True Then
     B = "������"
  End If
   '������ ���� ���� ���
  If Option1.Value = 0 And Option2.Value = 0 Or Option3.Value = 0 And Option4.Value = 0 And Option5.Value = 0 Then
    MsgBox "������ ���� �ʾҽ��ϴ�. ������ ���ּ���.", vbOKOnly, "�˸�"
   Else
    '������ �� ���
     C = A + B
    MsgBox (C + " �̽ñ���!"), vbInformation, "�˸�"
  End If
  End Sub

