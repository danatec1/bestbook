VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   3750
   ClientTop       =   2445
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ   ��"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��  ��ȿ"
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��  �ؿ�"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��  â��"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��  �¿�"
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��  ����"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "��  ����"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "���� �� ����� �����ϴ� �ŷ�Ʈ�� ��� �����Ͻʽÿ�."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'�� ��Ʈ�ѹ迭���� üũ�Ǿ������� �˻��Ͽ�, ��º����� ���� �ݴϴ�.
For i = 1 To 6 Step 1
   If Check1(i - 1).Value = Checked Then
       A = Check1(i - 1).Caption
   Else
       A = ""
   End If
       B = B & " " & A
Next i
   MsgBox (B & "���� �����Ͻô±���."), vbInformation, "�˸�"
End Sub


