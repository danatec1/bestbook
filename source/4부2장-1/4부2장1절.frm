VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ   ��"
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  '���� ����
      Height          =   1695
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "� ��1 :"
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "���� ��1 :"
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���� ��1 :"
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Click()
   If Combo2.ListIndex = 2 Then
       Label3.Visible = True
       Combo3.Visible = True
   End If
 
End Sub

Private Sub Command1_Click()
    Dim Check As String
    A = Combo1.Text
    If Combo2.ListIndex = 2 Then
        B = Combo3.Text
    Else
        B = Combo2.Text
    End If
    Check = MsgBox("����� " & A & ", " & B & "�� �����ϼ̽��ϴ�.�½��ϱ�?", _
                      vbYesNo + vbInformation, "�˸�")
     If Check = vbYes Then
        Label4.Caption = A & Chr(13) & B
     Else
        Combo1.SetFocus
        Combo3.Visible = False
        Label3.Visible = False
     End If
End Sub

Private Sub Form_Load()
  Combo1.AddItem "ǰ�� ����", 0
  Combo1.AddItem "�������", 1
  Combo1.AddItem "�����ȹ��", 2
  Combo1.AddItem "O.R.", 3
  Combo2.AddItem "������ ����", 0
  Combo2.AddItem "�����а���", 1
  Combo2.AddItem "�", 2
  Combo3.Visible = False
  Label3.Visible = False
    Combo3.AddItem "Ź��", 0
   Combo3.AddItem "�౸", 1
   Combo3.AddItem "�߱�", 2
End Sub

