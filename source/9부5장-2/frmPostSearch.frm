VERSION 5.00
Begin VB.Form frmPostSerch 
   Caption         =   "�����ȣ �˻��� Ver 1.0"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdSearch 
      Caption         =   "ã���ּ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   5
      Top             =   2340
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Height          =   1125
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   2625
      Begin VB.Label Label2 
         Caption         =   "���� ã���� �ϴ� ���� �̸��� �Է��ϰ� 'ã��'��ư�� ��������.��) ����='�б���'"
         Height          =   555
         Left            =   210
         TabIndex        =   4
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.TextBox txtDong 
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   480
      Width           =   1545
   End
   Begin VB.ListBox QueryResult 
      Height          =   2760
      Left            =   2820
      TabIndex        =   0
      Top             =   120
      Width           =   5025
   End
   Begin VB.Label Result 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '���� ����
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1230
      TabIndex        =   7
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label Label3 
      Caption         =   "�����ȣ"
      Height          =   285
      Left            =   210
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "��(��)��"
      Height          =   285
      Left            =   210
      TabIndex        =   1
      Top             =   510
      Width           =   705
   End
End
Attribute VB_Name = "frmPostSerch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents Rs As Recordset
Attribute Rs.VB_VarHelpID = -1

Private Sub cmdSearch_Click()

  Dim db As Connection
  Set db = New Connection
  'Ŭ���̾�Ʈ Ŀ�� ���
  db.CursorLocation = adUseClient
  '�����ͺ��̽� ����
  db.Open "dsn=�����ȣ;uid=;pwd=;"

  Dim YesNO As Integer

  If txtDong.Text = "" Then
    YesNO = MsgBox("���� ���� �Էµ��� �ʾҽ��ϴ�. ��ü�� �˻��Ͻðڽ��ϱ�?", vbYesNo + vbQuestion, "���� �Էµ��� �ʾҽ��ϴ�.")
    If YesNO = vbNo Then
      Exit Sub
    End If
  End If

  'SQL ���� �̿��Ͽ� ������ ����
  Set Rs = New Recordset
  Rs.Open "select �����ȣ,����,��ü�ּ� from �����ȣ where ���� LIKE '" + txtDong.Text + "%'", db, adOpenKeyset, adLockOptimistic

  QueryResult.Clear
  
  '�����ͺ��̽��� ũ�Ⱑ 0�����̸� �˻��� ���ڵ� ����
  If Rs.RecordCount > 0 Then
    Rs.MoveFirst
    Do While Not Rs.EOF
      '����Ʈ�ڽ� ��Ʈ�ѿ� [�����ȣ] ��ü �ּ��� ���·� ���ڵ带 �߰�
      QueryResult.AddItem "[" + Mid(Rs!�����ȣ, 2, 3) + "-" + Mid(Rs!�����ȣ, 4, 3) + "] " + Rs!��ü�ּ�
      Rs.MoveNext
    Loop
  Else
    MsgBox "�˻��� ���ڵ尡 �����ϴ�."
  End If
    
End Sub

Private Sub QueryResult_Click()
    Result.Caption = Mid(QueryResult.Text, 2, 3) + "-" + Mid(QueryResult.Text, 6, 3)
End Sub

Private Sub QueryResult_Scroll()
    QueryResult_Click
End Sub
