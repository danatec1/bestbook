VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "ProgressBar"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton Command1 
      Caption         =   "���α׷�����"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim I As Integer
  Dim ArrProgress(250) As String
   ProgressBar1.Min = LBound(ArrProgress)
   ProgressBar1.Max = UBound(ArrProgress)
   '���α׷������� Value �Ӽ��� Min �Ӽ������� �����մϴ�.
   ProgressBar1.Value = ProgressBar1.Min
   '����ؼ� �迭�� ��ȯ��ŵ�ϴ�.
   For I = LBound(ArrProgress) To UBound(ArrProgress)
            '�迭�� �ִ� �� �׸�鿡 ���� �ʱⰪ�� �����մϴ�.
         ArrProgress(I) = "Initial value" & I
         ProgressBar1.Value = I
   Next I
   ProgressBar1.Visible = False
   ProgressBar1.Value = ProgressBar1.Min
End Sub

