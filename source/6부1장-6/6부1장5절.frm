VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "��� ����"
   ClientHeight    =   3195
   ClientLeft      =   3915
   ClientTop       =   3060
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3360
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    '������ ��ü�� ���Ͽ� ��½� ����� ���۵Ǵ� ��ġ�� ����.
    Printer.CurrentX = 100
    Printer.CurrentY = 100
    '��½� ��¹��� �۲ø�� �۲� ũ�⸦ ����Ϸ��� �������κ����� �̸��� ũ��� ����.
    Printer.FontName = Text1.FontName
    Printer.FontSize = Text1.FontSize
    '�ؽ�Ʈ�ڽ��� ������ ���
    Printer.Print Text1.Text
    '�����ͷ��� ����� �Ϸ���� �˸�.
    Printer.EndDoc
End Sub

