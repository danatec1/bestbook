VERSION 5.00
Begin VB.Form frmFlash 
   Caption         =   "Window Flash Test"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Timer tmrFlash 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   240
      Top             =   240
   End
   Begin VB.CommandButton cmdFlashOff 
      Caption         =   "Flash Off"
      Height          =   735
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   1695
   End
   Begin VB.CommandButton cmdFlashOn 
      Caption         =   "Flash On"
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Flash Off Ŀ�ǵ��ư�� Ŭ���� �� ����
Private Sub cmdFlashOff_Click()
    tmrFlash.Enabled = False
End Sub

'Flash On Ŀ�ǵ��ư�� Ŭ���� �� ����
Private Sub cmdFlashOn_Click()
    tmrFlash.Enabled = True
End Sub

'Ÿ�̸� �Լ�
Private Sub tmrFlash_Timer()
    '���� �÷��� ��Ű�� �Լ��� �����ϴ� �κ�
    FlashWindow frmFlash.hwnd, True
End Sub


