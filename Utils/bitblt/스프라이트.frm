VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '���� ����
   Caption         =   "��������Ʈ ���ϸ��̼� ó��"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   388
   ScaleMode       =   3  '�ȼ�
   ScaleWidth      =   302
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.PictureBox picChar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   0
      Picture         =   "��������Ʈ.frx":0000
      ScaleHeight     =   1500
      ScaleWidth      =   4500
      TabIndex        =   5
      Top             =   3780
      Width           =   4560
   End
   Begin VB.PictureBox picBack 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1560
      Left            =   0
      Picture         =   "��������Ʈ.frx":7972
      ScaleHeight     =   1500
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   2205
      Width           =   4560
   End
   Begin VB.PictureBox picWork 
      Height          =   1500
      Left            =   0
      ScaleHeight     =   1440
      ScaleWidth      =   4440
      TabIndex        =   3
      Top             =   0
      Width           =   4500
   End
   Begin VB.Timer Timer1 
      Left            =   105
      Top             =   5355
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���ϸ��̼�"
      Height          =   540
      Left            =   3360
      TabIndex        =   2
      Top             =   1575
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��������Ʈ"
      Height          =   540
      Left            =   1785
      TabIndex        =   1
      Top             =   1575
      Width           =   1485
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��������Ʈ �ܰ�"
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   1575
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
                                             ByVal nWidth As Long, ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const SRCERASE = &H440328
Private Const SRCINVERT = &H660046
Dim MoveX, CharNum, I As Integer

Public Sub Animation(Process As String)
  '�۾� ȭ�鿡 ��� �׸��� ����
  If (Process = "All") Or (Process = "Back") Then BitBlt picWork.hDC, 0, 0, 300, 100, picBack.hDC, 0, 0, SRCCOPY
  '�۾� ȭ�鿡 ����ũ �׸��� ����(����� �����Ͽ� And ����)
  If (Process = "All") Or (Process = "Mask") Then BitBlt picWork.hDC, MoveX, 20, 60, 45, picChar.hDC, CharNum * 60, 45, SRCERASE
  '�۾� ȭ�鿡 ĳ���� �׸��� ����(Xor ����)
  If (Process = "All") Or (Process = "Char") Then BitBlt picWork.hDC, MoveX, 20, 60, 45, picChar.hDC, CharNum * 60, 0, SRCINVERT
  
  If (Process = "All") Or (Process = "Char") Then MoveX = MoveX + 15
  If MoveX >= 230 Then MoveX = 20
  If (Process = "All") Or (Process = "Char") Then CharNum = CharNum + 1
  If CharNum >= 5 Then CharNum = 0
End Sub

Private Sub Command1_Click()
  Select Case I
    Case 0:
      Animation "Back" '��� �׸� ����
    Case 1:
      Animation "Mask" '����ũ ����
    Case 2:
      Animation "Char" 'ĳ���� ����
  End Select
  I = I + 1
  If I >= 3 Then I = 0
End Sub

Private Sub Command2_Click()
  Animation "All"
End Sub

Private Sub Command3_Click()
  If Timer1.Enabled Then
    Command1.Enabled = True
    Command2.Enabled = True
    Timer1.Enabled = False
  Else
    Command1.Enabled = False
    Command2.Enabled = False
    Timer1.Enabled = True
  End If
End Sub

Private Sub Form_Load()
  I = 0
  MoveX = 20
  CharNum = 0
  Timer1.Enabled = False
  Timer1.Interval = 100
  Form1.Height = 2505
  
  '���� �߾ӿ� ����
  Form1.Top = (Screen.Height - Form1.Height) / 2
  Form1.Left = (Screen.Width - Form1.Width) / 2
End Sub

Private Sub Timer1_Timer()
  Animation "All"
End Sub
