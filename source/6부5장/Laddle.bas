Attribute VB_Name = "Laddle"
Public Type LeftPos  ' �����ٸ��� ������ ���� ��ġ�� ����
    X As Single
    Y As Single
End Type

Public Type RightPos  '�����ٸ��� ������ ���� ��ġ�� ����
   X As Single
   Y As Single
End Type

Public Type Result     '�ϳ��� ��ٸ� ���� ���۰� ����� ����
   ManName As String
   ResultNum As Integer
End Type

Public Const VerticalProgress = 0  '���� ��ٸ��� ���� �̵�
Public Const HorizontalLeftProgress = 1  '���� ��ٸ��� ���� �������� �̵�
Public Const HorizontalRightProgress = 2  '���� ��ٸ��� ���� ���������� �̵�
Public LaddleNum As Integer '���� ��ٸ��� ������ ����
Public Man As Integer   '������ ���� ����
Public Step As Integer  '�������� ���� ������ ����
Public NameF As Boolean  '�̸� ���� �÷���
Public RightPosition(100) As RightPos
Public LeftPosition(100) As LeftPos
Public RsltRcd(10) As Result
Public Vert As Integer   '���۵� ���� ���ڸ� ���� ���� ����
Public Vert1 As Integer '�ϳ��� ���� ���ο��� ����Ǿ� ���� ������ ���� ���� ����
Public Horiz As Integer  '������ ����ϱ� ���� ����
Public PM As Integer  '�������� ���� ������ ��Ÿ���� ���� ���� ����

Sub Main()
   Form1.Show
End Sub

Public Sub MakeLaddle(ManCnt As Integer)
  Dim Rslt As Integer
  Dim i As Integer
  Dim j As Integer
  '������ ���� ������� ��ε��Ѵ�.
  For i = 1 To (Man - 1)
     Unload Form1.Text1(i)
     Unload Form1.Shape1(i)
     Unload Form1.Label1(i)
  Next i
  
  For i = 1 To ((10 * Man) - 1)
    Unload Form1.Shape2(i)
  Next i
  For i = 1 To (ManCnt - 1)
     '��ٸ��� ���� ������� ������� ���߾� �ε��Ѵ�.
    Load Form1.Text1(i)
    Load Form1.Shape1(i)
    Load Form1.Label1(i)
  Next i
  For i = 1 To ((10 * ManCnt) - 1)
    Load Form1.Shape2(i)
  Next i
  
  For i = 0 To (ManCnt - 1)
   '������� �����Ѵ�.
     '���� ��ٸ�
    With Form1.Shape1(i)
      .Visible = True
      .Top = 700
      .Left = (Form1.Picture1.Width / (ManCnt + 1)) * (i + 1)
      .Height = Form1.Picture1.Height - 1400
      .Width = 50
    End With
    '��� �̸�
    With Form1.Text1(i)
       .Visible = True
       .Top = 700 - .Height
       .Left = Form1.Shape1(i).Left - (.Width / 2)
       .Text = "������" + CStr(i + 1)
    End With
  Next i
  
  Step = Form1.Shape1(0).Height / 48
  With Form1.Pointer
      .Top = Form1.Shape1(0).Top
      .Left = Form1.Shape1(0).Left - 25
      .Visible = True
  End With
  
  '���� ó���� 1�� �ڸ��� ���Ƿ� ���Ѵ�.
  Randomize
  Rslt = (Rnd() * ManCnt) + 1
  For i = 0 To (ManCnt - 1)
     With Form1.Label1(i)
        .Visible = False
        .Top = Form1.Shape1(i).Top + Form1.Shape1(i).Height + 100
        .Left = Form1.Shape1(i).Left - (.Width / 2)
        If (Rslt + i) > ManCnt Then Rslt = Rslt - ManCnt
        .Caption = CStr(Rslt + i)
     End With
  Next i
  LaddleNum = 0
  For i = 0 To (ManCnt - 2)
    '���� ��ٸ�
    For j = (0 + (10 * i)) To (9 + (10 * i))
       With Form1.Shape2(j)
          .Top = Form1.Shape1(i).Top + ((Form1.Shape1(i).Height / 12) * ((j Mod 10) + 1)) + ((Form1.Shape1(i).Height / 24) * (i Mod 2))
          .Left = Form1.Shape1(i).Left + 25
          .Height = 50
          .Width = Form1.Picture1.Width / (ManCnt + 1) + 10
          '���� Ȯ���� ��ġ
          Randomize
          If Int(Rnd() * 2) = 0 Then
              .Visible = False
          Else
              .Visible = True
              RightPosition(LaddleNum).X = .Left + .Width
              RightPosition(LaddleNum).Y = .Top
              LeftPosition(LaddleNum).X = .Left
              LeftPosition(LaddleNum).Y = .Top
              LaddleNum = LaddleNum + 1
          End If
       End With
     Next j
   Next i
  Man = ManCnt
End Sub

Public Sub ReturnResult()
  Dim ManCnt As Integer
  Dim i As Integer
  Dim ReturnString As String
  ManCnt = Form1.Shape1.Count
  For i = 0 To (ManCnt - 1)
     ReturnString = ReturnString + "     " + RsltRcd(i).ManName + "    :    " + CStr(RsltRcd(i).ResultNum) + "    " + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10)
  Next i
  MsgBox "<<���>>" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + ReturnString, , "���"
  
End Sub

