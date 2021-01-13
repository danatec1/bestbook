Attribute VB_Name = "Laddle"
Public Type LeftPos  ' 수평사다리의 오른쪽 끝점 위치를 저장
    X As Single
    Y As Single
End Type

Public Type RightPos  '수평사다리의 오른쪽 끝점 위치를 저장
   X As Single
   Y As Single
End Type

Public Type Result     '하나의 사다리 진행 시작과 결과를 저장
   ManName As String
   ResultNum As Integer
End Type

Public Const VerticalProgress = 0  '수직 사다리를 따라 이동
Public Const HorizontalLeftProgress = 1  '수평 사다리를 따라 왼쪽으로 이동
Public Const HorizontalRightProgress = 2  '수평 사다리를 따라 오른쪽으로 이동
Public LaddleNum As Integer '수평 사다리의 갯수를 저장
Public Man As Integer   '참가자 수를 저장
Public Step As Integer  '포인터의 진행 단위를 저장
Public NameF As Boolean  '이름 변경 플래그
Public RightPosition(100) As RightPos
Public LeftPosition(100) As LeftPos
Public RsltRcd(10) As Result
Public Vert As Integer   '시작된 라인 숫자를 위한 전역 변수
Public Vert1 As Integer '하나의 시작 라인에서 변경되어 가는 라인을 위한 전역 변수
Public Horiz As Integer  '수평선을 기억하기 위한 변수
Public PM As Integer  '포인터의 진행 방향을 나타내기 위한 전역 변수

Sub Main()
   Form1.Show
End Sub

Public Sub MakeLaddle(ManCnt As Integer)
  Dim Rslt As Integer
  Dim i As Integer
  Dim j As Integer
  '이전에 쓰인 자재들을 언로드한다.
  For i = 1 To (Man - 1)
     Unload Form1.Text1(i)
     Unload Form1.Shape1(i)
     Unload Form1.Label1(i)
  Next i
  
  For i = 1 To ((10 * Man) - 1)
    Unload Form1.Shape2(i)
  Next i
  For i = 1 To (ManCnt - 1)
     '사다리에 쓰일 자재들을 사람수에 맞추어 로드한다.
    Load Form1.Text1(i)
    Load Form1.Shape1(i)
    Load Form1.Label1(i)
  Next i
  For i = 1 To ((10 * ManCnt) - 1)
    Load Form1.Shape2(i)
  Next i
  
  For i = 0 To (ManCnt - 1)
   '자재들을 정렬한다.
     '수직 사다리
    With Form1.Shape1(i)
      .Visible = True
      .Top = 700
      .Left = (Form1.Picture1.Width / (ManCnt + 1)) * (i + 1)
      .Height = Form1.Picture1.Height - 1400
      .Width = 50
    End With
    '사람 이름
    With Form1.Text1(i)
       .Visible = True
       .Top = 700 - .Height
       .Left = Form1.Shape1(i).Left - (.Width / 2)
       .Text = "참가자" + CStr(i + 1)
    End With
  Next i
  
  Step = Form1.Shape1(0).Height / 48
  With Form1.Pointer
      .Top = Form1.Shape1(0).Top
      .Left = Form1.Shape1(0).Left - 25
      .Visible = True
  End With
  
  '난수 처리로 1번 자리를 임의로 구한다.
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
    '세로 사다리
    For j = (0 + (10 * i)) To (9 + (10 * i))
       With Form1.Shape2(j)
          .Top = Form1.Shape1(i).Top + ((Form1.Shape1(i).Height / 12) * ((j Mod 10) + 1)) + ((Form1.Shape1(i).Height / 24) * (i Mod 2))
          .Left = Form1.Shape1(i).Left + 25
          .Height = 50
          .Width = Form1.Picture1.Width / (ManCnt + 1) + 10
          '난수 확률로 배치
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
  MsgBox "<<결과>>" + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + ReturnString, , "결과"
  
End Sub

