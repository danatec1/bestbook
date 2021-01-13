VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  '단일 고정
   Caption         =   "사다리 게임"
   ClientHeight    =   7275
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture1 
      BackColor       =   &H0080FF80&
      Height          =   6375
      Left            =   0
      ScaleHeight     =   6315
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   720
      Width           =   8895
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   3360
         Top             =   3240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer1 
         Left            =   3720
         Top             =   5400
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Text            =   "참가자 이름을 넣으세요."
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Index           =   0
         Left            =   1080
         TabIndex        =   3
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00C000C0&
         FillColor       =   &H000040C0&
         Height          =   615
         Index           =   0
         Left            =   1680
         Top             =   2400
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C000C0&
         FillColor       =   &H000040C0&
         Height          =   615
         Index           =   0
         Left            =   480
         Top             =   1920
         Width           =   615
      End
      Begin VB.Shape Pointer 
         BackColor       =   &H00C000C0&
         FillColor       =   &H000040C0&
         Height          =   615
         Left            =   1800
         Top             =   480
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   3690
         Left            =   6000
         Picture         =   "6부5장사다리게임.frx":0000
         Stretch         =   -1  'True
         ToolTipText     =   "더블클릭하면 게임을 시작합니다."
         Top             =   2400
         Width           =   2595
      End
   End
   Begin MSComctlLib.Toolbar ToolBar1 
      Align           =   1  '위 맞춤
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "새 게임"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "참가자 이름 지정"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "나가기"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "6부5장사다리게임.frx":80EEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "6부5장사다리게임.frx":8117E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "6부5장사다리게임.frx":81412
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Game 
      Caption         =   "게임"
      Begin VB.Menu New 
         Caption         =   "새 게임"
         Shortcut        =   ^N
      End
      Begin VB.Menu Name 
         Caption         =   "참가자 이름 지정"
         Shortcut        =   ^P
      End
      Begin VB.Menu Quit 
         Caption         =   "나가기"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Info 
      Caption         =   "정보"
      Begin VB.Menu GameInfo 
         Caption         =   "게임정보"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu Help 
      Caption         =   "도움말"
      Begin VB.Menu Helpgame 
         Caption         =   "사다리게임"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Pointer.Height = 100 '포인터의 크기 조정
  Pointer.Width = 100
  Image1.Width = Picture1.Width + 20 '사다리를 덮을 그림의 크기 및 위치 조정
  Image1.Height = Picture1.Height - 300
  Image1.Left = Picture1.Left - 50
  Image1.Top = Picture1.Top + 300
  MakeLaddle 4  '디폴트로 4인용 사다리를 만듬
End Sub

Private Sub Helpgame_Click()
  CommonDialog1.HelpFile = "Game.hlp"
  CommonDialog1.HelpCommand = cdlHelpContents
  CommonDialog1.ShowHelp
End Sub

Private Sub Name_Click()
  Dim ManCnt As Integer
  Dim i As Integer
  If NameF = 0 Then
     ManCnt = Text1.Count
     For i = 0 To (ManCnt - 1)
         Text1(i).Enabled = True
     Next i
     NameF = 1
  Else
     ManCnt = Text1.Count
     For i = 0 To (ManCnt - 1)
        Text1(i).Enabled = False
     Next i
     NameF = 0
  End If
  
End Sub
Private Sub New_Click()
  Form2.Show vbModal
  
End Sub

Private Sub GameInfo_Click()
  MsgBox "사다리 게임 Version 1.0" + Chr$(13) + Chr$(10) + "제작 일자 : 1998년 11월 17일", vbInformation
End Sub

Private Sub Quit_Click()
  Dim Rslt As Variant
  Rslt = MsgBox("게임을 종료하시겠습니까?", vbInformation Or vbOKCancel, "종료 확인")
  If Rslt = vbOK Then
    End
  End If
  
End Sub

Private Sub Timer1_Timer()
  Dim i As Integer
  '수직선상에 있을 때
  If PM = VerticalProgress Then '수직운동 상황에서
    For i = 0 To LaddleNum
      '수직선 기준 오른쪽 수평선에 다달았을 때
      If ((Pointer.Top - LeftPosition(i).Y) < Step _
      And (Pointer.Top - LeftPosition(i).Y) > -Step) _
      And ((Pointer.Left - LeftPosition(i).X) < Step _
      And (Pointer.Left - LeftPosition(i).X) > -Step) Then
          Pointer.Move Pointer.Left + Step, LeftPosition(i).Y
          PM = HorizontalRightProgress
          Horiz = i
          Exit Sub
      End If
      '수직선 기준 왼쪽 수평선에 다달았을 때
     If ((Pointer.Top - RightPosition(i).Y) < Step _
      And (Pointer.Top - RightPosition(i).Y) > -Step) _
      And ((Pointer.Left - RightPosition(i).X) < Step _
      And (Pointer.Left - RightPosition(i).X) > -Step) Then
          Pointer.Move Pointer.Left - Step, RightPosition(i).Y
          PM = HorizontalLeftProgress
          Horiz = i
          Exit Sub
      End If
    Next i
    '포인터가 수직선의 끝에 다달았을 때
  If Pointer.Top >= (Shape1(Vert1).Top + Shape1(Vert1).Height) Then
    Vert = Vert + 1
    If Vert >= Shape1.Count Then
       Timer1.Enabled = False
       RsltRcd(Vert - 1).ResultNum = Label1(Vert1).Caption
       MsgBox RsltRcd(Vert - 1).ManName + ":" + CStr(RsltRcd(Vert - 1).ResultNum)
       ReturnResult
       Exit Sub
    End If
    RsltRcd(Vert - 1).ResultNum = Label1(Vert1).Caption
    MsgBox RsltRcd(Vert - 1).ManName + ":" + CStr(RsltRcd(Vert - 1).ResultNum)
    Vert1 = Vert
    Pointer.Move Shape1(Vert).Left + 25, Shape1(Vert).Top
    Exit Sub
  End If
  Pointer.Move Pointer.Left, Pointer.Top + Step
  Exit Sub
 End If
 '수평선 오른쪽 진행
 If PM = HorizontalRightProgress Then
     '오른쪽 끝 도달
   If (Pointer.Left - RightPosition(Horiz).X) < Step And _
      (Pointer.Left - RightPosition(Horiz).X) > -Step Then
      Vert1 = Vert1 + 1
      Pointer.Move Shape1(Vert1).Left, Pointer.Top + Step
      PM = VerticalProgress
      Exit Sub
    End If
    Pointer.Move Pointer.Left + Step, Pointer.Top
    Exit Sub
  End If
    ' 수평선 왼쪽 진행
  If PM = HorizontalLeftProgress Then
     '왼쪽 끝 도달
    If (Pointer.Left - LeftPosition(Horiz).X) < Step And _
       (Pointer.Left - LeftPosition(Horiz).X) > -Step Then
     Vert1 = Vert1 - 1
     Pointer.Move Shape1(Vert1).Left, Pointer.Top + Step
     PM = VerticalProgress
     Exit Sub
     End If
     Pointer.Move Pointer.Left - Step, Pointer.Top
     Exit Sub
   End If
   
   
   
End Sub

Private Sub ToolBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim ManCnt As Integer
  Dim i As Integer
  Select Case Button.Index
    Case 1 '새 게임 버튼을 클릭했을 때
      Form2.Show vbModal
    Case 2  '참가자 이름 지정 커튼을 클릭했을 때
      If NameF = 0 Then
         ManCnt = Text1.Count
         For i = 0 To (ManCnt - 1)
            Text1(i).Enabled = True
         Next i
         NameF = 1
      Else
         ManCnt = Text1.Count
         For i = 0 To (ManCnt - 1)
            Text1(i).Enabled = False
         Next i
         NameF = 0
      End If
    Case 4 '나가기 버튼을 클릭했을 때
      Dim Rslt As Variant
      Rslt = MsgBox("게임을 종료하시겠습니까?", vbInformation Or vbOKCancel, "종료 확인")
      If Rslt = vbOK Then
         End
      End If
  End Select
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then  '마우스 오른쪽 버튼을 클릭했을 때
    PopupMenu Game
  End If
  
End Sub

Private Sub Image1_DblClick()
  Dim ManCnt As Integer
  Dim i As Integer
  Dim j As Integer
  Image1.Visible = False  '게임을 시작하기 위해 그림을 없앤다.
  i = Label1.Count
  For j = 0 To (i - 1)
     Label1(j).Visible = True
  Next j
  ManCnt = Shape1.Count
  For i = 0 To (ManCnt - 1)
     RsltRcd(i).ManName = Text1(i).Text
  Next i
  Vert = 0
  Vert1 = 0
  PM = VerticalProgress
  Timer1.Interval = 50
  Timer1.Enabled = True
  
End Sub


