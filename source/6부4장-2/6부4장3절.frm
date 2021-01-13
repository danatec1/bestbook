VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Begin VB.Form Form1 
   Caption         =   "멀티미디어 플레이어"
   ClientHeight    =   5355
   ClientLeft      =   3705
   ClientTop       =   2850
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   5355
   ScaleWidth      =   6690
   Begin VB.CommandButton CmdEnd 
      Caption         =   "종    료"
      Height          =   375
      Left            =   5160
      TabIndex        =   12
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "Pause"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton CmdPrev 
      Caption         =   "Prev"
      Height          =   375
      Left            =   5160
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "Play"
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin MCI.MMControl MMControl1 
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   661
      _Version        =   393216
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "6부4장3절.frx":0000
      Left            =   2520
      List            =   "6부4장3절.frx":0002
      Style           =   2  '드롭다운 목록
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1770
      Left            =   2520
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   4800
      X2              =   4800
      Y1              =   240
      Y2              =   4200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      TabIndex        =   6
      Top             =   3840
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_DblClick()
   '파일리스트를 더블클릭하면 Open 명령을 내리게 함.
    CmdPlay_Click
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
  ' 콤보박스에 항목을 추가함.
  Combo1.AddItem "웨이브 파일(*.wav)"
  Combo1.AddItem "동영상 파일(*.avi)"
  Combo1.AddItem "CD 플레이어"
  '초기설정값으로 웨이브 파일을 지정.
  Combo1.ListIndex = 0
  MMControl1.Command = "Open"
  MMControl1.DeviceType = "WaveAudio"
  File1.Pattern = "*.wav"
End Sub

Private Sub Combo1_Click()
  MMControl1.Command = "Stop"
  MMControl1.Command = "Close"
  '콤보박스의 리스트들의 인덱스 값에 의해 장치의 유형을 선택하게 함.
  Select Case Combo1.ListIndex
  Case 0
         ' 장치유형을 WaveAudio로 설정.
         MMControl1.DeviceType = "WaveAudio"
         File1.Pattern = "*.wav"
         Drive1.Enabled = True
         Dir1.Enabled = True
         File1.Enabled = True
         CmdPlay.Enabled = True
  Case 1
        '장치유형을 AviVideo로 설정.
        MMControl1.DeviceType = "AviVideo"
        File1.Pattern = "*.avi"
        Drive1.Enabled = True
        Dir1.Enabled = True
        File1.Enabled = True
        CmdPlay.Enabled = True
  Case 2
        '장치유형을 CDAudio로 설정
        MMControl1.DeviceType = "CDAudio"
        MMControl1.FileName = ""
        Drive1.Enabled = False
        Dir1.Enabled = False
        File1.Enabled = False
        CmdPlay.Enabled = False
  End Select
End Sub

Private Sub CmdEnd_Click()
  MMControl1.Command = "Close"
  Unload Me
End Sub

Private Sub MMControl1_StatusUpdate()
  If MMControl1.DeviceType = "WaveAudio" Then
      ProgressBar1.Value = MMControl1.Position
      Label1.Caption = Format(ProgressBar1.Value / 100, "#0.00초")
  ElseIf MMControl1.DeviceType = "AviVideo" Then
      ProgressBar1.Value = MMControl1.Position
      Label1.Caption = Format(ProgressBar1.Value, "0") + "Frame"
  Else
      Label1.Caption = Format(MMControl1.Track, "0트랙")
  End If
End Sub

Private Sub CmdPlay_Click()
  If File1.FileName = "" Then
     MsgBox "파일을 선택하지 않았습니다. 파일을 선택하십시오.", vbOKOnly, "경고"
     Exit Sub
  End If
  MMControl1.Command = "Close"
  MMControl1.FileName = File1.Path + "/" + File1.FileName
  MMControl1.Command = "Open"
  ProgressBar1.Max = MMControl1.Length
  MMControl1.Command = "Play"
  
End Sub

Private Sub CmdNext_Click()
   MMControl1.Command = "Next"
End Sub

Private Sub CmdPause_Click()
   MMControl1.Command = "Pause"
End Sub

Private Sub CmdPrev_Click()
   MMControl1.Command = "Prev"
End Sub

Private Sub CmdStop_Click()
   MMControl1.Command = "Stop"
End Sub

