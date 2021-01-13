VERSION 5.00
Object = "{E54B6DC3-AE1F-11D1-A750-006097310C00}#1.0#0"; "GIFPLAY.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  '화면 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "나가기"
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin GIFPLAYLib.GifPlay GifPlay1 
      Height          =   1215
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   2143
      _StockProps     =   161
      BackColor       =   0
      AnimationGifFileName=   ""
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "요거 쓸만하죠...익스플로러처럼 로고가 뱅글뱅글 돌아가게 만들면 프로그램이 조금 돋보이겠죠"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Timer1_Timer()
     
  If Timer1.Enabled = True Then
     
    Call GifPlay1.LoadAnimationGifFile(App.Path & "\chair.gif")
            
    If GifPlay1.Play = False Then
      
     Else
     
     End If
      
     Timer1.Enabled = False
     
  End If
End Sub

