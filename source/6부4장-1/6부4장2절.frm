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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "프로그레스바"
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
   '프로그레스바의 Value 속성을 Min 속성값으로 설정합니다.
   ProgressBar1.Value = ProgressBar1.Min
   '계속해서 배열을 순환시킵니다.
   For I = LBound(ArrProgress) To UBound(ArrProgress)
            '배열에 있는 각 항목들에 대한 초기값을 설정합니다.
         ArrProgress(I) = "Initial value" & I
         ProgressBar1.Value = I
   Next I
   ProgressBar1.Visible = False
   ProgressBar1.Value = ProgressBar1.Min
End Sub

