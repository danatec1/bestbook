VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "확   인"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "김 경효"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CheckBox Check5 
      Caption         =   "권 해오"
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "임 창성"
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox Check3 
      Caption         =   "이 승욘"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CheckBox Check2 
      Caption         =   "최 진실"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "엄 정하"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "다음 중 당신이 좋아하는 탤런트를 모두 선택하십시요."
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 
   If Check1.Value = Checked Then
       A = "엄정하 "
   Else
       A = ""
   End If
   If Check2.Value = Checked Then
       B = "최진실 "
   Else
       B = ""
   End If
   If Check3.Value = Checked Then
       C = "이승욘 "
   Else
       C = ""
   End If
   If Check4.Value = Checked Then
       D = "임창성 "
   Else
       D = ""
   End If
   If Check5.Value = Checked Then
       E = "권해오 "
   Else
       E = ""
   End If
   If Check6.Value = Checked Then
       F = "김경효 "
   Else
       F = ""
   End If
        '선택을 하지 않았을 경우 메시지 보냄
   If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 And Check4.Value = 0 And Check5.Value = 0 And Check6.Value = 0 Then
      MsgBox "선택을 하지 않았습니다. 선택을 해주세용!!", vbOKOnly, "알림"
   Else
        '선택사항을 나타냄.
   G = A + B + C + D + E + F
   MsgBox (G + "씨를 좋아하시는군요."), vbInformation, "알림"
   End If
End Sub

