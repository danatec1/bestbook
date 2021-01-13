VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3555
   ClientLeft      =   3750
   ClientTop       =   2445
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "확   인"
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Top             =   2880
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   "김  경효"
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "권  해오"
      Height          =   375
      Index           =   4
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "임  창성"
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "이  승욘"
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "최  진실"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "엄  정하"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "다음 중 당신이 좋아하는 탤런트를 모두 선택하십시요."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'각 컨트롤배열마다 체크되었는지를 검사하여, 출력변수에 값을 줍니다.
For i = 1 To 6 Step 1
   If Check1(i - 1).Value = Checked Then
       A = Check1(i - 1).Caption
   Else
       A = ""
   End If
       B = B & " " & A
Next i
   MsgBox (B & "씨를 좋아하시는군요."), vbInformation, "알림"
End Sub


