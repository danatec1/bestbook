VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.ListBox List1 
      Height          =   1860
      ItemData        =   "5부5장1-1절.frx":0000
      Left            =   240
      List            =   "5부5장1-1절.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    List1.AddItem "그 때 또 다시 - 임 창정 "
    List1.AddItem "그녀와의 이별 - 김 현정 "
    List1.AddItem "믿지않는 전설 - 박 건후 "
    List1.AddItem "아뿔싸 - 윤 대관 "
End Sub

Private Sub List1_DblClick()
   Label1.Caption = List1.Text
End Sub

