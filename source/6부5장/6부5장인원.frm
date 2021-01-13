VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  '단일 고정
   Caption         =   "인원수 입력"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   3450
   StartUpPosition =   1  '소유자 가운데
   Begin VB.CommandButton Command1 
      Caption         =   "지 정"
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "6부5장인원.frx":0000
      Left            =   1080
      List            =   "6부5장인원.frx":0022
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "인원수"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  Form1.Image1.Visible = True
  If Len(Combo1.Text) > 0 Then
     MakeLaddle Combo1.Text
  Else
      MakeLaddle 4
  End If
  Unload Me
  
End Sub
