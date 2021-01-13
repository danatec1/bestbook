VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Pict1 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   3675
      TabIndex        =   4
      Top             =   960
      Width           =   3735
   End
   Begin VB.CommandButton cmdFourth 
      Caption         =   "4"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdThird 
      Caption         =   "3"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton cmdSecond 
      Caption         =   "2"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdFirst 
      Caption         =   "1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdFirst_Click()
  Form1.Print "저는 비주얼베이직 6.0입니다."
End Sub

Private Sub cmdFourth_Click()
  Form1.Print "우리는", "비주얼베이직", "6.0입니다."
End Sub

Private Sub cmdSecond_Click()
  Pict1.Print "당신은 비주얼베이직 6.0입니다."
End Sub

Private Sub cmdThird_Click()
  Print "우리는 비주얼베이직 6.0입니다."
End Sub
