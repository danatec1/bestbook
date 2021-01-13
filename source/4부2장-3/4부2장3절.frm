VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3435
   ClientTop       =   2640
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      LargeChange     =   10
      Left            =   840
      Max             =   500
      Min             =   -500
      TabIndex        =   2
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '단일 고정
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "현재값 :"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub HScroll1_Change()
    Label2.Caption = Str(HScroll1.Value)
End Sub

