VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   DrawMode        =   1  '검정
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Label Label6 
      Caption         =   "6-내부 단색"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "5-대시-점-점"
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "4-대시-점"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "3-점"
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "2-대시"
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "1-단색"
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin VB.Line Line6 
      BorderStyle     =   6  '내부 단색
      DrawMode        =   1  '검정
      X1              =   240
      X2              =   3120
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderStyle     =   5  '대시-점-점
      DrawMode        =   1  '검정
      X1              =   240
      X2              =   3120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderStyle     =   4  '대시-점
      DrawMode        =   1  '검정
      X1              =   240
      X2              =   3120
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line3 
      BorderStyle     =   3  '점
      DrawMode        =   1  '검정
      X1              =   240
      X2              =   3120
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderStyle     =   2  '대시
      DrawMode        =   1  '검정
      X1              =   240
      X2              =   3120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   1  '검정
      X1              =   240
      X2              =   3120
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
