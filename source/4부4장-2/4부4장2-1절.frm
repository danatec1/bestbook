VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3645
   ClientTop       =   2745
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      Index           =   2
      LargeChange     =   10
      Left            =   4080
      Max             =   255
      TabIndex        =   6
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      Index           =   1
      LargeChange     =   10
      Left            =   3480
      Max             =   255
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1935
      Index           =   0
      LargeChange     =   10
      Left            =   2880
      Max             =   255
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  '단색
      Height          =   1815
      Left            =   240
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "청색"
      Height          =   180
      Left            =   4080
      TabIndex        =   3
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "녹색"
      Height          =   180
      Left            =   3480
      TabIndex        =   2
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "적색"
      Height          =   180
      Left            =   2880
      TabIndex        =   1
      Top             =   2640
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub VScroll1_Change(Index As Integer)
  a = "RGB("
  For i = 0 To 2
    a = a & VScroll1(i).Value
    If i = 2 Then
      a = a & ")"
    Else
      a = a & ","
    End If
  Next i
    Label4.Caption = a
    Shape1.FillColor = RGB(VScroll1(0).Value, VScroll1(1).Value, VScroll1(2).Value)
End Sub

