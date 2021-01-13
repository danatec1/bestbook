VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3705
   ClientTop       =   3375
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.HScrollBar sc1Blue 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   255
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.HScrollBar sc1Green 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   255
      TabIndex        =   2
      Top             =   1320
      Width           =   1695
   End
   Begin VB.HScrollBar sc1Red 
      Height          =   255
      LargeChange     =   10
      Left            =   2760
      Max             =   255
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.PictureBox picBoard 
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   2235
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Blue"
      Height          =   180
      Left            =   2760
      TabIndex        =   6
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Green"
      Height          =   180
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Red"
      Height          =   180
      Left            =   2760
      TabIndex        =   4
      Top             =   720
      Width           =   330
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChangeColor()
   picBoard.BackColor = RGB(sc1Red.Value, sc1Green.Value, sc1Blue.Value)
End Sub
Private Sub sc1Blue_Change()
  ChangeColor
End Sub
Private Sub sc1Green_Change()
  ChangeColor
End Sub

Private Sub sc1Red_Change()
  ChangeColor
End Sub

