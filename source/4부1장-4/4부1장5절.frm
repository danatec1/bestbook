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
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Text2"
      Height          =   180
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Text1"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Width           =   60
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_GotFocus()
    Text1.BackColor = RGB(255, 0, 0)
    Label1.Caption = "Text1 has the focus"
End Sub

Private Sub Text1_LostFocus()
    Text1.BackColor = RGB(0, 0, 225)
    Label1.Caption = "Text1 doesn't have the focus"
End Sub

