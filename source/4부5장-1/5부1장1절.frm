VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3690
   ClientLeft      =   3000
   ClientTop       =   2550
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   3690
   ScaleWidth      =   5940
   Begin VB.CommandButton Command1 
      Caption         =   "평 균"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1440
      TabIndex        =   2
      Top             =   2400
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3201
      _Version        =   393216
      Rows            =   4
      Cols            =   5
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "점 수 :"
      Height          =   180
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Col = 1: MSFlexGrid1.Text = "국 어"
    MSFlexGrid1.Col = 2: MSFlexGrid1.Text = "영 어"
    MSFlexGrid1.Col = 3: MSFlexGrid1.Text = "수 학"
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 1: MSFlexGrid1.Text = "이기영"
    MSFlexGrid1.Row = 2: MSFlexGrid1.Text = "김동귀"
    MSFlexGrid1.Row = 3: MSFlexGrid1.Text = "조준상"
    MSFlexGrid1.Col = 1: MSFlexGrid1.Row = 1
End Sub

Private Sub MSFlexGrid1_Click()
   Text1.SetFocus
   Text1.Text = MSFlexGrid1.Text
End Sub

Private Sub Text1_Change()
  MSFlexGrid1.Text = Text1.Text
End Sub

Private Sub Command1_Click()
     For i = 1 To 3
         MSFlexGrid1.Row = i
         Sum = 0
         For j = 1 To 3
             MSFlexGrid1.Col = j
             Cell = Val(MSFlexGrid1.Text)
             Sum = Sum + Cell
         Next j
             MSFlexGrid1.Col = 4: MSFlexGrid1.Text = Sum / 3
     Next i
End Sub

