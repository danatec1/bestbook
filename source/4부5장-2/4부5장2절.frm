VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   3000
   ClientTop       =   2340
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6210
   Begin VB.CommandButton Command2 
      Caption         =   "셀  통 합"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "정   렬"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   2760
      TabIndex        =   2
      Top             =   2625
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1455
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   2566
      _Version        =   393216
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "입 력 :"
      Height          =   180
      Left            =   2040
      TabIndex        =   1
      Top             =   2640
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   MSHFlexGrid1.Rows = 10
   MSHFlexGrid1.Cols = 4
   MSHFlexGrid1.TextMatrix(0, 0) = "도시"
   MSHFlexGrid1.TextMatrix(0, 1) = "생산품"
   MSHFlexGrid1.TextMatrix(0, 2) = "회사"
   MSHFlexGrid1.TextMatrix(0, 3) = "평균임금"
End Sub

Private Sub MSHFlexGrid1_Click()
   Text1.SetFocus
   Text1.Text = MSHFlexGrid1.Text
End Sub

Private Sub Text1_Change()
  MSHFlexGrid1.Text = Text1.Text
End Sub

Private Sub Command1_Click()
  MSHFlexGrid1.Sort = 2
End Sub

Private Sub Command2_Click()
   MSHFlexGrid1.MergeCells = flexMergeFree
   MSHFlexGrid1.MergeRow(3) = True
   MSHFlexGrid1.MergeRow(2) = True
   MSHFlexGrid1.MergeCol(2) = True
   MSHFlexGrid1.MergeCol(1) = True
End Sub

