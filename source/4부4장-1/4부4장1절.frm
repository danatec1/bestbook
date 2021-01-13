VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3645
   ClientTop       =   2850
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "폰트변경"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
    On Error GoTo Font_Error
  
    CommonDialog1.Flags = cdlCFBoth
    CommonDialog1.ShowFont
    Label1.FontName = CommonDialog1.FontName
    Label1.FontSize = CommonDialog1.FontSize
    Label1.FontBold = CommonDialog1.FontBold
    Label1.FontItalic = CommonDialog1.FontItalic
    Label1.Caption = Text1.Text
Font_Error:

End Sub

