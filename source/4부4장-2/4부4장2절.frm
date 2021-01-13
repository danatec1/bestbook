VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3750
   ClientTop       =   2550
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "글자색 변경"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "배경색 변경"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '단일 고정
      Height          =   735
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
  On Error GoTo Color_Error
  
  CommonDialog1.DialogTitle = "배경색"
  CommonDialog1.Flags = cdlCCRGBInit
  CommonDialog1.ShowColor
  Label1.Caption = Text1.Text
  Label1.BackColor = CommonDialog1.Color

Color_Error:
  
End Sub

Private Sub Command2_Click()
  
  On Error GoTo Color_Error
  
  CommonDialog1.DialogTitle = "글자색"
  CommonDialog1.Flags = cdlCCRGBInit
  CommonDialog1.ShowColor
  Label1.Caption = Text1.Text
  Label1.ForeColor = CommonDialog1.Color

Color_Error:

End Sub

