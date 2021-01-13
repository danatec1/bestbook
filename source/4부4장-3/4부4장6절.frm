VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1245
   ClientLeft      =   3540
   ClientTop       =   1605
   ClientWidth     =   3825
   LinkTopic       =   "Form1"
   ScaleHeight     =   1245
   ScaleWidth      =   3825
   Begin VB.CommandButton Command1 
      Caption         =   "비주얼베이직 도움말 보기"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  CommonDialog1.HelpFile = "vb98.chm"
  CommonDialog1.HelpCommand = cdlHelpContext
  CommonDialog1.ShowHelp
End Sub

