VERSION 5.00
Begin VB.Form frmSecond 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   100
   ScaleMode       =   0  '사용자
   ScaleWidth      =   100
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtTest 
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdUnload 
      Caption         =   "Unload"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "frmSecond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdHide_Click()
    frmSecond.Hide
End Sub

Private Sub cmdUnload_Click()
    Unload Me
End Sub

