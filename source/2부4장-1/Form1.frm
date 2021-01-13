VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   2415
   End
   Begin VB.CommandButton cmdTest1 
      Caption         =   "cmdTest1"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    If cmdTest1.Enabled = True Then
        cmdTest1.Enabled = False
    Else
        cmdTest1.Enabled = True
    End If
End Sub

