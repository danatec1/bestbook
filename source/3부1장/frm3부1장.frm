VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.CheckBox chkStatus 
      Caption         =   "Text Box Visible"
      Height          =   525
      Left            =   390
      TabIndex        =   5
      Top             =   2100
      Width           =   1965
   End
   Begin VB.TextBox txtTest 
      Height          =   525
      Left            =   2970
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2070
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   210
      TabIndex        =   3
      Top             =   1680
      Width           =   4425
   End
   Begin VB.TextBox txtResult 
      Height          =   525
      Left            =   480
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   120
      Width           =   3915
   End
   Begin VB.CommandButton cmdStringTest 
      Caption         =   "StringTest"
      Height          =   525
      Left            =   2970
      TabIndex        =   1
      Top             =   870
      Width           =   1425
   End
   Begin VB.CommandButton cmdNumTest 
      Caption         =   "NumTest"
      Height          =   525
      Left            =   480
      TabIndex        =   0
      Top             =   900
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNumTest_Click()
    txtResult.Text = 123 + 456
End Sub

Private Sub cmdStringTest_Click()
    txtResult.Text = "123" + "456"
End Sub

Private Sub chkStatus_Click()
    Dim VisibleStatus As Boolean
    
    If chkStatus.Value = 1 Then
        VisibleStatus = True
    End If
    
    If chkStatus.Value = 0 Then
        VisibleStatus = False
    End If
    
    txtTest.Visible = VisibleStatus
    
End Sub


