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
   Begin VB.CommandButton cmdDate 
      Caption         =   "Date"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdTime 
      Caption         =   "Time"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdNow 
      Caption         =   "Now"
      Height          =   495
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtDate 
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox txtTime 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtNow 
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDate, MyNow, MyTime

Private Sub cmdDate_Click()
    txtDate.Text = Date
End Sub

Private Sub cmdNow_Click()
    txtNow.Text = Now
End Sub

Private Sub cmdTime_Click()
    txtTime = Time
End Sub

