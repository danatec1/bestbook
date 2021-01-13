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
   StartUpPosition =   3  'Windows 기본값
   Begin VB.TextBox txtFormatNow 
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtFormatTime 
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtFormatDate 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton cmdFormatNow 
      Caption         =   "Now"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton cmdFormatTime 
      Caption         =   "Time"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdFormatDate 
      Caption         =   "Date"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDate, MyNow, MyTime

Private Sub cmdFormatDate_Click()
    MyDate = Format(Date, "mm월 dd일")
    txtFormatDate.Text = MyDate
End Sub

Private Sub cmdFormatNow_Click()
    MyNow = Format(Now, "yy-dd-mm")
    txtFormatNow.Text = MyNow
End Sub

Private Sub cmdFormatTime_Click()
    MyTime = Format(Time, "h시m분s초")
    txtFormatTime.Text = MyTime
End Sub


