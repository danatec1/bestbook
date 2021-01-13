VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton cmdGetTime 
      Caption         =   "GetTime"
      Height          =   615
      Left            =   3480
      TabIndex        =   1
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label lblTime 
      Alignment       =   2  '가운데 맞춤
      BorderStyle     =   1  '단일 고정
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label lblSubject 
      Alignment       =   2  '가운데 맞춤
      BorderStyle     =   1  '단일 고정
      Caption         =   "진행시간 :"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyTimer As TimerClass

Private Sub cmdGetTime_Click()
    Dim Result As Date
    
    Result = MyTimer.GetTime
    
    lblTime.Caption = Hour(Result) & "(시간) " & _
        Minute(Result) & "(분) " & _
        Second(Result) & "(초)경과"
End Sub

Private Sub cmdReset_Click()
    
    Set MyTimer = New TimerClass
    
    MyTimer.Reset
    
End Sub

Private Sub Form_Load()

    Set MyTimer = New TimerClass
    
    MyTimer.Reset
    
End Sub
