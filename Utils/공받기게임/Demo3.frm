VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo 3 - Pong!"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9015
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   4200
      Top             =   2040
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   8640
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   240
      Top             =   2160
      Width           =   135
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4320
      Shape           =   3  'Circle
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Demo 3 - PONG!"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   $"Demo3.frx":0000
      BeginProperty Font 
         Name            =   "Radical"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim balldir As Integer
Dim youscore As Integer, compscore As Integer

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then Shape2.Top = Shape2.Top - 250
If KeyCode = vbKeyDown Then Shape2.Top = Shape2.Top + 250
If Shape2.Top < 0 Then Shape2.Top = 0
If Shape2.Top > 4560 Then Shape2.Top = 4560
End Sub

Private Sub Form_Load()
Randomize Timer
Shape1.Top = 2520
Shape1.Left = 4320
balldir = Int(Rnd * 4) + 1
End Sub

Private Sub Timer1_Timer()
'Bounce Ball...
If balldir = 1 Then Shape1.Top = Shape1.Top - 120: Shape1.Left = Shape1.Left - 120
If balldir = 2 Then Shape1.Top = Shape1.Top - 120: Shape1.Left = Shape1.Left + 120
If balldir = 3 Then Shape1.Top = Shape1.Top + 120: Shape1.Left = Shape1.Left + 120
If balldir = 4 Then Shape1.Top = Shape1.Top + 120: Shape1.Left = Shape1.Left - 120
'Bounce Ball of Walls...
If Shape1.Top <= 0 And balldir = 1 Then balldir = 4
If Shape1.Top <= 0 And balldir = 2 Then balldir = 3
If Shape1.Top >= 5280 And balldir = 3 Then balldir = 2
If Shape1.Top >= 5280 And balldir = 4 Then balldir = 1
'If Ball goes out your side then...
If Shape1.Left <= 0 Then
    compscore = compscore + 1
    MsgBox "You Lose!" & Chr$(13) & "Scores:" & Chr$(13) & "You=" & youscore & " , Computer=" & compscore, vbOKOnly, "You lose"
    Form_Load
End If
'If Ball goes out computers side then...
If Shape1.Left >= 8760 Then
    youscore = youscore + 1
    MsgBox "You Win!" & Chr$(13) & "Scores Board:" & Chr$(13) & "You= " & youscore & " , Computer= " & compscore, vbOKOnly, "You lose"
    Form_Load
End If
'Check if ball hits your bat...
If Shape1.Left >= 240 And Shape1.Left <= 360 Then
    If Shape1.Top - 120 >= Shape2.Top And Shape1.Top - 120 <= Shape2.Top + 975 Then
        If balldir = 1 Then balldir = 2
        If balldir = 4 Then balldir = 3
    End If
End If
'Check if ball hits Computers bat...
If Shape1.Left >= 8520 And Shape1.Left <= 8880 Then
    If Shape1.Top - 120 >= Shape3.Top And Shape1.Top - 120 <= Shape3.Top + 975 Then
        If balldir = 2 Then balldir = 1
        If balldir = 3 Then balldir = 4
    End If
End If
'Move computers bat up and down then check if it goes out the top or bottom of the screen...
If Shape1.Top + 120 <= Shape3.Top + Shape3.Height / 2 Then Shape3.Top = Shape3.Top - 120
If Shape1.Top + 120 >= Shape3.Top + Shape3.Height / 2 Then Shape3.Top = Shape3.Top + 120
If Shape3.Top < 0 Then Shape3.Top = 0
If Shape3.Top > 4560 Then Shape3.Top = 4560
End Sub
