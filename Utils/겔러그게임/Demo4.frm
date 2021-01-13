VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demo 4 - Space Invaders"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   600
      Top             =   5640
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   5640
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   2760
      Picture         =   "Demo4.frx":0000
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   5
      Left            =   840
      Picture         =   "Demo4.frx":0442
      Top             =   2160
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   4
      Left            =   2400
      Picture         =   "Demo4.frx":0884
      Top             =   2280
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   3
      Left            =   2040
      Picture         =   "Demo4.frx":0CC6
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   2
      Left            =   3360
      Picture         =   "Demo4.frx":1108
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   1
      Left            =   4080
      Picture         =   "Demo4.frx":154A
      Top             =   1080
      Width           =   480
   End
   Begin VB.Image Alien 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "Demo4.frx":198C
      Top             =   360
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4440
      X2              =   4440
      Y1              =   5280
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   8
      X1              =   2760
      X2              =   2760
      Y1              =   5760
      Y2              =   5640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   7
      X1              =   480
      X2              =   480
      Y1              =   3600
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   6
      X1              =   3240
      X2              =   3240
      Y1              =   2640
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   1200
      X2              =   1200
      Y1              =   4800
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   4
      X1              =   4320
      X2              =   4320
      Y1              =   1080
      Y2              =   960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   3720
      X2              =   3720
      Y1              =   4800
      Y2              =   4680
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   1560
      X2              =   1560
      Y1              =   2280
      Y2              =   2160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   2400
      X2              =   2400
      Y1              =   840
      Y2              =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   360
      X2              =   360
      Y1              =   480
      Y2              =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2160
      Picture         =   "Demo4.frx":1DCE
      Top             =   5280
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   4  'Dash-Dot
      FillStyle       =   0  'Solid
      Height          =   5775
      Left            =   120
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
If Image1.Picture <> Image2.Picture Then
    If KeyAscii = 44 Then Image1.Left = Image1.Left - 100
    If KeyAscii = 46 Then Image1.Left = Image1.Left + 100
    If Image1.Left < 240 Then Image1.Left = 240
    If Image1.Left > 4080 Then Image1.Left = 4080
    If Line2.Visible = False And KeyAscii = 32 Then
        Line2.Visible = True
        Line2.X1 = Image1.Left + 230
        Line2.X2 = Image1.Left + 230
        Line2.Y1 = Image1.Top - 360
        Line2.Y2 = Image1.Top
    End If
End If
End Sub

Private Sub Form_Load()
Randomize Timer
End Sub

Private Sub Timer1_Timer()
For x = 0 To 8
    Line1(x).Y1 = Line1(x).Y1 + 75
    Line1(x).Y2 = Line1(x).Y2 + 75
    If Line1(x).Y2 > 5760 Then Line1(x).Y2 = 240: Line1(x).Y1 = 120
Next x
End Sub

Private Sub Timer2_Timer()
' If your bullet is shooting then move it up...
Line2.Y1 = Line2.Y1 - 250
Line2.Y2 = Line2.Y2 - 250
'Check if bullet hits alien or hits the top of the screen
If Line2.Y1 < 120 Then Line2.Visible = False
For x = 0 To 5
    If Line2.X1 >= Alien(x).Left And Line2.X2 <= Alien(x).Left + 480 Then
        If Line2.Y1 >= Alien(x).Top And Line2.Y2 <= Alien(x).Top + 480 Then
            Alien(x).Visible = False
            Line2.Visible = False
            aliens = aliens + 1
            If aliens = 6 Then MsgBox "Congratulations!" & Chr$(13) & "You have finished Demo 4!", vbOKOnly, "Congratulations!": End
        End If
    End If
    'Move aliens down then left or right...
    Alien(x).Top = Alien(x).Top + 50
    If Int(Rnd * 2) = 0 Then Alien(x).Left = Alien(x).Left + 150 Else Alien(x).Left = Alien(x).Left - 150
    'Keep aliens away from walls and if it goes out the bottom return it to the top...
    If Alien(x).Top > 5400 Then Alien(x).Top = 120: Alien(x).Left = Int(Rnd * Me.Width - 240) + 120
    If Alien(x).Left < 120 Then Alien(x).Left = Alien(x).Left + 150
    If Alien(x).Left > 4200 Then Alien(x).Left = Alien(x).Left - 150
    'Check if aliens Collide with your Ship...
    If Alien(x).Left >= Image1.Left And Alien(x).Left <= Image1.Left + 480 Then
        If Alien(x).Top + 480 >= 5280 And Alien(x).Top + 480 <= 5760 Then
            Image1.Picture = Image2.Picture
        End If
    End If
Next x
End Sub
