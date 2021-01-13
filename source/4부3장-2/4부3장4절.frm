VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "그래픽 뷰어"
   ClientHeight    =   4260
   ClientLeft      =   3000
   ClientTop       =   2340
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5850
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      Height          =   1350
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2415
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Combo1.AddItem "아이콘 파일(*.ico)", 0
   Combo1.AddItem "그림 파일(*.pcx,*.jpg,*.bmp)", 1
   Combo1.AddItem "모든 파일(*.*)", 2
   Combo1.ListIndex = 1
End Sub

Private Sub Combo1_Click()
  Select Case Combo1.ListIndex
    Case 0
       File1.Pattern = "*.ico"
    Case 1
       File1.Pattern = "*.pcx;*.bmp;*.jpg"
    Case 2
       File1.Pattern = "*.*"
  End Select
End Sub

Private Sub Dir1_Change()
  File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
  Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
  d$ = File1.Path
  If Right(d$, 1) = "\" Then
     d$ = Left$(d$, 2)
  End If
    Image1.Picture = LoadPicture(d$ & "\" & File1.FileName)
End Sub

