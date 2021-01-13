VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3705
   ClientTop       =   2850
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.ListBox List1 
      Height          =   1140
      ItemData        =   "6부2장1절.frx":0000
      Left            =   2400
      List            =   "6부2장1절.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   360
      TabIndex        =   1
      Top             =   2400
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   360
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    List1.AddItem "조 준상", 0
    List1.AddItem "박 태상", 1
    List1.AddItem "김 석준", 2
    List1.AddItem "신 옥렬", 3
    List1.AddItem "정 혜연", 4
End Sub

Private Sub List1_Click()
  Label1.Caption = List1.List(List1.ListIndex)
  Select Case List1.ListIndex
    Case 0
       Image1.Picture = LoadPicture("C:\My Documents\vb6.0\Example\jjs.bmp")
    Case 1
       Image1.Picture = LoadPicture("C:\My Documents\vb6.0\Example\pts.bmp")
    Case 2
       Image1.Picture = LoadPicture("C:\My Documents\vb6.0\Example\ksj.bmp")
    Case 3
       Image1.Picture = LoadPicture("C:\My Documents\vb6.0\Example\syr.bmp")
    Case 4
       Image1.Picture = LoadPicture("C:\My Documents\vb6.0\Example\jhy.bmp")
  End Select
End Sub

