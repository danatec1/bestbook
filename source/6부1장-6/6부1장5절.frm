VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "출력 예제"
   ClientHeight    =   3195
   ClientLeft      =   3915
   ClientTop       =   3060
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.TextBox Text1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   3360
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
    '프린터 객체를 통하여 출력시 출력이 시작되는 위치를 지정.
    Printer.CurrentX = 100
    Printer.CurrentY = 100
    '출력시 출력문의 글꼴명과 글꼴 크기를 출력하려는 내용으로부터의 이름과 크기로 지정.
    Printer.FontName = Text1.FontName
    Printer.FontSize = Text1.FontSize
    '텍스트박스의 내용을 출력
    Printer.Print Text1.Text
    '프린터로의 출력이 완료됨을 알림.
    Printer.EndDoc
End Sub

