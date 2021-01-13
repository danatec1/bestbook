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
   Begin VB.CommandButton Command3 
      Caption         =   "종       료"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "4개 원소 자르기"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7개 원소 채우기"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DanyArray() As Date                    '동적 배열의 선언


Private Sub Command1_Click()
ReDim DanyArray(1 To 7) As Date            '배열의 크기를 변경
Dim i, MyDate
MyDate = #11/2/1998#                '배열의 원소를 초기화
For i = 1 To 7
   DanyArray(i) = MyDate
   MyDate = MyDate + 1
Next i
Combo1.Clear                'ComboBox의 목록을 지움
For i = 0 To 6                  'ComboBox의 내용을 배열의 원소로 채움
   Combo1.AddItem Str(DanyArray(i + 1))
Next i
End Sub

Private Sub Command2_Click()
   ReDim Preserve DanyArray(1 To 4) As Date      '배열의 크기를 변경
   Dim i, MyDate
   MyDate = #11/2/1998#
   Combo1.Clear              'ComboBox의 목록을 지움
   For i = 0 To 3                'ComboBox의 내용을 배열의 원소로 채움
         Combo1.AddItem Str(DanyArray(i + 1))
   Next i
End Sub

Private Sub Command3_Click()
End
End Sub


