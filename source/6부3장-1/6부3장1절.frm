VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3720
   ClientLeft      =   3915
   ClientTop       =   2850
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   4680
   Begin VB.OptionButton Option1 
      Caption         =   "둥근 정사각형"
      Height          =   255
      Index           =   5
      Left            =   2640
      TabIndex        =   5
      Top             =   3240
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "둥근 직사각형"
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      Caption         =   "   원"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   3
      Top             =   2280
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "타    원"
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Top             =   3240
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "정사각형"
      Height          =   255
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "직사각형"
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   600
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click(Index As Integer)
       Shape1.Shape = Index
End Sub

