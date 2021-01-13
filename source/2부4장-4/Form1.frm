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
   Begin VB.CommandButton cmdKeyUp 
      Caption         =   "cmdKeyUp"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdKeyDown 
      Caption         =   "cmdKeyDown"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdKeyPress 
      Caption         =   "cmdKeyPress"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtKeyUp 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtKeyDown 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtKeyPress 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
