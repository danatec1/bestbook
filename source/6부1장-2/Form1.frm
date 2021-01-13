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
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Height          =   2220
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub AddFormat(FormatName)
    
    X = Text1.Text
    List1.AddItem FormatName & "    " & Format(X, FormatName)

End Sub


Private Sub Command1_Click()
   
   AddFormat "General Number"
   AddFormat "Currency"
   AddFormat "Percent"
   AddFormat "Fixed"
   AddFormat "Standard"

End Sub
