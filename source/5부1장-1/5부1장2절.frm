VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3600
   ClientLeft      =   3330
   ClientTop       =   2625
   ClientWidth     =   6255
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuFile 
      Caption         =   "����(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "������(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "����(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "����(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu Separator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "�μ�(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu Separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileEnd 
         Caption         =   "����(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "����(&E)"
      Begin VB.Menu mnuEditCut 
         Caption         =   "�߶󳻱�(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "�ٿ��ֱ�(&A)"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "����(&V)"
      Begin VB.Menu mnuViewColor 
         Caption         =   "����ǥ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuViewState 
         Caption         =   "���� ǥ����(&W)"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpInf0 
         Caption         =   "�׸��� ����(&I)"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()

End Sub
