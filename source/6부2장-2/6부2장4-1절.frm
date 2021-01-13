VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3615
   ClientTop       =   2955
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Form_Click()
   Dim CX, CY, Radius, Limit
   ScaleMode = 3
   CY = ScaleHeight / 2
   CX = ScaleWidth / 2
   If CX > CY Then Limit = CY Else Limit = CX
   For Radius = 0 To Limit
        Circle (CX, CY), Radius, RGB(Rnd * 255, Rnd * 55, Rnd * 255)
   Next Radius
End Sub

