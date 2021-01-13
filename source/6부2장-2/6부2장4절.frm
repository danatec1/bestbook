VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   3090
   ClientTop       =   2340
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
   Dim CX, CY, F, F1, F2, I
   '좌표계의 단위를 픽셀로 지정
   ScaleMode = 3
   CX = ScaleWidth / 2
   CY = ScaleHeight / 2
   DrawWidth = 8
   For I = 50 To 10 Step -2
      F = I / 50
      F1 = 1 - F: F2 = 1 + F
      ForeColor = QBColor(I Mod 15)
      Line (CX * F1, CY * F1)-(CX * F2, CY * F2), , BF
   Next I
   DoEvents
   If CY > CX Then
       DrawWidth = ScaleWidth / 25
   Else
       DrawWidth = ScaleHeight / 25
   End If
   For I = 10 To 50 Step 2
         F = I / 50
         F1 = 1 - F: F2 = 1 + F
         Line (CX * F1, CY)-(CX, CY * F1)
         Line -(CX * F2, CY)
         Line -(CX, CY * F2)
         Line -(CX * F1, CY)
         ForeColor = QBColor(I Mod 15)
    Next I
    DoEvents
End Sub

