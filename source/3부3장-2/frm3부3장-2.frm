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
   Begin VB.CommandButton cmdText 
      Caption         =   "Text변경 테스트"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdBorderStyle 
      Caption         =   "BorderStyle 테스트"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton cmdVisible 
      Caption         =   "Visible 테스트"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdEnabled 
      Caption         =   "Enabled 테스트"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txtResult 
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblResult 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const EnabledTest = 0
Const VisibleTest = 1
Const BorderStyleTest = 2
Const TextTest = 3

Public Sub ChangeStatus(Task As Integer) '작업의 종류를 매개변수로 받음
    
    Select Case Task
        Case EnabledTest 'Enabled 의 상태를 바꾼다
            
            If lblResult.Enabled = False Then '레이블의 현재 상태를 점검
                lblResult.Enabled = True
            Else
                lblResult.Enabled = False
            End If
            
            If txtResult.Enabled = False Then '텍스트박스의 현재 상태를 점검
                txtResult.Enabled = True
            Else
                txtResult.Enabled = False
            End If
            
        Case VisibleTest 'Visible 의 상태를 바꾼다
            
            If lblResult.Visible = True Then '레이블의 현재 상태를 점검
                lblResult.Visible = False
            Else
                lblResult.Visible = True
            End If
            
            If txtResult.Visible = True Then '텍스트박스의 현재 상태를 점검
                txtResult.Visible = False
            Else
                txtResult.Visible = True
            End If
        
        Case BorderStyleTest 'BorderStyle 을 바꾼다
        
            If lblResult.BorderStyle = 0 Then '레이블의 현재 상태를 점검
                lblResult.BorderStyle = 1 '1-단일 고정
            Else
                lblResult.BorderStyle = 0 '0-없음
            End If
            
            If txtResult.BorderStyle = 0 Then '텍스트박스의 현재 상태를 점검
                txtResult.BorderStyle = 1 '1-단일 고정
            Else
                txtResult.BorderStyle = 0 '0-없음
            End If
        
        Case TextTest '레이블의 경우는 Caption을, 텍스트박스의 경우는 Text를 바꾼다
            
            If lblResult.Caption = "테스트 레이블" Then '레이블의 현재 상태를 점검
                lblResult.Caption = ""
            Else
                lblResult.Caption = "테스트 레이블"
            End If
            
            If txtResult.Text = "테스트 에디트" Then '텍스트박스의 현재 상태를 점검
                txtResult.Text = ""
            Else
                txtResult.Text = "테스트 에디트"
            End If
            
    End Select
End Sub


Private Sub cmdEnabled_Click()
    ChangeStatus EnabledTest 'Enabled속성을 바꾼다
End Sub

Private Sub cmdVisible_Click()
    Call ChangeStatus(VisibleTest)  'Visible속성을 바꾼다
End Sub

Private Sub cmdBorderStyle_Click()
    ChangeStatus BorderStyleTest 'BorderStyle을 바꾼다
End Sub

Private Sub cmdText_Click()
    Call ChangeStatus(TextTest)  'Text를 바꾼다
End Sub


