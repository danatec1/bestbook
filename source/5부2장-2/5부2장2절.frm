VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmTestCoolbar 
   BorderStyle     =   1  '단일 고정
   Caption         =   "쿨바 테스트"
   ClientHeight    =   2580
   ClientLeft      =   2835
   ClientTop       =   2940
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6585
   Begin VB.TextBox txtResult 
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "결   과"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   1588
      BandCount       =   4
      _CBWidth        =   6495
      _CBHeight       =   900
      _Version        =   "6.0.8169"
      Caption1        =   "Border 변경"
      Child1          =   "cmdChgBdr"
      MinHeight1      =   405
      Width1          =   975
      NewRow1         =   0   'False
      Caption2        =   "Enabled 변경"
      Child2          =   "cmdChgEnabled"
      MinHeight2      =   405
      Width2          =   1170
      NewRow2         =   0   'False
      Caption3        =   "Visible 변경"
      Child3          =   "cmdChgVisible"
      MinHeight3      =   405
      Width3          =   1815
      NewRow3         =   -1  'True
      Caption4        =   "Caption 변경"
      Child4          =   "txtChgCmdCaption"
      MinHeight4      =   405
      Width4          =   1335
      NewRow4         =   0   'False
      Begin VB.TextBox txtChgCmdCaption 
         Height          =   405
         Left            =   3135
         TabIndex        =   5
         Top             =   465
         Width           =   3270
      End
      Begin VB.CommandButton cmdChgVisible 
         Caption         =   "Visible 속성 바꾸기"
         Height          =   405
         Left            =   1215
         TabIndex        =   3
         Top             =   465
         Width           =   570
      End
      Begin VB.CommandButton cmdChgEnabled 
         Caption         =   "텍스트박스 Enabled 바꾸기"
         Height          =   405
         Left            =   6465
         TabIndex        =   2
         Top             =   30
         Width           =   90
      End
      Begin VB.CommandButton cmdChgBdr 
         Caption         =   "텍스트박스 BorderStyle 바꾸기"
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   30
         Width           =   3870
      End
   End
End
Attribute VB_Name = "frmTestCoolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChgBdr_Click()
    If txtResult.BorderStyle = 1 Then '결과 텍스트박스 단일 고정 형이면
        txtResult.BorderStyle = 0 '없음으로 변경
    Else '없음이면
        txtResult.BorderStyle = 1 '단일고정으로 변경
    End If
End Sub

Private Sub cmdChgEnabled_Click()
    
    If txtResult.Enabled = True Then '텍스트박스가 Enabled이면
        txtResult.Enabled = False 'False로
    Else 'False이면
        txtResult.Enabled = True 'False로
    End If
    
    If cmdResult.Enabled = True Then '커맨드버튼이 Enabled이면
        cmdResult.Enabled = False 'False로
    Else 'False이면
        cmdResult.Enabled = True 'True로
    End If
    
End Sub

Private Sub cmdChgVisible_Click()

    If txtResult.Visible = True Then  '텍스트박스가 Visible이면
        txtResult.Visible = False 'False로
    Else 'False이면
        txtResult.Visible = True 'False로
    End If
    
    If cmdResult.Visible = True Then '커맨드버튼이 Visible이면
        cmdResult.Visible = False 'False로
    Else 'False이면
        cmdResult.Visible = True 'True로
    End If

End Sub

Private Sub cmdResult_Click()
    Dim i As Integer '메세지 박스의 리턴값을 받을 정수
    i = MsgBox("정말 종료하시겠습니까?", vbInformation Or vbOKCancel, "종료 확인")
    If i = vbOK Then 'OK버튼을 누르면
        End '종료
    End If
End Sub

Private Sub coolTest_Resize()
    '쿨바 컨트롤의 크기가 변경될 때
    frmTestCoolbar.Height = coolTest.Height + 1100
                    '폼의 크기를 쿨바보다 1100크게
    '결과 텍스트박스와 커맨드버튼의 위치 결정
    txtResult.Top = coolTest.Top + coolTest.Height + 200
    cmdResult.Top = coolTest.Top + coolTest.Height + 200
End Sub

Private Sub txtChgCmdCaption_Change()
    cmdResult.Caption = txtChgCmdCaption.Text
        '텍스트박스의 문자열을 커맨드박스의 Caption으로
End Sub

