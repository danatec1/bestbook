VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   Caption         =   "������� Ŭ���̾�Ʈ"
   ClientHeight    =   3900
   ClientLeft      =   10590
   ClientTop       =   2010
   ClientWidth     =   5910
   Icon            =   "frmClient.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3900
   ScaleWidth      =   5910
   Begin MSWinsockLib.Winsock WSocket 
      Left            =   2010
      Top             =   1410
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "203.255.240.200"
      RemotePort      =   10001
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "����"
      Height          =   525
      Left            =   1620
      TabIndex        =   10
      Top             =   3090
      Width           =   1245
   End
   Begin VB.TextBox txtPort 
      Height          =   345
      Left            =   4770
      TabIndex        =   8
      Text            =   "10001"
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtHost 
      Height          =   345
      Left            =   2850
      TabIndex        =   7
      Text            =   "203.255.240.200"
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txtSay 
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2520
      Width           =   5415
   End
   Begin VB.TextBox txtHead 
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Text            =   "Ŭ���̾�Ʈ"
      Top             =   150
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   555
      Left            =   4470
      TabIndex        =   2
      Top             =   3090
      Width           =   1095
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "����"
      Height          =   555
      Left            =   3000
      TabIndex        =   1
      Top             =   3090
      Width           =   1305
   End
   Begin VB.TextBox txtBoard 
      Height          =   1815
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   630
      Width           =   5385
   End
   Begin VB.Label Label3 
      Caption         =   "��Ʈ"
      Height          =   255
      Left            =   4290
      TabIndex        =   9
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "ȣ��Ʈ"
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "���Ӹ�"
      Height          =   255
      Left            =   210
      TabIndex        =   3
      Top             =   180
      Width           =   585
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    
    If WSocket.State <> sckClosed Then WSocket.Close
    
    WSocket.Connect txtHost.Text, Val(txtPort.Text)
    cmdConnect.Enabled = False

End Sub

Private Sub cmdSend_Click()
    
    Dim Buffer As String
        
    Debug.Print "���" + Str(WSocket.State)
    'Debug.Print WSocket.State
    Buffer = txtHead + ":" + txtSay.Text
    WSocket.SendData Buffer

    txtBoard.Text = txtBoard.Text + Buffer + Chr(13) & Chr(10)

End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdSend_Click
        txtSay.SelStart = 0
        txtSay.SelLength = 1024
    End If
    
End Sub

Private Sub WSocket_Close()
    frmClient.Caption = "������� Ŭ���̾�Ʈ : ����"
    cmdConnect.Enabled = True
End Sub

Private Sub WSocket_Connect()
    frmClient.Caption = "������� Ŭ���̾�Ʈ : ����"
End Sub

Private Sub WSocket_DataArrival(ByVal bytesTotal As Long)
    Dim Buffer As String
    
    frmClient.Caption = "������� Ŭ���̾�Ʈ [����:" + Str(bytesTotal) + "����Ʈ]"
    WSocket.GetData Buffer, bybyte, 1024
    txtBoard.Text = txtBoard.Text + Buffer + Chr(13) & Chr(10)
    txtBoard.SelStart = Len(txtBoard.Text)
    txtBoard.SelLength = 0
    
End Sub

Private Sub WSocket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    frmClient.Caption = "������� ���� [����:" + Str(bytesSent) + "����Ʈ]"
End Sub
