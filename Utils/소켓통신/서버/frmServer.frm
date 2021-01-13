VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "������� ����"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton cmdListen 
      Caption         =   "���"
      Height          =   555
      Left            =   1530
      TabIndex        =   6
      Top             =   3090
      Width           =   1275
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
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Text            =   "����"
      Top             =   90
      Width           =   1545
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   555
      Left            =   4320
      TabIndex        =   2
      Top             =   3090
      Width           =   1245
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "����"
      Height          =   555
      Left            =   2910
      TabIndex        =   1
      Top             =   3090
      Width           =   1305
   End
   Begin VB.TextBox txtBoard 
      Height          =   1815
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  '����
      TabIndex        =   0
      Top             =   630
      Width           =   5385
   End
   Begin MSWinsockLib.Winsock WSocket 
      Left            =   2640
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   10001
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
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdListen_Click()
    
    Select Case WSocket.State
        Case sckListening
            frmServer.Caption = "������� ���� : �����"
        Case connected
            frmServer.Caption = "������� ���� : ������"
        Case Else
            WSocket.Close
            WSocket.Listen
            
            If WSocket.State = sckListening Then
                frmServer.Caption = "������� ���� : �����"
            End If
            
    End Select
    
End Sub

Private Sub cmdSend_Click()
    
    Dim Buffer As String
        
    'Debug.Print WSocket.State
    Buffer = txtHead + ":" + txtSay.Text
    WSocket.SendData Buffer

    txtBoard.Text = txtBoard.Text + Buffer + Chr(13) & Chr(10)
    txtBoard.SelStart = Len(txtBoard.Text)
    txtBoard.SelLength = 0

End Sub

Private Sub Form_Load()
    
    On Error GoTo aaa
    
    WSocket.Bind 10001, "203.255.240.200"
    WSocket.Listen
    Debug.Print WSocket.State
    
    Exit Sub
aaa:

    MsgBox "������ȣ" + Str(Number) + Chr(13) & Chr(10) + Description, vbExclamation, "�����߻�"

End Sub

Private Sub txtSay_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdSend_Click
        txtSay.SelStart = 0
        txtSay.SelLength = 1024
    End If
    
End Sub

Private Sub WSocket_Close()
    frmServer.Caption = "������� ���� : ���� ������ ����Ǿ����ϴ�."
End Sub

Private Sub WSocket_ConnectionRequest(ByVal requestID As Long)
   
    If WSocket.State <> sckClosed Then WSocket.Close
    
    WSocket.Accept requestID
    'WSocket.SendData "������ ������ �����Ǿ����ϴ�."
    frmServer.Caption = "������� ���� : ����"

End Sub

Private Sub WSocket_DataArrival(ByVal bytesTotal As Long)
    
    Dim Buffer As String
    
    frmServer.Caption = "������� ���� [����:" + Str(bytesTotal) + "����Ʈ]"
    WSocket.GetData Buffer, bybyte, 1024
    txtBoard.Text = txtBoard.Text + Buffer + Chr(13) & Chr(10)
    txtBoard.SelStart = Len(txtBoard.Text)
    txtBoard.SelLength = 0
    
End Sub


Private Sub WSocket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    frmServer.Caption = "������� ���� [����:" + Str(bytesSent) + "����Ʈ]"
End Sub
