VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "����"
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   4635
   StartUpPosition =   2  'ȭ�� ���
   Begin VB.CommandButton cmdEqual 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3600
      TabIndex        =   19
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdOff 
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "CLS"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   17
      Top             =   840
      Width           =   855
   End
   Begin VB.CommandButton cmdPoint 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   16
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdSign 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   15
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdPlus 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   14
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdMinus 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton cmdMulti 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   12
      Top             =   1440
      Width           =   855
   End
   Begin VB.CommandButton cmdDivision 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtScreen 
      Alignment       =   1  '������ ����
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   240
      Width           =   3375
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   960
      TabIndex        =   8
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   1800
      TabIndex        =   6
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   735
   End
   Begin VB.CommandButton cmdArray 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'�� ���α׷��� ������ ��������� �� �� �ִ°������α׷��Դϴ�.

'Flag�� ��� ���ڸ� �Է¹��� ������ �ƴ����� �Ǵ��ϴ� �����Դϴ�.
'Continue�� "="�� ������ ����� �ϴ� ������ ��������ڷ� ��� ������
          '�� �������� �Ǵ��ϴ� �����Դϴ�.
'Number�� ���α׷� ���� �ÿ� �ؽ�Ʈ�ڽ����� �Է¹��� ���ڵ��� ����
          '�ϴ� �����Դϴ�.
'Ispoint�� �ʱⰪ�� '0.'�̹Ƿ� �̰��� ���ڷ� �Է��� '0.'���� ���θ� �Ǵ�
          '�ϴ� �����Դϴ�
          
'�����ڰ� �ԷµǾ������� ���������� �����ϴ� ����
Dim Flag As Boolean
'���ڰ��� ���� ������ ����� �̷���������� ���� ���������� �����ϴ� ����
Dim Continue As Boolean
'����ڰ� .�� �Է��ߴ����� ���������� �����ϴ� ����
Dim Ispoint As Boolean
'���ڰ��� �Է¹޴� ���� ���������� ���� ���������� �����ϴ� ����
Dim opStatus As Boolean
'ù��° ���ڰ��� ���� ���ڿ��� �����ϴ� ����
Dim First As String
'��° ���ڰ��� ���� ���ڿ��� �����ϴ� ����
Dim Second As String
'�Ź� ����ڰ� ���� ���ڸ� ó���ϱ����� ����
Dim Number As String
'�����ڸ� �����ϴ� ����
Dim Opt As String
'������ ó���� ���� ����
Dim Sign As String
'�������� �����ϱ� ���� ����
Dim Previous As String
'����� ���������� ����
Dim Result As String

Private Sub cmdArray_Click(Index As Integer)
    
    '�������� ��ư�� ó���� ������ �ϱ����ؼ� ��Ʈ�ѹ迭�� ����߽��ϴ�.
    '���ڹ�ư�� ������ �� ��ư�� �ε������� ���� ���ڷ� ��ȯ���Ѽ�
    '����ڰ� ���� ��ư�� ���� ���ؿɴϴ�.
    Number = CStr(cmdArray(Index).Index)
    
    '��� �����ڸ� ���� �� Flag�� True�� �ٲ�� ���� ���� ���ڸ�
    '�Է¹޴� ���Դϴ�
    If Flag = True Then
    
        txtScreen.Text = ""
        txtScreen.Text = Number
        
    Else
        
        If txtScreen.Text = "0." Then
        
            If (Ispoint) Then
                '�ʱⰪ���� "0."�� �ƴ϶� ����ڰ� �Է��� '0.'�� ����� ó���Դϴ�.
                txtScreen.Text = txtScreen.Text + Number
            Else
                '�ʱⰪ���� "0."�� ����� ó���Դϴ�.
                txtScreen.Text = Number
            End If
            
        Else
            If (CStr(txtScreen.Text) = "0" And Number = "0") Then
                '0000...�� �ԷµǴ� ��� �����޼��� �߻���ŵ�ϴ�.
                MsgBox "�Ҽ����̳� �����ڸ� �Է��ϼ���"
            Else
                '���ڹ�ư�� ���� �� �ؽ�Ʈ�ڽ��� ������ ���ڿ��� �ٲ��
                '�����Ƿ� �� �ڿ� �� ���� ���ڸ� �̹� ����� ���ڿ� �ڿ� ���Դϴ�
                txtScreen.Text = txtScreen.Text + Number
            End If
        End If
    
    End If
    
    Flag = False
    
    '��� ���ڸ� �Է¹��� �غ� �մϴ�
    opStatus = True
    First = txtScreen.Text
    
End Sub

Private Sub cmdClear_Click()
    '�ؽ�Ʈ�ڽ��� ������ �ʱ�ȭ��ŵ�ϴ�
    Flag = False
    txtScreen.Text = "0."
    
    '��� ��������ڸ� �̿��� ����� �ƴ϶�� ���Դϴ�
    Continue = False
    opStatus = False
    
    '"0."�� �ʱⰪ�� ���� ������ �˸��ϴ�
    Ispoint = False
End Sub

Private Sub cmdDivision_Click()
    
    If opStatus = True Then
        Opt = "/"
        '��� �����ڸ� �������Ƿ� ���� ���ڸ� ���� �غ� �մϴ�
        Flag = True
        '��� ������ ������ ��������ڰ� �ִٸ� ��� ��������ڸ� �̿���
        '�����̹Ƿ� Preoperation()�� �̿��ؼ� ������ ��������ڿ� ����
        '����� �����մϴ�
        If (Continue) Then
            Call Preoperation
        End If
    
        '���� ����� Second�� �����ϰ� ������ ����� ���ڴ� First�� �����ϰ� �˴ϴ�
        Second = txtScreen.Text
    
        '������ �� ��������ڿ� ���� ������ ��ӵ� ���̹Ƿ� ���� ���õ�
        '��������ڸ� ���� �����ڷ� �����մϴ�
        Continue = True
        Previous = "/"
    Else
        '�����ڰ� �ι� �̻� ������ ��
        MsgBox "���ڸ� �Է��ϼ���"
    End If
    
    opStatus = False
    
End Sub

Private Sub cmdEqual_Click()
    
    If opStatus = True Then
        If Opt = "+" Then
            '���ڷ� ����Ǿ� �ִ� ���� Double������ ���ڷ� �ٲپ� �����մϴ�
            txtScreen.Text = CDbl(Second) + CDbl(First)
            ElseIf Opt = "-" Then
                txtScreen.Text = CDbl(Second) - CDbl(First)
            ElseIf Opt = "*" Then
                txtScreen.Text = CDbl(Second) * CDbl(First)
            ElseIf Opt = "/" Then
                If First = "0" Then
                '0���� ������ ��쿡 �����޼����� ���ϴ�
                    Beep
                    MsgBox "0���� ���� �� �����ϴ�.", vbOKOnly, "Error"
                    txtScreen.Text = "0."
                Else
                    txtScreen.Text = CDbl(Second) / CDbl(First)
                End If
            End If
    
            '��� ������� -1�� 1������ ���̸� �ؽ�Ʈ�ڽ� ȭ�鿡 "0"�� �ٿ��ݴϴ�
            '�ֳ��ϸ� ���� ��������� 0.6�̶�� �ؽ�Ʈ�ڽ�ȭ�鿡�� .6�̶��
            '���̱� �����Դϴ�
            If (CDbl(txtScreen.Text) > -1 And CDbl(txtScreen.Text) < 1) Then
                txtScreen.Text = "0" + txtScreen.Text
            End If

            '"="������ ������ ���̹Ƿ� �� �̻� ��� ��������ڿ� ���� ������
            '���� �ʽ��ϴ�
            Flag = True
            Second = txtScreen.Text
            Continue = False
            Ispoint = False
    Else
        '�����ڸ� �Է��ϰ� �ٽ� =�����ڸ� �Է��� ����̰ų�
        '=�����ڸ� �ι� �̻� Ŭ���� ���
        If txtScreen.Text = "0." Then
            '���� �Է¹ޱ� ���� �ʱⰪ�� '0.'�� �� =�����ڸ� Ŭ���� ���
            MsgBox "���ڸ� �Է��ϼ���"
        Else
            MsgBox "'=' ���� ���Ŀ��� ���ο� ����� �����Ͻʽÿ�"
        End If
    End If
    opStatus = False

End Sub

Private Sub cmdMinus_Click()

    If opStatus = True Then
        Opt = "-"
        Flag = True
        
        If (Continue) Then
            Call Preoperation
        End If
    
        Second = txtScreen.Text
    
        Continue = True
        Previous = "-"
    Else
        '�����ڰ� �� �� �̻� ��������
        MsgBox "���ڸ� �Է��ϼ���"
    End If
    opStatus = False

End Sub

Private Sub cmdMulti_Click()

    If opStatus = True Then
        Opt = "*"
        Flag = True

        If (Continue) Then Call Preoperation
            
        Second = txtScreen.Text
    
        Continue = True
        Previous = "*"
    Else
        '�����ڰ� �� �� �̻� ��������
        MsgBox "���ڸ� �Է��ϼ���"
    End If
    opStatus = False

End Sub

Private Sub cmdOff_Click()

    'OFF��ư�� ������ �� ���� ���α׷��� �����մϴ�
    Result = MsgBox("���⸦ �����Ͻðڽ��ϱ�?", _
                     vbOKCancel + vbQuestion, "����")
    'Result���� "���⸦ �����Ͻðڽ��ϱ�?"�� ���� "Ȯ��"�Ǵ� "���"��
    '������� ����˴ϴ�
    If Result = vbOK Then  '"Ȯ��"�� ���� ��쿡 ���α׷��� �����մϴ�
        End
    End If
    '"���"�� ���� ��쿡�� �����մϴ�
    
End Sub

Private Sub cmdPlus_Click()

    If opStatus = True Then
        Opt = "+"
        Flag = True

        If (Continue) Then Call Preoperation
        
        Second = txtScreen.Text
       
        Continue = True
        Previous = "+"
    Else
        '�����ڰ� �� �� �̻� ��������
        MsgBox "���ڸ� �Է��ϼ���"
    End If
    opStatus = False

End Sub


Private Sub cmdPoint_Click()

    '�Ҽ����� ���ڷ� ����� ���ڿ� �߰��մϴ�
    If opStatus = True Then
        If txtScreen.Text = "0." Or txtScreen.Text = "" Then
            txtScreen.Text = "0."
        Else
            txtScreen.Text = txtScreen.Text + "."
        End If
    
        '�Ҽ����� ���� �ڿ��� ��� ���ڸ� �Է¹��� �غ� �մϴ�
        Flag = False
        
        First = txtScreen.Text
        ''0.'�� �ʱⰪ�� ���� ���� �ƴ϶� ����ڿ� ���� �Է¹��� ������
        '�˸��ϴ�
        Ispoint = True
        
    Else
        '�Ҽ����� �� �� �̻� ��������
        MsgBox "���ڸ� ���� �Է��ϼ���"
    End If
    opStatus = False

End Sub

Private Sub cmdSign_Click()

    '"-"��ȣ�� ���õ� ��� �̸� ���ڷ� ����� ������ �� �տ� �߰��մϴ�
    If Not (txtScreen.Text = "0.") Then
        If (CDbl(txtScreen.Text) > 0) Then
           txtScreen.Text = "-" + txtScreen.Text
            First = txtScreen
        Else
            MsgBox "�����ڸ� �Է��ϼ���" ''-'��ȣ�� �� �� �̻� ���� ��
        End If
    End If

End Sub

Private Sub Form_Initialize()

    '���� �ε�Ǿ��� ���� �ƹ��� �Է��� ���� �ʾ����Ƿ� ���ڸ� �Է¹��� ��
    '�� �ϰ� �� �������� �ʱ�ȭ�մϴ�.
    Flag = False
    txtScreen.Text = "0."
    Continue = False
    Ispoint = False
    opStatus = False

End Sub

Private Function Preoperation()

    '��� ��������ڸ� ���� ����� �� ��츦 ó���ϴ� �Լ��Դϴ�
    '���� �����ڸ� ���� ������ ����� ��������ڸ� �̿��ؼ� ������ �� ��
    '�� ������� �ؽ�Ʈ�ڽ��� �����ݴϴ�
    Select Case Previous
        Case "+"
            txtScreen.Text = CDbl(Second) + CDbl(First)
        Case "-"
            txtScreen.Text = CDbl(Second) - CDbl(First)
        Case "*"
            txtScreen.Text = CDbl(Second) * CDbl(First)
        Case "/"
            If (First = "0") Then
                Beep
                MsgBox "0���� ���� �� �����ϴ�.", vbOKOnly, "Error"
                txtScreen.Text = "0."
            Else
                txtScreen.Text = CDbl(Second) / CDbl(First)
            End If

    End Select
    
    '��� ������� -1�� 1������ ���̸� �ؽ�Ʈ�ڽ� ȭ�鿡 "0"�� �ٿ��ݴϴ�
    '�ֳ��ϸ� ���� ��������� 0.6�̶�� �ؽ�Ʈ�ڽ�ȭ�鿡�� .6�̶��
    '���̱� �����Դϴ�
    If (CDbl(txtScreen.Text) > -1 And CDbl(txtScreen.Text) < 1) Then
       txtScreen.Text = "0" + txtScreen.Text
    End If
    
End Function
