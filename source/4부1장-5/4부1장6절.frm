VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5790
   ClientLeft      =   2160
   ClientTop       =   1500
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7380
   Begin VB.CommandButton Command2 
      Caption         =   "��   ��"
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ȯ   ��"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Frame Frame4 
      Caption         =   "���� ��1"
      Height          =   1815
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
      Begin VB.OptionButton Option1 
         Caption         =   "���� ����"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "������ ����"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "���� ȸ��"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "��  ��"
      Height          =   3015
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
      Begin VB.CheckBox Check2 
         Caption         =   "������ ����"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   33
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��ũ���̼�"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   32
         Top             =   2160
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "���������� ����"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   31
         Top             =   1800
         Width           =   1695
      End
      Begin VB.CheckBox Check2 
         Caption         =   "�û� ����"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   30
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "�ʱ� �Ͼ�"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   29
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "������ ����"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Chapel"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��  ��"
      Height          =   2175
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   2175
      Begin VB.CheckBox Check1 
         Caption         =   "M.I.S"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "���� ��ȹ��"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   25
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "O.R."
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "���� ����"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "ǰ�� ����"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "�Ż� ����"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2055
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   600
         TabIndex        =   18
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  '������ ����
         Height          =   270
         Left            =   120
         MaxLength       =   6
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   840
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   840
         TabIndex        =   14
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   840
         TabIndex        =   13
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "_"
         Height          =   180
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "�ֹε�Ϲ�ȣ :"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "�� �� :"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "�� �� :"
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�� �� :"
         Height          =   180
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "�� �� :"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  '���� ����
      Height          =   3615
      Left            =   4800
      TabIndex        =   10
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  
  'üũ�� ����������� B������ ����
  For i = 1 To 7 Step 1
     If Check2(i - 1).Value = Checked Then
        B = Check2(i - 1).Caption
        E = Chr(13)
     Else
        B = ""
        E = ""
     End If
        D = D & B & E
   Next i
   For j = 1 To 5 Step 1
      If Check1(j - 1).Value = 1 Then
         A = Check1(j - 1).Caption
         E = Chr(13)
      Else
         A = ""
         E = ""
      End If
         G = G & A & E
   Next j
   For k = 1 To 3
      If Option1(k - 1).Value = True Then
          W = Option1(k - 1).Caption
      Else
          W = ""
      End If
         T = T & W
   Next k
   '������ �޼����ڽ��� ����մϴ�.
  C = MsgBox((D & G & T & Chr(13) & "������ �����ϼ̽��ϴ�. �½��ϱ�?"), vbOKCancel, "�˸�")
       '�޼����ڽ����� '��'�� �����ϸ� Label1�� ���ð������ ǥ���ϰ�,
       ''�ƴϿ�'�� �����ϸ� üũ�� ǥ�ø� �� ����ϴ�.
  If C = vbOK Then
    Label7.Caption = Text3.Text & "   " & Text2.Text & "   " & Text1.Text & Chr(13) & _
                    "---------------------------------" & Chr(13) & _
                    "��û����" & "        " & Chr(13) & _
                    "---------------------------------" & Chr(13) & _
                      D & G & T & "           " & Chr(13) & _
                    "---------------------------------" & Chr(13) & _
                    "            " & "�� ��û���� : " & "����"
   Else
           For i = 1 To 7
             For j = 1 To 5
               For k = 1 To 3
                 Check2(i - 1).Value = 0
                 Check1(j - 1).Value = 0
                 Option1(k - 1).Value = False
               Next k
             Next j
           Next i
        End If
End Sub
Private Sub Command2_Click()

'���α׷��� �����մϴ�.
  End

End Sub
Private Sub Form_Load()
  Frame2.Visible = False
  Frame3.Visible = False
  Frame4.Visible = False
  
End Sub
Private Sub Text5_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
     Text6.SetFocus
  End If
     
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
     Frame2.Visible = True
     Frame3.Visible = True
     Frame4.Visible = True
  End If
     
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
     Text2.SetFocus
  End If
     
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
     Text3.SetFocus
  End If
     
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
     Text4.SetFocus
  End If
     
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = vbKeyReturn Then
     Text5.SetFocus
  End If
     
End Sub

