VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPost 
   BorderStyle     =   3  'ũ�� ���� ��ȭ ����
   Caption         =   "�����ȣ"
   ClientHeight    =   2910
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   5775
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   5775
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ݱ�"
      Height          =   300
      Left            =   4616
      TabIndex        =   10
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "���� ��ħ"
      Height          =   300
      Left            =   3462
      TabIndex        =   9
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "����"
      Height          =   300
      Left            =   2308
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "������Ʈ"
      Height          =   300
      Left            =   1154
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "�߰�"
      Height          =   300
      Left            =   0
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtAddr 
      DataField       =   "��ü�ּ�"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtNum 
      DataField       =   "�����ȣ"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtDong 
      DataField       =   "����"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  '�Ʒ� ����
      Height          =   330
      Left            =   0
      Top             =   2580
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=�����ȣ"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "�����ȣ"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select ����,�����ȣ,��ü�ּ� from �����ȣ"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "��ü�ּ�"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "�����ȣ"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  '������ ����
      Caption         =   "����"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   495
   End
End
Attribute VB_Name = "frmPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  '���� ó�� �ڵ带 �ִ� ��ġ�Դϴ�.
  '������ �����Ϸ��� ���� ���� �ּ����� ó���Ͻʽÿ�.
  '������ �������� ���⿡ ������ ó���ϴ� �ڵ带 �߰��Ͻʽÿ�.
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  '�� ���ڵ� ������ ���� ���ڵ� ��ġ�� ǥ���մϴ�.
  datPrimaryRS.Caption = "Record: " + Str(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub cmdAddNew_Click()
  '����ó�� ����
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  
  '����ó�� ����
  On Error GoTo DeleteErr
  
  '���ڵ� �����ϱ�
  atPrimaryRS.Recordset.Delete
  atPrimaryRS.Recordset.MoveNext
  '���ϳ��� ��� ������ ���ڵ�� �̵�
  If atPrimaryRS.Recordset.EOF Then
    atPrimaryRS.MoveLast
  End If
  
  Exit Sub
DeleteErr:
  '���� �޽��� ���
  MsgBox Err.Description

End Sub

Private Sub cmdRefresh_Click()
  
  '����ó�� ����
  On Error GoTo RefreshErr
  
  '������ ��Ʈ���� ���� ����
  datPrimaryRS.Refresh
  Exit Sub

RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  
  '����ó�� ����
  On Error GoTo UpdateErr

  '���ڵ� �����ϱ�
  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
  
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
