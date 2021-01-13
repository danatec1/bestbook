VERSION 5.00
Begin VB.UserControl MyControl 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.TextBox txtValueResult 
      Height          =   525
      Left            =   1950
      TabIndex        =   3
      Top             =   840
      Width           =   1245
   End
   Begin VB.TextBox txtvalue2 
      Height          =   525
      Left            =   360
      TabIndex        =   2
      Top             =   1290
      Width           =   1245
   End
   Begin VB.TextBox txtValue1 
      Height          =   525
      Left            =   390
      TabIndex        =   1
      Top             =   480
      Width           =   1245
   End
   Begin VB.CommandButton cmdGetValue 
      Caption         =   "Get Value"
      Height          =   525
      Left            =   2370
      TabIndex        =   0
      Top             =   2190
      Width           =   1935
   End
End
Attribute VB_Name = "MyControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'�⺻ �Ӽ� ��:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_ScrollBars = 0
'�Ӽ� ����:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
Dim m_ScrollBars As Integer
'�̺�Ʈ ����:
Event Click()
Attribute Click.VB_Description = "��ü���� ���콺 ���߸� �����ٰ� ���� �� �߻��մϴ�."
Event DblClick()
Attribute DblClick.VB_Description = "���콺 ���߸� ��ü���� ������ ���� �� �ٽ� ������ ������ �߻��մϴ�."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "��ü�� ��Ŀ���� ���� �� Ű�� ������ �߻��մϴ�."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "ANSIŰ�� ������ ������ ��� �߻��մϴ�."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "��ü�� ��Ŀ���� ���� �� Ű�� ������ �߻��մϴ�."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "��ü�� ��Ŀ���� ���� �� ���콺 ���߸� ������ �߻��մϴ�."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "���콺�� ������ ��� �߻��մϴ�."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "��ü�� ��Ŀ���� ���� �� ���콺 ���߸� ������ �߻��մϴ�."





Private Sub cmdGetValue_Click()
    txtValueResult.Text = txtValue1.Text + txtvalue2.Text
End Sub
'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=8,0,0,0
Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "��ü�� �ؽ�Ʈ�� �׷����� ǥ���ϱ� ���� ���Ǵ� ������ ��ȯ�ϰų� �����մϴ�."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=8,0,0,0
Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "��ü���� �ؽ�Ʈ�� �׷����� ǥ���ϴ� ������� ��ȯ�ϰų� �����մϴ�."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "����ڰ� ���� �̺�Ʈ�� ���� ��ü�� ������ �� �ִ����� ���θ� �����ϴ� ���� ��ȯ�ϰų� �����մϴ�."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=6,0,0,0
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font ��ü�� ��ȯ�մϴ�."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Label�̳� Shape�� ����� ���� �Ǵ� ������������ ���θ� ��Ÿ���ϴ�."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "��ü �׵θ� ������ ��ȯ�ϰų� �����մϴ�."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "��ü�� ������ �ٽ� �׸��� �մϴ�."
     
End Sub

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MemberInfo=7,0,0,0
Public Property Get ScrollBars() As Integer
Attribute ScrollBars.VB_Description = "��ü�� ����/���� ��ũ�� ���븦 ���������� ���θ� ��Ÿ���� ���� ��ȯ�ϰų� �����մϴ�."
    ScrollBars = m_ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As Integer)
    m_ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtValue1,txtValue1,-1,MaxLength
Public Property Get MaxLength1() As Long
Attribute MaxLength1.VB_Description = "��Ʈ�ѿ� �� �� �ִ� ������ �ִ���� ��ȯ�ϰų� �����մϴ�."
    MaxLength1 = txtValue1.MaxLength
End Property

Public Property Let MaxLength1(ByVal New_MaxLength1 As Long)
    txtValue1.MaxLength() = New_MaxLength1
    PropertyChanged "MaxLength1"
End Property

'���! �ּ����� �Ǿ� �ִ� ���� ���� �����ϰų� �������� ���ʽÿ�!
'MappingInfo=txtvalue2,txtvalue2,-1,MaxLength
Public Property Get MaxLength2() As Long
Attribute MaxLength2.VB_Description = "��Ʈ�ѿ� �� �� �ִ� ������ �ִ���� ��ȯ�ϰų� �����մϴ�."
    MaxLength2 = txtvalue2.MaxLength
End Property

Public Property Let MaxLength2(ByVal New_MaxLength2 As Long)
    txtvalue2.MaxLength() = New_MaxLength2
    PropertyChanged "MaxLength2"
End Property

'����� ���� ��Ʈ�ѿ� ���� �Ӽ��� �ʱ�ȭ�մϴ�.
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
    m_ScrollBars = m_def_ScrollBars
End Sub

'����ҿ��� �Ӽ����� �ε��մϴ�.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_ScrollBars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
    txtValue1.MaxLength = PropBag.ReadProperty("MaxLength1", 0)
    txtvalue2.MaxLength = PropBag.ReadProperty("MaxLength2", 0)
End Sub

'�Ӽ����� ����ҿ� ����մϴ�.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("ScrollBars", m_ScrollBars, m_def_ScrollBars)
    Call PropBag.WriteProperty("MaxLength1", txtValue1.MaxLength, 0)
    Call PropBag.WriteProperty("MaxLength2", txtvalue2.MaxLength, 0)
End Sub

