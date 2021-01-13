VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDataEnv 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows 기본값
   Begin VB.CommandButton Command1 
      Caption         =   "보고서 출력"
      Height          =   405
      Left            =   3630
      TabIndex        =   2
      Top             =   3390
      Width           =   1425
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "닫 기"
      Height          =   405
      Left            =   5160
      TabIndex        =   1
      Top             =   3390
      Width           =   1275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Bindings        =   "frmDataEnv.frx":0000
      Height          =   3195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   4
      AllowUserResizing=   1
      DataMember      =   "Command1"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0)._NumMapCols=   3
      _Band(0)._MapCol(0)._Name=   "우편번호"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "동명"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "전체주소"
      _Band(0)._MapCol(2)._RSIndex=   2
   End
End
Attribute VB_Name = "frmDataEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    
    '프로그램을 종료
    End
    
End Sub

Private Sub Command1_Click()
    DataReport1.Show
End Sub
