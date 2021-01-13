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
   Begin VB.CommandButton cmdCreate 
      Caption         =   "데이터베이스 만들기"
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2040
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCreate_Click()
    
    '데이터베이스, 테이블, 필드를 저장할 변수를 선언한다
    Dim MyDB As Database
    Dim MyTable As TableDef
    Dim MyField As Field

    '새로운 데이터베이스 파일을 생성한다
    Set MyDB = DBEngine.Workspaces(0).CreateDatabase("D:\가계부.MDB", dbLangKorean, dbEncrypt)
    '새로운 테이블을 생성한다
    Set MyTable = MyDB.CreateTableDef("가계부")

    '테이블에 필드를 추가한다
    Set MyField = MyTable.CreateField("ID", dbLong)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("DATE", dbText, 10)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("SECT", dbText, 20)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("ITEM", dbText, 50)
    MyTable.Fields.Append MyField
    Set MyField = MyTable.CreateField("AMOUNT", dbLong)
    MyTable.Fields.Append MyField
    
    'TableDefs 객체에 테이블을 추가한다
    MyDB.TableDefs.Append MyTable
    MyDB.Close
    DBEngine.Workspaces(0).Close

End Sub


