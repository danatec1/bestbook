VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim FirstTime As Date
Dim LastTime As Date

'시작시간을 초기화하는 함수
Public Function Reset()
    
    '현재의 시간을 저장한다
    FirstTime = Time
    
End Function

'시작시간과 현재 시간의 차를 리턴하는 함수
Public Function GetTime() As Date

    '결과를 저장하는변수
    Dim Result As Date
    
    '현재의 시간을 저장한다
    LastTime = Time
    
    '현재시간과 시작시간의 차를 구한다
    Result = LastTime - FirstTime
    
    '결과값을 리턴한다
    GetTime = Result
    
End Function
