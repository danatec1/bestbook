Attribute VB_Name = "Module1"
'API함수의 선언부
Public Declare Function FlashWindow Lib "user32" _
    (ByVal hwnd As Long, ByVal bInvert As Long) As Long

