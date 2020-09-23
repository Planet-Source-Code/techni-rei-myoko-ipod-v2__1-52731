Attribute VB_Name = "MouseWheel"
Option Explicit

Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type Msg
    hWnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Function ScrollMoved(Optional hWnd As Long) As Long
    Dim amsg As Msg
    GetMessage amsg, hWnd, 0, 0
    DispatchMessage amsg
    If amsg.message = 522 Then ScrollMoved = amsg.wParam / 65536
End Function
