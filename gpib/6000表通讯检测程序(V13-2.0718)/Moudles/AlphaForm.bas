Attribute VB_Name = "AlphaForm"
Option Explicit
'透明接口调用
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA As Long = &H2
Private Const WS_EX_LAYERED As Long = &H80000

Public Sub SetFormToAlpha(hwnd As Long, lngAlpha As Long)
    Dim tmpLog As Long
    If hwnd = 0 Then Exit Sub
    If lngAlpha >= 0 And lngAlpha <= 255 Then
        tmpLog = GetWindowLong(hwnd, GWL_EXSTYLE) '窗口属性
        Call SetWindowLong(hwnd, GWL_EXSTYLE, tmpLog Or WS_EX_LAYERED)
        Call SetLayeredWindowAttributes(hwnd, 0, lngAlpha, LWA_ALPHA)
    End If
End Sub
