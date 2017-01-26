Attribute VB_Name = "dataDim"
'…Ë÷√œ‡πÿ
Public IOSetting As Integer
Public GPIBnum As Integer
Public GPIBaddr As Integer
Public comport As Integer
Public baudrate As Long
Public databit As Integer
Public stopbit As Integer
Public cr As String
Public comPortOut As Integer
Public baudRateOut As Long
Public dataBitOut As Integer
Public stopBitOut As Integer
Public crOut As String
Public outEn As Boolean
Public IPaddr As String
Public IPport As Long
Public localIPPort As Long
Public multi(2, 11) As Double
Public div(2, 11) As Double
Public modiPara(2, 11, 1) As Double
Public valid(2, 11) As Integer
Public tSteady(2, 11) As Integer
Public aSteady(2, 11, 1) As Double
Public filter(2, 11) As Integer
Public filterDiv As Integer
Public filterCur As Integer
Public selfDef(7) As String
Public selfName(7) As String
Public selfValid(7) As Boolean
Public selfDef2(7) As String
Public selfName2(7) As String
Public selfValid2(7) As Boolean
Public rangeNum As Long
Public zero(2, 3) As String
Public digits(2, 1) As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
'Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
