Attribute VB_Name = "ModDelay"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' This timeout is set in the formLoad event to 10 seconds
Private m_timeout As Long                                 ' timeout in msec

Public lngStartTime As Long                                'time in msec
Public lngStartTime2 As Long                               'time in msec

Public Property Get TimeOut() As Integer
    TimeOut = m_timeout
End Property

Public Property Let TimeOut(ByVal vNewValue As Integer)
    m_timeout = vNewValue
End Property

Public Sub StartTimer()
    lngStartTime = timeGetTime()
End Sub

Public Function EndTimer() As Double
    EndTimer = timeGetTime() - lngStartTime
End Function

Public Sub delay(msdelay As Long)
   ' creates delay in ms
   Dim temp As Double
    lngStartTime2 = timeGetTime()
   Do Until (timeGetTime() - lngStartTime2) > (msdelay)
   Loop
End Sub
