Attribute VB_Name = "modWinSock"
Option Explicit

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

' This timeout is set in the formLoad event to 10 seconds
Private m_timeout As Long                                 ' timeout in msec

Public RemPort As Long

Global Const APP_CATEGORY = "Agilent\IntuiLink\34410A"

Public lngStartTime As Long                                'time in msec
Public lngStartTime2 As Long                               'time in msec

Public Sub WriteString(skt As CSocketMaster, ByVal cmd As String)
If skt.State = 7 Then
        skt.SendData cmd & vbCrLf
        OnOff False, 2
        FrmMain.Timer2.Enabled = True
End If
End Sub

Public Function ReadString(skt As CSocketMaster) As String
Busy = True
On Error GoTo errhdl
    Dim strData As String
    Dim numbBytes As Long
    Dim I As Long
    
    ' determines the number of passes to try to get data
    ' before calling for timeout
    Const iterations As Long = 200
    If m_timeout = 0 Then m_timeout = 200 ' timeout is in msec
    
    ' Uses the timeout value to determine how long to wait
    If (skt.State = sckConnected) Then
        numbBytes = skt.BytesReceived
        DoEvents
        'check repeatedly if there is new data.
        For I = 1 To iterations
        DoEvents
            If skt.BytesReceived > numbBytes Then Exit For
            delay m_timeout / iterations
        Next I

        ' Gets the data and Clears buffer
        skt.GetData strData, vbString
        
        If I > iterations Then
            If Lang = 1 Then
            err.Raise 999, skt, "Timeout occured. Timeout = " & str$(m_timeout) & "msec"
            FrmMain.LblErr.Text = "Timeout occured. Timeout = " & str$(m_timeout) & "msec"
            Else
            err.Raise 999, skt, "读取超时。 超时 = " & str$(m_timeout) & "毫秒"
            FrmMain.LblErr.Text = "读取超时。 超时 = " & str$(m_timeout) & "毫秒"
            End If
            FrmMain.Timer4 = True
        Else
            If Right$(strData, 1) = vbLf Then strData = Left$(strData, Len(strData) - 1)
            ReadString = strData
            
            OnOff False, 1
            FrmMain.Timer3.Enabled = True
        End If
    End If
errhdl: Exit Function
Busy = False
End Function

Public Function ReadNumber(skt As CSocketMaster) As Variant
    Dim strTemp As String
    strTemp = ReadString(skt)
    
    ReadNumber = IIf(RANGe = 0.1, (strTemp) * 1000, Val(strTemp))
    
    If ReadNumber = 9.9E+37 Then
        FrmMain.Timer4 = True
        FrmMain.LblErr.Text = IIf(Lang, "Waring:overload indication,please change the RANGE!", "警告：数据过载，请更换量程！")
    Else
        FrmMain.LblErr.Text = IO_Status(Winsock1) & "..."
    End If
End Function

Public Property Get TimeOut() As Integer
    TimeOut = m_timeout
End Property

Public Property Let TimeOut(ByVal vNewValue As Integer)
    m_timeout = vNewValue
End Property


Public Function IO_Status(skt As CSocketMaster) As String
    Dim status As Long
    
    ' Get the status code
    status = skt.State
    
    ' Return a text message for the status code
    Select Case status
        Case sckClosed                          '0
            IO_Status = IIf(Lang, "Closed", "关闭")
        Case sckOpen                            '1
            IO_Status = IIf(Lang, "Open", "打开")
        Case sckListening                       '2
            IO_Status = IIf(Lang, "Listening", "侦听")
        Case sckConnectionPending               '3
            IO_Status = IIf(Lang, "Connection Pending", "连接挂起")
        Case sckResolvingHost                   '4
            IO_Status = IIf(Lang, "Resolving Host", "解析域名")
        Case sckHostResolved                    '5
            IO_Status = IIf(Lang, "Host Resolved", "已识别主机")
        Case sckConnecting                      '6
            IO_Status = IIf(Lang, "Connecting", "正在连接")
        Case sckConnected                       '7
            IO_Status = IIf(Lang, "Connected", "已连接")
        Case sckClosing                         '8
            IO_Status = IIf(Lang, "Closing", "正在关闭")
        Case sckError                           '9
            IO_Status = IIf(Lang, "Error", "错误")
        Case Else
            IO_Status = IIf(Lang, "ERROR:" & vbCrLf & "Unknown Connection Status", "错误:" & vbCrLf & "未知连接状态")
    End Select
    
    IO_Status = IIf(Lang, "Connection Status: ", "连接状态:") & IO_Status
End Function

Public Function IO_Protocol(skt As CSocketMaster) As String
    Dim status As Long
    
    ' Get the status code
    status = skt.Protocol
    
    ' Return a text message for the status code
    Select Case status
        Case sckTCPProtocol                          '0
            IO_Protocol = "TCP/IP"
        Case sckUDPProtocol                             '1
            IO_Protocol = "UDP"
        Case Else
            IO_Protocol = "ERROR:" & vbCrLf & "未知协议"
    End Select
    
    'IO_Protocol = "Protocol: " & IO_Protocol
End Function



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

Public Sub OnOff(off As Boolean, Index As Integer)
Select Case Index
Case 0
FrmMain.ShpComOn.Visible = Not off
FrmMain.ShpComOff.Visible = off
Case 1
FrmMain.ShpRevOn.Visible = Not off
FrmMain.ShpRevOff.Visible = off
Case 2
FrmMain.ShpSendOn.Visible = Not off
FrmMain.ShpSendOff.Visible = off
Case 3
FrmMain.ShpErrOn.Visible = Not off
FrmMain.ShpErrOff.Visible = off
End Select
End Sub
