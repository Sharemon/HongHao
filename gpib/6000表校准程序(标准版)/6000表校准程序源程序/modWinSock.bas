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
            err.Raise 999, skt, "��ȡ��ʱ�� ��ʱ = " & str$(m_timeout) & "����"
            FrmMain.LblErr.Text = "��ȡ��ʱ�� ��ʱ = " & str$(m_timeout) & "����"
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
        FrmMain.LblErr.Text = IIf(Lang, "Waring:overload indication,please change the RANGE!", "���棺���ݹ��أ���������̣�")
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
            IO_Status = IIf(Lang, "Closed", "�ر�")
        Case sckOpen                            '1
            IO_Status = IIf(Lang, "Open", "��")
        Case sckListening                       '2
            IO_Status = IIf(Lang, "Listening", "����")
        Case sckConnectionPending               '3
            IO_Status = IIf(Lang, "Connection Pending", "���ӹ���")
        Case sckResolvingHost                   '4
            IO_Status = IIf(Lang, "Resolving Host", "��������")
        Case sckHostResolved                    '5
            IO_Status = IIf(Lang, "Host Resolved", "��ʶ������")
        Case sckConnecting                      '6
            IO_Status = IIf(Lang, "Connecting", "��������")
        Case sckConnected                       '7
            IO_Status = IIf(Lang, "Connected", "������")
        Case sckClosing                         '8
            IO_Status = IIf(Lang, "Closing", "���ڹر�")
        Case sckError                           '9
            IO_Status = IIf(Lang, "Error", "����")
        Case Else
            IO_Status = IIf(Lang, "ERROR:" & vbCrLf & "Unknown Connection Status", "����:" & vbCrLf & "δ֪����״̬")
    End Select
    
    IO_Status = IIf(Lang, "Connection Status: ", "����״̬:") & IO_Status
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
            IO_Protocol = "ERROR:" & vbCrLf & "δ֪Э��"
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
