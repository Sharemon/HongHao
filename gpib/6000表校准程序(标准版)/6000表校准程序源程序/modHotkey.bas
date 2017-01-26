Attribute VB_Name = "modHotkey"
'Option Explicit

'在窗口结构中为指定的窗口设置信息
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'从指定窗口的结构中取得信息
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'运行指定的进程
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'向系统注册一个指定的热键
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
'取消热键并释放占用的资源
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal ID As Long) As Long
'上述五个API函数是注册系统级热键所必需的，具体实现过程如后文所示

  '热键标志常数,用来判断当键盘按键被按下时是否命中了我们设定的热键
Public Const WM_HOTKEY = &H312
Public Const GWL_WNDPROC = (-4)

'定义系统的热键,原中断标示,被隐藏的项目句柄
Public preWinProc As Long, MyhWnd As Long, uVirtKey() As Long


'热键拦截过程
Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_HOTKEY Then HideDone wParam
    'If Len(CStr(wParam)) < 4 Then Debug.Print "wParam=" & wParam
    '如果不是热键,或者不是我们设置的热键,交还控制权给系统,继续监测热键
    WndProc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function

Public Sub HideDone(Index As Long)
On Error GoTo errhdl
Dim cmdstrs() As String
    Select Case Index
        Case 1           'C
            ReDim cmdstr(1 To 11)
            cmdstrs = Split(GetIni("cmdstr", "dispdown", "255,48,50,59,48,48,59,48,51,13,10", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            FrmMain.MSComm.OutPut = cmdstr
            delay 100
            ReDim cmdstr(1 To 11)
            cmdstrs = Split(GetIni("cmdstr", "dispup", "37,48,48,59,54,53,59,57,56,13,10", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            FrmMain.MSComm.OutPut = cmdstr
            
        Case 2           'ESC
            ReDim cmdstr(1 To 11)
            cmdstrs = Split(GetIni("cmdstr", "rundown", "255,48,50,59,48,48,59,48,50,13,10", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            delay 100
            ReDim cmdstr(1 To 11)
            cmdstrs = Split(GetIni("cmdstr", "runup", "255,48,50,59,48,48,59,49,50,13,10", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            FrmMain.MSComm.OutPut = cmdstr
        Case 3           'Z
            ReDim cmdstr(1 To 19)
            cmdstrs = Split(GetIni("cmdstr", "WK1down", "255,48,50,59,48,48,59,48,56,13,10,37,48,48,59,48,54,59,13", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            FrmMain.MSComm.OutPut = cmdstr
            delay 100
            ReDim cmdstr(1 To 11)
            cmdstrs = Split(GetIni("cmdstr", "WK1up", "255,48,50,59,48,48,59,49,56,13,10", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            FrmMain.MSComm.OutPut = cmdstr
        Case 4            'Enter
            ReDim cmdstr(1 To 11)
            cmdstrs = Split(GetIni("cmdstr", "WK2click", "255,48,50,59,48,48,59,48,57,13,10", App.Path & "\Config.ini"), ",")
            For i = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(i) = Val(cmdstrs(i - 1))
            Next
            FrmMain.MSComm.OutPut = cmdstr
            
        Case 5             'F1
            Dim Nul As Long
            NulOn = (FrmMain.shpNull.FillColor = FrmMain.ShpErrOn.FillColor)
            Nul = IIf(NulOn, 1, 0)
    
            If Nul = 0 Then
                If FrmMain.Check1.Value = 1 Then
                    ReDim cmdstr(1 To 19)
                    cmdstrs = Split(GetIni("cmdstr", "zeroclick1", "255,48,50,59,48,48,59,48,49,13,10,37,48,48,59,48,54,59,13", App.Path & "\Config.ini"), ",")
                    For i = LBound(cmdstr) To UBound(cmdstr)
                        cmdstr(i) = Val(cmdstrs(i - 1))
                    Next
                    FrmMain.MSComm.OutPut = cmdstr
                End If
    
                If Instru = 0 Then
                    WriteString Winsock1, "VOLT:DC:NULL:VAL 0"
                    WriteString Winsock1, "VOLT:DC:NULL:STAT ON"
                    FrmMain.shpNull.FillColor = FrmMain.ShpErrOn.FillColor
                Else
                    'frmMain.MSComm0.OutPut = ":VOLTage:DC:REFerence:STATe On" & vbCr
                    If FrmMain.shpNull.FillColor <> FrmMain.ShpErrOn.FillColor Then Base = Sign
                    FrmMain.shpNull.FillColor = FrmMain.ShpErrOn.FillColor
                End If

            Else
                If FrmMain.Check1.Value = 1 Then
                    ReDim cmdstr(1 To 19)
                    cmdstrs = Split(GetIni("cmdstr", "zeroclick0", "255,48,50,59,48,48,59,48,49,13,10,37,48,48,59,48,55,59,13", App.Path & "\Config.ini"), ",")
                    For i = LBound(cmdstr) To UBound(cmdstr)
                        cmdstr(i) = Val(cmdstrs(i - 1))
                    Next
                    FrmMain.MSComm.OutPut = cmdstr
                End If
                
                If Instru = 0 Then

                    WriteString Winsock1, "VOLT:DC:NULL:STAT Off"

                    FrmMain.shpNull.FillColor = FrmMain.ShpErrOff.FillColor
                Else
                    FrmMain.MSComm0.OutPut = ":VOLTage:DC:REFerence:STATe Off" & vbCr
                    Base = 0
                    FrmMain.shpNull.FillColor = FrmMain.ShpErrOff.FillColor
                End If
            End If
    
            ReadAuto = True
            If Instru = 0 Then
                FrmMain.Timer6.Enabled = True
            Else
                FrmMain.MSComm0.OutPut = ":FETCh?" & vbCr
            End If
        Case 6      'F2
            If FrmMain.Check1.Value = 1 Then
                ReDim cmdstr(1 To 11)
                cmdstrs = Split(GetIni("cmdstr", "runclick", "255,48,50,59,48,48,59,48,50,13,10", App.Path & "\Config.ini"), ",")
                For i = LBound(cmdstr) To UBound(cmdstr)
                    cmdstr(i) = Val(cmdstrs(i - 1))
                Next
                FrmMain.MSComm.OutPut = cmdstr
            End If
            
        Case 7      'F3
            If FrmMain.Check1.Value = 1 Then
                ReDim cmdstr(1 To 11)
                cmdstrs = Split(GetIni("cmdstr", "dispclick", "255,48,50,59,48,48,59,48,51,13,10", App.Path & "\Config.ini"), ",")
                For i = LBound(cmdstr) To UBound(cmdstr)
                    cmdstr(i) = Val(cmdstrs(i - 1))
                Next
                FrmMain.MSComm.OutPut = cmdstr
            End If
        Case 8      'F4
            ReadAuto = False
            FrmMain.Timer6.Enabled = False
            DeleteHotKey
            KeyChoose = 1
            Set frmChoose = New frmChoose
            frmChoose.show
            'frmMain.Command1_Click (3)
        Case 9      'F5
            DoEvents
            ReadAuto = False
            FrmMain.Timer6.Enabled = False
            DeleteHotKey
            KeyChoose = 2
            Set frmChoose = New frmChoose
            frmChoose.show
            'frmMain.Command1_Click (4)
        Case 10     'F6
            If FrmMain.Check1.Value = 1 Then FrmMain.MSComm.OutPut = Trim("%00;17;06" & vbCrLf)
        Case 11     'F7
            ReadAuto = False
            FrmMain.Timer6.Enabled = False
            If Instru = 0 Then WriteString Winsock1, "*RST"
            ReadAuto = True
            If Instru = 0 Then
                FrmMain.Timer6.Enabled = True
            Else
                FrmMain.MSComm0.OutPut = ":FETCh?" & vbCr
            End If
            If FrmMain.Check1.Value = 1 Then FrmMain.MSComm.OutPut = Trim("%00;17;07" & vbCrLf)
        
        Case 21
            Select Case Stall
            Case 2
            If frmChoose.Shape1.Top >= 1600 Then frmChoose.Shape1.Top = 800
            
            Case 3
            If frmChoose.Shape1.Top >= 2400 - frmChoose.Shape1.Height Then
                frmChoose.Shape1.Top = 1100
            ElseIf frmChoose.Shape1.Top >= 1100 Then
                frmChoose.Shape1.Top = 0
            End If
            
            Case 4
            If frmChoose.Shape1.Top >= 2400 - frmChoose.Shape1.Height Then
                frmChoose.Shape1.Top = 1550
            ElseIf frmChoose.Shape1.Top >= 1550 Then
                frmChoose.Shape1.Top = 750
            ElseIf frmChoose.Shape1.Top >= 750 Then
                frmChoose.Shape1.Top = 0
            End If
            
            Case 5
            If frmChoose.Shape1.Top >= 2400 - frmChoose.Shape1.Height Then
                frmChoose.Shape1.Top = 1650
            ElseIf frmChoose.Shape1.Top >= 1650 Then
                frmChoose.Shape1.Top = 1100
            ElseIf frmChoose.Shape1.Top >= 1100 Then
                frmChoose.Shape1.Top = 550
            ElseIf frmChoose.Shape1.Top >= 550 Then
                frmChoose.Shape1.Top = 0
            End If
            
            End Select
        Case 22
            Select Case Stall
            Case 2
            If frmChoose.Shape1.Top <= 800 Then frmChoose.Shape1.Top = 1600

            
            Case 3
            If frmChoose.Shape1.Top < 1100 Then
                frmChoose.Shape1.Top = 1100
            ElseIf frmChoose.Shape1.Top < 2400 - frmChoose.Shape1.Height Then
                frmChoose.Shape1.Top = 2400 - frmChoose.Shape1.Height
            End If
            
            Case 4
            If frmChoose.Shape1.Top < 750 Then
                frmChoose.Shape1.Top = 750
            ElseIf frmChoose.Shape1.Top < 1550 Then
                frmChoose.Shape1.Top = 1550
            ElseIf frmChoose.Shape1.Top < 2400 - frmChoose.Shape1.Height Then
                frmChoose.Shape1.Top = 2400 - frmChoose.Shape1.Height
            End If
            
            Case 5
            If frmChoose.Shape1.Top < 550 Then
                frmChoose.Shape1.Top = 550
            ElseIf frmChoose.Shape1.Top < 1100 Then
                frmChoose.Shape1.Top = 1100
            ElseIf frmChoose.Shape1.Top < 1650 Then
                frmChoose.Shape1.Top = 1650
            ElseIf frmChoose.Shape1.Top < 2400 - frmChoose.Shape1.Height Then
                frmChoose.Shape1.Top = 2400 - frmChoose.Shape1.Height
            End If
            
            End Select
        Case 23
            frmChoose.Command1_Click
    End Select
errhdl: If err.Number <> 0 Then MsgBox "Warning:Error" & err.Number & vbNewLine & err.Description
End Sub

Public Sub AddHotkey()
    ReDim uVirtKey(1 To 11)
    uVirtKey(1) = &H43
    uVirtKey(2) = &H1B
    uVirtKey(3) = &H5A
    uVirtKey(4) = &HD
    uVirtKey(5) = &H70
    uVirtKey(6) = &H71
    uVirtKey(7) = &H72
    uVirtKey(8) = &H73
    uVirtKey(9) = &H74
    uVirtKey(10) = &H75
    uVirtKey(11) = &H76
    Dim Modifiers As Long
    preWinProc = GetWindowLong(FrmMain.hwnd, GWL_WNDPROC)
    SetWindowLong FrmMain.hwnd, GWL_WNDPROC, AddressOf WndProc
    For i = 1 To 11
    RegisterHotKey FrmMain.hwnd, i, Modifiers, uVirtKey(i)
    Next i
End Sub

Public Sub DeleteHotKey()
SetWindowLong FrmMain.hwnd, GWL_WNDPROC, preWinProc
Busy = True
For i = 1 To 11
UnregisterHotKey FrmMain.hwnd, uVirtKey(i)
Next i
End Sub
