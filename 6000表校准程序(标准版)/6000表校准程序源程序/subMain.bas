Attribute VB_Name = "subMain"
Option Base 1

Public Winsock1 As New CSocketMaster
Public RmHost As String
Public Busy As Boolean
Public RANGe As Double
Public Filter As Long, FilterArry(1 To 5) As Long
Public Lang As Long
Public Adjnum(5) As Adj
Public RangArry(5) As Double
Public OutPut As Boolean
Public comingCMD As String
Public longBtn As Byte
Public TimeStart As Long, timenow As Long, TimeSpan As Long
Public Const H As Long = 1215
Public Const T As Long = 3900
Public Port0 As Long, Port1 As Long, Port2 As Long
Public CommIn As Boolean, CommOut As Boolean
Public Instru As Long
Public DDigits As Long, DDigitsArry(1 To 5) As Long
Public cmdstr() As Byte
Public NulOn As Boolean
Public ReadAuto As Boolean
Public KeyChoose As Long
Public numArry() As Double, Cnt As Long, ss As Long
Public Base As Double, Sign As Double, Outbit As Double
Public ShowAtStartup As Long
Public DelayTime(1 To 5) As Long, Tolerance(1 To 5) As Double, DelayTimeW As Long, ToleranceW As Double, ZeroToler(1 To 5) As Double, ZeroTolerW As Double, DepartToler(1 To 5) As Double, DepartTolerW As Double
Public AutoCali(1 To 3) As Boolean, dataArry As Dataset
Public AdjCnt As Long, AdjOut As New Collection, ZeroSendOnce As Boolean
Public RangeID(1 To 7) As Long, RangeIDIndex As Long
Public MultCons As Double, DivCons As Double
Public Data2_3(1 To 2) As Byte


Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Const TH32CS_SNAPPROCESS = &H2&
Public Const WM_CLOSE = &H10

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Public Type Dataset
Data As New Collection
timenow As New Collection
End Type

Public Type Adj
POS As Double
Neg As Double
End Type


Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As _
    String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As _
    Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, _
lpcbData As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
    As String, ByVal Reserved As Long, ByVal lpClass As Long, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal _
    lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal _
    lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, Source As Any, ByVal numBytes As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Sub Main()
Lang = IIf(Getsyslan = "CH", 0, 1)
WriteINI "UnloadMode", "Unload", "0", App.Path & "\Config.ini"
If App.PrevInstance = True Then
If Shell(App.Path & "\Reboot.exe", vbHide) = 1 Then End
End If

ShowAtStartup = GetSetting(App.EXEName, "Options", "在启动时显示提示", 1)

If Dir(App.Path & "\Themes\china.she") <> "" Then
SkinH_Attach
SkinH_AttachEx App.Path & "\Themes\china.she", ""
End If

captions(0).Ch = "置零": captions(0).En = "Null"
captions(1).Ch = "运行": captions(1).En = "Peak"
captions(2).Ch = "显示": captions(2).En = "Display"
captions(3).Ch = "位数": captions(3).En = "Decimal"
captions(4).Ch = "量程": captions(4).En = "Range"
captions(5).Ch = "打印": captions(5).En = "Print"
captions(6).Ch = "复位": captions(6).En = "Reset"

captions(8).Ch = "标准量程": captions(8).En = "Standard Range"
captions(9).Ch = "显示位数(+)": captions(9).En = "Disp_Digit(+)"
captions(10).Ch = "被较位数(-)": captions(10).En = "Cali_Digit(-)"
captions(11).Ch = "显示位数(-)": captions(11).En = "Disp_Digit(-)"
captions(12).Ch = "连续接收": captions(12).En = "ReceiveNonStop"
captions(13).Ch = "被较位数(+)": captions(13).En = "Cali_Digit(+)"

captions(14).Ch = "连接状态": captions(14).En = "Connection"
captions(15).Ch = "读取状态": captions(15).En = "Reading"
captions(16).Ch = "发送状态": captions(16).En = "Sending"
captions(17).Ch = "错误提示": captions(17).En = "Error"

captions(18).Ch = IIf(Instru, "仪器编号:", "远程IP地址:"): captions(18).En = IIf(Instru, "Instrument Number:", "Remote IP Address:")
captions(19).Ch = IIf(Instru, "校准端口:", "远程端口:"): captions(19).En = IIf(Instru, "Adjusting Port:", "Remote Port:")
captions(20).Ch = IIf(Instru, "读数端口:", "本地端口:"): captions(20).En = IIf(Instru, "Data Port:", "Loacal Port      :")
captions(21).Ch = IIf(Instru, "指令端口:", "传输协议:"): captions(21).En = IIf(Instru, "Command Port:", "Protocol   :")
RangArry(1) = 0.1: RangArry(2) = 1: RangArry(3) = 10: RangArry(4) = 100: RangArry(5) = 1000
Load frmSplash
frmSplash.show

RangeIDIndex = 4
End Sub

Public Sub OutTo6000(Number As String)
On Error GoTo errhdl
If Number <> "" Then
Dim DataOut(1 To 37) As Byte
DataOut(1) = 255
DataOut(2) = Data2_3(1)
DataOut(3) = Data2_3(2)
DataOut(4) = 59
Select Case RangeIDIndex
Case 1, 4
Number = CStr(Format(Number, IIf(Number > 0, "+0.000000000", "0.000000000")))
Case 2, 5
Number = CStr(Format(Number, IIf(Number > 0, "+00.00000000", "00.00000000")))
Case 3, 6
Number = CStr(Format(Number, IIf(Number > 0, "+000.0000000", "000.0000000")))
Case 7
Number = CStr(Format(Number, IIf(Number > 0, "+0000.000000", "0000.000000")))
End Select

Dim I As Integer
For I = 5 To 16
DataOut(I) = Asc(Mid(Number, I - 4, 1))
Next

DataOut(17) = 59
DataOut(18) = 48
DataOut(19) = 50
DataOut(20) = 59
DataOut(21) = 224
DataOut(22) = 0
DataOut(23) = 0
DataOut(24) = 72
DataOut(25) = 0
DataOut(26) = 0
DataOut(27) = 59
DataOut(28) = 48
DataOut(29) = 49
DataOut(30) = 59
DataOut(31) = 48
For I = 1 To 7
If RangeIDIndex = I Then
DataOut(32) = 48 + RangeID(I)
Exit For
End If
Next
DataOut(33) = 59
DataOut(34) = 48
DataOut(35) = 48
DataOut(36) = 13
DataOut(37) = 10
If FrmMain.MSComm.PortOpen = True Then FrmMain.MSComm.OutPut = DataOut
If FrmMain.MSComm1.PortOpen = True Then FrmMain.MSComm1.OutPut = DataOut
End If
errhdl: Exit Sub
End Sub

Public Sub showhideWK(show As Boolean)
Select Case show
Case True
    FrmMain.Height = 8685
    FrmMain.Frame3.Visible = True
    FrmMain.Frame3.Top = T
    FrmMain.Frame3.Height = H
    FrmMain.PB.Top = 3700
    FrmMain.Frame2.Top = 5280
    For I = 0 To 6
    If I <> 5 Then FrmMain.Command2(I).Top = 6960
    Next I
    FrmMain.Stb.Top = 7560
Case False
    FrmMain.Height = 8685 - H
    FrmMain.Frame3.Visible = False
    FrmMain.PB.Top = 3700 - H
    FrmMain.Frame2.Top = 5280 - H
    For I = 0 To 6
    If I <> 5 Then FrmMain.Command2(I).Top = 6960 - H
    Next I
    FrmMain.Stb.Top = 7560 - H
End Select
End Sub

Public Function Stable(Data As Collection) As Boolean
Dim s As Double, I As Integer, a As Double, err As Double
For I = 1 To Data.Count
s = s + Val(Data.Item(I))
Next
a = s / Data.Count
err = Abs(Data.Item(Data.Count) - a)
If err < IIf(RANGe = 0.1, 100, RANGe) * ToleranceW Then Stable = True
End Function

Public Sub updatefrmData(str As String)
Dim a() As String, str0 As String, str1 As String, str2 As String
    frmData.RText0.Text = ""
    frmData.RText1.Text = ""
    frmData.RText2.Text = ""
    frmData.RText0.Text = "No." & Chr(10)
    For I = 0 To 28
        PostMessage frmData.RText0.hwnd, WM_KEYDOWN, VK_NEXT, 0
        frmData.RText0.SelStart = Len(frmData.RText0.Text)
        frmData.RText0.SelText = Format(I, "00") & Chr(10)
    Next
    frmData.RText1 = "正向" & IIf(RANGe = 0.1, "(mV)", "(V)") & Chr(10)
    frmData.RText2 = "负向" & IIf(RANGe = 0.1, "(mV)", "(V)") & Chr(10)
    AdjOut.Add str
    ReDim a(AdjOut.Count)
    For I = 1 To AdjOut.Count
        a(I) = AdjOut.Item(I)
    Next
    Dim tmp() As Double
    tmp = DelDuplicates(a)
    tmp = NumSort(tmp, "up")
    For I = 1 To UBound(tmp)
        If Abs(tmp(I)) <= CDbl(IIf(RANGe = 0.1, 100, RANGe)) * ZeroTolerW Then
            If tmp(I) = AdjOut.Item(AdjOut.Count) Then
                str1 = "+0.0000000" & "  ☆"
                str2 = "-0.0000000" & "  ☆"
            Else
                str1 = "+0.0000000"
                str2 = "-0.0000000"
            End If
         Else
            If tmp(I) > 0 Then
                If tmp(I) = AdjOut.Item(AdjOut.Count) Then
                    str1 = "+" & Left(FormatNumber(tmp(I), 10, vbTrue), 9) & "  √"
                Else
                    str1 = "+" & Left(FormatNumber(tmp(I), 10, vbTrue), 9)
                End If
                str2 = vbNullString
            Else
                If tmp(I) = AdjOut.Item(AdjOut.Count) Then
                    str2 = Left(FormatNumber(tmp(I), 10, vbTrue), 10) & "  √"
                Else
                    str2 = Left(FormatNumber(tmp(I), 10, vbTrue), 10)
                End If
                str1 = vbNullString
            End If
        End If
        PostMessage frmData.RText1.hwnd, WM_KEYDOWN, VK_NEXT, 0
        PostMessage frmData.RText2.hwnd, WM_KEYDOWN, VK_NEXT, 0
        frmData.RText1.SelStart = Len(frmData.RText1.Text)
        frmData.RText2.SelStart = Len(frmData.RText2.Text)
        frmData.RText1.SelText = IIf(str1 = vbNullString, "", str1 & Chr(10))
        frmData.RText2.SelText = IIf(str2 = vbNullString, "", str2 & Chr(10))
    Next
End Sub


Public Function NumSort(ByRef a() As Double, Optional sort As String = "up") As Double()  '按绝对值大小排序
Dim Min As Long, Max As Long, num As Long, first As Long, last As Long, temp As Long, all As New Collection, steps As Long
Min = LBound(a)
Max = UBound(a)
all.Add a(Min)
steps = 1

For num = Min + 1 To Max
    last = all.Count
    If Abs(a(num)) < Abs(CDbl(all(1))) Then all.Add a(num), BEFORE:=1: GoTo nextnum '加到第一项
    If Abs(a(num)) > Abs(CDbl(all(last))) Then all.Add a(num), AFTER:=last: GoTo nextnum '加到最后一项
    first = 1

    Do While last > first + 1 '利用DO循环减少循环次数
    temp = (last + first) \ 2

    If Abs(a(num)) > Abs(CDbl(all(temp))) Then
        first = temp
    Else
        last = temp
        steps = steps + 1
    End If
    Loop
    all.Add a(num), BEFORE:=last '加到指定的索引
nextnum: steps = steps + 1
Next

For num = Min To Max
If sort = "UP" Or sort = "up" Then a(num) = CDbl(all(num - Min + 1)): steps = steps + 1 '升序
If sort = "DOWN" Or sort = "down" Then a(num) = CDbl(all(Max - num + 1)): steps = steps + 1 '降序
Next
NumSort = a
Set all = Nothing
End Function


Public Function DelDuplicates(a() As String) As Double()
Dim B As New Collection, c() As Double, I As Integer, j As Integer
Dim temp As String
For I = 1 To UBound(a)
    For j = I + 1 To UBound(a)
        If Abs(Val(a(I)) - Val(a(j))) < CDbl(IIf(RANGe = 0.1, 100, RANGe)) * DepartTolerW Then a(I) = "@"
    Next
Next
For I = 1 To UBound(a)
If a(I) <> "@" Then B.Add a(I)
Next
ReDim c(B.Count)
For I = 1 To B.Count
c(I) = B.Item(I)
Next
DelDuplicates = c
End Function

Public Sub updateCurve(ByVal Value As String)
Dim Scal As Long
Scal = CLng(Split(Format(Value, "Scientific"), "E")(1))
'frmCurve.Curve1.Value = 100 / (10 ^ (Scal + 1)) * CDbl(Value)
frmCurve.Curve1.Value = CDbl(Value)
End Sub

'关闭指定名称的进程
Public Sub KillProcess(sProcess As String)
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
        tPE.dwSize = Len(tPE)
        lNextProcess = Process32First(lSnapShot, tPE)
        Do While lNextProcess
            If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
                Dim lProcess As Long
                Dim lExitCode As Long
                lProcess = OpenProcess(1, False, tPE.th32ProcessID)
                TerminateProcess lProcess, lExitCode
                CloseHandle lProcess
            End If
            lNextProcess = Process32Next(lSnapShot, tPE)
        Loop
        CloseHandle (lSnapShot)
    End If
End Sub
