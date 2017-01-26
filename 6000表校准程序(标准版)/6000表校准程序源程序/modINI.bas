Attribute VB_Name = "modIni"
Option Explicit
'读出自定义INI文件
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'写入自定义INI文件
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'读出自定义INI文件中的单个区段间的所有键名和值
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
'读出自定义INI所有区段名
Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetIni(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim ResultString As String * 255
    If GetPrivateProfileString(ByVal SectionName, ByVal KeyWord, vbNullString, ByVal ResultString, ByVal Len(ResultString), ByVal FileName) > 0 Then '关键词的值不为空
        GetIni = Left(ResultString, InStr(ResultString, Chr(0)) - 1)
    Else    '将缺省值写入INI文件
        WritePrivateProfileString SectionName, KeyWord, DefString, FileName
        GetIni = DefString
    End If
End Function

Public Function GetKeyWord(ByVal SectionName As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim szBuf As String * 255
    If GetPrivateProfileSection(ByVal SectionName, ByVal szBuf, Len(szBuf), ByVal FileName) > 0 Then
        '同时获取键名和值
        GetKeyWord = Left(szBuf, InStr(szBuf, Chr(0)) - 1)
    Else
        WritePrivateProfileString SectionName, DefString, vbNullString, FileName
        GetKeyWord = DefString
    End If
End Function

Public Function GetKey(ByVal SectionName As String, ByVal DefString As String, ByVal FileName As String) As String
    Dim szBuf As String * 255, ResultString As String
    ResultString = GetKeyWord(ByVal SectionName, ByVal szBuf, ByVal FileName)
    If InStr(ResultString, "=") <> 0 Then
        GetKey = Left(ResultString, InStr(ResultString, "=") - 1)
    Else
        GetKey = DefString
    End If
End Function

Public Sub WriteINI(ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String)
WritePrivateProfileString lpApplicationName, lpKeyName, lpString, lpFileName
End Sub

Public Sub SaveINI(FileName As String)

If FileName = "" Then FileName = App.Path & "\Config.ini"

WriteINI "Adjust", "100mV(+)", Adjnum(1).POS, FileName
WriteINI "Adjust", "100mV(-)", Adjnum(1).Neg, FileName
Dim I As Integer
For I = 2 To 5
WriteINI "Adjust", RangArry(I) & "V" & "(+)", Adjnum(I).POS, FileName
WriteINI "Adjust", RangArry(I) & "V" & "(-)", Adjnum(I).Neg, FileName
Next I

WriteINI "Comm", "Port0", Port0, FileName
WriteINI "Comm", "Port1", Port1, FileName
WriteINI "Comm", "Port2", Port2, FileName
WriteINI "Comm", "CommIn", CInt(CommIn), FileName
WriteINI "Comm", "CommOut", CInt(CommOut), FileName

WriteINI "Lan", "RmHost", RmHost, FileName
WriteINI "Lan", "RmPort", RemPort, FileName

WriteINI "Custom", "DDigits", DDigits, FileName
WriteINI "Custom", "Instru", Instru, FileName
WriteINI "Custom", "Filter", Filter, FileName

Open App.Path & "\Configuration\LAN设置.txt" For Output As #3
Print #3, "[Lan]           Lan设置"
Print #3, "RmHost=" & RmHost & ";    仪器IP地址"
Print #3, "RmPort=5025;        仪器端口号(一般不需改动)"
Close #3

Open App.Path & "\Configuration\RS232端口设置.txt" For Output As #3
Print #3, "[Comm]          串口设置"
Print #3, "Port0=" & Port0 & ";    读数仪器选择为吉时利时的端口号"
Print #3, "Port1=" & Port1 & ";    指令端口号"
Print #3, "Port2=" & Port2 & ";    校准端口号"
Print #3, "CommIn=" & CInt(CommIn) & ";  是否从指令端口输入数据，0为不输入，非0为输入"
Print #3, "CommOut=" & CInt(CommOut) & ";  是否从校准端口输出数据，0为不输入，非0为输出"
Close #3

Open App.Path & "\Configuration\软件通用设置.txt" For Output As #3
Print #3, "[Custom]    用户通用软件设置"
Print #3, "DDigits=" & DDigits & ";  数据显示位数"
Print #3, "DDigits1=" & DDigitsArry(1) & ";  100mV显示位数"
Print #3, "DDigits2=" & DDigitsArry(2) & ";  1   V显示位数"
Print #3, "DDigits3=" & DDigitsArry(3) & ";  10  V显示位数"
Print #3, "DDigits4=" & DDigitsArry(4) & ";  100 V显示位数"
Print #3, "DDigits5=" & DDigitsArry(5) & ";  1000V显示位数"
Print #3, "Instru=" & Instru & ";   读数仪器选择，0为安捷伦，1为吉时利"
Print #3, "Filter=" & Filter & ";  正在使用的滤波值的大小"
Print #3, "Filter1=" & FilterArry(1) & ";  100mV滤波值"
Print #3, "Filter2=" & FilterArry(2) & ";  1   V滤波值"
Print #3, "Filter3=" & FilterArry(3) & ";  10  V滤波值"
Print #3, "Filter4=" & FilterArry(4) & ";  100 V滤波值"
Print #3, "Filter5=" & FilterArry(5) & ";  1000V滤波值"
Print #3, "DelayTime1=" & DelayTime(1) & ";  100mV稳定时间"
Print #3, "DelayTime2=" & DelayTime(2) & ";  1   V稳定时间"
Print #3, "DelayTime3=" & DelayTime(3) & ";  1   V稳定时间"
Print #3, "DelayTime4=" & DelayTime(4) & ";  100 V稳定时间"
Print #3, "DelayTime5=" & DelayTime(5) & ";  1000V稳定时间"
Print #3, "Tolerance1=" & Tolerance(1) & ";  120mV稳定门限"
Print #3, "Tolerance2=" & Tolerance(2) & ";  1   V稳定门限"
Print #3, "Tolerance3=" & Tolerance(3) & ";  10  V稳定门限"
Print #3, "Tolerance4=" & Tolerance(4) & ";  100 V稳定门限"
Print #3, "Tolerance5=" & Tolerance(5) & ";  1000V稳定门限"
Print #3, "MultCons=" & MultCons & ";  乘常数因子"
Print #3, "DivCons=" & DivCons & ";  除常数因子"
Print #3, "ZeroToler1=" & ZeroToler(1) & ";  100mV零位范围"
Print #3, "ZeroToler2=" & ZeroToler(2) & ";  1   V零位范围"
Print #3, "ZeroToler3=" & ZeroToler(3) & ";  10  V零位范围"
Print #3, "ZeroToler4=" & ZeroToler(4) & ";  100 V零位范围"
Print #3, "ZeroToler5=" & ZeroToler(5) & ";  1000V零位范围"
Print #3, "ZeroTolerW=" & ZeroTolerW & ";  零位范围"
Print #3, "DepartToler1=" & DepartToler(1) & ";  100mV区分门限"
Print #3, "DepartToler2=" & DepartToler(2) & ";  1   V区分门限"
Print #3, "DepartToler3=" & DepartToler(3) & ";  10  V区分门限"
Print #3, "DepartToler4=" & DepartToler(4) & ";  100 V区分门限"
Print #3, "DepartToler5=" & DepartToler(5) & ";  1000V区分门限"
Print #3, "DepartTolerW=" & DepartTolerW & ";  区分门限"
Print #3, "DataBit2=" & Data2_3(1) & ";  校准数据第二字节"
Print #3, "DataBit3=" & Data2_3(2) & ";  校准数据第三字节"
Print #3, "timer8=" & FrmMain.Timer8.Interval & ";  校准数据时间间隔"
Close #3

Open App.Path & "\Configuration\修正系数.txt" For Output As #3
Print #3, "[Adjust]        修正系数设置"
Print #3, "100mV(+)=" & Adjnum(1).POS & ";     量程为100mV的正测量结果的修正系数"
Print #3, "100mV(-)=" & Adjnum(1).Neg & ";     量程为100mV的负测量结果的修正系数"
Print #3, "1V(+)=" & Adjnum(2).POS & ";        量程为1V的正测量结果的修正系数"
Print #3, "1V(-)=" & Adjnum(2).Neg & ";        量程为1V的负测量结果的修正系数"
Print #3, "10V(+)=" & Adjnum(3).Neg & ";       量程为10V的正测量结果的修正系数"
Print #3, "10V(-)=" & Adjnum(3).POS & ";       量程为10V的负测量结果的修正系数"
Print #3, "100V(+)="; Adjnum(4).POS & ";      量程为100V的正测量结果的修正系数"
Print #3, "100V(-)=" & Adjnum(4).Neg & ";      量程为100V的负测量结果的修正系数"
Print #3, "1000V(+)=" & Adjnum(5).POS & ";     量程为1000V的正测量结果的修正系数"
Print #3, "1000V(-)=" & Adjnum(5).Neg & ";     量程为1000V的负测量结果的修正系数"
Close #3

Open App.Path & "\Configuration\量程段编号设置.txt" For Output As #3
Print #3, "[RangeID]        量程段编号设置"
Print #3, "Range0.001=" & RangeID(1) & ";     1mV量程段编号"
Print #3, "Range0.01=" & RangeID(2) & ";     10mV量程段编号"
Print #3, "Range0.1=" & RangeID(3) & ";     100mV量程段编号"
Print #3, "Range1=" & RangeID(4) & ";     1V量程段编号"
Print #3, "Range10=" & RangeID(5) & ";     10V量程段编号"
Print #3, "Range100=" & RangeID(6) & ";     100V量程段编号"
Print #3, "Range1000=" & RangeID(7) & ";     1000V量程段编号"
Close #3

End Sub

Public Sub ReadINI(FileName As String)
If FileName = "" Then FileName = App.Path & "Config.ini"
RmHost = GetIni("Lan", "RmHost", "169.254.4.10", FileName)
RemPort = Val(GetIni("Lan", "RmPort", 5025, FileName))
Port0 = Val(GetIni("Comm", "Port0", 1, FileName))
Port1 = Val(GetIni("Comm", "Port1", 5, FileName))
Port2 = Val(GetIni("Comm", "Port2", 6, FileName))
CommIn = CBool(GetIni("Comm", "CommIn", True, FileName))
CommOut = CBool(GetIni("Comm", "CommOut", True, FileName))
Adjnum(1).POS = GetIni("Adjust", "100mV(+)", 1, FileName)
Adjnum(1).Neg = GetIni("Adjust", "100mV(-)", 1, FileName)
Adjnum(2).POS = GetIni("Adjust", "1V(+)", 1, FileName)
Adjnum(2).Neg = GetIni("Adjust", "1V(-)", 1, FileName)
Adjnum(3).POS = GetIni("Adjust", "10V(+)", 1, FileName)
Adjnum(3).Neg = GetIni("Adjust", "10V(-)", 1, FileName)
Adjnum(4).POS = GetIni("Adjust", "100V(+)", 1, FileName)
Adjnum(4).Neg = GetIni("Adjust", "100V(-)", 1, FileName)
Adjnum(5).POS = GetIni("Adjust", "1000V(+)", 1, FileName)
Adjnum(5).Neg = GetIni("Adjust", "1000V(-)", 1, FileName)
DDigits = GetIni("Custom", "DDigits", 7, FileName)
DDigitsArry(1) = GetIni("Custom", "DDigits1", 5, FileName)
DDigitsArry(2) = GetIni("Custom", "DDigits2", 7, FileName)
DDigitsArry(3) = GetIni("Custom", "DDigits3", 6, FileName)
DDigitsArry(4) = GetIni("Custom", "DDigits4", 5, FileName)
DDigitsArry(5) = GetIni("Custom", "DDigits5", 5, FileName)
Instru = Val(GetIni("Custom", "Instru", 1, FileName))
Filter = Val(GetIni("Custom", "Filter", 10, FileName))
FilterArry(1) = Val(GetIni("Custom", "Filter1", 300, FileName))
FilterArry(2) = Val(GetIni("Custom", "Filter2", 100, FileName))
FilterArry(3) = Val(GetIni("Custom", "Filter3", 50, FileName))
FilterArry(4) = Val(GetIni("Custom", "Filter4", 50, FileName))
FilterArry(5) = Val(GetIni("Custom", "Filter5", 50, FileName))
DelayTime(1) = Val(GetIni("Custom", "DelayTime1", 3, FileName))
DelayTime(2) = Val(GetIni("Custom", "DelayTime2", 2, FileName))
DelayTime(3) = Val(GetIni("Custom", "DelayTime3", 3, FileName))
DelayTime(4) = Val(GetIni("Custom", "DelayTime4", 3, FileName))
DelayTime(5) = Val(GetIni("Custom", "DelayTime5", 3, FileName))
Tolerance(1) = Val(GetIni("Custom", "Tolerance1", 0.0003, FileName))
Tolerance(2) = Val(GetIni("Custom", "Tolerance2", 0.00001, FileName))
Tolerance(3) = Val(GetIni("Custom", "Tolerance3", 0.0003, FileName))
Tolerance(4) = Val(GetIni("Custom", "Tolerance4", 0.0003, FileName))
Tolerance(5) = Val(GetIni("Custom", "Tolerance5", 0.0003, FileName))
MultCons = Val(GetIni("Custom", "MultCons", 1, FileName))
DivCons = Val(GetIni("Custom", "DivCons", 1, FileName))
ZeroToler(1) = Val(GetIni("Custom", "ZeroToler1", 0.002, FileName))
ZeroToler(2) = Val(GetIni("Custom", "ZeroToler2", 0.002, FileName))
ZeroToler(3) = Val(GetIni("Custom", "ZeroToler3", 0.002, FileName))
ZeroToler(4) = Val(GetIni("Custom", "ZeroToler4", 0.002, FileName))
ZeroToler(5) = Val(GetIni("Custom", "ZeroToler5", 0.002, FileName))
ZeroTolerW = Val(GetIni("Custom", "ZeroTolerW", 0.002, FileName))
DepartToler(1) = Val(GetIni("Custom", "DepartToler1", 0.002, FileName))
DepartToler(2) = Val(GetIni("Custom", "DepartToler2", 0.002, FileName))
DepartToler(3) = Val(GetIni("Custom", "DepartToler3", 0.002, FileName))
DepartToler(4) = Val(GetIni("Custom", "DepartToler4", 0.002, FileName))
DepartToler(5) = Val(GetIni("Custom", "DepartToler5", 0.002, FileName))
DepartTolerW = Val(GetIni("Custom", "DepartTolerW", 0.002, FileName))
RangeID(1) = Val(GetIni("RangeID", "Range0.001", 0, FileName))
RangeID(2) = Val(GetIni("RangeID", "Range0.01", 1, FileName))
RangeID(3) = Val(GetIni("RangeID", "Range0.1", 2, FileName))
RangeID(4) = Val(GetIni("RangeID", "Range1", 3, FileName))
RangeID(5) = Val(GetIni("RangeID", "Range10", 4, FileName))
RangeID(6) = Val(GetIni("RangeID", "Range100", 5, FileName))
RangeID(7) = Val(GetIni("RangeID", "Range1000", 6, FileName))
Data2_3(1) = Val(GetIni("Custom", "DataBit2", 48, FileName))
Data2_3(2) = Val(GetIni("Custom", "DataBit3", 49, FileName))
End Sub
