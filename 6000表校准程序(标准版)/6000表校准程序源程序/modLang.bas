Attribute VB_Name = "modLanguage"
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long

Public Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
  wServicePackMinor As Integer
  wProductType As Byte
  OsName As String
  OsLanguage As String
End Type

Public Type Capt
    Ch As String
    En As String
End Type

Public captions(21) As Capt


Public Sub modBtnLang(cmdBtn As CommandButton, Caption As Capt)
Select Case Lang
    Case 0
        cmdBtn.Caption = Caption.Ch
    Case 1
        cmdBtn.Caption = Caption.En
End Select
End Sub

Public Sub modLblLang(Lbl As Label, Caption As Capt)
Select Case Lang
    Case 0
        Lbl.Caption = Caption.Ch
    Case 1
        Lbl.Caption = Caption.En
End Select
End Sub

Public Sub modFrmLang(frm As Frame, Caption As Capt)
Select Case Lang
    Case 0
        frm.Caption = Caption.Ch
    Case 1
        frm.Caption = Caption.En
End Select
End Sub

Public Sub langResize()
FrmMain.Label1.Left = -FrmMain.Label1.Width / 2 + FrmMain.TextDisp.Left + FrmMain.TextDisp.Width / 2
End Sub

Public Sub modLang()
Busy = True
FrmMain.lblAuto.Caption = IIf(Lang, "Auto", "自动")
FrmMain.Check3.Caption = IIf(Lang, "AutoCali", "自动校准")
FrmMain.mnuCurve.Caption = IIf(Lang, "Curve", "实时曲线")
FrmMain.mnuRangeID.Caption = IIf(Lang, "Range ID settings", "量程段编号设置")
FrmMain.Check2.Caption = IIf(Lang, "ZeroCali", "零位校准")
FrmMain.lblStable.Caption = IIf(Lang, "Stable", "稳定")
FrmMain.mnuTips.Caption = IIf(Lang, "Tips", "使用提示")
FrmMain.mnuDDigits.Caption = IIf(Lang, "Change the display digits", "更改显示位数")
FrmMain.mnuKEI.Caption = IIf(Lang, "Parameters of KEITHLEY(Reading Port)", "吉时利(读数口)参数设置")
FrmMain.Option1.Caption = IIf(Lang, "DDM Selection", "数字标准表选择")
FrmMain.mnuKEIT.Caption = IIf(Lang, "Switch to Keithley2000", "切换到吉时利2000")
FrmMain.mnuAgi.Caption = IIf(Lang, "Switch to Agilent34410A", "切换到安捷伦34410A")
FrmMain.mnuComset.Caption = IIf(Lang, "Comm Set", "串口设置")
FrmMain.Label3(4).Caption = IIf(Lang, "Key length:" & TimeSpan & "ms", "按键时长：" & TimeSpan & "ms")
FrmMain.Check1.Caption = IIf(Lang, "Send command to 6000 DMM", "发送命令到6000表")
FrmMain.mnuTheme.Caption = IIf(Lang, "Change theme", "更换主题")
FrmMain.mnuTH.Caption = IIf(Lang, "Look for one theme", "指定的主题")
FrmMain.mnuTh1.Caption = IIf(Lang, "Aero style", "Aero特效")
FrmMain.mnuTh2.Caption = IIf(Lang, "China style", "中国风")
FrmMain.mnuTh3.Caption = IIf(Lang, "Longhorn style", "Longhorn特效")
FrmMain.mnuHotKey.Caption = IIf(Lang, "Reboot hotkey", "重启快捷键")
FrmMain.mnuHlelp.Caption = IIf(Lang, "Help", "帮助")
FrmMain.mnuInstruc.Caption = IIf(Lang, "Instructions", "程序使用说明")
FrmMain.mnuAbout.Caption = IIf(Lang, "About", "关于")
FrmMain.mnuLanguage.Caption = IIf(Lang, "Language", "语言")
FrmMain.mnuImport.Caption = IIf(Lang, "Import setting", "从文件导入设置")
FrmMain.mnuExport.Caption = IIf(Lang, "Save setting", "保存设置到文件")
FrmMain.mnuFile.Caption = IIf(Lang, "File", "文件")
FrmMain.mnuExit.Caption = IIf(Lang, "Quit", "退出")
FrmMain.mnuNul.Caption = IIf(Lang, "Null", "置零")
FrmMain.mnuPeak.Caption = IIf(Lang, "Peak", "运行")
FrmMain.mnuDigit.Caption = IIf(Lang, "Digit", "位数")
FrmMain.mnuDisp.Caption = IIf(Lang, "Display", "显示")
FrmMain.mnuFunc.Caption = IIf(Lang, "Function", "功能")
FrmMain.mnuLanset.Caption = IIf(Lang, "Lan Set", "Lan设置")
FrmMain.mnuOptions.Caption = IIf(Lang, "Options", "设置")
FrmMain.mnuPrint.Caption = IIf(Lang, "Print", "打印")
FrmMain.mnuRange.Caption = IIf(Lang, "Range", "量程")
FrmMain.mnuReset.Caption = IIf(Lang, "Reset", "复位")
FrmMain.Frame2.Caption = IIf(Lang, "Pramaters and Status", "参数及状态信息")
FrmMain.LblNul.Caption = IIf(Lang, "Null", "置零")
FrmMain.Label1.Caption = IIf(Instru, IIf(Lang, "6000 Multimeter C.P(KEITHLEY 2000)", "6000数字表校准程序(吉时利2000)"), IIf(Lang, "6000 Multimeter C.P(Agilent34410A)", "6000数字表校准程序(安捷伦34410A)"))
FrmMain.Label3(5).Caption = IIf(Lang, "CaliMode:", "校准方式：")
FrmMain.lblZero.Caption = IIf(Lang, "Zero", "零位")
'C.P——Multimeter Calibration Procedure
FrmMain.mnuTol.Caption = IIf(Lang, "Tollerence Settings", "校准及门限设置")
FrmMain.mnuFilter.Caption = IIf(Lang, "Filter Settings", "滤波设置")
FrmMain.mnuAdjust.Caption = IIf(Lang, "Adjustment", "数据校正参数设置")
FrmMain.Label2(4).Caption = IIf(Lang, "Cali Data", "校准数据")

langResize
captions(18).Ch = IIf(Instru, "仪器编号:", "远程IP地址:"): captions(18).En = IIf(Instru, "Instrument Number:", "Remote IP Address:")
captions(19).Ch = IIf(Instru, "校准端口:", "远程端口:"): captions(19).En = IIf(Instru, "Adjusting Port:", "Remote Port:")
captions(20).Ch = IIf(Instru, "读数端口:", "本地端口:"): captions(20).En = IIf(Instru, "Data Port:", "Loacal Port      :")
captions(21).Ch = IIf(Instru, "指令端口:", "传输协议:"): captions(21).En = IIf(Instru, "Command Port:", "Protocol   :")

For I = 0 To 6
modBtnLang FrmMain.Command1(I), captions(I)
Next
For I = 8 To 13
If I <> 12 Then modBtnLang FrmMain.Command2(I - 7), captions(I)
Next
For I = 14 To 17
modLblLang FrmMain.Label2(I - 14), captions(I)
Next
For I = 18 To 21
modLblLang FrmMain.Label3(I - 18), captions(I)
Next
Dim bol As Boolean
bol = (Winsock1.State = sckConnected) Or (FrmMain.MSComm0.PortOpen = True)
FrmMain.Command2(0).Caption = IIf(bol, IIf(Lang, "Disconnect", "断开"), IIf(Lang, "Connect", "连接"))
Busy = False
End Sub

Public Function Getsyslan() As String
  Dim Ver As OSVERSIONINFO
  Ver.dwOSVersionInfoSize = 148
  GetVersionEx Ver
  With Ver
      Dim LocaleID As Long
      LocaleID = GetSystemDefaultLCID
      Select Case LocaleID
          Case &H804
              .OsLanguage = "CH"
          Case Else
              .OsLanguage = "EN"
      End Select
  End With
  Getsyslan = Ver.OsLanguage
End Function
