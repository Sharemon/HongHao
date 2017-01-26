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
FrmMain.lblAuto.Caption = IIf(Lang, "Auto", "�Զ�")
FrmMain.Check3.Caption = IIf(Lang, "AutoCali", "�Զ�У׼")
FrmMain.mnuCurve.Caption = IIf(Lang, "Curve", "ʵʱ����")
FrmMain.mnuRangeID.Caption = IIf(Lang, "Range ID settings", "���̶α������")
FrmMain.Check2.Caption = IIf(Lang, "ZeroCali", "��λУ׼")
FrmMain.lblStable.Caption = IIf(Lang, "Stable", "�ȶ�")
FrmMain.mnuTips.Caption = IIf(Lang, "Tips", "ʹ����ʾ")
FrmMain.mnuDDigits.Caption = IIf(Lang, "Change the display digits", "������ʾλ��")
FrmMain.mnuKEI.Caption = IIf(Lang, "Parameters of KEITHLEY(Reading Port)", "��ʱ��(������)��������")
FrmMain.Option1.Caption = IIf(Lang, "DDM Selection", "���ֱ�׼��ѡ��")
FrmMain.mnuKEIT.Caption = IIf(Lang, "Switch to Keithley2000", "�л�����ʱ��2000")
FrmMain.mnuAgi.Caption = IIf(Lang, "Switch to Agilent34410A", "�л���������34410A")
FrmMain.mnuComset.Caption = IIf(Lang, "Comm Set", "��������")
FrmMain.Label3(4).Caption = IIf(Lang, "Key length:" & TimeSpan & "ms", "����ʱ����" & TimeSpan & "ms")
FrmMain.Check1.Caption = IIf(Lang, "Send command to 6000 DMM", "�������6000��")
FrmMain.mnuTheme.Caption = IIf(Lang, "Change theme", "��������")
FrmMain.mnuTH.Caption = IIf(Lang, "Look for one theme", "ָ��������")
FrmMain.mnuTh1.Caption = IIf(Lang, "Aero style", "Aero��Ч")
FrmMain.mnuTh2.Caption = IIf(Lang, "China style", "�й���")
FrmMain.mnuTh3.Caption = IIf(Lang, "Longhorn style", "Longhorn��Ч")
FrmMain.mnuHotKey.Caption = IIf(Lang, "Reboot hotkey", "������ݼ�")
FrmMain.mnuHlelp.Caption = IIf(Lang, "Help", "����")
FrmMain.mnuInstruc.Caption = IIf(Lang, "Instructions", "����ʹ��˵��")
FrmMain.mnuAbout.Caption = IIf(Lang, "About", "����")
FrmMain.mnuLanguage.Caption = IIf(Lang, "Language", "����")
FrmMain.mnuImport.Caption = IIf(Lang, "Import setting", "���ļ���������")
FrmMain.mnuExport.Caption = IIf(Lang, "Save setting", "�������õ��ļ�")
FrmMain.mnuFile.Caption = IIf(Lang, "File", "�ļ�")
FrmMain.mnuExit.Caption = IIf(Lang, "Quit", "�˳�")
FrmMain.mnuNul.Caption = IIf(Lang, "Null", "����")
FrmMain.mnuPeak.Caption = IIf(Lang, "Peak", "����")
FrmMain.mnuDigit.Caption = IIf(Lang, "Digit", "λ��")
FrmMain.mnuDisp.Caption = IIf(Lang, "Display", "��ʾ")
FrmMain.mnuFunc.Caption = IIf(Lang, "Function", "����")
FrmMain.mnuLanset.Caption = IIf(Lang, "Lan Set", "Lan����")
FrmMain.mnuOptions.Caption = IIf(Lang, "Options", "����")
FrmMain.mnuPrint.Caption = IIf(Lang, "Print", "��ӡ")
FrmMain.mnuRange.Caption = IIf(Lang, "Range", "����")
FrmMain.mnuReset.Caption = IIf(Lang, "Reset", "��λ")
FrmMain.Frame2.Caption = IIf(Lang, "Pramaters and Status", "������״̬��Ϣ")
FrmMain.LblNul.Caption = IIf(Lang, "Null", "����")
FrmMain.Label1.Caption = IIf(Instru, IIf(Lang, "6000 Multimeter C.P(KEITHLEY 2000)", "6000���ֱ�У׼����(��ʱ��2000)"), IIf(Lang, "6000 Multimeter C.P(Agilent34410A)", "6000���ֱ�У׼����(������34410A)"))
FrmMain.Label3(5).Caption = IIf(Lang, "CaliMode:", "У׼��ʽ��")
FrmMain.lblZero.Caption = IIf(Lang, "Zero", "��λ")
'C.P����Multimeter Calibration Procedure
FrmMain.mnuTol.Caption = IIf(Lang, "Tollerence Settings", "У׼����������")
FrmMain.mnuFilter.Caption = IIf(Lang, "Filter Settings", "�˲�����")
FrmMain.mnuAdjust.Caption = IIf(Lang, "Adjustment", "����У����������")
FrmMain.Label2(4).Caption = IIf(Lang, "Cali Data", "У׼����")

langResize
captions(18).Ch = IIf(Instru, "�������:", "Զ��IP��ַ:"): captions(18).En = IIf(Instru, "Instrument Number:", "Remote IP Address:")
captions(19).Ch = IIf(Instru, "У׼�˿�:", "Զ�̶˿�:"): captions(19).En = IIf(Instru, "Adjusting Port:", "Remote Port:")
captions(20).Ch = IIf(Instru, "�����˿�:", "���ض˿�:"): captions(20).En = IIf(Instru, "Data Port:", "Loacal Port      :")
captions(21).Ch = IIf(Instru, "ָ��˿�:", "����Э��:"): captions(21).En = IIf(Instru, "Command Port:", "Protocol   :")

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
FrmMain.Command2(0).Caption = IIf(bol, IIf(Lang, "Disconnect", "�Ͽ�"), IIf(Lang, "Connect", "����"))
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
