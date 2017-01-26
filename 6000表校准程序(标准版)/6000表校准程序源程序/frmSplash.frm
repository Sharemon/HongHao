VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6960
      Top             =   3840
   End
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7200
      Begin VB.Image imgLogo 
         Height          =   2385
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label lblCopyright 
         Caption         =   "电话：83391064 83391020"
         Height          =   255
         Left            =   4440
         TabIndex        =   4
         Top             =   2880
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "网址：www.bjhhfa.com"
         Height          =   255
         Left            =   4440
         TabIndex        =   3
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblinfo 
         Caption         =   "info"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Visible         =   0   'False
         Width           =   6765
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "版本"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6480
         TabIndex        =   5
         Top             =   2520
         Width           =   510
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "平台"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6480
         TabIndex        =   6
         Top             =   2160
         Width           =   510
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "授权"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "北京弘豪福安仪器有限公司"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3315
         TabIndex        =   7
         Top             =   960
         Width           =   2700
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128 ' Maintenance string for PSS usage
End Type

Private Sub Form_Load()
'lblCopyright.Caption = IIf(Lang, "Tel:    ", "电话：") & "83391064 83391020"
lblCompany.Caption = IIf(Lang, "Website:", "网址：") & "www.bjhhfa.com"
lblCompanyProduct.Caption = IIf(Lang, "Beijing HHFA Instrument Co.,Ltd.", "北京弘豪福安仪器有限公司")
'lblPlatform.Left = lblVersion.Left
Dim xx As OSVERSIONINFO
xx.dwOSVersionInfoSize = 148
Dim StrT As String * 255
GetComputerName StrT, 255
GetVersionEx xx
    lblPlatform.Caption = IIf(Lang, "Platform:", "平台:") & "Windows" & xx.dwMajorVersion & "." & xx.dwMinorVersion
    lblLicenseTo.Caption = IIf(Lang, "Licensed to:" & StrT, "授权给：" & StrT)
    lblVersion.Caption = IIf(Lang, "Version:", "版本:") & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = IIf(Lang, "Caling for 6000DDM", "6000表校准程序")
End Sub

Private Sub Deploy()
Dim RN As Long, tmp As String
lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，请稍后……")
delay 100
Open App.Path & "\Config.ini" For Output As #2

lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，读取软件通用设置……")
lblinfo.Refresh
Open App.Path & "\Configuration\软件通用设置.txt" For Input As #1
Print #2, "[Custom]"
Do While Not EOF(1)
Line Input #1, tmp
RN = RN + 1
If RN > 1 Then Print #2, Split(tmp, ";")(0)
delay 10
Loop
Close #1

RN = 0
lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，读取LAN设置……")
lblinfo.Refresh
Open App.Path & "\Configuration\LAN设置.txt" For Input As #1
Print #2, "[Lan]"
Do While Not EOF(1)
Line Input #1, tmp
RN = RN + 1
If RN > 1 Then Print #2, Split(tmp, ";")(0)
delay 10
Loop
Close #1
 

RN = 0
lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，读取RS232端口设置……")
lblinfo.Refresh
Open App.Path & "\Configuration\RS232端口设置.txt" For Input As #1
Print #2, "[Comm]"
Do While Not EOF(1)
Line Input #1, tmp
RN = RN + 1
If RN > 1 Then Print #2, Split(tmp, ";")(0)
delay 10
Loop
Close #1
 

RN = 0
lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，读取按键代码……")
lblinfo.Refresh
Open App.Path & "\Configuration\按键代码.txt" For Input As #1
Print #2, "[cmdstr]"
Do While Not EOF(1)
Line Input #1, tmp
RN = RN + 1
If RN > 1 Then Print #2, Split(tmp, ";")(0)
delay 10
Loop
Close #1
 

RN = 0
lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，读取修正系数……")
lblinfo.Refresh
Open App.Path & "\Configuration\修正系数.txt" For Input As #1
Print #2, "[Adjust]"
Do While Not EOF(1)
Line Input #1, tmp
RN = RN + 1
If RN > 1 Then Print #2, Split(tmp, ";")(0)
delay 10
Loop
Close #1

RN = 0
lblinfo.Caption = IIf(Lang, "", "程序正在进行配置，读取量程段编号……")
lblinfo.Refresh
Open App.Path & "\Configuration\量程段编号设置.txt" For Input As #1
Print #2, "[RangeID]"
Do While Not EOF(1)
Line Input #1, tmp
RN = RN + 1
If RN > 1 Then Print #2, Split(tmp, ";")(0)
delay 10
Loop
Close #1
 

Close #2

ReadINI (App.Path & "\Config.ini")

On Error Resume Next
    Load frmTip
    frmTip.show
    Unload Me
End Sub

Private Sub Timer1_Timer()
lblinfo.Visible = True
lblinfo.Caption = ""
Deploy
Timer1.Enabled = False
End Sub
