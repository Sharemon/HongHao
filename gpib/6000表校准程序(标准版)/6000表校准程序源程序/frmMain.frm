VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "6000DDM Calitor"
   ClientHeight    =   7875
   ClientLeft      =   9075
   ClientTop       =   1560
   ClientWidth     =   7830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7830
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer8 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   7440
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   240
      TabIndex        =   45
      Text            =   "RangeDDM"
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin DDM6000_Calitor.progressbar PB 
      Height          =   200
      Left            =   240
      TabIndex        =   42
      Top             =   3640
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12582912
      Max             =   750
      Color2          =   16685363
   End
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Left            =   3840
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7200
      Top             =   7440
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   200
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   15
      TabIndex        =   39
      Top             =   3700
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4320
      Top             =   7440
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   240
      TabIndex        =   36
      Top             =   3900
      Width           =   7455
      Begin VB.OptionButton Option2 
         Caption         =   "1.0滤波"
         Height          =   375
         Index           =   2
         Left            =   6000
         TabIndex        =   52
         Top             =   760
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1/2滤波"
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   51
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "1/4滤波"
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   50
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check3 
         Caption         =   "AutoCali"
         Height          =   375
         Left            =   480
         TabIndex        =   44
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "ZeroCali"
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton WK2 
         BackColor       =   &H0000C000&
         Caption         =   "WK2"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton WK1 
         BackColor       =   &H0000C000&
         Caption         =   "WK1"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Timer TimePB 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4800
      Top             =   7440
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   5280
      Top             =   7440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Index           =   6
      Left            =   6360
      TabIndex        =   15
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Index           =   3
      Left            =   5160
      TabIndex        =   13
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   7440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6240
      Top             =   7440
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6720
      Top             =   7440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Index           =   4
      Left            =   2760
      TabIndex        =   14
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   12
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Index           =   1
      Left            =   1560
      TabIndex        =   11
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "参数及状态信息"
      Height          =   1575
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   7455
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "校准数据"
         Height          =   180
         Index           =   4
         Left            =   5160
         TabIndex        =   54
         Top             =   360
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "端口状态"
         Height          =   180
         Index           =   0
         Left            =   720
         TabIndex        =   53
         Top             =   360
         Width           =   720
      End
      Begin VB.Shape shpdataoff 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4800
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CaliMode"
         Height          =   180
         Index           =   5
         Left            =   4800
         TabIndex        =   43
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Key Length"
         Height          =   180
         Index           =   4
         Left            =   4800
         TabIndex        =   35
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Protocol"
         Height          =   180
         Index           =   3
         Left            =   2640
         TabIndex        =   23
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Local Port"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   22
         Top             =   1200
         Width           =   900
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Remote Port"
         Height          =   180
         Index           =   1
         Left            =   2640
         TabIndex        =   21
         Top             =   840
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "IP ADDR"
         Height          =   180
         Index           =   0
         Left            =   360
         TabIndex        =   20
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "错误提示"
         Height          =   180
         Index           =   3
         Left            =   4800
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "发送状态"
         Height          =   180
         Index           =   2
         Left            =   3600
         TabIndex        =   18
         Top             =   360
         Width           =   720
      End
      Begin VB.Shape ShpErrOff 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4320
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape ShpSendOff 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3240
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape ShpErrOn 
         BorderColor     =   &H00C00000&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4320
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape ShpSendOn 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3240
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "读取状态"
         Height          =   180
         Index           =   1
         Left            =   2160
         TabIndex        =   17
         Top             =   360
         Width           =   720
      End
      Begin VB.Shape ShpRevOff 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1800
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape ShpRevOn 
         BorderColor     =   &H000000FF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1800
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape ShpComOff 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   360
         Top             =   360
         Width           =   255
      End
      Begin VB.Shape ShpComOn 
         BorderColor     =   &H000000FF&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   360
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape shpdataon 
         BorderColor     =   &H00000000&
         FillColor       =   &H00C000C0&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   4800
         Top             =   360
         Width           =   255
      End
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   1560
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   960
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
      InputMode       =   1
   End
   Begin MSComDlg.CommonDialog so 
      Left            =   480
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm0 
      Left            =   2160
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   840
         TabIndex        =   46
         Top             =   2000
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "to6000"
         Height          =   375
         Left            =   6240
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "frmMain.frx":0442
         Left            =   0
         List            =   "frmMain.frx":0444
         TabIndex        =   30
         Text            =   "Digits"
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   0
         TabIndex        =   29
         Text            =   "Range"
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   16
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   6
         Left            =   6240
         TabIndex        =   8
         Top             =   2160
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   5
         Left            =   5760
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   4
         Left            =   4680
         TabIndex        =   6
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   3
         Left            =   3600
         TabIndex        =   5
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   2
         Left            =   2520
         TabIndex        =   4
         Top             =   2760
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox LblErr 
         BorderStyle     =   0  'None
         Height          =   270
         Left            =   3000
         TabIndex        =   32
         Top             =   2000
         Width           =   2790
      End
      Begin VB.TextBox TextDisp 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1455
         HideSelection   =   0   'False
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   5775
      End
      Begin VB.Label lblCali 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   4920
         TabIndex        =   49
         Top             =   2400
         Width           =   90
      End
      Begin VB.Shape shpZero 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   2640
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblZero 
         AutoSize        =   -1  'True
         Caption         =   "Zero"
         Height          =   180
         Left            =   3000
         TabIndex        =   48
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label lblAuto 
         AutoSize        =   -1  'True
         Caption         =   "Auto"
         Height          =   180
         Left            =   4080
         TabIndex        =   47
         Top             =   2400
         Width           =   360
      End
      Begin VB.Shape ShpAuto 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3720
         Top             =   2400
         Width           =   255
      End
      Begin VB.Shape ShpStable 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1440
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblStable 
         AutoSize        =   -1  'True
         Caption         =   "Stable"
         Height          =   180
         Left            =   1800
         TabIndex        =   40
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label LblNul 
         AutoSize        =   -1  'True
         Caption         =   "Null"
         Height          =   180
         Left            =   720
         TabIndex        =   31
         Top             =   2400
         Width           =   360
      End
      Begin VB.Shape shpNull 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   360
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1000V"
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   4
         Left            =   6600
         TabIndex        =   28
         Top             =   1800
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "100 V"
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   3
         Left            =   6600
         TabIndex        =   27
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "10  V"
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   2
         Left            =   6600
         TabIndex        =   26
         Top             =   1320
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "1   V"
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   1
         Left            =   6600
         TabIndex        =   25
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "100mV"
         ForeColor       =   &H00808080&
         Height          =   180
         Index           =   0
         Left            =   6600
         TabIndex        =   24
         Top             =   840
         Width           =   450
      End
      Begin VB.Shape shpRange 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   4
         Left            =   6240
         Top             =   1800
         Width           =   255
      End
      Begin VB.Shape shpRange 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   6240
         Top             =   1560
         Width           =   255
      End
      Begin VB.Shape shpRange 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   6240
         Top             =   1320
         Width           =   255
      End
      Begin VB.Shape shpRange 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   6240
         Top             =   1080
         Width           =   255
      End
      Begin VB.Shape shpRange 
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H00808080&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   6240
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "安捷伦34410A数字万用表"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1560
         TabIndex        =   1
         Top             =   240
         Width           =   3510
      End
   End
   Begin VB.Label Stb 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   240
      TabIndex        =   33
      Top             =   7560
      Width           =   90
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Quit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuRangeID 
         Caption         =   "RangeID"
      End
      Begin VB.Menu Option1 
         Caption         =   "Option1"
         Begin VB.Menu mnuKEIT 
            Caption         =   "Keithley"
         End
         Begin VB.Menu mnuAgi 
            Caption         =   "Agilent"
         End
      End
      Begin VB.Menu mnuKEI 
         Caption         =   "mnuKEI"
      End
      Begin VB.Menu mnuComset 
         Caption         =   "Comm Set"
      End
      Begin VB.Menu mnuLanset 
         Caption         =   "Lan Set"
      End
      Begin VB.Menu mnuTheme 
         Caption         =   "Theme"
         Begin VB.Menu mnuTh1 
            Caption         =   "aero.she"
         End
         Begin VB.Menu mnuTh2 
            Caption         =   "chine.she"
         End
         Begin VB.Menu mnuTh3 
            Caption         =   "longhorn.she"
         End
         Begin VB.Menu mnuTH 
            Caption         =   "Look for a Theme"
         End
      End
      Begin VB.Menu mnuTol 
         Caption         =   "Tollerence"
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Filter"
      End
      Begin VB.Menu mnuAdjust 
         Caption         =   "Adjustment"
      End
      Begin VB.Menu mnuDDigits 
         Caption         =   "DDigits"
         Begin VB.Menu mnu4 
            Caption         =   "4"
         End
         Begin VB.Menu mnu5 
            Caption         =   "5"
         End
         Begin VB.Menu mnu6 
            Caption         =   "6"
         End
         Begin VB.Menu mnu7 
            Caption         =   "7"
         End
         Begin VB.Menu mnu8 
            Caption         =   "8"
         End
         Begin VB.Menu mnu9 
            Caption         =   "9"
         End
      End
   End
   Begin VB.Menu mnuFunc 
      Caption         =   "Functions"
      Begin VB.Menu mnuNul 
         Caption         =   "Null"
      End
      Begin VB.Menu mnuPeak 
         Caption         =   "Peak"
      End
      Begin VB.Menu mnuDisp 
         Caption         =   "Display"
      End
      Begin VB.Menu mnuDigit 
         Caption         =   "Digit"
      End
      Begin VB.Menu mnuRange 
         Caption         =   "Range"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuReset 
         Caption         =   "Reset"
      End
   End
   Begin VB.Menu mnuLanguage 
      Caption         =   "Language"
      Begin VB.Menu mnuCh 
         Caption         =   "中文"
      End
      Begin VB.Menu mnuEn 
         Caption         =   "English"
      End
   End
   Begin VB.Menu mnuHotKey 
      Caption         =   "RebootHotkey"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuCurve 
      Caption         =   "Curve"
   End
   Begin VB.Menu mnuHlelp 
      Caption         =   "Help"
      Begin VB.Menu mnuInstruc 
         Caption         =   "Instructions"
      End
      Begin VB.Menu mnuTips 
         Caption         =   "Tips"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Private ZeroCali As Boolean
Private NPLC As Double
Private Digits As Long
Private numArray() As Variant
Private num As String, LastNum As Double
Private twc As Boolean
Private SendOnce As Boolean
Private StableBool As Boolean
Private ErrTrackNum As Integer

Private moveTrue As Boolean
Private x1 As Double
Private y1 As Double

Private Const CCM_FIRST = &H2000
Private Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Private Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Private Const WM_USER = &H400
Private Const PBM_SETBARCOLOR = (WM_USER + 9)
Const CB_SHOWDROPDOWN = &H14F

Private Sub Check1_Click()
OutPut = CBool(IIf(CBool(Check1.Value), 1, 0))
modLang
If Check1.Value = 0 Then
    PB.Visible = False
    Picture1.Visible = False
    showhideWK False
Else
    showhideWK True
    PB.Value = 0
End If

        Label3(0).Caption = Label3(0).Caption & Instru
        Label3(1).Caption = Label3(1).Caption & Port2 & CommState(MSComm)
        Label3(2).Caption = Label3(2).Caption & Port0 & CommState(MSComm0)
        Label3(3).Caption = Label3(3).Caption & Port1 & CommState(MSComm1)
        Label3(5).Caption = Label3(5).Caption & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
        
End Sub

Private Sub Check2_Click()
ZeroCali = Check2.Value
End Sub

Private Sub Check3_Click()
AutoCali(1) = Not AutoCali(1)
For I = 1 To 5
    If RANGe = RangArry(I) Then
        DelayTimeW = DelayTime(I)
        ToleranceW = Tolerance(I)
        ZeroTolerW = ZeroToler(I)
        DepartTolerW = DepartToler(I)
    End If
Next I
    Label3(5).Caption = IIf(Lang, "CaliMode:", "校准方式：") & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
If AutoCali(1) = False Then ShpAuto.FillColor = &H80&
If AutoCali(1) = True And AutoCali(3) = True Then ShpAuto.FillColor = vbRed
If AutoCali(1) = True And AutoCali(2) = True And AutoCali(3) = True Then
    lblCali.Caption = IIf(Lang, "Calibrating", "正在校准")
Else
    lblCali.Caption = ""
End If
End Sub

Private Sub Combo1_Click()
On Error Resume Next
If Check1.Value = 0 Then
    DoEvents
    RANGe = 0.1 * (10 ^ Combo1.ListIndex)
Else
Select Case Combo1.ListIndex
Case 0, 1, 2
    RANGe = 0.1
Case 3, 4, 5, 6
    RANGe = 10 ^ (Combo1.ListIndex - 3)
Case 7
    twc = True
    Combo1.RemoveItem 7
    Combo1.AddItem IIf(Lang, "standard instrument range", "以下为标准表量程")
    Combo1.AddItem "100mV"
    Combo1.AddItem "1   V"
    Combo1.AddItem "10  V"
    Combo1.AddItem "100 V"
    Combo1.AddItem "1000V"
    Combo1.AddItem IIf(Lang, "Range ID settings", "更改量程段编号")
    Dim ret
    ret = SendMessage(Combo1.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&)
    Set ret = Nothing
Case 8, 9, 10, 11, 12
    RANGe = 10 ^ (Combo1.ListIndex - 9)
Case 13
    mnuRangeID_Click
End Select

If Combo1.ListIndex <= 6 Then
    ReDim cmdstr(1 To 10)
    cmdstrs = Split("37,48,48,59,49,50,59,48," & CStr((48 + RangeID(Combo1.ListIndex + 1))) & ",13", ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    MSComm.OutPut = cmdstr
    RangeIDIndex = Combo1.ListIndex + 1
Select Case Combo1.ListIndex
Case 0
    Text1.Text = IIf(Lang, "1mV#" & RangeID(1), "校准表量程:1mV#" & Format(RangeID(1), "00"))
Case 1
    Text1.Text = IIf(Lang, "10mV#" & RangeID(1), "校准表量程:10mV#" & Format(RangeID(2), "00"))
Case 2
    Text1.Text = IIf(Lang, "100mV#" & RangeID(1), "校准表量程:100mV#" & Format(RangeID(3), "00"))
Case 3
    Text1.Text = IIf(Lang, "1V#" & RangeID(1), "校准表量程:1V#" & Format(RangeID(4), "00"))
Case 4
    Text1.Text = IIf(Lang, "10V#" & RangeID(1), "校准表量程:10V#" & Format(RangeID(5), "00"))
Case 5
    Text1.Text = IIf(Lang, "100V#" & RangeID(1), "校准表量程:100V#" & Format(RangeID(6), "00"))
Case 6
    Text1.Text = IIf(Lang, "1000V#" & RangeID(1), "校准表量程:1000V#" & Format(RangeID(7), "00"))
End Select
End If

End If

If twc = False Then

Combo1.Visible = False
Command1(5).Visible = True

    If Instru = 0 And Winsock1.State = sckConnected Then
        WriteString Winsock1, "VOLT:DC:RANG " & RANGe
        delay 10
        WriteString Winsock1, "VOLT:DC:RANG?"
        RANGe = Val(ReadString(Winsock1))
        Debug.Print RANGe
    Else
        If MSComm0.PortOpen = True Then MSComm0.OutPut = ":VOLTage:DC:RANGe " & RANGe & vbCr
    End If
    For I = 0 To 4
        shpRange(I).FillColor = 8421504
        Label4(I).ForeColor = 8421504
        If RANGe = RangArry(I + 1) Then
            Filter = FilterArry(I + 1)
            DDigits = DDigitsArry(I + 1)
            shpRange(I).FillColor = ShpErrOn.FillColor
            Label4(I).ForeColor = vbBlack
            DelayTimeW = DelayTime(I + 1)
            ToleranceW = Tolerance(I + 1)
            ZeroTolerW = ZeroToler(I + 1)
            DepartTolerW = DepartToler(I + 1)
        End If
    Next
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
ReadAuto = True
Timer6.Enabled = True
Else
twc = False
End If
For I = 0 To 2
    If Option2(I).Value = True Then Option2_Click (I)
Next I
End Sub

Private Sub Combo1_GotFocus()
SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

Private Sub Combo2_Click()
On Error Resume Next
If Instru = 0 And Winsock1.State = sckConnected Then
    WriteString Winsock1, "VOLT:DC:NPLC?"
    NPLC = Val(ReadString(Winsock1))
    Select Case Combo2.ListIndex
    Case 0
        WriteString Winsock1, "VOLT:DC:NPLC 1"
        Digits = 6
    Case 1
        WriteString Winsock1, "VOLT:DC:NPLC 10"
        Digits = 7
    End Select
Else
    If MSComm0.PortOpen = True Then MSComm0.OutPut = ":VOLTage:DC:DIGits " & (Combo2.ListIndex + 4) & vbCr
    Digits = Combo2.ListIndex + 4
End If


ChangeDigits (Combo2.ListIndex)

Combo2.Visible = False
Command1(4).Visible = True

ReadAuto = True

Timer6.Enabled = True

End Sub

Private Sub Combo2_GotFocus()
Dim ret
ret = SendMessage(Combo2.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&)
Set ret = Nothing
End Sub

Private Sub Combo3_Click()
RANGe = 10 ^ (Combo3.ListIndex - 1)

Combo3.Visible = False
Command2(4).Visible = True

    If Instru = 0 And Winsock1.State = sckConnected Then
        WriteString Winsock1, "VOLT:DC:RANG " & RANGe
        WriteString Winsock1, "VOLT:DC:RANG?"
        RANGe = Val(ReadString(Winsock1))
    Else
        If MSComm0.PortOpen = True Then MSComm0.OutPut = ":VOLTage:DC:RANGe " & RANGe & vbCr
    End If
    For I = 0 To 4
        shpRange(I).FillColor = 8421504
        Label4(I).ForeColor = 8421504
        If RANGe = RangArry(I + 1) Then
            Filter = FilterArry(I + 1)
            DDigits = DDigitsArry(I + 1)
            shpRange(I).FillColor = ShpErrOn.FillColor
            Label4(I).ForeColor = vbBlack
            DelayTimeW = DelayTime(I + 1)
            ToleranceW = Tolerance(I + 1)
            ZeroTolerW = ZeroToler(I + 1)
            DepartTolerW = DepartToler(I + 1)
        End If
    Next
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
ReadAuto = True
Timer6.Enabled = True
End Sub

Private Sub Combo3_GotFocus()
SendMessage Combo3.hwnd, CB_SHOWDROPDOWN, True, ByVal 0&
End Sub

Public Sub Command1_Click(Index As Integer)
cmd1_Click Index
End Sub

Public Sub Command2_Click(Index As Integer)
cmd2_Click (Index)
End Sub

Private Sub Command2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Stb.Caption = IIf(Lang, "Connect with your instrument by LAN", "打开与仪器的连接")
Case 1
Stb.Caption = IIf(Lang, "Enable the automatic calibration Function", "开启自动校准功能")
Case 2
Stb.Caption = IIf(Lang, "Increase displaying digits of the software", "增加软件显示的位数")
Case 3
Stb.Caption = IIf(Lang, "Decrease displaying digits of the 6000DDM", "减少被较表显示的位数")
Case 4
Stb.Caption = IIf(Lang, "Decrease displaying digits of the software", "减少软件显示的位数")
Case 5
Stb.Caption = IIf(Lang, "", "")
Case 6
Stb.Caption = IIf(Lang, "Increase displaying digits of the 6000DDM", "增加被较表显示的位数")
End Select
End Sub

Private Sub Form_Load()
On Error GoTo errhdl
NulOn = True
AutoCali(1) = False
AutoCali(2) = False
AutoCali(3) = False
moveTrue = False
RANGe = 1
Timer6.Interval = CLng(IIf(Instru, 50, 20))
modLang
mnuKEI.Visible = CLng(Instru)
    MSComm0.CommPort = Port0
    MSComm.CommPort = Port2
    MSComm1.CommPort = Port1
        
    Check1.Visible = CommOut
    If CommOut = True Then
        Check1.Value = 1
        Check1_Click
    Else
        showhideWK False
    End If
    Check1.Visible = False
Unload frmTip
Load frmData
frmData.show
Timer8.Interval = Val(GetIni("Custom", "timer8", 100, App.Path & "\Config.ini"))
Timer8.Enabled = True
ErrTrackNum = 1
MSComm.PortOpen = CommOut
ErrTrackNum = 2
If CommIn = True Then MSComm1.PortOpen = True
errhdl: If err.Number = 8002 Then PortErrHdl
End Sub

Private Sub Form_LostFocus()
'DeleteHotKey
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
x1 = X
y1 = Y
If x1 > Frame1.Left + Frame1.Width Then
    If frmData.Visible = False Then
        frmData.Visible = True
    Else
        frmData.Visible = False
    End If
Else
    moveTrue = True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If moveTrue = True Then
    FrmMain.Move X - x1 + Me.Left, Y - y1 + Me.Top
    For Each frm In Forms
    If frm.name = "frmData" Then
        frm.Left = FrmMain.Left + FrmMain.Width
        frm.Top = FrmMain.Top
    ElseIf frm.name = "frmCurve" Then
        frm.Left = FrmMain.Left - frmCurve.Width
        frm.Top = FrmMain.Top + FrmMain.Height / 2 - frm.Height / 2
    End If
    Next
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
moveTrue = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'DeleteHotKey
    SaveINI (App.Path & "\Config.ini")
    WriteINI "UnloadMode", "Unload", "1", App.Path & "\Config.ini"
    ShellExecute Me.hwnd, "Open", App.Path & "\Reboot.exe", 0, 0, 0
    Busy = False
    mnuExit_Click
End Sub

Private Sub Form_Resize()
    frmData.WindowState = FrmMain.WindowState
End Sub

Private Sub LblErr_GotFocus()
HideCaret LblErr.hwnd
End Sub

Private Sub mnu4_Click()
For I = 1 To 5
If RANGe = RangArry(I) Then DDigitsArry(I) = 4
Next
DDigits = 4
End Sub

Private Sub mnu5_Click()
For I = 1 To 5
If RANGe = RangArry(I) Then DDigitsArry(I) = 5
Next
DDigits = 5
End Sub

Private Sub mnu6_Click()
For I = 1 To 5
If RANGe = RangArry(I) Then DDigitsArry(I) = 6
Next
DDigits = 6
End Sub

Private Sub mnu7_Click()
For I = 1 To 5
If RANGe = RangArry(I) Then DDigitsArry(I) = 7
Next
DDigits = 7
End Sub

Private Sub mnu8_Click()
For I = 1 To 5
If RANGe = RangArry(I) Then DDigitsArry(I) = 8
Next
DDigits = 8
End Sub

Private Sub mnu9_Click()
For I = 1 To 5
If RANGe = RangArry(I) Then DDigitsArry(I) = 9
Next
DDigits = 9
End Sub

Private Sub mnuAbout_Click()
mnuInstruc_Click
End Sub

Private Sub mnuAdjust_Click()
Filter = 0
Adjust.show
End Sub

Private Sub mnuAgi_Click()
If Command2(0).Caption = "Disconnect" Or Command2(0).Caption = "断开" Then
MsgBox IIf(Lang, "Please disable cunrrent connection before further actions ", "请先断开连接再重新选择仪器！")
Exit Sub
Else
Timer6.Enabled = False
Instru = 0
mnuKEI.Visible = CLng(Instru)
langResize
modLang
Label3(0).Caption = Label3(0).Caption & IIf(CBool(Instru), Instru, Winsock1.RemoteHostIP)
Label3(1).Caption = Label3(1).Caption & IIf(CBool(Instru), Port2, Winsock1.RemotePort)
Label3(2).Caption = Label3(2).Caption & IIf(CBool(Instru), Port0, Winsock1.LocalPort)
Label3(3).Caption = Label3(3).Caption & IIf(CBool(Instru), Port1, IO_Protocol(Winsock1))
Timer6.Enabled = True
End If
End Sub

Private Sub mnuCh_Click()
Stb.Caption = ""
Lang = 0
modLang
Label3(0).Caption = Label3(0).Caption & IIf(CBool(Instru), Instru, Winsock1.RemoteHostIP)
Label3(1).Caption = Label3(1).Caption & IIf(CBool(Instru), Port2 & CommState(MSComm), Winsock1.RemotePort)
Label3(2).Caption = Label3(2).Caption & IIf(CBool(Instru), Port0 & CommState(MSComm0), Winsock1.LocalPort)
Label3(3).Caption = Label3(3).Caption & IIf(CBool(Instru), Port1 & CommState(MSComm1), IO_Protocol(Winsock1))
Label3(5).Caption = Label3(5).Caption & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
For I = 0 To 2
Option2(I).Caption = Replace(Option2(I).Caption, "Filter", "滤波")
Next
FrmMain.Refresh
End Sub

Private Sub mnuComset_Click()
If MSComm.PortOpen = True Then MSComm.PortOpen = False
If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
Load frmComm
frmComm.show
End Sub

Private Sub mnuCurve_Click()
If frmCurve.Visible = False Then
frmCurve.show
frmCurve.Left = FrmMain.Left - frmCurve.Width
Else
Unload frmCurve
End If
End Sub

Private Sub mnuDigit_Click()
Command1_Click (3)
End Sub

Private Sub mnuDisp_Click()
Command1_Click (2)
End Sub

Private Sub mnuEn_Click()
Stb.Caption = ""
Lang = 1
modLang

Label3(0).Caption = Label3(0).Caption & IIf(CBool(Instru), Instru, Winsock1.RemoteHostIP)
Label3(1).Caption = Label3(1).Caption & IIf(CBool(Instru), Port2 & CommState(MSComm), Winsock1.RemotePort)
Label3(2).Caption = Label3(2).Caption & IIf(CBool(Instru), Port0 & CommState(MSComm0), Winsock1.LocalPort)
Label3(3).Caption = Label3(3).Caption & IIf(CBool(Instru), Port1 & CommState(MSComm1), IO_Protocol(Winsock1))
Label3(5).Caption = Label3(5).Caption & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
FrmMain.Refresh
For I = 0 To 2
Option2(I).Caption = Replace(Option2(I).Caption, "滤波", "Filter")
Next
End Sub

Private Sub mnuExit_Click()
    OutPut = False
    ReadAuto = False
    Check1_Click
    Check1_Click
    Timer6.Enabled = False
    If MSComm0.PortOpen = True Then MSComm0.PortOpen = False
    If MSComm.PortOpen = True Then MSComm.PortOpen = False
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    If Busy = False Then
    Timer2.Enabled = False
    Timer3.Enabled = False
    Winsock1.CloseSck
    Set Winsock1 = Nothing
    Dim frm As Object
    For Each frm In Forms
    Unload frm
    Next
    Else
    FrmMain.LblErr.Text = IIf(Lang, "Busy now,Please do this later", "操作繁忙，请稍后执行该操作")
    End If
End Sub

Private Sub mnuExport_Click()
so.Filter = IIf(Lang, "Text Files(*.txt)", "文本文件(*.txt)") & "|*.txt|" & IIf(Lang, "Data Files(*.dat)", "数据文件(*.dat)") & "|*.dat|" & IIf(Lang, "INI Files(*.ini)", "Windows初始化文件(*.ini)") & "|*.ini"
so.FilterIndex = 3
so.ShowSave
If so.FileName <> "" Then SaveINI (so.FileName)
End Sub

Private Sub mnuFilter_Click()
frmFilter.show
End Sub

Private Sub mnuHotKey_Click()
AddHotkey
End Sub

Private Sub mnuImport_Click()
so.Filter = IIf(Lang, "Text Files(*.txt)", "文本文件(*.txt)") & "|*.txt|" & IIf(Lang, "Data Files(*.dat)", "数据文件(*.dat)") & "|*.dat|" & IIf(Lang, "INI Files(*.ini)", "Windows初始化文件(*.ini)") & "|*.ini"
so.FilterIndex = 3
so.ShowOpen
ReadINI (so.FileName)
End Sub

Private Sub mnuInstruc_Click()
MsgBox IIf(Lang, "Any question please contract BJHHFA LTD(83391064).", "有任何疑问请联系北京弘豪福安仪器有限公司（83391064）"), vbInformation
End Sub

Private Sub mnuKEI_Click()
'DeleteHotKey
Set KEI = New KEI
KEI.show
End Sub

Private Sub mnuKEIT_Click()
If Command2(0).Caption = "Disconnect" Or Command2(0).Caption = "断开" Then
MsgBox IIf(Lang, "Please disable cunrrent connection before further actions ", "请先断开连接再重新选择仪器！")
Exit Sub
Else
Timer6.Enabled = False
Instru = 1
mnuKEI.Visible = CLng(Instru)
langResize
modLang
Label3(0).Caption = Label3(0).Caption & IIf(CBool(Instru), Instru, Winsock1.RemoteHostIP)
Label3(1).Caption = Label3(1).Caption & IIf(CBool(Instru), Port2, Winsock1.RemotePort)
Label3(2).Caption = Label3(2).Caption & IIf(CBool(Instru), Port0, Winsock1.LocalPort)
Label3(3).Caption = Label3(3).Caption & IIf(CBool(Instru), Port1, IO_Protocol(Winsock1))
Timer6.Enabled = True
End If
End Sub

Private Sub mnuLanset_Click()
'DeleteHotKey
Lanset.show
End Sub

Private Sub mnuNul_Click()
Command1_Click (0)
End Sub

Private Sub mnuPeak_Click()
Command1_Click (1)
End Sub

Private Sub mnuPrint_Click()
Command1_Click (5)
End Sub

Private Sub mnuRange_Click()
Command1_Click (4)
End Sub

Private Sub mnuRangeID_Click()
Load frmRANGeID
frmRANGeID.show
End Sub

Private Sub mnuReset_Click()
Command1_Click (6)
End Sub

Private Sub mnuTH_Click()
so.Filter = IIf(Lang, "Theme Files(*.she)", "主题文件(*.she)") & "|*.she"
so.InitDir = App.Path
so.ShowOpen
If so.FileName <> "" Then
SkinH_Attach
SkinH_AttachEx so.FileName, ""
End If
End Sub

Private Sub mnuTh1_Click()
FileName = App.Path & "\Themes\aero.she"
If Dir(FileName) <> "" Then
SkinH_Attach
SkinH_AttachEx FileName, ""
End If
End Sub

Private Sub mnuTh2_Click()
FileName = App.Path & "\Themes\china.she"
If Dir(FileName) <> "" Then
SkinH_Attach
SkinH_AttachEx FileName, ""
End If
End Sub

Private Sub mnuTh3_Click()
FileName = App.Path & "\Themes\longhorn.she"
If Dir(FileName) <> "" Then
SkinH_Attach
SkinH_AttachEx FileName, ""
End If
End Sub

Private Sub mnuTips_Click()
'DeleteHotKey
ShowAtStartup = 1
Set frmTip = Nothing
Set frmTip = New frmTip
Load frmTip
frmTip.show
End Sub

Private Sub mnuTol_Click()
Set frmToler = New frmToler
Load frmToler
frmToler.show
End Sub

Private Sub MSComm0_OnComm()
On Error Resume Next
Timer1.Enabled = False
If ReadAuto = True Then
Select Case MSComm0.CommEvent
Dim sbyte() As Byte
Case comEvReceive
    sbyte = MSComm0.Input
    OnOff False, 2
    FrmMain.Timer3.Enabled = True
    For I = 0 To UBound(sbyte)
        num = num & Chr(sbyte(I))
        If sbyte(I) = 13 Then
            num = Replace(num, vbCr, "")
            If Len(CStr(num)) > 8 Or num Like "*9.9E*" Then
            If num Like "*9.9E*" Then
                StableBool = False
                FrmMain.Timer4 = True
                TextDisp.Text = "OVR.FLW"
                FrmMain.LblErr.Text = IIf(Lang, "Waring:overload indication,please change the RANGE!", "警告：数据过载，请更换量程！")
            Else
                Dim numTmp As String
                numTmp = FormatNumber(Avg, 12, vbTrue)
                DisplayInTextBox numTmp
                FrmMain.LblErr.Text = IIf(Lang, "Displaying Digits:", "显示位数:") & DDigits & "     " & IIf(Lang, "Cali Digits:", "被较表位数:") & Digits + 1
                If frmCurve.Visible = True Then updateCurve num
            End If
            End If
            num = ""
        End If
    Next
End Select
Timer6.Enabled = True
MSComm0.InBufferCount = 0
End If
End Sub

Private Sub MScomm1_OnComm()
Dim tempstr() As Byte, str1 As String
If MSComm1.CommEvent = comEvReceive Then
tempstr = MSComm1.Input
For I = 0 To UBound(tempstr)
    str1 = str1 & "," & CStr(Val(tempstr(I)))
Next I

str1 = Replace(str1, ",", "")
MSComm1.InBufferCount = 0
CompareSending str1, MSComm0
If (MSComm.PortOpen = True And OutPut = True) Then MSComm.OutPut = tempstr
End If
End Sub

Private Sub Option2_Click(Index As Integer)
    For I = 1 To 5
        If RANGe = RangArry(I) Then Filter = FilterArry(I)
    Next
    Select Case Index
    Case 0
        Filter = 1 / 4 * Filter
    Case 1
        Filter = 1 / 2 * Filter
    Case 2
        Filter = Filter
    End Select
    Cnt = 0
    ss = 0
    ReDim numArry(1 To Filter)
End Sub

Private Sub TextDisp_GotFocus()
HideCaret TextDisp.hwnd
End Sub

Private Sub TimePB_Timer()
On Error Resume Next
timenow = timeGetTime
TimeSpan = timenow - TimeStart
PB.Value = TimeSpan
Picture1.Visible = True
Picture1.Width = (PB.Value / PB.Max) * PB.Width
Label3(4).Caption = IIf(Lang, "Key length:" & TimeSpan & "ms", "按键时长：" & TimeSpan & "ms")
End Sub

Private Sub Timer1_Timer()
If MSComm0.PortOpen = True Then
    MSComm0.OutBufferCount = 0
    MSComm0.OutPut = ":READ?" & vbCr
    OnOff False, 1
    FrmMain.Timer2.Enabled = True
    Timer6.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
OnOff True, 2
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
OnOff True, 1
Timer3.Enabled = False
End Sub

Private Sub Initiate(Optional Running As Boolean)
ReadAuto = False
Timer6.Enabled = False
Timer2.Enabled = False
Timer3.Enabled = False
Digits = 7
If Instru = 0 Then
If Running = False Then
WriteString Winsock1, "CONF:VOLT:DC"
WriteString Winsock1, "VOLT:DC:NPLC 10"
WriteString Winsock1, "VOLT:DC:RANG:AUTO OFF"
WriteString Winsock1, "VOLT:DC:NULL:STAT OFF"
WriteString Winsock1, "VOLT:DC:RANG?"
RANGe = ReadNumber(Winsock1)
End If
Else
RANGe = 1
MSComm0.OutPut = ":CONF:VOLT:DC" & vbCr
MSComm0.OutPut = ":VOLT:DC:NPLCycles 1" & vbCr
MSComm0.OutPut = ":VOLT:DC:AVERAGE:State On" & vbCr
MSComm0.OutPut = ":VOLT:AVERAGE:TCONtrol Moving" & vbCr
MSComm0.OutPut = ":VOLTage:DC:DIGits " & Digits & vbCr
MSComm0.OutPut = ":VOLTage:DC:RANGe " & RANGe & vbCr
'MSComm0.OutPut = "Initiate" & vbCr
End If
For I = 0 To 4
    shpRange(I).FillColor = 8421504
    Label4(I).ForeColor = 8421504
    
    If RANGe = RangArry(I + 1) Then
    Filter = FilterArry(I + 1)
    ReDim numArry(1 To Filter)
    DDigits = DDigitsArry(I + 1)
    RangeIDIndex = I + 3
    shpRange(I).FillColor = ShpErrOn.FillColor
    Label4(I).ForeColor = vbBlack
    DelayTimeW = DelayTime(I + 1)
    ToleranceW = Tolerance(I + 1)
    Text1.Text = IIf(Lang, "1mV#" & RangeID(1), "校准表量程:" & RANGe & "V#" & Format(RangeID(I + 3), "00"))
    End If
    
Next
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
End Sub

Private Sub Timer4_Timer()
OnOff True, 4
Timer4.Enabled = False
End Sub

Private Function Avg() As Double
Dim Add As Double
LastNum = Val(num)

If Filter <> 0 Then
ss = ss + 1
If ss > Filter Then ss = 1
Cnt = Cnt + 1
numArry(ss) = num
If Cnt >= Filter Then Cnt = Filter

For I = 1 To Cnt
Add = Add + numArry(I)
Next
Avg = Add / IIf(Cnt < Filter, Cnt - 1, Cnt)
Else
Avg = num
End If

'滤波完毕
Add = 0

Sign = Avg

Avg = (Avg - Base) * MultCons / DivCons
Dim Zero As Boolean
Zero = (Avg > 0)
For j = 0 To 4
    If RANGe = RangArry(j + 1) Then
        If RANGe = 0.1 And Instru = 1 Then Avg = 1000 * Avg
        Avg = IIf(Zero, Avg * Adjnum(j + 1).POS, Avg * Adjnum(j + 1).Neg)
        Exit For
    End If
Next
'乘常数改正

Outbit = Avg

dataArry.Data.Add Outbit
dataArry.timenow.Add timeGetTime



For I = 1 To dataArry.Data.Count

If timeGetTime - dataArry.timenow.Item(I) > DelayTimeW * 1000 Then
dataArry.Data.Remove (I)
dataArry.timenow.Remove (I)
Else
Exit For
End If
Next I
'数据随时间的序列

If Abs(LastNum) > CDbl(IIf(Instru, RANGe, CDbl(IIf(RANGe = 0.1, 100, RANGe)))) * ZeroTolerW Then
    shpZero.FillColor = &H80&
Else
    shpZero.FillColor = vbRed
End If

StableBool = Stable(dataArry.Data)

If StableBool = True Then
    ShpStable.FillColor = ShpErrOn.FillColor
    
    '____________________________________________________
    
    If Abs(LastNum) > CDbl(IIf(Instru, RANGe, CDbl(IIf(RANGe = 0.1, 100, RANGe)))) * ZeroTolerW Then
        If SendOnce = False And AutoCali(1) = True And AutoCali(3) = True Then
            WK2_Click
            SendOnce = True
        End If
        
        ZeroCali = CBool(IIf(Check2.Value = 1, 1, 0))
        WK1.BackColor = &HC000&
    Else
        If ZeroCali = True And AutoCali(2) = True Then
            ZeroCali = False
            Timer7.Interval = IIf(DelayTimeW <= 3, DelayTimeW * 1000, 3000)
            Timer7.Enabled = True
        End If
    End If
    '___________________________________________________

    
Else
    Timer7.Enabled = False
    ZeroCali = CBool(IIf(Check2.Value = 1, 1, 0))
    ShpStable.FillColor = ShpErrOff.FillColor
    WK1.BackColor = &HC000&
    WK2.BackColor = &HC000&
    SendOnce = False
End If
'判断数据的稳定性

End Function

Private Sub AutoRead()
On Error Resume Next
Do While ReadAuto = True
DoEvents
Busy = True
WriteString Winsock1, "READ?"
num = ReadNumber(Winsock1)
Dim Zero As Boolean
Zero = (num > 0)
For I = 0 To 4
If RANGe = RangArry(I + 1) Then
num = IIf(Zero, num * Adjnum(I + 1).POS, num * Adjnum(I + 1).Neg)
Exit For
End If
Next
TextDisp.Text = FormatNumber$(Avg, DDigits, vbTrue)
TextDisp.Refresh
delay 1
Busy = False
Loop
End Sub

Private Sub Timer5_Timer()
Select Case longBtn
Case 128
    
Case 64
    Command1(1).BackColor = vbRed
    Command1(1).Refresh
    delay 200
    Command1(1).BackColor = &H8000000F
Case 32
    Command1(2).BackColor = vbRed
    Command1(2).Refresh
    delay 200
    Command1(2).BackColor = &H8000000F
Case 16
    'MsgBox "没有长按键功能", vbOKOnly
Case 8
    'MsgBox "没有长按键功能", vbOKOnly
Case 4
    Command1(5).BackColor = vbRed
    Command1(5).Refresh
    delay 200
    Command1(5).BackColor = &H8000000F
Case 2
    Command1(6).BackColor = vbRed
    Command1(6).Refresh
    delay 200
    Command1(6).BackColor = &H8000000F
End Select
Timer5.Enabled = False
TimePB.Enabled = False
End Sub

Private Sub Timer6_Timer()
On Error Resume Next
If ReadAuto = True Then
    If Instru = 0 And Winsock1.State = sckConnected Then
        WriteString Winsock1, "READ?"
        num = ReadNumber(Winsock1)
        If IsNumeric(num) Then
            Dim numTmp As String
            If num Like "*9.9E*" Then
                StableBool = False
                TextDisp.Text = "OVR.FLW"
                TextDisp.Refresh
                FrmMain.LblErr.Text = IIf(Lang, "Waring:overload indication,please change the RANGE!", "警告：数据过载，请更换量程！")
            Else
                numTmp = FormatNumber(Avg, 12, vbTrue)
                If frmCurve.Visible = True Then updateCurve numTmp
                DisplayInTextBox numTmp
                If Winsock1.State = sckConnected Then FrmMain.LblErr.Text = IIf(Lang, "Displaying Digits:", "显示位数:") & DDigits & "     " & IIf(Lang, "Cali Digits:", "被较表位数:") & Digits + 1
            End If
        End If
    Else
        MSComm0.OutBufferCount = 0
        MSComm0.OutPut = ":READ?" & vbCr
        OnOff False, 1
        FrmMain.Timer2.Enabled = True
        Timer6.Enabled = False
        Timer1.Enabled = True
    End If
End If
End Sub

Private Sub Timer7_Timer()
WK1_Click
Timer7.Enabled = False
End Sub

Private Sub Timer8_Timer()
If OutPut = True Then OutTo6000 (Outbit)
shpdataoff.Visible = Not shpdataoff.Visible
End Sub

Private Sub WK1_Click()
    ReDim cmdstr(1 To 19)
    cmdstrs = Split(GetIni("cmdstr", "zeroclick1", "255,48,50,59,48,48,59,48,49,13,10,37,48,48,59,48,54,59,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
        cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    Base = Sign
    MSComm.OutPut = cmdstr
    shpNull.FillColor = ShpErrOn.FillColor
    WK1.BackColor = vbRed
    AutoCali(3) = True
    If AutoCali(1) = True And AutoCali(3) = True Then
        ShpAuto.FillColor = vbRed
        If AutoCali(2) = True Then lblCali.Caption = IIf(Lang, "Calibrating", "正在校准")
    End If
    updatefrmData ("0")
    ZeroSendOnce = True
End Sub

Private Sub WK1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReDim cmdstr(1 To 19)
    cmdstrs = Split(GetIni("cmdstr", "WK1down", "255,48,50,59,48,48,59,48,56,13,10,37,48,48,59,48,54,59,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
MSComm.OutPut = cmdstr
AutoCali(2) = True
TimeStart = timeGetTime
TimePB.Enabled = True

PB.Visible = True
Picture1.Visible = True
Picture1.Width = 1
End Sub

Private Sub WK1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
TimePB.Enabled = False
PB.Visible = False
Picture1.Visible = False
If TimeSpan > 500 Then
ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "WK1up", "255,48,50,59,48,48,59,49,56,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
MSComm.OutPut = cmdstr
'updatefrmData ("0")
'ZeroSendOnce = True
End If
End Sub

Private Sub WK2_Click()
ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "WK2click", "255,48,50,59,48,48,59,48,57,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
MSComm.OutPut = cmdstr
Dim s As String
If dataArry.Data.Count >= 1 Then
s = FormatNumber(dataArry.Data.Item(dataArry.Data.Count), 10, vbTrue)
frmData.Visible = True
updatefrmData (s)
Else
If MSComm.PortOpen = True Then MSComm.OutPut = Trim("%00;05;10" & vbCr)
End If
WK2.BackColor = vbRed
End Sub

Private Function CommState(Comm As MSComm) As String
CommState = "(" & IIf(Comm.PortOpen, IIf(Lang, "Active", "工作中"), IIf(Lang, "Idle", "空闲")) & ")"
End Function

Private Sub CompareSending(str As String, Comm As MSComm)
If MSComm.PortOpen = True Then
Comm.OutBufferCount = 0

If str Like "*545913*" Then
Comm.OutPut = ":VOLTage:DC:REFerence:STATe On" & vbCr
Base = Sign
shpNull.FillColor = ShpErrOn.FillColor

ElseIf str Like "*555913*" Then
Comm.OutPut = ":VOLTage:DC:REFerence:STATe Off" & vbCr
Base = 0
shpNull.FillColor = ShpErrOff.FillColor
Base = 0


ElseIf str Like "*48594853*" Then
Comm.OutPut = ":VOLTage:DC:DIGits 4" & vbCr '选择了位数5

ElseIf str Like "*48594854*" Then Comm.OutPut = ":VOLTage:DC:DIGits 5" & vbCr   '选择了位数6

ElseIf str Like "*48594855*" Then Comm.OutPut = ":VOLTage:DC:DIGits 6" & vbCr   '选择了位数7

ElseIf str Like "*48594856*" Then Comm.OutPut = ":VOLTage:DC:DIGits 7" & vbCr   '选择了位数8

ElseIf str Like "*4950594848*" Then
Filter = FilterArry(1)
Cnt = 0
ss = 0
ReDim numArry(1 To Filter)
Comm.OutPut = ":VOLTage:DC:RANGe 0.1" & vbCr     '选择了量程00
For I = 0 To 4
    shpRange(I).FillColor = 8421504
    Label4(I).ForeColor = 8421504
Next
shpRange(0).FillColor = ShpErrOn.FillColor
Label4(0).ForeColor = vbBlack


ElseIf str Like "*4950594849*" Then
Filter = FilterArry(2)
Comm.OutPut = ":VOLTage:DC:RANGe 1" & vbCr       '选择了量程01
For I = 0 To 4
    shpRange(I).FillColor = 8421504
    Label4(I).ForeColor = 8421504
Next
shpRange(1).FillColor = ShpErrOn.FillColor
Label4(1).ForeColor = vbBlack

ElseIf str Like "*4950594850*" Then
Filter = FilterArry(3)
Comm.OutPut = ":VOLTage:DC:RANGe 10" & vbCr      '选择了量程02
For I = 0 To 4
    shpRange(I).FillColor = 8421504
    Label4(I).ForeColor = 8421504
Next
shpRange(2).FillColor = ShpErrOn.FillColor
Label4(2).ForeColor = vbBlack

Else
'MsgBox IIf(Lang, "Key information transmission failed, please try again", "按键信息传送失败，请重试")
End If
End If
End Sub

Public Sub cmd1_Click(Index As Integer, Optional Zero As Boolean)
On Error Resume Next
Timer5.Enabled = False
TimePB.Enabled = False
PB.Visible = False
Picture1.Visible = False
Busy = True
Select Case Index
Case 0
    Timer6.Enabled = False
    ReadAuto = False
    Dim Nul As Long
        NulOn = (shpNull.FillColor = ShpErrOn.FillColor)
        Nul = IIf(NulOn, 1, 0)
    
    If Nul = 0 Then
        'If Check1.Value = 1 Then
            ReDim cmdstr(1 To 19)
            cmdstrs = Split(GetIni("cmdstr", "zeroclick1", "255,48,50,59,48,48,59,48,49,13,10,37,48,48,59,48,54,59,13", App.Path & "\Config.ini"), ",")
            For I = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(I) = Val(cmdstrs(I - 1))
            Next
            MSComm.OutPut = cmdstr
    
        'End If
    
            Base = Sign
            shpNull.FillColor = ShpErrOn.FillColor
    Else
        
        'If Check1.Value = 1 Then
            ReDim cmdstr(1 To 19)
            cmdstrs = Split(GetIni("cmdstr", "zeroclick0", "255,48,50,59,48,48,59,48,49,13,10,37,48,48,59,48,55,59,13", App.Path & "\Config.ini"), ",")
            For I = LBound(cmdstr) To UBound(cmdstr)
                cmdstr(I) = Val(cmdstrs(I - 1))
            Next
            MSComm.OutPut = cmdstr
        'End If
            Base = 0
            shpNull.FillColor = ShpErrOff.FillColor
    End If
    
    ReadAuto = True
    Timer6.Enabled = True
    If Instru = 1 And MSComm0.PortOpen = True Then MSComm0.OutPut = ":FETCh?" & vbCr

    
Case 1
    'If Check1.Value = 1 Then
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "runclick", "255,48,50,59,48,48,59,48,50,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    MSComm.OutPut = cmdstr
    'End If
    

    ReadAuto = True
    If Instru = 0 Then
        Timer6.Enabled = True
    Else
        If MSComm0.PortOpen = True Then MSComm0.OutPut = ":FETCh?" & vbCr
    End If
    AutoCali(2) = False
    AutoCali(3) = False
    ShpAuto.FillColor = &H80&
Case 2
    'If Check1.Value = 1 Then
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "dispclick", "255,48,50,59,48,48,59,48,51,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr
    frmData.RText1.Text = "正向"
    frmData.RText2.Text = "负向"
    frmData.Label3.Caption = ""
    For I = 1 To AdjOut.Count
    AdjOut.Remove (AdjOut.Count)
    WK1.BackColor = &HC000&
    WK2.BackColor = &HC000&
    AutoCali(2) = False
    AutoCali(3) = False
    ShpAuto.FillColor = &H80&
    Next
    'End If
    
Case 3
    DoEvents
    ReadAuto = False
    Timer6.Enabled = False
    If Combo2.Visible = False Then
    Command1(4).Visible = False
    Combo2.Visible = True
    Combo2.Left = Command1(4).Left
    Combo2.Top = 2880
    Select Case Check1.Value
    Case 0
    Combo2.Clear
    If Instru = 0 Then
    Combo2.AddItem "5.5"
    Combo2.AddItem "6.5"
    Else
    Combo2.AddItem "4"
    Combo2.AddItem "5"
    Combo2.AddItem "6"
    Combo2.AddItem "7"
    End If
    Case 1
    Combo2.Clear
    Combo2.AddItem "5"
    Combo2.AddItem "6"
    Combo2.AddItem "7"
    Combo2.AddItem "8"
    End Select
    Combo2.Text = IIf(Lang, "Digits Selection", "位数选择")
    Combo2.SetFocus
    Else
    Combo2.Visible = False
    Command1(4).Visible = True
    End If
Case 4
    DoEvents
    ReadAuto = False
    Timer6.Enabled = False
    If Combo1.Visible = False Then
    Command1(5).Visible = False
    Combo1.Visible = True
    Combo1.Left = Command1(5).Left
    Combo1.Top = 2880
    Combo1.Clear
    Select Case Check1.Value
    Case 0
    Combo1.AddItem "100mV"
    For I = 0 To 4
    If I <> 0 Then Combo1.AddItem RangArry(I + 1) & "V"
    Next
    Case 1
    Combo1.AddItem "1  mV"
    Combo1.AddItem "10 mV"
    Combo1.AddItem "100mV"
    Combo1.AddItem "1   V"
    Combo1.AddItem "10  V"
    Combo1.AddItem "100 V"
    Combo1.AddItem "1000V"
    End Select
    
    Combo1.Text = IIf(Lang, "Range Selection", "量程选择")
    Combo1.SetFocus
    Else
    Combo1.Visible = False
    Command1(5).Visible = True
    End If

Case 5
    If MSComm.PortOpen = True Then MSComm.OutPut = Trim("%00;17;06" & vbCr)
    If frmData.Visible = True Then frmData.Label3.Caption = IIf(Lang, "Data has been saved", "数据已保存")
    AutoCali(2) = False
    AutoCali(3) = False
    ShpAuto.FillColor = &H80&
Case 6
    ReadAuto = False
    Timer6.Enabled = False
    Initiate True
    ReadAuto = True
    If Instru = 0 Then
        Timer6.Enabled = True
    Else
        If MSComm0.PortOpen = True Then MSComm0.OutPut = ":FETCh?" & vbCr
    End If
    If MSComm.PortOpen = True Then MSComm.OutPut = Trim("%00;17;07" & vbCr)
End Select
Timer5.Enabled = False
Timer6.Enabled = True
Busy = False
errhdl: If err.Number <> 0 Then Debug.Print "Error:" & err.Number & vbNewLine & "Error description:" & err.Description
End Sub

Private Sub Command1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd1_MouseDown (Index)
End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Stb.Caption = IIf(Lang, "Enable the Null mode", "开启Null模式")
Case 1
Stb.Caption = IIf(Lang, "Enable the Peak mode", "开启运行模式")
Case 2
Stb.Caption = IIf(Lang, "", "")
Case 3
Stb.Caption = IIf(Lang, "Change the decimal digits", "更改仪器小数位数")
Case 4
Stb.Caption = IIf(Lang, "Change the range of your instrument", "更改仪器量程")
Case 5
Stb.Caption = IIf(Lang, "Enable the PRINT function", "启用仪器打印键功能")
Case 6
Stb.Caption = IIf(Lang, "Reset the instrument", "复位仪器")
End Select
End Sub

Private Sub Command1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
cmd1_MouseUp (Index)
If AutoCali(1) = True Then
    lblCali.Caption = IIf(Lang, "Waiting", "等待校准")
Else
    lblCali.Caption = ""
End If
End Sub

Public Sub cmd2_Click(Index As Integer)
On Error GoTo errhdl
Select Case Index


    Case 0
LblErr.Text = ""
If Instru = 0 Then
    Dim I As Long
    
    If (Winsock1.State = sckClosed) Then
        ' Invoke the Connect method to initiate a connection.
        Winsock1.Connect RmHost, RemPort
    Else
        FrmMain.Command2(0).Caption = IIf(Lang, "Connect", "连接")
        Winsock1.CloseSck
        'Set winsock1 = Nothing
        ReadAuto = False
        Timer6 = False
        OnOff True, 0
        TextDisp.Text = "O.F.F      "
        LblErr.Text = IO_Status(Winsock1) & "..."
        Exit Sub
    End If
    
    DoEvents
    
    LblErr.Text = ""
    LblErr.Text = IO_Status(Winsock1) & "..."
    'Str$((i \ 10) + 1) & " - " &
    
    ' Test to see if connected and
    ' Wait until the connection is made and then write to the message text box
    Busy = True
    For I = 1 To 50
        delay 100
        'If (i \ 10) = CDbl(i) / 10 Then Me.txtError.SelText = Str$(i) & " - " & IO_Status(Winsock1) & vbcr
        If (I \ 10) = CDbl(I) / 10 Then LblErr.Text = IO_Status(Winsock1) & "..."
        'Me.txtError.Refresh
        If Winsock1.State = sckConnected Then
        LblErr.Text = IO_Status(Winsock1) & "..."
        For j = 18 To 21
        modLblLang Label3(j - 18), captions(j)
        Next
        modLang
        Label3(0).Caption = Label3(0).Caption & Winsock1.RemoteHostIP
        Label3(1).Caption = Label3(1).Caption & Winsock1.RemotePort
        Label3(2).Caption = Label3(2).Caption & Winsock1.LocalPort
        Label3(3).Caption = Label3(3).Caption & IO_Protocol(Winsock1)
        Label3(5).Caption = Label3(5).Caption & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
        FrmMain.Command2(0).Caption = IIf(Lang, "Disconnect", "断开")

        OnOff False, 0
        Initiate
        ReadAuto = True
        Timer6.Enabled = True
        Exit Sub
        End If
        DoEvents
    Next I
    Busy = False
    
Else
    If MSComm0.PortOpen = True Then
    ReadAuto = False
    Timer6.Enabled = False
    MSComm0.OutPut = ":SYSTem:LOCal" & vbCr
    End If
    ErrTrackNum = 3
    MSComm0.PortOpen = Not MSComm0.PortOpen
    FrmMain.Command2(0).Caption = IIf(MSComm0.PortOpen, IIf(Lang, "Disconnect", "断开"), IIf(Lang, "Connect", "连接"))
    If MSComm0.PortOpen = True Then
    MSComm0.InBufferCount = 0
    MSComm0.OutBufferCount = 0
    Initiate
        For j = 18 To 21
            modLblLang Label3(j - 18), captions(j)
        Next

        modLang
        Label3(0).Caption = Label3(0).Caption & Instru
        Label3(1).Caption = Label3(1).Caption & Port2 & CommState(MSComm)
        Label3(2).Caption = Label3(2).Caption & Port0 & CommState(MSComm0)
        Label3(3).Caption = Label3(3).Caption & Port1 & CommState(MSComm1)
        Label3(5).Caption = Label3(5).Caption & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
    ReadAuto = True
    Timer6.Enabled = True
    OnOff False, 0
    Else
    Timer6.Enabled = False
    Timer1.Enabled = False
    delay Timer5.Interval
    ReadAuto = False
    OnOff True, 0
    TextDisp.Text = "O.F.F      "
    LblErr.Text = IIf(Lang, "Connection Disabled", "已断开连接")
        modLang
        Label3(0).Caption = Label3(0).Caption & Instru
        Label3(1).Caption = Label3(1).Caption & Port2 & CommState(MSComm)
        Label3(2).Caption = Label3(2).Caption & Port0 & CommState(MSComm0)
        Label3(3).Caption = Label3(3).Caption & Port1 & CommState(MSComm1)
        Label3(5).Caption = Label3(5).Caption & IIf(AutoCali(1), IIf(Lang, "Auto", "自动"), IIf(Lang, "", "手动"))
    End If
    
End If

If Instru = 0 Then LblErr.Text = IIf(Lang, "Error:connection time out...", "错误:连接超时..."): Winsock1.CloseSck: Exit Sub

    Case 1
    
    DoEvents
    ReadAuto = False
    Timer6.Enabled = False
    If Combo3.Visible = False Then
    Command2(4).Visible = False
    Combo3.Visible = True
    Combo3.Left = Command2(4).Left
    Combo3.Top = 7000
    Combo3.Clear

    Combo3.AddItem "100mV"
    For I = 0 To 4
    If I <> 0 Then Combo3.AddItem RangArry(I + 1) & "V"
    Next
    
    
    Combo3.Text = IIf(Lang, "Standdard Range", "标准量程")
    Combo3.SetFocus
    Else
    Combo3.Visible = False
    Command2(2).Visible = True
    End If
    
    Case 2
    If DDigits < 9 Then
        DDigits = DDigits + 1
    Else
        DDigits = 9
    End If
    For I = 1 To 5
    If RANGe = RangArry(I) Then DDigitsArry(I) = DDigits
    Next
    
    Case 3
    If Digits > 4 Then
        Digits = Digits - 1
    Else
        Digits = 4
    End If
    ChangeDigits (Digits - 4)
    
    Case 4
    If DDigits > 4 Then
        DDigits = DDigits - 1
    Else
        DDigits = 4
    End If
    For I = 1 To 5
    If RANGe = RangArry(I) Then DDigitsArry(I) = DDigits
    Next
    
    Case 6
    If Digits < 7 Then
        Digits = Digits + 1
    Else
        Digits = 7
    End If
    ChangeDigits (Digits - 4)
    
End Select
errhdl: If err.Number = 8002 Then PortErrHdl
End Sub

Public Sub cmd1_MouseDown(Index As Integer)
Command1(Index).MousePointer = 15
On Error Resume Next
'If Check1.Value = 1 Then
Select Case Index
Case 0
    longBtn = 128
    
Case 1
    longBtn = 64
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "rundown", "255,48,50,59,48,48,59,48,50,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr
Case 2
    longBtn = 32
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "dispdown", "255,48,50,59,48,48,59,48,51,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr
Case 3
    longBtn = 16
Case 4
    longBtn = 8
Case 5
    longBtn = 4
Case 6
    longBtn = 2
    ReDim cmdstr(1 To 21)
    cmdstrs = Split(GetIni("cmdstr", "resetdown", "255,48,50,59,48,48,59,48,55,13,10,37,48,48,59,49,55,59,48,55,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr
End Select
If Index <> 0 And Index <> 3 And Index <> 4 Then
Timer5.Enabled = True
TimeStart = timeGetTime
TimePB.Enabled = True
PB.Visible = True
Picture1.Visible = True
Picture1.Width = 1
End If
'End If
End Sub

Public Sub cmd1_MouseUp(Index As Integer)
Command1(Index).MousePointer = 0
On Error Resume Next
 
If TimeSpan > 500 Then
Select Case Index
Case 1
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "runup", "255,48,50,59,48,48,59,49,50,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr

Case 2
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "dispup", "37,48,48,59,54,53,59,57,56,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr

Case 5
    ReDim cmdstr(1 To 11)
    cmdstrs = Split(GetIni("cmdstr", "printup", "255,48,50,59,48,48,59,49,54,13,10", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr
Case 6
End Select
End If
TimeSpan = 0
PB.Visible = False
Picture1.Visible = False
End Sub

Private Sub ChangeDigits(Index As Integer)
Select Case Index
Case 0     '5位
    Digits = 4
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits5", "37,48,48,59,49,48,59,48,53,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next

Case 1    '6位
    Digits = 5
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits6", "37,48,48,59,49,48,59,48,54,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next

Case 2    '7位
    Digits = 6
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits7", "37,48,48,59,49,48,59,48,55,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next

Case 3   '8位
    Digits = 7
    ReDim cmdstr(1 To 10)
    cmdstrs = Split(GetIni("cmdstr", "Digits8", "37,48,48,59,49,48,59,48,56,13", App.Path & "\Config.ini"), ",")
    For I = LBound(cmdstr) To UBound(cmdstr)
    cmdstr(I) = Val(cmdstrs(I - 1))
    Next
    
End Select
If MSComm.PortOpen = True Then MSComm.OutPut = cmdstr
End Sub

Private Sub DisplayInTextBox(numDisp As String)
Select Case RangeIDIndex
Case 1, 4
numDisp = CStr(Format(numDisp, IIf(numDisp > 0, "+0.000000000", "0.000000000")))
Case 2, 5
numDisp = CStr(Format(numDisp, IIf(numDisp > 0, "+00.00000000", "00.00000000")))
Case 3, 6
numDisp = CStr(Format(numDisp, IIf(numDisp > 0, "+000.0000000", "000.0000000")))
Case 7
numDisp = CStr(Format(numDisp, IIf(numDisp > 0, "+0000.000000", "0000.000000")))
End Select
Dim GetDigits As Long, numTmp As String
numTmp = Mid(numDisp, 2, DDigits + 1)
numTmp = Split(numTmp, ".")(1)
GetDigits = Len(numTmp)
numDisp = FormatNumber(numDisp, GetDigits, vbTrue)
TextDisp.Text = numDisp
TextDisp.Refresh
End Sub

Private Sub PortErrHdl()
Select Case ErrTrackNum
Case 1
    If MsgBox(IIf(Lang, "Port error,will you correct the port setting?", "预设的校准端口号为" & Port2 & ",但是该端口并不存在于本机上，是否重新设置该端口号？" & vbNewLine & "选择否将退出程序。"), vbYesNo) = vbYes Then mnuComset_Click
Case 2
    If MsgBox(IIf(Lang, "Port error,will you correct the port setting?", "预设的指令端口号为" & Port1 & ",但是该端口并不存在于本机上，是否重新设置该端口号？" & vbNewLine & "选择否将退出程序。"), vbYesNo) = vbYes Then mnuComset_Click
Case 3
    If MsgBox(IIf(Lang, "Port error,will you correct the port setting?", "预设的吉时利端口号为" & Port0 & ",但是该端口并不存在于本机上，是否重新设置该端口号？"), vbYesNo) = vbYes Then mnuKEI_Click
End Select
End Sub
