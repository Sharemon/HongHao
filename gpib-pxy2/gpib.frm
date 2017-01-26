VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "仪表通信工具"
   ClientHeight    =   8265
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8415
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   8415
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      TabIndex        =   52
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton clear 
      Caption         =   "清 零"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      TabIndex        =   51
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6240
      TabIndex        =   49
      Text            =   "稳定指示："
      Top             =   1750
      Width           =   975
   End
   Begin VB.TextBox steadyShow 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7320
      TabIndex        =   48
      Top             =   1800
      Width           =   375
   End
   Begin VB.ComboBox floatContent 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   38
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox insName 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   600
      TabIndex        =   37
      Text            =   "吉时利2000"
      Top             =   1730
      Width           =   1455
   End
   Begin VB.TextBox realShow 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3720
      TabIndex        =   36
      Text            =   "+0.00000000E+00"
      Top             =   1730
      Width           =   2410
   End
   Begin VB.TextBox realShow 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2160
      TabIndex        =   35
      Text            =   "0.00000000"
      Top             =   1730
      Width           =   1575
   End
   Begin VB.TextBox mask 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4725
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   280
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   9480
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "自定义按键"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   17
      Top             =   6240
      Width           =   7935
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   5950
         TabIndex        =   42
         Text            =   "4"
         Top             =   650
         Width           =   150
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   4390
         TabIndex        =   41
         Text            =   "3"
         Top             =   650
         Width           =   150
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2830
         TabIndex        =   40
         Text            =   "2"
         Top             =   650
         Width           =   150
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   1250
         TabIndex        =   39
         Text            =   "1"
         Top             =   650
         Width           =   150
      End
      Begin VB.CommandButton selfDefined 
         BackColor       =   &H00008000&
         Caption         =   "多功能"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   4
         Left            =   6480
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton selfDefined 
         BackColor       =   &H8000000D&
         Caption         =   "button4"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   3
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton selfDefined 
         BackColor       =   &H8000000D&
         Caption         =   "button3"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton selfDefined 
         BackColor       =   &H8000000D&
         Caption         =   "button2"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   1
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton selfDefined 
         BackColor       =   &H8000000D&
         Caption         =   "button1"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   0
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "参数设置"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   3240
      TabIndex        =   12
      Top             =   3480
      Width           =   4935
      Begin VB.CheckBox speed 
         Height          =   180
         Left            =   1800
         TabIndex        =   43
         Top             =   2040
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox outSwitch 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   33
         Top             =   1440
         Value           =   1  'Checked
         Width           =   375
      End
      Begin VB.OptionButton filterHalf 
         Caption         =   "1/4 滤波"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   3360
         TabIndex        =   32
         Top             =   2040
         Width           =   1455
      End
      Begin VB.OptionButton filterHalf 
         Caption         =   "1/2 滤波"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   31
         Top             =   1590
         Width           =   1455
      End
      Begin VB.OptionButton filterHalf 
         Caption         =   "1.0 滤波"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   30
         Top             =   1260
         Width           =   1455
      End
      Begin VB.OptionButton filterHalf 
         Caption         =   "2.0 滤波"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   29
         Top             =   810
         Width           =   1455
      End
      Begin VB.OptionButton filterHalf 
         Caption         =   "4.0 滤波"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3360
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox DCVIOChoice 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "gpib.frx":0000
         Left            =   1440
         List            =   "gpib.frx":000D
         TabIndex        =   27
         Text            =   "请先连接仪表"
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox range2 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   14
         Text            =   "请先连接仪表"
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "正常速度发送："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "修正数据输出："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label10 
         Caption         =   "测量项目："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "量程选择："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
   End
   Begin VB.Timer sdTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   9720
      Top             =   2760
   End
   Begin VB.CommandButton holdOn 
      Caption         =   "保持数据"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
   End
   Begin MSCommLib.MSComm spOut 
      Left            =   9480
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton inBitsDown 
      Caption         =   "仪表位数 -"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton inBitsUp 
      Caption         =   "仪表位数 +"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9000
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer stTime 
      Left            =   9360
      Top             =   2760
   End
   Begin VB.Timer rdTime 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   9000
      Top             =   2760
   End
   Begin VB.CommandButton function 
      Caption         =   "OUMU4"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton function 
      Caption         =   "DCI"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton function 
      Caption         =   "DCV"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton function 
      Caption         =   "置 零"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   1740
   End
   Begin MSCommLib.MSComm sp 
      Left            =   9000
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
   End
   Begin VB.CommandButton bitDown 
      Caption         =   "显示位数 -"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton bitUp 
      Caption         =   "显示位数 +"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10560
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "运行状态"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   240
      TabIndex        =   10
      Top             =   3480
      Width           =   2655
      Begin VB.CommandButton connect 
         Caption         =   "连 接"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   480
         Width           =   1875
      End
      Begin VB.ComboBox freq 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "gpib.frx":0020
         Left            =   1440
         List            =   "gpib.frx":0022
         TabIndex        =   47
         Text            =   "5"
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox bitChoice 
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   7.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "gpib.frx":0024
         Left            =   1440
         List            =   "gpib.frx":003A
         TabIndex        =   25
         Text            =   "8"
         Top             =   1530
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Hz"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   46
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "采样频率："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label interfereShow 
         Caption         =   "无"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   1005
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "接口选择："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   1040
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "显示位数："
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   1095
      End
   End
   Begin VB.TextBox modifiedShow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   56.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1740
      Left            =   480
      TabIndex        =   23
      Text            =   "000.0000"
      Top             =   345
      Width           =   7335
   End
   Begin VB.Label LabInfo 
      Caption         =   "提示信息"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   9015
   End
   Begin VB.Menu file 
      Caption         =   "文件"
      Begin VB.Menu save 
         Caption         =   "保存"
      End
      Begin VB.Menu open 
         Caption         =   "打开"
      End
      Begin VB.Menu exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu setting 
      Caption         =   "设置"
      Begin VB.Menu choice 
         Caption         =   "选择万用表"
         Index           =   0
         Begin VB.Menu keithley 
            Caption         =   "吉时利"
            Begin VB.Menu J2000 
               Caption         =   "2000"
            End
         End
         Begin VB.Menu agilent 
            Caption         =   "安捷伦"
            Begin VB.Menu A34401A 
               Caption         =   "34401A"
            End
            Begin VB.Menu A34410A 
               Caption         =   "34410A"
            End
         End
         Begin VB.Menu fluke 
            Caption         =   "福禄克"
            Begin VB.Menu A8846A 
               Caption         =   "8846A"
            End
         End
      End
      Begin VB.Menu IOChoiceEntry 
         Caption         =   "选择接口"
         Begin VB.Menu IOChoice 
            Caption         =   "GPIB"
            Index           =   0
         End
         Begin VB.Menu IOChoice 
            Caption         =   "串口"
            Index           =   1
         End
         Begin VB.Menu IOChoice 
            Caption         =   "LAN"
            Index           =   2
         End
      End
      Begin VB.Menu GPIBSetting 
         Caption         =   "GPIB设置"
      End
      Begin VB.Menu usartSetting 
         Caption         =   "串口设置"
      End
      Begin VB.Menu LANSetting 
         Caption         =   "LAN口设置"
      End
      Begin VB.Menu steadySetting 
         Caption         =   "稳定设置"
      End
      Begin VB.Menu filterSetting 
         Caption         =   "滤波设置"
      End
      Begin VB.Menu modifySetting 
         Caption         =   "修正设置"
      End
      Begin VB.Menu self 
         Caption         =   "自定义按键"
      End
   End
   Begin VB.Menu VIOChoice 
      Caption         =   "测量值选择"
      Visible         =   0   'False
      Begin VB.Menu DCVIO 
         Caption         =   "置零"
         Index           =   0
      End
      Begin VB.Menu DCVIO 
         Caption         =   "直流电压"
         Index           =   1
      End
      Begin VB.Menu DCVIO 
         Caption         =   "直流电流"
         Index           =   2
      End
      Begin VB.Menu DCVIO 
         Caption         =   "电阻"
         Index           =   3
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助"
      Begin VB.Menu test 
         Caption         =   "测试模块"
      End
      Begin VB.Menu about 
         Caption         =   "关于"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'不共享的变量
Private rangeN As Double
Private bits As Integer
Private instruName As String

'连接与测量相关
Dim status As Long
Dim defrm As Long
Dim vi As Long
Dim strRes As String * 200
Dim writeStr As String * 50
Dim wrFlag As Boolean
Dim outFlag As Boolean
Public connectFlag As Boolean
Dim recflag As Boolean   'sp's receive wait flag
Dim zeroFlag As Boolean
Public VIO As Integer
Dim memory() As Double
Dim memCnt As Integer 'memory's serial number
Dim fullFlag As Boolean
Dim steadyCnt As Long
Dim min, max, stdAverage As Double
Dim setMinMax As Boolean
Dim reviseArr(1, 19) As Double
Dim reviseNum As Integer
Dim sendStringR, sendStringM As String 'Out232 send string
Dim valTemp As Double  'u,m,k,g multiple
Dim bias As Double
Dim biasStandard As Double
Dim sdInter As Integer
Dim firstStd As Boolean

Private Sub delay_ms(time As Integer)
    Dim start As Double
    start = timeGetTime()
    While (timeGetTime() - start < time)
        DoEvents
    Wend
End Sub

'Private Sub Set_Revise()
'    Dim i As Integer
'    reviseNum = 0
'    For i = 0 To 19
'        If (valid(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), i) = 1) Then
'            reviseArr(0, reviseNum) = standard(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), i)
'            reviseArr(1, reviseNum) = realM(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), i)
'            reviseNum = reviseNum + 1
'        End If
'    Next
'End Sub

'修正滤波暂存器的大小
Public Sub Set_Memory()
    Dim temp As Integer ', size As Integer, i As Integer
    If (VIO = 0 Or VIO = 1 Or VIO = 2) Then
        temp = filter(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1)) * 0.5 ^ filterDiv * 4
        If (temp <= 0) Then temp = 1
        filterCur = temp
        Erase memory
        ReDim memory(0 To temp - 1)
        fullFlag = False
        memCnt = 0
    End If
End Sub

'命令读取
Private Sub readCommand(name As String)
    Dim strTemp As String
    Dim str() As String
    Dim Index As Integer
    Dim i, j, k As Integer
    Open App.Path & "\command.ini" For Input As #1
    
    Input #1, strTemp
    Do
    Input #1, strTemp
    str = Split(strTemp, ":", 2)
    If (str(0) = Left(name, 3)) Then
        Index = Val(str(1))
        'Exit Do
    End If
    Loop While (strTemp <> "//finish")
    
    'zero
    For i = 0 To 2
        Input #1, strTemp
        For j = 0 To Index
            For k = 0 To 3
                Input #1, strTemp
                zero(i, k) = strTemp
            Next
        Next
        Do
        Input #1, strTemp
        Loop While (strTemp <> "//finish")
    Next
    
    'digits
    For i = 0 To 2
        Input #1, strTemp
        For j = 0 To Index
            For k = 0 To 1
                Input #1, strTemp
                digits(i, k) = strTemp
            Next
        Next
        Do
        Input #1, strTemp
        Loop While (strTemp <> "//finish")
    Next
    
    Close #1
End Sub


'保存到ini
Private Sub save2ini(filename As String)
    Open filename For Output As #1
    Print #1, "[InstrumentName]"
    Print #1, "Name:" + instruName
    Print #1, "[IOSetting]"
    Print #1, "IO:" + str(dataDim.IOSetting)
    Print #1, "[GPIBSetting]"
    Print #1, "GPIBAddr:" + str(dataDim.GPIBaddr)
    Print #1, "[UsartSetting]"
    Print #1, "ComPort:" + str(dataDim.comport)
    Print #1, "BaudRate:" + str(dataDim.baudrate)
    Print #1, "DataBit:" + str(dataDim.databit)
    Print #1, "StopBit:" + str(dataDim.stopbit)
    Print #1, "cr:" + dataDim.cr
    Print #1, "[UsartSetting]"
    Print #1, "ComPort:" + str(dataDim.comPortOut)
    Print #1, "BaudRate:" + str(dataDim.baudRateOut)
    Print #1, "DataBit:" + str(dataDim.dataBitOut)
    Print #1, "StopBit:" + str(dataDim.stopBitOut)
    Print #1, "cr:" + dataDim.crOut
    Print #1, "OutEn:" + str(IIf(dataDim.outEn, 1, 0))
    Print #1, "[LANSetting]"
    Print #1, "IPAddr:" + dataDim.IPaddr
    Print #1, "IPPort:" + str(dataDim.IPport)
    Print #1, "localIPPort:" + str(dataDim.localIPPort)
    Print #1, "[SteadySetting]"
    Dim i As Integer, strTemp As String, j As Integer, k As Integer
    strTemp = ""
    For i = 0 To 35
        strTemp = strTemp + " " + Trim(str(tSteady(Fix(i / 12), i Mod 12)))
        If (i = 11 Or i = 23) Then
            strTemp = strTemp + vbCrLf + "        "
        End If
    Next
    Print #1, "tSteady:" + strTemp
    strTemp = ""
    For i = 0 To 35
        strTemp = strTemp + " " + Trim(str(aSteady(Fix(i / 12), i Mod 12, 0)))
        If (i = 11 Or i = 23) Then
            strTemp = strTemp + vbCrLf + "         "
        End If
    Next
    Print #1, "aSteady-:" + strTemp
    strTemp = ""
    For i = 0 To 35
        strTemp = strTemp + " " + Trim(str(aSteady(Fix(i / 12), i Mod 12, 1)))
        If (i = 11 Or i = 23) Then
            strTemp = strTemp + vbCrLf + "         "
        End If
    Next
    Print #1, "aSteady+:" + strTemp
    Print #1, "[FilterSetting]"
    strTemp = ""
    For i = 0 To 35
        strTemp = strTemp + " " + Trim(str(filter(Fix(i / 12), i Mod 12)))
        If (i = 11 Or i = 23) Then
            strTemp = strTemp + vbCrLf + "       "
        End If
    Next
    Print #1, "filter:" + strTemp
    Print #1, "filterHalf:" + str(filterDiv)
    Print #1, "[SelfDef]"
    For i = 0 To 7
        Print #1, "def" + Trim(str(i)) + ":" + dataDim.selfDef(i)
    Next
    For i = 0 To 7
        Print #1, "defName" + Trim(str(i)) + ":" + dataDim.selfName(i)
    Next
    For i = 0 To 7
        Print #1, "defValid" + Trim(str(i)) + ":" + IIf(dataDim.selfValid(i), "1", "0")
    Next
    For i = 0 To 7
        Print #1, "2def" + Trim(str(i)) + ":" + dataDim.selfDef2(i)
    Next
    For i = 0 To 7
        Print #1, "2defName" + Trim(str(i)) + ":" + dataDim.selfName2(i)
    Next
    For i = 0 To 7
        Print #1, "2defValid" + Trim(str(i)) + ":" + IIf(dataDim.selfValid2(i), "1", "0")
    Next
    Print #1, "[Range&Bits]"
    Print #1, "Range:" + str(rangeNum)
    Print #1, "Bits:" + str(bits)
    Print #1, "[Modify]"
    strTemp = ""
    For i = 0 To 2
        For j = 0 To 11
            For k = 0 To 1
                strTemp = strTemp + " " + Trim(str(dataDim.modiPara(i, j, k)))
            Next
            If (Not (i = 2 And j = 11)) Then
                strTemp = strTemp + vbCrLf + "         "
            End If
        Next
    Next
    Print #1, "modiPara:" + strTemp
    strTemp = ""
    For i = 0 To 2
        For j = 0 To 11
             strTemp = strTemp + " " + Trim(str(dataDim.valid(i, j)))
            If (Not (i = 2 And j = 11)) Then
                strTemp = strTemp + vbCrLf + "      "
            End If
        Next
    Next
    Print #1, "valid:" + strTemp
    strTemp = ""
    For i = 0 To 35
        strTemp = strTemp + " " + Trim(str(dataDim.multi(Fix(i / 12), i Mod 12)))
        If (i = 11 Or i = 23) Then
            strTemp = strTemp + vbCrLf + "      "
        End If
    Next
    Print #1, "multi:" + strTemp
    strTemp = ""
    For i = 0 To 35
        strTemp = strTemp + " " + Trim(str(dataDim.div(Fix(i / 12), i Mod 12)))
        If (i = 11 Or i = 23) Then
            strTemp = strTemp + vbCrLf + "    "
        End If
    Next
    Print #1, "div:" + strTemp
    Close #1
End Sub

'从文件打开
Private Sub openFromFile(filename As String)
    Dim strTemp As String
    Dim i As Integer, j As Integer
    Dim strSplit() As String
    
    Open filename For Input As #1
    Line Input #1, strTemp
    Line Input #1, strTemp
    instruName = Right(strTemp, Len(strTemp) - 5)
    Line Input #1, strTemp
    Line Input #1, strTemp
    dataDim.IOSetting = Val(Right(strTemp, Len(strTemp) - 4))
    Line Input #1, strTemp
    Line Input #1, strTemp
    dataDim.GPIBaddr = Val(Right(strTemp, Len(strTemp) - 10))
    Line Input #1, strTemp
    Line Input #1, strTemp
    dataDim.comport = Val(Right(strTemp, Len(strTemp) - 9))
    Line Input #1, strTemp
    dataDim.baudrate = Val(Right(strTemp, Len(strTemp) - 10))
    Line Input #1, strTemp
    dataDim.databit = Val(Right(strTemp, Len(strTemp) - 9))
    Line Input #1, strTemp
    dataDim.stopbit = Val(Right(strTemp, Len(strTemp) - 9))
    Line Input #1, strTemp
    dataDim.cr = Right(strTemp, Len(strTemp) - 3)
    Line Input #1, strTemp
    Line Input #1, strTemp
    dataDim.comPortOut = Val(Right(strTemp, Len(strTemp) - 9))
    Line Input #1, strTemp
    dataDim.baudRateOut = Val(Right(strTemp, Len(strTemp) - 10))
    Line Input #1, strTemp
    dataDim.dataBitOut = Val(Right(strTemp, Len(strTemp) - 9))
    Line Input #1, strTemp
    dataDim.stopBitOut = Val(Right(strTemp, Len(strTemp) - 9))
    Line Input #1, strTemp
    dataDim.crOut = Right(strTemp, Len(strTemp) - 3)
    Line Input #1, strTemp
    dataDim.outEn = IIf(Val(Right(strTemp, Len(strTemp) - 7)) = 1, True, False)
    Line Input #1, strTemp
    Line Input #1, strTemp
    dataDim.IPaddr = Right(strTemp, Len(strTemp) - 7)
    Line Input #1, strTemp
    dataDim.IPport = Val(Right(strTemp, Len(strTemp) - 8))
    Line Input #1, strTemp
    dataDim.localIPPort = Val(Right(strTemp, Len(strTemp) - 13))
    Line Input #1, strTemp
    For i = 0 To 2
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 9)
        strSplit() = Split(strTemp, " ", 12)
        For j = 0 To 11
            tSteady(i, j) = Val(strSplit(j))
        Next
    Next
    For i = 0 To 2
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 10)
        strSplit() = Split(strTemp, " ", 12)
        For j = 0 To 11
            aSteady(i, j, 0) = Val(strSplit(j))
        Next
    Next
    For i = 0 To 2
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 10)
        strSplit() = Split(strTemp, " ", 12)
        For j = 0 To 11
            aSteady(i, j, 1) = Val(strSplit(j))
        Next
    Next
    Line Input #1, strTemp
    For i = 0 To 2
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 8)
        strSplit() = Split(strTemp, " ", 12)
        For j = 0 To 11
            filter(i, j) = Val(strSplit(j))
        Next
    Next
    Line Input #1, strTemp
    filterDiv = Val(Right(strTemp, Len(strTemp) - 12))
    Line Input #1, strTemp
    For i = 0 To 7
        Line Input #1, strTemp
        dataDim.selfDef(i) = Right(strTemp, Len(strTemp) - 5)
    Next
    For i = 0 To 7
        Line Input #1, strTemp
        dataDim.selfName(i) = Right(strTemp, Len(strTemp) - 9)
    Next
    For i = 0 To 7
        Line Input #1, strTemp
        dataDim.selfValid(i) = Right(strTemp, Len(strTemp) - 10)
    Next
    For i = 0 To 7
        Line Input #1, strTemp
        dataDim.selfDef2(i) = Right(strTemp, Len(strTemp) - 6)
    Next
    For i = 0 To 7
        Line Input #1, strTemp
        dataDim.selfName2(i) = Right(strTemp, Len(strTemp) - 10)
    Next
    For i = 0 To 7
        Line Input #1, strTemp
        dataDim.selfValid2(i) = Right(strTemp, Len(strTemp) - 11)
    Next
    Line Input #1, strTemp
    Line Input #1, strTemp
    dataDim.rangeNum = Val(Right(strTemp, Len(strTemp) - 7))
    Line Input #1, strTemp
    bits = Val(Right(strTemp, Len(strTemp) - 6))
    Line Input #1, strTemp
    For i = 0 To 35
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 10)
        strSplit() = Split(strTemp, " ", 2)
        For j = 0 To 1
            dataDim.modiPara(Fix(i / 12), i Mod 12, j) = Val(strSplit(j))
        Next
    Next
    For i = 0 To 35
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 7)
        dataDim.valid(Fix(i / 12), i Mod 12) = Val(strTemp)
    Next
    For i = 0 To 2
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 7)
        strSplit() = Split(strTemp, " ", 12)
        For j = 0 To 11
            dataDim.multi(i, j) = Val(strSplit(j))
        Next
    Next
    For i = 0 To 2
        Line Input #1, strTemp
        strTemp = Right(strTemp, Len(strTemp) - 5)
        strSplit() = Split(strTemp, " ", 12)
        For j = 0 To 11
            dataDim.div(i, j) = Val(strSplit(j))
        Next
    Next
    Close #1
End Sub

Private Sub A34401A_Click()
    instruName = "安捷伦34401A"
    Me.insName.Text = "安捷伦34401A"
    Me.LANSetting.Visible = False
    Me.IOChoice(2).Visible = False
    If (dataDim.IOSetting = 2) Then dataDim.IOSetting = 1
    Call readCommand(instruName)
End Sub

Private Sub A34410A_Click()
    instruName = "安捷伦34410A"
    Me.insName.Text = "安捷伦34410A"
    Me.LANSetting.Visible = True
    Me.IOChoice(2).Visible = True
    Call readCommand(instruName)
End Sub

Private Sub A8846A_Click()
    instruName = "福禄克8846A"
    Me.insName.Text = "福禄克8846A"
    Me.LANSetting.Visible = True
    Me.IOChoice(2).Visible = True
    Call readCommand(instruName)
End Sub

Private Sub bitChoice_Click()
    bits = Val(Me.bitChoice.Text)
End Sub


Private Sub clear_Click()
'    Call wrIO(":READ?", True)
'    bias = Val(strRes)
    bias = biasStandard
    Me.modifiedShow.Text = ""
    Call Set_Memory
    zeroFlag = True
End Sub

Private Sub DCVIO_Click(Index As Integer)
    Call function_Click(Index)
End Sub

Private Sub DCVIOChoice_Click()
    Call function_Click(Me.DCVIOChoice.ListIndex + 1)
End Sub


Private Sub holdOn_Click()
    Me.holdOn.Enabled = False
    If (connectFlag) Then
        If (Me.rdTime.Enabled) Then
            While (wrFlag)
            Wend
            Me.rdTime.Enabled = False
            Me.stTime.Enabled = False
            Me.function(0).Enabled = False
'            Me.sdTimer.Interval = Me.rdTime.Interval
'            Me.sdTimer.Enabled = True
            If (dataDim.selfValid(6)) Then
                Call delay_ms(sdInter)
                Me.spOut.Output = dataDim.selfDef(6) + vbCrLf
            End If
            Me.holdOn.BackColor = vbGreen
        Else
            Me.rdTime.Enabled = True
            Me.stTime.Enabled = True
            Me.function(0).Enabled = True
'            Me.sdTimer.Enabled = False
            Me.holdOn.BackColor = &H8000000F
        End If
    End If
Me.holdOn.Enabled = True
End Sub

Private Sub IOChoice_Click(Index As Integer)
    If (connectFlag) Then
        While wrFlag
            DoEvents
        Wend
        If (dataDim.IOSetting = 1) Then
            If (Len(Me.sp.Input) <> 0) Then strbuf = Me.sp.Input
        End If
        Me.rdTime.Enabled = False
        Call closeDev
        dataDim.IOSetting = Index
        Call openDev
        Me.rdTime.Enabled = True
    Else
        dataDim.IOSetting = Index
    End If
    Select Case dataDim.IOSetting
    Case 0
        Me.interfereShow.Caption = "GPIB"
    Case 1
        Me.interfereShow.Caption = "串口"
    Case 2
        Me.interfereShow.Caption = "LAN"
    End Select
End Sub

Private Sub J2000_Click()
    instruName = "吉时利2000"
    Me.insName.Text = "吉时利2000"
    Me.LANSetting.Visible = False
    Me.IOChoice(2).Visible = False
    If (dataDim.IOSetting = 2) Then dataDim.IOSetting = 1
    Call readCommand(instruName)
End Sub

Private Sub bitDown_Click()
'    If (bits > 4) Then
'        bits = bits - 1
'        Me.bitsShow.Caption = Trim(str(bits))
'        If (Not connectFlag) Then
'            Dim strTemp As String, i As Integer
'            strTemp = "0."
'            For i = 0 To bits - 1
'                strTemp = strTemp + "0"
'            Next
'            Me.modifiedShow = strTemp
'            Me.realShow(0) = strTemp
'        End If
'    Else
'        MsgBox ("位数不能小于4")
'    End If
End Sub

Private Sub bitUp_Click()
'    If (bits < 9) Then
'        bits = bits + 1
'        Me.bitsShow.Caption = Trim(str(bits))
'        If (Not connectFlag) Then
'            Dim strTemp As String, i As Integer
'            strTemp = "0."
'            For i = 0 To bits - 1
'                strTemp = strTemp + "0"
'            Next
'            Me.modifiedShow = strTemp
'            Me.realShow(0) = strTemp
'        End If
'    Else
'        MsgBox ("位数不能大于9")
'    End If
End Sub

Private Sub openDev()
On Error GoTo ErrorHandler
    Select Case dataDim.IOSetting
    Case 0
        status = viOpenDefaultRM(defrm)
        If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
        status = viOpen(defrm, "GPIB" + Trim(str(dataDim.GPIBnum)) + "::" + Trim(str(dataDim.GPIBaddr)) + "::INSTR", 0, 100, vi)
        If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Case 1
        Me.sp.PortOpen = True
        status = 1
    Case 2
        If (Me.Winsock.State <> 7) Then
            Me.Winsock.Close
            Me.Winsock.connect dataDim.IPaddr, dataDim.IPport
            Dim start As Long
            start = timeGetTime()
            While ((timeGetTime() - start) < 500)
                DoEvents
            Wend
            If (Me.Winsock.State = 7) Then
                status = 1
            Else
                status = -1
            End If
        End If
    End Select
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    status = -1
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub closeDev()
On Error GoTo ErrorHandler
    Select Case dataDim.IOSetting
    Case 0
        status = viClose(vi)
        If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
    Case 1
        Me.sp.PortOpen = False
    Case 2
        Me.Winsock.Close
    End Select
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub connect_Click()
On Error GoTo ErrorHandler
    If (Not connectFlag) Then
        Call openDev
        If (status >= VI_SUCCESS) Then
            Call wrIO("*IDN?", True)
            Me.LabInfo.Caption = "连接成功：The IDN String is: " & strRes
            If (dataDim.IOSetting <> 0 And Left(instruName, 3) = "福禄克") Then
                Call wrIO("syst:rem", False)
            End If
            connectFlag = True
            recflag = True
            Me.connect.BackColor = vbGreen
            Me.DCVIOChoice.Text = "DCV"
            Call function_Click(1)
            Me.sdTimer.Enabled = True
        End If
    Else
        If (Me.holdOn.BackColor = vbGreen) Then Call holdOn_Click
        Me.sdTimer.Enabled = False
        Call closeDev
        Me.connect.BackColor = &H8000000F
        If (zeroFlag) Then Call function_Click(0)
        zeroFlag = False
        connectFlag = False
        Me.rdTime.Enabled = False
        Me.mask.Visible = False
    End If
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub exit_Click()
    Unload Me
End Sub

Private Sub filterHalf_Click(Index As Integer)
    filterDiv = Index
    Call Set_Memory
End Sub

Private Sub filterSetting_Click()
    frmFilter.Show
End Sub

Private Sub Form_Load()
'    SkinH_Attach
'    SkinH_AttachEx App.Path & "\aero.she", ""
    Dim i As Integer, j As Integer, k As Integer

    'read file.ini
    If (Dir(App.Path & "\Config.ini") <> "") Then
        Call openFromFile(App.Path & "\Config.ini")
    Else
        instruName = "吉时利"
        dataDim.IOSetting = 0
        dataDim.GPIBaddr = 16
        dataDim.comport = 1
        dataDim.baudrate = 9600
        dataDim.databit = 8
        dataDim.stopbit = 2
        dataDim.cr = "NONE"
        dataDim.comPortOut = 2
        dataDim.baudRateOut = 9600
        dataDim.dataBitOut = 8
        dataDim.stopBitOut = 2
        dataDim.crOut = "NONE"
        dataDim.outEn = False
        dataDim.IPaddr = "169.254.4.10"
        dataDim.IPport = 5025
        dataDim.localIPPort = 2000
        For i = 0 To 2
            For j = 0 To 11
                aSteady(i, j, 0) = 0.003
                aSteady(i, j, 1) = 0.003
                tSteady(i, j) = 12 - j * 1
                filter(i, j) = 20
            Next
        Next
        For i = 0 To 6
            selfDef(i) = "1234567890"
            selfName(i) = "123"
            selfValid(i) = True
        Next
        For i = 0 To 4
            selfName(i) = "123"
        Next
        filterDiv = 0
        rangeNum = 0
        bits = 7
        For i = 0 To 2
            For j = 0 To 11
                dataDim.modiPara(i, j, 0) = 1
                dataDim.modiPara(i, j, 1) = 1
                dataDim.valid(i, j) = 1
            Next
        Next
        For i = 0 To 2
            For j = 0 To 11
                dataDim.multi(i, j) = 2
                div(i, j) = 2
            Next
        Next
        Call save2ini(App.Path & "\Config.ini")
    End If
    Call readCommand(instruName)
    
    'initialize windows
    Me.insName = instruName
    If (instruName = "吉时利2000" Or instruName = "安捷伦34401A") Then
        Me.LANSetting.Visible = False
        Me.IOChoice(2).Visible = False
        If (dataDim.IOSetting = 2) Then dataDim.IOSetting = 0
    Else
        Me.LANSetting.Visible = True
        Me.IOChoice(2).Visible = True
    End If
    Me.filterHalf(filterDiv).value = True
    Me.bitChoice.Text = Trim(str(bits))
    strTemp = "0."
    For i = 0 To bits - 1
        strTemp = strTemp + "0"
    Next
    For i = 19 To 0 Step -1
        Me.freq.AddItem (Trim(str(i + 1)))
    Next
    For i = 0 To 3
        Me.selfDefined(i).Caption = dataDim.selfName(i)
    Next
    
    Select Case dataDim.IOSetting
    Case 0
        Me.interfereShow.Caption = "GPIB"
    Case 1
        Me.interfereShow.Caption = "串口"
    Case 2
        Me.interfereShow.Caption = "LAN"
    End Select
    frmMain.floatContent.clear
    For i = 0 To 7
        Me.floatContent.AddItem dataDim.selfName2(i)
    Next
    Me.floatContent.ListIndex = 0
    
    'var initialize
    dataDim.GPIBnum = 0
    wrFlag = False
    connectFlag = False
    VIO = 0
    zeroFlag = False
    Me.sp.CommPort = dataDim.comport
    Me.sp.Settings = Trim(str(dataDim.baudrate)) + "," + Left(dataDim.cr, 1) + "," + Trim(str(dataDim.databit)) + "," + Trim(str(dataDim.stopbit))
    Me.sp.InputMode = comInputModeText
    Me.spOut.CommPort = dataDim.comPortOut
    Me.sp.RThreshold = 1
    Me.spOut.Settings = Trim(str(dataDim.baudRateOut)) + "," + Left(dataDim.crOut, 1) + "," + Trim(str(dataDim.dataBitOut)) + "," + Trim(str(dataDim.stopBitOut))
    If (dataDim.outEn) Then Me.spOut.PortOpen = True
    Me.sdTimer.Interval = 100
    sdInter = 90
    firstStd = False
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    Call save2ini(App.Path & "\Config.ini")
    If (Me.spOut.PortOpen) Then Me.spOut.PortOpen = False
End Sub


Private Sub wrIO(content As String, wrIOFlag As Boolean)
On Error GoTo ErrorHandler
    Select Case dataDim.IOSetting
    Case 0
        status = viVPrintf(vi, content + Chr$(10), 0)
        If (status < VI_SUCCESS) Then Exit Sub
        If (wrIOFlag) Then
            status = viVScanf(vi, "%t", strRes)
            If (status < VI_SUCCESS) Then GoTo VisaErrorHandler
        End If
    Case 1
        Me.sp.OutBufferCount = 0
        Me.sp.InBufferCount = 0
        Me.sp.Output = content + vbCrLf
        If (wrIOFlag) Then
            start = timeGetTime()
            While ((Me.sp.InBufferCount = 0) And (timeGetTime() - start < 2000))
                DoEvents
            Wend
            Call delay_ms(100)
            strRes = Me.sp.Input
            Me.sp.InBufferCount = 0
        End If
    Case 2
        recflag = False
        Winsock.SendData content & vbCrLf
        If (wrIOFlag) Then
            While ((Not recflag) And (timeGetTime() - start < 2000))
                DoEvents
            Wend
        End If
    End Select
    Exit Sub
ErrorHandler:
    Me.LabInfo.Caption = Error$
    Exit Sub
VisaErrorHandler:
    Dim strVisaErr As String * 200
    Call viStatusDesc(defrm, status, strVisaErr)
    Me.LabInfo.Caption = strVisaErr
    Exit Sub
End Sub

Private Sub freq_Click()
    Me.rdTime.Interval = CInt(1000 / Val(Me.freq.Text))
End Sub

Private Sub function_Click(Index As Integer)
    If (connectFlag) Then
        If (dataDim.IOSetting = 1) Then
            Me.rdTime.Enabled = False
            If (Len(Me.sp.Input) <> 0) Then strbuf = Me.sp.Input
        Else
            If (dataDim.IOSetting = 0) Then
                While wrFlag
    '               DoEvents
                Wend
                Me.rdTime.Enabled = False
            Else
                Me.rdTime.Enabled = False
                Me.Winsock.GetData strbuf
            End If
        End If
        Select Case Index
        Case 0
            If (Not zeroFlag) Then
'                Call wrIO(":READ?", True)
'                bias = Val(strRes)
                bias = biasStandard
                Me.modifiedShow.Text = ""
                If (dataDim.selfValid(5)) Then
                    Call delay_ms(sdInter)
                    Me.spOut.Output = dataDim.selfDef(5) + vbCrLf
                End If
            Else
                bias = 0
            End If
            zeroFlag = Not zeroFlag
            If (Not zeroFlag) Then
                Me.function(Index).BackColor = &H8000000F
            Else
                Me.function(Index).BackColor = vbGreen
            End If
            Call Set_Memory
        Case 1
            VIO = 0
            Call wrIO(":MEAS:VOLT:DC?", True)
            Me.range2.clear
            Me.range2.AddItem "100uV"
            Me.range2.AddItem "1mV"
            Me.range2.AddItem "10mV"
            Me.range2.AddItem "100mV"
            Me.range2.AddItem "1V"
            Me.range2.AddItem "10V"
            Me.range2.AddItem "100V"
            Me.range2.AddItem "1000V"
        Case 2
            VIO = 1
            'read measure value and average
            Call wrIO(":MEAS:CURR:DC?", True)
            Me.range2.clear
            Me.range2.AddItem "1uA"
            Me.range2.AddItem "10uA"
            Me.range2.AddItem "100uA"
            Me.range2.AddItem "1mA"
            Me.range2.AddItem "10mA"
            Me.range2.AddItem "100mA"
            Me.range2.AddItem "1A"
            Me.range2.AddItem "3A"
            Me.range2.AddItem "10A"
        Case 3
            VIO = 2
            'read measure value and average
            Call wrIO(":MEASure:FRES?", True)
            Me.range2.clear
            Me.range2.AddItem "1Ω"
            Me.range2.AddItem "10Ω"
            Me.range2.AddItem "100Ω"
            Me.range2.AddItem "1KΩ"
            Me.range2.AddItem "10KΩ"
            Me.range2.AddItem "100KΩ"
            Me.range2.AddItem "1MΩ"
            Me.range2.AddItem "10MΩ"
            Me.range2.AddItem "100MΩ"
            Me.range2.AddItem "1GΩ"
            Me.range2.AddItem "10GΩ"
            Me.range2.AddItem "100GΩ"
        End Select
        Me.range2.ListIndex = (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1)
        If (Index <> 0) Then
            Me.function(Index).BackColor = vbGreen
            Dim i As Integer
            For i = 1 To 3
                If (i <> Index) Then
                    Me.function(i).BackColor = &H8000000F
                End If
            Next
        End If
        rdTime.Enabled = True
     Else
        MsgBox ("请先连接设备")
    End If
    Exit Sub
End Sub

Private Sub GPIBSetting_Click()
    frmGPIB.Show
End Sub


Private Sub inBitsDown_Click()
'Dim strTemp  As String, i As Double
'    If (connectFlag) Then
'        If (dataDim.IOSetting = 1) Then
'            Me.rdTime.Enabled = False
'            If (Len(Me.sp.Input) <> 0) Then strbuf = Me.sp.Input
'        Else
'            While wrFlag
''               DoEvents
'            Wend
'            Me.rdTime.Enabled = False
'        End If
'        Call wrIO(dataDim.digits(VIO, 0), True)
'        i = Val(strRes)
'        If (instruName = "安捷伦") Then
'            i = i * 10
'        Else
'            i = i - 1
'        End If
'        'i = i - 1
'        'If (i <= 3) Then i = 4
'        Call wrIO(dataDim.digits(VIO, 1) + " " + Trim(str(i)), False)
'        Call wrIO(dataDim.digits(VIO, 0), True)
'        If (instruName = "安捷伦") Then
'            Me.InsDigits.Caption = Trim(str(Round(1 + Log(rangeN / Val(strRes)) / Log(10))))
'        Else
'            Me.InsDigits.Caption = Trim(str(Val(strRes)))
'        End If
'        Me.rdTime.Enabled = True
'    Else
'        MsgBox ("请先打开设备")
'    End If
End Sub

Private Sub inBitsUp_Click()
'Dim strTemp  As String, i As Double
'    If (connectFlag) Then
'        If (dataDim.IOSetting = 1) Then
'            Me.rdTime.Enabled = False
'            If (Len(Me.sp.Input) <> 0) Then strbuf = Me.sp.Input
'        Else
'            While wrFlag
''               DoEvents
'            Wend
'            Me.rdTime.Enabled = False
'        End If
'        Call wrIO(dataDim.digits(VIO, 0), True)
'        i = Val(strRes)
'        If (instruName = "安捷伦") Then
'            If (rangeN / i * 10 < 2000000) Then i = i / 10
'        Else
'            i = i + 1
'        End If
'        'If (i >= 8) Then i = 7
'        Call wrIO(dataDim.digits(VIO, 1) + " " + Trim(str(i)), False)
'        Call wrIO(dataDim.digits(VIO, 0), True)
'        If (instruName = "安捷伦") Then
'            Me.InsDigits.Caption = Trim(str(Round(1 + Log(rangeN / Val(strRes)) / Log(10))))
'        Else
'            Me.InsDigits.Caption = Trim(str(Val(strRes)))
'        End If
'    Me.rdTime.Enabled = True
'    Else
'        MsgBox ("请先打开设备")
'    End If
End Sub


'Private Sub IO_Click(Index As Integer)
'    If (connectFlag) Then
'        While wrFlag
'            DoEvents
'        Wend
'        If (dataDim.IOSetting = 1) Then
'            If (Len(Me.sp.Input) <> 0) Then strbuf = Me.sp.Input
'        End If
'        Me.rdTime.Enabled = False
'        Call closeDev
'        dataDim.IOSetting = Index
'        Call openDev
'        Me.rdTime.Enabled = True
'    Else
'        dataDim.IOSetting = Index
'    End If
'End Sub



Private Sub LANSetting_Click()
    frmLAN.Show
End Sub


Private Sub modifiedShow_GotFocus()
    HideCaret Me.modifiedShow.hwnd
End Sub

Private Sub modifySetting_Click()
    frmModify.Show
End Sub

Private Sub ok_Click()
    If (dataDim.selfValid(7)) Then
        Call delay_ms(sdInter)
        Me.spOut.Output = dataDim.selfDef(7) + vbCrLf
    End If
End Sub

Private Sub open_Click()
    Dim saveDir As String
    saveDir = App.Path + "\Instrument"
    If (Dir(saveDir, vbDirectory) = "") Then
        MkDir (saveDir)
    End If
    Me.CommonDialog1.DialogTitle = "打开文件"
    Me.CommonDialog1.InitDir = saveDir
    Me.CommonDialog1.filter = "配置文件(*.ini)|*.ini"
    Me.CommonDialog1.ShowOpen
    If (Me.CommonDialog1.filename <> vbNullString) Then
        Call openFromFile(Me.CommonDialog1.filename)
        Me.LabInfo.Caption = "打开成功"
    End If
End Sub


Private Sub range2_Click()
On Error Resume Next
If (connectFlag) Then
   If (dataDim.IOSetting = 1) Then
            Me.rdTime.Enabled = False
            If (Len(Me.sp.Input) <> 0) Then strbuf = Me.sp.Input
    Else
            If (dataDim.IOSetting = 0) Then
                While wrFlag
    '               DoEvents
                Wend
                Me.rdTime.Enabled = False
            Else
                Me.rdTime.Enabled = False
                Me.Winsock.GetData strbuf
            End If
    End If
    Select Case Left(Right(Me.range2.Text, 2), 1)
        Case "u"
            valTemp = 0.000001
        Case "m"
            valTemp = 0.001
        Case "K"
            valTemp = 1000
        Case "M"
            valTemp = 1000000
        Case "G"
            valTemp = 1000000000
        Case Else
            valTemp = 1
    End Select
    Dim Index As Integer
    Index = Me.range2.ListIndex
    If (connectFlag) Then
            Select Case VIO
            Case 0
                rangeN = 0.0001 * 10 ^ Index
                Call wrIO(":VOLT:RANG" + str(rangeN), False)
                start = timeGetTime()
                While (timeGetTime() - start < 10)
                    DoEvents
                Wend
                Call wrIO(":VOLT:RANG?", True)
                If (CInt(Val(strRes) / rangeN) <> 1) Then
                    Me.range2.ListIndex = (Index + CInt(Log(Val(strRes) / rangeN) / Log(10)))
                    Me.LabInfo.Caption = "量程之外"
                Else
                    rangeNum = rangeNum + (Index - Round((rangeNum Mod 10 ^ (2 * VIO + 3)) / 10 ^ (2 * VIO + 1), 0)) * 10 ^ (2 * VIO + 1)
                End If
            Case 1
                rangeN = 0.000001 * 10 ^ Index
                If (rangeN = 10) Then rangeN = 3
                If (rangeN > 10) Then rangeN = 10
                Call wrIO(":CURR:RANG" + str(rangeN), False)
                start = timeGetTime()
                While (timeGetTime() - start < 10)
                    DoEvents
                Wend
                Call wrIO(":CURR:RANG?", True)
                If (CInt(Val(strRes) / rangeN) <> 1) Then
                    If (Val(strRes) < 3) Then
                        If (rangeN <> 10) Then
                            Me.range2.ListIndex = (Index + CInt(Log(Val(strRes) / rangeN) / Log(10)))
                        Else
                            Me.range2.ListIndex = (Index - 1 + CInt(Log(Val(strRes) / rangeN) / Log(10)))
                        End If
                    Else
                        If (CInt(Val(strRes)) = 3) Then Me.range2.ListIndex = 7
                        If (Val(strRes) = 10) Then Me.range2.ListIndex = 8
                    End If
                    Me.LabInfo.Caption = "量程之外"
                Else
                    rangeNum = rangeNum + (Index - Round((rangeNum Mod 10 ^ (2 * VIO + 3)) / 10 ^ (2 * VIO + 1), 0)) * 10 ^ (2 * VIO + 1)
                End If
            Case 2
                rangeN = 1 * 10 ^ Index
                Call wrIO(":FRES:RANG" + str(rangeN), False)
                start = timeGetTime()
                While (timeGetTime() - start < 10)
                    DoEvents
                Wend
                Call wrIO(":FRES:RANG?", True)
                If (CInt(Val(strRes) / rangeN) <> 1) Then
                    Me.range2.ListIndex = (Index + CInt(Log(Val(strRes) / rangeN) / Log(10)))
                    Me.LabInfo.Caption = "量程之外"
                Else
                    rangeNum = rangeNum + (Index - Round((rangeNum Mod 10 ^ (2 * VIO + 3)) / 10 ^ (2 * VIO + 1), 0)) * 10 ^ (2 * VIO + 1)
                End If
            End Select
            'rangeNum = rangeNum + (Index - Round((rangeNum Mod 10 ^ (2 * VIO + 3)) / 10 ^ (2 * VIO + 1), 0)) * 10 ^ (2 * VIO + 1)
        Else
            rangeNum = rangeNum + (Index - (rangeNum Mod 1000) / 10) * 10
        End If
        Call Set_Memory
        memCnt = 0
'        Call Set_Revise
        setMinMax = False
        steadyCnt = 0
        Me.stTime.Interval = tSteady(VIO, (rangeNum Mod 10 ^ (2 * VIO + 3)) / 10 ^ (2 * VIO + 1)) * 1000
        Me.stTime.Enabled = True
        Me.steadyShow.BackColor = vbRed
        stdAverage = 0
        recflag = True
        If (zeroFlag) Then Call function_Click(0)
        Me.rdTime.Enabled = True
    End If
End Sub

Private Sub save_Click()
    Dim saveDir As String
    saveDir = App.Path + "\Instrument"
    If (Dir(saveDir, vbDirectory) = "") Then
        MkDir (saveDir)
    End If
    Me.CommonDialog1.DialogTitle = "保存文件"
    Me.CommonDialog1.InitDir = saveDir
    Me.CommonDialog1.filter = "配置文件(*.ini)|*.ini"
    Me.CommonDialog1.ShowSave
    If (Me.CommonDialog1.filename <> vbNullString) Then
        Call save2ini(Me.CommonDialog1.filename)
        Me.LabInfo.Caption = "存储成功"
    End If
End Sub

Private Sub sdTimer_Timer()
    If (dataDim.outEn) Then
        outFlag = True
        If (Me.outSwitch.value = 1) Then
            Me.spOut.Output = Format(Val(sendStringM), IIf((Val(sendStringM) > 0), "+0.00000000E+00", "0.00000000E+00")) + vbCrLf
        Else
            Me.spOut.Output = Format(Val(sendStringR), IIf((Val(sendStringR) > 0), "+0.00000000E+00", "0.00000000E+00")) + vbCrLf
        End If
        outFlag = False
        
'        '232 output
'        If ((Me.holdOn.BackColor = vbGreen) And dataDim.selfValid(6)) Then
'            Call delay_ms(sdInter)
'            outFlag = True
'            Me.spOut.Output = dataDim.selfDef(6) + vbCrLf
'            outFlag = False
'        End If
'        If (zeroFlag And dataDim.selfValid(5)) Then
'            Call delay_ms(sdInter)
'            outFlag = True
'            Me.spOut.Output = dataDim.selfDef(5) + vbCrLf
'            outFlag = False
'        End If
'        If ((Me.steadyShow.FillColor = vbGreen) And dataDim.selfValid(4)) Then
'            Call delay_ms(sdInter)
'            outFlag = True
'            Me.spOut.Output = dataDim.selfDef(4) + vbCrLf
'            outFlag = False
'        End If
    End If
End Sub

Private Sub self_Click()
    frmSelf.Show
End Sub

Private Sub selfDefined_Click(Index As Integer)
    Select Case Index
    Case 0, 1, 2, 3
        Call delay_ms(sdInter)
        Me.spOut.Output = dataDim.selfDef(Index) + vbCrLf
    Case 4
        If (Me.floatContent.Visible = True) Then
            Me.floatContent.Visible = False
            If (Me.floatContent.Text <> "") Then
                Call delay_ms(sdInter)
                Me.spOut.Output = dataDim.selfDef2(Me.floatContent.ListIndex) + vbCrLf
            End If
        Else
            Me.floatContent.Visible = True
        End If
    End Select
End Sub

Private Sub speed_Click()
    If (Me.speed.value = 0) Then
        sdInter = 40
        Me.sdTimer.Interval = 50
    Else
        sdInter = 90
        Me.sdTimer.Interval = 100
    End If
End Sub

Private Sub steadySetting_Click()
    frmSteady.Show
End Sub

Private Function average(value As Double) As Double
    memory(memCnt) = value
    Dim sum As Double, num As Integer, i As Integer
    sum = 0
    num = IIf(fullFlag, filterCur, memCnt + 1)
    For i = 0 To num - 1
        sum = sum + memory(i)
    Next
    average = sum / num
    memCnt = memCnt + 1
    If (memCnt > filterCur - 1) Then
        fullFlag = True
        memCnt = 0
    End If
End Function

Private Sub isSteady(value As Double)
    If (value < 9.9E+35) Then
        If (Not setMinMax) Then
            min = value
            max = value
            setMinMax = True
        End If
        If (value > max) Then max = value
        If (value < min) Then min = value
        steadyCnt = steadyCnt + 1
        stdAverage = stdAverage + (value - stdAverage) / steadyCnt
        'if any value exceed the bound, turn red immediately
        Dim up, down As Double
        down = stdAverage * (1 - dataDim.aSteady(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), 0))
        up = stdAverage * (1 + dataDim.aSteady(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), 1))
        If (stdAverage > 0) Then
            If (min > down And max < up) Then
            
            Else
                Me.steadyShow.BackColor = vbRed
                firstStd = False
            End If
        Else
            If (min > up And max < down) Then
               
            Else
                Me.steadyShow.BackColor = vbRed
                firstStd = False
            End If
        End If
    End If
End Sub

Private Sub rdTime_Timer()
    'cannot interrupt
    wrFlag = True
    Call wrIO(":READ?", True)
    wrFlag = False
    'cannot interrupt
    Dim realValue As Double, modifyValue As Double
    realValue = Val(strRes)
    Me.realShow(1).Text = "\  " + Left(strRes, 15)
    If (dataDim.valid(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1)) = 1) Then
        If (modifyValue > 0) Then
            modifyValue = modifyValue * dataDim.modiPara(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), 0)
        Else
            modifyValue = modifyValue * dataDim.modiPara(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), 1)
        End If
    End If
    modifyValue = average(realValue)
    modifyValue = modifyValue * dataDim.multi(VIO, 0) / div(VIO, 0)
    Call isSteady(modifyValue)
    biasStandard = modifyValue
    modifyValue = modifyValue - bias
    Call display(modifyValue, realValue)
    '232 output
    sendStringR = str(realValue)
    sendStringM = str(modifyValue)
'    outFlag = True
'    If (dataDim.outEn) Then
'        If (Me.outSwitch.value = 0) Then
'            Me.spOut.Output = Format(realValue, IIf((realValue > 0), "+0.00000000E+00", "0.00000000E+00")) + vbCrLf
'        Else
'            Me.spOut.Output = Format(modifyValue, IIf((modifyValue > 0), "+0.00000000E+00", "0.00000000E+00")) + vbCrLf
'        End If
'    End If
'    outFlag = False
'    If ((Me.holdOn.BackColor = vbGreen) And dataDim.selfValid(6)) Then
'        Call delay_ms(10)
'        outFlag = True
'        Me.spOut.Output = dataDim.selfDef(6) + vbCrLf
'        outFlag = False
'    End If
'    If (zeroFlag And dataDim.selfValid(5)) Then
'        Call delay_ms(10)
'        outFlag = True
'        Me.spOut.Output = dataDim.selfDef(5) + vbCrLf
'        outFlag = False
'    End If
'    If ((Me.steadyShow.FillColor = vbGreen) And dataDim.selfValid(4)) Then
'        Call delay_ms(10)
'        outFlag = True
'        Me.spOut.Output = dataDim.selfDef(4) + vbCrLf
'        outFlag = False
'    End If
    'dot blink
    If (modifyValue < 9.9E+35) Then
        Me.mask.Visible = True
        Call delay_ms(100)
        Me.mask.Visible = False
    Else
        mask.Visible = False
    End If
End Sub

Private Sub display(value As Double, real As Double)
    Dim num As Integer, i As Integer, numFormat, numFormata As String
    If (value < 9.9E+35) Then
        num = Val(Me.range2.Text)
        Select Case num
            Case 1
                numFormat = "0."
                i = 1
                Me.mask.Left = 3335 + (7 - bits) * 695
            Case 3, 10
                numFormat = "00."
                i = 2
                Me.mask.Left = 4030 + (7 - bits) * 695
            Case 100
                numFormat = "000."
                Me.mask.Left = 4725 + (7 - bits) * 695
                i = 3
            Case 1000
                numFormat = "0000."
                i = 4
                Me.mask.Left = 5420 + (7 - bits) * 695
        End Select
        Dim a, b, temp As Integer
        b = bits - i
        a = 9 - i
        temp = i
        numFormata = numFormat
        For i = temp To bits - 1
            numFormat = numFormat + "0"
        Next
        For i = temp To 8
            numFormata = numFormata + "0"
        Next
        Me.modifiedShow.Text = Format(Fix(value / valTemp * 10 ^ b) / 10 ^ b, numFormat)
        Me.realShow(0).Text = Format(Fix(real / valTemp * 10 ^ a) / 10 ^ a, numFormata)
    Else
        Me.modifiedShow.Text = "OVER FLOW"
        Me.realShow(0).Text = "OVER FLOW"
    End If
End Sub

Private Sub stTime_Timer()
    Dim up, down As Double
    down = stdAverage * (1 - dataDim.aSteady(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), 0))
    up = stdAverage * (1 + dataDim.aSteady(VIO, (rangeNum Mod (10 ^ (2 * VIO + 3))) / 10 ^ (2 * VIO + 1), 1))
    If (stdAverage > 0) Then
        If (min > down And max < up) Then
            Me.steadyShow.BackColor = vbGreen
            If (Not firstStd) Then
                If (dataDim.selfValid(4)) Then
                    Call delay_ms(sdInter)
                    Me.spOut.Output = dataDim.selfDef(4) + vbCrLf
                End If
                firstStd = True
            End If
        Else
            Me.steadyShow.BackColor = vbRed
            firstStd = False
        End If
    Else
        If (min > up And max < down) Then
            Me.steadyShow.BackColor = vbGreen
            If (Not firstStd) Then
                If (dataDim.selfValid(4)) Then
                    Call delay_ms(sdInter)
                    Me.spOut.Output = dataDim.selfDef(4) + vbCrLf
                End If
                firstStd = True
            End If
        Else
            Me.steadyShow.BackColor = vbRed
            firstStd = False
        End If
    End If
    setMinMax = False
    steadyCnt = 0
    stdAverage = 0
End Sub

Private Sub test_Click()
    frmTest.Show
End Sub

Private Sub usartSetting_Click()
    frmUsart.Show
End Sub

'Private Function revise(value As Double) As Double
'    Dim i As Integer
'    For i = 0 To reviseNum - 2
'        If (value < reviseArr(1, i) And value > reviseArr(1, i + 1)) Then
'            revise = reviseArr(0, i + 1) + (reviseArr(0, i) - reviseArr(0, i + 1)) / (reviseArr(1, i) - reviseArr(1, i + 1)) * (value - reviseArr(1, i + 1))
'            Exit For
'        End If
'    Next
'    If (value > reviseArr(1, 0) Or value < reviseArr(1, reviseNum - 1)) Then revise = value
'End Function

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    If (Not recflag) Then
        Winsock.GetData strRes
    Else
        Winsock.GetData strbuf
    End If
    recflag = True
End Sub
