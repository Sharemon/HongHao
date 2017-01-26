VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmDataShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据分析"
   ClientHeight    =   10788
   ClientLeft      =   6192
   ClientTop       =   1188
   ClientWidth     =   6828
   Icon            =   "FrmDataShow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10788
   ScaleWidth      =   6828
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   135
      Top             =   0
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "清零"
         Height          =   375
         Left            =   5640
         TabIndex        =   136
         Top             =   130
         Width           =   855
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   2
         Left            =   4800
         TabIndex        =   142
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "重复帧数："
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   141
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "接收帧数："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   140
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "有效帧数："
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   139
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   138
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Height          =   255
         Index           =   1
         Left            =   3000
         TabIndex        =   137
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "图形显示"
      Height          =   255
      Index           =   1
      Left            =   4320
      TabIndex        =   130
      Top             =   720
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "二进制"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   129
      Top             =   720
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "关闭"
      Height          =   495
      Index           =   1
      Left            =   4680
      TabIndex        =   4
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "暂停"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog ComnDiaFile 
      Left            =   3240
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "保存文件为"
      Filter          =   "文本文件|*.txt"
      Flags           =   33792
   End
   Begin VB.Frame FrmDataSaveStatus 
      Caption         =   "数据记录"
      Height          =   2175
      Left            =   120
      TabIndex        =   5
      Top             =   8520
      Width           =   6615
      Begin VB.CommandButton CmdSave 
         Caption         =   "选择路径..."
         Height          =   495
         Index           =   2
         Left            =   5280
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TxtSaveAddress 
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   5055
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "停止记录"
         Height          =   495
         Index           =   1
         Left            =   5280
         TabIndex        =   14
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "开始记录"
         Height          =   495
         Index           =   0
         Left            =   5280
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "保存规则选择"
         Height          =   1335
         Index           =   1
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Width           =   2415
         Begin VB.OptionButton OptionSavegz 
            Caption         =   "逐条记录数据到文本中"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton OptionSavegz 
            Caption         =   "连续记录数据到文本中"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   840
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "记录形式选择"
         Height          =   1335
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton OptionSaveStyle 
            Caption         =   "十六进制形式保存"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton OptionSaveStyle 
            Caption         =   "ASCII制形式保存"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   2295
         End
         Begin VB.OptionButton OptionSaveStyle 
            Caption         =   "二进制形式保存"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   2295
         End
      End
   End
   Begin VB.Label Label3 
      Caption         =   ")"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   134
      Top             =   750
      Width           =   135
   End
   Begin VB.Shape Deng 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   180
      Index           =   1
      Left            =   6120
      Shape           =   3  'Circle
      Top             =   750
      Width           =   180
   End
   Begin VB.Label Label3 
      Caption         =   ",0"
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   133
      Top             =   750
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "(1"
      Height          =   255
      Index           =   0
      Left            =   5400
      TabIndex        =   132
      Top             =   750
      Width           =   255
   End
   Begin VB.Shape Deng 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   180
      Index           =   0
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   750
      Width           =   180
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "  7   6   5   4   3   2   1   0"
      Height          =   180
      Left            =   2805
      TabIndex        =   131
      Top             =   1005
      Width           =   3975
   End
   Begin VB.Shape DengStyle1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   3485
      Shape           =   3  'Circle
      Top             =   1230
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   1
      Left            =   2760
      TabIndex        =   127
      Top             =   1380
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   2
      Left            =   2760
      TabIndex        =   126
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   3
      Left            =   2760
      TabIndex        =   125
      Top             =   1740
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   4
      Left            =   2760
      TabIndex        =   124
      Top             =   1920
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   5
      Left            =   2760
      TabIndex        =   123
      Top             =   2100
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   6
      Left            =   2760
      TabIndex        =   122
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   7
      Left            =   2760
      TabIndex        =   121
      Top             =   2460
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   8
      Left            =   2760
      TabIndex        =   120
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   9
      Left            =   2760
      TabIndex        =   119
      Top             =   2820
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   10
      Left            =   2760
      TabIndex        =   118
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   11
      Left            =   2760
      TabIndex        =   117
      Top             =   3180
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   12
      Left            =   2760
      TabIndex        =   116
      Top             =   3360
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   13
      Left            =   2760
      TabIndex        =   115
      Top             =   3540
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   14
      Left            =   2760
      TabIndex        =   114
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   15
      Left            =   2760
      TabIndex        =   113
      Top             =   3900
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   16
      Left            =   2760
      TabIndex        =   112
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   17
      Left            =   2760
      TabIndex        =   111
      Top             =   4260
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   18
      Left            =   2760
      TabIndex        =   110
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   19
      Left            =   2760
      TabIndex        =   109
      Top             =   4620
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   20
      Left            =   2760
      TabIndex        =   108
      Top             =   4800
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   21
      Left            =   2760
      TabIndex        =   107
      Top             =   4980
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   22
      Left            =   2760
      TabIndex        =   106
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   23
      Left            =   2760
      TabIndex        =   105
      Top             =   5340
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   24
      Left            =   2760
      TabIndex        =   104
      Top             =   5520
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   25
      Left            =   2760
      TabIndex        =   103
      Top             =   5700
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   26
      Left            =   2760
      TabIndex        =   102
      Top             =   5880
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   27
      Left            =   2760
      TabIndex        =   101
      Top             =   6060
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   28
      Left            =   2760
      TabIndex        =   100
      Top             =   6240
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   29
      Left            =   2760
      TabIndex        =   99
      Top             =   6420
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   30
      Left            =   2760
      TabIndex        =   98
      Top             =   6600
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   31
      Left            =   2760
      TabIndex        =   97
      Top             =   6780
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   32
      Left            =   2760
      TabIndex        =   96
      Top             =   6960
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   33
      Left            =   2760
      TabIndex        =   95
      Top             =   7140
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   34
      Left            =   2760
      TabIndex        =   94
      Top             =   7320
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   35
      Left            =   2760
      TabIndex        =   93
      Top             =   7500
      Width           =   3975
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   36
      Left            =   2760
      TabIndex        =   92
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   0
      Left            =   600
      TabIndex        =   91
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   1
      Left            =   600
      TabIndex        =   90
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   2
      Left            =   600
      TabIndex        =   89
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   3
      Left            =   600
      TabIndex        =   88
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   4
      Left            =   600
      TabIndex        =   87
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   5
      Left            =   600
      TabIndex        =   86
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   6
      Left            =   600
      TabIndex        =   85
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   7
      Left            =   600
      TabIndex        =   84
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   8
      Left            =   600
      TabIndex        =   83
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   9
      Left            =   600
      TabIndex        =   82
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   10
      Left            =   600
      TabIndex        =   81
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   11
      Left            =   600
      TabIndex        =   80
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   12
      Left            =   600
      TabIndex        =   79
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   13
      Left            =   600
      TabIndex        =   78
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   14
      Left            =   600
      TabIndex        =   77
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   15
      Left            =   600
      TabIndex        =   76
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   16
      Left            =   600
      TabIndex        =   75
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   17
      Left            =   600
      TabIndex        =   74
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   18
      Left            =   600
      TabIndex        =   73
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   19
      Left            =   600
      TabIndex        =   72
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   20
      Left            =   600
      TabIndex        =   71
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   21
      Left            =   600
      TabIndex        =   70
      Top             =   4980
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   22
      Left            =   600
      TabIndex        =   69
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   23
      Left            =   600
      TabIndex        =   68
      Top             =   5340
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   24
      Left            =   600
      TabIndex        =   67
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   25
      Left            =   600
      TabIndex        =   66
      Top             =   5700
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   26
      Left            =   600
      TabIndex        =   65
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   27
      Left            =   600
      TabIndex        =   64
      Top             =   6060
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   28
      Left            =   600
      TabIndex        =   63
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   29
      Left            =   600
      TabIndex        =   62
      Top             =   6420
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   30
      Left            =   600
      TabIndex        =   61
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   31
      Left            =   600
      TabIndex        =   60
      Top             =   6780
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   32
      Left            =   600
      TabIndex        =   59
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   33
      Left            =   600
      TabIndex        =   58
      Top             =   7140
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   34
      Left            =   600
      TabIndex        =   57
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   35
      Left            =   600
      TabIndex        =   56
      Top             =   7500
      Width           =   1095
   End
   Begin VB.Label LabHex 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   36
      Left            =   600
      TabIndex        =   55
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   0
      Left            =   1680
      TabIndex        =   54
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   1
      Left            =   1680
      TabIndex        =   53
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   2
      Left            =   1680
      TabIndex        =   52
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   3
      Left            =   1680
      TabIndex        =   51
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   4
      Left            =   1680
      TabIndex        =   50
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   5
      Left            =   1680
      TabIndex        =   49
      Top             =   2100
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   6
      Left            =   1680
      TabIndex        =   48
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   7
      Left            =   1680
      TabIndex        =   47
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   8
      Left            =   1680
      TabIndex        =   46
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   9
      Left            =   1680
      TabIndex        =   45
      Top             =   2820
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   10
      Left            =   1680
      TabIndex        =   44
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   11
      Left            =   1680
      TabIndex        =   43
      Top             =   3180
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   12
      Left            =   1680
      TabIndex        =   42
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   13
      Left            =   1680
      TabIndex        =   41
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   14
      Left            =   1680
      TabIndex        =   40
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   15
      Left            =   1680
      TabIndex        =   39
      Top             =   3900
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   16
      Left            =   1680
      TabIndex        =   38
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   17
      Left            =   1680
      TabIndex        =   37
      Top             =   4260
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   18
      Left            =   1680
      TabIndex        =   36
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   19
      Left            =   1680
      TabIndex        =   35
      Top             =   4620
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   20
      Left            =   1680
      TabIndex        =   34
      Top             =   4800
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   21
      Left            =   1680
      TabIndex        =   33
      Top             =   4980
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   22
      Left            =   1680
      TabIndex        =   32
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   23
      Left            =   1680
      TabIndex        =   31
      Top             =   5340
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   24
      Left            =   1680
      TabIndex        =   30
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   25
      Left            =   1680
      TabIndex        =   29
      Top             =   5700
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   26
      Left            =   1680
      TabIndex        =   28
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   27
      Left            =   1680
      TabIndex        =   27
      Top             =   6060
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   28
      Left            =   1680
      TabIndex        =   26
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   29
      Left            =   1680
      TabIndex        =   25
      Top             =   6420
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   30
      Left            =   1680
      TabIndex        =   24
      Top             =   6600
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   31
      Left            =   1680
      TabIndex        =   23
      Top             =   6780
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   32
      Left            =   1680
      TabIndex        =   22
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   33
      Left            =   1680
      TabIndex        =   21
      Top             =   7140
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   34
      Left            =   1680
      TabIndex        =   20
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00B0B0B0&
      Height          =   180
      Index           =   35
      Left            =   1680
      TabIndex        =   19
      Top             =   7500
      Width           =   1095
   End
   Begin VB.Label LabASC 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   36
      Left            =   1680
      TabIndex        =   18
      Top             =   7680
      Width           =   1095
   End
   Begin VB.Label LabNO 
      Alignment       =   1  'Right Justify
      Height          =   6735
      Left            =   240
      TabIndex        =   17
      Top             =   1200
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ASCII"
      Height          =   255
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "十六进制"
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "字节号"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   615
   End
   Begin VB.Shape DengStyle0 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00000040&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   3485
      Shape           =   3  'Circle
      Top             =   1230
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label LabBin 
      Alignment       =   2  'Center
      BackColor       =   &H00D0D0D0&
      Height          =   180
      Index           =   0
      Left            =   2760
      TabIndex        =   128
      Top             =   1200
      Width           =   3975
   End
End
Attribute VB_Name = "FrmDataShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ISSelected As Boolean

Private Sub CmdOK_Click(Index As Integer)
Select Case Index
Case 0
    If CmdOK(0).Caption = "继续" Then
        CmdOK(0).Caption = "暂停"
        ISDataShow = True
    ElseIf CmdOK(0).Caption = "Continue" Then
        CmdOK(0).Caption = "Pause"
        ISDataShow = True
    ElseIf CmdOK(0).Caption = "暂停" Then
        CmdOK(0).Caption = "继续"
        ISDataShow = False
    ElseIf CmdOK(0).Caption = "Pause" Then
        CmdOK(0).Caption = "Continue"
        ISDataShow = False
    End If
Case 1
    ISDataShow = False
    Unload Me
End Select
End Sub

Private Sub CmdSave_Click(Index As Integer)
Select Case Index
Case 0
If OpenFileName = "" Then
    MsgBox "请先选择保存路径"
Else
    If OptionSavegz(0).Value = True Then
        DatasSaveStyle = 2
    ElseIf OptionSavegz(1).Value = True Then
        DatasSaveStyle = 3
    End If
End If
Case 1
    DatasSaveStyle = 1
Case 2
On Error GoTo ErrHandler
ComnDiaFile.InitDir = App.Path
ComnDiaFile.FileName = Format(Date, "yyyy-mm-dd") & "_" & Format(Now, "hh-mm-ss") & "数据记录.txt"
ComnDiaFile.Filter = "文本文件(*.txt)|*.txt"
ComnDiaFile.FilterIndex = 1
ComnDiaFile.ShowSave
OpenFileName = ComnDiaFile.FileName
ComnDiaFile.FileName = ""
TxtSaveAddress = OpenFileName
RecordNumber = 0
ErrHandler:
DatasSaveStyle = 1
Exit Sub
End Select
End Sub

Private Sub Command1_Click()
ReceiveCounts = 0
ReceiveTrueCounts = 0
End Sub

Private Sub Form_Click()
For i = 0 To 36
    If i Mod 2 = 0 Then
        Me.LabASC(i).BackColor = &HD0D0D0
        Me.LabBin(i).BackColor = &HD0D0D0
        Me.LabHex(i).BackColor = &HD0D0D0
    Else
        Me.LabASC(i).BackColor = &HB0B0B0
        Me.LabBin(i).BackColor = &HB0B0B0
        Me.LabHex(i).BackColor = &HB0B0B0
    End If
Next
End Sub

Private Sub Form_Load()
ISSelected = False
For i = 1 To 37
    Me.LabNO.Caption = Me.LabNO.Caption & i & vbCrLf
    Me.LabASC(i - 1).Font.Name = "宋体"
    Me.LabASC(i - 1).Font.Size = 9
    Me.LabBin(i - 1).Font.Name = "宋体"
    Me.LabBin(i - 1).Font.Size = 9
    Me.LabHex(i - 1).Font.Name = "宋体"
    Me.LabHex(i - 1).Font.Size = 9
Next
Me.LabNO.Font.Name = "宋体"
Me.LabNO.Font.Size = 9
Me.Label2.Font.Name = "宋体"
Me.Label2.Font.Size = 9
If Lan = 0 Then
    FrmDataShow.Caption = "接收数据分析"
    Label1(0).Caption = "字节号"
    Label1(1).Caption = "十六进制"
    Label1(2).Caption = "ASCII"
    Option1(0).Caption = "二进制"
    Option1(1).Caption = "图形显示"
    CmdOK(0).Caption = "暂停"
    CmdOK(1).Caption = "关闭"
    FrmDataSaveStatus.Caption = "数据记录 - 记录数据未开启"
    Frame2(0).Caption = "记录形式选择"
    OptionSaveStyle(0).Caption = "十六进制形式保存"
    OptionSaveStyle(1).Caption = "ASCII形式保存"
    OptionSaveStyle(2).Caption = "二进制形式保存"
    Frame2(1).Caption = "保存规则选择"
    OptionSavegz(0).Caption = "逐条记录数据到文本中"
    OptionSavegz(1).Caption = "连续记录数据到文本中"
    CmdSave(0).Caption = "开始记录"
    CmdSave(1).Caption = "停止记录"
    CmdSave(2).Caption = "选择路径..."
    Label4(0).Caption = "接收帧数："
    Label4(1).Caption = "有效帧数："
    Label4(2).Caption = "重复帧数："
    Command1.Caption = "清零"
ElseIf Lan = 1 Then
    FrmDataShow.Caption = "Receive data analysis"
    Label1(0).Caption = "NO."
    Label1(1).Caption = "Hex"
    Label1(2).Caption = "ASCII"
    Option1(0).Caption = "Binary"
    Option1(1).Caption = "Graphical"
    CmdOK(0).Caption = "Pause"
    CmdOK(1).Caption = "Exit"
    FrmDataSaveStatus.Caption = "Data save - Closed"
    Frame2(0).Caption = "Save style"
    OptionSaveStyle(0).Caption = "Save by Hexadecimal"
    OptionSaveStyle(1).Caption = "Save by ASCII"
    OptionSaveStyle(2).Caption = "Save by Binary"
    Frame2(1).Caption = "Save rule"
    OptionSavegz(0).Caption = "Save One by one"
    OptionSavegz(1).Caption = "Consecutive save"
    CmdSave(0).Caption = "Start"
    CmdSave(1).Caption = "Stop"
    CmdSave(2).Caption = "Address..."
    Label4(0).Caption = "Receive:"
    Label4(1).Caption = "Effective:"
    Label4(2).Caption = "Repeat:"
    Command1.Caption = "Clear"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
OpenFileName = ""
ISDataShow = False
End Sub

Private Sub LabASC_Click(Index As Integer)
If LabASC(Index).BackColor = &HD0D0D0 Or LabASC(Index).BackColor = &HB0B0B0 Then
    For i = 0 To 36
        If i Mod 2 = 0 Then
            Me.LabASC(i).BackColor = &HD0D0D0
            Me.LabBin(i).BackColor = &HD0D0D0
            Me.LabHex(i).BackColor = &HD0D0D0
        Else
            Me.LabASC(i).BackColor = &HB0B0B0
            Me.LabBin(i).BackColor = &HB0B0B0
            Me.LabHex(i).BackColor = &HB0B0B0
        End If
    Next
    Me.LabASC(Index).BackColor = &HFF8080
    Me.LabBin(Index).BackColor = &HFF8080
    Me.LabHex(Index).BackColor = &HFF8080
Else
        If Index Mod 2 = 0 Then
            Me.LabASC(Index).BackColor = &HD0D0D0
            Me.LabBin(Index).BackColor = &HD0D0D0
            Me.LabHex(Index).BackColor = &HD0D0D0
        Else
            Me.LabASC(Index).BackColor = &HB0B0B0
            Me.LabBin(Index).BackColor = &HB0B0B0
            Me.LabHex(Index).BackColor = &HB0B0B0
        End If
End If
End Sub

Private Sub LabBin_Click(Index As Integer)
If LabBin(Index).BackColor = &HD0D0D0 Or LabBin(Index).BackColor = &HB0B0B0 Then
    For i = 0 To 36
        If i Mod 2 = 0 Then
            Me.LabASC(i).BackColor = &HD0D0D0
            Me.LabBin(i).BackColor = &HD0D0D0
            Me.LabHex(i).BackColor = &HD0D0D0
        Else
            Me.LabASC(i).BackColor = &HB0B0B0
            Me.LabBin(i).BackColor = &HB0B0B0
            Me.LabHex(i).BackColor = &HB0B0B0
        End If
    Next
    Me.LabASC(Index).BackColor = &HFF8080
    Me.LabBin(Index).BackColor = &HFF8080
    Me.LabHex(Index).BackColor = &HFF8080
Else
        If Index Mod 2 = 0 Then
            Me.LabASC(Index).BackColor = &HD0D0D0
            Me.LabBin(Index).BackColor = &HD0D0D0
            Me.LabHex(Index).BackColor = &HD0D0D0
        Else
            Me.LabASC(Index).BackColor = &HB0B0B0
            Me.LabBin(Index).BackColor = &HB0B0B0
            Me.LabHex(Index).BackColor = &HB0B0B0
        End If
End If
End Sub

Private Sub LabHex_Click(Index As Integer)
If LabHex(Index).BackColor = &HD0D0D0 Or LabHex(Index).BackColor = &HB0B0B0 Then
    For i = 0 To 36
        If i Mod 2 = 0 Then
            Me.LabASC(i).BackColor = &HD0D0D0
            Me.LabBin(i).BackColor = &HD0D0D0
            Me.LabHex(i).BackColor = &HD0D0D0
        Else
            Me.LabASC(i).BackColor = &HB0B0B0
            Me.LabBin(i).BackColor = &HB0B0B0
            Me.LabHex(i).BackColor = &HB0B0B0
        End If
    Next
    Me.LabASC(Index).BackColor = &HFF8080
    Me.LabBin(Index).BackColor = &HFF8080
    Me.LabHex(Index).BackColor = &HFF8080
Else
        If Index Mod 2 = 0 Then
            Me.LabASC(Index).BackColor = &HD0D0D0
            Me.LabBin(Index).BackColor = &HD0D0D0
            Me.LabHex(Index).BackColor = &HD0D0D0
        Else
            Me.LabASC(Index).BackColor = &HB0B0B0
            Me.LabBin(Index).BackColor = &HB0B0B0
            Me.LabHex(Index).BackColor = &HB0B0B0
        End If
End If
End Sub

Private Sub Option1_Click(Index As Integer)
If Option1(0).Value = True Then
    BinOrDeng = False
    If Me.DengStyle0().Count > 1 And Me.DengStyle1().Count > 1 Then
        DengStyle0(i).Visible = False
        DengStyle1(i).Visible = False
        For i = 1 To 295
            Unload Me.DengStyle0(i)
            Unload Me.DengStyle1(i)
        Next i
    End If
Else
BinOrDeng = True
    If Me.DengStyle0().Count <> 296 And Me.DengStyle1().Count <> 296 Then
        DengStyle0(i).Visible = True
        DengStyle1(i).Visible = True
        For i = 1 To 295
            Load DengStyle0(i)
            DengStyle0(i).Top = LabBin(Int(i / 8)).Top + 25
            DengStyle0(i).Left = 3485 + 385 * (i Mod 8)
            DengStyle0(i).Visible = True
            DengStyle0(i).ZOrder
            Load DengStyle1(i)
            DengStyle1(i).Top = LabBin(Int(i / 8)).Top + 25
            DengStyle1(i).Left = 3485 + 385 * (i Mod 8)
            DengStyle1(i).Visible = True
            DengStyle1(i).ZOrder
        Next i
        For i = 0 To 36
            Me.LabBin(i).Caption = ""
        Next
    End If
End If
End Sub
