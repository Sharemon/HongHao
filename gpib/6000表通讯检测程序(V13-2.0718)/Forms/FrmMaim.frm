VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   Caption         =   " "
   ClientHeight    =   6492
   ClientLeft      =   6408
   ClientTop       =   3192
   ClientWidth     =   6840
   Icon            =   "FrmMaim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6492
   ScaleWidth      =   6840
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton CmdDataShow 
      Caption         =   "数据分析"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   180
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1575
      Left            =   600
      TabIndex        =   177
      Top             =   720
      Width           =   5055
      Begin VB.TextBox TxtShowInfo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.2
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   60
         TabIndex        =   179
         Top             =   1236
         Width           =   4932
      End
      Begin VB.TextBox TextShow 
         Alignment       =   1  'Right Justify
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   49.2
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1575
         Left            =   0
         TabIndex        =   178
         Top             =   0
         Width           =   5055
      End
   End
   Begin VB.Timer TimRecover 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6120
      Top             =   0
   End
   Begin VB.Timer TimReset 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   4920
      Top             =   0
   End
   Begin VB.Timer TimInstroction 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5520
      Top             =   0
   End
   Begin VB.CommandButton ComExit 
      Caption         =   "退 出"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton CmdMode 
      Caption         =   "调试模式"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   135
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton CmdReceive 
      Caption         =   "连续接收"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   25
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton ComComOpen 
      Caption         =   "打开端口"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   143
      Top             =   5400
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4740
      LargeChange     =   100
      Left            =   12715
      Max             =   9500
      Min             =   15
      SmallChange     =   50
      TabIndex        =   58
      Top             =   1250
      Value           =   15
      Width           =   255
   End
   Begin VB.Frame FrmInstruct 
      Caption         =   "控制指令COMr简表(供调试使用，YY代表仪表通讯识别号，←┘表示回车符)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5892
      Left            =   6840
      TabIndex        =   27
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox CombInstrumentNo 
         Enabled         =   0   'False
         Height          =   300
         ItemData        =   "FrmMaim.frx":0442
         Left            =   4920
         List            =   "FrmMaim.frx":04A0
         TabIndex        =   140
         Text            =   "01"
         Top             =   510
         Width           =   975
      End
      Begin VB.OptionButton OptionOne 
         Caption         =   "控制单个仪表                                               仪表识别号："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   550
         Width           =   4932
      End
      Begin VB.OptionButton OptionALL 
         Caption         =   "控制当前串口的所有仪表"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   240
         Value           =   -1  'True
         Width           =   4575
      End
      Begin VB.PictureBox Picture1 
         Height          =   4812
         Left            =   0
         ScaleHeight     =   4764
         ScaleWidth      =   7104
         TabIndex        =   28
         Top             =   1080
         Width           =   7155
         Begin VB.PictureBox PicOrder 
            Height          =   14292
            Left            =   -15
            ScaleHeight     =   14244
            ScaleWidth      =   7728
            TabIndex        =   29
            Top             =   -15
            Width           =   7770
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   13
               ItemData        =   "FrmMaim.frx":051C
               Left            =   3360
               List            =   "FrmMaim.frx":0526
               TabIndex        =   189
               Text            =   "00"
               Top             =   12960
               Width           =   735
            End
            Begin VB.CommandButton Command7 
               Caption         =   "4线/2线电阻测量方式"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   187
               Top             =   12960
               Width           =   2100
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   12
               ItemData        =   "FrmMaim.frx":0532
               Left            =   3360
               List            =   "FrmMaim.frx":053C
               TabIndex        =   184
               Text            =   "00"
               Top             =   12600
               Width           =   735
            End
            Begin VB.CommandButton Command6 
               Caption         =   "反转电流(6000-11)型"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   182
               Top             =   12600
               Width           =   2100
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   11
               ItemData        =   "FrmMaim.frx":0548
               Left            =   3360
               List            =   "FrmMaim.frx":0555
               TabIndex        =   169
               Text            =   "00"
               Top             =   12240
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   10
               ItemData        =   "FrmMaim.frx":0565
               Left            =   3360
               List            =   "FrmMaim.frx":056F
               TabIndex        =   168
               Text            =   "00"
               Top             =   11880
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   9
               ItemData        =   "FrmMaim.frx":057B
               Left            =   3360
               List            =   "FrmMaim.frx":058B
               TabIndex        =   167
               Text            =   "00"
               Top             =   11520
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   8
               ItemData        =   "FrmMaim.frx":059F
               Left            =   3360
               List            =   "FrmMaim.frx":05B8
               TabIndex        =   166
               Text            =   "00"
               Top             =   11160
               Width           =   735
            End
            Begin VB.CommandButton Cmdgaugeshowstyle 
               Caption         =   "测量仪显示类型"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   164
               Top             =   12240
               Width           =   2100
            End
            Begin VB.CommandButton Cmdpanelnoyes 
               Caption         =   "面板按键禁止/有效"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   163
               Top             =   11880
               Width           =   2100
            End
            Begin VB.CommandButton Comflicker 
               Caption         =   "数据闪烁方式选择"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   162
               Top             =   11520
               Width           =   2100
            End
            Begin VB.CommandButton Cmdlight 
               Caption         =   "显示亮度调整"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   161
               Top             =   11160
               Width           =   2100
            End
            Begin VB.CommandButton Command5 
               Caption         =   "A/D转换置位"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   154
               Top             =   10800
               Width           =   2100
            End
            Begin VB.CommandButton Command4 
               Caption         =   "COMr29"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   151
               Top             =   10440
               Width           =   2100
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   7
               ItemData        =   "FrmMaim.frx":05D8
               Left            =   3360
               List            =   "FrmMaim.frx":05E8
               TabIndex        =   149
               Text            =   "00"
               Top             =   6480
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   6
               ItemData        =   "FrmMaim.frx":05FC
               Left            =   3360
               List            =   "FrmMaim.frx":0630
               TabIndex        =   145
               Text            =   "01"
               Top             =   5760
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   1
               ItemData        =   "FrmMaim.frx":0674
               Left            =   3360
               List            =   "FrmMaim.frx":067E
               TabIndex        =   139
               Text            =   "00"
               Top             =   4260
               Width           =   735
            End
            Begin VB.TextBox TxtCOMr14 
               Height          =   270
               Left            =   3360
               TabIndex        =   138
               Text            =   "01"
               Top             =   5040
               Width           =   735
            End
            Begin VB.TextBox TxtTimes 
               Height          =   270
               Left            =   3360
               TabIndex        =   136
               Text            =   "10"
               Top             =   1800
               Width           =   615
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   5
               ItemData        =   "FrmMaim.frx":068A
               Left            =   3360
               List            =   "FrmMaim.frx":06A3
               TabIndex        =   132
               Text            =   "04"
               Top             =   7530
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   4
               ItemData        =   "FrmMaim.frx":06C3
               Left            =   3360
               List            =   "FrmMaim.frx":06D0
               TabIndex        =   130
               Text            =   "01"
               Top             =   7160
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   3
               ItemData        =   "FrmMaim.frx":06E0
               Left            =   3360
               List            =   "FrmMaim.frx":06FF
               TabIndex        =   128
               Text            =   "01"
               Top             =   6100
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   2
               ItemData        =   "FrmMaim.frx":0727
               Left            =   3360
               List            =   "FrmMaim.frx":0737
               TabIndex        =   125
               Text            =   "05"
               Top             =   3570
               Width           =   735
            End
            Begin VB.ComboBox Combo1 
               Height          =   276
               Index           =   0
               ItemData        =   "FrmMaim.frx":074B
               Left            =   3360
               List            =   "FrmMaim.frx":0752
               TabIndex        =   122
               Text            =   "01"
               Top             =   1400
               Width           =   735
            End
            Begin VB.CommandButton CmdRangeFast 
               Caption         =   "量程快速选择"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   118
               Top             =   4680
               Width           =   2100
            End
            Begin VB.CommandButton CmdRecover 
               Caption         =   "恢复设定滤波时间"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   57
               Top             =   10080
               Width           =   2100
            End
            Begin VB.CommandButton Cmdlengthen 
               Caption         =   "滤波时间加长"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   56
               Top             =   9720
               Width           =   2100
            End
            Begin VB.CommandButton CmdShorten 
               Caption         =   "滤波时间缩短"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   55
               Top             =   9360
               Width           =   2100
            End
            Begin VB.CommandButton CmdFilterReset 
               Caption         =   "滤波器复位"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   54
               Top             =   9000
               Width           =   2100
            End
            Begin VB.CommandButton CmdFull 
               Caption         =   "串口输出满位(8位)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   53
               Top             =   8280
               Width           =   2100
            End
            Begin VB.CommandButton CmdAlike 
               Caption         =   "输出与显示位数相同"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   52
               Top             =   7920
               Width           =   2100
            End
            Begin VB.CommandButton CmdSendExtremum 
               Caption         =   "发送极值数据类型"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   51
               Top             =   7560
               Width           =   2100
            End
            Begin VB.CommandButton CmdSendMeter 
               Caption         =   "发送测量数据类型"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   50
               Top             =   7200
               Width           =   2100
            End
            Begin VB.CommandButton CmdKeyFunction 
               Caption         =   "串口按键功能"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   49
               Top             =   6120
               Width           =   2100
            End
            Begin VB.CommandButton CmdGallery 
               Caption         =   "信号输入通道切换"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   48
               Top             =   5760
               Width           =   2100
            End
            Begin VB.CommandButton CmdCutBasic 
               Caption         =   "切换至基本单位"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   47
               Top             =   5400
               Width           =   2100
            End
            Begin VB.CommandButton CmdCutUsers 
               Caption         =   "切换至用户单位*"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   46
               Top             =   5040
               Width           =   2100
            End
            Begin VB.CommandButton CmdRange 
               Caption         =   "量程选择"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   45
               Top             =   4320
               Width           =   2100
            End
            Begin VB.CommandButton CmdCut 
               Caption         =   "位数(分辨率)切换"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   44
               Top             =   3960
               Width           =   2100
            End
            Begin VB.CommandButton CmdDigit 
               Caption         =   "位数(分辨率)选择 "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   43
               Top             =   3600
               Width           =   2100
            End
            Begin VB.CommandButton CmdReset 
               Caption         =   "按键发送状态复位"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   42
               Top             =   3240
               Width           =   2100
            End
            Begin VB.CommandButton CmdEliminate 
               Caption         =   "消除极值数据"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   41
               Top             =   2880
               Width           =   2100
            End
            Begin VB.CommandButton CmdAbsolute 
               Caption         =   "置绝对零"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   40
               Top             =   2520
               Width           =   2100
            End
            Begin VB.CommandButton CmdRelative 
               Caption         =   "置相对零"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   39
               Top             =   2160
               Width           =   2100
            End
            Begin VB.CommandButton CmdQuench 
               Caption         =   "仪表显示短时熄灭"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   38
               Top             =   1800
               Width           =   2100
            End
            Begin VB.CommandButton CmdOutputType 
               Caption         =   "输出数据类型选择"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   37
               Top             =   1440
               Width           =   2100
            End
            Begin VB.CommandButton CmdOnce 
               Caption         =   "启动数据单次输出"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   36
               Top             =   360
               Width           =   2100
            End
            Begin VB.CommandButton CmdInitialize 
               Caption         =   "串口通讯初始化"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   35
               Top             =   0
               Width           =   2100
            End
            Begin VB.CommandButton CmdStopContinuation 
               Caption         =   "停止数据连续输出"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   34
               Top             =   1080
               Width           =   2100
            End
            Begin VB.CommandButton CmdContinuation 
               Caption         =   "启动数据连续输出"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   33
               Top             =   720
               Width           =   2100
            End
            Begin VB.CommandButton Command1 
               Caption         =   "SP输出控制"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   32
               Top             =   6480
               Width           =   2100
            End
            Begin VB.CommandButton Command2 
               Caption         =   "COMr19"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   31
               Top             =   6840
               Width           =   2100
            End
            Begin VB.CommandButton Command3 
               Caption         =   "COMr24"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   840
               TabIndex        =   30
               Top             =   8640
               Width           =   2100
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;36;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   83
               Left            =   4560
               TabIndex        =   190
               Top             =   12960
               Width           =   1332
            End
            Begin VB.Label Label10 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   188
               Top             =   12960
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 36"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   82
               Left            =   0
               TabIndex        =   186
               Top             =   12960
               Width           =   1212
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;35;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   81
               Left            =   4560
               TabIndex        =   185
               Top             =   12600
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   183
               Top             =   12600
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 35"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   80
               Left            =   0
               TabIndex        =   181
               Top             =   12600
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;34;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   79
               Left            =   4560
               TabIndex        =   176
               Top             =   12240
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;33;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   78
               Left            =   4560
               TabIndex        =   175
               Top             =   11880
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;32;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   77
               Left            =   4560
               TabIndex        =   174
               Top             =   11520
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;31;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   76
               Left            =   4560
               TabIndex        =   173
               Top             =   11160
               Width           =   1335
            End
            Begin VB.Label Label8 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   172
               Top             =   12240
               Width           =   372
            End
            Begin VB.Label Label7 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   171
               Top             =   11880
               Width           =   372
            End
            Begin VB.Label Label6 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   170
               Top             =   11520
               Width           =   372
            End
            Begin VB.Label Label5 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3000
               TabIndex        =   165
               Top             =   11160
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 34"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   75
               Left            =   0
               TabIndex        =   160
               Top             =   12240
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 33"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   74
               Left            =   0
               TabIndex        =   159
               Top             =   11880
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 32"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   73
               Left            =   0
               TabIndex        =   158
               Top             =   11520
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 31"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   72
               Left            =   0
               TabIndex        =   157
               Top             =   11160
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 30"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   71
               Left            =   0
               TabIndex        =   156
               Top             =   10800
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;30←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   70
               Left            =   4560
               TabIndex        =   155
               Top             =   10800
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 29"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   69
               Left            =   0
               TabIndex        =   153
               Top             =   10440
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;29←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   65
               Left            =   4560
               TabIndex        =   152
               Top             =   10440
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   64
               Left            =   3000
               TabIndex        =   150
               Top             =   6528
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   58
               Left            =   3000
               TabIndex        =   144
               Top             =   5796
               Width           =   372
            End
            Begin VB.Label Label4 
               Caption         =   "(01-99)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Left            =   3960
               TabIndex        =   137
               Top             =   1805
               Width           =   612
            End
            Begin VB.Label LabNote 
               Caption         =   $"FrmMaim.frx":075A
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2172
               Left            =   120
               TabIndex        =   134
               Top             =   13380
               Width           =   5652
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   68
               Left            =   3000
               TabIndex        =   133
               Top             =   7596
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   67
               Left            =   3000
               TabIndex        =   131
               Top             =   7200
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   66
               Left            =   3000
               TabIndex        =   129
               Top             =   6156
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   63
               Left            =   3000
               TabIndex        =   127
               Top             =   5088
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   62
               Left            =   3000
               TabIndex        =   126
               Top             =   4320
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   61
               Left            =   3000
               TabIndex        =   124
               Top             =   3636
               Width           =   372
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   60
               Left            =   3000
               TabIndex        =   123
               Top             =   1836
               Width           =   252
            End
            Begin VB.Label Label1 
               Caption         =   "KK="
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   59
               Left            =   3000
               TabIndex        =   121
               Top             =   1452
               Width           =   252
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;28←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   57
               Left            =   4560
               TabIndex        =   120
               Top             =   10080
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 28"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   56
               Left            =   0
               TabIndex        =   119
               Top             =   10080
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;27←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   55
               Left            =   4560
               TabIndex        =   117
               Top             =   9720
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;26←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   54
               Left            =   4560
               TabIndex        =   116
               Top             =   9360
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;25←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   53
               Left            =   4560
               TabIndex        =   115
               Top             =   9000
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;24←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   52
               Left            =   4560
               TabIndex        =   114
               Top             =   8640
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;23←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   51
               Left            =   4560
               TabIndex        =   113
               Top             =   8280
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;22←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   50
               Left            =   4560
               TabIndex        =   112
               Top             =   7920
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;21;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   49
               Left            =   4560
               TabIndex        =   111
               Top             =   7560
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;20;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   48
               Left            =   4560
               TabIndex        =   110
               Top             =   7200
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;19←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   47
               Left            =   4560
               TabIndex        =   109
               Top             =   6840
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;18;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   46
               Left            =   4560
               TabIndex        =   108
               Top             =   6480
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;17;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   45
               Left            =   4560
               TabIndex        =   107
               Top             =   6120
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;16;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   44
               Left            =   4560
               TabIndex        =   106
               Top             =   5760
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;15←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   43
               Left            =   4560
               TabIndex        =   105
               Top             =   5400
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;14;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   42
               Left            =   4560
               TabIndex        =   104
               Top             =   5040
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;13←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   41
               Left            =   4560
               TabIndex        =   103
               Top             =   4680
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;12;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   40
               Left            =   4560
               TabIndex        =   102
               Top             =   4320
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;11←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   39
               Left            =   4560
               TabIndex        =   101
               Top             =   3960
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;10;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   4560
               TabIndex        =   100
               Top             =   3600
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;09←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   37
               Left            =   4560
               TabIndex        =   99
               Top             =   3240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;08←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   36
               Left            =   4560
               TabIndex        =   98
               Top             =   2880
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;07←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   4560
               TabIndex        =   97
               Top             =   2520
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;06←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   4560
               TabIndex        =   96
               Top             =   2160
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;05;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   33
               Left            =   4560
               TabIndex        =   95
               Top             =   1800
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;04;KK←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   32
               Left            =   4560
               TabIndex        =   94
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;03←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   31
               Left            =   4560
               TabIndex        =   93
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;02←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   30
               Left            =   4560
               TabIndex        =   92
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;01←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   29
               Left            =   4560
               TabIndex        =   91
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "%YY;00←┘"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   28
               Left            =   4560
               TabIndex        =   90
               Top             =   0
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 27"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   27
               Left            =   0
               TabIndex        =   88
               Top             =   9720
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 26"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   26
               Left            =   0
               TabIndex        =   87
               Top             =   9360
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 25"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   25
               Left            =   0
               TabIndex        =   86
               Top             =   9000
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 24"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   24
               Left            =   0
               TabIndex        =   85
               Top             =   8640
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 23"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   23
               Left            =   0
               TabIndex        =   84
               Top             =   8280
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 22"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   22
               Left            =   0
               TabIndex        =   83
               Top             =   7920
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 21"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   0
               TabIndex        =   82
               Top             =   7560
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 20"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   0
               TabIndex        =   81
               Top             =   7200
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 19"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   0
               TabIndex        =   80
               Top             =   6840
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 18"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   0
               TabIndex        =   79
               Top             =   6480
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 17"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   0
               TabIndex        =   78
               Top             =   6120
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 16"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   16
               Left            =   0
               TabIndex        =   77
               Top             =   5760
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 15"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   0
               TabIndex        =   76
               Top             =   5400
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 14"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   0
               TabIndex        =   75
               Top             =   5040
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 13"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   0
               TabIndex        =   74
               Top             =   4680
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 12"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   12
               Left            =   0
               TabIndex        =   73
               Top             =   4320
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 11"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   72
               Top             =   3960
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 10"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   10
               Left            =   0
               TabIndex        =   71
               Top             =   3600
               Width           =   1335
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 09"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   9
               Left            =   0
               TabIndex        =   70
               Top             =   3240
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 08"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   0
               TabIndex        =   69
               Top             =   2880
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 07"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   0
               TabIndex        =   68
               Top             =   2520
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 06"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   67
               Top             =   2160
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 05"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   66
               Top             =   1800
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 04"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   65
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 03"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   64
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 02"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   63
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 01"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   62
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   " COMr 00"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   61
               Top             =   0
               Width           =   1215
            End
         End
      End
      Begin VB.Label LabTitle 
         Caption         =   "指令编号nn         指令内容                    设定范围             ASCII指令格式     "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   840
         Width           =   5895
      End
   End
   Begin VB.Timer TimerFre 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   2880
      Top             =   0
   End
   Begin VB.Frame FrmInf 
      Caption         =   "参数及信息"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   16
      Top             =   3600
      Width           =   6615
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "FrmMaim.frx":07E2
         Left            =   4560
         List            =   "FrmMaim.frx":07FB
         TabIndex        =   147
         Text            =   "100mS"
         Top             =   1275
         Width           =   1575
      End
      Begin VB.ComboBox CombCom 
         Height          =   300
         ItemData        =   "FrmMaim.frx":0830
         Left            =   1200
         List            =   "FrmMaim.frx":0832
         TabIndex        =   0
         Top             =   780
         Width           =   1332
      End
      Begin VB.ComboBox ComBaudRate 
         Height          =   300
         ItemData        =   "FrmMaim.frx":0834
         Left            =   4560
         List            =   "FrmMaim.frx":0859
         TabIndex        =   1
         Text            =   "9600"
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "发送指令间隔："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3000
         TabIndex        =   148
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Shape ShpSendOn 
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3600
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "读取状态"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   2400
         TabIndex        =   24
         Top             =   245
         Width           =   720
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "发送状态"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   4080
         TabIndex        =   23
         Top             =   245
         Width           =   720
      End
      Begin VB.Shape ShpErrOn 
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5280
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "错误指示"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   5760
         TabIndex        =   22
         Top             =   245
         Width           =   720
      End
      Begin VB.Label LabErr 
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   120
         TabIndex        =   21
         Top             =   555
         Width           =   6300
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "当前串口："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   900
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "当前波特率："
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   3000
         TabIndex        =   19
         Top             =   840
         Width           =   1080
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "默认设置：   9600;N;8;1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   120
         TabIndex        =   18
         Top             =   1320
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label LabelOnTop 
         AutoSize        =   -1  'True
         Caption         =   "端口状态"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   720
         TabIndex        =   17
         Top             =   245
         Width           =   720
      End
      Begin VB.Shape ShpSendOff 
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   3600
         Top             =   300
         Width           =   255
      End
      Begin VB.Shape ShpErrOff 
         FillColor       =   &H00000080&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   5280
         Top             =   300
         Width           =   255
      End
      Begin VB.Shape ShpComOn 
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   240
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape ShpRevOn 
         BorderColor     =   &H00004000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1920
         Top             =   300
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape ShpRevOff 
         BorderColor     =   &H00004000&
         FillColor       =   &H00008000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   1920
         Top             =   300
         Width           =   255
      End
      Begin VB.Shape ShpComOff 
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   135
         Left            =   240
         Top             =   300
         Width           =   255
      End
   End
   Begin VB.CommandButton ComZero 
      Caption         =   "置零"
      DisabledPicture =   "FrmMaim.frx":08A7
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "FrmMaim.frx":342B
      TabIndex        =   2
      ToolTipText     =   "F1"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton ComRun 
      Caption         =   "运行"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      Picture         =   "FrmMaim.frx":5FAF
      TabIndex        =   3
      ToolTipText     =   "F2"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton ComShow 
      Caption         =   "显示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      Picture         =   "FrmMaim.frx":8A07
      TabIndex        =   4
      ToolTipText     =   "F3"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton ComDigits 
      Caption         =   "位数"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Picture         =   "FrmMaim.frx":B409
      TabIndex        =   5
      ToolTipText     =   "F4"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton ComRange 
      Caption         =   "量程"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      Picture         =   "FrmMaim.frx":DDEB
      TabIndex        =   9
      ToolTipText     =   "F5"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton ComPrint 
      Caption         =   "打印"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      Picture         =   "FrmMaim.frx":107C1
      TabIndex        =   8
      ToolTipText     =   "F6"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton ComReset 
      Caption         =   "复位"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      Picture         =   "FrmMaim.frx":1312E
      TabIndex        =   7
      ToolTipText     =   "F7"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Timer TimerCheck 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3840
      Top             =   0
   End
   Begin VB.Timer TimerErr 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   4320
      Top             =   0
   End
   Begin VB.Timer TimerReceive 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3360
      Top             =   0
   End
   Begin MSComDlg.CommonDialog ComnDiaFile 
      Left            =   840
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "保存文件为"
      Filter          =   "文本文件|*.txt"
      Flags           =   33792
   End
   Begin VB.Timer TimDetection 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin MSCommLib.MSComm MSComPort 
      Left            =   240
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   2048
      OutBufferSize   =   1024
      ParityReplace   =   0
      RThreshold      =   1
      InputMode       =   1
   End
   Begin VB.Label LabInfo 
      AutoSize        =   -1  'True
      Height          =   300
      Left            =   120
      TabIndex        =   146
      Top             =   6120
      Width           =   6735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "SP2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   220
      TabIndex        =   142
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "SP1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   220
      TabIndex        =   141
      Top             =   1200
      Width           =   375
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   5
      Left            =   5760
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   4
      Left            =   5760
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   3
      Left            =   5760
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   5760
      Top             =   1440
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "mV"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   18
      Left            =   6120
      TabIndex        =   26
      Top             =   1155
      Width           =   270
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   1560
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   600
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Shape ShpPeakOff 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   1560
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "置零"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   960
      TabIndex        =   15
      Top             =   2355
      Width           =   360
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "运行"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   14
      Top             =   2355
      Width           =   390
   End
   Begin VB.Shape ShpZeroOff 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   600
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "6000精密数字测量仪"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   1440
      TabIndex        =   13
      Top             =   300
      Width           =   3300
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "Ω"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   7
      Left            =   6120
      TabIndex        =   12
      Top             =   1395
      Width           =   135
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "KΩ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   8
      Left            =   6120
      TabIndex        =   11
      Top             =   900
      Width           =   255
   End
   Begin VB.Label LabelOnTop 
      AutoSize        =   -1  'True
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   225
      Index           =   9
      Left            =   6120
      TabIndex        =   10
      Top             =   660
      Width           =   105
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   120
      Y1              =   3480
      Y2              =   120
   End
   Begin VB.Line Line4 
      X1              =   6720
      X2              =   6720
      Y1              =   3480
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6720
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   6720
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   5760
      Top             =   1440
      Width           =   255
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   5760
      Top             =   1200
      Width           =   255
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   5760
      Top             =   960
      Width           =   255
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   5760
      Top             =   720
      Width           =   255
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   6
      Left            =   240
      Shape           =   2  'Oval
      Top             =   960
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape ShpOn 
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   7
      Left            =   240
      Shape           =   2  'Oval
      Top             =   1560
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00400040&
      FillColor       =   &H00008000&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   2
      Left            =   240
      Shape           =   2  'Oval
      Top             =   960
      Width           =   195
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00400040&
      FillColor       =   &H00000080&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   3
      Left            =   240
      Shape           =   2  'Oval
      Top             =   1560
      Width           =   195
   End
   Begin VB.Menu MenFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu MenOpenCom 
         Caption         =   "打开端口(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenExit 
         Caption         =   "退出(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenFun 
      Caption         =   "功能(&U)"
      Begin VB.Menu MenF 
         Caption         =   "置零"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenF 
         Caption         =   "运行"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenF 
         Caption         =   "显示"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenF 
         Caption         =   "位数"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenF 
         Caption         =   "量程"
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenF 
         Caption         =   "打印"
         Index           =   5
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenRes 
         Caption         =   "复位"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MenShowData 
         Caption         =   "查看数据"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenToExcel 
         Caption         =   "将记录数据导入Excel"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MenBar10 
         Caption         =   "-"
      End
      Begin VB.Menu MenSend 
         Caption         =   "发送数据"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu Language 
      Caption         =   "语言(&L)"
      Begin VB.Menu Chinese 
         Caption         =   "中文(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu English 
         Caption         =   "English(&E)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MenHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu MenInstr 
         Caption         =   "仪表使用说明(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu MenBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenAbout 
         Caption         =   "关于(&A)"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MenPopup 
      Caption         =   "右键菜单"
      Visible         =   0   'False
      Begin VB.Menu MenMode 
         Caption         =   "调试模式"
      End
      Begin VB.Menu MenBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MenDataShow 
         Caption         =   "查看数据"
      End
      Begin VB.Menu MenBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MenSupport 
         Caption         =   "技术支持"
      End
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "托盘菜单"
      Visible         =   0   'False
      Begin VB.Menu MenuShow 
         Caption         =   "显示窗体"
      End
      Begin VB.Menu MenBar6 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "关于我们"
      End
      Begin VB.Menu MenBar7 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu ReceiveMenu 
      Caption         =   "接收区右键菜单"
      Visible         =   0   'False
      Begin VB.Menu MenuChar 
         Caption         =   "以字符显示"
      End
      Begin VB.Menu MenuHex 
         Caption         =   "以十六进制显示"
      End
      Begin VB.Menu MenBar8 
         Caption         =   "-"
      End
      Begin VB.Menu MenuStop 
         Caption         =   "停止显示"
      End
      Begin VB.Menu MenuReceive 
         Caption         =   "清空接收区"
      End
      Begin VB.Menu MenBar9 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSend 
         Caption         =   "清空发送字符数"
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumArray(36) As String
Dim II As Integer
Dim JJ As Integer
Dim ReceiveAll As String
Dim ReceiveShow As String
Dim StringIn() As Byte
Dim StringTemp(36) As Byte
Dim Median As Integer
Dim Instrument As String
Dim Range As String
Dim Channel As String
Dim DotPosition As Integer
Dim LightStatus As Byte
Dim StateLight As Byte
Dim StateByte5 As Byte
Dim Xianshileixing As String
Dim Anjiancaozuo As String
Dim Fanzhuandianliu As String
Dim Dianzu4or2 As String
Dim BinDatas As String  '接收数据二进制表示
Dim HexDatas As String  '接收数据十六进制表示
Dim ASCIIDatas As String  '接收数据ASCII表示
Dim ABinDatas As String '一个字节的二进制
Dim TxtShowStyleInfo As String
Dim ReceiveFlag As Boolean
Dim StringFlag As Boolean
Dim iZ As Long
'Dim RecordFlag As Boolean
Dim MyPath As String
Dim K As Integer
Dim YY As String
Dim ReceiveType As String
Dim ShowType As String
Dim JudgeData As String
Dim ResetCount As Integer
Public FrmSendShow As Boolean
Private Type PortSettings
    Baud As String
End Type

Dim COMS() As String '串口号数组

Private RT As Integer
Private typSettings As PortSettings
Private sSubKey As String
Private sKeyValue As String
Private hnd As Long
Private Ports() As Variant
Private sSettings As String
Private sPortNum As String
Private sPath As String
Private InputSignal As String

Private Const lMainKey As Long = HKEY_LOCAL_MACHINE
Private Const lLength As Long = 1024
Private Const sSettingsKey As String = "Settings"
Private Const sPortKey As String = "COM"
Private Const sPathKey As String = "Path"


Private Const NOTIFYICON_VERSION = 3
Private Const NOTIFYICON_OLDVERSION = 0

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4

Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
Private Const NIIF_GUID = &H4

Private myData As NOTIFYICONDATA

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type

Private Tooltips As New Collection
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Chinese_Click()
If Lan = 1 Then
If TxtShowInfo.Text = "Port Status：Closed" Then TxtShowInfo.Text = "串口状态：关闭"
LabelOnTop(0).Caption = "6000精密数字测量仪"
LabelOnTop(1).Caption = "置零"
LabelOnTop(2).Caption = "运行"
LabelOnTop(3).Caption = "发送指令间隔："
FrmMain.Caption = "6000表通讯检测程序(V13-2.0718)"
ComZero.Caption = "置零"
ComRun.Caption = "运行"
ComShow.Caption = "显示"
ComDigits.Caption = "位数"
ComRange.Caption = "量程"
ComPrint.Caption = " 打印"
ComReset.Caption = "复位"
ComComOpen.Caption = "打开端口"
CmdDataShow.Caption = "数据分析"
ComExit.Caption = "退 出"
LabInfo.Caption = "准备完毕"
MenFile.Caption = "文件(&F)"
MenOpenCom.Caption = "打开端口(&O)"
MenExit.Caption = "退出&Q)"
MenFun.Caption = "功能(&U)"
MenF.Item(0).Caption = "置零"
MenF.Item(1).Caption = "运行"
MenF.Item(2).Caption = "显示"
MenF.Item(3).Caption = "位数"
MenF.Item(4).Caption = "量程"
MenF.Item(5).Caption = "打印"
MenRes.Caption = "复位"
MenToExcel.Caption = "将记录数据导入Excel"
MenShowData.Caption = "查看数据"
MenSend.Caption = "发送数据"
Language.Caption = "语言(&L)"
MenHelp.Caption = "帮助(&H)"
MenInstr.Caption = "仪表使用说明(&R)"
MenAbout.Caption = "关于(&A)"
FrmInf.Caption = "参数及信息"
LabelOnTop(10).Caption = "端口状态"
LabelOnTop(11).Caption = "读取状态"
LabelOnTop(12).Caption = "发送状态"
LabelOnTop(13).Caption = "错误指示"
LabelOnTop(14).Caption = "当前串口："
LabelOnTop(15).Caption = "当前波特率："
LabelOnTop(16).Caption = "默认设置：   9600;N;8;1"
If CmdMode.Caption = "Detection mode" Then
   CmdMode.Caption = "检测模式"
ElseIf CmdMode.Caption = "Debug mode" Then
   CmdMode.Caption = "调试模式"
End If
If CmdReceive.Caption = "Receive by Command" Then
   CmdReceive.Caption = "指令接收"
ElseIf CmdReceive.Caption = "Receive Continuously" Then
   CmdReceive.Caption = "连续接收"
End If
If ReceiveType = "Instructions" Then
ReceiveType = "指令"
ElseIf ReceiveType = "Continuous" Then
ReceiveType = "连续"
End If
If ShowType = "Debug" Then
ShowType = "调试模式"
ElseIf ShowType = "Detect" Then
ShowType = "检测模式"
End If
MenuShow.Caption = "显示窗体"
MenuAbout.Caption = "关于"
MenuExit.Caption = "退出"
CmdInitialize.Caption = "串口通讯初始化"
CmdOnce.Caption = "启动数据单次输出"
CmdContinuation.Caption = "启动数据连续输出"
CmdStopContinuation.Caption = "停止数据连续输出"
CmdOutputType.Caption = "输出数据类型选择"
CmdQuench.Caption = "仪表显示短时熄灭"
CmdRelative.Caption = "置相对零"
CmdAbsolute.Caption = "置绝对零"
CmdEliminate.Caption = "消除极值数据"
CmdReset.Caption = "按键发送状态复位"
CmdDigit.Caption = "位数(分辨率)选择"
CmdCut.Caption = "位数(分辨率)切换"
CmdRange.Caption = "量程选择"
CmdRangeFast.Caption = "量程快速选择"
CmdCutUsers.Caption = "切换至用户单位"
CmdCutBasic.Caption = "切换至基本单位"
CmdGallery.Caption = "信号输入通道切换"
CmdKeyFunction.Caption = "串口按键功能"
CmdSendMeter.Caption = "发送测量数据类型"
CmdSendExtremum.Caption = "发送极值数据类型"
CmdAlike.Caption = "输出与显示位数相同"
CmdFull.Caption = "串口输出满位(8位)"
CmdFilterReset.Caption = "滤波器复位"
CmdShorten.Caption = "滤波器时间缩短"
Cmdlengthen.Caption = "滤波器时间加长"
CmdRecover.Caption = "恢复设定滤波时间"
Cmdlight.Caption = "显示亮度调整"
Command1.Caption = "SP输出控制"
Command5.Caption = "A/D转换置位"
Comflicker.Caption = "数据闪烁方式选择"
Cmdpanelnoyes.Caption = "面板按键禁止/有效"
Cmdgaugeshowstyle.Caption = "测量仪显示类型"
Command6.Caption = "反转电流(6000-11)型"
Command7.Caption = "4线/2线电阻测量方式"
LabNote.Caption = "注意:两次发送指令的间隔时间应大于显示刷新率的间隔时间，如显示刷新率6次/秒，指令间隔时间应大于0.18秒，建议两次指令间隔时间大于0.2秒。"
OptionALL.Caption = "控制当前串口的所有仪表"
OptionOne.Caption = "控制单个仪表                                              仪表识别号："
LabTitle.Caption = "指令编号nn         指令内容                    设定范围             ASCII指令格式     "
FrmInstruct.Caption = "控制指令COMr简表(供调试使用，YY代表仪表通讯识别号，←┘表示回车符)"

FrmDataShow.Caption = "接收数据分析"
FrmDataShow.Label1(0).Caption = "字节号"
FrmDataShow.Label1(1).Caption = "十六进制"
FrmDataShow.Label1(2).Caption = "ASCII"
FrmDataShow.Option1(0).Caption = "二进制"
FrmDataShow.Option1(1).Caption = "图形显示"
If FrmDataShow.CmdOK(0).Caption = "Pause" Then FrmDataShow.CmdOK(0).Caption = "暂停"
If FrmDataShow.CmdOK(0).Caption = "Continue" Then FrmDataShow.CmdOK(0).Caption = "继续"
FrmDataShow.CmdOK(1).Caption = "关闭"
FrmDataShow.FrmDataSaveStatus.Caption = "数据记录 - 记录数据未开启"
FrmDataShow.Frame2(0).Caption = "记录形式选择"
FrmDataShow.OptionSaveStyle(0).Caption = "十六进制形式保存"
FrmDataShow.OptionSaveStyle(1).Caption = "ASCII形式保存"
FrmDataShow.OptionSaveStyle(2).Caption = "二进制形式保存"
FrmDataShow.Frame2(1).Caption = "保存规则选择"
FrmDataShow.OptionSavegz(0).Caption = "逐条记录数据到文本中"
FrmDataShow.OptionSavegz(1).Caption = "连续记录数据到文本中"
FrmDataShow.CmdSave(0).Caption = "开始记录"
FrmDataShow.CmdSave(1).Caption = "停止记录"
FrmDataShow.CmdSave(2).Caption = "选择路径..."
FrmDataShow.Label4(0).Caption = "接收帧数："
FrmDataShow.Label4(1).Caption = "有效帧数："
FrmDataShow.Label4(2).Caption = "重复帧数："
FrmDataShow.Command1.Caption = "清零"
Lan = 0
End If
Call lanResize
End Sub

Private Sub CmdAbsolute_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";07" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdAbsolute_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "置绝对零位，[%YY;07←┘]，收到指令后仪表将调出校准时的绝对零位数据，仪表此后显示的数据为绝对值数据。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set absolute zero,[%YY;07←┘], After receiving instruction instrument will bring up the absolute zero calibration data, instrument data, then display the data for the absolute value."
End If
End Sub

Private Sub CmdAlike_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";22" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdAlike_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "置串口输出数据与仪表显示数据位数相同，[%YY;22←┘]，串口通讯输出或打印输出的位数与显示位数相同。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Serial output data set and the same instrument display data bits, [%YY;22←┘], Serial communication output or printout of the median and display the same median."
End If
End Sub

Private Sub CmdContinuation_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";02" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdContinuation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "启动数据连续输出，[%YY;02←┘]，编号为YY仪表收到该指令后连续发送测量数据，在串口通讯波特率足够快时，其数据输出的刷新率与显示刷新率同步，除特殊用途外不建议使用本方式进行数据通讯。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Start data continuous output, [%YY;02←┘], code instrumentation YY after receiving the order to send a continuous measurement data, except for special purposes, does not recommend using this approach to data communication"
End If
End Sub

Private Sub CmdCut_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";11" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdCut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "显示位数(分辨率)快速切换，[%YY;11←┘]，收到指令后仪表将调整到下一个位数的显示分辨率。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Show the median (resolution) fast switching, [% YY;11←┘], after receiving instructions to adjust the instrument to the next digit display resolution."
End If
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabInfo.Caption = ""
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabInfo.Caption = ""
End Sub

Private Sub CmdCutBasic_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";15" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdCutBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "切换至基本测量单位(mV、V、Ω、mV/V)，[%YY;15←┘]，与COMr14指令相似，需添加相应的工作模块。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Switch to basic measurement unit (mV,V,Ω,mV/V), [%YY;15←┘], and COMr14 command similar to the work you need to add the corresponding module."
End If
End Sub

Private Sub CmdCutUsers_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";14;" & Trim(TxtCOMr14.Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdCutUsers_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "切换至用户设定单位*，[%YY;14;KK←┘]，需附加相应的工作模块，收到指令后仪表将显示数据切换至用户设定的单位*，KK为用户单位编号。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Switch to the user to set the unit *, [%YY;14;KK←┘], need to add the appropriate modules of work, after receiving instructions to switch to the meter will display the data the user to set the unit *, KK is the user unit code."
End If
End Sub

Private Sub CmdDataShow_Click()
If ISDataShow = False Then
    ISDataShow = True
    BinOrDeng = False
    FrmMain.Left = (Screen.Width - FrmMain.Width - FrmDataShow.Width) / 2
    FrmDataShow.Left = FrmMain.Left + FrmMain.Width
    FrmDataShow.Top = (Screen.Height - FrmDataShow.Height) / 2
    FrmDataShow.Show
Else
    ISDataShow = False
    Unload FrmDataShow
    FrmMain.Left = (Screen.Width - FrmMain.Width) / 2
End If
End Sub

Private Sub CmdDigit_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";10;" & Combo1(2).Text & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdDigit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "显示位数(分辨率)选择，[%YY;10;KK←┘]，KK=05-08为设定要显示的位数。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Show the median (resolution) to select,[%YY;10;KK←┘], KK = 05-08 for the set of bits to be displayed"
End If
End Sub

Private Sub CmdEliminate_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";08" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdEliminate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "消除极值数据，[%YY;08←┘]，收到指令后仪表将清除此前所保留的极值数据，以便进行后续测量工作。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Elimination of extreme value data,[%YY;08←┘], after receiving instructions to clear the instrument previously reserved extreme data, to facilitate later measurements."
End If
End Sub

Private Sub CmdFilterReset_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";25" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdFilterReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "滤波器复位，[%YY;25←┘]，仪表内部滤波器硬件和软件初始化，以便快速跟踪变化的活跃信号，这个指令在滤波时间设定较大时想过比较明显，在滤波时间小于0.5秒时作用有限，执行指令后可能会出现短时间的数据波动时正常现象。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Filter reset, [%YY;25←┘], instrument hardware and software within the filter is initialized in order to fast track changes in the active signal."
End If
End Sub

Private Sub CmdFull_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";23" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdFull_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "置串口发送数据为满为数据(8位)输出，[%YY;23←┘]，仪表收到指令后以最大的位数(8位)输出的数据。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set serial COM to send data over the data (8 bit) output, [%YY;23←┘], instrument instruction received the greatest number of bits (8 bits) output data."
End If
End Sub

Private Sub CmdGallery_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";16;" & Combo1(6).Text & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdGallery_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "信号输入通道切换(需添加多通道输入转换模块)，[%YY;16←┘]，收到指令后仪表切换至第KK输入通道。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Signal input channel switching (to add multi-channel input conversion module), [%YY;16←┘], after receiving instructions switching to KK instrument input channels."
End If
End Sub

Private Sub Cmdgaugeshowstyle_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";34;" & Trim(Combo1(11).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Cmdgaugeshowstyle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "显示类型选择指令，[%YY;34;KK ]，KK取值范围是0~2，本指令仅影响仪表显示的内容，对于串口数据通讯的传输和指令接收均不产生影响。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Display and select the type of instruction, [% YY; 34; KK], the range is 0 to 2, the instruction affects only the instrument displays the contents of the receiver does not have an impact, for the transmission of serial data communication and instruction."
End If
End Sub

Private Sub CmdInitialize_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";08" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdInitialize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "串口通讯初始化，[%YY;00←┘]，该指令使串口输出终止，等待新的命令。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Serial initializing,[%YY;00←┘], this instruction to terminate, waiting for a new serial output command"
End If
End Sub

Private Sub CmdKeyFunction_Click()
      
On Error GoTo ErrHndl

   MSComPort.Output = Trim("%" & YY & ";17;" & Trim(Combo1(3).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Cmdlengthen_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";27" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Cmdlengthen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "滤波器滤波时间加大2倍，[%YY;27←┘]，可以使测量数据的稳定性增加。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Filter filter time increased 2 times, [%YY;27←┘], can increase the stability of the measurement data."
End If
End Sub

Private Sub Cmdlight_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";31;" & Trim(Combo1(8).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Cmdlight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "测量仪显示亮度调整，[%YY;31;KK ]，KK取值范围是1~6，收到本指令后仪表根据KK值调整显示器的亮度，KK=1时显示亮度最暗，KK=0或KK=6时显示亮度最大。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Meter display brightness adjustment, [% YY; 31; KK, ?] KK range is 1 to 6, after receiving this instruction, instrument KK value adjustments on the brightness of the display, the display brightness is the darkest KK, = 1, KK, =0 or KK = 6, the display brightness."
End If
End Sub

Private Sub CmdMode_Click()
If CmdMode.Caption = "调试模式" Or CmdMode.Caption = "Debug mode" Then
    If Lan = 0 Then
      CmdMode.Caption = "检测模式"
      ShowType = "调试模式"
    ElseIf Lan = 1 Then
      CmdMode.Caption = "Detection mode"
      ShowType = "Debug"
    End If
    TimInstroction = False
    If MSComPort.PortOpen = True Then
    MSComPort.Output = Trim("%00;02" & vbCr)
    End If
    OptionALL.Value = True
    Picture1.Enabled = True
    FrmMain.Width = 13140
    LabInfo.Width = 12910
    If ISDataShow = True Then
        FrmMain.Left = (Screen.Width - FrmDataShow.Width - FrmMain.Width) / 2
        FrmDataShow.Left = FrmMain.Left + FrmMain.Width
    Else
        FrmMain.Left = (Screen.Width - FrmMain.Width) / 2
    End If
ElseIf CmdMode.Caption = "检测模式" Or CmdMode.Caption = "Detection mode" Then
    If Lan = 0 Then
       CmdMode.Caption = "调试模式"
       ShowType = "检测模式"
    ElseIf Lan = 1 Then
       CmdMode.Caption = "Debug mode"
       ShowType = "Detect"
    End If
    OptionALL.Value = False
    OptionOne.Value = False
    Picture1.Enabled = False
    FrmMain.Width = 6930
    LabInfo.Width = 6850
    If ISDataShow = True Then
        FrmMain.Left = (Screen.Width - FrmDataShow.Width - FrmMain.Width) / 2
        FrmDataShow.Left = FrmMain.Left + FrmMain.Width
    Else
        FrmMain.Left = (Screen.Width - FrmMain.Width) / 2
    End If
End If
End Sub

Private Sub CmdMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "打开调试模式(关闭检测模式)或者打开检测模式(关闭调试模式)。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Open the debug mode (turn off detection mode) or open detection mode (turn off debug mode)."
End If
End Sub

Private Sub CmdDataShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "查看数据接收，并可保存接收的数据。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Show the data received, the received data can be saved."
End If
End Sub

Private Sub CmdOnce_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";01" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdOnce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "启动串口数据单次输出，[%YY;01←┘]，编号为YY仪表收到该指令后输出一帧测量数据并进入指令等待状态，在串口设定于指令输出方式(br48=1)时最常用的数据输出指令。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Start a single serial data output,[%YY;00←┘],number is YY instrument after receiving the order data and output a measurement wait state into the instruction."
End If
End Sub

Private Sub CmdOutputType_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";04;01" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdOutputType_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "通讯输出数据类型选择，[%YY;04←┘]，编号为YY仪表收到该指令后按KK所要求的类型(格式)输出数据，常规产品只提供一种格式。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Communication output data type selection,[%YY;04←┘], numbering as YY instrument after receiving the order requested by KK type (format) output data, only one format of conventional products."
End If
End Sub

Private Sub Cmdpanelnoyes_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";33;" & Trim(Combo1(10).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Cmdpanelnoyes_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "面板按键操作禁止/有效指令，[%YY;33;KK ]，KK取值范围是0~1,00允许按键，01禁止按键。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Panel key operation to prohibit / effective instruction,% YY; 33; KK], the range is 0 to 1,00 to allow key, 01 prohibit button."
End If
End Sub

Private Sub CmdQuench_Click()
   
On Error GoTo ErrHndl
   
   TextShow.Text = ""
   MSComPort.Output = Trim("%" & YY & ";05;" & Trim(TxtTimes.Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdQuench_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "仪表显示短时熄灭(闪烁)，[%YY;05；KK←┘]，KK为仪表显示熄灭的时间(KK X 0.01秒)，然后重新显示，该指令对仪表的测量和工作无任何影响，主要用于警示。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Short-term instruments show off (flashing),[%YY;05；KK←┘], KK is the instrument display off time (KK X 0.01 seconds), then re-show, the instruction on the instrument's measurement and work without any effect, mainly used for warning."
End If
End Sub

Private Sub CmdRange_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";12;" & Combo1(1).Text & vbCr)
   TimerCheckStart
   TimReset.Enabled = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdRange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "量程转换，[%YY;12;KK←┘]，KK为量程识别编号(量程识别编号不能更改)。收到指令后仪表将更换至量程识别号为KK量程。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Range conversion, [%YY;12;KK ←┘], KK identification number for the range (range identification number can not be changed). Instrument will be replaced after receipt of order to the range identifier for the KK scale."
End If
End Sub

Private Sub CmdRangeFast_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";13" & vbCr)
   TimerCheckStart
   TimReset.Enabled = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdRangeFast_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "量程快速转换，[%YY;13←┘]，收到指令后仪表切换至下一个量程。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Fast switching range, [%YY;13←┘], after receiving instructions switching to the next meter range."
End If
End Sub

Private Sub CmdReceive_Click()
MSComPort.InBufferCount = 0
MSComPort.OutBufferCount = 0
delay 100
If Picture1.Enabled = True Then Picture1.Enabled = False
If OptionALL.Value = True Then OptionALL.Value = False
If OptionOne.Value = True Then OptionOne.Value = False
If CmdReceive.Caption = "指令接收" Or CmdReceive.Caption = "Receive by Command" Then
   If MSComPort.PortOpen = True Then
   RT = 1
    If Lan = 0 Then
       CmdReceive.Caption = "连续接收"
       ReceiveType = "指令"
    ElseIf Lan = 1 Then
       CmdReceive.Caption = "Receive Continuously"
       ReceiveType = "Instructions"
    End If
    MSComPort.Output = Trim("%00;03" & vbCr)
    TimInstroction = True
  End If
ElseIf CmdReceive.Caption = "连续接收" Or CmdReceive.Caption = "Receive Continuously" Then
   If MSComPort.PortOpen = True Then
   RT = 2
    If Lan = 0 Then
       CmdReceive.Caption = "指令接收"
       ReceiveType = "连续"
    ElseIf Lan = 1 Then
        CmdReceive.Caption = "Receive by Command"
       ReceiveType = "Continuous"
    End If
    TimInstroction = False
    MSComPort.Output = Trim("%00;02" & vbCr)
  End If
End If
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
End Sub

Private Sub CmdReceive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "用于在检测模式下指令接收和连续接收相互切换。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "To receive instruction in the detection mode and continuous reception of switching between"
End If
End Sub

Private Sub CmdRecover_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";28" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdRecover_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "恢复设定滤波时间，[%YY;28←┘]，恢复仪表开机或重新启动时所设定的时间。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set the filter to restore the time, [%YY;28←┘], to resume the boot or restart the instrument when the set time."
End If
End Sub

Private Sub CmdRelative_Click()

On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";06" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdRelative_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "置相对零，[%YY;06←┘]，收到指令后仪表将此时的绝对测量数据作为相对零位，仪表此后显示的数据为相对值数据。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set relative to zero,[%YY;06←┘], after the instrument will receive instructions at this time as the relative absolute zero measurements, instrumentation, then display the data for the absolute value of the data."
End If
End Sub

Private Sub CmdReset_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";09" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "按键状态复位，[%YY;09←┘]，收到指令后仪表将状态字节3中的bit0-5置0，以便进行后续按键状态判定。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "State reset button,[%YY;09←┘], after the instrument will receive the state bit 0-5 in the instruction byte 3 set to 0, in order to determine the status of follow-up button"
End If
End Sub


Private Sub CmdStart_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%00;02" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdStart_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "启动数据的输出（此命令只适用于某些特殊仪器）"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Start data output（This command applies only to specific instruments）"
End If
End Sub

Private Sub CmdStop_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%00;03" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdStop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "停止数据的输出（此命令只适用于某些特殊仪器）"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Stop data output（This command applies only to specific instruments）"
End If
End Sub

Private Sub CmdSendExtremum_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";21;" & Trim(Combo1(5).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdSendMeter_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";20;" & Trim(Combo1(4).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdSendMeter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "发送测量数据类型选择，[%YY;20;KK←┘]，KK所对应的显示内容如下：[%YY;20;01←┘]显示绝对值数据，[%YY;20;02←┘]显示相对值数据，[%YY;20;03←┘]相对零位的绝对值数据。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Send measurement data type selection, [%YY;20;KK←┘], KK corresponding to the display as follows: [%YY;20;01←┘] shows the absolute value of the data, [%YY;20;02←┘] shows the relative zero data, [%YY;20;03←┘] the absolute value of the data relative zero."
End If
End Sub

Private Sub CmdShorten_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";26" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdShorten_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "滤波器滤波时间缩小2倍，[%YY;26←┘]，用于快速跟踪阶跃信号，仪表执行指令后可能会出现数据波动现象。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Filter filter time reduced 2 times, [%YY;26←┘], for fast tracking step signals, meters may occur after the implementation of instruction data fluctuation."
End If
End Sub

Private Sub CmdStopContinuation_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";03" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub CmdStopContinuation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "停止数据连续输出，[%YY;03←┘]，编号为YY仪表收到该指令后停止发送测量数据，进入指令等待状态。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Stop data continuous output,[%YY;03←┘], numbered YY instrument after receiving the order to stop sending measurement data into the command wait state."
End If
End Sub

Private Sub CombInstrumentNo_Click()
If Lan = 0 Then
ShowType = "调试模式"
ElseIf Lan = 1 Then
ShowType = "Debug Mode"
End If
YY = Trim(CombInstrumentNo.List(CombInstrumentNo.ListIndex))
End Sub

Private Sub Combo2_Click()
TimInstroction.Interval = CDbl(Left(Combo2.Text, Len(Combo2.Text) - 2))
End Sub

Private Sub Comflicker_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";32;" & Trim(Combo1(9).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Comflicker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "显示闪烁方式选择，[%YY;32;KK ]，KK取值范围是0~3，收到本指令后仪表根据KK值调整显示器的闪烁类型。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Flashing mode selection [% YY; 32; KK], the range is 0 to 3, after receiving this instruction, instrument according to the KK value adjustments on the monitor flashing type."
End If
End Sub

Private Sub Command1_Click()
      
On Error GoTo ErrHndl

   MSComPort.Output = Trim("%" & YY & ";18;" & Trim(Combo1(7).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "输出继电器控制"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Relay control"
End If
End Sub

Private Sub Command5_Click()
   
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";30" & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "A/D转换置位"
ElseIf Lan = 1 Then
   LabInfo.Caption = "A/D Convertion"
End If
End Sub

Private Sub Command6_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";35;" & Trim(Combo1(12).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Command7_Click()
On Error GoTo ErrHndl
   
   MSComPort.Output = Trim("%" & YY & ";36;" & Trim(Combo1(13).Text) & vbCr)
   TimerCheckStart
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, myData
Call SetWindowLong(Me.hWnd, GWL_WNDPROC, OrgWinRet)
MenExit_Click
End Sub

Private Sub FrmInstruct_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "6000仪表的众多功能和状态可以通过串口指令进行控制和调整，COMr指令格式为：[%YY;nn;KK←┘]，YY为00时所有的仪表均执行所发送的指令内容。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "6000 instrument state can be many functions and commands through the serial COM to control and adjust, COMr instruction format:[%YY;nn;KK←┘], when YY is 00 when all the instruments are the implementation of Directive"
End If
End Sub

Private Sub MenRes_Click()
ComReset_Click
End Sub

Private Sub MenSend_Click()
If FrmSendShow = False Then
    FrmSend.Show vbModeless
    FrmSend.Left = FrmMain.Left + FrmMain.Width - FrmSend.Width
    If Lan = 1 Then
    FrmSend.Caption = "Send Instruction"
    FrmSend.OptSendASC.Caption = "ASCII Instruction"
    FrmSend.OptSendHex.Caption = "Hex Instruction"
    FrmSend.LabelOnTop(20).Caption = "Send cycle"
    FrmSend.LabelOnTop(21).Caption = "Ms time"
    FrmSend.ChkSent.Caption = "Automatically sent"
    FrmSend.CmdEmpty.Caption = "Clear fill"
    FrmSend.CmdManual.Caption = "Manual Send"
    FrmSend.CmdExit.Caption = "Exit"
    ElseIf Lan = 0 Then
    FrmSend.Caption = "发送指令(用于命令调试)"
    FrmSend.OptSendASC.Caption = "ASCII指令"
    FrmSend.OptSendHex.Caption = "十六进制指令"
    FrmSend.LabelOnTop(20).Caption = "自动发送周期"
    FrmSend.LabelOnTop(21).Caption = "毫秒/次"
    FrmSend.ChkSent.Caption = "自动发送"
    FrmSend.CmdEmpty.Caption = "清空重填"
    FrmSend.CmdManual.Caption = "手动发送"
    FrmSend.CmdExit.Caption = "退出"
    End If
    FrmSendShow = True
ElseIf FrmSendShow = True Then
    Unload FrmSend
    Set FrmSend = Nothing
    FrmSendShow = False
End If
End Sub
'
Private Sub MenToExcel_Click()
On Error GoTo ErrHandler
'If OpenFileName = "" Then
   ComnDiaFile.ShowOpen
   OpenFileName = ComnDiaFile.FileName
   If OpenFileName = "" Then Exit Sub
   TxtToExcel OpenFileName, Space(1)
'End If
ErrHandler:
    Exit Sub

End Sub

Private Sub MenShowData_Click()
CmdDataShow_Click
End Sub

Private Sub MenMode_Click()
CmdMode_Click
End Sub

Private Sub MenDataShow_Click()
CmdDataShow_Click
End Sub

Private Sub MenuAbout_Click()
   FrmAbout.Show vbModal
End Sub

Private Sub MenuExit_Click()
Unload Me
End Sub

Sub TxtToExcel(txtFile As String, DistanceChar As String)
    If Dir((Left(txtFile, Len(txtFile) - 4) & ".xls")) <> "" Then
      MsgBox "文件" & (Left(txtFile, Len(txtFile) - 4) & ".xls") & "已存在，可能您已经转换过了。"
      OpenFileName = ""
      Exit Sub
    End If
On Error GoTo L

    '建立excel对象
    Dim XlApp As Excel.Application
    Dim XlWb As Excel.Workbook
    Dim XlSt As Excel.Worksheet
    Set XlApp = CreateObject("Excel.Application")
    If XlApp Is Nothing Then
       MsgBox "请检查本机是否安装了Microsoft Excel软件"
       Exit Sub
    End If
    Set XlWb = XlApp.Workbooks.Add
    Set XlSt = XlWb.Worksheets(1)
    XlApp.ActiveWorkbook.SaveAs (Left(txtFile, Len(txtFile) - 4) & ".xls")
    XlApp.Visible = False
    XlApp.Rows.HorizontalAlignment = xlVAlignCenter
    '开始转换
    Dim i As Integer, j As Integer, linenext As String, strb() As String
    j = 1
    XlSt.Columns(1).ColumnWidth = 6
    XlSt.Columns(2).ColumnWidth = 37
    XlSt.Columns(3).ColumnWidth = 12
    XlSt.Columns(4).ColumnWidth = 12
    XlSt.Columns(5).ColumnWidth = 12
    Open txtFile For Input As #1
        Do Until EOF(1)
            Line Input #1, linenext
            strb = Split(linenext, DistanceChar)
            For i = 0 To UBound(strb)
                If j = 1 Then
                XlSt.Range(XlSt.Cells(1, 1), XlSt.Cells(1, 3)).Merge
                XlSt.Cells(j, 1) = "注：分别显示的是：编号、原始数据、仪表显示数据、日期、时间。"
                Else
                XlSt.Cells(j, i + 1) = strb(i)
                End If
            Next
            j = j + 1
        Loop
    Close #1
    XlApp.ActiveWorkbook.Close True
    XlApp.Quit
    Set XlSt = Nothing
    Set XlWb = Nothing
    Set XlApp = Nothing
    OpenFileName = ""
    MsgBox "导入成功"
L:
    Exit Sub
    MsgBox "转换中出现错误"
    Set XlSt = Nothing
    If Not XlWb Is Nothing Then
       XlWb.Close
       Set XlWb = Nothing
    End If
    If Not XlApp Is Nothing Then
    Shell "cmd.exe /c taskkill /f /im excel.exe"
    End If
    OpenFileName = ""
End Sub

Private Sub MenuShow_Click()
With FrmMain
 Select Case CLng(lp_id)
   Case WM_LBUTTONUP '左键
    If .WindowState = vbMinimized Then
     Status = STA_NORMAL
     .Visible = True
     .SetFocus
     .WindowState = vbNormal
    Else
     Status = STA_MIN
     .WindowState = vbMinimized
     .Visible = False
    End If
 End Select
End With
End Sub

Private Sub English_Click()
If Lan = 0 Then
If TxtShowInfo.Text = "串口状态：关闭" Then TxtShowInfo.Text = "Port Status：Closed"
LabelOnTop(0).Caption = "6000 Digital Meter"
LabelOnTop(1).Caption = "Zero"
LabelOnTop(2).Caption = "Run"
LabelOnTop(3).Caption = "Command Delay:"
FrmMain.Caption = "6000 Digital Meter(V13-2.0718)"
ComZero.Caption = "Zero"
ComRun.Caption = "Run"
ComShow.Caption = "Display"
ComDigits.Caption = "Decimal"
ComRange.Caption = "Range"
ComPrint.Caption = "Print"
ComReset.Caption = "Reset"
ComComOpen.Caption = "Open Port"
CmdDataShow.Caption = "Data analysis"
ComExit.Caption = "Exit"
LabInfo.Caption = "Readly"
MenFile.Caption = "File(&F)"
MenOpenCom.Caption = "Open Port(&O)"
MenExit.Caption = "Exit(&Q)"
MenFun.Caption = "Function(&U)"
MenF.Item(0).Caption = "Zero"
MenF.Item(1).Caption = "Run"
MenF.Item(2).Caption = "Display"
MenF.Item(3).Caption = "Decimal"
MenF.Item(4).Caption = "Range"
MenF.Item(5).Caption = "Print"
MenRes.Caption = "Reset"
MenToExcel.Caption = "Data To Excel"
MenShowData.Caption = "Data show"
MenSend.Caption = "Send data"
Language.Caption = "Language(&L)"
MenHelp.Caption = "Help(&H)"
MenInstr.Caption = "Instruction(&R)"
MenAbout.Caption = "About(&A)"
FrmInf.Caption = "Parameters information"
LabelOnTop(10).Caption = "Port"
LabelOnTop(11).Caption = "Receive"
LabelOnTop(12).Caption = "Send"
LabelOnTop(13).Caption = "Error"
LabelOnTop(14).Caption = "Port:"
LabelOnTop(15).Caption = "Baud rate:"
LabelOnTop(16).Caption = "Default:          9600;N;8;1"
If CmdMode.Caption = "检测模式" Then
   CmdMode.Caption = "Detection mode"
ElseIf CmdMode.Caption = "调试模式" Then
   CmdMode.Caption = "Debug mode"
End If
If CmdReceive.Caption = "指令接收" Then
    CmdReceive.Caption = "Receive by Command"
ElseIf CmdReceive.Caption = "连续接收" Then
   CmdReceive.Caption = "Receive Continuously"
End If
If ReceiveType = "指令" Then
ReceiveType = "Instructions"
ElseIf ReceiveType = "连续" Then
ReceiveType = "Continuous"
End If
If ShowType = "调试模式" Then
ShowType = "Debug"
ElseIf ShowType = "检测模式" Then
ShowType = "Detect"
End If
MenuShow.Caption = "Display window"
MenuAbout.Caption = "About"
MenuExit.Caption = "exit"
CmdInitialize.Caption = "Port initialize"
CmdOnce.Caption = "Single Output"
CmdContinuation.Caption = "Continuous output"
CmdStopContinuation.Caption = "Stop data output"
CmdOutputType.Caption = "Output Type"
CmdQuench.Caption = "Meter Quench"
CmdRelative.Caption = "Relative to zero"
CmdAbsolute.Caption = "Absolute zero"
CmdEliminate.Caption = "Eliminate extreme"
CmdReset.Caption = "Reset button"
CmdDigit.Caption = "Digit select"
CmdCut.Caption = "Digit Cutover"
CmdRange.Caption = "Range selection"
CmdRangeFast.Caption = "Quick selection"
CmdCutUsers.Caption = "Switch to user unit"
CmdCutBasic.Caption = "Switch to basic unit"
CmdGallery.Caption = "Channel Select"
CmdKeyFunction.Caption = "Key features"
CmdSendMeter.Caption = "Data type"
CmdSendExtremum.Caption = "Extremum Type"
CmdAlike.Caption = "Simultaneous"
CmdFull.Caption = "Full bit output"
CmdFilterReset.Caption = "Filter Reset"
CmdShorten.Caption = "Shorter filter time"
Cmdlengthen.Caption = "Longer filter time"
CmdRecover.Caption = "Reset filter time"
Cmdlight.Caption = "Brightness"
Command1.Caption = "SP Output Control"
Command5.Caption = "A/D Conversion"
Comflicker.Caption = "Data flicker"
Cmdpanelnoyes.Caption = "Key Ban/Allow"
Cmdgaugeshowstyle.Caption = "Display style"
Command6.Caption = "reverse current"
Command7.Caption = "resistance measurement"
LabNote.Caption = "Note: the two instruction delivery interval should be larger than the display refresh rate intervals, such as the display refresh rate of 6 times per second, command interval time should be more than 0.18 seconds, two instruction intervals greater than 0.2 seconds."
OptionALL.Caption = "Control all instruments available"
OptionOne.Caption = "Control single  Instrument                                       No.    :"
LabTitle.Caption = "Instructions         InstructionS                Set range                Instruction"
FrmInstruct.Caption = "Instruction List (YY: instrument No., ←┘: carriage return)"

FrmDataShow.Caption = "Receive data analysis"
FrmDataShow.Label1(0).Caption = "NO."
FrmDataShow.Label1(1).Caption = "Hex"
FrmDataShow.Label1(2).Caption = "ASCII"
FrmDataShow.Option1(0).Caption = "Binary"
FrmDataShow.Option1(1).Caption = "Graphical"
If FrmDataShow.CmdOK(0).Caption = "暂停" Then FrmDataShow.CmdOK(0).Caption = "Pause"
If FrmDataShow.CmdOK(0).Caption = "继续" Then FrmDataShow.CmdOK(0).Caption = "Continue"
FrmDataShow.CmdOK(1).Caption = "Exit"
FrmDataShow.FrmDataSaveStatus.Caption = "Data save - Closed"
FrmDataShow.Frame2(0).Caption = "Save style"
FrmDataShow.OptionSaveStyle(0).Caption = "Save by Hexadecimal"
FrmDataShow.OptionSaveStyle(1).Caption = "Save by ASCII"
FrmDataShow.OptionSaveStyle(2).Caption = "Save by Binary"
FrmDataShow.Frame2(1).Caption = "Save rule"
FrmDataShow.OptionSavegz(0).Caption = "Save One by one"
FrmDataShow.OptionSavegz(1).Caption = "Consecutive save"
FrmDataShow.CmdSave(0).Caption = "Start"
FrmDataShow.CmdSave(1).Caption = "Stop"
FrmDataShow.CmdSave(2).Caption = "Address..."
FrmDataShow.Label4(0).Caption = "Receive:"
FrmDataShow.Label4(1).Caption = "Effective:"
FrmDataShow.Label4(2).Caption = "Repeat:"
FrmDataShow.Command1.Caption = "Clear"
Lan = 1
End If
Call lanResize
End Sub

Private Sub ComBaudRate_Click()
   typSettings.Baud = ComBaudRate.Text
   UpdatePortSettings
End Sub

Private Sub ComBaudRate_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub CombCom_Click()
    UpdatePortSettings
End Sub

Private Sub CombCom_KeyPress(KeyAscii As Integer)
     KeyAscii = 0
End Sub

Private Sub ComComOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "打开已设置好的端口"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Open the COM has been well set"
End If
End Sub

Private Sub ComExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "退出程序"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Exit"
End If
End Sub

Private Sub ComRange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "通道建  主要功能：查看或更换数据通道"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Channel Function: To view or change data channels"
End If
End Sub

Private Sub ComDigits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "单位建  主要功能：进行单位变换，显示不同单位制下的测量结果数据"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Units features: flat exchange, display the data under different unit system results"
End If
End Sub

Private Sub ComRun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   If ShpOn(3).Visible = False Then
      LabInfo.Caption = "关闭键  主要功能：在常规测量状态下用于关闭仪表显示，降低电能消耗"
   Else
      LabInfo.Caption = "清零键  主要功能：在峰值测量状态下用于清除所保持的峰值数据"
   End If
ElseIf Lan = 1 Then
   If ShpOn(3).Visible = False Then
      LabInfo.Caption = "Close Function: Measurement instruments in the show closed, reducing power consumption"
   Else
      LabInfo.Caption = "Clear Function: Measured in peak condition for the peak data maintained by removal"
   End If
End If
End Sub

Private Sub ComShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "峰值键  主要功能：开启或关闭峰值测量状态，峰值测量状态时<峰值>指示灯亮"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Peak  Function: On or off peak measurement, peak state <peak> indicator light"
End If
End Sub

Private Sub ComPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "打印键  主要功能：打印测量数据"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Prinf Function: Print measurements"
End If
End Sub

Private Sub ComReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "复位建  主要功能：重新复位启动仪表，用于参数设定、查看、力值标定等操作"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Reset  Function: Re-start the meter reset for the parameter settings, view, force the value of calibration and other operations"
End If
End Sub

Private Sub ComZero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "置零键  主要功能：使数据显示器置零或还原，置零状态使<置零>指示灯点亮"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Zero  Function: Zero or restore data, zero state <Zero> indicator light"
End If
End Sub


Private Sub CmdShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "停止或继续显示接收区中的内容"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Stop or continue to show the contents of the receiving area"
End If
End Sub

Private Sub CmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "清空接收区中的内容"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Empty the contents of the reception area"
End If
End Sub

Private Sub AutoClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "选中则自动清空接收区中的内容"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Selected automatically empty the contents of the receiving area"
End If
End Sub

Private Sub ComReset_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;07" & vbCr)
   TimerCheckStart
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
   TimReset.Enabled = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComPrint_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;06" & vbCr)
   TimerCheckStart
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
   TimReset.Enabled = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComRange_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;05" & vbCr)
   TimerCheckStart
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComDigits_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;04" & vbCr)
   TimerCheckStart
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComShow_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;03" & vbCr)
   TimerCheckStart
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComRun_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;02" & vbCr)
   TimerCheckStart
If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComZero_Click()
On Error GoTo ErrHndl

MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;01" & vbCr)
   TimerCheckStart

If ReceiveType = "指令" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComComOpen_Click()
On Error GoTo ErrHndl
    If MSComPort.PortOpen = False Then
        MSComPort.PortOpen = True
        CombCom.Enabled = False
        ComBaudRate.Enabled = False
        ShpComOn.Visible = True
        TimerReceive.Enabled = True
        TimDetection.Enabled = True
        LabErr.Caption = ""
        If Lan = 0 Then
        MenOpenCom.Caption = "关闭端口" & "(&O)"
        ComComOpen.Caption = "关闭端口"
        FrmMain.Caption = "6000精密数字测量仪(V13-2.0718)"
        ShowType = "检测模式"
        ReceiveType = "指令"
        ElseIf Lan = 1 Then
        ComComOpen.Caption = "Close Port"
        MenOpenCom.Caption = "Close Port" & "(&O)"
        ShowType = "Detect"
        ReceiveType = "Instructions"
      End If
      MSComPort.Output = Trim("%00;03" & vbCr)
      RT = 1
      TimInstroction = True
      With MSComPort
             .InBufferCount = 0
             .OutBufferCount = 0
      End With
      ReceiveCounts = 0    '接收帧数
      ReceiveTrueCounts = 0 '接收有效帧数
    Else
    TimRecover.Enabled = False
    If TimInstroction.Enabled = True Then TimInstroction.Enabled = False
        MSComPort.PortOpen = False
        FrmMain.Caption = "6000精密数字测量仪(V13-2.0718)"
        LabErr.Caption = ""
        TextShow.Text = ""
        ShpComOn.Visible = False
        ShpRevOn.Visible = False
        For i% = 0 To 5
            If ShpOn(i).Visible Then ShpOn(i).Visible = False
        Next i
        ShpErrOn.Visible = False

        If Lan = 0 Then
          ComComOpen.Caption = "打开端口"
          MenOpenCom.Caption = "打开端口" & "(&O)"
        ElseIf Lan = 1 Then
          ComComOpen.Caption = "Open Port"
          MenOpenCom.Caption = "Open Port" & "(&O)"
        End If

        CombCom.Enabled = True
        ComBaudRate.Enabled = True
        TimerReceive.Enabled = False
        TimDetection.Enabled = False
        LabErr.Caption = ""
        With MSComPort
             .InBufferCount = 0
             .OutBufferCount = 0
        End With
    End If

ShpComOn.Visible = CBool(IIf(MSComPort.PortOpen = True, 1, 0))

ErrHndl:
    Select Case Err.Number
    Case comPortAlreadyOpen
     If Lan = 0 Then
        LabErr.Caption = "提示：串口可能已被占用，请重新检查设置"
        CloseMSC

     ElseIf Lan = 1 Then
        LabErr.Caption = "Serial may have been occupied"
        CloseMSC
     End If
    End Select
    Exit Sub
End Sub

Private Sub CloseMSC()
If MSComPort.PortOpen = True Then
   MSComPort.PortOpen = False
End If
End Sub

Private Sub ComExit_Click()
   MenExit_Click
End Sub

Private Sub SaveType()
    If OpenFileName = "" Then
        Exit Sub
    End If
    
    RecordNumber = RecordNumber + 1
    Open OpenFileName For Append As #1
    If RecordNumber = 1 Then
        Print #1, "分别显示的是：编号、原始数据、仪表显示数据、日期、时间"
    End If
    If FrmDataShow.OptionSaveStyle(0).Value = True Then
        Print #1, RecordNumber; Spc(1); HexDatas; Spc(1); ReceiveShow; Spc(1); Format(Date, "YYYY.MM.DD"); Spc(1); Time
    ElseIf FrmDataShow.OptionSaveStyle(1).Value = True Then
        Print #1, RecordNumber; Spc(1); ASCIIDatas; Spc(1); ReceiveShow; Spc(1); Format(Date, "YYYY.MM.DD"); Spc(1); Time
    ElseIf FrmDataShow.OptionSaveStyle(2).Value = True Then
        Print #1, RecordNumber; Spc(1); BinDatas; Spc(1); ReceiveShow; Spc(1); Format(Date, "YYYY.MM.DD"); Spc(1); Time
    End If
    Close #1
    If DatasSaveStyle = 2 Then
        If Lan = 0 Then
        FrmDataShow.FrmDataSaveStatus.Caption = "数据记录 - " & "当前为逐条记录数据形式，已记录" & RecordNumber & "条数据"
        Else
        FrmDataShow.FrmDataSaveStatus.Caption = "Data save - " & "Save one by one" & RecordNumber
        End If
        DatasSaveStyle = 1
    ElseIf DatasSaveStyle = 3 Then
        If Lan = 0 Then
        FrmDataShow.FrmDataSaveStatus.Caption = "数据记录 - " & "当前为连续记录数据形式，已记录" & RecordNumber & "条数据"
        Else
        FrmDataShow.FrmDataSaveStatus.Caption = "Data save - " & "Consecutive save" & RecordNumber
        End If
    End If
    Exit Sub
End Sub

Private Sub Form_Initialize()
Dim sPortSet As String
Dim iComma As Long
Dim iC As Long
On Error GoTo ErrHndl
    sPortSet = MSComPort.Settings
'
    iComma = InStr(1, sPortSet, ",", vbBinaryCompare)
    With typSettings
        .Baud = Mid$(sPortSet, 1, iComma - 1)
        ComBaudRate.Text = .Baud
    End With
    SearchCOM
    
    ReceiveFlag = True

    ISDataShow = False
    StringFlag = False
    Picture1.Enabled = False
    If Lan = 0 Then
    FrmMain.Caption = "6000表通讯检测程序(V13-2.0718)"
    ElseIf Lan = 1 Then
    FrmMain.Caption = "6000 Digital Meter(V13-2.0718)"
    End If
    YY = "00"
    Languages
    FrmSendShow = False
    Call ButtonInstruct
    Call SendDateStyle
    Call ButtonCOMr30
    Call ShowTray
    Call CmdgaugeshowstyleShow
    Call CmdfanzhuandianliuShow
    Call Cmd4xian2xianShow
    With MSComPort
       .InputLen = 37
       .RThreshold = 37
       .InBufferCount = 0
       .OutBufferCount = 0
    End With
    '-----------
ErrHndl:
    If Err.Number = 53 Then
    End If
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
Call lanResize
End Sub

Private Sub SearchCOM()
Dim hKey As Long, rName As String, rData As String
Dim zuixiao As Integer
RegOpenKeyEx HKEY_LOCAL_MACHINE, "HARDWARE\DEVICEMAP\SERIALCOMM", 0, KEY_ALL_ACCESS, hKey
If hKey <> 0 Then
    Dim Cnt As Long, rType As Long
    Cnt = 0
    rName = Space(20): rData = Space(20)
    While RegEnumValue(hKey, Cnt, rName, 20, 0, rType, ByVal rData, 20) <> 259&
        Cnt = Cnt + 1
        CombCom.AddItem Mid(Left(rData, InStr(rData, Chr(0)) - 1), 4, Len(Left(rData, InStr(rData, Chr(0)) - 1)))
    Wend
    If CombCom.ListCount > 1 Then
        zuixiao = CombCom.List(0)
        For i = 0 To CombCom.ListCount - 2
            If CombCom.List(i) > CombCom.List(i + 1) Then
                zuixiao = CombCom.List(i + 1)
            End If
        Next
        CombCom.Text = zuixiao
    Else
    CombCom.Text = CombCom.List(0)
    End If
    RegCloseKey hKey
End If
End Sub

Private Sub Languages()
If LogoLan = 1 Then
English_Click
ElseIf LogoLan = 0 Then
Chinese_Click
End If
End Sub

Private Sub mnuShow_Click()
If Status = STA_MIN Then
    Status = STA_NORMAL
    Me.Visible = True
    Me.WindowState = vbNormal
Else
    Status = STA_MIN
    Me.WindowState = vbMinimized
    Me.Visible = False
End If
End Sub

Private Sub ShowTray()
OrgWinRet = GetWindowLong(Me.hWnd, GWL_WNDPROC)
With myData
    .cbSize = Len(myData)
    .hWnd = Me.hWnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_INFO Or NIF_MESSAGE
    .uCallbackMessage = TRAY_CALLBACK
    .hIcon = Me.Icon
    .szTip = "6000精密数字测量仪" & vbCr & "测试软件" & vbCr & "V1.0.0" & vbNullChar
    .dwState = 0
    .dwStateMask = 0
    If LogoLan = 0 Then
    .szInfoTitle = "欢迎使用" & vbNullChar
    ElseIf LogoLan = 1 Then
    .szInfoTitle = "Welcome" & vbNullChar
    End If
    If LogoLan = 0 Then
    .szInfo = "单击本图标将显示/隐藏主程序。" & vbNullChar
    ElseIf LogoLan = 1 Then
    .szInfo = "Click the icon to display/hide the program。" & vbNullChar
    End If
    .dwInfoFlags = NIIF_INFO
    .uTimeout = 100
End With
Shell_NotifyIcon NIM_ADD, myData
glWinRet = SetWindowLong(Me.hWnd, GWL_WNDPROC, AddressOf CallbackMsgs)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button And vbRightButton Then PopupMenu MenPopup
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   LabInfo.Caption = ""
End Sub


Private Sub MenAbout_Click()
   FrmAbout.Show vbModal
End Sub

Private Sub MenExit_Click()
   If MSComPort.PortOpen = True Then MSComPort.PortOpen = False
   If FrmSendShow = True Then
   Unload FrmSend
   Set FrmSend = Nothing
   End If
   If ISDataShow = True Then
   Unload FrmDataShow
   Set FrmDataShow = Nothing
   End If
   Set FrmMain = Nothing
   End
End Sub

Private Sub MenF_Click(Index As Integer)
   Select Case Index
      Case 0
         Call ComZero_Click
      Case 1
         Call ComRun_Click
      Case 2
         Call ComShow_Click
      Case 3
         Call ComDigits_Click
      Case 4
         Call ComRange_Click
      Case 5
         Call ComPrint_Click
   End Select
End Sub

Private Sub MenInstr_Click()
MsgBox IIf(Lan, "Thanks for using our pruducts. Please read the manual.", "感谢使用北京弘豪福安仪器有限公司产品，请阅读产品说明书。"), vbInformation
End Sub

Private Sub MenOpenCom_Click()
    ComComOpen_Click
End Sub

Private Sub MenSupport_Click()
    FrmSupport.Show vbModal
End Sub

Private Sub MSComPort_OnComm()
   With MSComPort
      Select Case .CommEvent
         Case comEvReceive
            ReceiveFlag = False
            DealWithData
            TimRecover.Enabled = False
         Case comEventFrame
         Debug.Print "错误"
            If RT = 1 Then
            TimInstroction.Enabled = False
            MSComPort.Output = Trim("%00;02" & vbCr)
            delay 500
            MSComPort.Output = Trim("%00;03" & vbCr)
            TimInstroction.Enabled = True
            End If
            Exit Sub
         Case Else
         TextShow.Text = ""
      End Select
   End With
   
End Sub

Private Sub DealWithData()

Dim iX As Long
Dim iY As Long
Dim StrElement As Byte
Dim LightElement(7) As Byte
Dim ZeroSwitch As Boolean
Dim iZero As Integer
Dim receiveData As String
On Error GoTo ErrHndl
BinDatas = ""
HexDatas = ""
ASCIIDatas = ""
If MSComPort.InBufferCount < 37 Then Exit Sub
   Counter = MSComPort.InBufferCount
   StringIn = MSComPort.Input
   iX = UBound(StringIn)

   For iY = 0 To iX
    
      StrElement = StringIn(iY)
      If ISDataShow Then
        FrmDataShow.LabHex(iY).Caption = IIf(Len(Hex(StrElement)) = 2, Hex(StrElement), "0" & Hex(StrElement))
        FrmDataShow.LabASC(iY).Caption = Chr(StrElement)
        If BinOrDeng = False Then
            ABinDatas = ""
            For i = 1 To 8
              ABinDatas = ABinDatas & "   " & Mid(DEC_to_BIN(CLng(StrElement), 8), i, 1)
            Next i
            FrmDataShow.LabBin(iY).Caption = ABinDatas
        Else
            If FrmDataShow.DengStyle0.Count = 296 And FrmDataShow.DengStyle1.Count = 296 Then
                For i = 1 To 8
                  If Mid(DEC_to_BIN(CLng(StrElement), 8), i, 1) = 1 Then
                    FrmDataShow.DengStyle1(iY * 8 + i - 1).Visible = True
                  Else
                    FrmDataShow.DengStyle1(iY * 8 + i - 1).Visible = False
                  End If
                Next i
            End If
        End If
              HexDatas = HexDatas & Hex(StrElement) & " "
              ASCIIDatas = ASCIIDatas & Chr(StrElement)
              For i = 1 To 8
                  BinDatas = BinDatas & Mid(DEC_to_BIN(CLng(Str(StrElement)), 8), i, 1)
              Next i
              BinDatas = BinDatas & " "
      End If
      
      If StringFlag = False Then
          If StrElement = 255 Then
            StringFlag = True
            ReceiveCounts = ReceiveCounts + 1
            FrmDataShow.Label5(0).Caption = ReceiveCounts
          End If
      End If
      If iY = 21 Then
        If Mid(DEC_to_BIN(CLng(StrElement), 8), 4, 1) = 0 Then
            ReceiveTrueCounts = ReceiveTrueCounts + 1
            FrmDataShow.Label5(1).Caption = ReceiveTrueCounts
            FrmDataShow.Label5(2).Caption = ReceiveCounts - ReceiveTrueCounts
        End If
      End If
      If StringFlag = True Then
         If StrElement = 255 Then
         iZ = 0
         LabErr.Caption = ""
         End If
         StringTemp(iZ) = StrElement
         iZ = iZ + 1
      End If
      If iZ = 36 Then
         ShpRevOn.Visible = Not ShpRevOn.Visible
         ZeroSwitch = False

'------------------------------------------------read data----------------------------------------------------------
   For II = 0 To 36
    If ISDataShow = True Then
        
    End If
    NumArray(II) = Chr(StringTemp(II))
    receiveData = receiveData & NumArray(II)
   Next II
   JudgeData = receiveData
   StateByte5 = StringTemp(24)
   LightStatus = StringTemp(20)
   StateLight = StringTemp(21)
   Instrument = NumArray(27) & NumArray(28)
   Range = NumArray(30) & NumArray(31)
   Channel = NumArray(33) & NumArray(34)
   Median = StringTemp(23) And &HF&
   If NumArray(5) = "E" Then
      If NumArray(4) = "+" Then
         temp$ = NumArray(5) + NumArray(7)
      ElseIf NumArray(4) = "-" Then
         temp$ = NumArray(4) + NumArray(5) + NumArray(7)
      End If
   Else
         If NumArray(5) = "0" Then
            If NumArray(6) = "0" Then
               If NumArray(7) = "0" Then
                  If NumArray(8) = "." Then
                     For JJ = 7 To 5 + Median
                        temp$ = temp$ + NumArray(JJ)
                     Next JJ
                  Else
                     For JJ = 8 To 5 + Median
                        temp$ = temp$ + NumArray(JJ)
                     Next JJ
                  End If
               ElseIf NumArray(7) = "." Then
                  For JJ = 6 To 5 + Median
                     temp$ = temp$ + NumArray(JJ)
                  Next JJ
               Else
                  For JJ = 7 To 5 + Median
                     temp$ = temp$ + NumArray(JJ)
                  Next JJ
               End If
            ElseIf NumArray(6) = "." Then
               For JJ = 5 To 5 + Median
                  temp$ = temp$ + NumArray(JJ)
               Next JJ
            Else
               For JJ = 6 To 5 + Median
                  temp$ = temp$ + NumArray(JJ)
               Next JJ
            End If
         Else
            For JJ = 5 To 5 + Median
               temp$ = temp$ + NumArray(JJ)
            Next JJ
         End If
      If NumArray(4) = "-" Then
           temp$ = "-" + temp$
      End If
      MSComPort.InBufferCount = 0
   End If
   TextShow.Text = temp
   If Lan = 0 Then
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "00" Then
       Xianshileixing = "数显01"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "01" Then
       Xianshileixing = "数显02"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "10" Then
       Xianshileixing = "数显03"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 5, 1) = "0" Then
       Anjiancaozuo = "键允许"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 5, 1) = "1" Then
       Anjiancaozuo = "键禁止"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 4, 1) = "0" Then
       Fanzhuandianliu = "+"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 4, 1) = "1" Then
       Fanzhuandianliu = "-"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "0" Then
       Dianzu4or2 = "Ω4"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "1" Then
       Dianzu4or2 = "Ω2"
    End If
    TxtShowStyleInfo = Xianshileixing & " " & Anjiancaozuo & " " & Dianzu4or2 & Fanzhuandianliu
   ElseIf Lan = 1 Then
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "00" Then
       Xianshileixing = "Data01"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "01" Then
       Xianshileixing = "Data02"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "10" Then
       Xianshileixing = "Data03"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 5, 1) = "0" Then
       Anjiancaozuo = "Key+"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 5, 1) = "1" Then
       Anjiancaozuo = "Key-"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 4, 1) = "0" Then
       Fanzhuandianliu = "+"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 4, 1) = "1" Then
       Fanzhuandianliu = "-"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "0" Then
       Dianzu4or2 = "Ω4"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "1" Then
       Dianzu4or2 = "Ω2"
    End If
    TxtShowStyleInfo = Xianshileixing & " " & Anjiancaozuo & " " & Dianzu4or2 & Fanzhuandianliu
   End If
   
   If RT = 1 Then TimInstroction = True
   If Lan = 0 Then
   TxtShowInfo.Text = "br42=" & Instrument & "  Lr32=" & Range & "  CH=" & Channel & "  " & ReceiveType & "  " & TxtShowStyleInfo
   ElseIf Lan = 1 Then
   TxtShowInfo.Text = "br42=" & Instrument & " Lr32=" & Range & " CH=" & Channel & " " & ReceiveType & " " & TxtShowStyleInfo
   End If
   
   LightStatus = StringTemp(20)
   StateLight = StringTemp(21)
   Instrument = NumArray(27) & NumArray(28)
   Range = NumArray(30) & NumArray(31)
   Channel = NumArray(33) & NumArray(34)
   Median = StringTemp(23) And &HF&
'---------------------------------------------Lights Status----------------------------------------------------------

   LightElement(0) = (LightStatus And &H1&) And &H1&
   LightElement(1) = (LightStatus And &H2&) And &H2&
   LightElement(2) = (LightStatus And &H4&) And &H4&
   LightElement(3) = (LightStatus And &H8&) And &H8&
   LightElement(4) = (LightStatus And &H10&) And &H10&
   LightElement(5) = (LightStatus And &H20&) And &H20&
   LightElement(6) = (StateLight And &H1&) And &H1&
   LightElement(7) = (StateLight And &H2&) And &H2&
   For i = 0 To 7
      If LightElement(i) = 0 Then
         ShpOn(i).Visible = False
      Else
         ShpOn(i).Visible = True
      End If
   Next i
'-------------------------------------------------end----------------------------------------------------------------
         
         StringFlag = False
         
         iZ = 0
      End If
   Next iY
          ReceiveShow = temp
          If DatasSaveStyle = 2 Then
            SaveType
          ElseIf DatasSaveStyle = 3 Then
            SaveType
          End If

ErrHndl:
    Select Case Err.Number
        Case comPortNotOpen
            MsgBox "The COM Is Not Open. Open The COM And Then Retry.", vbOKOnly, Err.Description
        Case Else
            Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
    End Select
End Sub


Private Function TimerCheckStart()
   ShpSendOn.Visible = True
   TimerCheck.Enabled = True
End Function

Private Sub OptionALL_Click()
If Lan = 0 Then
ShowType = "调试模式"
ElseIf Lan = 1 Then
ShowType = "Debug Mode"
End If
Picture1.Enabled = True
CombInstrumentNo.Enabled = False
YY = "00"
End Sub

Private Sub OptionALL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "选中此项，则下列命令控制当前串口下连接的所有仪表。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Select the command and control, under the current serial connecting all meter."
End If
End Sub

Private Sub OptionOne_Click()
Picture1.Enabled = True
CombInstrumentNo.Enabled = True
YY = Trim(CombInstrumentNo.Text)
End Sub

Private Sub OptionOne_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "选中此项，则下列命令只控制所选定的仪表识别号的仪表，当前串口的其他仪表不受影响。"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Select the following command and control of the selected only the instrument meter recognition, the serial other instrument is not affected."
End If
End Sub



Private Sub TimDetection_Timer()
If JudgeData <> "" Then
   JudgeData = ""
Else
   TextShow.Text = "···················"
   If Lan = 0 Then
       TxtShowInfo = "当前接收模式：" & ReceiveType
   ElseIf Lan = 1 Then
      TxtShowInfo = "Receive mode:" & ReceiveType
   End If
   For i% = 0 To 7
       If ShpOn(i).Visible Then ShpOn(i).Visible = False
   Next i
   ShpRevOn.Visible = False
   If ReceiveType = "连续" Or ReceiveType = "Continuous" Then
       MSComPort.Output = Trim("%00;02" & vbCr)
   End If
End If
End Sub

Private Sub TimerCheck_Timer()
   If ShpSendOn.Visible = True Then
      ShpSendOn.Visible = False
      TimerCheck.Enabled = False
   End If
End Sub

Private Sub TimerErr_Timer()
   TimerErr.Enabled = False
   ShpErrOn.Visible = False
End Sub

Private Sub TimerFre_Timer()
   If ShpRevOn.Visible = True Then
      ShpRevOn.Visible = False
      TimerFre.Enabled = False
   End If
End Sub

Private Sub TimerReceive_Timer()
If Lan = 0 Then
   If ReceiveFlag Then LabErr.Caption = "提示：没有数据接收，请检查是否已开启数据输出或端口设置是否有效"
ElseIf Lan = 1 Then
   If ReceiveFlag Then LabErr.Caption = "Error: No data is received, check the COM settings"
End If
   TimerReceive.Enabled = False

End Sub


'---------------------------------------------------------START---------------------------------------------------------
Private Sub UpdatePortSettings()
    Dim sPortSet As String

On Error GoTo ErrHndl

        With typSettings
            sPortSet = .Baud & ","
            sPortSet = sPortSet & "n,8,1"
        End With
        If MSComPort.PortOpen Then
            MSComPort.PortOpen = False
            DoEvents
            MSComPort.Settings = sPortSet
            MSComPort.CommPort = CombCom.Text
            MSComPort.PortOpen = True
        Else
            MSComPort.Settings = sPortSet
            MSComPort.CommPort = CombCom.Text
        End If
    Exit Sub
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub SendDateStyle()
    Set tooltip = New cToolTip
    With tooltip
        .Create CmdSendExtremum, " 指令内容                     显示数据类型          " & vbCr & "[%YY;21;04←┘]               瞬态最大值数据" & vbCr & "[%YY;21;05←┘]               瞬态最小值数据" & vbCr & "[%YY;21;06←┘]               瞬态最大值与最小值的差值数据" & vbCr & "[%YY;21;07←┘]               平均最大值数据" & vbCr & "[%YY;21;08←┘]               平均最小值数据" & vbCr & "[%YY;21;09←┘]               平均最大值与最小值的差值数据" & vbCr & "[%YY;21;10←┘]               显示平均值数据(由极致状态切换至平均值状态)" & vbCr & vbCr & "[%YY;21;01←┘]               显示绝对平均值数据(在平均值状态下有效)" & vbCr & "[%YY;21;02←┘]               现实相对平均值数据(在平均值状态下有效)" & vbCr & "[%YY;21;03←┘]               无效", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "COMr21 发送极值测量数据类型选择，[%YY;21;KK←┘]，KK所对应的显示内容如下：", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "宋体", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, CmdSendExtremum.Name
End Sub

Private Sub ButtonInstruct()
    Set tooltip = New cToolTip
    With tooltip
        .Create CmdKeyFunction, "    指令内容            说明" & vbCr & "[%YY;17;01←┘]   等效《置零》按键 " & vbCr & "[%YY;17;02←┘]   等效《运行》按键 " & vbCr & "[%YY;17;03←┘]   等效《显示》按键 " & vbCr & "[%YY;17;04←┘]   等效《位数》按键 " & vbCr & "[%YY;17;05←┘]   等效《量程》按键 " & vbCr & "[%YY;17;06←┘]   等效《打印》按键 " & vbCr & "[%YY;17;07←┘]   等效《复位》按键 " & vbCr & "[%YY;17;08←┘]   等效《WK1》按键 " & vbCr & "[%YY;17;09←┘]   等效《WK2》按键 ", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "COMr17 串口按键指令，[%YY;17;KK←┘]，KK所对应的按键功能如下：", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "宋体", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, CmdKeyFunction.Name
End Sub

Private Sub ButtonCOMr30()
    Set tooltip = New cToolTip
    With tooltip
        .Create Command1, "    指令内容            说明" & vbCr & "[%YY;18;00←┘]   SP1和SP2灯熄灭，输出继电器断开 " & vbCr & "[%YY;18;01←┘]   SP1灯亮，SP1继电器闭合 " & vbCr & "[%YY;18;02←┘]   SP2灯亮，SP2继电器闭合 " & vbCr & "[%YY;18;03←┘]   SP1、SP2灯均亮，继电器均闭合", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "COMr18 串口按键指令，[%YY;18;KK←┘]，KK所对应的按键功能如下：", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "宋体", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Command1.Name
End Sub

Private Sub CmdfanzhuandianliuShow()
    Set tooltip = New cToolTip
    With tooltip
        .Create Command6, "（本指令适用于6000-11型电阻测量方式或特别指明的产品)" & vbCr & vbCr & "    指令内容            说明" & vbCr & "[%YY;35;00←┘]   设定电阻正向测量电流(常规方式) " & vbCr & "[%YY;35;01←┘]   设定电阻反向测量电流(反转方式) ", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "激励电流方向指令，[%YY;35;KK ]，KK取值范围是0~1", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "宋体", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Command6.Name
End Sub

Private Sub Cmd4xian2xianShow()
    Set tooltip = New cToolTip
    With tooltip
        .Create Command7, "     （本指令适用于6000-11型电阻测量方式或特别指明的产品)" & vbCr & vbCr & "    指令内容            说明" & vbCr & "[%YY;36;00←┘]   设定为4线电阻测量方式(开机默认方式) " & vbCr & "[%YY;36;01←┘]   设定为2线电阻测量方式 ", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "4线/2线电阻测量方式设定指令，[%YY;36;KK ]，KK取值范围是0~1", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "宋体", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Command7.Name
End Sub

Private Sub CmdgaugeshowstyleShow()
    Set tooltip = New cToolTip
    With tooltip
        .Create Cmdgaugeshowstyle, "    指令内容            说明" & vbCr & "[%YY;34;00←┘]   显示测量数据(正常显示方式)" & vbCr & "[%YY;34;01←┘]   显示仪表当前的量程识别号Lr32，〖LH=XX〗XX=Lr32" & vbCr & "[%YY;34;02←┘]   显示仪表通讯识别号br42，〖Ud=XX〗XX=br42", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "显示类型选择指令，[%YY;34;KK ]，KK取值范围是0~2", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "宋体", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Cmdgaugeshowstyle.Name
End Sub

Private Sub TimInstroction_Timer()
   MSComPort.Output = Trim("%00;01" & vbCr)
   TimerCheckStart
End Sub


Private Sub TimReset_Timer()
If JudgeData <> "" Then
   ResetCount = ResetCount + 1
   JudgeData = ""
   If ResetCount = 2 Then
   TimReset.Enabled = False
   ResetCount = 0
   End If
Else
  If MSComPort.PortOpen = True Then
   MSComPort.Output = Trim("%00;02" & vbCr)
  End If
End If
End Sub
'
Private Sub VScroll1_Change()
PicOrder.Top = -1 * VScroll1.Value
End Sub

Private Sub lanResize()
LabelOnTop(0).Left = Frame1.Left + (Frame1.Width / 2 - LabelOnTop(0).Width / 2)
End Sub

Private Sub Light1()
ShpRevOff.Visible = False
delay 5
ShpRevOn.Visible = True
delay 5
ShpRevOff.Visible = True
delay 5
ShpRevOn.Visible = False
End Sub

Private Sub VScroll1_Scroll()
PicOrder.Top = -1 * VScroll1.Value
End Sub
