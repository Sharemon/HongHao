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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton CmdDataShow 
      Caption         =   "���ݷ���"
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
      Caption         =   "�� ��"
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
      Caption         =   "����ģʽ"
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
      Caption         =   "��������"
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
      Caption         =   "�򿪶˿�"
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
      Caption         =   "����ָ��COMr���(������ʹ�ã�YY�����Ǳ�ͨѶʶ��ţ�������ʾ�س���)"
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
         Caption         =   "���Ƶ����Ǳ�                                               �Ǳ�ʶ��ţ�"
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
         Caption         =   "���Ƶ�ǰ���ڵ������Ǳ�"
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
               Caption         =   "4��/2�ߵ��������ʽ"
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
               Caption         =   "��ת����(6000-11)��"
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
               Caption         =   "��������ʾ����"
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
               Caption         =   "��尴����ֹ/��Ч"
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
               Caption         =   "������˸��ʽѡ��"
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
               Caption         =   "��ʾ���ȵ���"
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
               Caption         =   "A/Dת����λ"
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
               Caption         =   "���̿���ѡ��"
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
               Caption         =   "�ָ��趨�˲�ʱ��"
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
               Caption         =   "�˲�ʱ��ӳ�"
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
               Caption         =   "�˲�ʱ������"
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
               Caption         =   "�˲�����λ"
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
               Caption         =   "���������λ(8λ)"
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
               Caption         =   "�������ʾλ����ͬ"
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
               Caption         =   "���ͼ�ֵ��������"
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
               Caption         =   "���Ͳ�����������"
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
               Caption         =   "���ڰ�������"
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
               Caption         =   "�ź�����ͨ���л�"
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
               Caption         =   "�л���������λ"
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
               Caption         =   "�л����û���λ*"
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
               Caption         =   "����ѡ��"
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
               Caption         =   "λ��(�ֱ���)�л�"
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
               Caption         =   "λ��(�ֱ���)ѡ�� "
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
               Caption         =   "��������״̬��λ"
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
               Caption         =   "������ֵ����"
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
               Caption         =   "�þ�����"
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
               Caption         =   "�������"
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
               Caption         =   "�Ǳ���ʾ��ʱϨ��"
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
               Caption         =   "�����������ѡ��"
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
               Caption         =   "�������ݵ������"
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
               Caption         =   "����ͨѶ��ʼ��"
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
               Caption         =   "ֹͣ�����������"
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
               Caption         =   "���������������"
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
               Caption         =   "SP�������"
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
               Caption         =   "%YY;36;KK����"
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
               Caption         =   "%YY;35;KK����"
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
               Caption         =   "%YY;34;KK����"
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
               Caption         =   "%YY;33;KK����"
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
               Caption         =   "%YY;32;KK����"
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
               Caption         =   "%YY;31;KK����"
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
               Caption         =   "%YY;30����"
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
               Caption         =   "%YY;29����"
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
               Caption         =   "%YY;28����"
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
               Caption         =   "%YY;27����"
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
               Caption         =   "%YY;26����"
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
               Caption         =   "%YY;25����"
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
               Caption         =   "%YY;24����"
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
               Caption         =   "%YY;23����"
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
               Caption         =   "%YY;22����"
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
               Caption         =   "%YY;21;KK����"
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
               Caption         =   "%YY;20;KK����"
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
               Caption         =   "%YY;19����"
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
               Caption         =   "%YY;18;KK����"
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
               Caption         =   "%YY;17;KK����"
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
               Caption         =   "%YY;16;KK����"
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
               Caption         =   "%YY;15����"
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
               Caption         =   "%YY;14;KK����"
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
               Caption         =   "%YY;13����"
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
               Caption         =   "%YY;12;KK����"
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
               Caption         =   "%YY;11����"
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
               Caption         =   "%YY;10;KK����"
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
               Caption         =   "%YY;09����"
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
               Caption         =   "%YY;08����"
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
               Caption         =   "%YY;07����"
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
               Caption         =   "%YY;06����"
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
               Caption         =   "%YY;05;KK����"
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
               Caption         =   "%YY;04;KK����"
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
               Caption         =   "%YY;03����"
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
               Caption         =   "%YY;02����"
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
               Caption         =   "%YY;01����"
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
               Caption         =   "%YY;00����"
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
         Caption         =   "ָ����nn         ָ������                    �趨��Χ             ASCIIָ���ʽ     "
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
      Caption         =   "��������Ϣ"
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
         Caption         =   "����ָ������"
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
         Caption         =   "��ȡ״̬"
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
         Caption         =   "����״̬"
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
         Caption         =   "����ָʾ"
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
         Caption         =   "��ǰ���ڣ�"
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
         Caption         =   "��ǰ�����ʣ�"
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
         Caption         =   "Ĭ�����ã�   9600;N;8;1"
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
         Caption         =   "�˿�״̬"
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
      Caption         =   "����"
      DisabledPicture =   "FrmMaim.frx":08A7
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ʾ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "λ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ӡ"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��λ"
      BeginProperty Font 
         Name            =   "����"
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
      DialogTitle     =   "�����ļ�Ϊ"
      Filter          =   "�ı��ļ�|*.txt"
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
      Caption         =   "����"
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
      Caption         =   "����"
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
      Caption         =   "6000�������ֲ�����"
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
      Caption         =   "��"
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
      Caption         =   "K��"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu MenOpenCom 
         Caption         =   "�򿪶˿�(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu MenBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MenExit 
         Caption         =   "�˳�(&Q)"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MenFun 
      Caption         =   "����(&U)"
      Begin VB.Menu MenF 
         Caption         =   "����"
         Index           =   0
         Shortcut        =   {F1}
      End
      Begin VB.Menu MenF 
         Caption         =   "����"
         Index           =   1
         Shortcut        =   {F2}
      End
      Begin VB.Menu MenF 
         Caption         =   "��ʾ"
         Index           =   2
         Shortcut        =   {F3}
      End
      Begin VB.Menu MenF 
         Caption         =   "λ��"
         Index           =   3
         Shortcut        =   {F4}
      End
      Begin VB.Menu MenF 
         Caption         =   "����"
         Index           =   4
         Shortcut        =   {F5}
      End
      Begin VB.Menu MenF 
         Caption         =   "��ӡ"
         Index           =   5
         Shortcut        =   {F6}
      End
      Begin VB.Menu MenBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MenRes 
         Caption         =   "��λ"
         Shortcut        =   {F7}
      End
      Begin VB.Menu MenBar3 
         Caption         =   "-"
      End
      Begin VB.Menu MenShowData 
         Caption         =   "�鿴����"
         Shortcut        =   {F8}
      End
      Begin VB.Menu MenToExcel 
         Caption         =   "����¼���ݵ���Excel"
         Shortcut        =   {F9}
      End
      Begin VB.Menu MenBar10 
         Caption         =   "-"
      End
      Begin VB.Menu MenSend 
         Caption         =   "��������"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu Language 
      Caption         =   "����(&L)"
      Begin VB.Menu Chinese 
         Caption         =   "����(&C)"
         Shortcut        =   ^C
      End
      Begin VB.Menu English 
         Caption         =   "English(&E)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MenHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu MenInstr 
         Caption         =   "�Ǳ�ʹ��˵��(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu MenBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MenAbout 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu MenPopup 
      Caption         =   "�Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu MenMode 
         Caption         =   "����ģʽ"
      End
      Begin VB.Menu MenBar4 
         Caption         =   "-"
      End
      Begin VB.Menu MenDataShow 
         Caption         =   "�鿴����"
      End
      Begin VB.Menu MenBar5 
         Caption         =   "-"
      End
      Begin VB.Menu MenSupport 
         Caption         =   "����֧��"
      End
   End
   Begin VB.Menu TrayMenu 
      Caption         =   "���̲˵�"
      Visible         =   0   'False
      Begin VB.Menu MenuShow 
         Caption         =   "��ʾ����"
      End
      Begin VB.Menu MenBar6 
         Caption         =   "-"
      End
      Begin VB.Menu MenuAbout 
         Caption         =   "��������"
      End
      Begin VB.Menu MenBar7 
         Caption         =   "-"
      End
      Begin VB.Menu MenuExit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu ReceiveMenu 
      Caption         =   "�������Ҽ��˵�"
      Visible         =   0   'False
      Begin VB.Menu MenuChar 
         Caption         =   "���ַ���ʾ"
      End
      Begin VB.Menu MenuHex 
         Caption         =   "��ʮ��������ʾ"
      End
      Begin VB.Menu MenBar8 
         Caption         =   "-"
      End
      Begin VB.Menu MenuStop 
         Caption         =   "ֹͣ��ʾ"
      End
      Begin VB.Menu MenuReceive 
         Caption         =   "��ս�����"
      End
      Begin VB.Menu MenBar9 
         Caption         =   "-"
      End
      Begin VB.Menu MenuSend 
         Caption         =   "��շ����ַ���"
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
Dim BinDatas As String  '�������ݶ����Ʊ�ʾ
Dim HexDatas As String  '��������ʮ�����Ʊ�ʾ
Dim ASCIIDatas As String  '��������ASCII��ʾ
Dim ABinDatas As String 'һ���ֽڵĶ�����
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

Dim COMS() As String '���ں�����

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
If TxtShowInfo.Text = "Port Status��Closed" Then TxtShowInfo.Text = "����״̬���ر�"
LabelOnTop(0).Caption = "6000�������ֲ�����"
LabelOnTop(1).Caption = "����"
LabelOnTop(2).Caption = "����"
LabelOnTop(3).Caption = "����ָ������"
FrmMain.Caption = "6000��ͨѶ������(V13-2.0718)"
ComZero.Caption = "����"
ComRun.Caption = "����"
ComShow.Caption = "��ʾ"
ComDigits.Caption = "λ��"
ComRange.Caption = "����"
ComPrint.Caption = " ��ӡ"
ComReset.Caption = "��λ"
ComComOpen.Caption = "�򿪶˿�"
CmdDataShow.Caption = "���ݷ���"
ComExit.Caption = "�� ��"
LabInfo.Caption = "׼�����"
MenFile.Caption = "�ļ�(&F)"
MenOpenCom.Caption = "�򿪶˿�(&O)"
MenExit.Caption = "�˳�&Q)"
MenFun.Caption = "����(&U)"
MenF.Item(0).Caption = "����"
MenF.Item(1).Caption = "����"
MenF.Item(2).Caption = "��ʾ"
MenF.Item(3).Caption = "λ��"
MenF.Item(4).Caption = "����"
MenF.Item(5).Caption = "��ӡ"
MenRes.Caption = "��λ"
MenToExcel.Caption = "����¼���ݵ���Excel"
MenShowData.Caption = "�鿴����"
MenSend.Caption = "��������"
Language.Caption = "����(&L)"
MenHelp.Caption = "����(&H)"
MenInstr.Caption = "�Ǳ�ʹ��˵��(&R)"
MenAbout.Caption = "����(&A)"
FrmInf.Caption = "��������Ϣ"
LabelOnTop(10).Caption = "�˿�״̬"
LabelOnTop(11).Caption = "��ȡ״̬"
LabelOnTop(12).Caption = "����״̬"
LabelOnTop(13).Caption = "����ָʾ"
LabelOnTop(14).Caption = "��ǰ���ڣ�"
LabelOnTop(15).Caption = "��ǰ�����ʣ�"
LabelOnTop(16).Caption = "Ĭ�����ã�   9600;N;8;1"
If CmdMode.Caption = "Detection mode" Then
   CmdMode.Caption = "���ģʽ"
ElseIf CmdMode.Caption = "Debug mode" Then
   CmdMode.Caption = "����ģʽ"
End If
If CmdReceive.Caption = "Receive by Command" Then
   CmdReceive.Caption = "ָ�����"
ElseIf CmdReceive.Caption = "Receive Continuously" Then
   CmdReceive.Caption = "��������"
End If
If ReceiveType = "Instructions" Then
ReceiveType = "ָ��"
ElseIf ReceiveType = "Continuous" Then
ReceiveType = "����"
End If
If ShowType = "Debug" Then
ShowType = "����ģʽ"
ElseIf ShowType = "Detect" Then
ShowType = "���ģʽ"
End If
MenuShow.Caption = "��ʾ����"
MenuAbout.Caption = "����"
MenuExit.Caption = "�˳�"
CmdInitialize.Caption = "����ͨѶ��ʼ��"
CmdOnce.Caption = "�������ݵ������"
CmdContinuation.Caption = "���������������"
CmdStopContinuation.Caption = "ֹͣ�����������"
CmdOutputType.Caption = "�����������ѡ��"
CmdQuench.Caption = "�Ǳ���ʾ��ʱϨ��"
CmdRelative.Caption = "�������"
CmdAbsolute.Caption = "�þ�����"
CmdEliminate.Caption = "������ֵ����"
CmdReset.Caption = "��������״̬��λ"
CmdDigit.Caption = "λ��(�ֱ���)ѡ��"
CmdCut.Caption = "λ��(�ֱ���)�л�"
CmdRange.Caption = "����ѡ��"
CmdRangeFast.Caption = "���̿���ѡ��"
CmdCutUsers.Caption = "�л����û���λ"
CmdCutBasic.Caption = "�л���������λ"
CmdGallery.Caption = "�ź�����ͨ���л�"
CmdKeyFunction.Caption = "���ڰ�������"
CmdSendMeter.Caption = "���Ͳ�����������"
CmdSendExtremum.Caption = "���ͼ�ֵ��������"
CmdAlike.Caption = "�������ʾλ����ͬ"
CmdFull.Caption = "���������λ(8λ)"
CmdFilterReset.Caption = "�˲�����λ"
CmdShorten.Caption = "�˲���ʱ������"
Cmdlengthen.Caption = "�˲���ʱ��ӳ�"
CmdRecover.Caption = "�ָ��趨�˲�ʱ��"
Cmdlight.Caption = "��ʾ���ȵ���"
Command1.Caption = "SP�������"
Command5.Caption = "A/Dת����λ"
Comflicker.Caption = "������˸��ʽѡ��"
Cmdpanelnoyes.Caption = "��尴����ֹ/��Ч"
Cmdgaugeshowstyle.Caption = "��������ʾ����"
Command6.Caption = "��ת����(6000-11)��"
Command7.Caption = "4��/2�ߵ��������ʽ"
LabNote.Caption = "ע��:���η���ָ��ļ��ʱ��Ӧ������ʾˢ���ʵļ��ʱ�䣬����ʾˢ����6��/�룬ָ����ʱ��Ӧ����0.18�룬��������ָ����ʱ�����0.2�롣"
OptionALL.Caption = "���Ƶ�ǰ���ڵ������Ǳ�"
OptionOne.Caption = "���Ƶ����Ǳ�                                              �Ǳ�ʶ��ţ�"
LabTitle.Caption = "ָ����nn         ָ������                    �趨��Χ             ASCIIָ���ʽ     "
FrmInstruct.Caption = "����ָ��COMr���(������ʹ�ã�YY�����Ǳ�ͨѶʶ��ţ�������ʾ�س���)"

FrmDataShow.Caption = "�������ݷ���"
FrmDataShow.Label1(0).Caption = "�ֽں�"
FrmDataShow.Label1(1).Caption = "ʮ������"
FrmDataShow.Label1(2).Caption = "ASCII"
FrmDataShow.Option1(0).Caption = "������"
FrmDataShow.Option1(1).Caption = "ͼ����ʾ"
If FrmDataShow.CmdOK(0).Caption = "Pause" Then FrmDataShow.CmdOK(0).Caption = "��ͣ"
If FrmDataShow.CmdOK(0).Caption = "Continue" Then FrmDataShow.CmdOK(0).Caption = "����"
FrmDataShow.CmdOK(1).Caption = "�ر�"
FrmDataShow.FrmDataSaveStatus.Caption = "���ݼ�¼ - ��¼����δ����"
FrmDataShow.Frame2(0).Caption = "��¼��ʽѡ��"
FrmDataShow.OptionSaveStyle(0).Caption = "ʮ��������ʽ����"
FrmDataShow.OptionSaveStyle(1).Caption = "ASCII��ʽ����"
FrmDataShow.OptionSaveStyle(2).Caption = "��������ʽ����"
FrmDataShow.Frame2(1).Caption = "�������ѡ��"
FrmDataShow.OptionSavegz(0).Caption = "������¼���ݵ��ı���"
FrmDataShow.OptionSavegz(1).Caption = "������¼���ݵ��ı���"
FrmDataShow.CmdSave(0).Caption = "��ʼ��¼"
FrmDataShow.CmdSave(1).Caption = "ֹͣ��¼"
FrmDataShow.CmdSave(2).Caption = "ѡ��·��..."
FrmDataShow.Label4(0).Caption = "����֡����"
FrmDataShow.Label4(1).Caption = "��Ч֡����"
FrmDataShow.Label4(2).Caption = "�ظ�֡����"
FrmDataShow.Command1.Caption = "����"
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
   LabInfo.Caption = "�þ�����λ��[%YY;07����]���յ�ָ����Ǳ�����У׼ʱ�ľ�����λ���ݣ��Ǳ�˺���ʾ������Ϊ����ֵ���ݡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set absolute zero,[%YY;07����], After receiving instruction instrument will bring up the absolute zero calibration data, instrument data, then display the data for the absolute value."
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
   LabInfo.Caption = "�ô�������������Ǳ���ʾ����λ����ͬ��[%YY;22����]������ͨѶ������ӡ�����λ������ʾλ����ͬ��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Serial output data set and the same instrument display data bits, [%YY;22����], Serial communication output or printout of the median and display the same median."
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
   LabInfo.Caption = "�����������������[%YY;02����]�����ΪYY�Ǳ��յ���ָ����������Ͳ������ݣ��ڴ���ͨѶ�������㹻��ʱ�������������ˢ��������ʾˢ����ͬ������������;�ⲻ����ʹ�ñ���ʽ��������ͨѶ��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Start data continuous output, [%YY;02����], code instrumentation YY after receiving the order to send a continuous measurement data, except for special purposes, does not recommend using this approach to data communication"
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
   LabInfo.Caption = "��ʾλ��(�ֱ���)�����л���[%YY;11����]���յ�ָ����Ǳ���������һ��λ������ʾ�ֱ��ʡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Show the median (resolution) fast switching, [% YY;11����], after receiving instructions to adjust the instrument to the next digit display resolution."
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
   LabInfo.Caption = "�л�������������λ(mV��V������mV/V)��[%YY;15����]����COMr14ָ�����ƣ��������Ӧ�Ĺ���ģ�顣"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Switch to basic measurement unit (mV,V,��,mV/V), [%YY;15����], and COMr14 command similar to the work you need to add the corresponding module."
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
   LabInfo.Caption = "�л����û��趨��λ*��[%YY;14;KK����]���踽����Ӧ�Ĺ���ģ�飬�յ�ָ����Ǳ���ʾ�����л����û��趨�ĵ�λ*��KKΪ�û���λ��š�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Switch to the user to set the unit *, [%YY;14;KK����], need to add the appropriate modules of work, after receiving instructions to switch to the meter will display the data the user to set the unit *, KK is the user unit code."
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
   LabInfo.Caption = "��ʾλ��(�ֱ���)ѡ��[%YY;10;KK����]��KK=05-08Ϊ�趨Ҫ��ʾ��λ����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Show the median (resolution) to select,[%YY;10;KK����], KK = 05-08 for the set of bits to be displayed"
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
   LabInfo.Caption = "������ֵ���ݣ�[%YY;08����]���յ�ָ����Ǳ������ǰ�������ļ�ֵ���ݣ��Ա���к�������������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Elimination of extreme value data,[%YY;08����], after receiving instructions to clear the instrument previously reserved extreme data, to facilitate later measurements."
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
   LabInfo.Caption = "�˲�����λ��[%YY;25����]���Ǳ��ڲ��˲���Ӳ���������ʼ�����Ա���ٸ��ٱ仯�Ļ�Ծ�źţ����ָ�����˲�ʱ���趨�ϴ�ʱ����Ƚ����ԣ����˲�ʱ��С��0.5��ʱ�������ޣ�ִ��ָ�����ܻ���ֶ�ʱ������ݲ���ʱ��������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Filter reset, [%YY;25����], instrument hardware and software within the filter is initialized in order to fast track changes in the active signal."
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
   LabInfo.Caption = "�ô��ڷ�������Ϊ��Ϊ����(8λ)�����[%YY;23����]���Ǳ��յ�ָ���������λ��(8λ)��������ݡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set serial COM to send data over the data (8 bit) output, [%YY;23����], instrument instruction received the greatest number of bits (8 bits) output data."
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
   LabInfo.Caption = "�ź�����ͨ���л�(����Ӷ�ͨ������ת��ģ��)��[%YY;16����]���յ�ָ����Ǳ��л�����KK����ͨ����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Signal input channel switching (to add multi-channel input conversion module), [%YY;16����], after receiving instructions switching to KK instrument input channels."
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
   LabInfo.Caption = "��ʾ����ѡ��ָ�[%YY;34;KK ]��KKȡֵ��Χ��0~2����ָ���Ӱ���Ǳ���ʾ�����ݣ����ڴ�������ͨѶ�Ĵ����ָ����վ�������Ӱ�졣"
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
   LabInfo.Caption = "����ͨѶ��ʼ����[%YY;00����]����ָ��ʹ���������ֹ���ȴ��µ����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Serial initializing,[%YY;00����], this instruction to terminate, waiting for a new serial output command"
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
   LabInfo.Caption = "�˲����˲�ʱ��Ӵ�2����[%YY;27����]������ʹ�������ݵ��ȶ������ӡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Filter filter time increased 2 times, [%YY;27����], can increase the stability of the measurement data."
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
   LabInfo.Caption = "��������ʾ���ȵ�����[%YY;31;KK ]��KKȡֵ��Χ��1~6���յ���ָ����Ǳ����KKֵ������ʾ�������ȣ�KK=1ʱ��ʾ�������KK=0��KK=6ʱ��ʾ�������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Meter display brightness adjustment, [% YY; 31; KK, ?] KK range is 1 to 6, after receiving this instruction, instrument KK value adjustments on the brightness of the display, the display brightness is the darkest KK, = 1, KK, =0 or KK = 6, the display brightness."
End If
End Sub

Private Sub CmdMode_Click()
If CmdMode.Caption = "����ģʽ" Or CmdMode.Caption = "Debug mode" Then
    If Lan = 0 Then
      CmdMode.Caption = "���ģʽ"
      ShowType = "����ģʽ"
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
ElseIf CmdMode.Caption = "���ģʽ" Or CmdMode.Caption = "Detection mode" Then
    If Lan = 0 Then
       CmdMode.Caption = "����ģʽ"
       ShowType = "���ģʽ"
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
   LabInfo.Caption = "�򿪵���ģʽ(�رռ��ģʽ)���ߴ򿪼��ģʽ(�رյ���ģʽ)��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Open the debug mode (turn off detection mode) or open detection mode (turn off debug mode)."
End If
End Sub

Private Sub CmdDataShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "�鿴���ݽ��գ����ɱ�����յ����ݡ�"
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
   LabInfo.Caption = "�����������ݵ��������[%YY;01����]�����ΪYY�Ǳ��յ���ָ������һ֡�������ݲ�����ָ��ȴ�״̬���ڴ����趨��ָ�������ʽ(br48=1)ʱ��õ��������ָ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Start a single serial data output,[%YY;00����],number is YY instrument after receiving the order data and output a measurement wait state into the instruction."
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
   LabInfo.Caption = "ͨѶ�����������ѡ��[%YY;04����]�����ΪYY�Ǳ��յ���ָ���KK��Ҫ�������(��ʽ)������ݣ������Ʒֻ�ṩһ�ָ�ʽ��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Communication output data type selection,[%YY;04����], numbering as YY instrument after receiving the order requested by KK type (format) output data, only one format of conventional products."
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
   LabInfo.Caption = "��尴��������ֹ/��Чָ�[%YY;33;KK ]��KKȡֵ��Χ��0~1,00��������01��ֹ������"
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
   LabInfo.Caption = "�Ǳ���ʾ��ʱϨ��(��˸)��[%YY;05��KK����]��KKΪ�Ǳ���ʾϨ���ʱ��(KK X 0.01��)��Ȼ��������ʾ����ָ����Ǳ�Ĳ����͹������κ�Ӱ�죬��Ҫ���ھ�ʾ��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Short-term instruments show off (flashing),[%YY;05��KK����], KK is the instrument display off time (KK X 0.01 seconds), then re-show, the instruction on the instrument's measurement and work without any effect, mainly used for warning."
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
   LabInfo.Caption = "����ת����[%YY;12;KK����]��KKΪ����ʶ����(����ʶ���Ų��ܸ���)���յ�ָ����Ǳ�����������ʶ���ΪKK���̡�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Range conversion, [%YY;12;KK ����], KK identification number for the range (range identification number can not be changed). Instrument will be replaced after receipt of order to the range identifier for the KK scale."
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
   LabInfo.Caption = "���̿���ת����[%YY;13����]���յ�ָ����Ǳ��л�����һ�����̡�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Fast switching range, [%YY;13����], after receiving instructions switching to the next meter range."
End If
End Sub

Private Sub CmdReceive_Click()
MSComPort.InBufferCount = 0
MSComPort.OutBufferCount = 0
delay 100
If Picture1.Enabled = True Then Picture1.Enabled = False
If OptionALL.Value = True Then OptionALL.Value = False
If OptionOne.Value = True Then OptionOne.Value = False
If CmdReceive.Caption = "ָ�����" Or CmdReceive.Caption = "Receive by Command" Then
   If MSComPort.PortOpen = True Then
   RT = 1
    If Lan = 0 Then
       CmdReceive.Caption = "��������"
       ReceiveType = "ָ��"
    ElseIf Lan = 1 Then
       CmdReceive.Caption = "Receive Continuously"
       ReceiveType = "Instructions"
    End If
    MSComPort.Output = Trim("%00;03" & vbCr)
    TimInstroction = True
  End If
ElseIf CmdReceive.Caption = "��������" Or CmdReceive.Caption = "Receive Continuously" Then
   If MSComPort.PortOpen = True Then
   RT = 2
    If Lan = 0 Then
       CmdReceive.Caption = "ָ�����"
       ReceiveType = "����"
    ElseIf Lan = 1 Then
        CmdReceive.Caption = "Receive by Command"
       ReceiveType = "Continuous"
    End If
    TimInstroction = False
    MSComPort.Output = Trim("%00;02" & vbCr)
  End If
End If
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
End Sub

Private Sub CmdReceive_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "�����ڼ��ģʽ��ָ����պ����������໥�л���"
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
   LabInfo.Caption = "�ָ��趨�˲�ʱ�䣬[%YY;28����]���ָ��Ǳ�������������ʱ���趨��ʱ�䡣"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set the filter to restore the time, [%YY;28����], to resume the boot or restart the instrument when the set time."
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
   LabInfo.Caption = "������㣬[%YY;06����]���յ�ָ����Ǳ���ʱ�ľ��Բ���������Ϊ�����λ���Ǳ�˺���ʾ������Ϊ���ֵ���ݡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Set relative to zero,[%YY;06����], after the instrument will receive instructions at this time as the relative absolute zero measurements, instrumentation, then display the data for the absolute value of the data."
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
   LabInfo.Caption = "����״̬��λ��[%YY;09����]���յ�ָ����Ǳ�״̬�ֽ�3�е�bit0-5��0���Ա���к�������״̬�ж���"
ElseIf Lan = 1 Then
   LabInfo.Caption = "State reset button,[%YY;09����], after the instrument will receive the state bit 0-5 in the instruction byte 3 set to 0, in order to determine the status of follow-up button"
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
   LabInfo.Caption = "�������ݵ������������ֻ������ĳЩ����������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Start data output��This command applies only to specific instruments��"
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
   LabInfo.Caption = "ֹͣ���ݵ������������ֻ������ĳЩ����������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Stop data output��This command applies only to specific instruments��"
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
   LabInfo.Caption = "���Ͳ�����������ѡ��[%YY;20;KK����]��KK����Ӧ����ʾ�������£�[%YY;20;01����]��ʾ����ֵ���ݣ�[%YY;20;02����]��ʾ���ֵ���ݣ�[%YY;20;03����]�����λ�ľ���ֵ���ݡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Send measurement data type selection, [%YY;20;KK����], KK corresponding to the display as follows: [%YY;20;01����] shows the absolute value of the data, [%YY;20;02����] shows the relative zero data, [%YY;20;03����] the absolute value of the data relative zero."
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
   LabInfo.Caption = "�˲����˲�ʱ����С2����[%YY;26����]�����ڿ��ٸ��ٽ�Ծ�źţ��Ǳ�ִ��ָ�����ܻ�������ݲ�������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Filter filter time reduced 2 times, [%YY;26����], for fast tracking step signals, meters may occur after the implementation of instruction data fluctuation."
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
   LabInfo.Caption = "ֹͣ�������������[%YY;03����]�����ΪYY�Ǳ��յ���ָ���ֹͣ���Ͳ������ݣ�����ָ��ȴ�״̬��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Stop data continuous output,[%YY;03����], numbered YY instrument after receiving the order to stop sending measurement data into the command wait state."
End If
End Sub

Private Sub CombInstrumentNo_Click()
If Lan = 0 Then
ShowType = "����ģʽ"
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
   LabInfo.Caption = "��ʾ��˸��ʽѡ��[%YY;32;KK ]��KKȡֵ��Χ��0~3���յ���ָ����Ǳ����KKֵ������ʾ������˸���͡�"
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
   LabInfo.Caption = "����̵�������"
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
   LabInfo.Caption = "A/Dת����λ"
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
   LabInfo.Caption = "6000�Ǳ���ڶ๦�ܺ�״̬����ͨ������ָ����п��ƺ͵�����COMrָ���ʽΪ��[%YY;nn;KK����]��YYΪ00ʱ���е��Ǳ��ִ�������͵�ָ�����ݡ�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "6000 instrument state can be many functions and commands through the serial COM to control and adjust, COMr instruction format:[%YY;nn;KK����], when YY is 00 when all the instruments are the implementation of Directive"
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
    FrmSend.Caption = "����ָ��(�����������)"
    FrmSend.OptSendASC.Caption = "ASCIIָ��"
    FrmSend.OptSendHex.Caption = "ʮ������ָ��"
    FrmSend.LabelOnTop(20).Caption = "�Զ���������"
    FrmSend.LabelOnTop(21).Caption = "����/��"
    FrmSend.ChkSent.Caption = "�Զ�����"
    FrmSend.CmdEmpty.Caption = "�������"
    FrmSend.CmdManual.Caption = "�ֶ�����"
    FrmSend.CmdExit.Caption = "�˳�"
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
      MsgBox "�ļ�" & (Left(txtFile, Len(txtFile) - 4) & ".xls") & "�Ѵ��ڣ��������Ѿ�ת�����ˡ�"
      OpenFileName = ""
      Exit Sub
    End If
On Error GoTo L

    '����excel����
    Dim XlApp As Excel.Application
    Dim XlWb As Excel.Workbook
    Dim XlSt As Excel.Worksheet
    Set XlApp = CreateObject("Excel.Application")
    If XlApp Is Nothing Then
       MsgBox "���鱾���Ƿ�װ��Microsoft Excel���"
       Exit Sub
    End If
    Set XlWb = XlApp.Workbooks.Add
    Set XlSt = XlWb.Worksheets(1)
    XlApp.ActiveWorkbook.SaveAs (Left(txtFile, Len(txtFile) - 4) & ".xls")
    XlApp.Visible = False
    XlApp.Rows.HorizontalAlignment = xlVAlignCenter
    '��ʼת��
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
                XlSt.Cells(j, 1) = "ע���ֱ���ʾ���ǣ���š�ԭʼ���ݡ��Ǳ���ʾ���ݡ����ڡ�ʱ�䡣"
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
    MsgBox "����ɹ�"
L:
    Exit Sub
    MsgBox "ת���г��ִ���"
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
   Case WM_LBUTTONUP '���
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
If TxtShowInfo.Text = "����״̬���ر�" Then TxtShowInfo.Text = "Port Status��Closed"
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
If CmdMode.Caption = "���ģʽ" Then
   CmdMode.Caption = "Detection mode"
ElseIf CmdMode.Caption = "����ģʽ" Then
   CmdMode.Caption = "Debug mode"
End If
If CmdReceive.Caption = "ָ�����" Then
    CmdReceive.Caption = "Receive by Command"
ElseIf CmdReceive.Caption = "��������" Then
   CmdReceive.Caption = "Receive Continuously"
End If
If ReceiveType = "ָ��" Then
ReceiveType = "Instructions"
ElseIf ReceiveType = "����" Then
ReceiveType = "Continuous"
End If
If ShowType = "����ģʽ" Then
ShowType = "Debug"
ElseIf ShowType = "���ģʽ" Then
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
FrmInstruct.Caption = "Instruction List (YY: instrument No., ����: carriage return)"

FrmDataShow.Caption = "Receive data analysis"
FrmDataShow.Label1(0).Caption = "NO."
FrmDataShow.Label1(1).Caption = "Hex"
FrmDataShow.Label1(2).Caption = "ASCII"
FrmDataShow.Option1(0).Caption = "Binary"
FrmDataShow.Option1(1).Caption = "Graphical"
If FrmDataShow.CmdOK(0).Caption = "��ͣ" Then FrmDataShow.CmdOK(0).Caption = "Pause"
If FrmDataShow.CmdOK(0).Caption = "����" Then FrmDataShow.CmdOK(0).Caption = "Continue"
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
   LabInfo.Caption = "�������úõĶ˿�"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Open the COM has been well set"
End If
End Sub

Private Sub ComExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "�˳�����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Exit"
End If
End Sub

Private Sub ComRange_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "ͨ����  ��Ҫ���ܣ��鿴���������ͨ��"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Channel Function: To view or change data channels"
End If
End Sub

Private Sub ComDigits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "��λ��  ��Ҫ���ܣ����е�λ�任����ʾ��ͬ��λ���µĲ����������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Units features: flat exchange, display the data under different unit system results"
End If
End Sub

Private Sub ComRun_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   If ShpOn(3).Visible = False Then
      LabInfo.Caption = "�رռ�  ��Ҫ���ܣ��ڳ������״̬�����ڹر��Ǳ���ʾ�����͵�������"
   Else
      LabInfo.Caption = "�����  ��Ҫ���ܣ��ڷ�ֵ����״̬��������������ֵķ�ֵ����"
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
   LabInfo.Caption = "��ֵ��  ��Ҫ���ܣ�������رշ�ֵ����״̬����ֵ����״̬ʱ<��ֵ>ָʾ����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Peak  Function: On or off peak measurement, peak state <peak> indicator light"
End If
End Sub

Private Sub ComPrint_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "��ӡ��  ��Ҫ���ܣ���ӡ��������"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Prinf Function: Print measurements"
End If
End Sub

Private Sub ComReset_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "��λ��  ��Ҫ���ܣ����¸�λ�����Ǳ����ڲ����趨���鿴����ֵ�궨�Ȳ���"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Reset  Function: Re-start the meter reset for the parameter settings, view, force the value of calibration and other operations"
End If
End Sub

Private Sub ComZero_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "�����  ��Ҫ���ܣ�ʹ������ʾ�������ԭ������״̬ʹ<����>ָʾ�Ƶ���"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Zero  Function: Zero or restore data, zero state <Zero> indicator light"
End If
End Sub


Private Sub CmdShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "ֹͣ�������ʾ�������е�����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Stop or continue to show the contents of the receiving area"
End If
End Sub

Private Sub CmdClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "��ս������е�����"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Empty the contents of the reception area"
End If
End Sub

Private Sub AutoClear_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "ѡ�����Զ���ս������е�����"
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
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
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
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
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
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComDigits_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;04" & vbCr)
   TimerCheckStart
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComShow_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;03" & vbCr)
   TimerCheckStart
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComRun_Click()
On Error GoTo ErrHndl
MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;02" & vbCr)
   TimerCheckStart
If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
ErrHndl:
    Debug.Print "Error No. " & Err.Number & ": Error Description: " & Err.Description
End Sub

Private Sub ComZero_Click()
On Error GoTo ErrHndl

MSComPort.OutBufferCount = 0
delay 100
   MSComPort.Output = Trim("%00;17;01" & vbCr)
   TimerCheckStart

If ReceiveType = "ָ��" Or ReceiveType = "Instructions" Then TimInstroction = True
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
        MenOpenCom.Caption = "�رն˿�" & "(&O)"
        ComComOpen.Caption = "�رն˿�"
        FrmMain.Caption = "6000�������ֲ�����(V13-2.0718)"
        ShowType = "���ģʽ"
        ReceiveType = "ָ��"
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
      ReceiveCounts = 0    '����֡��
      ReceiveTrueCounts = 0 '������Ч֡��
    Else
    TimRecover.Enabled = False
    If TimInstroction.Enabled = True Then TimInstroction.Enabled = False
        MSComPort.PortOpen = False
        FrmMain.Caption = "6000�������ֲ�����(V13-2.0718)"
        LabErr.Caption = ""
        TextShow.Text = ""
        ShpComOn.Visible = False
        ShpRevOn.Visible = False
        For i% = 0 To 5
            If ShpOn(i).Visible Then ShpOn(i).Visible = False
        Next i
        ShpErrOn.Visible = False

        If Lan = 0 Then
          ComComOpen.Caption = "�򿪶˿�"
          MenOpenCom.Caption = "�򿪶˿�" & "(&O)"
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
        LabErr.Caption = "��ʾ�����ڿ����ѱ�ռ�ã������¼������"
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
        Print #1, "�ֱ���ʾ���ǣ���š�ԭʼ���ݡ��Ǳ���ʾ���ݡ����ڡ�ʱ��"
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
        FrmDataShow.FrmDataSaveStatus.Caption = "���ݼ�¼ - " & "��ǰΪ������¼������ʽ���Ѽ�¼" & RecordNumber & "������"
        Else
        FrmDataShow.FrmDataSaveStatus.Caption = "Data save - " & "Save one by one" & RecordNumber
        End If
        DatasSaveStyle = 1
    ElseIf DatasSaveStyle = 3 Then
        If Lan = 0 Then
        FrmDataShow.FrmDataSaveStatus.Caption = "���ݼ�¼ - " & "��ǰΪ������¼������ʽ���Ѽ�¼" & RecordNumber & "������"
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
    FrmMain.Caption = "6000��ͨѶ������(V13-2.0718)"
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
    .szTip = "6000�������ֲ�����" & vbCr & "�������" & vbCr & "V1.0.0" & vbNullChar
    .dwState = 0
    .dwStateMask = 0
    If LogoLan = 0 Then
    .szInfoTitle = "��ӭʹ��" & vbNullChar
    ElseIf LogoLan = 1 Then
    .szInfoTitle = "Welcome" & vbNullChar
    End If
    If LogoLan = 0 Then
    .szInfo = "������ͼ�꽫��ʾ/����������" & vbNullChar
    ElseIf LogoLan = 1 Then
    .szInfo = "Click the icon to display/hide the program��" & vbNullChar
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
MsgBox IIf(Lan, "Thanks for using our pruducts. Please read the manual.", "��лʹ�ñ�����������������޹�˾��Ʒ�����Ķ���Ʒ˵���顣"), vbInformation
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
         Debug.Print "����"
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
       Xianshileixing = "����01"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "01" Then
       Xianshileixing = "����02"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 7, 2) = "10" Then
       Xianshileixing = "����03"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 5, 1) = "0" Then
       Anjiancaozuo = "������"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 5, 1) = "1" Then
       Anjiancaozuo = "����ֹ"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 4, 1) = "0" Then
       Fanzhuandianliu = "+"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 4, 1) = "1" Then
       Fanzhuandianliu = "-"
    End If
    If Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "0" Then
       Dianzu4or2 = "��4"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "1" Then
       Dianzu4or2 = "��2"
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
       Dianzu4or2 = "��4"
    ElseIf Mid(DEC_to_BIN(CLng(Str(StateByte5)), 8), 3, 1) = "1" Then
       Dianzu4or2 = "��2"
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
ShowType = "����ģʽ"
ElseIf Lan = 1 Then
ShowType = "Debug Mode"
End If
Picture1.Enabled = True
CombInstrumentNo.Enabled = False
YY = "00"
End Sub

Private Sub OptionALL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Lan = 0 Then
   LabInfo.Caption = "ѡ�д��������������Ƶ�ǰ���������ӵ������Ǳ�"
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
   LabInfo.Caption = "ѡ�д������������ֻ������ѡ�����Ǳ�ʶ��ŵ��Ǳ���ǰ���ڵ������Ǳ���Ӱ�졣"
ElseIf Lan = 1 Then
   LabInfo.Caption = "Select the following command and control of the selected only the instrument meter recognition, the serial other instrument is not affected."
End If
End Sub



Private Sub TimDetection_Timer()
If JudgeData <> "" Then
   JudgeData = ""
Else
   TextShow.Text = "��������������������������������������"
   If Lan = 0 Then
       TxtShowInfo = "��ǰ����ģʽ��" & ReceiveType
   ElseIf Lan = 1 Then
      TxtShowInfo = "Receive mode:" & ReceiveType
   End If
   For i% = 0 To 7
       If ShpOn(i).Visible Then ShpOn(i).Visible = False
   Next i
   ShpRevOn.Visible = False
   If ReceiveType = "����" Or ReceiveType = "Continuous" Then
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
   If ReceiveFlag Then LabErr.Caption = "��ʾ��û�����ݽ��գ������Ƿ��ѿ������������˿������Ƿ���Ч"
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
        .Create CmdSendExtremum, " ָ������                     ��ʾ��������          " & vbCr & "[%YY;21;04����]               ˲̬���ֵ����" & vbCr & "[%YY;21;05����]               ˲̬��Сֵ����" & vbCr & "[%YY;21;06����]               ˲̬���ֵ����Сֵ�Ĳ�ֵ����" & vbCr & "[%YY;21;07����]               ƽ�����ֵ����" & vbCr & "[%YY;21;08����]               ƽ����Сֵ����" & vbCr & "[%YY;21;09����]               ƽ�����ֵ����Сֵ�Ĳ�ֵ����" & vbCr & "[%YY;21;10����]               ��ʾƽ��ֵ����(�ɼ���״̬�л���ƽ��ֵ״̬)" & vbCr & vbCr & "[%YY;21;01����]               ��ʾ����ƽ��ֵ����(��ƽ��ֵ״̬����Ч)" & vbCr & "[%YY;21;02����]               ��ʵ���ƽ��ֵ����(��ƽ��ֵ״̬����Ч)" & vbCr & "[%YY;21;03����]               ��Ч", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "COMr21 ���ͼ�ֵ������������ѡ��[%YY;21;KK����]��KK����Ӧ����ʾ�������£�", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "����", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, CmdSendExtremum.Name
End Sub

Private Sub ButtonInstruct()
    Set tooltip = New cToolTip
    With tooltip
        .Create CmdKeyFunction, "    ָ������            ˵��" & vbCr & "[%YY;17;01����]   ��Ч�����㡷���� " & vbCr & "[%YY;17;02����]   ��Ч�����С����� " & vbCr & "[%YY;17;03����]   ��Ч����ʾ������ " & vbCr & "[%YY;17;04����]   ��Ч��λ�������� " & vbCr & "[%YY;17;05����]   ��Ч�����̡����� " & vbCr & "[%YY;17;06����]   ��Ч����ӡ������ " & vbCr & "[%YY;17;07����]   ��Ч����λ������ " & vbCr & "[%YY;17;08����]   ��Ч��WK1������ " & vbCr & "[%YY;17;09����]   ��Ч��WK2������ ", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "COMr17 ���ڰ���ָ�[%YY;17;KK����]��KK����Ӧ�İ����������£�", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "����", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, CmdKeyFunction.Name
End Sub

Private Sub ButtonCOMr30()
    Set tooltip = New cToolTip
    With tooltip
        .Create Command1, "    ָ������            ˵��" & vbCr & "[%YY;18;00����]   SP1��SP2��Ϩ������̵����Ͽ� " & vbCr & "[%YY;18;01����]   SP1������SP1�̵����պ� " & vbCr & "[%YY;18;02����]   SP2������SP2�̵����պ� " & vbCr & "[%YY;18;03����]   SP1��SP2�ƾ������̵������պ�", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "COMr18 ���ڰ���ָ�[%YY;18;KK����]��KK����Ӧ�İ����������£�", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "����", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Command1.Name
End Sub

Private Sub CmdfanzhuandianliuShow()
    Set tooltip = New cToolTip
    With tooltip
        .Create Command6, "����ָ��������6000-11�͵��������ʽ���ر�ָ���Ĳ�Ʒ)" & vbCr & vbCr & "    ָ������            ˵��" & vbCr & "[%YY;35;00����]   �趨���������������(���淽ʽ) " & vbCr & "[%YY;35;01����]   �趨���跴���������(��ת��ʽ) ", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "������������ָ�[%YY;35;KK ]��KKȡֵ��Χ��0~1", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "����", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Command6.Name
End Sub

Private Sub Cmd4xian2xianShow()
    Set tooltip = New cToolTip
    With tooltip
        .Create Command7, "     ����ָ��������6000-11�͵��������ʽ���ر�ָ���Ĳ�Ʒ)" & vbCr & vbCr & "    ָ������            ˵��" & vbCr & "[%YY;36;00����]   �趨Ϊ4�ߵ��������ʽ(����Ĭ�Ϸ�ʽ) " & vbCr & "[%YY;36;01����]   �趨Ϊ2�ߵ��������ʽ ", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "4��/2�ߵ��������ʽ�趨ָ�[%YY;36;KK ]��KKȡֵ��Χ��0~1", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "����", 9, Italic:=flase
    End With
    Tooltips.Add tooltip, Command7.Name
End Sub

Private Sub CmdgaugeshowstyleShow()
    Set tooltip = New cToolTip
    With tooltip
        .Create Cmdgaugeshowstyle, "    ָ������            ˵��" & vbCr & "[%YY;34;00����]   ��ʾ��������(������ʾ��ʽ)" & vbCr & "[%YY;34;01����]   ��ʾ�Ǳ�ǰ������ʶ���Lr32����LH=XX��XX=Lr32" & vbCr & "[%YY;34;02����]   ��ʾ�Ǳ�ͨѶʶ���br42����Ud=XX��XX=br42", TTBalloonIfActive Or TTSpeak, False, TTIconNone, "��ʾ����ѡ��ָ�[%YY;34;KK ]��KKȡֵ��Χ��0~2", vbBlack, RGB(240, 248, 255), 0, 30000
        .SubstituteFont "����", 9, Italic:=flase
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
