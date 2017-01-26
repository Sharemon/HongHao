VERSION 5.00
Begin VB.Form Lanset 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   3585
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      TabIndex        =   9
      Text            =   "TCP/IP"
      Top             =   1725
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2000
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   2
      Left            =   1560
      TabIndex        =   6
      Text            =   "2000"
      Top             =   1275
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Text            =   "5025"
      Top             =   765
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Text            =   "169.254.4.10"
      Top             =   300
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   1800
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   630
   End
End
Attribute VB_Name = "Lanset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RmHost = Text1(0).Text
RemPort = Val(Text1(1).Text)
Winsock1.LocalPort = Val(Text1(2).Text)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = IIf(Lang, "Lan settings", "LAN设置")
Text1(0).Text = RmHost
Text1(1).Text = RemPort
Label1(0).Caption = IIf(Lang, "IP of the instrument:", "仪器IP:")
Label1(1).Caption = IIf(Lang, "Port of the instrument:", "仪器端口：")
Label1(2).Caption = IIf(Lang, "Local port:", "本地端口：")
Label1(3).Caption = IIf(Lang, "IP protocol:", "IP协议：")
Combo1.AddItem "TCP/IP"
Combo1.AddItem "UDP"
Command1.Caption = IIf(Lang, "OK", "确定")
Command2.Caption = IIf(Lang, "Cancel", "取消")
End Sub
